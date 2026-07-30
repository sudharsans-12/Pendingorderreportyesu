# -*- coding: utf-8 -*-
# VERSION: v1 - Email Sender via SMTP
import smtplib
import os
import io
import base64
import requests
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import excel_formatter

def _smtp_connect_and_login(smtp_config):
    """Connect to SMTP server and perform login, forcing LOGIN/PLAIN to avoid auth negotiation loop."""
    import base64
    host = smtp_config.get("host")
    port = int(smtp_config.get("port", 587))
    user = smtp_config.get("user")
    password = smtp_config.get("password")
    use_tls = smtp_config.get("use_tls", True)
    
    if use_tls:
        server = smtplib.SMTP(host, port, timeout=15)
        server.ehlo()
        server.starttls()
        server.ehlo()
    else:
        server = smtplib.SMTP_SSL(host, port, timeout=15)
        
    # Perform manual custom login to prevent smtplib's infinite challenge-response negotiation loops
    try:
        # Initiate AUTH LOGIN
        code, resp = server.docmd("AUTH", "LOGIN")
        if code == 334:
            # Send Base64 Username
            user_b64 = base64.b64encode(user.encode('utf-8')).decode('utf-8')
            code, resp = server.docmd(user_b64)
            if code == 334:
                # Send Base64 Password
                pass_b64 = base64.b64encode(password.encode('utf-8')).decode('utf-8')
                code, resp = server.docmd(pass_b64)
                if code == 235:
                    return server
                else:
                    raise smtplib.SMTPAuthenticationError(code, f"Authentication failed: {resp}")
            else:
                raise smtplib.SMTPAuthenticationError(code, f"Username rejected or failed: {resp}")
        else:
            # Fallback to standard login if AUTH LOGIN is not supported directly
            server.login(user, password)
            return server
    except smtplib.SMTPAuthenticationError as e:
        server.quit()
        raise e
    except Exception as e:
        # Fallback to standard login
        try:
            # Force SMTP AUTH to basic LOGIN/PLAIN to avoid server negotiation loops (fixes infinite AUTH loops)
            if hasattr(server, 'esmtp_features') and 'auth' in server.esmtp_features:
                auth_methods = server.esmtp_features['auth'].split()
                basic_auths = [m for m in auth_methods if m.upper() in ['LOGIN', 'PLAIN']]
                if basic_auths:
                    server.esmtp_features['auth'] = ' '.join(basic_auths)
            server.login(user, password)
            return server
        except Exception as fallback_err:
            server.quit()
            raise fallback_err

def test_smtp_connection(host, port, user, password, use_tls=True):
    """Test connection to the SMTP server."""
    try:
        missing = []
        if not str(host or "").strip():
            missing.append("SMTP Server Host")
        if not str(user or "").strip():
            missing.append("SMTP Username")
        if not str(password or "").strip():
            missing.append("SMTP Password")
        if missing:
            return False, f"SMTP configuration details are incomplete. Missing: {', '.join(missing)}."
        cfg = {"host": host, "port": port, "user": user, "password": password, "use_tls": use_tls}
        server = _smtp_connect_and_login(cfg)
        server.quit()
        return True, "Successfully connected to SMTP server!"
    except Exception as e:
        return False, f"SMTP Connection failed: {str(e)}"


def test_brevo_connection(api_key):
    """
    Verify a Brevo (formerly Sendinblue) API key by calling their account
    endpoint. Used as an alternative to SMTP when Office 365 blocks
    Authenticated SMTP / legacy auth.
    """
    try:
        if not str(api_key or "").strip():
            return False, "Brevo API Key is required."
        resp = requests.get(
            "https://api.brevo.com/v3/account",
            headers={"api-key": api_key, "Accept": "application/json"},
            timeout=15
        )
        if resp.status_code == 200:
            data = resp.json()
            acct_email = data.get("email", "")
            return True, f"Successfully connected to Brevo!{f' (Account: {acct_email})' if acct_email else ''}"
        elif resp.status_code == 401:
            return False, "Brevo connection failed: Invalid API key (401 Unauthorized)."
        else:
            return False, f"Brevo connection failed: {resp.status_code} - {resp.text[:200]}"
    except Exception as e:
        return False, f"Brevo connection failed: {str(e)}"


def _send_via_brevo(api_key, sender_email, sender_name, to_email, cc_email, subject, html_body,
                     attachment_bytes=None, attachment_filename=None):
    """
    Send an email through the Brevo transactional email API (HTTPS, API-key
    based) instead of SMTP. sender_email must be a verified sender in the
    Brevo account.
    """
    try:
        if not str(api_key or "").strip():
            return False, "Brevo API Key is required."
        if not str(sender_email or "").strip():
            return False, "Sender Email is required (must be a verified sender in Brevo)."

        to_list = [{"email": e.strip()} for e in str(to_email or "").split(",") if e.strip()]
        if not to_list:
            return False, "Recipient 'To' email address is required."

        payload = {
            "sender": {"email": sender_email, "name": sender_name or "Graas Team"},
            "to": to_list,
            "subject": subject,
            "htmlContent": html_body
        }
        cc_list = [{"email": e.strip()} for e in str(cc_email or "").split(",") if e.strip()]
        if cc_list:
            payload["cc"] = cc_list
        if attachment_bytes is not None and attachment_filename:
            payload["attachment"] = [{
                "content": base64.b64encode(attachment_bytes).decode("utf-8"),
                "name": attachment_filename
            }]

        resp = requests.post(
            "https://api.brevo.com/v3/smtp/email",
            headers={"api-key": api_key, "Content-Type": "application/json", "Accept": "application/json"},
            json=payload,
            timeout=30
        )
        if resp.status_code in (200, 201):
            return True, "Email sent successfully via Brevo!"
        else:
            return False, f"Brevo send failed: {resp.status_code} - {resp.text[:300]}"
    except Exception as e:
        return False, f"Brevo send failed: {str(e)}"


# ============================================================================
# Google OAuth2 / Gmail API - "Sign in with Google" support
# ============================================================================
# This lets the app send email as a Gmail/Google Workspace account via
# Google's official API instead of SMTP - no app password, no Office 365
# tenant policy issues. Since the app sends automatically/on a schedule with
# nobody present to click "Allow" each time, the flow is: sign in with Google
# ONCE to get a refresh token, then the app silently exchanges that refresh
# token for a fresh access token on every send (refresh tokens don't expire
# unless revoked).

GOOGLE_GMAIL_SEND_SCOPE = "https://www.googleapis.com/auth/gmail.send"

def build_google_auth_url(client_id, redirect_uri):
    """Build the Google OAuth consent URL that requests Gmail-send access."""
    from urllib.parse import urlencode
    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": GOOGLE_GMAIL_SEND_SCOPE,
        "access_type": "offline",   # required to get a refresh_token back
        "prompt": "consent",        # forces a refresh_token even on repeat sign-ins
    }
    return "https://accounts.google.com/o/oauth2/v2/auth?" + urlencode(params)


def exchange_google_code_for_tokens(client_id, client_secret, redirect_uri, code):
    """Exchange a one-time OAuth authorization code for access + refresh tokens."""
    try:
        resp = requests.post(
            "https://oauth2.googleapis.com/token",
            data={
                "code": code,
                "client_id": client_id,
                "client_secret": client_secret,
                "redirect_uri": redirect_uri,
                "grant_type": "authorization_code",
            },
            timeout=15
        )
        if resp.status_code == 200:
            return True, resp.json()
        return False, resp.text[:300]
    except Exception as e:
        return False, str(e)


def _get_google_access_token(client_id, client_secret, refresh_token):
    """Exchange a stored refresh token for a fresh (short-lived) access token."""
    resp = requests.post(
        "https://oauth2.googleapis.com/token",
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
        },
        timeout=15
    )
    if resp.status_code == 200:
        token = resp.json().get("access_token")
        if token:
            return token
        raise Exception("Google did not return an access token.")
    raise Exception(f"Failed to refresh Google access token: {resp.status_code} - {resp.text[:200]}")


def test_google_connection(client_id, client_secret, refresh_token):
    """Verify Google OAuth credentials by attempting to refresh an access token."""
    try:
        missing = []
        if not str(client_id or "").strip():
            missing.append("Google Client ID")
        if not str(client_secret or "").strip():
            missing.append("Google Client Secret")
        if not str(refresh_token or "").strip():
            missing.append("Refresh Token (click 'Sign in with Google' first)")
        if missing:
            return False, f"Google configuration incomplete. Missing: {', '.join(missing)}."
        _get_google_access_token(client_id, client_secret, refresh_token)
        return True, "Successfully connected to Google - access token refreshed OK!"
    except Exception as e:
        return False, f"Google connection failed: {str(e)}"


def _send_via_gmail_api(client_id, client_secret, refresh_token, sender_email, to_email, cc_email,
                         subject, html_body, attachment_bytes=None, attachment_filename=None):
    """Send an email via the Gmail API using OAuth2 (no SMTP, no app password)."""
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    try:
        if not to_email or "@" not in to_email:
            return False, "Recipient 'To' email address is required and must be valid."

        access_token = _get_google_access_token(client_id, client_secret, refresh_token)

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = to_email
        if cc_email:
            msg["Cc"] = cc_email
        msg["Subject"] = subject
        msg.attach(MIMEText(html_body, "html"))
        if attachment_bytes is not None and attachment_filename:
            part = MIMEApplication(attachment_bytes, _subtype="xlsx")
            part.add_header("Content-Disposition", "attachment", filename=attachment_filename)
            msg.attach(part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        resp = requests.post(
            "https://gmail.googleapis.com/gmail/v1/users/me/messages/send",
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            json={"raw": raw},
            timeout=30
        )
        if resp.status_code in (200, 202):
            return True, "Email sent successfully via Gmail!"
        return False, f"Gmail send failed: {resp.status_code} - {resp.text[:300]}"
    except Exception as e:
        return False, f"Gmail send failed: {str(e)}"


def send_seller_report_email(smtp_config, seller_name, recipient_email, seller_df, discrepancies_df=None):
    """
    Generate an Excel sheet for the seller, build a nice HTML summary, and send it via SMTP.
    """
    if seller_df.empty:
        return False, "No data available for this seller."

    if not recipient_email or "@" not in recipient_email:
        return False, f"Invalid or missing recipient email address: '{recipient_email}'"

    host = smtp_config.get("host")
    port = int(smtp_config.get("port", 587))
    user = smtp_config.get("user")
    password = smtp_config.get("password")
    use_tls = smtp_config.get("use_tls", True)
    sender_email = smtp_config.get("sender_email", user)

    # == 1. Create the Excel Attachment in Memory ==============================
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        seller_df.to_excel(writer, sheet_name="Pending Orders", index=False)
        excel_formatter.format_data_sheet(writer.sheets["Pending Orders"], seller_df)
        
        if discrepancies_df is not None and not discrepancies_df.empty:
            # Filter discrepancies for this seller's orders
            order_ids = set(seller_df.iloc[:, 0].dropna().apply(lambda x: str(x).strip()).tolist())
            
            # Find the Order ID column in discrepancies
            disc_id_col = next((c for c in discrepancies_df.columns if "order" in c.lower()), "")
            if disc_id_col:
                seller_disc = discrepancies_df[discrepancies_df[disc_id_col].astype(str).str.strip().isin(order_ids)]
                if not seller_disc.empty:
                    seller_disc.to_excel(writer, sheet_name="Status Discrepancies", index=False)
                    excel_formatter.format_data_sheet(writer.sheets["Status Discrepancies"], seller_disc)
                    
    excel_data = excel_buffer.getvalue()

    # == 2. Build HTML Body ==================================================-
    total_orders = len(seller_df)
    
    # Check if there are urgent SLAs (e.g. within today/tomorrow)
    # Since we don't have standard date parsing here, we just display the first 5 pending orders as a preview
    preview_df = seller_df.head(10)
    
    # Convert preview to HTML table
    table_html = preview_df.to_html(classes="table", index=False, border=0)
    
    # Custom CSS style for email
    email_html = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                color: #333333;
                background-color: #f9f9f9;
                margin: 0;
                padding: 20px;
            }}
            .container {{
                max-width: 700px;
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 30px;
                margin: 0 auto;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
            }}
            .header {{
                border-bottom: 2px solid #0066cc;
                padding-bottom: 15px;
                margin-bottom: 20px;
            }}
            .header h2 {{
                color: #0066cc;
                margin: 0;
                font-size: 24px;
            }}
            .summary-box {{
                background-color: #f0f7ff;
                border-left: 4px solid #0066cc;
                padding: 15px;
                margin-bottom: 20px;
                border-radius: 0 4px 4px 0;
            }}
            .summary-title {{
                font-weight: bold;
                font-size: 16px;
                color: #004499;
                margin-bottom: 5px;
            }}
            .table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 15px;
                margin-bottom: 20px;
                font-size: 13px;
            }}
            .table th {{
                background-color: #f2f2f2;
                color: #444444;
                font-weight: bold;
                text-align: left;
                padding: 10px;
                border-bottom: 1px solid #dddddd;
            }}
            .table td {{
                padding: 10px;
                border-bottom: 1px solid #eeeeee;
            }}
            .footer {{
                font-size: 12px;
                color: #888888;
                border-top: 1px solid #e0e0e0;
                padding-top: 15px;
                margin-top: 30px;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h2>Daily Pending Order & SLA Report</h2>
                <p style="margin: 5px 0 0 0; color: #666666;">Store: <strong>{seller_name}</strong></p>
            </div>
            
            <p>Dear Seller partner,</p>
            <p>Please find attached the daily Pending Order and SLA Status validation report for your store. Kindly review the details to ensure prompt fulfillment and address any status mismatches highlighted.</p>
            
            <div class="summary-box">
                <div class="summary-title">Report Summary</div>
                Total Pending Orders: <strong>{total_orders}</strong><br>
                Please check the attached Excel sheet for the full list and any discrepancies identified.
            </div>

            <h3>Pending Orders Preview (Top 10 Rows)</h3>
            {table_html}

            <p style="font-size: 14px;"><strong>Note:</strong> A complete report with all order statuses and validation checks has been attached to this email as an Excel spreadsheet.</p>

            <p>Best regards,<br>
            <strong>Operations & Analytics Team</strong></p>
            
            <div class="footer">
                This is an automated report generated by the Status Validation Analyzer. Please do not reply directly to this email.
            </div>
        </div>
    </body>
    </html>
    """

    # == 3. Build Multipart Message ==========================================-
    msg = MIMEMultipart()
    msg["From"] = f"Operations Team <{sender_email}>"
    msg["To"] = recipient_email
    msg["Subject"] = f"Pending Order & SLA Report - {seller_name}"

    msg.attach(MIMEText(email_html, "html"))

    # Attach Excel file
    attachment = MIMEApplication(excel_data, _subtype="xlsx")
    attachment.add_header("Content-Disposition", "attachment", filename=f"Pending_Orders_Report_{seller_name.replace(' ', '_')}.xlsx")
    msg.attach(attachment)

    # == 4. Send Email via SMTP ==============================================-
    try:
        server = _smtp_connect_and_login(smtp_config)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        return True, "Email sent successfully!"
    except Exception as e:
        return False, f"Failed to send email: {str(e)}"

def _build_country_pivot_email_html(country, pivot_df, summary_df, ref_date_dmy):
    """
    Builds an inline-styled HTML block (pivot table + colored Order Summary
    box side-by-side) for the country report email body - a direct mirror of
    the internal dashboard's Excel "Summary" sheet:
      - Column headers use the pivot's actual column names (e.g. "Channel",
        "OMS Order Status") rather than relabeled text.
      - Status text (e.g. "Not in OMS") is shown exactly as-is, never renamed.
      - Date columns are colored red/orange/green (breached/handover-today/
        within-SLA) relative to the reference date - uniformly for every row,
        with no special-case override for any particular status.
      - The right-hand summary box shows all 6 dashboard metrics with the
        same colors as the dashboard (Overdue=red, Handover Today=orange,
        Order status at NEW=dark red, Within SLA=green, Not reflected in
        OMS=pale green, Unpaid Orders=grey).
    Falls back to a simple "no data" message if the pivot is empty.
    """
    import pandas as pd
    from datetime import datetime

    if pivot_df is None or pivot_df.empty:
        return "<p><em>No pending order data available for this country.</em></p>"

    try:
        today_dt = datetime.strptime(ref_date_dmy, "%d-%m-%Y") if ref_date_dmy else datetime.today()
    except Exception:
        today_dt = datetime.today()

    cols = list(pivot_df.columns)
    marketplace_col = cols[0]
    status_col = cols[1] if len(cols) > 1 else cols[0]
    date_cols = [c for c in cols[2:] if c != "Grand Total"]
    has_grand_total = "Grand Total" in cols

    def _cell_color(col_name):
        # Uniform date-based coloring for every row - no per-status override,
        # matching the dashboard exactly.
        try:
            d = datetime.strptime(str(col_name), "%d-%m-%Y")
        except Exception:
            return None
        if d.date() < today_dt.date():
            return "#FF0000"  # red - overdue/breached
        elif d.date() == today_dt.date():
            return "#FFA500"  # orange - handover today
        else:
            return "#92D050"  # green - within SLA

    # Pre-compute rowspans so repeated Channel values are visually merged,
    # same as the dashboard's merged cells.
    marketplace_counts = pivot_df[marketplace_col].astype(str).value_counts()
    printed_marketplace = set()

    body_rows = []
    for _, row in pivot_df.iterrows():
        mp_val = str(row[marketplace_col])
        status_val = str(row[status_col]) if status_col in row else ""
        is_grand_total_row = (mp_val.strip() == "Grand Total")

        row_html = "<tr>"
        if is_grand_total_row:
            row_html += '<td colspan="2" style="border:1px solid #bfbfbf;padding:6px 10px;font-weight:700;background:#f2f2f2;">Grand Total</td>'
        else:
            if mp_val not in printed_marketplace:
                rowspan = int(marketplace_counts.get(mp_val, 1))
                row_html += f'<td rowspan="{rowspan}" style="border:1px solid #bfbfbf;padding:6px 10px;font-weight:600;background:#ffffff;vertical-align:middle;">{mp_val}</td>'
                printed_marketplace.add(mp_val)
            # Status text shown exactly as-is - "Not in OMS" stays "Not in OMS".
            row_html += f'<td style="border:1px solid #bfbfbf;padding:6px 10px;background:#ffffff;">{status_val}</td>'

        for dc in date_cols:
            val = row.get(dc, 0)
            try:
                val_int = int(val) if pd.notna(val) else 0
            except Exception:
                val_int = 0
            val_disp = "" if val_int == 0 else str(val_int)
            style = "border:1px solid #bfbfbf;padding:6px 10px;text-align:center;"
            if is_grand_total_row:
                style += "font-weight:700;background:#f2f2f2;"
            elif val_disp:
                bg = _cell_color(dc)
                if bg:
                    style += f"background:{bg};color:#ffffff;font-weight:600;"
                else:
                    style += "background:#ffffff;"
            else:
                style += "background:#ffffff;"
            row_html += f'<td style="{style}">{val_disp}</td>'

        if has_grand_total:
            gt_val = row.get("Grand Total", 0)
            try:
                gt_disp = int(gt_val) if pd.notna(gt_val) else 0
            except Exception:
                gt_disp = 0
            row_html += f'<td style="border:1px solid #bfbfbf;padding:6px 10px;text-align:center;font-weight:700;background:#f2f2f2;">{gt_disp}</td>'
        row_html += "</tr>"
        body_rows.append(row_html)

    header_html = (
        '<tr>'
        f'<th style="border:1px solid #bfbfbf;padding:8px 10px;background:#ffffff;color:#000000;text-align:left;">{marketplace_col}</th>'
        f'<th style="border:1px solid #bfbfbf;padding:8px 10px;background:#ffffff;color:#000000;text-align:left;">{status_col}</th>'
    )
    for dc in date_cols:
        header_html += f'<th style="border:1px solid #bfbfbf;padding:8px 10px;background:#ffffff;color:#000000;">{dc}</th>'
    if has_grand_total:
        header_html += '<th style="border:1px solid #bfbfbf;padding:8px 10px;background:#ffffff;color:#000000;">Grand Total</th>'
    header_html += '</tr>'

    pivot_table_html = f"""
    <table style="border-collapse:collapse;width:100%;font-size:13px;">
        <tr><td colspan="{2 + len(date_cols) + (1 if has_grand_total else 0)}" style="background:#9E1B32;color:#ffffff;padding:10px 12px;font-size:16px;font-weight:700;text-align:center;">PUMA {country} - Pending Orders</td></tr>
        {header_html}
        {''.join(body_rows)}
    </table>
    """

    # -- Order Summary box (right side) - all 6 dashboard metrics, same colors --
    metrics_dict = summary_df.set_index("Metric")["Count"].to_dict() if summary_df is not None and not summary_df.empty else {}
    summary_rows = [
        ("Overdue", metrics_dict.get("Overdue (SLA breached)", 0), "#FF0000", "#ffffff"),
        ("Handover Today", metrics_dict.get("Handover today (Today SLA)", 0), "#FFA500", "#ffffff"),
        ("Order status at NEW", metrics_dict.get("Order Status at New", 0), "#C00000", "#ffffff"),
        ("Within SLA", metrics_dict.get("Within SLA (Future)", 0), "#92D050", "#ffffff"),
        ("Not reflected in OMS", metrics_dict.get("Not reflecting in OM", 0), "#E2EFDA", "#000000"),
        ("Unpaid Orders", metrics_dict.get("Unpaid Orders", 0), "#D9D9D9", "#000000"),
    ]
    summary_rows_html = ""
    for label, count, bg_color, text_color in summary_rows:
        summary_rows_html += f"""
        <tr>
            <td style="padding:8px 12px;border:1px solid #bfbfbf;background:{bg_color};color:{text_color};font-weight:600;">{label}</td>
            <td style="padding:8px 12px;border:1px solid #bfbfbf;background:#ffffff;text-align:center;font-weight:700;">{count}</td>
        </tr>
        """
    summary_box_html = f"""
    <table style="border-collapse:collapse;width:100%;font-size:13px;">
        {summary_rows_html}
    </table>
    """

    combined_html = f"""
    <table style="width:100%;border-collapse:collapse;margin-top:15px;margin-bottom:15px;">
        <tr>
            <td style="vertical-align:top;padding-right:15px;">{pivot_table_html}</td>
            <td style="vertical-align:top;width:280px;">{summary_box_html}</td>
        </tr>
    </table>
    """
    return combined_html


def send_country_report_email(smtp_config, country, to_email, cc_email, excel_bytes, ref_date_str, pivot_df=None, summary_df=None, urgent_orders_df=None):
    """
    Send the country-specific Excel report to Ops/Seller, via either SMTP or
    the Brevo API depending on smtp_config["provider"] ("smtp" or "brevo").
    Body includes the Pending Orders pivot table + Order Summary box (built
    from pivot_df/summary_df). `urgent_orders_df` is accepted for backward
    compatibility but no longer used for the body content.
    """
    from datetime import datetime
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication

    if not to_email or "@" not in to_email:
        return False, "Recipient 'To' email address is required and must be valid."

    date_suffix = ref_date_str if ref_date_str else datetime.today().strftime('%d-%m-%Y')
    subject = f"PUMA - {country} Pending order Report on {date_suffix}"
    filename = f"Pending_order_report_-_PUMA_{country}_{date_suffix}.xlsx"

    pivot_summary_html = _build_country_pivot_email_html(country, pivot_df, summary_df, date_suffix)

    email_html = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                color: #333333;
                background-color: #ffffff;
                margin: 0;
                padding: 20px;
            }}
            .content {{
                max-width: 1000px;
            }}
            .highlight-note {{
                background-color: #1E88E5;
                color: #ffffff;
                font-weight: 600;
                padding: 10px 14px;
                border-radius: 4px;
                display: inline-block;
                margin-top: 10px;
            }}
            .footer {{
                font-size: 12px;
                color: #888888;
                border-top: 1px solid #e0e0e0;
                padding-top: 15px;
                margin-top: 30px;
            }}
        </style>
    </head>
    <body>
        <div class="content">
            <p>Hi Ops Team,</p>
            <p>find the attached Pending Orders Report. Kindly ensure all orders are processed and shipped on time to avoid any cancellations.</p>
            <p>Below are orders that require immediate attention:</p>

            {pivot_summary_html}

            <p><span class="highlight-note">Please prioritize shipping to avoid cancellations.</span></p>

            <p style="margin-top: 25px;">Thanks &amp; Regards,<br>
            <strong>Graas Team</strong></p>

            <div class="footer">
                This is an automated report. Please do not reply directly to this email.
            </div>
        </div>
    </body>
    </html>
    """

    provider = smtp_config.get("provider", "smtp")

    # -- Path 1: Brevo API (recommended when Office 365 SMTP AUTH is blocked) --
    if provider == "brevo":
        ok, msg = _send_via_brevo(
            api_key=smtp_config.get("api_key", ""),
            sender_email=smtp_config.get("sender_email", ""),
            sender_name="Graas Team",
            to_email=to_email,
            cc_email=cc_email,
            subject=subject,
            html_body=email_html,
            attachment_bytes=excel_bytes,
            attachment_filename=filename
        )
        if ok:
            return True, f"Report email for PUMA {country} sent successfully via Brevo!"
        return False, msg

    # -- Path 2: Google / Gmail API (OAuth2, "Sign in with Google") --
    if provider == "google":
        ok, msg = _send_via_gmail_api(
            client_id=smtp_config.get("google_client_id", ""),
            client_secret=smtp_config.get("google_client_secret", ""),
            refresh_token=smtp_config.get("google_refresh_token", ""),
            sender_email=smtp_config.get("sender_email", ""),
            to_email=to_email,
            cc_email=cc_email,
            subject=subject,
            html_body=email_html,
            attachment_bytes=excel_bytes,
            attachment_filename=filename
        )
        if ok:
            return True, f"Report email for PUMA {country} sent successfully via Gmail!"
        return False, msg

    # -- Path 3: SMTP (Office 365 / other) --
    user = smtp_config.get("user")
    # "From" is always the SMTP login/sender email - never a separately
    # entered address - per business requirement.
    sender_email = smtp_config.get("sender_email", user)

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = to_email

    recipients = [email.strip() for email in to_email.split(",") if email.strip()]
    if cc_email:
        msg["Cc"] = cc_email
        recipients += [email.strip() for email in cc_email.split(",") if email.strip()]

    msg["Subject"] = subject
    msg.attach(MIMEText(email_html, "html"))

    attachment = MIMEApplication(excel_bytes, _subtype="xlsx")
    attachment.add_header("Content-Disposition", "attachment", filename=filename)
    msg.attach(attachment)

    try:
        server = _smtp_connect_and_login(smtp_config)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        return True, f"Report email for PUMA {country} sent successfully!"
    except Exception as e:
        return False, f"Failed to send email: {str(e)}"


def send_discrepancies_to_slack_email(smtp_config, discrepancies_df, ref_date_str):
    """
    Send the status discrepancies Excel report to the Slack channel email integration address.
    """
    import smtplib
    import io
    from datetime import datetime
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication
    import pandas as pd
    import excel_formatter

    if discrepancies_df.empty:
        return True, "No discrepancies found today to send to Slack."

    slack_email = "order-related-issues-aaaausyacvowbr5tw6wbivuota@graas-talk.slack.com"
    
    host = smtp_config.get("host")
    port = int(smtp_config.get("port", 587))
    user = smtp_config.get("user")
    password = smtp_config.get("password")
    use_tls = smtp_config.get("use_tls", True)
    sender_email = smtp_config.get("sender_email", user)

    date_suffix = ref_date_str if ref_date_str else datetime.today().strftime('%d-%m-%Y')
    subject = f"Status Discrepancies Report - {date_suffix}"
    
    # Generate Excel in memory
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        discrepancies_df.to_excel(writer, sheet_name="Status Discrepancies", index=False)
        excel_formatter.format_data_sheet(writer.sheets["Status Discrepancies"], discrepancies_df)
    excel_bytes = excel_buffer.getvalue()

    total_discrepancies = len(discrepancies_df)
    
    # HTML body for Slack integration
    email_html = f"""
    <html>
    <body>
        <h3>Daily Status Discrepancies Report ({date_suffix})</h3>
        <p>Total Discrepancies Found: <strong>{total_discrepancies}</strong></p>
        <p>Please find attached the detailed Excel status discrepancies sheet.</p>
    </body>
    </html>
    """

    msg = MIMEMultipart()
    msg["From"] = f"Operations Team <{sender_email}>"
    msg["To"] = slack_email
    msg["Subject"] = subject
    msg.attach(MIMEText(email_html, "html"))

    # Attach Excel
    attachment = MIMEApplication(excel_bytes, _subtype="xlsx")
    attachment.add_header("Content-Disposition", "attachment", filename=f"Status_Discrepancies_{date_suffix}.xlsx")
    msg.attach(attachment)

    try:
        server = _smtp_connect_and_login(smtp_config)
        server.sendmail(sender_email, [slack_email], msg.as_string())
        server.quit()
        return True, "Discrepancies report shared successfully to Slack group!"
    except Exception as e:
        return False, f"Failed to share to Slack: {str(e)}"


def send_discrepancies_to_slack_webhook(webhook_url, discrepancies_df, ref_date_str, country=None, app_download_url=None):
    """
    Post a Status Discrepancies summary directly to a Slack channel via an
    Incoming Webhook (https://api.slack.com/messaging/webhooks) - a simple
    HTTPS POST, no SMTP/email involved.

    NOTE: Slack Incoming Webhooks cannot carry file attachments (that needs
    the separate Slack Files API with a bot token) - so this posts a
    formatted text summary (counts + breakdown by Validation Result) with a
    reminder to check the full report in the app, rather than an attached
    Excel file.
    """
    try:
        if not str(webhook_url or "").strip():
            return False, "Slack Webhook URL is required."
        if not webhook_url.startswith("https://hooks.slack.com/"):
            return False, "That doesn't look like a valid Slack Incoming Webhook URL (should start with https://hooks.slack.com/)."

        from datetime import datetime
        date_suffix = ref_date_str if ref_date_str else datetime.today().strftime('%d-%m-%Y')
        title = f"📋 Status Discrepancies{f' - {country}' if country else ''} ({date_suffix})"

        if discrepancies_df is None or discrepancies_df.empty:
            blocks = [
                {"type": "header", "text": {"type": "plain_text", "text": title}},
                {"type": "section", "text": {"type": "mrkdwn", "text": "✅ No status discrepancies found today."}},
            ]
        else:
            total = len(discrepancies_df)
            breakdown_lines = []
            if "Validation Result" in discrepancies_df.columns:
                counts = discrepancies_df["Validation Result"].value_counts()
                for label, cnt in counts.items():
                    breakdown_lines.append(f"• *{label}*: {cnt}")
            breakdown_text = "\n".join(breakdown_lines) if breakdown_lines else "_(no breakdown available)_"

            blocks = [
                {"type": "header", "text": {"type": "plain_text", "text": title}},
                {"type": "section", "text": {"type": "mrkdwn", "text": f"*Total Discrepancies:* {total}"}},
                {"type": "divider"},
                {"type": "section", "text": {"type": "mrkdwn", "text": breakdown_text}},
            ]
            if app_download_url:
                blocks.append({"type": "divider"})
                blocks.append({"type": "section", "text": {"type": "mrkdwn", "text": f"<{app_download_url}|Open the full report in the app>"}})

        payload = {
            "text": title,  # fallback text for notifications
            "blocks": blocks
        }

        resp = requests.post(webhook_url, json=payload, timeout=15)
        if resp.status_code == 200 and resp.text.strip().lower() == "ok":
            return True, "Discrepancies summary posted to Slack successfully!"
        return False, f"Slack webhook failed: {resp.status_code} - {resp.text[:300]}"
    except Exception as e:
        return False, f"Failed to post to Slack: {str(e)}"


# ============================================================================
# Slack Bot Token - actual Excel FILE upload
# ============================================================================
# Incoming Webhooks (above) can only post text/Block Kit messages - they
# cannot carry file attachments. To post the real Excel file into a Slack
# channel (not just a text summary), Slack requires a Bot Token with the
# `files:write` scope and their current 3-step upload flow (the old single-
# call `files.upload` endpoint was deprecated).

def test_slack_bot_token(bot_token):
    """Verify a Slack Bot Token is valid by calling auth.test."""
    try:
        if not str(bot_token or "").strip():
            return False, "Slack Bot Token is required."
        resp = requests.post(
            "https://slack.com/api/auth.test",
            headers={"Authorization": f"Bearer {bot_token}"},
            timeout=15
        )
        data = resp.json()
        if data.get("ok"):
            team = data.get("team", "")
            return True, f"Successfully connected to Slack!{f' (Workspace: {team})' if team else ''}"
        return False, f"Slack token check failed: {data.get('error', 'unknown error')}"
    except Exception as e:
        return False, f"Slack connection failed: {str(e)}"


def send_excel_to_slack_channel(bot_token, channel_id, excel_bytes, filename, initial_comment=None):
    """
    Upload and share an Excel file directly into a Slack channel using a Bot
    Token, via Slack's current 3-step external upload flow:
      1. files.getUploadURLExternal - get a one-time upload URL
      2. POST the raw file bytes to that URL
      3. files.completeUploadExternal - finalize and share it into the channel
    """
    try:
        if not str(bot_token or "").strip():
            return False, "Slack Bot Token is required."
        if not str(channel_id or "").strip():
            return False, "Slack Channel ID is required."

        headers = {"Authorization": f"Bearer {bot_token}"}

        # Step 1: get an upload URL
        step1 = requests.post(
            "https://slack.com/api/files.getUploadURLExternal",
            headers=headers,
            data={"filename": filename, "length": len(excel_bytes)},
            timeout=20
        )
        step1_data = step1.json()
        if not step1_data.get("ok"):
            return False, f"Slack upload-URL request failed: {step1_data.get('error', 'unknown error')}"
        upload_url = step1_data["upload_url"]
        file_id = step1_data["file_id"]

        # Step 2: upload the raw bytes
        step2 = requests.post(upload_url, files={"file": (filename, excel_bytes)}, timeout=30)
        if step2.status_code != 200:
            return False, f"Slack file upload failed: {step2.status_code} - {step2.text[:200]}"

        # Step 3: complete the upload and share it into the channel
        step3 = requests.post(
            "https://slack.com/api/files.completeUploadExternal",
            headers={**headers, "Content-Type": "application/json"},
            json={
                "files": [{"id": file_id, "title": filename}],
                "channel_id": channel_id,
                "initial_comment": initial_comment or ""
            },
            timeout=20
        )
        step3_data = step3.json()
        if step3_data.get("ok"):
            return True, "Excel file shared to Slack successfully!"
        return False, f"Slack upload-completion failed: {step3_data.get('error', 'unknown error')}"
    except Exception as e:
        return False, f"Failed to share file to Slack: {str(e)}"
