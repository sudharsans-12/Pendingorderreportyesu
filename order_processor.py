# -*- coding: utf-8 -*-
# VERSION: v6 - MCP Integration Update
import pandas as pd
import numpy as np
import io
import re
import urllib.request
import urllib.error
import json
import os
import requests

# Collects non-fatal warnings from file loading (e.g. malformed CSV rows that
# had to be skipped) so the app can surface them to the user instead of
# silently dropping data or crashing the whole run. Cleared at the start of
# each process_and_validate_orders() call.
FILE_LOAD_WARNINGS = []
from datetime import datetime

def get_google_sheet_download_url(url):
    pattern = r"/spreadsheets/d/([a-zA-Z0-9-_]+)"
    match = re.search(pattern, url)
    if match:
        spreadsheet_id = match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
    return None

def download_google_sheet(url):
    download_url = get_google_sheet_download_url(url)
    if not download_url:
        download_url = url
    try:
        req = urllib.request.Request(
            download_url,
            headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        )
        with urllib.request.urlopen(req) as response:
            data = response.read()
        return io.BytesIO(data)
    except urllib.error.HTTPError as e:
        if e.code in (401, 403):
            raise ValueError(
                "Failed to download Google Sheet: Access Denied (HTTP 401/403 Unauthorized). "
                "Please ensure that the sharing settings of your Google Sheet are set to "
                "'Anyone with the link can view' or 'Anyone with the link can edit' (restricted sheets require auth, which is not supported)."
            )
        raise ValueError(f"Failed to download Google Sheet: HTTP Error {e.code} - {e.reason}")
    except Exception as e:
        raise ValueError(f"Failed to download Google Sheet from URL: {str(e)}")

def compute_sla_status(sla_date, ref_date):
    if _is_blank(sla_date) or _is_blank(ref_date):
        return ""
    s_dt = extract_date(sla_date)
    r_dt = extract_date(ref_date)
    if not (len(s_dt) == 10 and len(r_dt) == 10):
        return ""
    if s_dt < r_dt:
        return "Breached"
    elif s_dt == r_dt:
        return "Today"
    else:
        return "Future"


def _clean_str(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip()

def _normalize_status_val(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip().lower().replace("_", " ").replace("-", " ")

def _clean_order_id(val):
    """
    Standardizes the order number as a string. Handles potential float scientific 
    notations introduced by Excel for numbers larger than 10 digits.
    """
    s = _clean_str(val)
    if not s or s.lower() in ("nan", "none", "nat"):
        return ""
    
    # Check if the string was converted to scientific notation (e.g. 2.3E+14)
    if "e" in s.lower() and "+" in s:
        try:
            f_val = float(s)
            s = str(int(f_val))
        except Exception:
            pass
            
    # Remove float decimal point representation
    if s.endswith(".0"):
        s = s[:-2]
        
    return s.upper()

def parse_country_and_channel(nickname):
    nick = str(nickname).strip().upper()
    # Extract Country
    if nick.endswith("SG") or "-SG" in nick or "_SG" in nick:
        country = "SG"
    elif nick.endswith("MY") or "-MY" in nick or "_MY" in nick:
        country = "MY"
    elif nick.endswith("PH") or "-PH" in nick or "_PH" in nick:
        country = "PH"
    else:
        # Default fallback checks
        if "SG" in nick:
            country = "SG"
        elif "MY" in nick:
            country = "MY"
        elif "PH" in nick:
            country = "PH"
        else:
            country = "UNKNOWN"
            
    # Extract Channel dynamically
    # Clean PUMA_ or PUMA- prefix
    clean_nick = re.sub(r'^PUMA[_-]', '', nick, flags=re.IGNORECASE)
    clean_nick = re.sub(r'^PUMA', '', clean_nick, flags=re.IGNORECASE)
    
    # Split by delimiters like - or _ or spaces
    parts = re.split(r'[_-]', clean_nick)
    first_part = parts[0].strip().title() if parts else "Other"
    
    # Standard mapping for common names to look nice
    mapping = {
        "Shopee": "Shopee",
        "Lazada": "Lazada",
        "Zalora": "Zalora",
        "Tiktok": "TikTok",
        "Salesforce": "Salesforce",
        "Shopify": "Shopify",
        "Decathlon": "Decathlon",
        "Amazon": "Amazon"
    }
    
    channel = mapping.get(first_part, first_part)
    if not channel or channel.upper() in ("SG", "MY", "PH", "UNKNOWN"):
        channel = "Other"
        
    return country, channel

def _find_column(df, candidates):
    """Find a column in df that matches any of the candidate names case-insensitively and ignoring underscores/spaces."""
    cols = list(df.columns)
    for cand in candidates:
        cand_norm = cand.lower().replace(" ", "").replace("_", "").replace("-", "")
        for col in cols:
            col_norm = str(col).lower().replace(" ", "").replace("_", "").replace("-", "")
            if col_norm == cand_norm:
                return col
    return None

def _build_seller_email_map(df_contacts):
    """
    Build a {seller_name (cleaned) -> email} lookup from an uploaded contacts
    file. Expects a seller/store name column (e.g. "Seller Name", "Store",
    "Nickname") and an email column (e.g. "Email", "Seller Email"). Returns an
    empty dict if no contacts file was supplied or the required columns
    aren't found.
    """
    email_map = {}
    if df_contacts is None or df_contacts.empty:
        return email_map

    name_col = _find_column(df_contacts, [
        "Seller Name", "Seller", "Store Name", "Store", "Nickname", "Shop Name", "Shop"
    ])
    email_col = _find_column(df_contacts, [
        "Email", "Seller Email", "Email Address", "Contact Email", "Recipient Email"
    ])
    if not name_col or not email_col:
        return email_map

    for _, row in df_contacts.iterrows():
        seller_name = _clean_str(row.get(name_col, ""))
        email = _clean_str(row.get(email_col, ""))
        if seller_name and email and "@" in email:
            email_map[seller_name.strip().lower()] = email

    return email_map

def _is_blank(val):
    """Check if a value is null, empty string, or nan/nat strings produced by Excel loading."""
    s = str(val).strip().replace('\r', '').replace('\n', '')
    return not s or s.lower() in ("nan", "none", "nat", "null", "undefined", "nat", "#n/a")

def normalize_ean(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    if digits:
        if len(digits) >= 13:
            return digits[-13:]
        else:
            return digits.zfill(13)
    return s

def extract_date(val):
    """Extract a YYYY-MM-DD date from various timestamp formats. Returns #N/A if blank."""
    if _is_blank(val):
        return "#N/A"
    s = str(val).strip()
    if len(s) >= 10:
        if '-' in s:
            parts = s.split(' ')[0].split('-')
            if len(parts) == 3:
                if len(parts[0]) == 4:
                    return f"{parts[0]}-{parts[1]}-{parts[2]}"
                else:
                    return f"{parts[2]}-{parts[1]}-{parts[0]}"
        elif '/' in s:
            parts = s.split(' ')[0].split('/')
            if len(parts) == 3:
                if len(parts[2]) == 4:
                    return f"{parts[2]}-{parts[1]}-{parts[0]}"
                else:
                    return f"{parts[0]}-{parts[1]}-{parts[2]}"
    return s if s else "#N/A"

def load_file_safely(file):
    """Load uploaded file object (CSV or Excel) into a DataFrame. Scans sheets if Excel."""
    if file is None:
        return pd.DataFrame()
    if isinstance(file, pd.DataFrame):
        return file
    
    # If a string path is passed, open it and read as BytesIO
    if isinstance(file, str):
        import os
        if not os.path.exists(file):
            return pd.DataFrame()
        with open(file, 'rb') as f:
            file_bytes = f.read()
        filename = os.path.basename(file)
        file = io.BytesIO(file_bytes)
        file.name = filename

    try:
        file.seek(0)
    except Exception:
        pass
        
    name = getattr(file, "name", "google_sheet.xlsx").lower()
    
    try:
        if name.endswith(".csv"):
            # Detect delimiter safely by reading the first line/chunk
            delim = ','
            try:
                file.seek(0)
                first_bytes = file.read(2048)
                first_line = first_bytes.decode('utf-8-sig', errors='ignore').split('\n')[0]
                comma_count = first_line.count(',')
                semicolon_count = first_line.count(';')
                tab_count = first_line.count('\t')
                
                max_count = comma_count
                if semicolon_count > max_count:
                    delim = ';'
                    max_count = semicolon_count
                if tab_count > max_count:
                    delim = '\t'
                    max_count = tab_count
            except Exception:
                pass
            finally:
                try:
                    file.seek(0)
                except Exception:
                    pass

            try:
                file.seek(0)
                df = pd.read_csv(file, sep=delim, dtype=str)
                if not df.empty:
                    return df.dropna(how="all").reset_index(drop=True)
            except Exception:
                pass

            # FIX: previously, if the standard read failed (e.g. a row has
            # more fields than the header - typically an unescaped comma
            # inside a text field like an address or product name), the
            # fallback below just retried the exact same strict parsing and
            # crashed the ENTIRE run, blocking every other file too. Instead,
            # try again allowing malformed rows to be skipped (not silently -
            # every skipped row is recorded in FILE_LOAD_WARNINGS so it can be
            # surfaced to the user) rather than failing the whole file.
            try:
                file.seek(0)
                skipped_lines = []

                def _capture_bad_line(bad_line):
                    skipped_lines.append(bad_line)
                    return None  # tells pandas to drop this line and continue

                df = pd.read_csv(file, sep=delim, dtype=str, engine="python", on_bad_lines=_capture_bad_line)
                if skipped_lines:
                    FILE_LOAD_WARNINGS.append(
                        f"'{name}': skipped {len(skipped_lines)} malformed row(s) that had an "
                        f"unexpected number of fields (likely an unescaped comma in a text field "
                        f"like an address or product name). These rows were NOT included in the "
                        f"report - please check the source file for orders that may be missing."
                    )
                if not df.empty:
                    return df.dropna(how="all").reset_index(drop=True)
            except Exception:
                pass
            
            # Ultimate fallback: parse with Python's built-in csv module
            # directly, bypassing pandas' C/Python tokenizer entirely. This
            # cannot raise a "tokenizing data" error no matter how ragged the
            # rows are - rows with too many fields are skipped (recorded as a
            # warning), rows with too few are padded with blanks.
            file.seek(0)
            raw = file.read()
            text = raw.decode("utf-8", errors="replace") if isinstance(raw, bytes) else raw
            import csv as _csv_module
            reader = _csv_module.reader(io.StringIO(text), delimiter=delim)
            all_rows = list(reader)
            if not all_rows:
                return pd.DataFrame()

            header = all_rows[0]
            n_cols = len(header)
            clean_rows = []
            skipped_count = 0
            for r in all_rows[1:]:
                if len(r) == n_cols:
                    clean_rows.append(r)
                elif len(r) > n_cols:
                    skipped_count += 1
                else:
                    clean_rows.append(r + [""] * (n_cols - len(r)))

            if skipped_count:
                FILE_LOAD_WARNINGS.append(
                    f"'{name}': skipped {skipped_count} malformed row(s) that had more fields than "
                    f"the header (likely an unescaped comma in a text field like an address or "
                    f"product name). These rows were NOT included in the report - please check the "
                    f"source file for orders that may be missing."
                )
            df = pd.DataFrame(clean_rows, columns=header, dtype=str) if clean_rows else pd.DataFrame(columns=header)
            return df.dropna(how="all").reset_index(drop=True)
        else:
            # Excel - Scan all sheets to find the first non-empty one
            try:
                file.seek(0)
                xl = pd.ExcelFile(file)
            except Exception:
                file.seek(0)
                raw = file.read()
                xl = pd.ExcelFile(io.BytesIO(raw))
                
            # Optimized fallback: try first sheet first (99% of cases contain data in sheet 1)
            if xl.sheet_names:
                first_sheet = xl.sheet_names[0]
                try:
                    df = xl.parse(first_sheet, dtype=str)
                    if df is not None and not df.empty:
                        df_clean = df.dropna(how="all").reset_index(drop=True)
                        if not df_clean.empty:
                            return df_clean
                except Exception:
                    pass
                    
            # Loop fallback for secondary sheets
            for sheet in xl.sheet_names[1:]:
                try:
                    df = xl.parse(sheet, dtype=str)
                    if df is not None and not df.empty:
                        df_clean = df.dropna(how="all").reset_index(drop=True)
                        if not df_clean.empty:
                            return df_clean
                except Exception:
                    continue
            
            # Final fallback
            if xl.sheet_names:
                return xl.parse(xl.sheet_names[0], dtype=str)
            return pd.DataFrame()
    except Exception as e:
        raise ValueError(f"Failed to read file {getattr(file, 'name', 'Google Sheet')}: {str(e)}")

def fetch_pending_from_mcp():
    """
    Fetch pending orders list directly from PUMA MCP Database.
    """
    token_file = r"C:\Users\Yesuraja\.gemini\antigravity\brain\abf6c61b-3147-45f7-90e4-f03458ddd1ae\scratch\token_data.json"
    if not os.path.exists(token_file):
        raise ValueError("MCP Token not found. Please authenticate with the MCP server in the sidebar first.")
        
    with open(token_file, "r") as f:
        token_data = json.load(f)
        
    access_token = token_data.get("access_token")
    if not access_token:
        raise ValueError("Access token not found in token_data.json. Please re-authenticate.")
        
    mcp_url = "https://mcp.graas.ai/mcp/GED"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }
    
    # 1. Initialize session
    init_payload = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "initialize",
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {
                "name": "Antigravity-Client",
                "version": "1.0.0"
            }
        }
    }
    
    try:
        r_init = requests.post(mcp_url, json=init_payload, headers=headers, timeout=15)
        if r_init.status_code == 401:
            raise PermissionError("Access token expired or unauthorized. Please re-authenticate in the sidebar.")
    except requests.exceptions.RequestException as e:
        raise ValueError(f"Failed to connect to MCP server: {str(e)}")
        
    query = """
    SELECT 
        oi.ORDER_ID AS "orderID",
        oi.SELLER_ID AS "merchantID",
        oi.ORDER_CREATED_REPORT_TS AS "timeOrderCreated",
        oi.ORDER_ID AS "orderNumber",
        oi.ORDER_STATUS AS "orderStatus",
        oi.ORDER_ITEM_STATUS AS "orderItems.orderStatus",
        oi.PAYMENT_STATUS AS "paymentStatus",
        o.PAYMENT_METHOD AS "paymentMethods",
        oi.SELLER_SKU AS "orderItems.customSKU",
        NULL AS "shippingDeadLine",
        o.SHIPPING_METHOD AS "courierName",
        NULL AS "airwaybill",
        oi.SHIPPING_STATUS AS "omsStatus",
        oi.CHANNEL_NAME AS "storeName"
    FROM ORDER_ITEMS_METRICS oi
    LEFT JOIN ORDER_METRICS o ON oi.ORDER_ID = o.ORDER_ID
    WHERE oi.ORDER_STATUS IN ('UNPAID', 'INITIATED', 'PROCESSING', 'ACCEPTED', 'AWAITING_COLLECTION', 'AWAITING_SHIPMENT')
    """
    
    call_payload = {
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "query_data",
            "arguments": {
                "sql_query": query,
                "question": "Fetch pending orders list for seller"
            }
        }
    }
    
    try:
        r = requests.post(mcp_url, json=call_payload, headers=headers, timeout=30)
        if r.status_code == 401:
            raise PermissionError("Access token expired or unauthorized. Please re-authenticate in the sidebar.")
            
        if r.status_code != 200:
            raise ValueError(f"MCP Server returned status code {r.status_code}: {r.text}")
            
        res = r.json()
        if "error" in res:
            raise ValueError(f"MCP Server Error: {res['error'].get('message', 'Unknown error')}")
            
        result_data = res.get("result", {})
        structured = result_data.get("structuredContent", {})
        
        columns = []
        rows = []
        
        if structured:
            columns = structured.get("columns", [])
            rows = structured.get("rows", [])
        else:
            # Fallback to content[0].text
            content_list = result_data.get("content", [])
            if content_list:
                try:
                    parsed = json.loads(content_list[0].get("text", "{}"))
                    if isinstance(parsed, dict) and "rows" in parsed:
                        columns = parsed.get("columns", [])
                        rows = parsed.get("rows", [])
                except Exception:
                    pass
                    
        if not rows:
            return pd.DataFrame()
            
        df = pd.DataFrame(rows, columns=columns)
        
        # Construct mapped DataFrame matching expected structure of Pending Order Report
        df_pending = pd.DataFrame()
        df_pending["Order ID"] = df["orderID"].astype(str)
        # Clean channel name to PUMA-like or simple format
        df_pending["Store Name"] = df["storeName"].fillna(df["merchantID"]).astype(str)
        # Empty string for SLA (will be enriched from TC Report)
        df_pending["SLA"] = ""
        # Format date
        df_pending["Order Date"] = df["timeOrderCreated"].astype(str).apply(lambda x: x.split("T")[0] if "T" in x else x)
        
        return df_pending
    except Exception as e:
        raise ValueError(f"Failed to fetch pending orders from DB: {str(e)}")

def run_gsheet_oms_validation(df_pending, df_oms, df_contacts=None):
    # Find ID and nickname/store column in Pending Order Report
    pend_id_col = _find_column(df_pending, ["order_id", "order_number", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    pend_sla_col = _find_column(df_pending, ["mp_sla_date", "SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date", "mp_sla_date_updated"])
    pend_store_col = _find_column(df_pending, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
    
    is_tiktok_ph = lambda x: str(x).strip().lower().replace(" ", "").replace("-", "").replace("_", "") == "tiktokph"
    if pend_store_col and pend_store_col in df_pending.columns:
        df_pending = df_pending[~df_pending[pend_store_col].apply(is_tiktok_ph)].copy()
        
    oms_store_col = _find_column(df_oms, ["store", "nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
    if oms_store_col and oms_store_col in df_oms.columns:
        df_oms = df_oms[~df_oms[oms_store_col].apply(is_tiktok_ph)].copy()

    if not pend_id_col:
        raise KeyError(f"Could not find 'Order ID' column in Pending Order Report. Available: {list(df_pending.columns)}")
    df_pending[pend_id_col] = df_pending[pend_id_col].apply(_clean_order_id)
    
    target_sla_col = pend_sla_col if pend_sla_col else "SLA"
    if target_sla_col not in df_pending.columns:
        df_pending[target_sla_col] = ""
        
    target_store_col = pend_store_col if pend_store_col else "Store Name"
    if target_store_col not in df_pending.columns:
        df_pending[target_store_col] = "Default Store"
        
    # Clean store column to remove PUMA_ prefix case-insensitively
    df_pending[target_store_col] = df_pending[target_store_col].apply(
        lambda x: re.sub(r'puma_', '', str(x).strip(), flags=re.IGNORECASE)
    )

    # OMS Report columns (Sales Order file)
    oms_id_col = _find_column(df_oms, ["order_no", "order_id", "order_number", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    oms_status_col = _find_column(df_oms, ["order_status", "OMS Status", "Order Status", "Status", "OMS_Status"])
    
    if not oms_id_col:
        raise KeyError(f"Could not find 'Order ID' column in OMS Report. Available: {list(df_oms.columns)}")
    df_oms[oms_id_col] = df_oms[oms_id_col].apply(_clean_order_id)

    oms_status_map = {}
    oms_order_status_fallback_map = {}
    oms_ean_col = _find_column(df_oms, ["ean", "EAN", "Ean", "item_sku", "SKU"])
    
    if oms_id_col:
        for _, row in df_oms.iterrows():
            oid = _clean_order_id(row[oms_id_col])
            if not oid:
                continue
            ean = normalize_ean(row.get(oms_ean_col)) if oms_ean_col else ""
            key = oid + ean
            stat_val = _clean_str(row.get(oms_status_col, "")) if oms_status_col else ""
            oms_status_map[key] = stat_val
            oms_order_status_fallback_map[oid] = stat_val

    df_pending["OMS Order Status"] = ""
    df_pending["Final Remarks"] = ""
    df_pending["Correct Order Number"] = df_pending[pend_id_col]
    df_pending["SLA Source"] = "Pending Report"
    
    if "sla_status" not in df_pending.columns:
        df_pending["sla_status"] = ""
        
    # Determine reference date
    ref_date = datetime.today().strftime('%Y-%m-%d')
    temp_ref_date = None
    if "sla_status" in df_pending.columns:
        today_rows = df_pending[df_pending["sla_status"].astype(str).str.strip().str.lower() == "today"]
        if not today_rows.empty:
            for val in today_rows[target_sla_col]:
                if not _is_blank(val):
                    p_dt = extract_date(val)
                    if len(p_dt) == 10:
                        temp_ref_date = p_dt
                        break
    if not temp_ref_date and target_sla_col in df_pending.columns:
        for val in df_pending[target_sla_col]:
            if not _is_blank(val):
                p_dt = extract_date(val)
                if len(p_dt) == 10:
                    temp_ref_date = p_dt
                    break
    if temp_ref_date:
        ref_date = temp_ref_date

    # Compute sla status
    for idx, row in df_pending.iterrows():
        sla_val = row[target_sla_col]
        sla_status_val = row.get("sla_status", "")
        if _is_blank(sla_status_val):
            curr_sla = df_pending.at[idx, target_sla_col]
            calculated_status = compute_sla_status(curr_sla, ref_date)
            if calculated_status:
                df_pending.at[idx, "sla_status"] = calculated_status

    discrepancies = []
    pushed_count = 0
    not_pushed_count = 0
    unpaid_count = 0
    
    # Check what status column GSheet has for orders
    pend_status_col = _find_column(df_pending, ["order_status", "status", "Item Status", "orderItems.orderStatus", "TC Status"])
    pend_sku_col = _find_column(df_pending, ["seller_sku", "sellerSku", "Seller SKU", "SellerSKU", "custom_sku", "customSku", "sku", "item_sku", "SKU", "orderItems.customSKU"])
    
    # Check if there is an order number column and/or oms_pushed column in df_pending
    pend_num_col = _find_column(df_pending, ["order_number", "Order Number", "Order_No", "Correct Order Number"])
    oms_pushed_col = _find_column(df_pending, ["oms_pushed", "omsPushed", "OMS Pushed", "OMS_pushed"])
    if not oms_pushed_col:
        df_pending["oms_pushed"] = ""
        oms_pushed_col = "oms_pushed"
    
    # Resolve payment columns if available
    pend_pay_status_col = _find_column(df_pending, ["payment_status", "Payment Status", "Payment_Status", "PaymentStatus", "Payment"])
    pend_pay_method_col = _find_column(df_pending, ["payment_methods", "Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"])

    for idx, row in df_pending.iterrows():
        order_id = row[pend_id_col]
        order_id_str = str(order_id).strip()
        store_val = _clean_str(row.get(target_store_col, ""))
        sku_val = normalize_ean(row.get(pend_sku_col)) if pend_sku_col else ""
        key = order_id_str + sku_val
        
        pay_stat_val = _clean_str(row.get(pend_pay_status_col, "")) if pend_pay_status_col else ""
        pay_meth_val = _clean_str(row.get(pend_pay_method_col, "")) if pend_pay_method_col else ""
        
        # Resolve order number for Zalora matching
        ord_num = _clean_order_id(row.get(pend_num_col)) if pend_num_col else ""
        match_id = ord_num if ord_num else order_id_str
        
        is_gsheet_cancelled = False
        if pend_status_col:
            pend_stat_raw = _clean_str(row.get(pend_status_col, ""))
            pend_norm = _normalize_status_val(pend_stat_raw)
            if "cancel" in pend_norm:
                is_gsheet_cancelled = True
                
        is_in_oms = False
        oms_stat = ""
        
        # Match using order number + SKU, order ID + SKU, or fallback order number/ID
        if match_id and (match_id + sku_val) in oms_status_map:
            is_in_oms = True
            oms_stat = oms_status_map[match_id + sku_val]
        elif key in oms_status_map:
            is_in_oms = True
            oms_stat = oms_status_map[key]
        elif match_id and match_id in oms_order_status_fallback_map:
            is_in_oms = True
            oms_stat = oms_order_status_fallback_map[match_id]
        elif order_id_str and order_id_str in oms_order_status_fallback_map:
            is_in_oms = True
            oms_stat = oms_order_status_fallback_map[order_id_str]
            
        if is_in_oms:
            df_pending.at[idx, "OMS Order Status"] = oms_stat
            df_pending.at[idx, "Final Remarks"] = "Successfully Pushed to OMS"
            pushed_count += 1
            if oms_pushed_col:
                df_pending.at[idx, oms_pushed_col] = "Pushed"
            
            if pend_status_col:
                pend_stat_raw = _clean_str(row.get(pend_status_col, ""))
                if pend_stat_raw:
                    pend_norm = _normalize_status_val(pend_stat_raw)
                    oms_norm = _normalize_status_val(oms_stat)
                    
                    # Rule 1: Cancelled status check
                    is_tc_cancelled = ("cancel" in pend_norm)
                    is_oms_cancelled = ("cancel" in oms_norm)
                    if (is_tc_cancelled or is_oms_cancelled) and (is_tc_cancelled != is_oms_cancelled):
                        is_oms_returned = ("return" in oms_norm)
                        if not (is_tc_cancelled and is_oms_returned):
                            discrepancies.append({
                                "Order ID": order_id_str,
                                "Nickname": store_val,
                                "SKU": sku_val,
                                "Payment Status": pay_stat_val,
                                "Payment Method": pay_meth_val,
                                "Validation Result": "Cancelled Status Mismatch",
                                "TC Order Status": pend_stat_raw,
                                "TC Item Status": pend_stat_raw,
                                "OMS Order Status": oms_stat,
                                "OMS Line Status": oms_stat,
                                "Details": f"Cancelled status mismatch: Pending is '{pend_stat_raw}', OMS is '{oms_stat}'."
                            })

                    # Rule 2: OMS Packed and TC/GSheet Status if New can highlight
                    if "packed" in oms_norm and "new" in pend_norm:
                        discrepancies.append({
                            "Order ID": order_id_str,
                            "Nickname": store_val,
                            "SKU": sku_val,
                            "Payment Status": pay_stat_val,
                            "Payment Method": pay_meth_val,
                            "Validation Result": "OMS Packed but TC New",
                            "TC Order Status": pend_stat_raw,
                            "TC Item Status": pend_stat_raw,
                            "OMS Order Status": oms_stat,
                            "OMS Line Status": oms_stat,
                            "Details": f"OMS status is Packed, but Pending status is '{pend_stat_raw}' (should be READY TO SHIP or ACCEPTED/PICKED)."
                        })

                    # Rule 3: OMS Status is Shipped and TC/GSheet status is NEW, READY TO SHIP, ACCEPTED/PICKED, Cancelled
                    if "shipped" in oms_norm:
                        is_tc_invalid = False
                        for val in ["new", "ready to ship", "accepted", "picked", "cancel"]:
                            if val in pend_norm:
                                is_tc_invalid = True
                                break
                        if is_tc_invalid:
                            discrepancies.append({
                                "Order ID": order_id_str,
                                "Nickname": store_val,
                                "SKU": sku_val,
                                "Payment Status": pay_stat_val,
                                "Payment Method": pay_meth_val,
                                "Validation Result": "OMS Shipped but TC Status Invalid",
                                "TC Order Status": pend_stat_raw,
                                "TC Item Status": pend_stat_raw,
                                "OMS Order Status": oms_stat,
                                "OMS Line Status": oms_stat,
                                "Details": f"OMS status is Shipped, but Pending status is '{pend_stat_raw}'."
                            })

                    # Rule 4: TC/GSheet is Delivered and OMS is not Delivered
                    if "delivered" in pend_norm and "delivered" not in oms_norm:
                        discrepancies.append({
                            "Order ID": order_id_str,
                            "Nickname": store_val,
                            "SKU": sku_val,
                            "Payment Status": pay_stat_val,
                            "Payment Method": pay_meth_val,
                            "Validation Result": "Need to push Delivered Status",
                            "TC Order Status": pend_stat_raw,
                            "TC Item Status": pend_stat_raw,
                            "OMS Order Status": oms_stat,
                            "OMS Line Status": oms_stat,
                            "Details": f"Pending status is Delivered, but OMS status is '{oms_stat}'."
                        })

                    # Rule 5: TC/GSheet is Returned and OMS is not Returned
                    tc_is_return = ("returned" in pend_norm or "return" in pend_norm)
                    oms_is_delivered = ("delivered" in oms_norm)
                    if tc_is_return and oms_is_delivered:
                        pass # Ignored per rule: Returned Requested/Accepted vs OMS Delivered is not a status mismatch
                    elif tc_is_return and not ("returned" in oms_norm or "return" in oms_norm):
                        discrepancies.append({
                            "Order ID": order_id_str,
                            "Nickname": store_val,
                            "SKU": sku_val,
                            "Payment Status": pay_stat_val,
                            "Payment Method": pay_meth_val,
                            "Validation Result": "TC Returned but OMS not Returned",
                            "TC Order Status": pend_stat_raw,
                            "TC Item Status": pend_stat_raw,
                            "OMS Order Status": oms_stat,
                            "OMS Line Status": oms_stat,
                            "Details": f"Pending status is Returned, but OMS status is '{oms_stat}'."
                        })

                    # Rule 6: TC/GSheet Delivery Failed and OMS is not Returned
                    if "failed" in pend_norm and not ("returned" in oms_norm or "return" in oms_norm):
                        discrepancies.append({
                            "Order ID": order_id_str,
                            "Nickname": store_val,
                            "SKU": sku_val,
                            "Payment Status": pay_stat_val,
                            "Payment Method": pay_meth_val,
                            "Validation Result": "TC Delivery Failed but OMS not Returned",
                            "TC Order Status": pend_stat_raw,
                            "TC Item Status": pend_stat_raw,
                            "OMS Order Status": oms_stat,
                            "OMS Line Status": oms_stat,
                            "Details": f"Pending status is Delivery Failed, but OMS status is '{oms_stat}'."
                        })
        else:
            # Check if this order is an Unpaid Order (payment status is NOT_INITIATED, etc. and method is not COD)
            is_unpaid = False
            if pend_pay_status_col:
                pay_stat = _clean_str(row.get(pend_pay_status_col, ""))
                pay_meth = _clean_str(row.get(pend_pay_method_col, "")) if pend_pay_method_col else ""
                is_cod = any(term in pay_meth.lower() for term in ["cod", "cash on delivery", "cashondelivery"])
                is_pending = (pay_stat.lower() in ("pending", "unpaid", "awaiting", "not_initiated", "not_initiate", "not initiated"))
                if is_pending and not is_cod:
                    is_unpaid = True
            
            if is_unpaid:
                df_pending.at[idx, "OMS Order Status"] = "Not in OMS - Unpaid orders"
                df_pending.at[idx, "Final Remarks"] = "Not Pushed to OMS - Unpaid orders"
                if oms_pushed_col:
                    df_pending.at[idx, oms_pushed_col] = "Not pushed - Unpaid Orders"
                unpaid_count += 1
            else:
                if oms_pushed_col:
                    df_pending.at[idx, oms_pushed_col] = "Not Pushed"
                if is_gsheet_cancelled:
                    # Rule 1 exception: If TC status cancelled and OMS order not found, ignore!
                    df_pending.at[idx, "OMS Order Status"] = "Not in OMS"
                    df_pending.at[idx, "Final Remarks"] = "Cancelled (Ignored)"
                else:
                    df_pending.at[idx, "OMS Order Status"] = "Not in OMS"
                    df_pending.at[idx, "Final Remarks"] = "Not Pushed to OMS"
                    not_pushed_count += 1
                
                discrepancies.append({
                    "Order ID": order_id_str,
                    "Nickname": store_val,
                    "SKU": sku_val,
                    "Payment Status": pay_stat_val,
                    "Payment Method": pay_meth_val,
                    "Validation Result": "Not Pushed to OMS",
                    "TC Order Status": pend_stat_raw if pend_status_col else "N/A",
                    "TC Item Status": pend_stat_raw if pend_status_col else "N/A",
                    "OMS Order Status": "Not in OMS",
                    "OMS Line Status": "Not in OMS",
                    "Details": "Order ID is present in Pending SLA report but missing from OMS Report."
                })

    # Format SLA Date column
    if target_sla_col in df_pending.columns:
        df_pending[target_sla_col] = df_pending[target_sla_col].fillna("#N/A")
        df_pending[target_sla_col] = df_pending[target_sla_col].apply(extract_date)

    df_discrepancies = pd.DataFrame(discrepancies) if discrepancies else pd.DataFrame(columns=[
        "Order ID", "Nickname", "Seller SKU", "Validation Result", "TC Order Status", "TC Item Status", "OMS Order Status", "OMS Line Status", "Details"
    ])
    if not df_discrepancies.empty and "SKU" in df_discrepancies.columns:
        df_discrepancies = df_discrepancies.rename(columns={"SKU": "Seller SKU"})

    # Seller grouping
    seller_email_map = _build_seller_email_map(df_contacts)
    seller_groups = {}
    stores = df_pending[target_store_col].unique()
    for store in stores:
        store_clean = _clean_str(store)
        store_df = df_pending[df_pending[target_store_col] == store].copy()
        matched_email = seller_email_map.get(store_clean.strip().lower(), "")
        store_df["Seller Email"] = matched_email
        seller_groups[store_clean] = {
            "df": store_df,
            "email": matched_email
        }

    # Country-specific reports
    country_reports = {}
    for country in ["SG", "MY", "PH"]:
        country_reports[country] = {
            "raw_df": pd.DataFrame(),
            "pivot_df": pd.DataFrame(),
            "summary_df": pd.DataFrame()
        }

    df_pending["Order Date"] = df_pending[target_sla_col].apply(extract_date) if target_sla_col in df_pending.columns else "Unknown"

    for country in ["SG", "MY", "PH"]:
        c_rows = []
        for idx, row in df_pending.iterrows():
            store_val = _clean_str(row[target_store_col])
            c_code, chan = parse_country_and_channel(store_val)
            if c_code == country:
                final_rem = _clean_str(row.get("Final Remarks", ""))
                oms_stat = _clean_str(row.get("OMS Order Status", ""))
                if oms_stat.lower() == "shipped":
                    continue
                sla_val = row.get(target_sla_col)
                if _is_blank(sla_val) or str(sla_val).strip() == "#N/A":
                    continue
                row_dict = row.to_dict()
                row_dict["Country"] = country
                row_dict["Channel"] = f"{chan} {country}"
                c_rows.append(row_dict)
                
        country_df = pd.DataFrame(c_rows)
        if not country_df.empty:
            pivot_df = country_df.pivot_table(
                index=["Channel", "OMS Order Status"],
                columns="Order Date",
                values="Correct Order Number",
                aggfunc="count",
                fill_value=0
            )
            new_cols = []
            for col in pivot_df.columns:
                try:
                    dt = pd.to_datetime(col)
                    new_cols.append(dt.strftime('%d-%m-%Y'))
                except Exception:
                    new_cols.append(col)
            pivot_df.columns = new_cols
            pivot_df["Grand Total"] = pivot_df.sum(axis=1)
            pivot_df.loc[("Grand Total", ""), :] = pivot_df.sum(axis=0)
            pivot_df = pivot_df.reset_index()

            summary_metrics = [
                {"Metric": "Overdue (SLA breached)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "breached").sum()) if "sla_status" in country_df else 0},
                {"Metric": "Handover today (Today SLA)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "today").sum()) if "sla_status" in country_df else 0},
                {"Metric": "Order Status at New", "Count": int((country_df["OMS Order Status"].astype(str).str.strip().str.lower() == "new").sum())},
                {"Metric": "Within SLA (Future)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "future").sum()) if "sla_status" in country_df else 0},
                {"Metric": "Not reflecting in OM", "Count": int((country_df["OMS Order Status"] == "Not in OMS").sum())},
                {"Metric": "Unpaid Orders", "Count": int(country_df["Final Remarks"].astype(str).str.contains("Unpaid", case=False).sum()) if "Final Remarks" in country_df else 0}
            ]
            summary_df = pd.DataFrame(summary_metrics)
            
            cols_to_drop = ["Correct Order Number", "SLA Source", "Order Date", "Country", "Channel"]
            country_df_export = country_df.drop(columns=[c for c in cols_to_drop if c in country_df.columns])
            
            country_reports[country] = {
                "raw_df": country_df_export,
                "pivot_df": pivot_df,
                "summary_df": summary_df
            }

    ref_date_dmy = ""
    try:
        ref_dt = datetime.strptime(ref_date, "%Y-%m-%d")
        ref_date_dmy = ref_dt.strftime("%d-%m-%Y")
    except Exception:
        ref_date_dmy = ref_date

    summary = {
        "total_pending_orders": len(df_pending),
        "enriched_sla_count": 0,
        "blank_sla_not_found": 0,
        "total_discrepancies": len(df_discrepancies),
        "cancelled_mismatches": 0,
        "packed_mismatches": 0,
        "pushed_count": pushed_count,
        "not_pushed_count": not_pushed_count,
        "unpaid_count": unpaid_count,
        "total_sellers": len(seller_groups),
        "mode": "gsheet_oms"
    }

    return {
        "enriched_pending_df": df_pending,
        "discrepancies_df": df_discrepancies,
        "summary": summary,
        "seller_groups": seller_groups,
        "pending_order_id_col": target_sla_col,
        "country_reports": country_reports,
        "ref_date_dmy": ref_date_dmy
    }

def detect_dataframe_channel(df):
    columns = list(df.columns)
    columns_lower = [str(c).strip().lower() for c in columns]
    
    # 1. Lazada check
    if any("lazada" in c or "lazadasku" in c or "lazada sku" in c for c in columns_lower):
        country = "PH"
        for col in columns:
            if str(col).strip().lower() in ("shippingcountry", "shipping country", "country"):
                for val in df[col].dropna():
                    val_str = str(val).strip().upper()
                    if val_str:
                        if "SINGAPORE" in val_str or val_str == "SG":
                            country = "SG"
                        elif "MALAYSIA" in val_str or val_str == "MY":
                            country = "MY"
                        elif "PHILIPPINES" in val_str or val_str == "PH":
                            country = "PH"
                        else:
                            country = val_str
                        break
                break
        return f"Lazada {country}"

    # 2. Shopee check
    if any("shopee rebate" in c or "shopee_rebate" in c for c in columns_lower):
        country = "MY"
        for col in columns:
            if "php" in str(col).lower():
                country = "PH"
                break
        for col in columns:
            if str(col).strip().lower() in ("country", "shipping country"):
                for val in df[col].dropna():
                    val_str = str(val).strip().upper()
                    if val_str in ("SG", "MY", "PH"):
                        country = val_str
                        break
                break
        return f"Shopee {country}"

    # 3. Zalora check
    if any("zalora" in c or "zalora sku" in c or "zalorasku" in c for c in columns_lower):
        country = "PH"
        for col in columns:
            if "currency" in str(col).lower() or "order currency" in str(col).lower():
                for val in df[col].dropna():
                    val_str = str(val).strip().upper()
                    if "SGD" in val_str:
                        country = "SG"
                    elif "MYR" in val_str:
                        country = "MY"
                    elif "PHP" in val_str:
                        country = "PH"
                    break
                break
        if country == "PH":
            for col in columns:
                if str(col).strip().lower() in ("shipping country", "shippingcountry", "country"):
                    for val in df[col].dropna():
                        val_str = str(val).strip().upper()
                        if val_str in ("SG", "MY", "PH"):
                            country = val_str
                            break
                    break
        return f"Zalora {country}"

    # 4. TikTok check
    possible_id_cols = ["order id", "order_id", "orderid", "order number", "order_number", "ordernumber"]
    target_id_col = None
    for col in columns:
        if str(col).strip().lower() in possible_id_cols:
            target_id_col = col
            break
    if not target_id_col:
        for col in columns:
            if "order" in str(col).lower() or "id" in str(col).lower():
                target_id_col = col
                break
    if target_id_col:
        for val in df[target_id_col].dropna().head(10):
            val_str = str(val).strip()
            if len(val_str) == 18 and val_str.isdigit():
                return "TikTok MY"

    # Default fallback
    store_col = None
    for col in columns:
        if str(col).strip().lower() in ("nickname", "store name", "store", "seller", "seller name", "marketplace", "shop name", "shop"):
            store_col = col
            break
    if store_col:
        for val in df[store_col].dropna():
            store_val = str(val).strip()
            if store_val and store_val != "Default Store":
                c_code, chan = parse_country_and_channel(store_val)
                return f"{chan} {c_code}"

    return "Marketplace Default"

def resolve_marketplace_order_id_col(channel_name, columns):
    chan_lower = channel_name.lower()
    if "shopee" in chan_lower:
        for col in columns:
            if str(col).strip().lower() == "order id":
                return col
        for col in columns:
            if "order id" in str(col).strip().lower() or "order_id" in str(col).strip().lower() or "orderid" in str(col).strip().lower():
                return col
    elif "lazada" in chan_lower:
        for col in columns:
            if str(col).strip().lower() == "ordernumber":
                return col
        for col in columns:
            if "ordernumber" in str(col).strip().lower() or "order number" in str(col).strip().lower():
                return col
    elif "zalora" in chan_lower:
        for col in columns:
            if str(col).strip().lower() == "order number":
                return col
        for col in columns:
            if "order number" in str(col).strip().lower() or "ordernumber" in str(col).strip().lower():
                return col
    elif "tiktok" in chan_lower:
        for col in columns:
            if str(col).strip().lower() == "order id":
                return col
        for col in columns:
            if "order id" in str(col).strip().lower() or "order_id" in str(col).strip().lower() or "orderid" in str(col).strip().lower():
                return col
                
    for name in ["order_id", "order_number", "order id", "order no", "order number", "order_no", "orderid"]:
        for col in columns:
            if str(col).strip().lower() == name:
                return col
    return columns[0] if len(columns) > 0 else None

def is_ignored_tiktok(val):
    if not val:
        return False
    s = str(val).strip().lower().replace(" ", "").replace("-", "").replace("_", "")
    if s in ("tiktokph", "tiktoksg", "tiktoksingapore", "tiktokpumasingapore", "tiktokpumasg"):
        return True
    if "tiktok" in s and ("ph" in s or "sg" in s or "singapore" in s):
        return True
    return False

def load_and_preprocess_marketplace_files(marketplace_files):
    if not isinstance(marketplace_files, list):
        marketplace_files = [marketplace_files]
        
    processed_dfs = []
    for f in marketplace_files:
        sub_df = load_file_safely(f)
        if sub_df.empty:
            continue
            
        channel_detected = detect_dataframe_channel(sub_df)
        if is_ignored_tiktok(channel_detected):
            continue
        ord_id_col = resolve_marketplace_order_id_col(channel_detected, sub_df.columns)
        if not ord_id_col:
            continue
            
        sku_col = _find_column(sub_df, ["custom_sku", "customSku", "sku", "item_sku", "SKU", "sellerSku", "seller SKU", "Zalora SKU"])
        date_col = _find_column(sub_df, ["date", "created", "order_date", "Order Date", "timeOrderCreated", "Created at"])
        
        clean_rows = []
        for idx, row in sub_df.iterrows():
            raw_id = str(row.get(ord_id_col, "")).strip()
            if not raw_id or raw_id.lower() in ("nan", "null", "none", "platform unique order id.", "order id", "order no"):
                continue
            if "tiktok" in channel_detected.lower():
                if not (len(raw_id) == 18 and raw_id.isdigit()):
                    continue
                    
            clean_id = _clean_order_id(raw_id)
            sku_val = str(row.get(sku_col, "")).strip() if sku_col else ""
            date_val = extract_date(row.get(date_col, "")) if date_col else "Unknown"
            
            clean_rows.append({
                "Correct Order Number": clean_id,
                "Store Name": channel_detected,
                "SKU": sku_val,
                "Order Date": date_val,
                "Final Remarks": "",
                "OMS Order Status": "N/A",
                "SLA Source": "Marketplace Report",
                "SLA": "",
                "sla_status": ""
            })
            
        if clean_rows:
            processed_dfs.append(pd.DataFrame(clean_rows))
            
    if processed_dfs:
        return pd.concat(processed_dfs, ignore_index=True)
    else:
        return pd.DataFrame(columns=["Correct Order Number", "Store Name", "SKU", "Order Date", "Final Remarks", "OMS Order Status", "SLA Source", "SLA", "sla_status"])

def run_tc_marketplace_reconciliation(df_tc, df_marketplace):
    if df_tc is None or df_tc.empty:
        raise ValueError("TC Order Report is empty or was not uploaded. Please upload the TC Order Report.")
    tc_id_col = _find_column(df_tc, ["order_number", "order_id", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    if not tc_id_col:
        raise KeyError(f"Could not find 'Order ID' column in TC Order Report. Available: {list(df_tc.columns)}")

    is_tiktok_ph = lambda x: str(x).strip().lower().replace(" ", "").replace("-", "").replace("_", "") == "tiktokph"
    tc_store_col = _find_column(df_tc, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
    if tc_store_col and tc_store_col in df_tc.columns:
        df_tc = df_tc[~df_tc[tc_store_col].apply(is_tiktok_ph)].copy()

    df_tc[tc_id_col] = df_tc[tc_id_col].apply(_clean_order_id)
    tc_ids = set(df_tc[tc_id_col].dropna().astype(str).str.strip().tolist())

    discrepancies = []
    reflected_count = 0
    missing_count = 0
    target_store_col = "Store Name"

    for idx, row in df_marketplace.iterrows():
        order_id_str = str(row["Correct Order Number"]).strip()
        store_val = str(row["Store Name"]).strip()
        sku_val = str(row.get("SKU", "")).strip()
        
        if order_id_str in tc_ids:
            df_marketplace.at[idx, "Final Remarks"] = "Imported to TC"
            reflected_count += 1
        else:
            df_marketplace.at[idx, "Final Remarks"] = "Order missing in TC"
            missing_count += 1
            
            discrepancies.append({
                "Order ID": order_id_str,
                "Nickname": store_val,
                "Seller SKU": sku_val,
                "Validation Result": "Order missing in TC",
                "TC Order Status": "Missing",
                "TC Item Status": "Missing",
                "OMS Order Status": "N/A",
                "OMS Line Status": "N/A",
                "Details": "Order is present in Marketplace reports but completely missing from TC Order Report."
            })

    df_discrepancies = pd.DataFrame(discrepancies) if discrepancies else pd.DataFrame(columns=[
        "Order ID", "Nickname", "Seller SKU", "Validation Result", "TC Order Status", "TC Item Status", "OMS Order Status", "OMS Line Status", "Details"
    ])

    seller_groups = {}
    stores = df_marketplace[target_store_col].unique()
    for store in stores:
        store_clean = _clean_str(store)
        store_df = df_marketplace[df_marketplace[target_store_col] == store].copy()
        store_df["Seller Email"] = ""
        seller_groups[store_clean] = {
            "df": store_df,
            "email": ""
        }

    country_reports = {}
    for country in ["SG", "MY", "PH"]:
        country_reports[country] = {
            "raw_df": pd.DataFrame(),
            "pivot_df": pd.DataFrame(),
            "summary_df": pd.DataFrame()
        }

    for country in ["SG", "MY", "PH"]:
        c_rows = []
        for idx, row in df_marketplace.iterrows():
            store_val = str(row[target_store_col]).strip()
            if store_val.endswith(country):
                row_dict = row.to_dict()
                row_dict["Country"] = country
                row_dict["Channel"] = store_val
                c_rows.append(row_dict)
                
        country_df = pd.DataFrame(c_rows)
        if not country_df.empty:
            pivot_df = country_df.pivot_table(
                index=["Channel", "Final Remarks"],
                columns="Order Date",
                values="Correct Order Number",
                aggfunc="count",
                fill_value=0
            )
            pivot_df.index.names = ["Channel", "OMS Order Status"]
            
            cols_to_drop_pivot = [col for col in pivot_df.columns if str(col).lower().strip() in ("nan", "none", "nat", "null", "#n/a", "")]
            pivot_df = pivot_df.drop(columns=cols_to_drop_pivot, errors="ignore")
            
            new_cols = []
            for col in pivot_df.columns:
                try:
                    dt = pd.to_datetime(col)
                    new_cols.append(dt.strftime('%d-%m-%Y'))
                except Exception:
                    new_cols.append(col)
            pivot_df.columns = new_cols
            pivot_df["Grand Total"] = pivot_df.sum(axis=1)
            pivot_df.loc[("Grand Total", ""), :] = pivot_df.sum(axis=0)
            pivot_df = pivot_df.reset_index()

            summary_metrics = [
                {"Metric": "Reflected in TC", "Count": int((country_df["Final Remarks"] == "Imported to TC").sum())},
                {"Metric": "Missing from TC", "Count": int((country_df["Final Remarks"] == "Order missing in TC").sum())},
                {"Metric": "Order Status at New", "Count": 0},
                {"Metric": "Within SLA (Future)", "Count": 0},
                {"Metric": "Not reflecting in OM", "Count": 0}
            ]
            summary_df = pd.DataFrame(summary_metrics)
            
            country_df_missing = country_df[country_df["Final Remarks"] == "Order missing in TC"].copy()
            if not country_df_missing.empty:
                country_df_export = pd.DataFrame({
                    "Order ID": country_df_missing["Correct Order Number"],
                    "Channel Name": country_df_missing[target_store_col]
                })
            else:
                country_df_export = pd.DataFrame(columns=["Order ID", "Channel Name"])
            
            country_reports[country] = {
                "raw_df": country_df_export,
                "pivot_df": pivot_df,
                "summary_df": summary_df
            }

    ref_date_str = datetime.today().strftime('%d-%m-%Y')
    df_missing = df_marketplace[df_marketplace["Final Remarks"] == "Order missing in TC"].copy()
    if not df_missing.empty:
        df_missing_export = pd.DataFrame({
            "Order ID": df_missing["Correct Order Number"],
            "Channel Name": df_missing[target_store_col]
        })
    else:
        df_missing_export = pd.DataFrame(columns=["Order ID", "Channel Name"])

    summary = {
        "total_pending_orders": len(df_marketplace),
        "enriched_sla_count": 0,
        "blank_sla_not_found": 0,
        "total_discrepancies": len(df_discrepancies),
        "cancelled_mismatches": 0,
        "packed_mismatches": 0,
        "pushed_count": reflected_count,
        "not_pushed_count": missing_count,
        "unpaid_count": 0,
        "total_sellers": len(seller_groups),
        "all_imported_to_tc": (len(df_discrepancies) == 0),
        "mode": "tc_marketplace"
    }

    return {
        "enriched_pending_df": df_missing_export,
        "discrepancies_df": df_discrepancies,
        "summary": summary,
        "seller_groups": seller_groups,
        "pending_order_id_col": "Order ID",
        "country_reports": country_reports,
        "ref_date_dmy": ref_date_str
    }

def run_tc_oms_reconciliation(df_tc, df_marketplace, df_oms):
    if df_tc is None or df_tc.empty:
        raise ValueError("TC Order Report is empty or was not uploaded. Please upload the TC Order Report.")
    import re
    # Find column names in TC
    tc_id_col = _find_column(df_tc, ["order_number", "order_id", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    tc_num_col = _find_column(df_tc, ["order_no", "Order No", "Order Number", "Order Number / Reference No"])
    tc_status_col = _find_column(df_tc, ["order_status", "TC Status", "Order Status", "Status", "TC_Status"])
    tc_item_status_col = _find_column(df_tc, ["order_item_status", "item_status", "line_item_status", "order_status"])
    tc_store_col = _find_column(df_tc, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
    tc_sku_col = _find_column(df_tc, ["seller_sku", "sellerSku", "Seller SKU", "SellerSKU", "custom_sku", "customSku", "sku"])
    tc_sla_col = _find_column(df_tc, ["time_to_ship_dead_line", "order_sla", "SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date"])
    
    tc_pay_status_col = _find_column(df_tc, ["payment_status", "Payment Status", "Payment_Status", "PaymentStatus", "Payment"])
    tc_pay_method_col = _find_column(df_tc, ["payment_methods", "Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"])

    # Find column names in OMS
    oms_id_col = _find_column(df_oms, ["order_no", "order_id", "order_number", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    oms_status_col = _find_column(df_oms, ["order_status", "OMS Status", "Order Status", "Status", "OMS_Status"])
    oms_line_status_col = _find_column(df_oms, ["line_status", "OMS Line Status", "item_status", "order_item_status", "Line Status", "Line_Status", "oms_line_status"])
    oms_ean_col = _find_column(df_oms, ["ean", "EAN", "Ean", "item_sku", "SKU"])
    oms_store_col = _find_column(df_oms, ["store", "nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
    oms_pay_status_col = _find_column(df_oms, ["payment_status", "Payment Status", "Payment_Status", "PaymentStatus", "Payment"]) if df_oms is not None else None
    oms_pay_method_col = _find_column(df_oms, ["payment_methods", "Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"]) if df_oms is not None else None

    # Filter out TikTok PH and TikTok SG/Singapore from input datasets
    def is_ignored_tiktok(val):
        if not val:
            return False
        s = str(val).strip().lower().replace(" ", "").replace("-", "").replace("_", "")
        if s in ("tiktokph", "tiktoksg", "tiktoksingapore", "tiktokpumasingapore", "tiktokpumasg"):
            return True
        if "tiktok" in s and ("ph" in s or "sg" in s or "singapore" in s):
            return True
        return False

    if not df_tc.empty and tc_store_col and tc_store_col in df_tc.columns:
        df_tc = df_tc[~df_tc[tc_store_col].apply(is_ignored_tiktok)].copy()
    if df_marketplace is not None and not df_marketplace.empty:
        mp_store_col = _find_column(df_marketplace, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
        if mp_store_col and mp_store_col in df_marketplace.columns:
            df_marketplace = df_marketplace[~df_marketplace[mp_store_col].apply(is_ignored_tiktok)].copy()
    if df_oms is not None and not df_oms.empty and oms_store_col and oms_store_col in df_oms.columns:
        df_oms = df_oms[~df_oms[oms_store_col].apply(is_ignored_tiktok)].copy()

    # Clean IDs
    if not df_tc.empty and tc_id_col:
        df_tc[tc_id_col] = df_tc[tc_id_col].apply(_clean_order_id)
    if not df_tc.empty and tc_num_col:
        df_tc[tc_num_col] = df_tc[tc_num_col].apply(_clean_order_id)
    if df_oms is not None and not df_oms.empty and oms_id_col:
        df_oms[oms_id_col] = df_oms[oms_id_col].apply(_clean_order_id)

    # Build bidirectional maps for Zalora
    tc_id_to_num = {}
    tc_num_to_id = {}
    if not df_tc.empty and tc_id_col and tc_num_col and tc_id_col != tc_num_col:
        for row in df_tc.to_dict('records'):
            oid = row.get(tc_id_col)
            onum = row.get(tc_num_col)
            if oid and onum:
                tc_id_to_num[oid] = onum
                tc_num_to_id[onum] = oid

    # Build OMS status and payment lookup maps with concatenated Order ID + SKU keys
    oms_status_map = {}
    oms_line_status_map = {}
    oms_order_status_fallback_map = {}
    oms_line_status_fallback_map = {}
    oms_pay_status_map = {}
    oms_pay_method_map = {}
    if df_oms is not None and not df_oms.empty and oms_id_col:
        for row in df_oms.to_dict('records'):
            oid = _clean_order_id(row.get(oms_id_col))
            if not oid:
                continue
            ean = normalize_ean(row.get(oms_ean_col)) if oms_ean_col else ""
            sku_raw = _clean_str(row.get(oms_ean_col, "")).lower() if oms_ean_col else ""
            stat_val = _clean_str(row.get(oms_status_col, "")) if oms_status_col else ""
            line_stat_val = _clean_str(row.get(oms_line_status_col, "")) if oms_line_status_col else stat_val
            
            if ean:
                oms_status_map[oid + ean] = stat_val
                oms_line_status_map[oid + ean] = line_stat_val
            if sku_raw:
                oms_status_map[oid + sku_raw] = stat_val
                oms_line_status_map[oid + sku_raw] = line_stat_val
                
            oms_order_status_fallback_map[oid] = stat_val
            oms_line_status_fallback_map[oid] = line_stat_val
            if oms_pay_status_col:
                oms_pay_status_map[oid] = _clean_str(row.get(oms_pay_status_col, ""))
            if oms_pay_method_col:
                oms_pay_method_map[oid] = _clean_str(row.get(oms_pay_method_col, ""))

    ref_date_str = datetime.today().strftime('%d-%m-%Y')
    tc_active_statuses = {"new", "ready to ship", "accepted/picked", "picked", "accepted"}

    main_rows = []
    discrepancy_rows = []
    unpaid_count = 0
    pushed_count = 0
    not_pushed_count = 0

    # Mode 4 Check: If Marketplace reports are uploaded alongside TC and OMS (Order Flow Check)
    mode_name = "order_status_reconciliation"
    if df_marketplace is not None and not df_marketplace.empty:
        mode_name = "order_flow_check"
        tc_ids = set()
        if not df_tc.empty and tc_id_col:
            tc_ids.update(df_tc[tc_id_col].dropna().astype(str).str.strip().tolist())
        if not df_tc.empty and tc_num_col:
            tc_ids.update(df_tc[tc_num_col].dropna().astype(str).str.strip().tolist())
            
        for mp_row in df_marketplace.to_dict('records'):
            mp_oid = str(mp_row.get("Correct Order Number", "")).strip()
            if not mp_oid:
                continue
            mp_store = str(mp_row.get("Store Name", "")).strip()
            mp_sku = str(mp_row.get("SKU", "")).strip()
            
            if mp_oid not in tc_ids:
                flow_disc = {
                    "Order ID": mp_oid,
                    "Store Name": mp_store,
                    "Seller SKU": mp_sku,
                    "TC Order Status": "Missing in TC",
                    "TC Item Status": "Missing in TC",
                    "Payment Status": "N/A",
                    "Payment Method": "N/A",
                    "SLA Date": "Unknown",
                    "SLA": "Unknown",
                    "sla_status": "Unknown",
                    "OMS Order Status": "N/A",
                    "OMS Line Status": "N/A",
                    "Validation Result": "Order missing in TC",
                    "Details": "Order is present in Marketplace reports but completely missing from TC Order Report.",
                    "Final Remarks": "Order missing in TC",
                    "Correct Order Number": mp_oid,
                    "SLA Source": "Marketplace Report"
                }
                main_rows.append(flow_disc)
                discrepancy_rows.append(flow_disc)

    # Process all rows from df_tc for complete status reconciliation (optimized using to_dict)
    tc_rows = df_tc.to_dict('records')

    for row in tc_rows:
        oid_str = _clean_order_id(row.get(tc_id_col, ""))
        if not oid_str:
            continue

        tc_stat = _clean_str(row.get(tc_status_col, "")) if tc_status_col else ""
        tc_item_stat = _clean_str(row.get(tc_item_status_col, "")) if tc_item_status_col else ""
        tc_stat_norm = _normalize_status_val(tc_stat)
        tc_item_stat_norm = _normalize_status_val(tc_item_stat)

        sku_val = _clean_str(row.get(tc_sku_col, "")) if tc_sku_col else ""
        ean_val = normalize_ean(sku_val)
        sku_raw_val = sku_val.lower()
        
        pay_stat = _clean_str(row.get(tc_pay_status_col, "")) if tc_pay_status_col else ""
        pay_meth = _clean_str(row.get(tc_pay_method_col, "")) if tc_pay_method_col else ""
        if not pay_stat and oid_str in oms_pay_status_map:
            pay_stat = oms_pay_status_map[oid_str]
        if not pay_meth and oid_str in oms_pay_method_map:
            pay_meth = oms_pay_method_map[oid_str]

        # Check if this order is an Unpaid Order (payment status is NOT_INITIATED, etc. and method is not COD)
        is_unpaid = False
        if pay_stat:
            is_cod = any(term in pay_meth.lower() for term in ["cod", "cash on delivery", "cashondelivery"])
            is_pending = (pay_stat.lower() in ("pending", "unpaid", "awaiting", "not_initiated", "not_initiate", "not initiated"))
            if is_pending and not is_cod:
                is_unpaid = True

        oms_stat = ""
        oms_line_stat = ""
        key1 = oid_str + ean_val if ean_val else ""
        key2 = oid_str + sku_raw_val if sku_raw_val else ""
        onum = tc_id_to_num.get(oid_str, "")
        key3 = onum + ean_val if onum and ean_val else ""
        key4 = onum + sku_raw_val if onum and sku_raw_val else ""
        
        if key1 and key1 in oms_status_map:
            oms_stat = oms_status_map[key1]
            oms_line_stat = oms_line_status_map.get(key1, oms_stat)
        elif key2 and key2 in oms_status_map:
            oms_stat = oms_status_map[key2]
            oms_line_stat = oms_line_status_map.get(key2, oms_stat)
        elif oid_str in oms_order_status_fallback_map:
            oms_stat = oms_order_status_fallback_map[oid_str]
            oms_line_stat = oms_line_status_fallback_map.get(oid_str, oms_stat)
        elif key3 and key3 in oms_status_map:
            oms_stat = oms_status_map[key3]
            oms_line_stat = oms_line_status_map.get(key3, oms_stat)
        elif key4 and key4 in oms_status_map:
            oms_stat = oms_status_map[key4]
            oms_line_stat = oms_line_status_map.get(key4, oms_stat)
        elif onum and onum in oms_order_status_fallback_map:
            oms_stat = oms_order_status_fallback_map[onum]
            oms_line_stat = oms_line_status_fallback_map.get(onum, oms_stat)

        if not oms_line_stat:
            oms_line_stat = oms_stat if oms_stat else "Not in OMS"

        store_val = _clean_str(row.get(tc_store_col, "")) if tc_store_col else "Default Store"
        store_val = re.sub(r'puma_', '', store_val, flags=re.IGNORECASE)

        sla_raw = row.get(tc_sla_col, "") if tc_sla_col else ""
        sla_date = extract_date(sla_raw) if _clean_str(sla_raw) else "Unknown"
        sla_status_str = compute_sla_status(sla_date, ref_date_str) if sla_date != "Unknown" else "Unknown"

        # Reconcile strictly between TC Item Status (fallback TC Order Status) and OMS Line Status (fallback OMS Order Status)
        item_norm = tc_item_stat_norm if tc_item_stat_norm else tc_stat_norm
        line_norm = _clean_str(oms_line_stat).lower() if oms_line_stat else _clean_str(oms_stat).lower()

        # NOTE on tc_is_return: "returned" (final/terminal) is treated separately
        # from in-progress return states like "RETURN ACCEPTED", "RETURN
        # REQUESTED", "RETURN SHIPPED". Per the Status Mismatch reference sheet,
        # only the terminal "Returned" status should be compared against OMS
        # (and flagged if OMS hasn't caught up); in-progress return states are
        # always ignored regardless of what OMS currently shows.
        tc_is_return_final = (item_norm.strip() == "returned")
        tc_is_return = ("return" in item_norm)
        tc_is_return_progress = tc_is_return and not tc_is_return_final
        tc_is_cancelled = ("cancel" in item_norm or "lost" in item_norm or "refund" in item_norm)
        tc_is_failed = ("failed" in item_norm)
        tc_is_delivered = ("delivered" in item_norm)
        tc_is_new = ("new" in item_norm)
        tc_is_ready = ("ready" in item_norm or "to_ship" in item_norm or tc_is_new)
        is_active_in_tc = tc_is_new or tc_is_ready
        tc_is_blank = (item_norm == "")
        tc_is_plain_shipped = ("shipped" in item_norm) and not (
            tc_is_new or tc_is_ready or tc_is_delivered or tc_is_return or tc_is_cancelled or tc_is_failed
        )

        oms_is_returned = ("return" in line_norm)
        oms_is_cancelled = ("cancel" in line_norm or "void" in line_norm or "refund" in line_norm)
        oms_is_delivered = ("delivered" in line_norm)
        oms_is_shipped = ("shipped" in line_norm or "ship" in line_norm or "sent" in line_norm)
        oms_is_packed = ("packed" in line_norm or "pack" in line_norm)

        val_result = "OK"
        details = "Order item status matched between TC and OMS."
        final_remarks = f"Successfully Pushed to OMS ({oms_stat})" if oms_stat else "Successfully Pushed to OMS"
        is_disc = False

        if is_unpaid and (not oms_stat or oms_stat == "Not in OMS"):
            val_result = "Not Pushed to OMS - Unpaid orders"
            details = "Order is unpaid in TC and not pushed to OMS."
            final_remarks = "Not Pushed to OMS - Unpaid orders"
            oms_stat = "Not in OMS - Unpaid orders"
            oms_line_stat = "Not in OMS - Unpaid orders"
            is_disc = False
            unpaid_count += 1

        elif tc_is_return_final:
            # FIX: previously only checked oms_is_shipped/oms_is_packed, so
            # Returned(TC) + Delivered(OMS) silently passed as "Ignored". Per
            # the reference sheet this combination must be flagged.
            if oms_is_shipped or oms_is_packed or oms_is_delivered:
                val_result = "Need to push Returned Status"
                details = f"TC Item Status is Returned, but OMS Line Status is '{oms_line_stat}'."
                final_remarks = "Need to push Returned Status"
                is_disc = True
            else:
                val_result = "Returned (Ignored)"
                details = f"TC Item Status is '{tc_item_stat or tc_stat}', OMS Line Status is '{oms_line_stat}' (Returned/Cancelled in OMS is a valid end-state - Ignored)."
                final_remarks = f"Successfully Pushed to OMS ({oms_stat})" if oms_stat else "Returned (Ignored)"
                is_disc = False

        elif tc_is_return_progress:
            # Return Requested/Accepted/Shipped are in-progress states in TC -
            # OMS can legitimately show a variety of statuses while the return
            # is being processed, so this is never treated as a mismatch.
            val_result = "Returned (Ignored)"
            details = f"TC Item Status is '{tc_item_stat or tc_stat}' (return in progress), OMS Line Status is '{oms_line_stat}' - Ignored."
            final_remarks = f"Successfully Pushed to OMS ({oms_stat})" if oms_stat else "Returned (Ignored)"
            is_disc = False

        elif tc_is_failed and (oms_is_returned or oms_is_cancelled or oms_is_delivered):
            # Delivery failed orders returned to warehouse - Ignored
            val_result = "Delivery Failed (Returned)"
            details = f"TC Item Status is Delivery Failed, OMS Line Status is '{oms_line_stat}' (Delivery Failed orders returned to warehouse - Ignored)."
            final_remarks = f"Successfully Pushed to OMS ({oms_stat})" if oms_stat else "Delivery Failed (Returned)"
            is_disc = False

        elif tc_is_cancelled:
            if oms_is_cancelled or oms_is_returned or not oms_stat or oms_stat == "Not in OMS":
                # Cancelled or Lost in TC and Returned/Cancelled in OMS - Ignored
                val_result = "Cancelled/Returned (Ignored)"
                details = f"TC Item Status is Cancelled, OMS Line Status is '{oms_line_stat}' (Normal cancellation/return - Ignored)."
                final_remarks = "Cancelled/Returned (Ignored)"
                is_disc = False
            else:
                val_result = "Cancelled Status Mismatch"
                details = f"TC Item Status is Cancelled, but OMS Line Status shows '{oms_line_stat}'."
                final_remarks = "Cancelled Status Mismatch"
                is_disc = True

        elif is_active_in_tc and (not oms_stat or oms_stat == "Not in OMS"):
            val_result = "Not Pushed to OMS"
            details = "Paid or COD order is present in TC but missing from OMS Report."
            final_remarks = "Not Pushed to OMS"
            is_disc = True
            not_pushed_count += 1

        elif not oms_stat or oms_stat == "Not in OMS":
            val_result = "Not in OMS"
            details = f"TC Item Status is '{tc_item_stat or tc_stat}', but order is missing from OMS Report."
            final_remarks = "Not in OMS"
            is_disc = True
            not_pushed_count += 1

        elif tc_is_blank and oms_is_delivered:
            # TC Item Status missing but OMS already shows Delivered - flag per
            # reference sheet (previously fell through to a silent "OK").
            val_result = "Need to push Returned Status"
            details = f"TC Item Status is missing/blank, but OMS Line Status is '{oms_line_stat}'."
            final_remarks = "Need to push Returned Status"
            is_disc = True

        elif oms_is_cancelled and not (tc_is_cancelled or tc_is_return_final or tc_is_return_progress or tc_is_failed or tc_is_delivered):
            val_result = "Cancelled Status Mismatch"
            details = f"OMS Line Status is Cancelled, but TC Item Status is '{tc_item_stat or tc_stat}'."
            final_remarks = "Cancelled Status Mismatch"
            is_disc = True

        elif tc_is_delivered and not (oms_is_delivered or oms_is_returned or oms_is_cancelled):
            # FIX: Delivered(TC) + Returned(OMS) and Delivered(TC) + Cancelled(OMS)
            # are both valid "OK" end-states per the reference sheet and must
            # not be flagged - only an in-transit OMS status (Shipped/Packed)
            # while TC already shows Delivered is a genuine mismatch.
            val_result = "Need to push Delivered Status"
            details = f"TC Item Status is Delivered, but OMS Line Status is '{oms_line_stat}'."
            final_remarks = "Need to push Delivered Status"
            is_disc = True

        elif tc_is_plain_shipped and oms_is_delivered:
            # TC still shows plain "Shipped" while OMS has already moved to
            # Delivered - previously no branch caught this at all.
            val_result = "Need to push Shipped status to TC"
            details = f"TC Item Status is Shipped, but OMS Line Status is already '{oms_line_stat}'."
            final_remarks = "Need to push Shipped status to TC"
            is_disc = True

        elif oms_is_shipped and tc_is_ready and not tc_is_delivered:
            val_result = "OMS Shipped but TC Status Invalid"
            details = f"OMS Line Status is Shipped, but Pending item status in TC is '{tc_item_stat or tc_stat}'."
            final_remarks = "OMS Shipped but TC Status Invalid"
            is_disc = True

        elif oms_is_packed and tc_is_new and not tc_is_delivered:
            val_result = "OMS Packed but TC New"
            details = "OMS Line Status is Packed, but TC Item Status is New."
            final_remarks = "OMS Packed but TC New"
            is_disc = True

        elif tc_is_failed and not (oms_is_returned or oms_is_cancelled):
            val_result = "TC Delivery Failed but OMS not Returned"
            details = f"TC Item Status is Delivery Failed, but OMS Line Status is '{oms_line_stat}'."
            final_remarks = "TC Delivery Failed but OMS not Returned"
            is_disc = True

        # If it's not a real discrepancy, and TC Item Status and OMS Line
        # Status are effectively the same status text (e.g. both "Cancelled",
        # both "Returned"), show a plain "OK" instead of a descriptive
        # ignored-reason label - the descriptive labels are only useful when
        # the two sides actually differ.
        if not is_disc:
            _squash = lambda s: re.sub(r'[\s_\-/]+', '', s.strip().lower())
            if _squash(item_norm) == _squash(line_norm):
                val_result = "OK"
                details = "TC Item Status and OMS Line Status match."
                final_remarks = f"Successfully Pushed to OMS ({oms_stat})" if oms_stat else "OK"

        # Check if it was successfully pushed to OMS
        is_pushed = oms_stat and oms_stat != "Not in OMS" and "not in oms" not in str(oms_stat).lower()
        if is_pushed:
            pushed_count += 1

        row_data = {
            "Order ID": oid_str,
            "Store Name": store_val,
            "Seller SKU": sku_val,
            "TC Order Status": tc_stat,
            "TC Item Status": tc_item_stat,
            "Payment Status": pay_stat,
            "Payment Method": pay_meth,
            "SLA Date": sla_date,
            "SLA": sla_status_str,
            "sla_status": sla_status_str,
            "OMS Order Status": oms_stat if oms_stat else "Not in OMS",
            "OMS Line Status": oms_line_stat,
            "Validation Result": val_result,
            "Details": details,
            "Final Remarks": final_remarks,
            "Correct Order Number": oid_str,
            "SLA Source": "TC Order Report"
        }

        main_rows.append(row_data)
        if is_disc:
            discrepancy_rows.append(row_data)

    main_df = pd.DataFrame(main_rows) if main_rows else pd.DataFrame(columns=[
        "Order ID", "Store Name", "Seller SKU", "TC Order Status", "TC Item Status",
        "Payment Status", "Payment Method", "SLA Date", "SLA", "sla_status",
        "OMS Order Status", "Validation Result", "Details", "Final Remarks"
    ])
    
    df_discrepancies = pd.DataFrame(discrepancy_rows) if discrepancy_rows else pd.DataFrame(columns=main_df.columns)
    
    target_store_col = "Store Name"
    total_discrepancies = len(df_discrepancies)

    seller_groups = {}
    country_reports = {}

    reflected_count = len(main_df) - total_discrepancies
    missing_count = total_discrepancies

    summary = {
        "total_pending_orders": len(main_df),
        "enriched_sla_count": len(main_df),
        "blank_sla_not_found": 0,
        "total_discrepancies": total_discrepancies,
        "cancelled_mismatches": len([r for r in main_rows if "Cancelled" in r["Validation Result"]]),
        "packed_mismatches": len([r for r in main_rows if "Packed" in r["Validation Result"]]),
        "pushed_count": pushed_count,
        "not_pushed_count": not_pushed_count,
        "unpaid_count": unpaid_count,
        "total_sellers": len(seller_groups),
        "all_imported_to_tc": (total_discrepancies == 0),
        "mode": mode_name
    }

    # Build seller groupings
    seller_groups = {}
    if not main_df.empty:
        for s_name, sub_grp in main_df.groupby(target_store_col):
            s_name_clean = _clean_str(s_name)
            seller_groups[s_name_clean] = {"df": sub_grp, "email": ""}

    # Country-specific reports
    country_reports = {}
    for country in ["SG", "MY", "PH"]:
        country_reports[country] = {
            "raw_df": pd.DataFrame(),
            "pivot_df": pd.DataFrame(),
            "summary_df": pd.DataFrame()
        }

    if not main_df.empty:
        main_df["_country"] = main_df[target_store_col].apply(lambda x: parse_country_and_channel(x)[0])
        for country, country_grp in main_df.groupby("_country"):
            if country not in ["SG", "MY", "PH"]:
                continue
            c_rows = []
            for idx, row in country_grp.iterrows():
                store_val = _clean_str(row[target_store_col])
                c_code, chan = parse_country_and_channel(store_val)
                row_dict = row.to_dict()
                row_dict["Country"] = country
                row_dict["Channel"] = f"{chan} {country}"
                c_rows.append(row_dict)
                
            country_df = pd.DataFrame(c_rows)
            if not country_df.empty:
                col_to_use = "Final Remarks"
                pivot_df = country_df.pivot_table(
                    index=["Channel", "OMS Order Status"],
                    columns=col_to_use,
                    values="Correct Order Number",
                    aggfunc="count",
                    fill_value=0
                )
                pivot_df["Grand Total"] = pivot_df.sum(axis=1)
                pivot_df.loc[("Grand Total", ""), :] = pivot_df.sum(axis=0)
                pivot_df = pivot_df.reset_index()
                
                summary_metrics = [
                    {"Metric": "Overdue (SLA breached)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "breached").sum())},
                    {"Metric": "Handover today (Today SLA)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "today").sum())},
                    {"Metric": "Order Status at New", "Count": int((country_df["OMS Order Status"].astype(str).str.strip().str.lower() == "new").sum())},
                    {"Metric": "Within SLA (Future)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "future").sum())},
                    {"Metric": "Not reflecting in OM", "Count": int((country_df["OMS Order Status"] == "Not in OMS").sum())},
                    {"Metric": "Unpaid Orders", "Count": 0}
                ]
                summary_df = pd.DataFrame(summary_metrics)
                
                cols_to_drop = ["Correct Order Number", "SLA Source", "Order Date", "Country", "Channel", "_country"]
                country_df_export = country_df.drop(columns=[c for c in cols_to_drop if c in country_df.columns])
                
                country_reports[country] = {
                    "raw_df": country_df_export,
                    "pivot_df": pivot_df,
                    "summary_df": summary_df
                }

    reflected_count = len(main_df) - total_discrepancies
    missing_count = total_discrepancies

    summary = {
        "total_pending_orders": len(main_df),
        "enriched_sla_count": len(main_df),
        "blank_sla_not_found": 0,
        "total_discrepancies": total_discrepancies,
        "cancelled_mismatches": len([r for r in main_rows if "Cancelled" in r["Validation Result"]]),
        "packed_mismatches": len([r for r in main_rows if "Packed" in r["Validation Result"]]),
        "pushed_count": reflected_count,
        "not_pushed_count": missing_count,
        "unpaid_count": 0,
        "total_sellers": len(seller_groups),
        "all_imported_to_tc": (total_discrepancies == 0),
        "mode": "tc_oms"
    }

    return {
        "enriched_pending_df": main_df,
        "discrepancies_df": df_discrepancies,
        "summary": summary,
        "seller_groups": seller_groups,
        "pending_order_id_col": "Order ID",
        "country_reports": country_reports,
        "ref_date_dmy": ref_date_str
    }

def process_and_validate_orders(pending_file, tc_file, *args, **kwargs):
    FILE_LOAD_WARNINGS.clear()

    # Resolve marketplace_file and oms_file based on positional args length
    marketplace_file = None
    oms_file = None
    contacts_file = None
    
    if len(args) == 1:
        # 3 positional arguments: process_and_validate_orders(pending, tc, oms)
        oms_file = args[0]
    elif len(args) >= 2:
        # 4+ positional arguments: process_and_validate_orders(pending, tc, marketplace, oms, [contacts])
        marketplace_file = args[0]
        oms_file = args[1]
        if len(args) >= 3:
            contacts_file = args[2]
            
    # Fallback to keyword arguments if not set by positional args
    if marketplace_file is None:
        marketplace_file = kwargs.get('marketplace_file', None)
    if oms_file is None:
        oms_file = kwargs.get('oms_file', None)
    if contacts_file is None:
        contacts_file = kwargs.get('contacts_file', None)

    # Load all dataframes
    df_pending = pd.DataFrame()
    df_tc = pd.DataFrame()
    df_marketplace = pd.DataFrame()
    df_oms = pd.DataFrame()
    
    # Load Pending Report
    if isinstance(pending_file, str) and pending_file == "mcp":
        df_pending = fetch_pending_from_mcp()
    elif pending_file is not None:
        if isinstance(pending_file, pd.DataFrame):
            df_pending = pending_file
        else:
            if isinstance(pending_file, str) and (pending_file.startswith("http://") or pending_file.startswith("https://")):
                pending_file = download_google_sheet(pending_file)
            df_pending = load_file_safely(pending_file)
            
    # Load TC Report
    if isinstance(tc_file, list):
        tc_dfs = []
        for f in tc_file:
            sub_df = load_file_safely(f)
            if not sub_df.empty:
                tc_dfs.append(sub_df)
        df_tc = pd.concat(tc_dfs, ignore_index=True) if tc_dfs else pd.DataFrame()
    elif tc_file is not None:
        df_tc = load_file_safely(tc_file)
        
    # Load Marketplace Reports using platform-specific rules
    if marketplace_file is not None:
        df_marketplace = load_and_preprocess_marketplace_files(marketplace_file)
        
    # Load OMS Report
    if isinstance(oms_file, list):
        oms_dfs = []
        for f in oms_file:
            sub_df = load_file_safely(f)
            if not sub_df.empty:
                oms_dfs.append(sub_df)
        df_oms = pd.concat(oms_dfs, ignore_index=True) if oms_dfs else pd.DataFrame()
    elif oms_file is not None:
        df_oms = load_file_safely(oms_file)
        
    df_contacts = load_file_safely(contacts_file) if contacts_file is not None else pd.DataFrame()

    has_pending = not df_pending.empty
    has_tc = not df_tc.empty
    has_marketplace = not df_marketplace.empty
    has_oms = not df_oms.empty

    # New Mode: TC Report (+ Marketplace Report) + OMS Report (but no Pending/GSheet report)
    if has_tc and has_oms and not has_pending:
        result = run_tc_oms_reconciliation(df_tc, df_marketplace, df_oms)
        result["file_load_warnings"] = list(FILE_LOAD_WARNINGS)
        return result

    # Mode 2: TC Order Report + Marketplace Reports alone
    elif (has_tc or has_marketplace) and not has_pending and not has_oms:
        result = run_tc_marketplace_reconciliation(df_tc, df_marketplace)
        result["file_load_warnings"] = list(FILE_LOAD_WARNINGS)
        return result
        
    # Mode 1: GSheet + OMS Report Alone
    elif has_pending and has_oms and not has_tc and not has_marketplace:
        result = run_gsheet_oms_validation(df_pending, df_oms, df_contacts)
        result["file_load_warnings"] = list(FILE_LOAD_WARNINGS)
        return result
        
    # Default Mode 3/4: standard validation with combined TC + Marketplace reports
    else:
        combined_tc_dfs = []
        if not df_tc.empty:
            combined_tc_dfs.append(df_tc)
        if not df_marketplace.empty:
            combined_tc_dfs.append(df_marketplace)
            
        if combined_tc_dfs:
            df_tc = pd.concat(combined_tc_dfs, ignore_index=True)
            
        result = run_standard_validation(df_pending, df_tc, df_oms, df_contacts)
        result["file_load_warnings"] = list(FILE_LOAD_WARNINGS)
        return result

def run_standard_validation(df_pending, df_tc, df_oms, df_contacts):

    # If the g sheet / pending file is not uploaded, generate df_pending from TC Report
    # pending orders from TC with status of NEW, READY TO SHIP & ACCEPTED/PICKED
    if df_pending.empty:
        tc_id_col = _find_column(df_tc, ["order_number", "order_id", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
        tc_item_status_col = _find_column(df_tc, ["order_item_status", "item_status", "line_item_status", "order_status"])
        tc_status_col = _find_column(df_tc, ["order_status", "TC Status", "Order Status", "Status", "TC_Status"])
        tc_sla_col = _find_column(df_tc, ["time_to_ship_dead_line", "order_sla", "SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date"])
        tc_store_col = _find_column(df_tc, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
        
        status_col = tc_item_status_col if tc_item_status_col else tc_status_col
        pending_statuses = {"new", "ready to ship", "accepted/picked", "picked", "accepted"}
        
        tc_pending_rows = []
        if status_col and tc_id_col:
            for _, row in df_tc.iterrows():
                stat = _normalize_status_val(row.get(status_col))
                is_pending = False
                for p_stat in pending_statuses:
                    if p_stat in stat:
                        is_pending = True
                        break
                if is_pending:
                    tc_pending_rows.append(row)
                    
        if tc_pending_rows:
            df_pending = pd.DataFrame(tc_pending_rows)
            # Map required columns
            df_pending["Order ID"] = df_pending[tc_id_col].apply(_clean_order_id)
            if tc_sla_col and tc_sla_col in df_pending.columns:
                df_pending["SLA"] = df_pending[tc_sla_col]
            else:
                df_pending["SLA"] = ""
            if tc_store_col and tc_store_col in df_pending.columns:
                df_pending["Store Name"] = df_pending[tc_store_col]
            else:
                df_pending["Store Name"] = "Default Store"

    has_pending = not df_pending.empty

    # Find store column in SLA Report first to ignore TikTok PH
    pend_store_col = None
    if has_pending:
        pend_store_col = _find_column(df_pending, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
        if pend_store_col and pend_store_col in df_pending.columns:
            # Filter out TikTok PH orders case-insensitively
            is_tiktok_ph = lambda x: str(x).strip().lower().replace(" ", "").replace("-", "").replace("_", "") == "tiktokph"
            df_pending = df_pending[~df_pending[pend_store_col].apply(is_tiktok_ph)].copy()
            
            # Filter df_tc if store column is present
            tc_store_col = _find_column(df_tc, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
            if tc_store_col and tc_store_col in df_tc.columns:
                df_tc = df_tc[~df_tc[tc_store_col].apply(is_tiktok_ph)].copy()
                
            # Filter df_oms if store column is present
            oms_store_col = _find_column(df_oms, ["store", "nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])
            if oms_store_col and oms_store_col in df_oms.columns:
                df_oms = df_oms[~df_oms[oms_store_col].apply(is_tiktok_ph)].copy()

        if df_pending.empty:
            raise ValueError("Pending Order Report (SLA Report) has no rows after filtering out TikTok PH.")
        
        has_pending = not df_pending.empty

    # == 2. Standardize Columns ===============================================
    pend_id_col = None
    pend_sla_col = None
    target_sla_col = "SLA"
    target_store_col = "Store Name"

    if has_pending:
        # Pending Order columns
        pend_id_col = _find_column(df_pending, ["order_id", "order_number", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
        pend_sla_col = _find_column(df_pending, ["mp_sla_date", "SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date", "mp_sla_date_updated"])
        if not pend_id_col:
            raise KeyError(f"Could not find 'Order ID' column in Pending Order Report. Available: {list(df_pending.columns)}")
        df_pending[pend_id_col] = df_pending[pend_id_col].apply(_clean_order_id)
        
        target_sla_col = pend_sla_col if pend_sla_col else "SLA"
        if target_sla_col not in df_pending.columns:
            df_pending[target_sla_col] = ""

        target_store_col = pend_store_col if pend_store_col else "Store Name"
        if target_store_col not in df_pending.columns:
            df_pending[target_store_col] = "Default Store"
        oms_pushed_col = _find_column(df_pending, ["oms_pushed", "omsPushed", "OMS Pushed", "OMS_pushed"])
        if not oms_pushed_col:
            df_pending["oms_pushed"] = ""
            oms_pushed_col = "oms_pushed"
    
    # TC Report columns (All file)
    tc_id_col = _find_column(df_tc, ["order_number", "order_id", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    tc_num_col = _find_column(df_tc, ["order_number", "order_id"]) # specifically find order_number column
    tc_status_col = _find_column(df_tc, ["order_status", "TC Status", "Order Status", "Status", "TC_Status"])
    
    # SLA Lookup targets time_to_ship_dead_line as primary, then fallback
    tc_sla_col = _find_column(df_tc, ["time_to_ship_dead_line", "order_sla", "SLA", "SLA Date", "SLA_Date", "Ship By Date", "ship_by_date"])
    
    tc_pay_status_col = _find_column(df_tc, ["payment_status", "Payment Status", "Payment_Status", "PaymentStatus", "Payment"])
    tc_pay_method_col = _find_column(df_tc, ["payment_methods", "Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"])
    
    # OMS Report columns (Sales Order file)
    oms_id_col = _find_column(df_oms, ["order_no", "order_id", "order_number", "Order ID", "Order No", "Order Number", "Order_No", "Order_ID"])
    oms_status_col = _find_column(df_oms, ["order_status", "OMS Status", "Order Status", "Status", "OMS_Status"])
    oms_pay_status_col = _find_column(df_oms, ["Payment Status", "Payment_Status", "PaymentStatus", "Payment"])
    oms_pay_method_col = _find_column(df_oms, ["Payment Method", "Payment_Method", "PaymentMethod", "Payment Type"])

    # Raise error if critical columns are missing
    if not tc_id_col:
        raise KeyError(f"Could not find 'Order ID' column in TC Report (All file). Available: {list(df_tc.columns)}")
    if not oms_id_col:
        raise KeyError(f"Could not find 'Order ID' column in OMS Report (Sales Order file). Available: {list(df_oms.columns)}")

    # Clean Order IDs to ensure matches (retaining large string values correctly)
    df_tc[tc_id_col] = df_tc[tc_id_col].apply(_clean_order_id)
    df_oms[oms_id_col] = df_oms[oms_id_col].apply(_clean_order_id)
    if tc_num_col:
        df_tc[tc_num_col] = df_tc[tc_num_col].apply(_clean_order_id)

    # == 3. SLA Enrichment & Pushed Status ====================================
    # Build TC mappings
    tc_sla_map = {}
    if tc_sla_col:
        tc_sla_map = df_tc.set_index(tc_id_col)[tc_sla_col].dropna().to_dict()

    tc_payment_status = {}
    if tc_pay_status_col:
        tc_payment_status = df_tc.set_index(tc_id_col)[tc_pay_status_col].dropna().to_dict()

    tc_payment_method = {}
    if tc_pay_method_col:
        tc_payment_method = df_tc.set_index(tc_id_col)[tc_pay_method_col].dropna().to_dict()

    # Build maps for original statuses and SKUs from TC to show in Status Discrepancies sheet
    tc_order_status_map = {}
    if tc_status_col:
        tc_order_status_map = df_tc.set_index(tc_id_col)[tc_status_col].dropna().to_dict()

    tc_item_status_col = _find_column(df_tc, ["order_item_status", "item_status", "line_item_status", "order_status"])
    tc_item_status_map = {}
    if tc_item_status_col:
        tc_item_status_map = df_tc.set_index(tc_id_col)[tc_item_status_col].dropna().to_dict()

    tc_custom_sku_col = _find_column(df_tc, ["seller_sku", "sellerSku", "Seller SKU", "SellerSKU", "custom_sku", "customSku", "custom_SKU", "sku"])
    tc_sku_map = {}
    if tc_custom_sku_col:
        tc_sku_map = df_tc.set_index(tc_id_col)[tc_custom_sku_col].dropna().to_dict()

    # Build bidirectional ID to package-number mappings from TC report (crucial for Zalora package lookup)
    tc_id_to_num = {}
    tc_num_to_id = {}
    if tc_id_col and tc_num_col and tc_id_col != tc_num_col:
        for _, row in df_tc.iterrows():
            oid = row[tc_id_col]
            onum = row[tc_num_col]
            if oid and onum:
                tc_id_to_num[oid] = onum
                tc_num_to_id[onum] = oid

    # Build OMS status lookup maps
    oms_status_map = df_oms.set_index(oms_id_col)[oms_status_col].dropna().to_dict() if oms_status_col else {}
    oms_pay_status_map = df_oms.set_index(oms_id_col)[oms_pay_status_col].dropna().to_dict() if oms_pay_status_col else {}
    oms_pay_method_map = df_oms.set_index(oms_id_col)[oms_pay_method_col].dropna().to_dict() if oms_pay_method_col else {}

    enriched_sla_count = 0
    blank_sla_not_found = 0
    pushed_count = 0
    not_pushed_count = 0
    unpaid_count = 0
    ref_date = datetime.today().strftime('%Y-%m-%d')
    missing_tc_discrepancies = []

    # Build set of all unique order IDs/numbers in TC Report for quick cross-checking
    tc_order_ids_set = set()
    if tc_id_col and tc_id_col in df_tc.columns:
        tc_order_ids_set.update(df_tc[tc_id_col].dropna().astype(str).str.strip().tolist())
    # Add mapped Zalora package numbers
    tc_order_ids_set.update(tc_id_to_num.keys())
    tc_order_ids_set.update(tc_num_to_id.keys())

    if has_pending:
        # Add output columns to Pending Order Report
        df_pending["Correct Order Number"] = df_pending[pend_id_col]
        df_pending["SLA Source"] = "Pending Report"
        df_pending["OMS Order Status"] = ""
        df_pending["Final Remarks"] = ""

        # Clean Store Name / nickname column to remove PUMA_ prefix case-insensitively
        if target_store_col in df_pending.columns:
            df_pending[target_store_col] = df_pending[target_store_col].apply(
                lambda x: re.sub(r'puma_', '', str(x).strip(), flags=re.IGNORECASE)
            )

        # Determine reference date from "Today" rows or data
        temp_ref_date = None
        if "sla_status" in df_pending.columns:
            today_rows = df_pending[df_pending["sla_status"].astype(str).str.strip().str.lower() == "today"]
            if not today_rows.empty:
                for val in today_rows[target_sla_col]:
                    if not _is_blank(val):
                        p_dt = extract_date(val)
                        if len(p_dt) == 10:
                            temp_ref_date = p_dt
                            break
        if not temp_ref_date:
            if target_sla_col in df_pending.columns:
                for val in df_pending[target_sla_col]:
                    if not _is_blank(val):
                        p_dt = extract_date(val)
                        if len(p_dt) == 10:
                            temp_ref_date = p_dt
                            break
        if temp_ref_date:
            ref_date = temp_ref_date

        # Ensure sla_status column exists in df_pending
        pend_sku_col = _find_column(df_pending, ["seller_sku", "sellerSku", "Seller SKU", "SellerSKU", "custom_sku", "customSku", "sku", "item_sku", "SKU", "orderItems.customSKU"])
        for idx, row in df_pending.iterrows():
            order_id = row[pend_id_col]
            sla_val = row[target_sla_col]
            
            # Clean string formats for checking
            order_id_str = str(order_id).strip()
            
            # ── SLA Check ──
            if _is_blank(sla_val): # Blank SLA (recognizing "NaT", "nan" loaded via dtype=str)
                tc_sla_val = _clean_str(tc_sla_map.get(order_id, ""))
                if not _is_blank(tc_sla_val):
                    df_pending.at[idx, target_sla_col] = tc_sla_val
                    df_pending.at[idx, "SLA Source"] = "Enriched from TC"
                    enriched_sla_count += 1
                else:
                    df_pending.at[idx, target_sla_col] = "#N/A"
                    df_pending.at[idx, "SLA Source"] = "Missing (Not in TC)"
                    blank_sla_not_found += 1
            else:
                df_pending.at[idx, "SLA Source"] = "Pending Report"

            # Update/Compute sla_status for blanks/enriched or if not present
            sla_status_val = row.get("sla_status", "")
            if _is_blank(sla_status_val):
                curr_sla = df_pending.at[idx, target_sla_col]
                calculated_status = compute_sla_status(curr_sla, ref_date)
                if calculated_status:
                    df_pending.at[idx, "sla_status"] = calculated_status

            # ── OMS Status & Final Remarks Check (Checking both SLA ID and Mapped TC Order Number) ──
            tc_mapped_num = tc_id_to_num.get(order_id, "")
            tc_mapped_num_str = str(tc_mapped_num).strip() if tc_mapped_num else ""
            
            is_in_oms = False
            oms_stat = ""
            
            if order_id:
                if order_id in oms_status_map:
                    is_in_oms = True
                    oms_stat = oms_status_map[order_id]
                elif tc_mapped_num_str and tc_mapped_num_str in oms_status_map:
                    is_in_oms = True
                    oms_stat = oms_status_map[tc_mapped_num_str]
                
            # Perform TC cross-check
            is_in_tc = (order_id_str in tc_order_ids_set) or (tc_mapped_num_str and tc_mapped_num_str in tc_order_ids_set)
            
            pend_sku_val = ""
            if pend_sku_col and pend_sku_col in row:
                pend_sku_val = row.get(pend_sku_col, "")
            if not is_in_tc:
                # Flag missing in TC
                df_pending.at[idx, "Final Remarks"] = "Order missing in TC"
                df_pending.at[idx, "OMS Order Status"] = oms_stat if is_in_oms else "Not in OMS"
                if oms_pushed_col:
                    df_pending.at[idx, oms_pushed_col] = "Pushed" if is_in_oms else "Not Pushed"
                
                # Append to discrepancy list
                store_val = _clean_str(row.get(target_store_col, ""))
                missing_tc_discrepancies.append({
                    "Order ID": order_id_str,
                    "Nickname": store_val,
                    "SKU": pend_sku_val,
                    "Validation Result": "Order missing in TC",
                    "TC Order Status": "Missing",
                    "TC Item Status": "Missing",
                    "OMS Order Status": oms_stat if is_in_oms else "Not in OMS",
                    "OMS Line Status": oms_stat if is_in_oms else "Not in OMS",
                    "Details": "Order is present in Pending SLA report but completely missing from TC Report."
                })
            else:
                if is_in_oms:
                    df_pending.at[idx, "OMS Order Status"] = oms_stat
                    df_pending.at[idx, "Final Remarks"] = "Successfully Pushed to OMS"
                    pushed_count += 1
                    if oms_pushed_col:
                        df_pending.at[idx, oms_pushed_col] = "Pushed"
                else:
                    df_pending.at[idx, "OMS Order Status"] = "Not in OMS"
                    if oms_pushed_col:
                        df_pending.at[idx, oms_pushed_col] = "Not Pushed"
                    
                    # Retrieve payment status & method from TC (All file) as fallback
                    pay_status = _clean_str(tc_payment_status.get(order_id, ""))
                    pay_method = _clean_str(tc_payment_method.get(order_id, ""))
                    
                    is_cod = any(term in pay_method.lower() for term in ["cod", "cash on delivery", "cashondelivery"])
                    is_pending = (pay_status.lower() in ("pending", "unpaid", "awaiting", "not_initiated", "not_initiate", "not initiated"))
                    
                    if is_pending and not is_cod:
                        df_pending.at[idx, "OMS Order Status"] = "Not in OMS - Unpaid orders"
                        df_pending.at[idx, "Final Remarks"] = "Not Pushed to OMS - Unpaid orders"
                        if oms_pushed_col:
                            df_pending.at[idx, oms_pushed_col] = "Not pushed - Unpaid Orders"
                        unpaid_count += 1
                    else:
                        df_pending.at[idx, "Final Remarks"] = "Not Pushed to OMS"
                        not_pushed_count += 1
                        
                        # Retrieve original status and SKU from TC for "Not Pushed to OMS" discrepancy
                        tc_sku_val = tc_sku_map.get(order_id_str, "")
                        if not tc_sku_val and tc_mapped_num_str:
                            tc_sku_val = tc_sku_map.get(tc_mapped_num_str, "")
                        if not tc_sku_val:
                            tc_sku_val = pend_sku_val
                            
                        tc_ord_status = tc_order_status_map.get(order_id_str, "")
                        if not tc_ord_status and tc_mapped_num_str:
                            tc_ord_status = tc_order_status_map.get(tc_mapped_num_str, "")
                        if not tc_ord_status:
                            tc_ord_status = pay_status if pay_status else "Paid/COD"
                            
                        tc_itm_status = tc_item_status_map.get(order_id_str, "")
                        if not tc_itm_status and tc_mapped_num_str:
                            tc_itm_status = tc_item_status_map.get(tc_mapped_num_str, "")
                        if not tc_itm_status:
                            tc_itm_status = pay_status if pay_status else "Paid/COD"
                        
                        # Add to discrepancies
                        store_val = _clean_str(row.get(target_store_col, ""))
                        missing_tc_discrepancies.append({
                            "Order ID": order_id_str,
                            "Nickname": store_val,
                            "SKU": tc_sku_val,
                            "Payment Status": pay_status,
                            "Payment Method": pay_method,
                            "Validation Result": "Not Pushed to OMS",
                            "TC Order Status": tc_ord_status,
                            "TC Item Status": tc_itm_status,
                            "OMS Order Status": "Not in OMS",
                            "OMS Line Status": "Not in OMS",
                            "Details": "Paid or COD order is present in TC but missing from OMS Report."
                        })

        # Format SLA Date column to show only Date (no Time) in output report
        if target_sla_col in df_pending.columns:
            df_pending[target_sla_col] = df_pending[target_sla_col].fillna("#N/A")
            df_pending[target_sla_col] = df_pending[target_sla_col].apply(extract_date)

    # == 4. Status Discrepancy Validations ====================================
    # Identify item-level columns
    tc_custom_sku_col = _find_column(df_tc, ["custom_sku", "customSku", "custom_SKU", "sku"])
    tc_item_status_col = _find_column(df_tc, ["order_item_status", "item_status", "line_item_status", "order_status"])
    
    # In OMS report SKU is in ean header
    oms_ean_col = _find_column(df_oms, ["ean", "sku", "OMS sku", "EAN"])
    oms_line_status_col = _find_column(df_oms, ["line_status", "order_status", "item_status"])

    # Fallbacks if columns are not resolved
    if not tc_custom_sku_col:
        tc_custom_sku_col = tc_id_col
    if not tc_item_status_col:
        tc_item_status_col = tc_status_col
        
    if not oms_ean_col:
        oms_ean_col = oms_id_col
    if not oms_line_status_col:
        oms_line_status_col = oms_status_col

    tc_lookup = {}
    for _, row in df_tc.iterrows():
        oid = _clean_order_id(row[tc_id_col])
        sku = _clean_str(row[tc_custom_sku_col])
        sku_norm = normalize_ean(sku)
        item_status = _clean_str(row[tc_item_status_col]) if tc_item_status_col else ""
        order_status = _clean_str(row[tc_status_col]) if tc_status_col else ""
        if oid:
            val_dict = {
                "Order ID": oid,
                "SKU": sku,
                "Item Status": item_status,
                "Order Status": order_status,
                "Row": row.to_dict()
            }
            tc_lookup[oid + sku_norm] = val_dict
            if tc_num_col:
                onum = _clean_order_id(row[tc_num_col])
                if onum and onum != oid:
                    tc_lookup[onum + sku_norm] = val_dict

    oms_lookup = {}
    for _, row in df_oms.iterrows():
        oid_raw = _clean_order_id(row[oms_id_col])
        oid = tc_num_to_id.get(oid_raw, oid_raw)
        ean = _clean_str(row[oms_ean_col])
        ean_norm = normalize_ean(ean)
        line_status = _clean_str(row[oms_line_status_col]) if oms_line_status_col else ""
        order_status = _clean_str(row[oms_status_col]) if oms_status_col else ""
        if oid_raw:
            val_dict = {
                "Order ID": oid,
                "SKU": ean,
                "Line Status": line_status,
                "Order Status": order_status,
                "Row": row.to_dict()
            }
            oms_lookup[oid_raw + ean_norm] = val_dict
            if oid and oid != oid_raw:
                oms_lookup[oid + ean_norm] = val_dict

    all_keys = set(tc_lookup.keys()) & set(oms_lookup.keys()) # intersect line item keys
    discrepancies = []

    # Get TC store column for nickname lookup
    tc_store_col = _find_column(df_tc, ["nickname", "Store Name", "Store", "Seller", "Seller Name", "Marketplace", "Shop Name", "Shop"])

    for key in all_keys:
        tc_item_status = tc_lookup[key]["Item Status"]
        tc_order_status = tc_lookup[key]["Order Status"]
        
        oms_line_status = oms_lookup[key]["Line Status"]
        oms_order_status = oms_lookup[key]["Order Status"]
        
        oid = tc_lookup[key]["Order ID"]
        sku = tc_lookup[key]["SKU"]
        
        nickname_val = ""
        if tc_store_col and tc_store_col in tc_lookup[key]["Row"]:
            nickname_val = tc_lookup[key]["Row"][tc_store_col]
            
        # Normalize for checks
        tc_item_norm = _normalize_status_val(tc_item_status)
        oms_line_norm = _normalize_status_val(oms_line_status)

        # Rule 1: Cancelled status check
        # NOTE: "lost" (e.g. LOST_BY_3PL) and "refund" are treated the same as
        # "cancel" here - previously only "cancel" was checked, so LOST_BY_3PL
        # rows never triggered any discrepancy at all (Status Mismatch sheet
        # row: LOST_BY_3PL / SHIPPED -> Status Mismatch).
        is_tc_cancelled = ("cancel" in tc_item_norm or "lost" in tc_item_norm or "refund" in tc_item_norm)
        is_oms_cancelled = ("cancel" in oms_line_norm)
        if (is_tc_cancelled or is_oms_cancelled) and (is_tc_cancelled != is_oms_cancelled):
            # Exceptions per the Status Mismatch reference sheet - all of these
            # are valid "OK" / ignored combinations, not mismatches:
            #  - TC Cancelled + OMS Returned
            #  - TC Cancelled + order missing/blank in OMS ("Not in OMS")
            #  - TC Delivered + OMS Cancelled
            #  - TC Returned (final) + OMS Cancelled
            is_oms_returned = ("return" in oms_line_norm)
            is_oms_not_in_oms = (oms_line_norm == "" or "not in oms" in oms_line_norm)
            is_tc_delivered_chk = ("delivered" in tc_item_norm)
            is_tc_returned_final_chk = (tc_item_norm.strip() == "returned")
            skip_flag = (
                (is_tc_cancelled and (is_oms_returned or is_oms_not_in_oms))
                or (is_oms_cancelled and (is_tc_delivered_chk or is_tc_returned_final_chk))
            )
            if not skip_flag:
                discrepancies.append({
                    "Order ID": oid,
                    "Nickname": nickname_val,
                    "SKU": sku,
                    "Validation Result": "Cancelled Status Mismatch",
                    "TC Order Status": tc_order_status,
                    "TC Item Status": tc_item_status,
                    "OMS Order Status": oms_order_status,
                    "OMS Line Status": oms_line_status,
                    "Details": f"Cancelled status mismatch: TC is '{tc_item_status}', OMS is '{oms_line_status}'."
                })
                
        # Rule 2: OMS Packed and TC Status if New can highlight
        if "packed" in oms_line_norm and "new" in tc_item_norm:
            discrepancies.append({
                "Order ID": oid,
                "Nickname": nickname_val,
                "SKU": sku,
                "Validation Result": "OMS Packed but TC New",
                "TC Order Status": tc_order_status,
                "TC Item Status": tc_item_status,
                "OMS Order Status": oms_order_status,
                "OMS Line Status": oms_line_status,
                "Details": f"OMS status is Packed, but TC status is '{tc_item_status}' (should be READY TO SHIP or ACCEPTED/PICKED)."
            })

        # Rule 3: If OMS Status is shipped and TC status with NEW, READY TO SHIP, ACCEPTED/PICKED & Cancelled can highlight
        if "shipped" in oms_line_norm:
            is_tc_invalid = False
            for val in ["new", "ready to ship", "accepted", "picked", "cancel"]:
                if val in tc_item_norm:
                    is_tc_invalid = True
                    break
            if is_tc_invalid:
                discrepancies.append({
                    "Order ID": oid,
                    "Nickname": nickname_val,
                    "SKU": sku,
                    "Validation Result": "OMS Shipped but TC Status Invalid",
                    "TC Order Status": tc_order_status,
                    "TC Item Status": tc_item_status,
                    "OMS Order Status": oms_order_status,
                    "OMS Line Status": oms_line_status,
                    "Details": f"OMS status is Shipped, but TC status is '{tc_item_status}'."
                })

        # Rule 4: If TC Status is Delivered and OMS status not Delivered can highlight
        # FIX: Delivered/Returned and Delivered/Cancelled are both valid terminal
        # combinations per the Status Mismatch reference sheet (Validation Result:
        # "OK", not included in the mismatch sheet) - only flag when OMS is still
        # showing an in-transit state (e.g. Shipped/Packed), not when it has
        # already reached another valid end-state (Returned/Cancelled).
        if "delivered" in tc_item_norm and not any(x in oms_line_norm for x in ["delivered", "return", "cancel"]):
            discrepancies.append({
                "Order ID": oid,
                "Nickname": nickname_val,
                "SKU": sku,
                "Validation Result": "Need to push Delivered Status",
                "TC Order Status": tc_order_status,
                "TC Item Status": tc_item_status,
                "OMS Order Status": oms_order_status,
                "OMS Line Status": oms_line_status,
                "Details": f"TC status is Delivered, but OMS status is '{oms_line_status}'."
            })

        # Rule 5: If TC Status is Returned and OMS Status not Returned can highlight
        # FIX: Returned/Cancelled is a valid "OK" combination per the reference
        # sheet and should not be flagged (previously it was).
        if "returned" in tc_item_norm and not ("returned" in oms_line_norm or "return" in oms_line_norm or "cancel" in oms_line_norm):
            discrepancies.append({
                "Order ID": oid,
                "Nickname": nickname_val,
                "SKU": sku,
                "Validation Result": "TC Returned but OMS not Returned",
                "TC Order Status": tc_order_status,
                "TC Item Status": tc_item_status,
                "OMS Order Status": oms_order_status,
                "OMS Line Status": oms_line_status,
                "Details": f"TC status is Returned, but OMS status is '{oms_line_status}'."
            })

        # Rule 5b: TC Item Status is blank/missing but OMS already shows Delivered.
        # Per the Status Mismatch reference sheet this should be flagged as
        # "Need to push Returned Status" (previously this fell through every
        # rule above and was silently marked OK).
        if not tc_item_norm and "delivered" in oms_line_norm:
            discrepancies.append({
                "Order ID": oid,
                "Nickname": nickname_val,
                "SKU": sku,
                "Validation Result": "Need to push Returned Status",
                "TC Order Status": tc_order_status,
                "TC Item Status": tc_item_status,
                "OMS Order Status": oms_order_status,
                "OMS Line Status": oms_line_status,
                "Details": f"TC Item Status is missing/blank, but OMS status is '{oms_line_status}'."
            })

        # Rule 5c: TC Item Status is plain "Shipped" (not New/Ready/Accepted/
        # Picked/Cancelled/Returned/Delivered/Failed) but OMS already shows
        # Delivered. Per the reference sheet: "Need to push Shipped status to TC".
        # Previously no rule caught this combination at all.
        tc_is_plain_shipped = "shipped" in tc_item_norm and not any(
            x in tc_item_norm for x in ["new", "ready", "accepted", "picked", "cancel", "return", "delivered", "failed", "lost"]
        )
        if tc_is_plain_shipped and "delivered" in oms_line_norm:
            discrepancies.append({
                "Order ID": oid,
                "Nickname": nickname_val,
                "SKU": sku,
                "Validation Result": "Need to push Shipped status to TC",
                "TC Order Status": tc_order_status,
                "TC Item Status": tc_item_status,
                "OMS Order Status": oms_order_status,
                "OMS Line Status": oms_line_status,
                "Details": f"TC status is Shipped, but OMS status is already '{oms_line_status}'."
            })

        # Rule 6: If TC Status Delivery Failed and OMS not Returned can highlight
        if "failed" in tc_item_norm and not ("returned" in oms_line_norm or "return" in oms_line_norm):
            discrepancies.append({
                "Order ID": oid,
                "Nickname": nickname_val,
                "SKU": sku,
                "Validation Result": "TC Delivery Failed but OMS not Returned",
                "TC Order Status": tc_order_status,
                "TC Item Status": tc_item_status,
                "OMS Order Status": oms_order_status,
                "OMS Line Status": oms_line_status,
                "Details": f"TC status is Delivery Failed, but OMS status is '{oms_line_status}'."
            })

    # == Shopee Partial Cancellation Mismatch Filter ==
    # PERF FIX: previously, for every order with a cancel mismatch, this
    # re-scanned the *entire* tc_lookup/oms_lookup dictionaries (twice, via
    # `next(...)` and again via list comprehensions) to find that order's
    # items. That's O(orders * total_line_items) - on large TC/OMS reports
    # (tens of thousands of rows) this was the single biggest slowdown in
    # report generation. Grouping by Order ID once up-front makes the whole
    # filter O(total_line_items) instead.
    filtered_discrepancies = []
    from collections import defaultdict
    order_discs = defaultdict(list)
    for disc in discrepancies:
        order_discs[disc["Order ID"]].append(disc)

    tc_items_by_order = defaultdict(list)
    for k, v in tc_lookup.items():
        tc_items_by_order[v["Order ID"]].append(v)

    oms_items_by_order = defaultdict(list)
    for k, v in oms_lookup.items():
        oms_items_by_order[v["Order ID"]].append(v)

    for oid, discs in order_discs.items():
        has_cancel_mismatch = any(d["Validation Result"] == "Cancelled Status Mismatch" for d in discs)
        if has_cancel_mismatch:
            tc_items = tc_items_by_order.get(oid, [])
            oms_items = oms_items_by_order.get(oid, [])

            # Check if this is a Shopee order
            sample_item = tc_items[0] if tc_items else (oms_items[0] if oms_items else None)
            is_shopee = False
            if sample_item is not None:
                s_val = ""
                if tc_items and tc_store_col and tc_store_col in tc_items[0]["Row"]:
                    s_val = tc_items[0]["Row"][tc_store_col]
                elif oms_items and "nickname" in oms_items[0]["Row"]:
                    s_val = oms_items[0]["Row"].get("nickname", "")
                elif oms_items and "store" in oms_items[0]["Row"]:
                    s_val = oms_items[0]["Row"].get("store", "")

                _, chan = parse_country_and_channel(s_val)
                if chan == "Shopee":
                    is_shopee = True

            if is_shopee:
                # Get all items in TC and OMS for this Order ID to see if it's a partial cancellation
                tc_cancelled_count = sum(1 for item in tc_items if "cancel" in _normalize_status_val(item["Item Status"]))
                oms_cancelled_count = sum(1 for item in oms_items if "cancel" in _normalize_status_val(item["Line Status"]))
                
                total_tc_items = len(tc_items)
                total_oms_items = len(oms_items)
                
                # Check if cancellation is partial
                is_partial_tc = (0 < tc_cancelled_count < total_tc_items)
                is_partial_oms = (0 < oms_cancelled_count < total_oms_items)
                
                # If either side is partially cancelled, we ignore the Cancelled Status Mismatch
                if is_partial_tc or is_partial_oms:
                    # Filter out Cancelled Status Mismatch for this order
                    for d in discs:
                        if d["Validation Result"] != "Cancelled Status Mismatch":
                            filtered_discrepancies.append(d)
                    continue

        filtered_discrepancies.extend(discs)

    # Append missing in TC discrepancies
    filtered_discrepancies.extend(missing_tc_discrepancies)

    df_discrepancies = pd.DataFrame(filtered_discrepancies) if filtered_discrepancies else pd.DataFrame(columns=[
        "Order ID", "Nickname", "Seller SKU", "Validation Result", "TC Order Status", "TC Item Status", "OMS Order Status", "OMS Line Status", "Details"
    ])
    if not df_discrepancies.empty and "SKU" in df_discrepancies.columns:
        df_discrepancies = df_discrepancies.rename(columns={"SKU": "Seller SKU"})

    # == 5. Seller Contact Map & Grouping =====================================
    seller_groups = {}
    if has_pending:
        email_map = {}
        if not df_contacts.empty:
            c_store = _find_column(df_contacts, ["Store Name", "Store", "Seller", "Shop Name", "Shop"])
            c_email = _find_column(df_contacts, ["Seller Email", "Email", "SellerEmail", "Email Address"])
            if c_store and c_email:
                for _, row in df_contacts.iterrows():
                    store_key = _clean_str(row[c_store]).lower()
                    email_val = _clean_str(row[c_email])
                    if store_key and email_val:
                        email_map[store_key] = email_val

        # Group enriched pending orders by Seller
        stores = df_pending[target_store_col].unique()

        for store in stores:
            store_clean = _clean_str(store)
            store_key = store_clean.lower()
            
            store_df = df_pending[df_pending[target_store_col] == store].copy()
            
            mapped_email = email_map.get(store_key, "")
            store_df["Seller Email"] = mapped_email
            
            seller_groups[store_clean] = {
                "df": store_df,
                "email": mapped_email
            }

    # == 6. Country-specific datasets & Pivot Tables ==========================
    country_reports = {}
    
    # Initialize empty reports for all countries
    for country in ["SG", "MY", "PH"]:
        country_reports[country] = {
            "raw_df": pd.DataFrame(),
            "pivot_df": pd.DataFrame(),
            "summary_df": pd.DataFrame()
        }

    if has_pending:
        # Pre-calculate clean SLA Date for columns
        date_col = target_sla_col
        if date_col and date_col in df_pending.columns:
            df_pending["Order Date"] = df_pending[date_col].apply(extract_date)
        else:
            df_pending["Order Date"] = "Unknown"
        
        for country in ["SG", "MY", "PH"]:
            c_rows = []
            for idx, row in df_pending.iterrows():
                store_val = _clean_str(row[target_store_col])
                c_code, chan = parse_country_and_channel(store_val)
                if c_code == country:
                    final_rem = _clean_str(row.get("Final Remarks", ""))
                    oms_stat = _clean_str(row.get("OMS Order Status", ""))
                    
                    # Ignore Unpaid orders and OMS shipped orders in country wise output reports
                    if oms_stat.lower() == "shipped":
                        continue
                    sla_val = row.get(target_sla_col)
                    if _is_blank(sla_val) or str(sla_val).strip() == "#N/A":
                        continue
                        
                    # Accept all channels dynamically
                    if country in ["SG", "MY", "PH"]:
                        row_dict = row.to_dict()
                        row_dict["Country"] = country
                        row_dict["Channel"] = f"{chan} {country}"
                        c_rows.append(row_dict)
                        
            country_df = pd.DataFrame(c_rows)
            if not country_df.empty:
                # Pivot table: Channel & OMS status in Rows, Order date in Columns, Order number (count) in Values
                pivot_df = country_df.pivot_table(
                    index=["Channel", "OMS Order Status"],
                    columns="Order Date",
                    values="Correct Order Number",
                    aggfunc="count",
                    fill_value=0
                )
                
                # Drop columns representing blank/invalid SLA dates (including nan, nat, #N/A)
                cols_to_drop_pivot = [col for col in pivot_df.columns if str(col).lower().strip() in ("nan", "none", "nat", "null", "#n/a", "")]
                pivot_df = pivot_df.drop(columns=cols_to_drop_pivot, errors="ignore")
                
                # Format columns of Pivot Table as DD-MM-YYYY
                new_cols = []
                for col in pivot_df.columns:
                    try:
                        dt = pd.to_datetime(col)
                        new_cols.append(dt.strftime('%d-%m-%Y'))
                    except Exception:
                        new_cols.append(col)
                pivot_df.columns = new_cols
                
                # Add Grand Total column
                pivot_df["Grand Total"] = pivot_df.sum(axis=1)
                # Add Grand Total row
                pivot_df.loc[("Grand Total", ""), :] = pivot_df.sum(axis=0)
                pivot_df = pivot_df.reset_index()
                
                # Highlight Summary metrics
                summary_metrics = [
                    {"Metric": "Overdue (SLA breached)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "breached").sum()) if "sla_status" in country_df else 0},
                    {"Metric": "Handover today (Today SLA)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "today").sum()) if "sla_status" in country_df else 0},
                    {"Metric": "Order Status at New", "Count": int((country_df["OMS Order Status"].astype(str).str.strip().str.lower() == "new").sum())},
                    {"Metric": "Within SLA (Future)", "Count": int((country_df["sla_status"].astype(str).str.strip().str.lower() == "future").sum()) if "sla_status" in country_df else 0},
                    {"Metric": "Not reflecting in OM", "Count": int((country_df["OMS Order Status"] == "Not in OMS").sum())},
                    {"Metric": "Unpaid Orders", "Count": int(country_df["Final Remarks"].astype(str).str.contains("Unpaid", case=False).sum()) if "Final Remarks" in country_df else 0}
                ]
                summary_df = pd.DataFrame(summary_metrics)
                
                # Drop unwanted columns from raw sheet data
                cols_to_drop = ["Correct Order Number", "SLA Source", "Order Date", "Country", "Channel"]
                country_df_export = country_df.drop(columns=[c for c in cols_to_drop if c in country_df.columns])
                
                country_reports[country] = {
                    "raw_df": country_df_export,
                    "pivot_df": pivot_df,
                    "summary_df": summary_df
                }

    # Summary metrics
    summary = {
        "total_pending_orders": len(df_pending) if has_pending else 0,
        "enriched_sla_count": enriched_sla_count,
        "blank_sla_not_found": blank_sla_not_found,
        "total_discrepancies": len(df_discrepancies),
        "cancelled_mismatches": int((df_discrepancies["Validation Result"] == "Cancelled Sync Mismatch").sum()),
        "packed_mismatches": int((df_discrepancies["Validation Result"] == "OMS Packed but TC New").sum()),
        "pushed_count": pushed_count,
        "not_pushed_count": not_pushed_count,
        "unpaid_count": unpaid_count,
        "total_sellers": len(seller_groups),
        "mode": "standard"
    }

    ref_date_dmy = ""
    try:
        ref_dt = datetime.strptime(ref_date, "%Y-%m-%d")
        ref_date_dmy = ref_dt.strftime("%d-%m-%Y")
    except Exception:
        ref_date_dmy = ref_date

    return {
        "enriched_pending_df": df_pending,
        "discrepancies_df": df_discrepancies,
        "summary": summary,
        "seller_groups": seller_groups,
        "pending_order_id_col": target_sla_col if has_pending else "",
        "country_reports": country_reports,
        "ref_date_dmy": ref_date_dmy,
        "tc_lookup": tc_lookup,
        "oms_lookup": oms_lookup
    }

