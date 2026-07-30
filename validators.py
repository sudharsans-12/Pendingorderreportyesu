# -*- coding: utf-8 -*-
import re
import pandas as pd


def _safe_num(val):
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def _safe_str(val):
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except (TypeError, ValueError):
        pass
    return str(val).strip()


def _normalise_status(status):
    s = _safe_str(status).lower()
    if s in ("active", "1", "enabled", "yes", "y", "live", "listed"):
        return "Active"
    if s in ("inactive", "0", "disabled", "no", "n", "delisted",
             "unlisted", "deleted", "removed"):
        return "Inactive"
    return _safe_str(status)


def _is_valid_sku(sku):
    """Seller SKU must be exactly 13 digits."""
    return bool(re.fullmatch(r'\d{13}', _safe_str(sku)))


def _normalise_article_no(val):
    """
    Normalise Article No for cross-file matching.
    Standardise separators (space/hyphen -> underscore), uppercase,
    strip leading/trailing underscores.
    Same logic as file_loaders._normalise_article_no, kept local to
    avoid cross-module import dependency.
    """
    s = _safe_str(val)
    if not s:
        return ""
    s = s.strip().upper()
    s = re.sub(r'[\s\-]+', '_', s)
    s = s.strip('_')
    return s


# == Lookup builders ==========================================================-

def _build_article_map(content):
    """SKU -> Article No (normalised). Tries multiple candidate columns."""
    article_map = {}
    if content.empty or "SKU" not in content.columns:
        return article_map
    art_col = next(
        (c for c in ["Article No", "ArticleNo", "Article Number",
                      "Color_No", "Color_No.1", "Style#", "STYLE#",
                      "Style #", "STYLE #"]
         if c in content.columns),
        next((c for c in content.columns
              if "article" in c.lower() or "color" in c.lower()
              or "style" in c.lower()), "")
    )
    if art_col:
        skus = content["SKU"].tolist()
        art_vals = content[art_col].tolist()
        for sku, art in zip(skus, art_vals):
            sku_s = _safe_str(sku)
            if sku_s:
                article_map[sku_s] = _normalise_article_no(art)
    return article_map


def _build_ecom_map(zecom, mp_name):
    """Article No -> Ecom Status for the given marketplace."""
    ecom_map = {}
    if zecom.empty or "Article No" not in zecom.columns:
        return ecom_map
    mp_key = mp_name.split()[0].lower()
    ecom_col = next(
        (c for c in zecom.columns
         if c.startswith("Ecom_") and mp_key in c.lower()), ""
    )
    if not ecom_col:
        return ecom_map
    arts = zecom["Article No"].tolist()
    ecom_vals = zecom[ecom_col].tolist()
    for art, val in zip(arts, ecom_vals):
        art_norm = _normalise_article_no(art)
        if art_norm:
            ecom_map[art_norm] = _safe_str(val)
    return ecom_map


def _build_tc_map(tc_inv):
    tc_map = {}
    parent_fallback = {}

    if tc_inv.empty or "SKU" not in tc_inv.columns:
        return tc_map

    skus = tc_inv["SKU"].tolist()
    tc_skus = tc_inv["TC SKU"].tolist() if "TC SKU" in tc_inv.columns else skus
    tc_statuses = tc_inv["TC Status"].tolist() if "TC Status" in tc_inv.columns else ["Unknown"] * len(skus)
    max_0s = tc_inv["Max 0"].tolist() if "Max 0" in tc_inv.columns else ["No"] * len(skus)

    for sku, tc_sku_raw, tc_status, max_0 in zip(skus, tc_skus, tc_statuses, max_0s):
        sku_s = _safe_str(sku)
        if not sku_s:
            continue
        entry = {
            "TC SKU":    _safe_str(tc_sku_raw),
            "TC Status": _safe_str(tc_status),
            "Max 0":     _safe_str(max_0),
        }
        if "-" in sku_s:
            tc_map[sku_s] = entry
            parent_base = sku_s.rsplit("-", 1)[0]
            if parent_base not in tc_map:
                parent_fallback[parent_base] = entry
        else:
            if sku_s not in tc_map:
                tc_map[sku_s] = entry

    for parent, entry in parent_fallback.items():
        if parent not in tc_map:
            tc_map[parent] = entry

    return tc_map


def _build_stock_map(all_df, apply_buffer=False):
    stock_map = {}
    if all_df.empty or "SKU" not in all_df.columns:
        return stock_map

    skus = all_df["SKU"].tolist()
    tc_stocks = all_df["TC Stock"].tolist() if "TC Stock" in all_df.columns else [0] * len(skus)
    reserved_stocks = all_df["Reserved Stock"].tolist() if "Reserved Stock" in all_df.columns else [0] * len(skus)

    for sku, tc_val, reserved_val in zip(skus, tc_stocks, reserved_stocks):
        sku_s = _safe_str(sku)
        if sku_s and sku_s not in stock_map:
            tc = _safe_num(tc_val)
            tc = max(tc, 0)
            if apply_buffer:
                tc = max(tc - 1, 0)
            stock_map[sku_s] = {
                "TC Stock":       tc,
                "Reserved Stock": _safe_num(reserved_val),
            }
    return stock_map


def _build_excl_map(exclusion):
    excl_map = {}
    if exclusion is None or exclusion.empty:
        return excl_map
    if "Article No" not in exclusion.columns:
        return excl_map

    art_nos = exclusion["Article No"].tolist()
    statuses = exclusion["Exclusion Status"].tolist() if "Exclusion Status" in exclusion.columns else ["Inactive"] * len(art_nos)

    for raw_val, status in zip(art_nos, statuses):
        status_s = _safe_str(status)
        art = _normalise_article_no(raw_val)
        if art:
            excl_map[art] = status_s

        raw_clean = re.sub(r'\D', '', _safe_str(raw_val))
        if re.fullmatch(r'\d{13}', raw_clean):
            excl_map[raw_clean] = status_s

    return excl_map


def _build_launch_map(zecom):
    launch_map = {}
    if zecom.empty or "Article No" not in zecom.columns:
        return launch_map
    if "Launch Date" not in zecom.columns:
        return launch_map

    arts = zecom["Article No"].tolist()
    launch_dates = zecom["Launch Date"].tolist()

    for art, ld in zip(arts, launch_dates):
        art_norm = _normalise_article_no(art)
        if art_norm:
            if pd.notna(ld) and str(ld).strip() not in ("", "NaT", "nan"):
                try:
                    launch_map[art_norm] = str(pd.to_datetime(ld).date())
                except Exception:
                    launch_map[art_norm] = _safe_str(ld)
            else:
                launch_map[art_norm] = ""
    return launch_map


def _needs_buffer(mp_name):
    """Buffer -1 stock only for Lazada PH and TikTok MY."""
    return mp_name in ("Lazada PH", "TikTok MY")


# == Exclusion override ========================================================

def _apply_exclusion(article_no, tc_stock, excl_map, max_0, sku=None):
    """
    Check exclusion by Article No first, then by raw SKU (13-digit)
    as a fallback — so exclusion works even without a Content File
    bridging SKU -> Article No.
    """
    match_key = None
    if article_no and article_no in excl_map:
        match_key = article_no
    elif sku and sku in excl_map:
        match_key = sku

    if match_key is None:
        return None
    excl_status = excl_map[match_key]
    if excl_status == "Inactive":
        # If Max 0 is already Yes, no need to set max 0 again
        max_action = "" if max_0 == "Yes" else "Set max 0"
        return ("Inactive", "Inactive as per AM Request", max_action)
    if excl_status == "Active":
        if tc_stock >= 1:
            ma = "Remove max" if max_0 == "Yes" else ""
            return ("Active", "Active as per AM Request", ma)
        else:
            ma = "Remove max" if max_0 == "Yes" else ""
            return ("Inactive", "AM Request Active but 0 Stock", ma)
    return None


# == SKU-level logic ==========================================================-

def _sku_logic(mp_status, mp_stock, ecom_status, tc_status,
               tc_stock, reserved, max_0, article_no, excl_map, sku=None):
    """
    ecom_status here is the normalised value (Inactive for future launch too).
    """
    excl = _apply_exclusion(article_no, tc_stock, excl_map, max_0, sku=sku)
    if excl:
        final_status, comment, max_action = excl
    else:
        if ecom_status == "Inactive":
            final_status = "Inactive"
            comment      = "Due to Ecom No"
        elif tc_stock == 0:
            final_status = "Inactive"
            comment      = "Due to 0 Stock"
        else:
            final_status = "Active"
            comment      = "Ecom Yes with Stock"

        max_action = ""
        if comment == "Due to Ecom No" and max_0 == "No":
            max_action = "Set max 0"
        elif comment in ("Due to 0 Stock", "Ecom Yes with Stock") and max_0 == "Yes":
            max_action = "Remove max"

    mp_norm  = _normalise_status(mp_status)
    tc_norm  = _normalise_status(tc_status)
    fin_norm = final_status

    final_check = (mp_norm == tc_norm == fin_norm)
    stock_check = (mp_stock == tc_stock)

    if not final_check:
        remarks = "Change to Active" if final_status == "Active" else "Change to Inactive"
    elif not stock_check:
        if final_status == "Active":
            remarks = "Due to Reserved Stock" if reserved != 0 else "Make Impact"
        else:
            remarks = "Stock not pushed due to Inactive Status"
    else:
        remarks = "All Good"

    push_0 = "Yes" if (tc_stock <= 0 and mp_stock > 0) else ""

    return {
        "Final Status":  final_status,
        "Comments":      comment,
        "Final Check":   str(final_check),
        "Stock Check":   str(stock_check),
        "Remarks":       remarks,
        "Max Setup":     max_action,
        "Update 0":      push_0,
    }


# == SKU-level validation (Lazada + Zalora) ====================================

def run_sku_validation(data, country):
    content   = data.get("content",   pd.DataFrame())
    tc_inv    = data.get("tc_inv",    pd.DataFrame())
    zecom     = data.get("zecom",     pd.DataFrame())
    all_df    = data.get("all_file",  pd.DataFrame())
    exclusion = data.get("exclusion", pd.DataFrame())

    excl_map    = _build_excl_map(exclusion)
    article_map = _build_article_map(content)
    tc_map      = _build_tc_map(tc_inv)
    launch_map  = _build_launch_map(zecom)

    mp_sources = {
        "Lazada " + country: data.get("lazada", pd.DataFrame()),
        "Zalora " + country: data.get("zalora", pd.DataFrame()),
    }

    rows = []
    for mp_name, df in mp_sources.items():
        if df is None or df.empty or "SKU" not in df.columns:
            continue

        apply_buffer = _needs_buffer(mp_name)
        ecom_map  = _build_ecom_map(zecom, mp_name)
        stock_map = _build_stock_map(all_df, apply_buffer)

        records = df.to_dict('records')
        for r in records:
            sku        = _safe_str(r.get("SKU", ""))
            mp_status  = _safe_str(r.get("MP Status", "Unknown"))
            mp_stock   = _safe_num(r.get("MP Stock", 0))
            article_no = article_map.get(sku, "")
            ecom_st    = ecom_map.get(article_no, "Inactive") if article_no else "Inactive"
            tc_data    = tc_map.get(sku, {"TC SKU": "", "TC Status": "Unknown", "Max 0": "No"})
            sd         = stock_map.get(sku, {"TC Stock": 0.0, "Reserved Stock": 0.0})
            excl_lbl   = excl_map.get(article_no, "") or excl_map.get(sku, "")
            launch_dt  = launch_map.get(article_no, "") if article_no else ""

            # Invalid SKU check — must be exactly 13 digits
            if not _is_valid_sku(sku):
                rows.append({
                    "Marketplace":    mp_name,
                    "Seller SKU":     sku,
                    "TC SKU":         tc_data["TC SKU"] if tc_data["TC SKU"] else "#N/A",
                    "Article No":     article_no if article_no else "#N/A",
                    "MP Status":      mp_status if mp_status else "#N/A",
                    "TC Status":      "#N/A",
                    "e-com (Yes/No)": "#N/A",
                    "Launch Date":    launch_dt if launch_dt else "#N/A",
                    "Exclusion":      excl_lbl if excl_lbl else "#N/A",
                    "ECOM Status":    "#N/A",
                    "MP Stock":       mp_stock,
                    "TC Stock":       "#N/A",
                    "Reserved Stock": "#N/A",
                    "Max 0":          "#N/A",
                    "Final Status":   "Invalid",
                    "Comments":       "Invalid SKU",
                    "Final Check":    "False",
                    "Stock Check":    "False",
                    "Remarks":        "Invalid SKU",
                    "Max Setup":      "#N/A",
                    "Update 0":       "#N/A",
                })
                continue

            # Normalise ecom for logic: future launch = Inactive for logic
            ecom_for_logic = "Inactive" if ecom_st.startswith("Inactive") else ecom_st

            result = _sku_logic(
                mp_status=mp_status,
                mp_stock=mp_stock,
                ecom_status=ecom_for_logic,
                tc_status=tc_data["TC Status"],
                tc_stock=sd["TC Stock"],
                reserved=sd["Reserved Stock"],
                max_0=tc_data["Max 0"],
                article_no=article_no,
                excl_map=excl_map,
                sku=sku,
            )
            rows.append({
                "Marketplace":    mp_name,
                "Seller SKU":     sku,
                "TC SKU":         tc_data["TC SKU"],
                "Article No":     article_no,
                "MP Status":      mp_status,
                "TC Status":      _normalise_status(tc_data["TC Status"]),
                "e-com (Yes/No)": "Yes" if ecom_st == "Active" else "No",
                "Launch Date":    launch_dt,
                "Exclusion":      excl_lbl,
                "ECOM Status":    ecom_st,
                "MP Stock":       mp_stock,
                "TC Stock":       sd["TC Stock"],
                "Reserved Stock": sd["Reserved Stock"],
                "Max 0":          tc_data["Max 0"],
                **result,
            })

    return pd.DataFrame(rows) if rows else pd.DataFrame()


# == PID-level logic (Shopee + TikTok) ========================================

def run_pid_validation(data, country):
    """
    Output columns:
    Marketplace, SellerSku, TC SKU, Product ID, Article No, MP Status,
    TC Status, e-com (Yes/No), Launch Date, Exclusion, ECOM Status,
    Final Status, Comments, Final Check, Dual Status, Consolidated SUM QTY,
    MP Stock, TC Stock, Reserved Stock, Max 0, Stock Check,
    Remarks, Max Setup, Update 0
    """
    content   = data.get("content",   pd.DataFrame())
    tc_inv    = data.get("tc_inv",    pd.DataFrame())
    zecom     = data.get("zecom",     pd.DataFrame())
    all_df    = data.get("all_file",  pd.DataFrame())
    exclusion = data.get("exclusion", pd.DataFrame())

    excl_map    = _build_excl_map(exclusion)
    article_map = _build_article_map(content)
    tc_map      = _build_tc_map(tc_inv)
    launch_map  = _build_launch_map(zecom)

    mp_sources = {
        "Shopee " + country: data.get("shopee", pd.DataFrame()),
    }
    if country == "MY":
        mp_sources["TikTok MY"] = data.get("tiktok", pd.DataFrame())

    rows = []

    for mp_name, df in mp_sources.items():
        if df is None or df.empty or "SKU" not in df.columns:
            continue

        apply_buffer = _needs_buffer(mp_name)
        ecom_map  = _build_ecom_map(zecom, mp_name)
        stock_map = _build_stock_map(all_df, apply_buffer)

        # == Step 1: Enrich each SKU row ==================================-
        enriched = []
        records = df.to_dict('records')
        for r in records:
            sku       = _safe_str(r.get("SKU", ""))
            pid       = _safe_str(r.get("Product ID", sku))
            mp_status = _safe_str(r.get("MP Status", "Unknown"))
            mp_stock  = _safe_num(r.get("MP Stock", 0))
            art       = article_map.get(sku, "")
            ecom_st   = ecom_map.get(art, "Inactive") if art else "Inactive"
            # Normalise ecom for logic
            ecom_logic = "Inactive" if ecom_st.startswith("Inactive") else ecom_st
            td        = tc_map.get(sku, {"TC SKU": "", "TC Status": "Unknown", "Max 0": "No"})
            sd        = stock_map.get(sku, {"TC Stock": 0.0, "Reserved Stock": 0.0})
            excl_lbl  = excl_map.get(art, "") or excl_map.get(sku, "")
            launch_dt = launch_map.get(art, "") if art else ""
            sku_valid = _is_valid_sku(sku)
            enriched.append({
                "SKU":            sku,
                "Product ID":     pid,
                "MP Status":      mp_status,
                "MP Stock":       mp_stock,
                "Article No":     art,
                "Ecom Status":    ecom_st,
                "Ecom Logic":     ecom_logic,
                "TC SKU":         td["TC SKU"],
                "TC Status":      td["TC Status"],
                "Max 0":          td["Max 0"],
                "TC Stock":       sd["TC Stock"],
                "Reserved Stock": sd["Reserved Stock"],
                "Exclusion":      excl_lbl,
                "Launch Date":    launch_dt,
                "SKU Valid":      sku_valid,
            })

        enriched_df = pd.DataFrame(enriched)
        if enriched_df.empty:
            continue

        # == Step 2: Dual Status per Product ID ==========================-
        dual_map = {}
        for pid, grp in enriched_df.groupby("Product ID", dropna=False):
            statuses = set(grp["Ecom Logic"].unique())
            dual_map[_safe_str(pid)] = (
                2 if ("Active" in statuses and "Inactive" in statuses) else 1
            )

        # == Step 3: Consolidated TC Stock per Product ID ==================
        consolidated_map = (
            enriched_df.groupby("Product ID")["TC Stock"].sum().to_dict()
        )

        # == Step 4: Per-SKU output row ====================================
        enriched_records = enriched_df.to_dict('records')
        for r in enriched_records:
            sku         = r["SKU"]
            pid         = r["Product ID"]
            mp_status   = r["MP Status"]
            mp_stock    = r["MP Stock"]
            article_no  = r["Article No"]
            ecom_st     = r["Ecom Status"]
            ecom_logic  = r["Ecom Logic"]
            tc_sku      = r["TC SKU"]
            tc_status   = r["TC Status"]
            max_0       = r["Max 0"]
            tc_stock    = r["TC Stock"]
            reserved    = r["Reserved Stock"]
            excl_lbl    = r["Exclusion"]
            launch_dt   = r["Launch Date"]
            sku_valid   = r["SKU Valid"]

            dual_status     = dual_map.get(pid, 1)
            consolidated_tc = consolidated_map.get(pid, 0.0)
            ecom_yn         = "Yes" if ecom_st == "Active" else "No"

            # Invalid SKU row
            if not sku_valid:
                rows.append({
                    "Marketplace":          mp_name,
                    "SellerSku":            sku,
                    "TC SKU":               tc_sku if tc_sku else "#N/A",
                    "Product ID":           pid if pid else "#N/A",
                    "Article No":           article_no if article_no else "#N/A",
                    "MP Status":            mp_status if mp_status else "#N/A",
                    "TC Status":            "#N/A",
                    "e-com (Yes/No)":       "#N/A",
                    "Launch Date":          launch_dt if launch_dt else "#N/A",
                    "Exclusion":            excl_lbl if excl_lbl else "#N/A",
                    "ECOM Status":          "#N/A",
                    "Final Status":         "Invalid",
                    "Comments":             "Invalid SKU",
                    "Final Check":          "False",
                    "Dual Status":          dual_status,
                    "Consolidated SUM QTY": consolidated_tc,
                    "MP Stock":             mp_stock,
                    "TC Stock":             "#N/A",
                    "Reserved Stock":       "#N/A",
                    "Max 0":                "#N/A",
                    "Stock Check":          "False",
                    "Remarks":              "Invalid SKU",
                    "Max Setup":            "#N/A",
                    "Update 0":             "#N/A",
                })
                continue

            # == Exclusion override ========================================
            excl = _apply_exclusion(article_no, consolidated_tc, excl_map, max_0, sku=sku)
            if excl:
                final_status, comment, max_action = excl
            else:
                # == Dual Status = 1 ======================================-
                if dual_status == 1:
                    if ecom_logic == "Inactive":
                        final_status = "Inactive"
                        comment      = "Due to Ecom No"
                    elif consolidated_tc == 0:
                        final_status = "Inactive"
                        comment      = "Due to 0 Stock"
                    else:
                        final_status = "Active"
                        comment      = "Ecom Yes with Stock"
                # == Dual Status = 2 ======================================-
                else:
                    if consolidated_tc == 0:
                        final_status = "Inactive"
                        comment      = "Due to 0 Stock"
                    elif ecom_logic == "Active":
                        final_status = "Active"
                        comment      = "Ecom Yes with Stock"
                    else:
                        final_status = "Active"
                        comment      = "Set max"

                # == Max Setup logic ======================================-
                max_action = ""
                if comment in ("Due to Ecom No", "Set max") and max_0 == "No":
                    max_action = "Set max"
                elif comment == "Ecom Yes with Stock" and max_0 == "Yes":
                    max_action = "Remove max"
                elif comment == "Due to 0 Stock":
                    if ecom_yn == "Yes" and max_0 == "Yes":
                        max_action = "Remove max"
                    elif ecom_yn in ("No", "") and max_0 == "No":
                        max_action = "Set max"

            # == Final Check & Stock Check ================================-
            mp_norm  = _normalise_status(mp_status)
            tc_norm  = _normalise_status(tc_status)
            fin_norm = final_status

            final_check = (mp_norm == tc_norm == fin_norm)
            stock_check = (mp_stock == tc_stock)

            # == Remarks ==================================================-
            if not final_check:
                remarks = "Update status to " + final_status
            elif not stock_check:
                if final_status == "Active":
                    if comment == "Set max":
                        remarks = "Set max product"
                    elif reserved != 0:
                        remarks = "Due to Reserved Stock"
                    else:
                        remarks = "Make Impact"
                else:
                    remarks = "Stock not pushed due to Inactive Status"
            else:
                remarks = "All Good"

            push_0 = "Yes" if (tc_stock <= 0 and mp_stock > 0) else ""

            rows.append({
                "Marketplace":          mp_name,
                "SellerSku":            sku,
                "TC SKU":               tc_sku,
                "Product ID":           pid,
                "Article No":           article_no,
                "MP Status":            mp_status,
                "TC Status":            _normalise_status(tc_status),
                "e-com (Yes/No)":       ecom_yn,
                "Launch Date":          launch_dt,
                "Exclusion":            excl_lbl,
                "ECOM Status":          ecom_st,
                "Final Status":         final_status,
                "Comments":             comment,
                "Final Check":          str(final_check),
                "Dual Status":          dual_status,
                "Consolidated SUM QTY": consolidated_tc,
                "MP Stock":             mp_stock,
                "TC Stock":             tc_stock,
                "Reserved Stock":       reserved,
                "Max 0":                max_0,
                "Stock Check":          str(stock_check),
                "Remarks":              remarks,
                "Max Setup":            max_action,
                "Update 0":             push_0,
            })

    return pd.DataFrame(rows) if rows else pd.DataFrame()
