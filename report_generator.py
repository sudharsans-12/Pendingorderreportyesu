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


def _normalise_article_no(val):
    """Standardise Article No separators/case for cross-file matching."""
    s = _safe_str(val)
    if not s:
        return ""
    s = s.strip().upper()
    s = re.sub(r'[\s\-]+', '_', s)
    s = s.strip('_')
    return s


def _build_article_map(content):
    article_map = {}
    if content.empty or "SKU" not in content.columns:
        return article_map
    art_col = next(
        (c for c in ["Article No", "Color_No", "Color_No.1", "ArticleNo"]
         if c in content.columns),
        next((c for c in content.columns
              if "article" in c.lower() or "color" in c.lower()), "")
    )
    if art_col:
        for _, r in content.iterrows():
            sku = _safe_str(r.get("SKU", ""))
            if sku:
                article_map[sku] = _normalise_article_no(r.get(art_col, ""))
    return article_map


def _build_ecom_map(zecom, mp_name):
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
    for _, r in zecom.iterrows():
        art = _normalise_article_no(r.get("Article No", ""))
        if art:
            ecom_map[art] = _safe_str(r.get(ecom_col, "Inactive"))
    return ecom_map


def _build_tc_map(tc_inv):
    """
    Child SKU (GED1708197-01) takes priority over Parent (GED1708197).
    Returns TC SKU (raw first-column value) along with TC Status and Max 0.
    """
    tc_map = {}
    parent_fallback = {}
    if tc_inv.empty or "SKU" not in tc_inv.columns:
        return tc_map
    for _, r in tc_inv.iterrows():
        sku = _safe_str(r.get("SKU", ""))
        if not sku:
            continue
        entry = {
            "TC SKU":    _safe_str(r.get("TC SKU", sku)),
            "TC Status": _safe_str(r.get("TC Status", "Unknown")),
            "Max 0":     _safe_str(r.get("Max 0", "No")),
        }
        if "-" in sku:
            tc_map[sku] = entry
            parent_base = sku.rsplit("-", 1)[0]
            if parent_base not in tc_map:
                parent_fallback[parent_base] = entry
        else:
            if sku not in tc_map:
                tc_map[sku] = entry
    for parent, entry in parent_fallback.items():
        if parent not in tc_map:
            tc_map[parent] = entry
    return tc_map


def _build_stock_map(all_df, apply_buffer=False):
    """Negative TC Stock -> 0. Buffer -1 for Lazada PH / TikTok MY."""
    stock_map = {}
    if all_df.empty or "SKU" not in all_df.columns:
        return stock_map
    for _, r in all_df.iterrows():
        sku = _safe_str(r.get("SKU", ""))
        if sku and sku not in stock_map:
            tc = _safe_num(r.get("TC Stock", 0))
            tc = max(tc, 0)
            if apply_buffer:
                tc = max(tc - 1, 0)
            stock_map[sku] = {
                "TC Stock":       tc,
                "Reserved Stock": _safe_num(r.get("Reserved Stock", 0)),
            }
    return stock_map


def _build_excl_map(exclusion):
    excl_map = {}
    if exclusion is None or exclusion.empty:
        return excl_map
    if "Article No" not in exclusion.columns:
        return excl_map
    for _, r in exclusion.iterrows():
        art = _normalise_article_no(r.get("Article No", ""))
        if art:
            excl_map[art] = _safe_str(r.get("Exclusion Status", "Inactive"))
    return excl_map


def _build_launch_map(zecom):
    """Article No -> Launch Date string (YYYY-MM-DD)."""
    launch_map = {}
    if zecom.empty or "Article No" not in zecom.columns:
        return launch_map
    if "Launch Date" not in zecom.columns:
        return launch_map
    for _, r in zecom.iterrows():
        art = _normalise_article_no(r.get("Article No", ""))
        if art:
            ld = r.get("Launch Date", "")
            if pd.notna(ld) and str(ld).strip() not in ("", "NaT", "nan"):
                try:
                    launch_map[art] = str(pd.to_datetime(ld).date())
                except Exception:
                    launch_map[art] = _safe_str(ld)
            else:
                launch_map[art] = ""
    return launch_map


def _needs_buffer(mp_name):
    return mp_name in ("Lazada PH", "TikTok MY")


def generate_status_report(data, country):
    all_df    = data.get("all_file",  pd.DataFrame())
    tc_inv    = data.get("tc_inv",    pd.DataFrame())
    content   = data.get("content",   pd.DataFrame())
    zecom     = data.get("zecom",     pd.DataFrame())
    exclusion = data.get("exclusion", pd.DataFrame())

    article_map = _build_article_map(content)
    excl_map    = _build_excl_map(exclusion)
    tc_map      = _build_tc_map(tc_inv)
    launch_map  = _build_launch_map(zecom)

    # Build mp_sources dynamically
    mp_sources = {}
    for key, label in [("lazada", "Lazada " + country),
                       ("shopee", "Shopee " + country),
                       ("zalora", "Zalora " + country)]:
        df = data.get(key, pd.DataFrame())
        if df is not None and not df.empty and "SKU" in df.columns:
            mp_sources[label] = df
    if country == "MY":
        df = data.get("tiktok", pd.DataFrame())
        if df is not None and not df.empty and "SKU" in df.columns:
            mp_sources["TikTok MY"] = df

    if not mp_sources:
        return pd.DataFrame()

    frames = []
    for mp_name, df in mp_sources.items():
        apply_buffer = _needs_buffer(mp_name)
        ecom_map  = _build_ecom_map(zecom, mp_name)
        stock_map = _build_stock_map(all_df, apply_buffer)

        for _, row in df.iterrows():
            sku = _safe_str(row.get("SKU", ""))
            if not sku:
                continue
            mp_status  = _safe_str(row.get("MP Status", "Unknown"))
            mp_stock   = _safe_num(row.get("MP Stock", 0))
            article_no = article_map.get(sku, "")
            ecom_st    = ecom_map.get(article_no, "Inactive") if article_no else "Inactive"
            launch_dt  = launch_map.get(article_no, "") if article_no else ""
            tc_data    = tc_map.get(sku, {"TC SKU": "", "TC Status": "Unknown", "Max 0": "No"})
            sd         = stock_map.get(sku, {"TC Stock": 0.0, "Reserved Stock": 0.0})
            excl_lbl   = excl_map.get(article_no, "") if article_no else ""

            frames.append({
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
            })

    return pd.DataFrame(frames) if frames else pd.DataFrame()
