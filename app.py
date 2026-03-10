import re
import json
from io import BytesIO
from datetime import date

import numpy as np
import pandas as pd
import requests
import streamlit as st

try:
    from dateutil.relativedelta import relativedelta
except Exception:
    relativedelta = None


APP_TITLE = "Account Management"
APP_SUBTITLE = "Upsell and Churn KPI Calculation"

KC_LIGHT_PINKISH_PURPLE = "#F4B7FF"
KC_VIBRANT_MAGENTA = "#EA66FF"
KC_BRIGHT_VIOLET = "#8548FF"
KC_DEEP_PURPLE = "#8D34F0"
KC_TEXT = "#15151A"

EMAIL_RE = re.compile(r"([a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,})", re.IGNORECASE)

ALLOWED_ANNUAL_BREAKDOWN_NUMS = [
    "9360", "14400", "11520", "12960", "10080", "10800", "38400", "34560", "30720", "26880"
]

ANNUAL_AMOUNT_THRESHOLD = 40.0
ANNUAL_UPGRADE_DESC_EXACT = "upgrade plan to yearly subscription."


# -------------------------
# Secrets and Auth
# -------------------------
def _require_secrets():
    missing = []

    if "mixpanel" not in st.secrets:
        missing.append("mixpanel")
    else:
        for k in ["project_id", "auth_header", "from_date"]:
            if k not in st.secrets["mixpanel"]:
                missing.append(f"mixpanel.{k}")

    if "auth" not in st.secrets:
        missing.append("auth")
    else:
        for k in ["username", "password"]:
            if k not in st.secrets["auth"]:
                missing.append(f"auth.{k}")

    if missing:
        st.error(
            "Missing required secrets. Add these keys in .streamlit/secrets.toml.\n\n"
            + "\n".join([f"- {m}" for m in missing])
        )
        st.stop()


def login_gate() -> bool:
    if st.session_state.get("authenticated"):
        return True

    st.markdown(
        """
        <div style="padding:14px;border-radius:16px;border:1px solid rgba(0,0,0,0.08);
                    background:linear-gradient(90deg, rgba(141,52,240,0.10), rgba(234,102,255,0.10), rgba(133,72,255,0.10));">
          <div style="font-size:1.1rem;font-weight:900;">Login</div>
          <div style="opacity:0.8;margin-top:4px;">Enter credentials to access the dashboard.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    u = st.text_input("Username", value="", key="login_user")
    p = st.text_input("Password", value="", type="password", key="login_pass")

    if st.button("Sign in", type="primary"):
        if u == str(st.secrets["auth"]["username"]) and p == str(st.secrets["auth"]["password"]):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Invalid username or password.")

    return False


# -------------------------
# Branding
# -------------------------
def inject_brand_css():
    css = f"""
    <style>
      .stApp {{
        color: {KC_TEXT};
      }}
      section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, rgba(133,72,255,0.10), rgba(244,183,255,0.10));
        border-right: 1px solid rgba(21,21,26,0.06);
      }}
      div.stButton > button {{
        border-radius: 12px !important;
        border: 0 !important;
        background: linear-gradient(90deg, {KC_DEEP_PURPLE}, {KC_BRIGHT_VIOLET}) !important;
        color: white !important;
        padding: 0.6rem 0.9rem !important;
        font-weight: 800 !important;
      }}
      div.stDownloadButton > button {{
        border-radius: 12px !important;
        border: 1px solid rgba(21,21,26,0.12) !important;
        background: white !important;
        color: {KC_TEXT} !important;
        font-weight: 800 !important;
      }}
      div[data-testid="stDataFrame"] {{
        border-radius: 16px;
        overflow: hidden;
        border: 1px solid rgba(21,21,26,0.08);
      }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


def render_header():
    c1, c2 = st.columns([1, 6])
    with c1:
        try:
            st.image("assets/KrispCallLogo.png", use_container_width=True)
        except Exception:
            pass
    with c2:
        st.markdown(
            f"""
            <div style="display:flex;gap:14px;align-items:center;padding:14px 16px;border-radius:16px;
                        background:linear-gradient(90deg, rgba(141,52,240,0.10), rgba(234,102,255,0.10), rgba(133,72,255,0.10));
                        border:1px solid rgba(21,21,26,0.06);">
              <div>
                <div style="font-size:1.25rem;font-weight:900;line-height:1.2;">{APP_TITLE}</div>
                <div style="opacity:0.85;margin-top:2px;">{APP_SUBTITLE}</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )


# -------------------------
# Utilities
# -------------------------
def _normalize_email(v):
    if v is None:
        return None
    s = str(v).strip().lower()
    if not s or s == "nan":
        return None
    return s


def _extract_email_from_text(txt):
    if txt is None:
        return None
    m = EMAIL_RE.search(str(txt))
    return m.group(1).lower() if m else None


def _prop_any(props: dict, keys: list[str]):
    for k in keys:
        if k in props and props.get(k) not in [None, ""]:
            return props.get(k)
    return None


def _time_to_epoch_seconds(v):
    if v is None:
        return None
    try:
        t = int(float(v))
        if t > 10**11:
            t //= 1000
        return t
    except Exception:
        dt = pd.to_datetime(v, errors="coerce", utc=True)
        if pd.isna(dt):
            return None
        return int(dt.value // 10**9)


def _epoch_to_dt_naive(series: pd.Series) -> pd.Series:
    t = pd.to_numeric(series, errors="coerce")
    if t.notna().all():
        if float(t.median()) > 1e11:
            t = (t // 1000)
        dt = pd.to_datetime(t, unit="s", utc=True, errors="coerce")
    else:
        dt = pd.to_datetime(series, errors="coerce", utc=True)
    return dt.dt.tz_convert(None)


def _row_is_candidate(desc_lower: str) -> bool:
    if not desc_lower:
        return False
    if "workspace subscription" in desc_lower:
        return True
    if "upgrade plan to yearly subscription" in desc_lower:
        return True
    if "number purchased" in desc_lower:
        return True
    if "agent added" in desc_lower and "," in desc_lower:
        return True
    if "number renew" in desc_lower:
        return True
    if EMAIL_RE.search(desc_lower):
        return True
    return False


def _connected_from_label(label) -> bool:
    if label is None:
        return False
    s = str(label).strip().lower()
    if not s or s == "nan":
        return False
    return ("connected" in s) and ("not connected" not in s)


def _connected_from_reach_status_or_label(status, label) -> bool:
    status_str = "" if status is None else str(status).strip().lower()

    if not status_str or status_str == "nan":
        return _connected_from_label(label)

    false_exact_values = {
        "not connected",
        "not answered",
        "voicemail",
        "hung up",
    }

    if status_str in false_exact_values:
        return False

    if "not connected" in status_str:
        return False

    return "connected" in status_str


def _tier_from_label(label):
    if label is None:
        return None
    s = str(label).lower()
    for k in ["bronze", "silver", "gold", "platinum", "vip"]:
        if k in s:
            return k.title()
    return None


def _add_month(ts):
    if pd.isna(ts):
        return pd.NaT
    ts = pd.Timestamp(ts)
    if relativedelta:
        return pd.Timestamp(ts.to_pydatetime() + relativedelta(months=1))
    return ts + pd.DateOffset(months=1)


def _add_year(ts):
    if pd.isna(ts):
        return pd.NaT
    ts = pd.Timestamp(ts)
    if relativedelta:
        return pd.Timestamp(ts.to_pydatetime() + relativedelta(years=1))
    return ts + pd.DateOffset(years=1)


def _clean_breakdown_str(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    s = str(x)
    s = s.replace("\u00a0", "")
    s = re.sub(r"\s+", "", s)
    return s


def _dt_to_date_only(x):
    if pd.isna(x):
        return pd.NaT
    try:
        return pd.Timestamp(x).date()
    except Exception:
        return pd.NaT


# -------------------------
# Mixpanel fetch. Reduced columns
# -------------------------
@st.cache_data(show_spinner=False, ttl=60 * 60, max_entries=3)
def fetch_mixpanel_npm(to_date: date):
    mp = st.secrets["mixpanel"]
    project_id = str(mp["project_id"]).strip()
    from_date = str(mp["from_date"]).strip()
    to_date_str = to_date.isoformat()

    events = ["New Payment Made"]
    event_array_json = json.dumps(events)

    base_url = mp.get("base_url", "https://data-eu.mixpanel.com")
    url = (
        f"{base_url}/api/2.0/export"
        f"?project_id={project_id}"
        f"&from_date={from_date}"
        f"&to_date={to_date_str}"
        f"&event={event_array_json}"
    )

    headers = {"accept": "text/plain", "authorization": str(mp["auth_header"]).strip()}

    kept = {}
    stats = {
        "lines_read": 0,
        "rows_kept_prefilter": 0,
        "rows_dedup_final": 0,
        "dupes_replaced": 0,
        "dupes_skipped": 0,
    }

    with requests.get(url, headers=headers, stream=True, timeout=240) as r:
        if r.status_code != 200:
            raise RuntimeError(f"Mixpanel export failed. Status {r.status_code}. Body: {r.text[:500]}")

        for line in r.iter_lines(decode_unicode=True):
            if not line:
                continue
            stats["lines_read"] += 1

            obj = json.loads(line)
            props = obj.get("properties") or {}

            rec = {
                "event": obj.get("event"),
                "distinct_id": props.get("distinct_id"),
                "time": props.get("time"),
                "$insert_id": props.get("$insert_id"),
                "mp_processing_time_ms": props.get("mp_processing_time_ms"),
                "$email": props.get("$email"),
                "Amount": props.get("Amount"),
                "Amount Description": props.get("Amount Description"),
                "Amount Breakdown": _prop_any(props, ["Amount Breakdown", "Amount breakdown", "AmountBreakdown"]),
                "Amount Breakdown by Unit": _prop_any(props, ["Amount Breakdown by Unit", "Amount breakdown by unit"]),
            }

            desc = rec.get("Amount Description")
            desc_lower = str(desc).lower().strip() if desc is not None else ""
            if not _row_is_candidate(desc_lower):
                continue

            stats["rows_kept_prefilter"] += 1

            event = rec.get("event")
            distinct_id = rec.get("distinct_id")
            insert_id = rec.get("$insert_id")
            time_s = _time_to_epoch_seconds(rec.get("time"))
            if event is None or distinct_id is None or insert_id is None or time_s is None:
                continue

            key = (event, distinct_id, time_s, insert_id)

            mp_pt = rec.get("mp_processing_time_ms")
            try:
                mp_pt_num = int(float(mp_pt)) if mp_pt is not None else None
            except Exception:
                mp_pt_num = None

            if key not in kept:
                kept[key] = (mp_pt_num, rec)
            else:
                old_mp_pt, _ = kept[key]
                if old_mp_pt is None and mp_pt_num is None:
                    kept[key] = (mp_pt_num, rec)
                    stats["dupes_replaced"] += 1
                elif old_mp_pt is None and mp_pt_num is not None:
                    kept[key] = (mp_pt_num, rec)
                    stats["dupes_replaced"] += 1
                elif old_mp_pt is not None and mp_pt_num is None:
                    stats["dupes_skipped"] += 1
                else:
                    if mp_pt_num >= old_mp_pt:
                        kept[key] = (mp_pt_num, rec)
                        stats["dupes_replaced"] += 1
                    else:
                        stats["dupes_skipped"] += 1

    rows = [v[1] for v in kept.values()]
    df = pd.DataFrame(rows) if rows else pd.DataFrame()
    stats["rows_dedup_final"] = len(df)
    return df, stats


# -------------------------
# Summary tables
# -------------------------
def _agg_metrics(g: pd.DataFrame) -> pd.Series:
    total = len(g)
    churn = int(g["Churned (AsOf MonthEnd)"].fillna(True).sum())
    churn_pct = (churn / total) if total else np.nan
    upsell_sum = float(g["Upsell Net Change"].fillna(0).sum())
    upsell_pos_sum = float(g["Upsell Positive Only"].fillna(0).sum())
    annual_active = int(g["Annual Active (AsOf MonthEnd)"].fillna(False).sum())
    return pd.Series(
        {
            "Accounts": total,
            "Churned": churn,
            "Churn %": churn_pct,
            "Annual Active": annual_active,
            "Upsell Net Change Sum": upsell_sum,
            "Upsell Positive Only Sum": upsell_pos_sum,
        }
    )


def summarize_basic(deals_enriched: pd.DataFrame, connected_only: bool) -> pd.DataFrame:
    df = deals_enriched.copy()
    if connected_only:
        df = df[df["Connected"] == True].copy()

    df = df[df["DealMonth"].notna()].copy()
    if df.empty:
        return pd.DataFrame()

    overall = df.groupby("DealMonth", as_index=False).apply(_agg_metrics).reset_index(drop=True)
    overall["Scope"] = "Overall"
    overall["Deal Owner"] = "All"

    if "Deal - Owner" in df.columns:
        by_owner = df.groupby(["DealMonth", "Deal - Owner"], as_index=False).apply(_agg_metrics).reset_index(drop=True)
        by_owner = by_owner.rename(columns={"Deal - Owner": "Deal Owner"})
        by_owner["Scope"] = "By Owner"
    else:
        by_owner = pd.DataFrame()

    out = pd.concat([overall, by_owner], ignore_index=True)
    out = out.sort_values(["DealMonth", "Scope", "Deal Owner"], kind="mergesort")
    return out


def summarize_owner_connected_status(deals_enriched: pd.DataFrame) -> pd.DataFrame:
    df = deals_enriched.copy()
    df = df[df["DealMonth"].notna()].copy()
    if df.empty:
        return pd.DataFrame()

    if "Deal - Owner" not in df.columns:
        df["Deal - Owner"] = "Unknown"
    if "Deal - Status" not in df.columns:
        df["Deal - Status"] = pd.NA

    group_cols = ["DealMonth", "Deal - Owner", "Connected", "Deal - Status"]
    out = df.groupby(group_cols, as_index=False).apply(_agg_metrics).reset_index(drop=True)
    out = out.rename(columns={"Deal - Owner": "Deal Owner", "Deal - Status": "Deal Status"})
    out = out.sort_values(["DealMonth", "Deal Owner", "Connected", "Deal Status"], kind="mergesort")
    return out


# -------------------------
# Core enrichment
# -------------------------
def build_enriched_deals(deals_df: pd.DataFrame, npm_df: pd.DataFrame):
    deal_date_col = "Deal - Deal created on"
    deal_email_col = "Person - Email"
    deal_owner_col = "Deal - Owner"
    deal_label_col = "Deal - Label"
    deal_reach_status_col = "Deal - Reach Status"
    deal_value_col = "Deal - Deal value"

    required_cols = [deal_date_col, deal_email_col, deal_owner_col, deal_label_col, deal_reach_status_col]
    missing = [c for c in required_cols if c not in deals_df.columns]
    if missing:
        raise KeyError(f"Deals file missing columns: {missing}")

    deals = deals_df.copy()
    deals["_deal_created_dt"] = pd.to_datetime(deals[deal_date_col], errors="coerce", utc=True).dt.tz_convert(None)
    if deals["_deal_created_dt"].isna().all():
        deals["_deal_created_dt"] = pd.to_datetime(deals[deal_date_col], errors="coerce")

    deals["DealMonth"] = deals["_deal_created_dt"].dt.to_period("M").dt.to_timestamp()
    deals["EmailKey"] = deals[deal_email_col].map(_normalize_email)

    deals["Tier"] = deals[deal_label_col].map(_tier_from_label)

    deals["Connected"] = [
        _connected_from_reach_status_or_label(status, label)
        for status, label in zip(deals[deal_reach_status_col], deals[deal_label_col])
    ]

    deals["_is_pipedrive_krispcall"] = (
        deals[deal_owner_col].astype(str).str.strip().str.lower() == "pipedrive krispcall"
    ).astype(int)

    if deal_value_col in deals.columns:
        deals["_deal_value_num"] = pd.to_numeric(deals[deal_value_col], errors="coerce").fillna(0.0)
    else:
        deals["_deal_value_num"] = 0.0

    deals["_dedup_key"] = deals["EmailKey"].fillna("__missing_email__") + "|" + deals["DealMonth"].astype(str)
    grp_counts = deals.groupby("_dedup_key")["_dedup_key"].transform("count")
    deals["Dedup Group Count"] = grp_counts
    deals["Dedup Dropped Duplicates"] = grp_counts.gt(1)

    deals_sorted = deals.sort_values(
        by=["_dedup_key", "_is_pipedrive_krispcall", "_deal_value_num", "_deal_created_dt"],
        ascending=[True, True, False, False],
        kind="mergesort",
    )
    deals_dedup = deals_sorted.drop_duplicates(subset=["_dedup_key"], keep="first").copy()
    deals_dedup["PrevDealMonth"] = (deals_dedup["DealMonth"] - pd.offsets.MonthBegin(1)).dt.to_period("M").dt.to_timestamp()

    if npm_df is None or npm_df.empty:
        out = deals_dedup.drop(columns=["_dedup_key"], errors="ignore").copy()
        summary_all = summarize_basic(out, False)
        summary_connected = summarize_basic(out, True)
        summary_owner_cs = summarize_owner_connected_status(out)
        return out, summary_all, summary_connected, summary_owner_cs

    npm = npm_df.copy()
    npm = npm.loc[:, ~npm.columns.duplicated()].copy()

    for col in ["Amount Breakdown", "Amount Breakdown by Unit", "Amount Description", "$email"]:
        if col not in npm.columns:
            npm[col] = pd.NA
    if "time" not in npm.columns:
        raise KeyError("Mixpanel export missing required column: time")
    if "Amount" not in npm.columns:
        raise KeyError("Mixpanel export missing required column: Amount")

    npm["PayDT"] = _epoch_to_dt_naive(npm["time"])
    npm["PayMonth"] = npm["PayDT"].dt.to_period("M").dt.to_timestamp()
    npm["AmountNum"] = pd.to_numeric(npm["Amount"], errors="coerce")

    npm["EmailKey"] = npm["$email"].map(_normalize_email)
    npm["EmailKey"] = npm["EmailKey"].fillna(npm["Amount Description"].map(_extract_email_from_text))
    npm_valid = npm.dropna(subset=["EmailKey", "PayDT"]).copy()

    desc = npm_valid["Amount Description"].astype(str)
    desc_lower = desc.str.lower().str.strip()
    breakdown_clean = npm_valid["Amount Breakdown"].apply(_clean_breakdown_str)

    annual_start = (npm_valid["AmountNum"].fillna(0) > ANNUAL_AMOUNT_THRESHOLD) & (
        desc_lower.str.contains("workspace subscription", na=False)
        | (desc_lower == ANNUAL_UPGRADE_DESC_EXACT)
    )

    desc_no_comma = ~desc.str.contains(",", na=False)
    contains_email = pd.Series(
        [(e in d) if (e and isinstance(d, str)) else False for e, d in zip(npm_valid["EmailKey"], desc_lower)],
        index=npm_valid.index,
    )

    breakdown_mask = breakdown_clean.apply(lambda x: any(n in x for n in ALLOWED_ANNUAL_BREAKDOWN_NUMS))
    annual_renew = (npm_valid["AmountNum"].fillna(0) > ANNUAL_AMOUNT_THRESHOLD) & (
        (contains_email & desc_no_comma) | breakdown_mask
    )

    annual_candidates = npm_valid[annual_start | annual_renew].copy()
    annual_candidates["Annual Payment Type"] = np.where(
        annual_start.loc[annual_candidates.index], "Subscription", "Renew"
    )
    annual_user_set = set(annual_candidates["EmailKey"].unique())

    cond_email_in_desc = contains_email
    cond_number_purchased = desc_lower.str.contains("number purchased", na=False)
    cond_agent_added = desc_lower.str.contains("agent added", na=False) & desc.str.contains(",", na=False)
    cond_workspace_sub = desc_lower.str.contains("workspace subscription", na=False)
    cond_number_renew_for_mapping = desc_lower.str.contains("number renew", na=False) & npm_valid["EmailKey"].isin(annual_user_set)

    renewal_mask = (
        cond_email_in_desc
        | cond_number_purchased
        | cond_agent_added
        | cond_workspace_sub
        | cond_number_renew_for_mapping
    )
    renewals = npm_valid[renewal_mask].copy()
    renewals_sorted = renewals.sort_values(["EmailKey", "PayMonth", "PayDT"], kind="mergesort")

    txn_count = (
        renewals_sorted.groupby(["EmailKey", "PayMonth"])
        .size()
        .rename("Renew Txn Count")
        .reset_index()
    )

    latest_rows = renewals_sorted.drop_duplicates(subset=["EmailKey", "PayMonth"], keep="last").copy()
    latest_rows = latest_rows.merge(txn_count, on=["EmailKey", "PayMonth"], how="left")
    latest_rows["Renew Multiple Flag"] = latest_rows["Renew Txn Count"].fillna(0).astype(int) > 1
    latest_map = latest_rows.set_index(["EmailKey", "PayMonth"])

    def _map_from_latest(email, month, col, default):
        if email is None or pd.isna(month):
            return default
        try:
            return latest_map.loc[(email, month), col]
        except Exception:
            return default

    deals_dedup["Current Month Renew Amount"] = [
        float(_map_from_latest(e, m, "AmountNum", np.nan)) if pd.notna(_map_from_latest(e, m, "AmountNum", np.nan)) else np.nan
        for e, m in zip(deals_dedup["EmailKey"], deals_dedup["DealMonth"])
    ]
    deals_dedup["Previous Month Renew Amount"] = [
        float(_map_from_latest(e, m, "AmountNum", np.nan)) if pd.notna(_map_from_latest(e, m, "AmountNum", np.nan)) else np.nan
        for e, m in zip(deals_dedup["EmailKey"], deals_dedup["PrevDealMonth"])
    ]

    deals_dedup["Current Month Renew Date"] = [
        _map_from_latest(e, m, "PayDT", pd.NaT)
        for e, m in zip(deals_dedup["EmailKey"], deals_dedup["DealMonth"])
    ]
    deals_dedup["Previous Month Renew Date"] = [
        _map_from_latest(e, m, "PayDT", pd.NaT)
        for e, m in zip(deals_dedup["EmailKey"], deals_dedup["PrevDealMonth"])
    ]
    deals_dedup["Current Month Renew Date"] = deals_dedup["Current Month Renew Date"].apply(_dt_to_date_only)
    deals_dedup["Previous Month Renew Date"] = deals_dedup["Previous Month Renew Date"].apply(_dt_to_date_only)

    deals_dedup["Current Month Renew Multiple Flag"] = [
        bool(_map_from_latest(e, m, "Renew Multiple Flag", False))
        for e, m in zip(deals_dedup["EmailKey"], deals_dedup["DealMonth"])
    ]
    deals_dedup["Previous Month Renew Multiple Flag"] = [
        bool(_map_from_latest(e, m, "Renew Multiple Flag", False))
        for e, m in zip(deals_dedup["EmailKey"], deals_dedup["PrevDealMonth"])
    ]

    deal_month_index = (
        deals_dedup[["EmailKey", "DealMonth"]]
        .dropna(subset=["EmailKey", "DealMonth"])
        .drop_duplicates()
        .sort_values(["EmailKey", "DealMonth"], kind="mergesort")
    )

    renewals_for_validity = renewals.copy()
    renewals_for_validity = renewals_for_validity[
        ~renewals_for_validity["Amount Description"].astype(str).str.contains("number renew", case=False, na=False)
    ].copy()

    valid_sorted = renewals_for_validity.sort_values(["EmailKey", "PayMonth", "PayDT"], kind="mergesort")
    valid_latest_in_month = valid_sorted.drop_duplicates(
        subset=["EmailKey", "PayMonth"], keep="last"
    )[["EmailKey", "PayMonth", "PayDT"]].copy()
    valid_latest_in_month = valid_latest_in_month.rename(columns={"PayMonth": "DealMonth", "PayDT": "Monthly PayDT In Month"})

    monthly_roll = (
        deal_month_index.merge(valid_latest_in_month, on=["EmailKey", "DealMonth"], how="left")
        .sort_values(["EmailKey", "DealMonth"], kind="mergesort")
    )
    monthly_roll["Latest Monthly PayDT (AsOf MonthEnd)"] = monthly_roll.groupby("EmailKey")["Monthly PayDT In Month"].ffill()

    deals_dedup = deals_dedup.merge(
        monthly_roll[["EmailKey", "DealMonth", "Latest Monthly PayDT (AsOf MonthEnd)"]],
        on=["EmailKey", "DealMonth"],
        how="left",
    )

    annual_simple = annual_candidates.dropna(subset=["EmailKey", "PayDT"]).copy()
    annual_simple = annual_simple.assign(PayMonth=annual_simple["PayDT"].dt.to_period("M").dt.to_timestamp())
    annual_simple = annual_simple.sort_values(["EmailKey", "PayMonth", "PayDT"], kind="mergesort")

    annual_latest_in_month = annual_simple.drop_duplicates(
        subset=["EmailKey", "PayMonth"], keep="last"
    )[["EmailKey", "PayMonth", "PayDT", "Annual Payment Type"]].copy()
    annual_latest_in_month = annual_latest_in_month.rename(columns={"PayMonth": "DealMonth", "PayDT": "Annual PayDT In Month"})

    annual_roll = (
        deal_month_index.merge(annual_latest_in_month, on=["EmailKey", "DealMonth"], how="left")
        .sort_values(["EmailKey", "DealMonth"], kind="mergesort")
    )
    annual_roll["Latest Annual PayDT (AsOf MonthEnd)"] = annual_roll.groupby("EmailKey")["Annual PayDT In Month"].ffill()
    annual_roll["Annual Payment Type (AsOf MonthEnd)"] = annual_roll.groupby("EmailKey")["Annual Payment Type"].ffill()

    deals_dedup = deals_dedup.merge(
        annual_roll[["EmailKey", "DealMonth", "Latest Annual PayDT (AsOf MonthEnd)", "Annual Payment Type (AsOf MonthEnd)"]],
        on=["EmailKey", "DealMonth"],
        how="left",
    )

    deals_dedup["NextMonthStart"] = (deals_dedup["DealMonth"] + pd.offsets.MonthBegin(1)).dt.normalize()

    deals_dedup["Monthly Valid Till (AsOf MonthEnd)"] = deals_dedup["Latest Monthly PayDT (AsOf MonthEnd)"].apply(_add_month)
    deals_dedup["Annual Valid Till (AsOf MonthEnd)"] = deals_dedup["Latest Annual PayDT (AsOf MonthEnd)"].apply(_add_year)

    deals_dedup["Subscription Valid Till (AsOf MonthEnd)"] = deals_dedup[
        ["Monthly Valid Till (AsOf MonthEnd)", "Annual Valid Till (AsOf MonthEnd)"]
    ].max(axis=1)

    deals_dedup["Annual Active (AsOf MonthEnd)"] = deals_dedup["Annual Valid Till (AsOf MonthEnd)"].notna() & (
        deals_dedup["Annual Valid Till (AsOf MonthEnd)"] >= deals_dedup["NextMonthStart"]
    )

    deals_dedup["Active Subscription (AsOf MonthEnd)"] = deals_dedup["Subscription Valid Till (AsOf MonthEnd)"].notna() & (
        deals_dedup["Subscription Valid Till (AsOf MonthEnd)"] >= deals_dedup["NextMonthStart"]
    )

    deals_dedup["Churned (AsOf MonthEnd)"] = ~deals_dedup["Active Subscription (AsOf MonthEnd)"]

    prev_amt = deals_dedup["Previous Month Renew Amount"].fillna(0.0)
    curr_amt = deals_dedup["Current Month Renew Amount"].fillna(0.0)

    is_annual_user_asof = deals_dedup["Latest Annual PayDT (AsOf MonthEnd)"].notna()
    eligible = (prev_amt > 0) & (~deals_dedup["Churned (AsOf MonthEnd)"]) & (~is_annual_user_asof)

    deals_dedup["Upsell Net Change"] = np.where(eligible, (curr_amt - prev_amt), 0.0)
    deals_dedup["Upsell Positive Only"] = np.where(deals_dedup["Upsell Net Change"] > 0, deals_dedup["Upsell Net Change"], 0.0)

    out = deals_dedup.drop(columns=["_dedup_key"], errors="ignore").copy()
    summary_all = summarize_basic(out, False)
    summary_connected = summarize_basic(out, True)
    summary_owner_cs = summarize_owner_connected_status(out)
    return out, summary_all, summary_connected, summary_owner_cs


def make_excel(
    deals_raw: pd.DataFrame,
    deals_enriched: pd.DataFrame,
    summary_all: pd.DataFrame,
    summary_connected: pd.DataFrame,
    summary_owner_cs: pd.DataFrame,
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        deals_enriched.to_excel(writer, sheet_name="Deals_enriched", index=False)
        summary_all.to_excel(writer, sheet_name="Summary_all", index=False)
        summary_connected.to_excel(writer, sheet_name="Summary_connected", index=False)
        summary_owner_cs.to_excel(writer, sheet_name="Summary_owner_connected_status", index=False)
        deals_raw.to_excel(writer, sheet_name="Deals_raw", index=False)
    return output.getvalue()


def kpi_row(summary_df: pd.DataFrame):
    if summary_df is None or summary_df.empty:
        st.info("No summary available.")
        return

    overall = summary_df[summary_df["Scope"] == "Overall"].copy()
    if overall.empty:
        st.info("No overall summary available.")
        return

    latest = overall.sort_values("DealMonth").iloc[-1]
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Accounts", int(latest["Accounts"]))
    c2.metric("Churned", int(latest["Churned"]))
    churn_pct = float(latest["Churn %"]) * 100 if pd.notna(latest["Churn %"]) else np.nan
    c3.metric("Churn %", f"{churn_pct:.2f}%" if pd.notna(churn_pct) else "NA")
    c4.metric("Annual Active", int(latest["Annual Active"]))
    c5.metric("Upsell Net Sum", f'{float(latest["Upsell Net Change Sum"]):,.2f}')
    c6.metric("Upsell Positive Only", f'{float(latest["Upsell Positive Only Sum"]):,.2f}')


def main():
    st.set_page_config(page_title="KrispCall | Account Management", page_icon="📞", layout="wide")
    _require_secrets()
    inject_brand_css()
    render_header()

    st.sidebar.image("assets/KrispCallLogo.png", use_container_width=True)
    st.sidebar.markdown("### Controls")

    if not login_gate():
        return

    end_date = st.sidebar.date_input("Payments to date", value=date.today())
    st.sidebar.caption("Mixpanel Export API. Start date comes from secrets mixpanel.from_date.")
    deals_file = st.sidebar.file_uploader("Upload Deals CSV", type=["csv"])
    fetch_btn = st.sidebar.button("Fetch payments", type="primary")
    focus_connected = st.sidebar.toggle("Connected only view", value=False)

    if deals_file is None:
        st.info("Upload your Deals CSV to begin.")
        return

    deals_raw = pd.read_csv(deals_file)

    if "npm_cached" not in st.session_state:
        st.session_state["npm_cached"] = None
    if "npm_stats" not in st.session_state:
        st.session_state["npm_stats"] = None

    if fetch_btn:
        with st.spinner("Fetching Mixpanel payments. Reduced columns, filtered, deduped..."):
            npm_df, stats = fetch_mixpanel_npm(end_date)
            st.session_state["npm_cached"] = npm_df
            st.session_state["npm_stats"] = stats
            st.success("Payments fetched.")

    npm_df = st.session_state.get("npm_cached")
    if npm_df is None:
        st.warning("Click Fetch payments to load Mixpanel events.")
        return

    stats = st.session_state.get("npm_stats") or {}
    st.sidebar.markdown("### Mixpanel fetch stats")
    st.sidebar.write(stats)

    with st.spinner("Building enriched dataset..."):
        deals_enriched, summary_all, summary_connected, summary_owner_cs = build_enriched_deals(deals_raw, npm_df)

    summary_view = summary_connected if focus_connected else summary_all
    kpi_row(summary_view)

    tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Visuals", "Deals enriched", "Payments preview"])

    with tab1:
        st.subheader("All deals")
        st.dataframe(summary_all, use_container_width=True)

        st.subheader("Connected only")
        st.dataframe(summary_connected, use_container_width=True)

        st.subheader("Owner + Connected + Deal Status breakdown")
        st.dataframe(summary_owner_cs, use_container_width=True)

    with tab2:
        overall = summary_view[summary_view["Scope"] == "Overall"].copy() if not summary_view.empty else pd.DataFrame()
        if overall.empty:
            st.info("No data to chart.")
        else:
            overall["DealMonth"] = pd.to_datetime(overall["DealMonth"], errors="coerce")
            overall = overall.sort_values("DealMonth")
            st.caption("Churn percentage over time")
            st.line_chart(overall.set_index("DealMonth")[["Churn %"]])
            st.caption("Churned accounts over time")
            st.line_chart(overall.set_index("DealMonth")[["Churned"]])
            st.caption("Upsell net change sum over time")
            st.line_chart(overall.set_index("DealMonth")[["Upsell Net Change Sum"]])
            st.caption("Upsell positive only sum over time")
            st.line_chart(overall.set_index("DealMonth")[["Upsell Positive Only Sum"]])

    with tab3:
        st.dataframe(deals_enriched, use_container_width=True)

    with tab4:
        st.caption("Reduced payments dataset after filter and dedupe. First 200 rows shown.")
        st.dataframe(npm_df.head(200), use_container_width=True)

    st.divider()
    st.subheader("Export")
    excel_bytes = make_excel(deals_raw, deals_enriched, summary_all, summary_connected, summary_owner_cs)
    st.download_button(
        "Download Excel workbook",
        data=excel_bytes,
        file_name="account_mgmt_upsell_churn_enriched.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
