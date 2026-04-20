import re
import json
from io import BytesIO
from datetime import date, timedelta

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
ANNUAL_SUBSCRIPTION_TEXT_MARKERS = ["workspace subscription", "advance,", "starter,", "stater,"]
EXCLUDED_UPSELL_DESCRIPTIONS = {"purchased credit", "credit purchased", "amount recharged"}


# -------------------------
# State
# -------------------------
def init_state():
    defaults = {
        "authenticated": False,
        "npm_cached": None,
        "npm_stats": None,
        "last_fetch_to_date": None,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


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
    if bool(st.session_state.get("authenticated", False)):
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

    if st.button("Sign in", type="primary", key="login_submit"):
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


def _prop_any(props, keys):
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


def _epoch_to_dt_naive(series):
    t = pd.to_numeric(series, errors="coerce")
    if t.notna().all():
        if float(t.median()) > 1e11:
            t = (t // 1000)
        dt = pd.to_datetime(t, unit="s", utc=True, errors="coerce")
    else:
        dt = pd.to_datetime(series, errors="coerce", utc=True)
    return dt.dt.tz_convert(None)


def _clean_breakdown_str(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    s = str(x)
    s = s.replace("\u00a0", "")
    s = s.replace(",", "")
    s = re.sub(r"\s+", "", s)
    return s


def _breakdown_mask_from_series(series):
    cleaned = series.apply(_clean_breakdown_str)
    mask = cleaned.apply(lambda x: any(n in x for n in ALLOWED_ANNUAL_BREAKDOWN_NUMS))
    return cleaned, mask


def _desc_matches_annual_subscription_terms(desc_lower):
    if not desc_lower:
        return False
    return any(marker in desc_lower for marker in ANNUAL_SUBSCRIPTION_TEXT_MARKERS)


def _annual_desc_series_mask(series):
    s = series.fillna("").astype(str).str.lower().str.strip()
    mask = pd.Series(False, index=s.index)
    for marker in ANNUAL_SUBSCRIPTION_TEXT_MARKERS:
        mask = mask | s.str.contains(re.escape(marker), na=False)
    return mask


def _excluded_credit_desc_series_mask(series):
    s = series.fillna("").astype(str).str.strip().str.lower()
    return s.isin(EXCLUDED_UPSELL_DESCRIPTIONS)


def _row_is_candidate(desc_lower, breakdown_clean=""):
    if not desc_lower:
        return False
    if _desc_matches_annual_subscription_terms(desc_lower):
        return True
    if "upgrade plan to yearly subscription" in desc_lower:
        return True
    if "number purchased" in desc_lower:
        return True
    if "agent added" in desc_lower:
        return True
    if "number renew" in desc_lower:
        return True
    if EMAIL_RE.search(desc_lower):
        return True
    if any(n in breakdown_clean for n in ALLOWED_ANNUAL_BREAKDOWN_NUMS):
        return True
    return False


def _connected_from_label(label):
    if label is None:
        return False
    s = str(label).strip().lower()
    if not s or s == "nan":
        return False
    return ("connected" in s) and ("not connected" not in s)


def _connected_from_reach_status_or_label(status, label):
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


def _dt_to_date_only(x):
    if pd.isna(x):
        return pd.NaT
    try:
        return pd.Timestamp(x).date()
    except Exception:
        return pd.NaT


def _safe_parse_date(v):
    try:
        return pd.to_datetime(v).date()
    except Exception:
        return None


def _month_start_from_date(v):
    if v is None:
        return pd.NaT
    return pd.Timestamp(v).to_period("M").to_timestamp()


def _prev_month_start(month_start):
    if pd.isna(month_start):
        return pd.NaT
    return (pd.Timestamp(month_start) - pd.offsets.MonthBegin(1)).to_period("M").to_timestamp()


def _aligned_prev_month_cutoff(report_month_start, churn_cutoff_date):
    if churn_cutoff_date is None or pd.isna(report_month_start):
        return None
    report_month_start = pd.Timestamp(report_month_start)
    prev_month_start = _prev_month_start(report_month_start)
    prev_month_end = prev_month_start + pd.offsets.MonthEnd(0)
    try:
        day = int(pd.Timestamp(churn_cutoff_date).day)
    except Exception:
        return None
    day = max(1, day)
    max_day = int(prev_month_end.day)
    aligned_day = min(day, max_day)
    return (prev_month_start + pd.Timedelta(days=aligned_day - 1)).date()


def _month_end_date(month_start):
    if pd.isna(month_start):
        return None
    return (pd.Timestamp(month_start) + pd.offsets.MonthEnd(0)).date()


def _to_date_safe(series):
    return pd.to_datetime(series, errors="coerce").dt.date


def _series_contains_email(email_series, desc_series_lower):
    out = []
    for e, d in zip(email_series, desc_series_lower):
        if e and isinstance(d, str):
            out.append(e in d)
        else:
            out.append(False)
    return pd.Series(out, index=desc_series_lower.index)


def _next_month_start(month_series):
    return (month_series + pd.offsets.MonthBegin(1)).dt.normalize()


def _month_end_asof_dt(month_series):
    return (
        month_series
        + pd.offsets.MonthEnd(0)
        + pd.Timedelta(days=1)
        - pd.Timedelta(microseconds=1)
    )


def _merge_latest_asof(deal_month_index, payments_df, payment_dt_col, extra_cols=None, prefix=""):
    base = (
        deal_month_index[["EmailKey", "DealMonth"]]
        .dropna(subset=["EmailKey", "DealMonth"])
        .drop_duplicates()
        .copy()
    )
    if base.empty:
        return base

    base["AsOfDT"] = _month_end_asof_dt(base["DealMonth"])
    pay_cols = ["EmailKey", payment_dt_col] + (extra_cols or [])
    pay = payments_df[pay_cols].dropna(subset=["EmailKey", payment_dt_col]).copy()

    out_dt_col = f"{prefix}{payment_dt_col}"
    out_extra_cols = {c: f"{prefix}{c}" for c in (extra_cols or [])}

    frames = []
    for email, base_g in base.groupby("EmailKey", sort=False, dropna=False):
        g = base_g.sort_values("AsOfDT", kind="mergesort").copy()

        pay_g = pay[pay["EmailKey"] == email].sort_values(payment_dt_col, kind="mergesort").reset_index(drop=True)

        g[out_dt_col] = pd.NaT
        for c, out_c in out_extra_cols.items():
            g[out_c] = pd.NA

        if not pay_g.empty:
            pay_times = pay_g[payment_dt_col].to_numpy(dtype="datetime64[ns]")
            asof_times = g["AsOfDT"].to_numpy(dtype="datetime64[ns]")
            idx = np.searchsorted(pay_times, asof_times, side="right") - 1
            valid = idx >= 0

            if valid.any():
                matched = pay_g.iloc[idx[valid]].reset_index(drop=True)
                g.loc[g.index[valid], out_dt_col] = matched[payment_dt_col].to_numpy()
                for c, out_c in out_extra_cols.items():
                    g.loc[g.index[valid], out_c] = matched[c].to_numpy()

        frames.append(g)

    merged = pd.concat(frames, ignore_index=True, sort=False)
    keep_cols = ["EmailKey", "DealMonth", out_dt_col] + list(out_extra_cols.values())
    return merged[keep_cols]


# -------------------------
# Mixpanel fetch
# -------------------------
@st.cache_data(show_spinner=False, ttl=60 * 60, max_entries=3)
def fetch_mixpanel_npm(to_date):
    mp = st.secrets["mixpanel"]
    project_id = str(mp["project_id"]).strip()

    configured_from_date = _safe_parse_date(str(mp["from_date"]).strip())
    min_annual_lookback_date = to_date - timedelta(days=400)

    if configured_from_date is None:
        effective_from_date = min_annual_lookback_date
    else:
        effective_from_date = min(configured_from_date, min_annual_lookback_date)

    from_date = effective_from_date.isoformat()
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
        "from_date_used": from_date,
        "to_date_used": to_date_str,
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

            amount_desc = props.get("Amount Description")
            amount_breakdown = _prop_any(props, ["Amount Breakdown", "Amount breakdown", "AmountBreakdown"])
            amount_breakdown_by_unit = _prop_any(props, ["Amount Breakdown by Unit", "Amount breakdown by unit"])
            breakdown_source = f"{'' if amount_breakdown is None else amount_breakdown} {' ' if amount_breakdown_by_unit is not None else ''}{'' if amount_breakdown_by_unit is None else amount_breakdown_by_unit}"
            breakdown_clean = _clean_breakdown_str(breakdown_source)
            desc_lower = str(amount_desc).lower().strip() if amount_desc is not None else ""

            stats["rows_kept_prefilter"] += 1

            rec = {
                "event": obj.get("event"),
                "distinct_id": props.get("distinct_id"),
                "time": props.get("time"),
                "$insert_id": props.get("$insert_id"),
                "mp_processing_time_ms": props.get("mp_processing_time_ms"),
                "$email": props.get("$email"),
                "Amount": props.get("Amount"),
                "Amount Description": amount_desc,
                "Amount Breakdown": amount_breakdown,
                "Amount Breakdown by Unit": amount_breakdown_by_unit,
            }

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
def _build_summary_metrics(df, group_cols, include_connected_pct=True):
    base_cols = [
        "Accounts", "Connected", "Churn Eligible Accounts", "Churn",
        "Upsell (Net)", "Upsell (Positive Only)", "Churn %"
    ]
    if include_connected_pct:
        base_cols.insert(2, "Connected %")

    if df.empty:
        cols = list(group_cols) + base_cols
        return pd.DataFrame(columns=cols)

    out = (
        df.groupby(group_cols, dropna=False, observed=False)
        .agg(
            Accounts=("EmailKey", "size"),
            Connected=("Connected", lambda s: int(pd.Series(s).fillna(False).sum())),
            **{
                "Churn Eligible Accounts": ("Churn Eligible", lambda s: int(pd.Series(s).fillna(False).sum())),
                "Churn": ("Churned (Reporting)", lambda s: int(pd.Series(s).fillna(False).sum())),
                "Upsell (Net)": ("Upsell Net Change", lambda s: float(pd.Series(s).fillna(0).sum())),
                "Upsell (Positive Only)": ("Upsell Positive Only", lambda s: float(pd.Series(s).fillna(0).sum())),
            }
        )
        .reset_index()
    )
    if include_connected_pct:
        out["Connected %"] = np.where(
            out["Accounts"] > 0,
            (out["Connected"] / out["Accounts"]) * 100,
            np.nan,
        )
    out["Churn %"] = np.where(
        out["Churn Eligible Accounts"] > 0,
        (out["Churn"] / out["Churn Eligible Accounts"]) * 100,
        np.nan,
    )
    return out


def summarize_overall(deals_enriched):
    df = deals_enriched.copy()
    df = df[df["DealMonth"].notna()].copy()
    return _build_summary_metrics(df, ["DealMonth"], include_connected_pct=True).sort_values(["DealMonth"], kind="mergesort")


def summarize_tier(deals_enriched):
    df = deals_enriched.copy()
    df = df[df["DealMonth"].notna()].copy()
    if "Tier" not in df.columns:
        df["Tier"] = pd.NA
    out = _build_summary_metrics(df, ["DealMonth", "Tier"], include_connected_pct=True)
    return out.sort_values(["DealMonth", "Tier"], kind="mergesort")


def summarize_owner(deals_enriched):
    df = deals_enriched.copy()
    df = df[df["DealMonth"].notna()].copy()
    if "Deal - Owner" not in df.columns:
        df["Deal - Owner"] = "Unknown"
    out = _build_summary_metrics(df, ["DealMonth", "Deal - Owner"], include_connected_pct=True)
    total_deal_value = _owner_total_deal_value(df)
    out = out.merge(total_deal_value, on=["DealMonth", "Deal - Owner"], how="left")
    out["Total Deal Value"] = pd.to_numeric(out["Total Deal Value"], errors="coerce").fillna(0.0)
    out = out.rename(columns={"Deal - Owner": "Deal Owner"})
    desired_cols = [
        "DealMonth", "Deal Owner", "Accounts", "Connected", "Connected %",
        "Total Deal Value", "Churn Eligible Accounts", "Churn", "Upsell (Net)",
        "Upsell (Positive Only)", "Churn %"
    ]
    existing_cols = [c for c in desired_cols if c in out.columns]
    return out[existing_cols].sort_values(["DealMonth", "Deal Owner"], kind="mergesort")


def summarize_tier_owner(deals_enriched):
    df = deals_enriched.copy()
    df = df[df["DealMonth"].notna()].copy()
    if "Tier" not in df.columns:
        df["Tier"] = pd.NA
    if "Deal - Owner" not in df.columns:
        df["Deal - Owner"] = "Unknown"
    out = _build_summary_metrics(df, ["DealMonth", "Tier", "Deal - Owner"], include_connected_pct=True)
    return out.rename(columns={"Deal - Owner": "Deal Owner"}).sort_values(["DealMonth", "Tier", "Deal Owner"], kind="mergesort")


def summarize_tier_owner_connected(deals_enriched):
    df = deals_enriched.copy()
    df = df[(df["DealMonth"].notna()) & (df["Connected"] == True)].copy()
    if "Tier" not in df.columns:
        df["Tier"] = pd.NA
    if "Deal - Owner" not in df.columns:
        df["Deal - Owner"] = "Unknown"
    out = _build_summary_metrics(df, ["DealMonth", "Tier", "Deal - Owner"], include_connected_pct=False)
    return out.rename(columns={"Deal - Owner": "Deal Owner"}).sort_values(["DealMonth", "Tier", "Deal Owner"], kind="mergesort")


def _owner_total_deal_value(df):
    if df.empty:
        return pd.DataFrame(columns=["DealMonth", "Deal - Owner", "Total Deal Value"])

    value_df = df.copy()
    if "Previous Month Renew Amount" not in value_df.columns:
        value_df["Previous Month Renew Amount"] = 0.0

    value_df["Previous Month Renew Amount"] = pd.to_numeric(
        value_df["Previous Month Renew Amount"], errors="coerce"
    ).fillna(0.0)

    if "Latest Annual PayDT (AsOf MonthEnd)" in value_df.columns:
        annual_mask = value_df["Latest Annual PayDT (AsOf MonthEnd)"].notna()
        value_df = value_df.loc[~annual_mask].copy()

    if value_df.empty:
        return pd.DataFrame(columns=["DealMonth", "Deal - Owner", "Total Deal Value"])

    return (
        value_df.groupby(["DealMonth", "Deal - Owner"], dropna=False, observed=False)["Previous Month Renew Amount"]
        .sum()
        .reset_index(name="Total Deal Value")
    )


def summarize_owner_connected(deals_enriched):
    df = deals_enriched.copy()
    df = df[(df["DealMonth"].notna()) & (df["Connected"] == True)].copy()
    if "Deal - Owner" not in df.columns:
        df["Deal - Owner"] = "Unknown"
    out = _build_summary_metrics(df, ["DealMonth", "Deal - Owner"], include_connected_pct=False)
    total_deal_value = _owner_total_deal_value(df)
    out = out.merge(total_deal_value, on=["DealMonth", "Deal - Owner"], how="left")
    out["Total Deal Value"] = pd.to_numeric(out["Total Deal Value"], errors="coerce").fillna(0.0)
    out = out.rename(columns={"Deal - Owner": "Deal Owner"})
    desired_cols = [
        "DealMonth", "Deal Owner", "Accounts", "Connected",
        "Total Deal Value", "Churn Eligible Accounts", "Churn",
        "Upsell (Net)", "Upsell (Positive Only)", "Churn %"
    ]
    existing_cols = [c for c in desired_cols if c in out.columns]
    return out[existing_cols].sort_values(["DealMonth", "Deal Owner"], kind="mergesort")


# -------------------------
# Core enrichment
# -------------------------
def build_enriched_deals(deals_df, npm_df, report_current_month, churn_cutoff_date=None):
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

    report_month_start = _month_start_from_date(report_current_month)
    prev_report_month_start = _prev_month_start(report_month_start)
    churn_prev_month_cutoff_date = _aligned_prev_month_cutoff(report_month_start, churn_cutoff_date)
    report_month_end_date = _month_end_date(report_month_start)

    deals = deals_df.copy()
    deals["_deal_created_dt"] = pd.to_datetime(deals[deal_date_col], errors="coerce", utc=True).dt.tz_convert(None)
    if deals["_deal_created_dt"].isna().all():
        deals["_deal_created_dt"] = pd.to_datetime(deals[deal_date_col], errors="coerce")

    deals["Original Deal Month"] = deals["_deal_created_dt"].dt.to_period("M").dt.to_timestamp()
    deals["DealMonth"] = pd.Timestamp(report_month_start)
    deals["PrevDealMonth"] = pd.Timestamp(prev_report_month_start)
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

    if npm_df is None or npm_df.empty:
        out = deals_dedup.drop(columns=["_dedup_key"], errors="ignore").copy()
        out["Report Current Month"] = pd.Timestamp(report_month_start).date()
        out["Report Previous Month"] = pd.Timestamp(prev_report_month_start).date()
        out["Churn Current Month Cutoff"] = churn_cutoff_date if churn_cutoff_date else pd.NaT
        out["Churn Previous Month Cutoff"] = churn_prev_month_cutoff_date if churn_prev_month_cutoff_date else pd.NaT
        out["Churn Eligible"] = True
        out["Churned (Reporting)"] = out.get("Churned (AsOf MonthEnd)", False)
        summary_overall = summarize_overall(out)
        summary_tier = summarize_tier(out)
        summary_owner = summarize_owner(out)
        summary_tier_owner = summarize_tier_owner(out)
        summary_tier_owner_connected = summarize_tier_owner_connected(out)
        return out, summary_overall, summary_tier, summary_owner, summary_tier_owner, summary_tier_owner_connected

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

    breakdown_source = (
        npm_valid["Amount Breakdown"].fillna("").astype(str)
        + " "
        + npm_valid["Amount Breakdown by Unit"].fillna("").astype(str)
    )
    breakdown_clean, breakdown_mask = _breakdown_mask_from_series(breakdown_source)

    desc_has_comma = desc.str.contains(",", na=False)
    desc_no_comma = ~desc_has_comma
    contains_email = _series_contains_email(npm_valid["EmailKey"], desc_lower)

    cond_agent_added_any = desc_lower.str.contains("agent added", na=False)
    cond_number_purchased_any = desc_lower.str.contains("number purchased", na=False)
    cond_number_renew_any = desc_lower.str.contains("number renew", na=False)
    cond_annual_subscription_text_any = _annual_desc_series_mask(npm_valid["Amount Description"])

    annual_start = (npm_valid["AmountNum"].fillna(0) > ANNUAL_AMOUNT_THRESHOLD) & (
        cond_annual_subscription_text_any | (desc_lower == ANNUAL_UPGRADE_DESC_EXACT)
    )

    annual_renew_original = (npm_valid["AmountNum"].fillna(0) > ANNUAL_AMOUNT_THRESHOLD) & (
        (contains_email & desc_no_comma) | breakdown_mask
    )

    annual_renew_action_with_breakdown = (npm_valid["AmountNum"].fillna(0) > ANNUAL_AMOUNT_THRESHOLD) & (
        breakdown_mask & (cond_agent_added_any | cond_number_purchased_any)
    )

    annual_renew = (npm_valid["AmountNum"].fillna(0) > ANNUAL_AMOUNT_THRESHOLD) & (
        (contains_email & desc_no_comma)
        | (
            breakdown_mask
            & (
                contains_email
                | cond_agent_added_any
                | cond_number_purchased_any
                | cond_number_renew_any
                | cond_annual_subscription_text_any
            )
        )
    )

    annual_candidates = npm_valid[annual_start | annual_renew].copy()
    annual_candidates["Annual Payment Type"] = np.where(
        annual_start.loc[annual_candidates.index], "Subscription", "Renew"
    )

    excluded_credit_mask = _excluded_credit_desc_series_mask(npm_valid["Amount Description"])
    annual_user_upsell_txns = npm_valid[(~annual_start) & (~annual_renew) & (~excluded_credit_mask)].copy()
    annual_user_upsell_by_month = (
        annual_user_upsell_txns
        .dropna(subset=["EmailKey", "PayMonth"])
        .groupby(["EmailKey", "PayMonth"], dropna=False, observed=False)["AmountNum"]
        .sum()
        .rename("Annual User Current Month Upsell Amount")
        .reset_index()
    )
    annual_user_upsell_map = annual_user_upsell_by_month.set_index(["EmailKey", "PayMonth"])[
        "Annual User Current Month Upsell Amount"
    ]

    cond_email_in_desc = contains_email
    cond_number_purchased = cond_number_purchased_any
    cond_agent_added = desc_lower.str.contains("agent added", na=False) & desc_has_comma
    cond_workspace_sub = cond_annual_subscription_text_any
    cond_number_renew = cond_number_renew_any
    annual_action_with_breakdown = breakdown_mask & (cond_agent_added_any | cond_number_purchased_any)

    renewal_mask = (
        cond_email_in_desc
        | cond_number_purchased
        | cond_agent_added
        | cond_workspace_sub
        | cond_number_renew
        | annual_action_with_breakdown
    )

    renewals = npm_valid[renewal_mask].copy()
    renewals_sorted = renewals.sort_values(["EmailKey", "PayMonth", "PayDT"], kind="mergesort")

    txn_count = (
        renewals_sorted.groupby(["EmailKey", "PayMonth"], dropna=False, observed=False)
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

    deals_dedup["Annual User Current Month Upsell Amount"] = [
        float(annual_user_upsell_map.get((e, m), 0.0))
        for e, m in zip(deals_dedup["EmailKey"], deals_dedup["DealMonth"])
    ]

    deal_month_index = (
        deals_dedup[["EmailKey", "DealMonth"]]
        .dropna(subset=["EmailKey", "DealMonth"])
        .drop_duplicates()
        .sort_values(["EmailKey", "DealMonth"], kind="mergesort")
    )

    monthly_latest_in_month = latest_rows[["EmailKey", "PayDT"]].copy()
    monthly_asof = _merge_latest_asof(
        deal_month_index,
        monthly_latest_in_month,
        payment_dt_col="PayDT",
        extra_cols=None,
        prefix="Latest Monthly "
    )
    monthly_asof = monthly_asof.rename(columns={"Latest Monthly PayDT": "Latest Monthly PayDT (AsOf MonthEnd)"})

    annual_for_asof = annual_candidates[["EmailKey", "PayDT", "Annual Payment Type"]].copy()
    annual_asof = _merge_latest_asof(
        deal_month_index,
        annual_for_asof,
        payment_dt_col="PayDT",
        extra_cols=["Annual Payment Type"],
        prefix="Latest Annual "
    )
    annual_asof = annual_asof.rename(
        columns={
            "Latest Annual PayDT": "Latest Annual PayDT (AsOf MonthEnd)",
            "Latest Annual Annual Payment Type": "Annual Payment Type (AsOf MonthEnd)",
        }
    )

    deals_dedup = deals_dedup.merge(
        monthly_asof,
        on=["EmailKey", "DealMonth"],
        how="left",
    )
    deals_dedup = deals_dedup.merge(
        annual_asof,
        on=["EmailKey", "DealMonth"],
        how="left",
    )

    deals_dedup["Current Month Renew DT Fallback"] = pd.to_datetime(
        deals_dedup["Current Month Renew Date"], errors="coerce"
    )

    same_month_fallback_mask = (
        deals_dedup["Current Month Renew DT Fallback"].notna()
        & (deals_dedup["Current Month Renew DT Fallback"].dt.to_period("M").dt.to_timestamp() == deals_dedup["DealMonth"])
    )

    deals_dedup.loc[same_month_fallback_mask, "Latest Monthly PayDT (AsOf MonthEnd)"] = (
        deals_dedup.loc[same_month_fallback_mask, "Latest Monthly PayDT (AsOf MonthEnd)"]
        .combine_first(deals_dedup.loc[same_month_fallback_mask, "Current Month Renew DT Fallback"])
    )

    deals_dedup["NextMonthStart"] = _next_month_start(deals_dedup["DealMonth"])

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

    deals_dedup["Report Current Month"] = pd.Timestamp(report_month_start).date()
    deals_dedup["Report Previous Month"] = pd.Timestamp(prev_report_month_start).date()
    deals_dedup["Churn Current Month Cutoff"] = churn_cutoff_date if churn_cutoff_date else pd.NaT
    deals_dedup["Churn Previous Month Cutoff"] = churn_prev_month_cutoff_date if churn_prev_month_cutoff_date else pd.NaT

    if churn_prev_month_cutoff_date is None:
        deals_dedup["Monthly Churn Eligible"] = True
        deals_dedup["Annual Churn Eligible"] = True
        deals_dedup["Churn Eligible"] = True
    else:
        prev_renew_date_series = pd.to_datetime(deals_dedup["Previous Month Renew Date"], errors="coerce").dt.date
        deals_dedup["Monthly Churn Eligible"] = prev_renew_date_series.notna() & (prev_renew_date_series <= churn_prev_month_cutoff_date)

        annual_due_date_series = pd.to_datetime(deals_dedup["Annual Valid Till (AsOf MonthEnd)"], errors="coerce").dt.date
        annual_user_series = deals_dedup["Latest Annual PayDT (AsOf MonthEnd)"].notna()
        annual_due_after_cutoff_in_report_month = (
            annual_user_series
            & pd.Series(annual_due_date_series, index=deals_dedup.index).notna()
            & (pd.Series(annual_due_date_series, index=deals_dedup.index) > churn_cutoff_date)
            & (pd.Series(annual_due_date_series, index=deals_dedup.index) <= report_month_end_date)
        )
        deals_dedup["Annual Churn Eligible"] = annual_user_series & (~annual_due_after_cutoff_in_report_month)

        deals_dedup["Churn Eligible"] = np.where(
            annual_user_series,
            deals_dedup["Annual Churn Eligible"],
            deals_dedup["Monthly Churn Eligible"],
        )

    deals_dedup["Churned (Reporting)"] = deals_dedup["Churned (AsOf MonthEnd)"] & deals_dedup["Churn Eligible"].fillna(False)

    prev_amt = deals_dedup["Previous Month Renew Amount"].fillna(0.0)
    curr_amt = deals_dedup["Current Month Renew Amount"].fillna(0.0)

    is_annual_user_asof = deals_dedup["Latest Annual PayDT (AsOf MonthEnd)"].notna()
    eligible = (prev_amt > 0) & (~deals_dedup["Churned (AsOf MonthEnd)"]) & (~is_annual_user_asof)

    monthly_upsell_net = np.where(eligible, (curr_amt - prev_amt), 0.0)
    annual_user_upsell_amount = pd.to_numeric(
        deals_dedup["Annual User Current Month Upsell Amount"], errors="coerce"
    ).fillna(0.0)

    deals_dedup["Upsell Net Change"] = np.where(
        is_annual_user_asof,
        annual_user_upsell_amount,
        monthly_upsell_net,
    )
    deals_dedup["Upsell Positive Only"] = np.where(
        deals_dedup["Upsell Net Change"] > 0,
        deals_dedup["Upsell Net Change"],
        0.0,
    )
    out = deals_dedup.drop(columns=["_dedup_key", "Current Month Renew DT Fallback"], errors="ignore").copy()
    summary_overall = summarize_overall(out)
    summary_tier = summarize_tier(out)
    summary_owner = summarize_owner(out)
    summary_tier_owner = summarize_tier_owner(out)
    summary_tier_owner_connected = summarize_tier_owner_connected(out)
    return out, summary_overall, summary_tier, summary_owner, summary_tier_owner, summary_tier_owner_connected


def _autosize_worksheet(ws, min_width=10, max_width=50, padding=2):
    for column_cells in ws.columns:
        lengths = []
        for cell in column_cells:
            value = cell.value
            if value is None:
                continue
            lengths.extend(len(line) for line in str(value).splitlines())
        if not lengths:
            continue
        adjusted_width = max(min_width, min(max(lengths) + padding, max_width))
        ws.column_dimensions[column_cells[0].column_letter].width = adjusted_width


@st.cache_data(show_spinner=False)
def make_excel(
    deals_raw,
    deals_enriched,
    summary_overall,
    summary_tier,
    summary_owner,
    summary_tier_owner,
    summary_tier_owner_connected,
):
    summary_owner_connected = summarize_owner_connected(deals_enriched)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        deals_enriched.to_excel(writer, sheet_name="Deals_enriched", index=False)
        summary_overall.to_excel(writer, sheet_name="Summary_overall", index=False)
        summary_tier.to_excel(writer, sheet_name="Summary_tier", index=False)
        summary_owner.to_excel(writer, sheet_name="Summary_owner", index=False)
        summary_tier_owner.to_excel(writer, sheet_name="Summary_tier_owner", index=False)
        summary_tier_owner_connected.to_excel(writer, sheet_name="Summary_tier_owner_conn", index=False)

        connected_sheet = writer.book["Summary_tier_owner_conn"]
        start_row = len(summary_tier_owner_connected.index) + 4
        connected_sheet.cell(row=start_row, column=1, value="Owner wise summary for connected users")
        summary_owner_connected.to_excel(writer, sheet_name="Summary_tier_owner_conn", index=False, startrow=start_row)

        deals_raw.to_excel(writer, sheet_name="Deals_raw", index=False)

        for ws in writer.book.worksheets:
            _autosize_worksheet(ws)
    return output.getvalue()


def kpi_row(summary_df):
    if summary_df is None or summary_df.empty:
        st.info("No summary available.")
        return

    latest = summary_df.sort_values("DealMonth").iloc[-1]
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Accounts", int(latest["Accounts"]))
    c2.metric("Connected", int(latest["Connected"]))
    c3.metric("Churn", int(latest["Churn"]))
    churn_pct = float(latest["Churn %"]) if pd.notna(latest["Churn %"]) else np.nan
    c4.metric("Churn %", f"{churn_pct:.2f}%" if pd.notna(churn_pct) else "NA")
    c5.metric("Upsell Net", f'{float(latest["Upsell (Net)"]):,.2f}')
    c6.metric("Upsell Positive Only", f'{float(latest["Upsell (Positive Only)"]):,.2f}')


def main():
    st.set_page_config(page_title="KrispCall | Account Management", page_icon="📞", layout="wide")
    init_state()
    _require_secrets()
    inject_brand_css()
    render_header()

    with st.sidebar:
        try:
            st.image("assets/KrispCallLogo.png", use_container_width=True)
        except Exception:
            pass
        st.markdown("### Controls")

    if not login_gate():
        return

    report_current_month_input = st.sidebar.date_input(
        "Current month",
        value=date.today().replace(day=1),
        help="This report month is applied to all uploaded users. The previous month is derived automatically.",
    )
    use_mid_month_churn = st.sidebar.toggle("Use mid-month churn cutoff", value=False)
    churn_cutoff_date = None
    if use_mid_month_churn:
        default_cutoff = _month_end_date(_month_start_from_date(report_current_month_input)) or date.today()
        churn_cutoff_date = st.sidebar.date_input(
            "Current month churn cutoff date",
            value=default_cutoff,
            help="Monthly churn will only consider users whose previous-month renewal date is on or before the aligned prior-month cutoff. Annual users whose annual due date falls after this cutoff within the current month are excluded from churn.",
        )
    end_date = st.sidebar.date_input("Payments to date", value=date.today())
    st.sidebar.caption("Mixpanel Export API. A minimum annual-history lookback is applied automatically.")
    deals_file = st.sidebar.file_uploader("Upload Deals CSV", type=["csv"])
    fetch_btn = st.sidebar.button("Fetch payments", type="primary")

    if deals_file is None:
        st.info("Upload your Deals CSV to begin.")
        return

    deals_raw = pd.read_csv(deals_file)

    needs_fetch = (
        st.session_state.get("npm_cached") is None
        or st.session_state.get("last_fetch_to_date") != end_date
    )

    if fetch_btn or needs_fetch:
        with st.spinner("Fetching Mixpanel payments..."):
            npm_df, stats = fetch_mixpanel_npm(end_date)
            st.session_state["npm_cached"] = npm_df
            st.session_state["npm_stats"] = stats
            st.session_state["last_fetch_to_date"] = end_date
            if fetch_btn:
                st.success("Payments fetched.")

    npm_df = st.session_state.get("npm_cached")
    fetch_stats = st.session_state.get("npm_stats")

    if npm_df is None:
        st.warning("Click Fetch payments to load Mixpanel events.")
        return

    if fetch_stats:
        st.sidebar.markdown("### Mixpanel fetch stats")
        st.sidebar.write(fetch_stats)

    with st.spinner("Building enriched dataset..."):
        deals_enriched, summary_overall, summary_tier, summary_owner, summary_tier_owner, summary_tier_owner_connected = build_enriched_deals(
            deals_raw,
            npm_df,
            report_current_month=report_current_month_input,
            churn_cutoff_date=churn_cutoff_date,
        )

    kpi_row(summary_overall)

    prev_month_display = _prev_month_start(_month_start_from_date(report_current_month_input))
    st.caption(
        f"Reporting current month: {pd.Timestamp(_month_start_from_date(report_current_month_input)).date()} | "
        f"Previous month: {pd.Timestamp(prev_month_display).date()} | "
        f"Churn cutoff: {churn_cutoff_date if churn_cutoff_date else 'Month end'}"
    )

    tab1, tab2, tab3, tab4 = st.tabs(["Summary", "Visuals", "Deals enriched", "Payments preview"])

    with tab1:
        st.subheader("Overall summary")
        st.dataframe(summary_overall, use_container_width=True)

        st.subheader("Tier wise summary")
        st.dataframe(summary_tier, use_container_width=True)

        st.subheader("Owner wise summary")
        st.dataframe(summary_owner, use_container_width=True)

        st.subheader("Tier and Owner wise summary")
        st.dataframe(summary_tier_owner, use_container_width=True)

        st.subheader("Tier and Owner wise summary. Connected only")
        st.dataframe(summary_tier_owner_connected, use_container_width=True)

    with tab2:
        overall = summary_overall.copy()
        if overall.empty:
            st.info("No data to chart.")
        else:
            overall["DealMonth"] = pd.to_datetime(overall["DealMonth"], errors="coerce")
            overall = overall.sort_values("DealMonth")
            st.caption("Churn percentage over time")
            st.line_chart(overall.set_index("DealMonth")[["Churn %"]])
            st.caption("Churn count over time")
            st.line_chart(overall.set_index("DealMonth")[["Churn"]])
            st.caption("Upsell net over time")
            st.line_chart(overall.set_index("DealMonth")[["Upsell (Net)"]])
            st.caption("Upsell positive only over time")
            st.line_chart(overall.set_index("DealMonth")[["Upsell (Positive Only)"]])

    with tab3:
        st.dataframe(deals_enriched, use_container_width=True)

    with tab4:
        st.caption("Reduced payments dataset after filter and dedupe. First 200 rows shown.")
        st.dataframe(npm_df.head(200), use_container_width=True)

    st.divider()
    st.subheader("Export")
    st.caption("The workbook below is generated from the current dataset and summaries.")

    excel_bytes = make_excel(
        deals_raw=deals_raw,
        deals_enriched=deals_enriched,
        summary_overall=summary_overall,
        summary_tier=summary_tier,
        summary_owner=summary_owner,
        summary_tier_owner=summary_tier_owner,
        summary_tier_owner_connected=summary_tier_owner_connected,
    )

    report_month_for_name = pd.Timestamp(_month_start_from_date(report_current_month_input)).strftime("%b_%Y")
    generated_date_for_name = date.today().isoformat()
    download_filename = f"AMKPI_{report_month_for_name}_{generated_date_for_name}.xlsx"

    st.download_button(
        "Download Excel workbook",
        data=excel_bytes,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_excel_workbook",
    )


if __name__ == "__main__":
    main()
