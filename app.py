import io
import hmac
import html
import datetime
import time
import os
import pandas as pd
import streamlit as st
import folium
from folium.plugins import Fullscreen
from branca.element import MacroElement
from jinja2 import Template
import json
import math
import re
import requests
from streamlit_folium import st_folium

st.set_page_config(page_title="IRUP Project Map", page_icon="🗺️", layout="wide")

# ---------------- Login (single shared account) ----------------
def require_login():
    """Blocks the whole app until the correct username/password (from st.secrets) is entered."""
    try:
        want_user = str(st.secrets["auth"]["username"])
        want_pass = str(st.secrets["auth"]["password"])
    except Exception:
        st.error("Login is not configured. Add an [auth] section with username and password to the app's "
                 "secrets (see .streamlit/secrets.toml.example). The app stays locked until this is set.")
        st.stop()

    if st.session_state.get("authed"):
        if st.sidebar.button("Log out"):
            st.session_state["authed"] = False
            st.rerun()
        return

    st.title("🔒 IRUP Project Map")
    if st.session_state.get("fails", 0) >= 5:
        st.error("Too many wrong attempts. Refresh the page to try again.")
        st.stop()
    with st.form("login"):
        user = st.text_input("Username")
        pw = st.text_input("Password", type="password")
        go = st.form_submit_button("Log in")
    if go:
        ok_user = hmac.compare_digest(user.encode(), want_user.encode())
        ok_pass = hmac.compare_digest(pw.encode(), want_pass.encode())
        if ok_user and ok_pass:
            st.session_state["authed"] = True
            st.session_state["fails"] = 0
            st.rerun()
        else:
            st.session_state["fails"] = st.session_state.get("fails", 0) + 1
            time.sleep(1)
            st.error("Wrong username or password.")
    st.stop()

# The login turns ON by itself as soon as an [auth] section (username + password) exists in the app's Secrets.
def login_configured():
    try:
        return "auth" in st.secrets
    except Exception:          # no secrets file / no secrets set
        return False

if login_configured():
    require_login()
else:
    st.sidebar.warning("No login is set: anyone with the link can open this app. "
                       "Add an [auth] section to the app's Secrets to turn on the password.")

st.title("🗺️ IRUP Project Mapping & Status Dashboard")
st.caption("Interactive map and filters powered by your IRUP SITELIST Excel sheet.")

DEFAULT_FILE = "IRUP TRACKER - 2026 OCTOBER 29.xlsx"

def load_tracker(uploaded_file):
    # The workbook has title/metadata rows above the actual headers.
    source = uploaded_file if uploaded_file is not None else DEFAULT_FILE
    if uploaded_file is None and not os.path.exists(DEFAULT_FILE):
        return None, "Upload your IRUP Tracker Excel file to get started."
    try:
        df = pd.read_excel(source, sheet_name="IRUP SITELIST", header=1)
    except Exception as exc:
        return None, f"Could not read the 'IRUP SITELIST' sheet: {exc}"

    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()].copy()

    # Remove repeated header rows and completely empty rows.
    # A row counts as a site if it has a SITE NO *or* a SITE NAME (some sites have no SITE NO yet).
    def _filled(col):
        if col not in df.columns:
            return pd.Series(False, index=df.index)
        return df[col].notna() & (df[col].astype(str).str.strip() != "")
    if "SITE NO" in df.columns:
        df = df[df["SITE NO"].astype(str).str.strip().str.upper() != "SITE NO"]
    df = df[_filled("SITE NO") | _filled("SITE NAME")]

    for col in ["LATITUDE", "LONGITUDE", "OVERALL PROGRESS", "DESIGN PROGRESS",
                "PERMITTING PROGESS", "PERMITTING PROGRESS", "EXTRACTION PROGRESS",
                "APPROX TONS", "DESIGN TONS", "TOTAL TONS EXTRACTED"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df, None

def clean_text(series):
    return series.fillna("(Blank)").astype(str).str.strip().replace("", "(Blank)")

uploaded = st.sidebar.file_uploader("Upload latest IRUP Tracker (.xlsx)", type=["xlsx"])
df, error = load_tracker(uploaded)
if error:
    st.info(error)
    st.stop()

st.sidebar.header("Search & Filters")
filter_columns = [
    "AREA", "PROVINCE", "CITY", "QTR", "APM ASSIGNED",
    "OVERALL MILESTONE", "DESIGN STATUS", "PERMITTING GENERAL STAUS",
    "EXTRACTION STATUS", "STATUS", "SUBCON"
]

# Quick stage filters.
#  - "cols": the stage counts as reached when any of those Excel cells is filled in (a date).
#  - "col"/"values": the stage matches when that column currently equals one of the values.
STAGE_RULES = {
    "Survey start":       {"cols": ["SURVEY START"]},
    "Survey done":        {"cols": ["SURVEY DONE DATE"]},
    "Design completed":   {"cols": ["DESIGN COMPLETION"]},
    "Design submitted":   {"cols": ["DESIGN SUBMISSION"]},
    "Design approved":    {"cols": ["DESIGN APPROVED"]},
    "On going permitting": {"col": "PERMITTING GENERAL STAUS", "values": ["ON GOING PERMITTING"]},
    "RFE (Ready for extraction)": {"col": "PERMITTING GENERAL STAUS", "values": ["RFE"]},
    "Extraction started": {"cols": ["EXTRACTION START"]},
    "Extraction done":    {"cols": ["EXTRACTION DONE"]},
}
SEARCH_COLUMNS = ["SITE NO", "SITE NAME", "CITY", "PROVINCE", "AREA", "APM ASSIGNED", "WBS CODE",
                  "PS NAME", "SUBCON", "DESIGN REMARKS/ ISSUE", "PERMITTING ISSUE / REMARKS", "EXTRACTION REMARKS"]

def stage_mask(frame, rule):
    """Rows that match one stage rule (see STAGE_RULES)."""
    if "cols" in rule:
        return has_value(frame, rule["cols"])
    col = rule["col"]
    if col not in frame.columns:
        return pd.Series(False, index=frame.index)
    norm = frame[col].fillna("").astype(str).str.upper().str.split().str.join(" ")
    return norm.isin(rule["values"])

def has_value(frame, columns):
    """True for rows where ANY of the given columns is filled in (blank, 0 and '-' count as empty)."""
    result = pd.Series(False, index=frame.index)
    for col in columns:
        if col in frame.columns:
            txt = frame[col].astype(str).str.strip().str.upper()
            result |= frame[col].notna() & ~txt.isin(["", "NAN", "NONE", "NAT", "0", "0.0", "-"])
    return result

def reset_filters():
    st.session_state["q_search"] = ""
    st.session_state["stage_sel"] = []
    st.session_state["stage_have"] = "Have reached"
    st.session_state["stage_match"] = "All selected"
    for col in filter_columns:
        st.session_state[f"flt_{col}"] = []
    st.session_state.pop("tons_rng", None)

st.sidebar.button("Reset all filters", on_click=reset_filters, use_container_width=True)

filtered = df.copy()
active_filters = {}

# 1) Free-text search (every word must appear somewhere in the site's details)
query = st.sidebar.text_input("Search", key="q_search",
                              placeholder="site no, name, city, APM, subcon, remarks...")
if query.strip():
    present = [c for c in SEARCH_COLUMNS if c in filtered.columns]
    haystack = filtered[present].fillna("").astype(str).agg(" ".join, axis=1).str.lower()
    for word in query.lower().split():
        filtered = filtered[haystack.loc[filtered.index].str.contains(word, regex=False)]

# 2) Quick stage filter (survey done, design submitted, permit secured, ...)
available_stages = [name for name, rule in STAGE_RULES.items()
                    if any(c in filtered.columns for c in (rule.get("cols") or [rule.get("col")]))]
stage_sel = st.sidebar.multiselect("Stage", available_stages, key="stage_sel",
                                   placeholder="e.g. Survey done, Design submitted",
                                   help="Date stages (survey, design, extraction) count as reached when the date is filled in the Excel. "
                                        "'On going permitting' and 'RFE' match the Permitting General Status column.")
selected_stage_cols = []
if stage_sel:
    have = st.sidebar.radio("Show sites that", ["Have reached", "Have NOT reached"], key="stage_have", horizontal=True)
    match = "All selected"
    if len(stage_sel) > 1:
        match = st.sidebar.radio("Match", ["All selected", "Any selected"], key="stage_match", horizontal=True)
    masks = [stage_mask(filtered, STAGE_RULES[name]) for name in stage_sel]
    if have == "Have NOT reached":
        masks = [~m for m in masks]
    keep = masks[0]
    for m in masks[1:]:
        keep = (keep & m) if match == "All selected" else (keep | m)
    filtered = filtered[keep]
    for name in stage_sel:
        selected_stage_cols += [c for c in (STAGE_RULES[name].get("cols") or [STAGE_RULES[name]["col"]]) if c in filtered.columns]

# 3) Column filters (dropdown lists)
st.sidebar.divider()
for col in filter_columns:
    if col in filtered.columns:
        options = sorted(clean_text(filtered[col]).unique().tolist(), key=str.casefold)
        selected = st.sidebar.multiselect(col.replace("_", " ").title(), options, key=f"flt_{col}")
        if selected:
            active_filters[col] = selected
            filtered = filtered[clean_text(filtered[col]).isin(selected)]

if "APPROX TONS" in filtered.columns:
    st.sidebar.divider()
    st.sidebar.caption("Optional approximate tons range")
    vals = filtered["APPROX TONS"].dropna()
    if not vals.empty and float(vals.min()) < float(vals.max()):
        low, high = st.sidebar.slider(
            "Approx. Tons", min_value=float(vals.min()), max_value=float(vals.max()),
            value=(float(vals.min()), float(vals.max())), key="tons_rng"
        )
        # Only filter when the slider was actually narrowed; at full range, sites with blank tons stay in.
        if (low, high) != (float(vals.min()), float(vals.max())):
            filtered = filtered[filtered["APPROX TONS"].between(low, high)]

# ---------------- Dashboard look & helpers ----------------
REGION_COLORS = {
    "GMM": "#548235", "NLZ": "#2F5597", "SLZ": "#BF8F00",
    "VIS": "#C55A11", "N. MIN": "#203864", "S. MIN": "#7030A0",
}
# Pin / status colour per OVERALL MILESTONE (exact values from the Excel column).
MILESTONE_COLORS = {
    "NOT STARTED":          "#9E9E9E",
    "ON GOING SURVEY":      "#4FC3F7",
    "ON GOING DESIGN":      "#1E88E5",
    "DESIGN COMPLETED":     "#3949AB",
    "DESIGN SUBMITTED":     "#8E24AA",
    "DESIGN APPROVED":      "#00897B",
    "ON GOING PERMITTING":  "#FFB300",
    "READY FOR EXTRACTION": "#7CB342",
    "ON GOING EXTRACTION":  "#2E7D32",
    "CANCELLED":            "#E53935",
}
UNKNOWN_COLOR = "#212121"

DASH_CSS = """<style>
.irup-h{font:800 21px Georgia,'Times New Roman',serif;letter-spacing:.6px;text-transform:uppercase;margin:8px 0 6px;color:inherit}
.irup-sub{font:600 11px Arial,sans-serif;letter-spacing:.6px;text-transform:uppercase;opacity:.6;margin:-2px 0 8px}
.irup-row{display:flex;flex-wrap:wrap;gap:10px;margin:6px 0 18px}
.irup-hov{position:relative}
.irup-card,.irup-tile,.irup-region{flex:1 1 118px;min-width:100px;text-align:center;padding:8px 8px 9px;box-sizing:border-box}
.irup-card{background:#fff;border:2px solid #1F3864;border-radius:14px;box-shadow:0 1px 3px rgba(0,0,0,.2)}
.irup-card .irup-lab,.irup-card .irup-val,.irup-tile .irup-lab,.irup-tile .irup-val{color:#1F3864}
.irup-tile{background:#DDEBF7;border-radius:6px}
.irup-region{color:#fff;border-radius:6px}
.irup-lab{font:700 11px Georgia,serif;letter-spacing:.4px;text-transform:uppercase;line-height:1.25;min-height:14px}
.irup-val{font:700 24px Georgia,serif;line-height:1.3}
.irup-hov:hover{filter:brightness(1.06);z-index:100000}
.irup-tip{display:none;position:absolute;top:100%;left:50%;transform:translateX(-50%);z-index:99999;margin-top:6px;min-width:230px;max-width:340px;max-height:270px;overflow-y:auto;background:#1F2937;color:#fff;border-radius:8px;padding:8px 10px;font:12px/1.45 Arial,sans-serif;text-align:left;text-transform:none;letter-spacing:0;font-weight:400;box-shadow:0 6px 18px rgba(0,0,0,.45);white-space:normal}
.irup-tip.tip-l{left:0;transform:none}
.irup-tip.tip-r{left:auto;right:0;transform:none}
.irup-tip-h{font-weight:700;border-bottom:1px solid rgba(255,255,255,.25);padding-bottom:3px;margin-bottom:3px}
.irup-hov:hover>.irup-tip{display:block}
.irup-tbl{border-collapse:separate;border-spacing:0;width:100%;font:13px Georgia,serif;text-align:center;margin-bottom:18px}
.irup-tbl th{background:#002060;color:#fff;padding:8px 6px;font-size:12px;letter-spacing:.4px;text-transform:uppercase}
.irup-tbl td{background:#DDEBF7;color:#1F3864;padding:8px 6px;border-bottom:3px solid #fff;position:relative}
.irup-tbl tr.irup-hl td{background:#C5E0B4}
.irup-tbl td.irup-first{font-weight:700;color:#fff}
.irup-sb-row{display:flex;align-items:center;gap:10px;margin:7px 0;font:13px Arial,sans-serif;position:relative}
.irup-dot{flex:0 0 12px;height:12px;border-radius:50%}
.irup-sb-name{flex:0 0 150px}
.irup-bar{flex:1;height:14px;background:rgba(128,128,128,.22);border-radius:7px;overflow:hidden}
.irup-bar>span{display:block;height:100%;border-radius:7px}
.irup-sb-n{flex:0 0 46px;text-align:right;font-weight:700}
</style>"""

def cell_text(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else f"{v:g}"
    text = str(v).replace("\xa0", " ").strip()
    if text[:10].count("-") == 2 and text[10:].strip() in ("", "00:00:00"):
        text = text[:10]
    return text

def is_progress_col(name):
    return "PROGRESS" in name.upper() or "PROGESS" in name.upper()

def site_labels(frame):
    """Escaped 'SITE NO — SITE NAME' text for every row, keyed by index."""
    labels = {}
    for idx, no, name in zip(frame.index,
                             frame["SITE NO"] if "SITE NO" in frame.columns else [None] * len(frame),
                             frame["SITE NAME"] if "SITE NAME" in frame.columns else [None] * len(frame)):
        no, name = cell_text(no), cell_text(name) or "Unnamed site"
        labels[idx] = html.escape(f"{no} — {name}" if no else name)
    return labels

def tip_html(subset, labels, align=""):
    if subset is None or len(subset) == 0:
        return ""
    shown = [labels.get(i, "") for i in list(subset.index)[:40]]
    body = "".join(f"<div>• {t}</div>" for t in shown)
    if len(subset) > 40:
        body += f'<div style="opacity:.7">+ {len(subset) - 40} more…</div>'
    return f'<div class="irup-tip {align}"><div class="irup-tip-h">{len(subset)} site(s)</div>{body}</div>'

def _align(i, n):
    return "tip-l" if i < n / 3 else ("tip-r" if i >= 2 * n / 3 and n > 2 else "")

def cards_row(specs, labels):
    """specs: dicts with label, value, optional subset / cls / style."""
    parts = []
    for i, sp in enumerate(specs):
        tip = tip_html(sp.get("subset"), labels, _align(i, len(specs)))
        hov = " irup-hov" if tip else ""
        parts.append(f'<div class="{sp.get("cls", "irup-tile")}{hov}" style="{sp.get("style", "")}">'
                     f'<div class="irup-lab">{sp["label"]}</div><div class="irup-val" style="{sp.get("val_style", "")}">{sp["value"]}</div>{tip}</div>')
    return '<div class="irup-row">' + "".join(parts) + "</div>"

def html_table(headers, rows, labels, highlight=()):
    head = "".join(f"<th>{h}</th>" for h in headers)
    body = ""
    for ri, r in enumerate(rows):
        tds = ""
        for ci, cell in enumerate(r):
            parts = list(cell) if isinstance(cell, tuple) else [cell]
            text, subset, style, cls = (parts + [None, "", ""][len(parts) - 1:])[:4]
            tip = tip_html(subset, labels, _align(ci, len(r)))
            if ri in highlight and ci > 0:
                style = re.sub(r"background:[^;]+;?", "", style or "")   # keep the green current-week row
            tds += f'<td class="{cls}{" irup-hov" if tip else ""}" style="{style}">{text}{tip}</td>'
        body += f'<tr class="{"irup-hl" if ri in highlight else ""}">{tds}</tr>'
    return f'<table class="irup-tbl"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>'

def num(frame, col):
    if col not in frame.columns:
        return pd.Series(0.0, index=frame.index)
    return pd.to_numeric(frame[col], errors="coerce").fillna(0)

def region_key(frame):
    if "AREA" not in frame.columns:
        return pd.Series("", index=frame.index)
    return frame["AREA"].fillna("").astype(str).str.strip().str.upper()

def region_names(frame):
    present = set(region_key(frame))
    extra = sorted(r for r in present if r and r not in REGION_COLORS)
    return list(REGION_COLORS) + extra

def region_cards(frame, mask, labels, value_fn=None):
    keys = region_key(frame)
    specs = []
    for r in region_names(frame):
        sub = frame[mask & (keys == r)]
        specs.append({"label": r, "value": f"{len(sub):,}", "subset": sub, "cls": "irup-region",
                      "style": f"background:{REGION_COLORS.get(r, '#546E7A')}"})
    return cards_row(specs, labels)

MANILA_TZ = datetime.timezone(datetime.timedelta(hours=8))
WEEK_RE = re.compile(r"^\s*(\d{1,2})\s*-\s*(20\d{2})\s*$")

def parse_week(v):
    m = WEEK_RE.match(str(v)) if v is not None else None
    if m and 1 <= int(m.group(1)) <= 53:
        return (int(m.group(2)), int(m.group(1)))
    return None

def weeknum_tuple(d):
    """(year, week) the way Excel's WEEKNUM numbers it (week 1 contains Jan 1, weeks start on Sunday)."""
    jan1 = datetime.date(d.year, 1, 1)
    offset = (jan1.weekday() + 1) % 7
    return (d.year, ((d - jan1).days + offset) // 7 + 1)

def week_tuples(frame, date_cols=(), week_cols=()):
    """(year, week) per row. Columns are tried in the order given, so list the REVISED column first:
    a site uses its revised date when it has one, otherwise the original target."""
    result = pd.Series([None] * len(frame), index=frame.index, dtype=object)
    for col in date_cols:
        if col in frame.columns:
            try:
                dt = pd.to_datetime(frame[col], errors="coerce", format="mixed")
            except Exception:
                dt = pd.to_datetime(frame[col], errors="coerce")
            for idx in dt.index[dt.notna() & (dt.dt.year >= 2020)]:
                if result[idx] is None:
                    result[idx] = weeknum_tuple(dt[idx].date())
    for col in week_cols:
        if col in frame.columns:
            for idx, v in frame[col].items():
                if result[idx] is None:
                    result[idx] = parse_week(v)
    return result

def week_monday(y, w):
    """Monday of that week (Mon-Sun brackets, like the Excel dashboard)."""
    jan1 = datetime.date(y, 1, 1)
    first_sunday = jan1 - datetime.timedelta(days=(jan1.weekday() + 1) % 7)
    return first_sunday + datetime.timedelta(days=7 * (w - 1) + 1)

def week_bracket(y, w):
    start = week_monday(y, w)
    return f"{start:%b.%d} - {start + datetime.timedelta(days=6):%b.%d}"

def timeline_html(frame, weeks, done_mask, labels, show_all=False):
    """Rows: date bracket | week | target | done | gap. Target = sites planned for that week."""
    today = datetime.datetime.now(MANILA_TZ).date()
    cur_key = weeknum_tuple(today - datetime.timedelta(days=today.weekday()))   # the Mon-Sun bracket that contains today
    present = [w for w in weeks if w is not None]
    if show_all and present:
        first, last = week_monday(*min(present)), week_monday(*max(present))
    else:
        first = week_monday(*cur_key) - datetime.timedelta(weeks=1)
        last = week_monday(*cur_key) + datetime.timedelta(weeks=4)
    rows, highlight = [], []
    d = first
    while d <= last:
        key = weeknum_tuple(d)
        in_wk = weeks.map(lambda x: x == key)
        target = frame[in_wk]
        done = target[done_mask.loc[target.index]]
        gap = target[~done_mask.loc[target.index]]
        if key == cur_key:
            highlight.append(len(rows))
        rows.append([week_bracket(*key), f"{key[1]}-{key[0]}",
                     (len(target), target, "background:#FCE4D6"), (len(done), done), (len(gap), gap)])
        d += datetime.timedelta(weeks=1)
    return html_table(["Date bracket", "Week", "Target", "Done", "Gap"], rows, labels, highlight)

# ---------------- Dashboard data (all sites in the current results, cancelled included) ----------------
def build_masks(f):
    m = {}
    m["survey_start"] = has_value(f, ["SURVEY START"])
    m["survey_done"] = has_value(f, ["SURVEY DONE DATE"])
    m["og_survey"] = m["survey_start"] & ~m["survey_done"]
    m["design_completed"] = has_value(f, ["DESIGN COMPLETION"])
    m["design_submitted"] = has_value(f, ["DESIGN SUBMISSION"])
    m["design_approved"] = has_value(f, ["DESIGN APPROVED"])
    m["og_permit"] = stage_mask(f, {"col": "PERMITTING GENERAL STAUS", "values": ["ON GOING PERMITTING"]})
    m["rfe"] = stage_mask(f, {"col": "PERMITTING GENERAL STAUS", "values": ["RFE"]})
    m["started"] = has_value(f, ["EXTRACTION START"])
    m["done"] = has_value(f, ["EXTRACTION DONE"])
    return m

def status_html(frame, labels):
    if "OVERALL MILESTONE" not in frame.columns or frame.empty:
        return ""
    keys = frame["OVERALL MILESTONE"].fillna("").astype(str).str.upper().str.split().str.join(" ")
    groups = [(name.title(), color, keys == name) for name, color in MILESTONE_COLORS.items()]
    groups.append(("Other / blank", UNKNOWN_COLOR, ~keys.isin(list(MILESTONE_COLORS))))
    counts = [int(mask.sum()) for _, _, mask in groups]
    biggest = max(counts) or 1
    rows = ""
    for (name, color, mask), n in zip(groups, counts):
        if n == 0 and name == "Other / blank":
            continue
        tip = tip_html(frame[mask], labels, "tip-l")
        rows += (f'<div class="irup-sb-row{" irup-hov" if tip else ""}"><span class="irup-dot" style="background:{color}"></span>'
                 f'<span class="irup-sb-name">{name}</span><span class="irup-bar"><span style="width:{n / biggest * 100:.1f}%;background:{color}"></span></span>'
                 f'<span class="irup-sb-n">{n:,}</span>{tip}</div>')
    return rows

scope = filtered
labels = site_labels(filtered)
M = build_masks(scope)
all_true = pd.Series(True, index=scope.index)

# ---------------- Milestone panel ----------------
st.markdown(DASH_CSS, unsafe_allow_html=True)
st.markdown('<div class="irup-h">Milestone</div>', unsafe_allow_html=True)

lat_ok = pd.to_numeric(scope["LATITUDE"], errors="coerce").notna() if "LATITUDE" in scope.columns else pd.Series(False, index=scope.index)
lon_ok = pd.to_numeric(scope["LONGITUDE"], errors="coerce").notna() if "LONGITUDE" in scope.columns else pd.Series(False, index=scope.index)
no_coords = scope[~(lat_ok & lon_ok)]
avg_progress = pd.to_numeric(scope.get("OVERALL PROGRESS", pd.Series(dtype=float)), errors="coerce").mean()

def _tons(col):
    return f"{num(scope, col).sum():,.2f}"

st.markdown(cards_row([
    {"label": "Sites", "value": f"{len(scope):,}", "cls": "irup-card", "style": "border-color:#1F3864"},
    {"label": "Missing coordinates", "value": f"{len(no_coords):,}", "subset": no_coords, "cls": "irup-card", "style": "border-color:#C00000"},
    {"label": "Approx. tons", "value": _tons("APPROX TONS"), "cls": "irup-card", "style": "border-color:#2F5597"},
    {"label": "Design tons", "value": _tons("DESIGN TONS"), "cls": "irup-card", "style": "border-color:#548235"},
    {"label": "Actual tons", "value": _tons("TOTAL TONS EXTRACTED"), "cls": "irup-card", "style": "border-color:#BF8F00"},
    {"label": "Avg. progress", "value": f"{avg_progress:.1%}" if pd.notna(avg_progress) else "N/A", "cls": "irup-card", "style": "border-color:#C55A11"},
], labels), unsafe_allow_html=True)

left, right = st.columns([3, 2], gap="large")
with left:
    st.markdown('<div class="irup-h" style="font-size:16px">Number of sites per region</div>', unsafe_allow_html=True)
    st.markdown(region_cards(scope, all_true, labels), unsafe_allow_html=True)
    st.markdown('<div class="irup-h" style="font-size:16px">Where the sites are now</div>', unsafe_allow_html=True)
    st.markdown(cards_row([
        {"label": "Submitted design", "value": int(M["design_submitted"].sum()), "subset": scope[M["design_submitted"]], "cls": "irup-card"},
        {"label": "Approved design", "value": int(M["design_approved"].sum()), "subset": scope[M["design_approved"]], "cls": "irup-card"},
        {"label": "On going permitting", "value": int(M["og_permit"].sum()), "subset": scope[M["og_permit"]], "cls": "irup-card"},
        {"label": "Ready for extraction", "value": int(M["rfe"].sum()), "subset": scope[M["rfe"]], "cls": "irup-card"},
        {"label": "Started extraction", "value": int(M["started"].sum()), "subset": scope[M["started"]], "cls": "irup-card"},
        {"label": "Done extraction", "value": int(M["done"].sum()), "subset": scope[M["done"]], "cls": "irup-card"},
    ], labels), unsafe_allow_html=True)
with right:
    st.markdown('<div class="irup-h" style="font-size:16px">Status breakdown</div>', unsafe_allow_html=True)
    st.markdown(status_html(filtered, labels), unsafe_allow_html=True)
st.caption("Hover over any number to see which sites it covers.")

# mapped sites (for the map below)
mapped = filtered.dropna(subset=["LATITUDE", "LONGITUDE"]).copy() if {"LATITUDE", "LONGITUDE"}.issubset(filtered.columns) else filtered.iloc[0:0].copy()

# ---------------- Travel time & distance between two sites ----------------

def haversine_km(lat1, lon1, lat2, lon2):
    r = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = p2 - p1
    dlmb = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlmb / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))

def fmt_duration(seconds):
    minutes = int(round(seconds / 60))
    h, mnt = divmod(minutes, 60)
    return f"{h} hr {mnt} min" if h else f"{mnt} min"

@st.cache_data(show_spinner=False, ttl=3600)
def get_driving_route(lat1, lon1, lat2, lon2):
    """Road route from the free public OSRM server (no API key). Returns (route_dict, error_text)."""
    url = (f"https://router.project-osrm.org/route/v1/driving/"
           f"{lon1},{lat1};{lon2},{lat2}?overview=full&geometries=geojson")
    try:
        resp = requests.get(url, timeout=12, headers={"User-Agent": "IRUP-Project-Map"})
        data = resp.json()
        if data.get("code") != "Ok" or not data.get("routes"):
            return None, data.get("message") or data.get("code") or "No route found"
        r = data["routes"][0]
        return {"distance_km": r["distance"] / 1000.0, "duration_s": r["duration"],
                "coords": [(c[1], c[0]) for c in r["geometry"]["coordinates"]]}, None
    except Exception as exc:
        return None, str(exc)

route = None
route_points = None
_sites = mapped.copy()
for _c in ["LATITUDE", "LONGITUDE"]:
    _sites[_c] = pd.to_numeric(_sites[_c], errors="coerce")
_sites = _sites.dropna(subset=["LATITUDE", "LONGITUDE"])

with st.expander("🚗 Travel Time & Distance", expanded=False):
    if len(_sites) < 2:
        st.caption("Need at least two sites with coordinates in the current filter to calculate travel.")
    else:
        def _label(i, r):
            no = str(r.get("SITE NO", "")).strip()
            no = "" if no.lower() in ("nan", "none") else no
            name = str(r.get("SITE NAME", "Unnamed site")).strip()
            return f"{no} — {name}" if no else name
        route_labels = {i: _label(i, r) for i, r in _sites.iterrows()}
        ids = list(route_labels.keys())
        t1, t2 = st.columns(2)
        from_id = t1.selectbox("From site", ids, index=None, format_func=lambda i: route_labels[i],
                               placeholder="Choose starting site", key="route_from")
        to_id = t2.selectbox("To site", ids, index=None, format_func=lambda i: route_labels[i],
                             placeholder="Choose destination site", key="route_to")
        if from_id is not None and to_id is not None:
            if from_id == to_id:
                st.warning("Choose two different sites.")
            else:
                a, b = _sites.loc[from_id], _sites.loc[to_id]
                lat1, lon1, lat2, lon2 = float(a["LATITUDE"]), float(a["LONGITUDE"]), float(b["LATITUDE"]), float(b["LONGITUDE"])
                straight = haversine_km(lat1, lon1, lat2, lon2)
                with st.spinner("Calculating route..."):
                    route, route_err = get_driving_route(lat1, lon1, lat2, lon2)
                route_points = ((lat1, lon1), (lat2, lon2))
                m1, m2, m3 = st.columns(3)
                if route:
                    m1.metric("Driving distance", f"{route['distance_km']:,.1f} km")
                    m2.metric("Est. driving time", fmt_duration(route["duration_s"]))
                    m3.metric("Straight-line distance", f"{straight:,.1f} km")
                    st.caption("Road route from OpenStreetMap data (OSRM). Time is based on typical road speeds and "
                               "does not include traffic, stops, or ferry crossings.")
                else:
                    est_road = straight * 1.3
                    m1.metric("Straight-line distance", f"{straight:,.1f} km")
                    m2.metric("Rough road distance (est.)", f"{est_road:,.1f} km")
                    m3.metric("Rough driving time (est.)", fmt_duration(est_road / 40 * 3600))
                    st.warning(f"Could not get a road route ({route_err}). Showing a rough estimate: straight-line "
                               "distance x 1.3 at an average of 40 km/h. The sites may be on different islands, "
                               "or the routing service may be unreachable.")


st.subheader("Project Site Map")
map_height = st.slider("Map height (pixels)", min_value=600, max_value=1200, value=850, step=50)
st.caption("For an even larger view, use your browser's full-screen mode (F11).")

def milestone_key(value):
    if pd.isna(value):
        return ""
    return " ".join(str(value).upper().split())

def pin_color(value):
    return MILESTONE_COLORS.get(milestone_key(value), UNKNOWN_COLOR)

def make_pin(color):
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="26" height="36" viewBox="0 0 26 36">'
        f'<path d="M13 1C6.4 1 1 6.3 1 12.8c0 8.6 12 22.2 12 22.2s12-13.6 12-22.2C25 6.3 19.6 1 13 1z" '
        f'fill="{color}" stroke="#ffffff" stroke-width="2"/>'
        f'<circle cx="13" cy="12.8" r="4.5" fill="#ffffff" fill-opacity="0.9"/></svg>'
    )
    html = f'<div class="irup-pin">{svg}</div>'
    return folium.DivIcon(html=html, icon_size=(26, 36), icon_anchor=(13, 36), popup_anchor=(0, -32))

def fmt(value, percent=False):
    """Readable text for popup values: no .0 on whole numbers, dates without time, HTML-safe."""
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return "N/A"
    if isinstance(value, str):
        text = value.replace("\xa0", " ").strip()
        if not text or text.upper() in ("NAN", "NAT", "NONE"):
            return "N/A"
        if text[:10].count("-") == 2 and text[10:].strip() in ("", "00:00:00"):
            text = text[:10]
        return html.escape(text)
    if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
        return value.strftime("%Y-%m-%d")
    try:
        n = float(value)
    except Exception:
        return html.escape(str(value))
    if percent:
        return f"{n:.1%}" if abs(n) <= 1 else f"{n:.1f}%"
    return f"{n:,.0f}" if n == int(n) else f"{n:,.2f}"

def popup_row(label, value, percent=False):
    return f"<b>{label}:</b> {fmt(value, percent)}<br>"

def popup_row_if_filled(label, value, max_len=None):
    text = fmt(value)
    if text == "N/A":
        return ""
    if max_len and len(text) > max_len:
        text = text[:max_len].rstrip() + "…"
    return f"<b>{label}:</b> {text}<br>"

def build_popup(row):
    total_mh = row.get("TOTAL MANHOLE", row.get("NO. OF MH ACTUAL"))
    main = "".join([
        f'<div style="font-size:14px;margin-bottom:4px;"><b>{fmt(row.get("SITE NAME", "Unnamed site"))}</b></div>',
        popup_row("Site No", row.get("SITE NO")),
        popup_row("Area", row.get("AREA")),
        f'<b>City/Province:</b> {fmt(row.get("CITY"))}, {fmt(row.get("PROVINCE"))}<br>',
        popup_row("APM", row.get("APM ASSIGNED")),
        popup_row("Overall Milestone", row.get("OVERALL MILESTONE")),
        popup_row("Overall Progress", row.get("OVERALL PROGRESS"), percent=True),
        popup_row("Total of MH", total_mh),
        popup_row("Accessible MH", row.get("MH ACCESSIBLE")),
        popup_row("N/ Accessible MH", row.get("MH N/ ACCESSIBLE")),
        popup_row("Approx. Tons", row.get("APPROX TONS")),
        popup_row("Design Tons", row.get("DESIGN TONS")),
        popup_row("ISP/OSP", row.get("ISP/OSP")),
        popup_row("ISP Extracted", row.get("ISP EXTRACTED")),
        popup_row("OSP Extracted", row.get("OSP EXTRACTED")),
        popup_row("Total Tons Extracted", row.get("TOTAL TONS EXTRACTED")),
        popup_row("Extraction Status", row.get("EXTRACTION STATUS")),
    ])
    return (f'<div style="font-family:Arial;font-size:12px;min-width:240px;max-height:340px;overflow-y:auto;">'
            f'{main}</div>')

class LegendControl(MacroElement):
    """Legend as a Leaflet control inside the map container, so it stays visible in full screen."""
    def __init__(self, html):
        super().__init__()
        self._html = json.dumps(html)
        self._template = Template("""
        {% macro script(this, kwargs) %}
        (function () {
            var legend = L.control({position: "bottomleft"});
            legend.onAdd = function () {
                var div = L.DomUtil.create("div", "irup-legend");
                div.innerHTML = {{ this._html }};
                L.DomEvent.disableClickPropagation(div);
                L.DomEvent.disableScrollPropagation(div);
                return div;
            };
            legend.addTo({{ this._parent.get_name() }});
        })();
        {% endmacro %}
        """)

class PinZoomScaler(MacroElement):
    """Shrinks/grows the pins as the map is zoomed out/in."""
    _template = Template("""
        {% macro script(this, kwargs) %}
        (function () {
            var map = {{ this._parent.get_name() }};
            function scalePins() {
                var f = Math.max(0.3, Math.min(1.2, 0.3 + (map.getZoom() - 5) * 0.17));
                document.documentElement.style.setProperty("--pin-scale", f);
            }
            map.on("zoomend", scalePins);
            scalePins();
        })();
        {% endmacro %}
    """)

PIN_CSS = "<style>.irup-pin{transform:scale(var(--pin-scale,0.64));transform-origin:50% 100%;}</style>"

def build_legend(counts):
    rows = ""
    for name, color in MILESTONE_COLORS.items():
        rows += (
            f'<div style="display:flex;align-items:center;margin:3px 0;">'
            f'<span style="display:inline-block;width:14px;height:14px;border-radius:50%;'
            f'background:{color};border:1px solid #fff;box-shadow:0 0 2px #555;margin-right:8px;"></span>'
            f'<span style="flex:1;">{name.title()}</span>'
            f'<span style="margin-left:10px;color:#666;">{counts.get(name, 0)}</span></div>'
        )
    if counts.get("OTHER", 0):
        rows += (
            f'<div style="display:flex;align-items:center;margin:3px 0;">'
            f'<span style="display:inline-block;width:14px;height:14px;border-radius:50%;'
            f'background:{UNKNOWN_COLOR};border:1px solid #fff;box-shadow:0 0 2px #555;margin-right:8px;"></span>'
            f'<span style="flex:1;">Other / Blank</span>'
            f'<span style="margin-left:10px;color:#666;">{counts["OTHER"]}</span></div>'
        )
    return (
        '<div style="background:rgba(255,255,255,0.95);'
        'padding:10px 14px;border-radius:8px;box-shadow:0 1px 6px rgba(0,0,0,0.4);'
        'font-family:Arial,sans-serif;font-size:12px;color:#222;min-width:200px;">'
        '<div style="font-weight:bold;margin-bottom:6px;font-size:13px;">Overall Milestone</div>'
        f'{rows}</div>'
    )

if mapped.empty:
    st.warning("No sites with valid latitude/longitude in the current filter. Check the coordinates or relax the filters.")
else:
    mapped["LATITUDE"] = pd.to_numeric(mapped["LATITUDE"], errors="coerce")
    mapped["LONGITUDE"] = pd.to_numeric(mapped["LONGITUDE"], errors="coerce")
    mapped = mapped.dropna(subset=["LATITUDE", "LONGITUDE"])
    center = [mapped["LATITUDE"].mean(), mapped["LONGITUDE"].mean()]
    # Standard OpenStreetMap tiles only; no API key is required.
    m = folium.Map(location=center, zoom_start=7, tiles="OpenStreetMap", control_scale=True)

    # Individual pins added straight to the map (no clustering), so no numbered bubbles when zoomed out.
    legend_counts = {}
    for _, row in mapped.iterrows():
        site_no = row.get("SITE NO", "N/A")
        site_name = row.get("SITE NAME", "Unnamed site")
        milestone = row.get("OVERALL MILESTONE", "N/A")
        key = milestone_key(milestone)
        legend_counts[key if key in MILESTONE_COLORS else "OTHER"] = legend_counts.get(key if key in MILESTONE_COLORS else "OTHER", 0) + 1
        popup_html = build_popup(row)
        folium.Marker(
            location=[float(row["LATITUDE"]), float(row["LONGITUDE"])],
            icon=make_pin(pin_color(milestone)),
            popup=folium.Popup(popup_html, max_width=360),
            tooltip=f"{site_no} — {site_name}",
        ).add_to(m)

    # Permanent legend (always visible, shows all milestones with counts for the current filter).
    # Route line + fit map to the selected pair of sites.
    if route_points:
        if route:
            folium.PolyLine(route["coords"], color="#1565C0", weight=5, opacity=0.85).add_to(m)
        else:
            folium.PolyLine(list(route_points), color="#1565C0", weight=4, opacity=0.8, dash_array="8 8").add_to(m)
        pts = list(route_points) + (route["coords"] if route else [])
        m.fit_bounds([[min(p[0] for p in pts), min(p[1] for p in pts)],
                      [max(p[0] for p in pts), max(p[1] for p in pts)]], padding=(40, 40))

    LegendControl(build_legend(legend_counts)).add_to(m)

    # Pins shrink when zooming out and grow when zooming in.
    m.get_root().header.add_child(folium.Element(PIN_CSS))
    PinZoomScaler().add_to(m)

    # Full-screen button (top-right of the map).
    Fullscreen(position="topright", title="Full screen", title_cancel="Exit full screen",
               force_separate_button=True).add_to(m)
    st_folium(m, use_container_width=True, height=map_height, returned_objects=[])

# ---------------- Progress updates (design / permit / extraction) ----------------
st.markdown('<div class="irup-h">Progress Updates</div>', unsafe_allow_html=True)
tab_d, tab_p, tab_e = st.tabs(["Design update", "Permit update", "Extraction update"])
cancelled_sites = stage_mask(scope, {"col": "OVERALL MILESTONE", "values": ["CANCELLED"]})
active = scope[~cancelled_sites]          # cancelled sites are not part of any target or gap
if cancelled_sites.any():
    st.caption(f"{int(cancelled_sites.sum())} cancelled site(s) are not counted in the targets, gaps and timelines below.")
PINK = "background:#FCE4D6"
CREAM = "background:#FFF2CC"

def week_block(title, weeks, done_mask, key):
    st.markdown(f'<div class="irup-h" style="font-size:16px">{title}</div>', unsafe_allow_html=True)
    show_all = st.checkbox("Show all weeks", key=key)
    st.markdown(timeline_html(active, weeks.loc[active.index], done_mask.loc[active.index], labels, show_all),
                unsafe_allow_html=True)

with tab_d:
    gap_sites = active[~M["design_approved"].loc[active.index]]
    st.markdown('<div class="irup-sub">Design milestone</div>', unsafe_allow_html=True)
    st.markdown(cards_row([
        {"label": "OG survey", "value": int(M["og_survey"].sum()), "subset": scope[M["og_survey"]]},
        {"label": "Survey done", "value": int(M["survey_done"].sum()), "subset": scope[M["survey_done"]]},
        {"label": "Completed", "value": int(M["design_completed"].sum()), "subset": scope[M["design_completed"]], "style": CREAM},
        {"label": "Submitted", "value": int(M["design_submitted"].sum()), "subset": scope[M["design_submitted"]]},
        {"label": "Approved", "value": int(M["design_approved"].sum()), "subset": scope[M["design_approved"]]},
        {"label": "Design gap", "value": len(gap_sites), "subset": gap_sites, "val_style": "color:#C00000"},
    ], labels), unsafe_allow_html=True)
    st.markdown('<div class="irup-h" style="font-size:16px">Approved design per region</div>', unsafe_allow_html=True)
    st.markdown(region_cards(scope, M["design_approved"], labels), unsafe_allow_html=True)
    week_block("Design timeline",
               week_tuples(scope, date_cols=["DESIGN REVISED TARGET COMPLETION", "DESIGN ORIGINAL TARGET"],
                           week_cols=["DESIGN WEEK TARGET"]),
               M["design_completed"], "tl_design")
    st.markdown('<div class="irup-h" style="font-size:16px">Design progress</div>', unsafe_allow_html=True)
    keys = region_key(scope)
    rows = []
    for r in region_names(scope):
        in_r = keys == r
        rows.append([(r, None, f"background:{REGION_COLORS.get(r, '#546E7A')}", "irup-first")] +
                    [(int((M[k] & in_r).sum()), scope[M[k] & in_r]) for k in
                     ["og_survey", "survey_done", "design_completed", "design_submitted", "design_approved"]])
    rows.append([("TOTAL", None, "font-weight:700")] + [(int(M[k].sum()), scope[M[k]], "font-weight:700") for k in
                ["og_survey", "survey_done", "design_completed", "design_submitted", "design_approved"]])
    st.markdown(html_table(["Region", "OG survey", "Survey done", "Comp", "Submitted", "Approved"], rows, labels), unsafe_allow_html=True)

with tab_p:
    permit_gap = active[~M["rfe"].loc[active.index]]
    st.markdown('<div class="irup-sub">Permitting milestone</div>', unsafe_allow_html=True)
    st.markdown(cards_row([
        {"label": "Target (sites)", "value": len(active), "subset": active, "style": PINK},
        {"label": "Ready for permitting", "value": int(M["design_approved"].sum()), "subset": scope[M["design_approved"]], "style": CREAM},
        {"label": "On going", "value": int(M["og_permit"].sum()), "subset": scope[M["og_permit"]]},
        {"label": "RFE", "value": int(M["rfe"].sum()), "subset": scope[M["rfe"]]},
        {"label": "Permit gap", "value": len(permit_gap), "subset": permit_gap, "val_style": "color:#C00000"},
    ], labels), unsafe_allow_html=True)
    st.markdown('<div class="irup-h" style="font-size:16px">Permit secured per region (RFE)</div>', unsafe_allow_html=True)
    st.markdown(region_cards(scope, M["rfe"], labels), unsafe_allow_html=True)
    week_block("Permit timeline",
               week_tuples(scope, date_cols=["NEW PERMITTING TARGET", "PERMITTING TARGET COMPLETION"],
                           week_cols=["PERMITTING WEEK TARGET"]),
               M["rfe"], "tl_permit")
    st.markdown('<div class="irup-h" style="font-size:16px">Permit progress</div>', unsafe_allow_html=True)
    permit_rows = []
    for name, sub_col, sec_col in [
        ("Barangay", "BRGY SUBMITTED", "BRGY PERMIT SECURED"), ("LGU", "LGU PERMIT SUBMITTED", "LGU PERMIT SECURED"),
        ("Traffic", "TRAFFIC submitted", "TRAFFIC PERMIT SECURED"), ("DPWH", "DPWH submitted", "DPWH SECURED"),
        ("HOA", "HOA SUBMITTED", "HOA SECURED"), ("Developer", "DEVELOPER SUBMITTED", "DEVELOPER secured"),
    ]:
        sub, sec = num(scope, sub_col), num(scope, sec_col)
        if sub.sum() == 0 and sec.sum() == 0:
            continue
        applied, still_og = scope[sub > 0], scope[sub > sec]
        permit_rows.append([(name, None, "font-weight:700"), (len(applied), applied), (len(still_og), still_og),
                            (f"{sub.sum():,.0f}", None), (f"{sec.sum():,.0f}", None)])
    if permit_rows:
        st.markdown(html_table(["Permit", "Sites applied", "Sites still on going", "Permits submitted", "Permits secured"],
                               permit_rows, labels), unsafe_allow_html=True)
        st.caption("'Still on going' = more permits submitted than secured for that site. Counts come from the submitted/secured columns of the tracker.")
    else:
        st.caption("No permit application data in the current results.")

with tab_e:
    ext_gap = active[(M["rfe"] & ~M["started"]).loc[active.index]]
    st.markdown('<div class="irup-sub">Extraction milestone</div>', unsafe_allow_html=True)
    st.markdown(cards_row([
        {"label": "Started", "value": int(M["started"].sum()), "subset": scope[M["started"]]},
        {"label": "Done", "value": int(M["done"].sum()), "subset": scope[M["done"]]},
        {"label": "ISP tons", "value": f"{num(scope, 'ISP EXTRACTED').sum():,.2f}", "style": PINK},
        {"label": "OSP tons", "value": f"{num(scope, 'OSP EXTRACTED').sum():,.2f}", "style": PINK},
        {"label": "Total tons", "value": f"{num(scope, 'TOTAL TONS EXTRACTED').sum():,.2f}"},
        {"label": "Gap (RFE not started)", "value": len(ext_gap), "subset": ext_gap, "val_style": "color:#C00000"},
    ], labels), unsafe_allow_html=True)
    st.markdown('<div class="irup-h" style="font-size:16px">Extraction done per region</div>', unsafe_allow_html=True)
    st.markdown(region_cards(scope, M["done"], labels), unsafe_allow_html=True)
    week_block("Extraction start timeline",
               week_tuples(scope, date_cols=["REVISED EXTRACTION START DATE", "EXTRACTION TARGET START DATE"],
                           week_cols=["REVISED START WEEK", "EXTRACTION TARGET START WEEK"]),
               M["started"], "tl_ext_start")
    week_block("Extraction end timeline",
               week_tuples(scope, date_cols=["REVISED EXTRACTION DONE DATE", "EXTRACTION TARGET DONE DATE"], week_cols=["EXTRACTION DONE TARGET WEEK"]),
               M["done"], "tl_ext_end")


st.divider()
st.subheader("Site List")
st.caption(f"{len(filtered):,} site(s) match the current search and filters. Tabs follow the sections of the Excel tracker; "
           "the first two columns stay visible while you scroll sideways.")

# ---- column groups, in the same order as the Excel tracker ----
GROUP_COLUMNS = {
    "Overview": [
        "SITE NO", "SITE NAME", "AREA", "CITY", "PROVINCE", "QTR", "APM ASSIGNED", "STATUS", "SUBCON",
        "OVERALL MILESTONE", "OVERALL PROGRESS", "APPROX TONS", "DESIGN TONS", "TOTAL TONS EXTRACTED",
        "LATITUDE", "LONGITUDE",
    ],
    "Survey & Design": [
        "SURVEY START", "SURVEY DONE DATE", "DESIGN STATUS", "DESIGN WEEK TARGET", "DESIGN ORIGINAL TARGET",
        "DESIGN REVISED TARGET COMPLETION", "DESIGN COMPLETION", "DESIGN SUBMISSION", "DESIGN APPROVED",
        "DESIGN PROGRESS", "ISP/OSP", "TOTAL WORKING DAYS", "NO. OF MH ACTUAL", "MH ACCESSIBLE",
        "MH N/ ACCESSIBLE", "ASPHALTED", "CEMENTED", "BRICKS", "MH INSIDE BLDG", "DESIGN REMARKS/ ISSUE",
    ],
    "Permitting": [
        "PERMITTING GENERAL STAUS", "PERMITTING PROGESS", "PERMITTING WEEK TARGET", "PERMITTING TARGET COMPLETION",
        "NEW PERMITTING TARGET",
        "HOA SUBMITTED", "HOA SECURED", "HOA START", "HOA END",
        "DEVELOPER SUBMITTED", "DEVELOPER secured", "DEVELOPER START", "DEVELOPER END",
        "BRGY SUBMITTED", "BRGY PERMIT SECURED", "BARANGAY START", "BRGY END",
        "LGU PERMIT SUBMITTED", "LGU PERMIT SECURED", "LGU START", "LGU END",
        "TRAFFIC submitted", "TRAFFIC PERMIT SECURED", "TRAFFIC START", "TRAFFIC END",
        "DPWH submitted", "DPWH SECURED", "DPWH START", "DPWH END", "PERMITTING ISSUE / REMARKS",
    ],
    "Extraction": [
        "EXTRACTION STATUS", "EXTRACTION  PROGRESS", "EXTRACTION TARGET START WEEK", "EXTRACTION TARGET START DATE",
        "REVISED START WEEK", "REVISED EXTRACTION START DATE", "EXTRACTION DONE TARGET WEEK",
        "EXTRACTION TARGET DONE DATE", "REVISED EXTRACTION DONE DATE", "REVISION REMARKS", "EXTRACTION START",
        "EXCAVATION TEAM QUANTITY", "EXTRACTION TEAM QUANTITY", "EXTRACTION DONE DAYS", "PS NAME",
        "TOTAL MANHOLE", "CIVIL WORKS MH", "DEWATERING MH", "NO DEWATERING MH", "TOTAL MANHOLE EXTRACTED",
        "MANHOLE EXTRACTED GAP", "ISP EXTRACTED", "OSP EXTRACTED", "TOTAL TONS EXTRACTED", "EXTRACTION DONE",
        "EXTRACTION REMARKS",
    ],
}
LONG_TEXT_COLUMNS = ["DESIGN REMARKS/ ISSUE", "PERMITTING ISSUE / REMARKS", "REVISION REMARKS", "EXTRACTION REMARKS"]

def prep_display(frame, columns):
    """Frame ready for st.dataframe: dates as YYYY-MM-DD, progress as %, mixed cells as clean text."""
    out = frame[columns].copy()
    config = {}
    for c in out.columns:
        col = out[c]
        if col.dtype == object:
            as_num = pd.to_numeric(col, errors="coerce")
            if as_num.notna().sum() == col.notna().sum() and col.notna().any():
                col = as_num          # object column that only holds numbers (e.g. after the header row was removed)
                out[c] = col
        if pd.api.types.is_numeric_dtype(col) and not pd.api.types.is_bool_dtype(col):
            if is_progress_col(c) and col.dropna().abs().max() <= 1.5:
                out[c] = col * 100
                config[c] = st.column_config.NumberColumn(c, format="%.1f%%")
            continue
        if pd.api.types.is_datetime64_any_dtype(col):
            out[c] = col.dt.strftime("%Y-%m-%d").fillna("")
        else:
            out[c] = col.map(cell_text)
    return out, config

def show_table(columns, height=460):
    cols = [c for c in columns if c in filtered.columns]
    if not cols:
        st.info("None of these columns exist in the uploaded tracker.")
        return
    key_cols = [c for c in ["SITE NO", "SITE NAME"] if c in filtered.columns]
    rest = [c for c in cols if c not in key_cols]
    table, config = prep_display(filtered, key_cols + rest)
    if key_cols:
        table = table.set_index(key_cols)   # index columns stay frozen on the left while scrolling
    st.dataframe(table, use_container_width=True, height=height, column_config=config)

excel_columns = [c for c in df.columns
                 if not str(c).startswith("Unnamed") and df[c].notna().any()]

tab_names = ["Overview", "Survey & Design", "Permitting", "Extraction", "All columns (Excel view)", "Site breakdown"]
t_over, t_design, t_permit, t_extract, t_all, t_site = st.tabs(tab_names)

with t_over:
    overview_cols = GROUP_COLUMNS["Overview"] + [c for c in selected_stage_cols if c not in GROUP_COLUMNS["Overview"]]
    show_table(overview_cols)
with t_design:
    show_table(GROUP_COLUMNS["Survey & Design"])
with t_permit:
    show_table(GROUP_COLUMNS["Permitting"])
with t_extract:
    show_table(GROUP_COLUMNS["Extraction"])
with t_all:
    st.caption("Every column of the IRUP SITELIST sheet, in the same order as the Excel file.")
    show_table(excel_columns, height=520)

with t_site:
    if filtered.empty:
        st.info("No sites match the current filters.")
    else:
        def _site_label(i, r):
            no = cell_text(r.get("SITE NO"))
            name = cell_text(r.get("SITE NAME")) or "Unnamed site"
            return f"{no} — {name}" if no else name
        site_pick_labels = {i: _site_label(i, r) for i, r in filtered.iterrows()}
        pick = st.selectbox("Choose a site", list(site_pick_labels.keys()), format_func=lambda i: site_pick_labels[i],
                            key="site_breakdown_pick")
        row = filtered.loc[pick]

        def section_table(title, columns):
            rows = []
            for c in columns:
                if c in LONG_TEXT_COLUMNS or c not in filtered.columns:
                    continue
                v = row[c]
                if is_progress_col(c) and isinstance(v, (int, float)) and not pd.isna(v) and abs(v) <= 1.5:
                    text = f"{v:.1%}"
                else:
                    text = cell_text(v)
                if text and text != "0":
                    rows.append({"Field": c.replace("  ", " ").title(), "Value": text})
            st.markdown(f"**{title}**")
            if rows:
                st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True,
                             height=min(35 * len(rows) + 38, 420))
            else:
                st.caption("No data filled in for this section.")

        left, right = st.columns(2)
        with left:
            section_table("Overview", GROUP_COLUMNS["Overview"])
            section_table("Survey & Design", GROUP_COLUMNS["Survey & Design"])
        with right:
            section_table("Permitting", GROUP_COLUMNS["Permitting"])
            section_table("Extraction", GROUP_COLUMNS["Extraction"])

        remark_cols = [c for c in LONG_TEXT_COLUMNS if c in filtered.columns and cell_text(row[c])]
        if remark_cols:
            st.markdown("**Remarks / Issues**")
            for c in remark_cols:
                with st.expander(c.title()):
                    st.text_area(c, value=cell_text(row[c]), height=220, disabled=True,
                                 key=f"remarks_{pick}_{c}", label_visibility="collapsed")

# ---- downloads ----
export_df = filtered[excel_columns].copy()
csv_data = export_df.to_csv(index=False).encode("utf-8-sig")
xlsx_buffer = io.BytesIO()
with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
    export_df.to_excel(writer, sheet_name="IRUP SITELIST (filtered)", index=False)
d1, d2, _ = st.columns([1, 1, 3])
d1.download_button("Download as Excel", xlsx_buffer.getvalue(), "irup_filtered_sites.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
d2.download_button("Download as CSV", csv_data, "irup_filtered_sites.csv", "text/csv")

st.caption("Map tiles require an internet connection. The app filters the workbook you upload and does not edit the original Excel file.")
