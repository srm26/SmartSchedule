# -*- coding: utf-8 -*-
"""
GES SmartSchedule - AI-Powered Electrical Labor Calendar
The page shows every schedule file in the SharePoint calendar folders; the sidebar picks the calendar to view.
The 15-minute Outlook sync runs on its own (sync_engine.py, started by startup.sh).
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import calendar as pycal
from datetime import timedelta
from html import escape

from schedule_config import CALENDARS, FACILITIES, OT_RULES, HOLIDAYS, DT_RULES, CONFIG_WARNINGS, outlook_categories
from schedule_logic import compute_overtime, ot_summary
from sync_engine import (graph_secrets_ok, get_graph_token, load_all, duplicate_ids, route_rows,
                         facility_warnings, sync_rows, remove_deleted_shifts, ensure_outlook_categories,
                         sweep_status, sweep_test_rows, set_sweep_test_rows, today, SYNC_MINUTES,
                         read_calendar_events, events_on_day)
import requests

# ─── PAGE CONFIG ─────────────────────────────────────────────────────
st.set_page_config(
    page_title="GES SmartSchedule",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  [data-testid="stSidebar"] { background: #1a1f36; }
  [data-testid="stSidebar"] * { color: #e2e8f0 !important; }
  [data-testid="stSidebar"] input { color: #1e293b !important; }
  [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {
    background: #2d3452; border: 1px solid #64748b; }
  [data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover {
    background: #3b4470; border-color: #e2e8f0; }
  [data-testid="stSidebar"] .stRadio label { font-size: 15px; padding: 6px 0; }
  /* Status boxes in the sidebar ("No file loaded", "Last check", file name): white box, dark text */
  [data-testid="stSidebar"] [data-testid="stAlertContainer"] {
    background: #fff !important; border: 1px solid #cbd5e1; border-left: 4px solid #f59e0b; }
  [data-testid="stSidebar"] [data-testid="stAlertContainer"] * { color: #1e293b !important; }
  .metric-card {
    background:#fff; border:1px solid #e2e8f0; border-radius:12px;
    padding:20px 24px; box-shadow:0 1px 4px rgba(0,0,0,.06);
  }
  .metric-label { font-size:12px; color:#64748b; font-weight:700;
    letter-spacing:.5px; text-transform:uppercase; margin-bottom:4px; }
  .metric-value { font-size:30px; font-weight:800; color:#1e293b; }
  .metric-delta { font-size:12px; color:#22c55e; margin-top:3px; }
  .section-title {
    font-size:19px; font-weight:700; color:#1e293b;
    border-left:4px solid #f59e0b; padding-left:10px; margin:20px 0 12px;
  }
  .fn-badge {
    display:inline-block; padding:2px 10px; border-radius:20px;
    font-size:11px; font-weight:700;
  }
  .fn-set     { background:#dbeafe; color:#1d4ed8; }
  .fn-strike  { background:#fee2e2; color:#b91c1c; }
  .fn-standby { background:#fef9c3; color:#854d0e; }
  .notif-card {
    background:#f8fafc; border-left:4px solid #f59e0b;
    border-radius:6px; padding:12px 16px; margin-bottom:8px;
  }
  .ai-banner {
    background:linear-gradient(135deg,#1e293b,#334155);
    border-radius:12px; padding:18px 22px; color:white; margin:14px 0;
  }
  div[data-testid="stTabs"] button { font-size:15px; font-weight:600; padding:10px 20px; }
</style>
""", unsafe_allow_html=True)

# ─── CONSTANTS ───────────────────────────────────────────────────────
FUNCTION_COLORS = {"SET": "#3b82f6", "STRIKE": "#ef4444", "STANDBY": "#f59e0b"}
ALL_VIEW        = "All calendars"

# ─── DATA HELPERS ────────────────────────────────────────────────────
def build_notifications(df, demo_contacts):
    rows = []
    union_email   = demo_contacts.get("union_email",   "union-dispatch@ibew.org")
    foreman_email = demo_contacts.get("foreman_email", "foreman@ges.com")
    for _, row in df.iterrows():
        date_str    = row["DATE"].strftime("%Y-%m-%d")
        notify_date = (row["DATE"] - timedelta(days=2)).strftime("%Y-%m-%d")
        show = row["SHOW"]; fn = row["FUNCTION"]
        loc  = row["LOCATION"]; time_s = row["TIME"]; foreman = row["FOREMAN"]
        rows.append({
            "Type": "Labor Call", "Recipient": f"Union Dispatch ({union_email})",
            "Show": show, "Date": date_str, "Notify By": notify_date,
            "Message": (f"Labor call for {int(row['TOTAL_CREW'])} electrician(s). "
                        f"Show: {show} | {fn} | {date_str} {time_s} | {loc}. Foreman: {foreman}."),
            "Channel": "Email", "Status": "Ready",
        })
        if foreman not in ("TBD", "nan", ""):
            rows.append({
                "Type": "Foreman Assignment", "Recipient": f"{foreman} ({foreman_email})",
                "Show": show, "Date": date_str, "Notify By": notify_date,
                "Message": f"You are Foreman for {show} | {fn} | {date_str} {time_s} | {loc}.",
                "Channel": "Email", "Status": "Ready",
            })
    return pd.DataFrame(rows)

def _build_ics(df_shifts, cal_name):
    if df_shifts.empty:
        return b""
    lines = [
        "BEGIN:VCALENDAR", "VERSION:2.0",
        f"PRODID:-//GES SmartSchedule//{cal_name}//EN",
        f"X-WR-CALNAME:GES {cal_name}",
    ]
    for _, r in df_shifts.iterrows():
        start_str = r["DATE"].strftime("%Y%m%d")
        end_val   = r["END_DATE"] if pd.notna(r["END_DATE"]) else r["DATE"]
        end_str   = (pd.Timestamp(end_val) + timedelta(days=1)).strftime("%Y%m%d")
        lines += [
            "BEGIN:VEVENT",
            f"DTSTART;VALUE=DATE:{start_str}",
            f"DTEND;VALUE=DATE:{end_str}",
            f"SUMMARY:[GES] {r['SHOW']} - {r['FUNCTION']}",
            f"DESCRIPTION:{r['TIME']} | {r['LOCATION']} | Foreman: {r['FOREMAN']}",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines).encode()

# ─── LOAD EVERY SCHEDULE FROM SHAREPOINT ─────────────────────────────
@st.cache_data(ttl=300, show_spinner="Reading the schedules from SharePoint...")
def _load_sharepoint():
    """(files, {path: df}, {path: error}). Also loads the admin settings into schedule_config."""
    files, parsed, errors, _ = load_all(get_graph_token())
    return files, parsed, errors

@st.cache_data(ttl=60, show_spinner="Reading the Outlook calendar...")
def _read_outlook(mailbox, start, end, time_zone):
    """Events in the calendar's mailbox (read-only), kept 1 minute; 'Reload from Outlook' clears it."""
    return read_calendar_events(get_graph_token(), mailbox, start, end, time_zone)

# Outlook color category (preset) -> event fill on the Outlook Calendar page
OUTLOOK_FILL = {"preset0": "#f7b9b9", "preset1": "#fcd3a5", "preset2": "#dcc0a3", "preset3": "#fbeea0",
                "preset4": "#bfe5bf", "preset5": "#a9e2df", "preset6": "#d8db9f", "preset7": "#b7d5fb",
                "preset8": "#d9c2f3", "preset9": "#f1aac4", "preset12": "#d9d9d9"}
DEFAULT_FILL = "#dbe4f0"

if not graph_secrets_ok():
    st.error("Microsoft Graph credentials are not configured (GRAPH_APP_ID, GRAPH_TENANT_ID, GRAPH_CLIENT_SECRET), "
             "so SharePoint can't be read.")
    st.stop()
try:
    sp_files, sp_parsed, sp_errors = _load_sharepoint()
except Exception as e:
    st.error(f"SharePoint error: {e}")
    st.stop()

for _k, _v in [("notifs", None), ("ai_done", False), ("view_shown", None)]:
    if _k not in st.session_state:
        st.session_state[_k] = _v
# Re-assign every run so the nav radio keeps its page
st.session_state.nav = st.session_state.get("nav", "Dashboard")

# ─── SIDEBAR ─────────────────────────────────────────────────────────
with st.sidebar:
    st.image("GES-logo-white.png", use_container_width=True)
    st.markdown("## GES SmartSchedule")
    st.caption("AI-Powered Electrical Labor Calendar")
    st.divider()

    view = st.selectbox("Calendar", [ALL_VIEW] + list(CALENDARS), key="view")
    if st.button("Refresh from SharePoint", use_container_width=True):
        _load_sharepoint.clear()
        st.rerun()
    st.divider()
    page = st.radio(
        "Navigation",
        ["Dashboard", "Schedule & Push", "Calendar View", "Outlook Calendar", "Crew Lookup", "Overtime"],
        key="nav", label_visibility="collapsed",
    )

    st.divider()
    st.markdown("**Auto-Sync**")
    st.caption(f"Runs every {SYNC_MINUTES} minutes, also when nobody has the app open.")
    _test_rows = sweep_test_rows()
    _test_on = st.checkbox("Sweep test mode", value=_test_rows > 0)
    _n = int(st.number_input("Rows to sync", min_value=1, max_value=50, value=max(_test_rows, 2), step=1)) if _test_on else 0
    if _n != _test_rows:
        set_sweep_test_rows(_n)
    _sts = sweep_status()
    if _sts.get("running"):
        st.info("Syncing now...")
    elif _sts.get("last_check"):
        st.success(f"Last check: {_sts['last_check']}")
        for entry in _sts.get("log", []):
            st.caption(entry)
        if _sts.get("error"):
            st.error(_sts["error"])
    else:
        st.caption("Waiting for first sweep...")

# ─── HEADER ──────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 5])
with col_logo:
    st.image("GES-logo.webp", width=100)
with col_title:
    st.markdown("## GES SmartSchedule")
    st.caption("AI-Powered Electrical Labor Calendar — replacing manual Union Outlook entry")
st.divider()

for w in CONFIG_WARNINGS:  # admin workbook rows that couldn't be read
    st.warning(w)
for name, err in sp_errors.items():
    st.error(f"**{name}** could not be read: {err}")

# ─── The rows of the chosen calendar (or all) ────────────────────────
_frames = [d for d in sp_parsed.values() if view == ALL_VIEW or d["CALENDAR"].iloc[0] == view]
if not _frames:
    st.info(f"No schedule files found for **{view}**. Save the schedule in the calendar's SharePoint folder, "
            f"then click Refresh from SharePoint.")
    st.stop()
df    = pd.concat(_frames, ignore_index=True)
stype = df["SCHEDULE_TYPE"].iloc[0]
view_label = view if view == ALL_VIEW else f"{view} calendar"
if st.session_state.view_shown != view:  # notifications were built for another view
    st.session_state.update(view_shown=view, notifs=None, ai_done=False)

# ═══════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ═══════════════════════════════════════════════════════════════════
if "Dashboard" in page:
    total  = len(df)
    shows  = df["SHOW"].nunique()
    crew   = int(df["TOTAL_CREW"].sum())
    dr_min = df["DATE"].min().strftime("%b %d, %Y")
    dr_max = df["DATE"].max().strftime("%b %d, %Y")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Total Shifts / Events</div>
            <div class="metric-value">{total}</div>
            <div class="metric-delta">{view_label}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Unique Shows</div>
            <div class="metric-value">{shows}</div>
            <div class="metric-delta">Across all locations</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Date Range</div>
            <div class="metric-value" style="font-size:18px">{dr_min}</div>
            <div class="metric-delta">through {dr_max}</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-label">Crew Calls (Total)</div>
            <div class="metric-value">{crew}</div>
            <div class="metric-delta">From crew count column</div>
        </div>""", unsafe_allow_html=True)

    st.divider()
    st.markdown('<div class="section-title">Schedule Files in SharePoint</div>', unsafe_allow_html=True)
    _files = pd.DataFrame([
        {"Calendar": f["calendar"], "File": f["path"], "Modified": f["modified_display"],
         "Shifts": len(sp_parsed[f["path"]]) if f["path"] in sp_parsed else 0,
         "From Today On": int((sp_parsed[f["path"]]["DATE"] >= today()).sum()) if f["path"] in sp_parsed else 0,
         "Status": "Error: " + sp_errors[f["path"]] if f["path"] in sp_errors
                   else ("OK" if f["path"] in sp_parsed else "Not a schedule (no By Day sheet)")}
        for f in sp_files if view == ALL_VIEW or f["calendar"] == view])
    st.dataframe(_files, use_container_width=True, hide_index=True)

    st.divider()
    st.markdown('<div class="section-title">Schedule — Recent 20 Rows</div>', unsafe_allow_html=True)
    recent = df.sort_values("DATE").tail(20)[
        ["DATE","SHOW","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW"]
    ].copy()
    recent["DATE"] = recent["DATE"].dt.strftime("%Y-%m-%d")
    st.dataframe(recent, use_container_width=True, hide_index=True)

    st.divider()
    fn_counts = df["FUNCTION"].value_counts().reset_index()
    fn_counts.columns = ["Function", "Count"]
    fig = px.bar(fn_counts, x="Function", y="Count", color="Function",
                 color_discrete_map=FUNCTION_COLORS, title="Shifts by Function Type")
    fig.update_layout(showlegend=False, height=300, margin=dict(t=40, b=10, l=10, r=10))
    st.plotly_chart(fig, use_container_width=True)


# ═══════════════════════════════════════════════════════════════════
# PAGE: SCHEDULE & PUSH
# ═══════════════════════════════════════════════════════════════════
elif "Schedule" in page:
    with st.expander("Notification Settings", expanded=False):
        dc1, dc2 = st.columns(2)
        with dc1:
            union_email   = st.text_input("Union dispatch email", value="union-dispatch@ibew-local.org")
        with dc2:
            foreman_email = st.text_input("Foreman email",        value="foreman@ges.com")
    demo_contacts = {"union_email": union_email, "foreman_email": foreman_email}

    tab_prev, tab_email, tab_push = st.tabs(["Preview", "Emails", "Push to Outlook"])

    # ── Preview ──────────────────────────────────────────────────────
    with tab_prev:
        st.caption(f"Source: **{view_label}** · {len(df):,} rows · {stype}")
        show_cols = ["DATE","CALENDAR","SHOW","FACILITY","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW"]
        d = df[show_cols].copy()
        d["DATE"] = d["DATE"].dt.strftime("%Y-%m-%d")
        st.dataframe(d, use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("""<div class="ai-banner">
            <b>AI Parsing Engine — Ready</b><br><br>
            Click Preview to review what will be created in Outlook before pushing.
        </div>""", unsafe_allow_html=True)

        if st.button("Preview Schedule", type="primary", use_container_width=True):
            st.session_state.notifs  = build_notifications(df, demo_contacts)
            st.session_state.ai_done = True

        if st.session_state.ai_done and st.session_state.notifs is not None:
            df_n = st.session_state.notifs
            c1, c2, c3 = st.columns(3)
            c1.metric("Calendar Events",     len(df))
            c2.metric("Email Notifications", len(df_n))
            c3.metric("Manual Steps",        "0")
            st.info("Switch to the Push to Outlook tab when ready.")

    # ── Emails ───────────────────────────────────────────────────────
    with tab_email:
        st.markdown('<div class="section-title">Email Notifications</div>', unsafe_allow_html=True)
        if st.session_state.notifs is None:
            st.info("Go to the Preview tab and click Preview Schedule first.")
            if st.button("Quick-generate with default settings"):
                st.session_state.notifs  = build_notifications(df, {})
                st.session_state.ai_done = True
                st.rerun()
        else:
            df_n = st.session_state.notifs
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Emails",    len(df_n))
            c2.metric("Labor Calls",     int((df_n["Type"] == "Labor Call").sum()))
            c3.metric("Foreman Notices", int((df_n["Type"] == "Foreman Assignment").sum()))
            st.divider()
            nf1, nf2 = st.columns(2)
            with nf1:
                type_sel = st.multiselect("Type", df_n["Type"].unique().tolist(),
                                          default=df_n["Type"].unique().tolist(), key="n_type")
            with nf2:
                top_n    = df_n["Show"].value_counts().head(15).index.tolist()
                show_sel = st.multiselect("Show", top_n, default=top_n[:5], key="n_show")
            df_nf = df_n[df_n["Type"].isin(type_sel) & df_n["Show"].isin(show_sel)]
            for _, r in df_nf.head(50).iterrows():
                st.markdown(f"""<div class="notif-card">
                    <b>{r['Type']}</b> &nbsp;|&nbsp; To: <b>{r['Recipient']}</b>
                    &nbsp;|&nbsp; {r['Date']} &nbsp; {r['Status']}<br>
                    <small style="color:#64748b;">{r['Message']}</small>
                </div>""", unsafe_allow_html=True)
            if len(df_nf) > 50:
                st.caption(f"Showing 50 of {len(df_nf)}.")
            st.download_button("Export email log (.csv)", df_nf.to_csv(index=False).encode(),
                               "ges_notifications.csv", "text/csv")

    # ── Push to Outlook ──────────────────────────────────────────────
    with tab_push:
        st.markdown('<div class="section-title">Push to Outlook Calendars</div>', unsafe_allow_html=True)
        st.caption("Auto-Sync does this every 15 minutes. Use this to push the shown calendar now.")
        dups   = duplicate_ids(list(sp_parsed.values()))
        future = df["DATE"] >= today()
        df_fut = df[future]
        mbx_s, err_s = route_rows(df_fut, dups)
        ok_s = err_s == ""
        if (~future).any():
            st.info(f"**{int((~future).sum())}** row(s) are dated before today and are not pushed. "
                    f"Past events already in Outlook stay as history.")
        if ok_s.any():
            dest = df_fut.loc[ok_s, "CALENDAR"].groupby(mbx_s[ok_s]).unique()
            st.info("Pushing to: " + "; ".join(f"**{', '.join(sorted(ks))}** → {m}" for m, ks in dest.items()))
        if not ok_s.all():
            st.error(f"**{int((~ok_s).sum())}** row(s) will NOT be pushed until fixed:")
            bad = df_fut.loc[~ok_s, ["DATE", "TIME", "SHOW", "FUNCTION", "FACILITY", "LOCATION", "FILE"]].assign(
                Problem=err_s[~ok_s])
            bad["DATE"] = bad["DATE"].dt.strftime("%Y-%m-%d")
            st.dataframe(bad, hide_index=True, use_container_width=True)
        fac_w = facility_warnings(df_fut)
        if (fac_w != "").any():
            st.warning(f"**{int((fac_w != '').sum())}** row(s) have a FACILITY note. They are still pushed to "
                       f"their folder's calendar, without a facility color:")
            w = df_fut.loc[fac_w != "", ["DATE", "SHOW", "FUNCTION", "FACILITY", "FILE"]].assign(Note=fac_w[fac_w != ""])
            w["DATE"] = w["DATE"].dt.strftime("%Y-%m-%d")
            st.dataframe(w, hide_index=True, use_container_width=True)

        col_test, _ = st.columns([1, 3])
        with col_test:
            if st.button("Test Connection"):
                with st.spinner("Checking each calendar..."):
                    try:
                        tok = get_graph_token()
                        for key, c in CALENDARS.items():
                            if not c["mailbox"]:
                                st.warning(f"{key}: no mailbox set")
                                continue
                            r = requests.get(f"https://graph.microsoft.com/v1.0/users/{c['mailbox']}/calendar",
                                             headers={"Authorization": f"Bearer {tok}"}, timeout=10)
                            (st.success if r.status_code == 200 else st.error)(
                                f"{key} → {c['mailbox']}: " + ("connected" if r.status_code == 200
                                                               else f"failed ({r.status_code})"))
                    except Exception as e:
                        st.error(f"Auth error: {e}")

        st.divider()
        st.caption(f"**{len(df_fut):,}** events from today on in **{view_label}**.")
        test_mode = st.checkbox("Test mode (push a few events only)", value=True)
        test_n    = st.number_input("Number of events to push", min_value=1, max_value=max(len(df_fut), 1),
                                    value=min(2, max(len(df_fut), 1)), step=1) if test_mode else len(df_fut)
        if test_mode:
            st.caption(f"Test mode: will push **{int(test_n)}** of {len(df_fut)} events.")

        if st.button("Push to Outlook", type="primary", use_container_width=True, disabled=df_fut.empty):
            try:
                token = get_graph_token()
            except Exception as e:
                st.error(f"Auth failed: {e}")
                st.stop()
            with st.spinner("Pushing to Outlook..."):
                for m in set(mbx_s[ok_s]):  # event colors; push continues either way
                    if (cat_err := ensure_outlook_categories(token, m)):
                        st.warning(f"Event colors: {cat_err}")
                stats, results, not_pushed = sync_rows(token, df_fut.head(int(test_n)), dups)
                results += [{"Shift": e, "Show": "—", "Status": f"Not pushed: {m}"} for e, m in not_pushed]
                # Remove shifts deleted from Excel (today on, every calendar), from a fresh read of SharePoint
                _, parsed_now, errors_now, _ = load_all(token)
                removed, problem = remove_deleted_shifts(token, parsed_now, errors_now)
                results += [{"Shift": e, "Show": "—", "Status": "Deleted (removed from Excel)" if ok else "Delete failed"}
                            for e, _, ok in removed]
            if problem:
                st.error(problem)
            failed = stats["failed"] + len(not_pushed) + sum(1 for *_, ok in removed if not ok)
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Created", stats["created"])
            c2.metric("Updated", stats["updated"])
            c3.metric("Deleted", sum(1 for *_, ok in removed if ok))
            c4.metric("Skipped", stats["skipped"])
            c5.metric("Failed",  failed)
            if results:
                st.dataframe(pd.DataFrame(results), use_container_width=True, hide_index=True)
            if failed == 0:
                st.success("Done. Check the Outlook calendars.")
            else:
                st.warning(f"{failed} event(s) failed or were not pushed. See table above.")


# ═══════════════════════════════════════════════════════════════════
# PAGE: CALENDAR VIEW
# ═══════════════════════════════════════════════════════════════════
elif page == "Outlook Calendar":
    st.markdown('<div class="section-title">Outlook Calendar (read-only)</div>', unsafe_allow_html=True)
    st.caption("What Outlook has right now, read straight from the calendar's mailbox. Nothing here changes Outlook. "
               "Hover over an event to see its notes, or pick a day below for full details.")

    _cals = [k for k, c in CALENDARS.items() if c["mailbox"]]
    if not _cals:
        st.info("No calendar has a MAILBOX on the admin Calendars sheet.")
        st.stop()
    if "oc_day" not in st.session_state:
        st.session_state.oc_day = today().date()

    def _oc_shift_month(step):
        d = pd.Timestamp(st.session_state.oc_day).replace(day=1) + pd.DateOffset(months=step)
        st.session_state.oc_day = d.date()

    c1, c2, c3, c4, c5 = st.columns([3, 1, 1, 1, 2])
    with c1:
        oc_cal = st.selectbox("Calendar", _cals, index=_cals.index(view) if view in _cals else 0, key="oc_cal")
    with c2:
        st.button("◀ Prev month", on_click=_oc_shift_month, args=(-1,), use_container_width=True)
    with c3:
        st.button("Today", on_click=lambda: st.session_state.update(oc_day=today().date()), use_container_width=True)
    with c4:
        st.button("Next month ▶", on_click=_oc_shift_month, args=(1,), use_container_width=True)
    with c5:
        if st.button("Reload from Outlook", use_container_width=True):
            _read_outlook.clear()

    cal_cfg = CALENDARS[oc_cal]
    sel_day = pd.Timestamp(st.session_state.oc_day)
    month1  = sel_day.replace(day=1)
    weeks   = pycal.Calendar(firstweekday=6).monthdatescalendar(month1.year, month1.month)  # weeks start Sunday
    events, oc_err = _read_outlook(cal_cfg["mailbox"], pd.Timestamp(weeks[0][0]), pd.Timestamp(weeks[-1][-1]),
                                   cal_cfg["time_zone"])
    if oc_err:
        st.error(oc_err)
    st.markdown(f"### {month1.strftime('%B %Y')}")
    st.caption(f"Mailbox **{cal_cfg['mailbox']}** · times in {cal_cfg['time_zone']} · "
               f"read at {pd.Timestamp.now(tz='America/Los_Angeles').strftime('%I:%M %p').lstrip('0')} PT")

    cat_colors = {name: OUTLOOK_FILL.get(preset, DEFAULT_FILL) for name, preset in outlook_categories(oc_cal).items()}

    def _when(e):
        if e["ALL_DAY"]:
            return "All day"
        return e["START"].strftime("%I:%M %p").lstrip("0")

    cells = []
    for week in weeks:
        row = []
        for d in week:
            d = pd.Timestamp(d)
            chips = []
            for e in events_on_day(events, d):
                fill = next((cat_colors[c] for c in e["CATEGORIES"] if c in cat_colors), DEFAULT_FILL)
                tip  = "\n".join(x for x in [f"{_when(e)}  {e['SUBJECT']}", e["LOCATION"], e["NOTES"]] if x)
                label = e["SUBJECT"] if e["ALL_DAY"] else f"{_when(e)} {e['SUBJECT']}"
                chips.append(f'<div class="oc-ev" style="background:{fill}" title="{escape(tip)}">{escape(label)}</div>')
            cls = "oc-day" + (" oc-out" if d.month != month1.month else "") \
                + (" oc-today" if d == today() else "") + (" oc-sel" if d == sel_day else "")
            row.append(f'<td class="{cls}"><div class="oc-num">{d.day}</div>{"".join(chips)}</td>')
        cells.append("<tr>" + "".join(row) + "</tr>")
    heads = "".join(f"<th>{n}</th>" for n in ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
    st.markdown(f"""
<style>
  .oc-tbl {{ width:100%; border-collapse:collapse; table-layout:fixed; }}
  .oc-tbl th {{ font-size:12px; color:#64748b; font-weight:600; text-align:left; padding:4px 6px; }}
  .oc-day {{ border:1px solid #e2e8f0; vertical-align:top; height:120px; padding:4px; }}
  .oc-day > div.oc-num {{ font-size:13px; font-weight:600; color:#1e293b; margin-bottom:2px; }}
  .oc-out {{ background:#f8fafc; }} .oc-out > div.oc-num {{ color:#94a3b8; }}
  .oc-today > div.oc-num {{ background:#2563eb; color:#fff; border-radius:10px; display:inline-block; padding:0 7px; }}
  .oc-sel {{ outline:2px solid #f59e0b; outline-offset:-2px; }}
  .oc-ev {{ font-size:11px; color:#1e293b; border-radius:3px; padding:1px 4px; margin-bottom:2px;
           white-space:nowrap; overflow:hidden; text-overflow:ellipsis; cursor:default; }}
</style>
<table class="oc-tbl"><tr>{heads}</tr>{"".join(cells)}</table>""", unsafe_allow_html=True)

    if cat_colors:
        st.markdown(" ".join(f'<span class="oc-ev" style="background:{c};display:inline-block;padding:2px 8px">'
                             f'{escape(n)}</span>' for n, c in cat_colors.items()), unsafe_allow_html=True)

    # ── One day's events, with the full notes ─────────────────────────
    st.divider()
    st.date_input("Show the events of", key="oc_day", format="MM/DD/YYYY")
    day_events = events_on_day(events, sel_day)
    if not day_events:
        st.caption(f"No events on {sel_day.strftime('%A %B %d, %Y')}.")
    for e in day_events:
        if e["ALL_DAY"]:
            last = e["END"] - timedelta(days=1)
            span = "All day" + (f", through {last.strftime('%a %b %d')}" if last > e["START"] else "")
        else:
            span = f"{e['START'].strftime('%a %b %d %I:%M %p')} – {e['END'].strftime('%a %b %d %I:%M %p')}"
        with st.expander(f"{_when(e)} · {e['SUBJECT']}", expanded=len(day_events) <= 3):
            st.markdown(f"**When:** {span}  \n**Location:** {e['LOCATION'] or '—'}  \n"
                        f"**Color category:** {', '.join(e['CATEGORIES']) or '—'}  \n"
                        f"**From the schedule:** {'Yes' if e['EVENT_ID'] else 'No (added in Outlook)'}  \n"
                        f"**Last changed in Outlook:** "
                        + (e["MODIFIED"].tz_convert("America/Los_Angeles").strftime("%b %d, %Y %I:%M %p") + " PT"
                           if e["MODIFIED"] is not None else "—"))
            st.text(e["NOTES"] or "(no notes)")

elif "Calendar" in page:
    st.markdown('<div class="section-title">Labor Calendar</div>', unsafe_allow_html=True)
    st.caption(f"Source: **{view_label}** · {stype}")

    # ── Filters ──────────────────────────────────────────────────────
    f1, f2, f3, f4 = st.columns([1, 2, 2, 3])
    with f1:
        years   = sorted(df["DATE"].dt.year.dropna().unique().tolist())
        yr_sel  = st.selectbox("Year", ["All"] + [str(y) for y in years])
    with f2:
        fn_opts = sorted(df["FUNCTION"].dropna().unique().tolist())
        fn_sel  = st.multiselect("Function", fn_opts, default=fn_opts)
    with f3:
        locs    = [l for l in sorted(df["LOCATION"].dropna().unique().tolist()) if l]
        loc_sel = st.selectbox("Location", ["All Locations"] + locs)
    with f4:
        shows    = sorted(df["SHOW"].dropna().unique().tolist())
        show_sel = st.selectbox("Show", ["All Shows"] + shows)

    df_filt = df.copy()
    if yr_sel != "All":
        df_filt = df_filt[df_filt["DATE"].dt.year == int(yr_sel)]
    if fn_sel:
        df_filt = df_filt[df_filt["FUNCTION"].isin(fn_sel)]
    if loc_sel != "All Locations":
        df_filt = df_filt[df_filt["LOCATION"] == loc_sel]
    if show_sel != "All Shows":
        df_filt = df_filt[df_filt["SHOW"] == show_sel]

    st.caption(f"**{len(df_filt)}** shifts matching filters")

    # ── Views ────────────────────────────────────────────────────────
    v_gantt, v_grid, v_table, v_monthly = st.tabs(
        ["Gantt", "Calendar Grid", "Table", "Monthly Summary"]
    )

    with v_gantt:
        if df_filt.empty:
            st.warning("No shifts match the current filters.")
        else:
            df_g = df_filt.copy()
            has_end = stype == "events" and df_g["END_DATE"].notna().any()
            if has_end:
                df_g["END_DATE"] = df_g["END_DATE"].fillna(df_g["DATE"])
                fig = px.timeline(
                    df_g.sort_values("DATE"), x_start="DATE", x_end="END_DATE",
                    y="SHOW", color="FUNCTION", color_discrete_map=FUNCTION_COLORS,
                    hover_data={"LOCATION": True, "FOREMAN": True, "TOTAL_CREW": True,
                                "DATE": False, "END_DATE": False},
                    title="Show Timeline",
                )
            else:
                df_g["_END"] = df_g["DATE"] + pd.Timedelta(hours=8)
                fig = px.timeline(
                    df_g.sort_values("DATE"), x_start="DATE", x_end="_END",
                    y="SHOW", color="FUNCTION", color_discrete_map=FUNCTION_COLORS,
                    hover_data={"LOCATION": True, "TIME": True, "FOREMAN": True,
                                "TOTAL_CREW": True, "DATE": False, "_END": False},
                    title="Labor Schedule",
                )
            fig.update_layout(
                height=max(420, len(df_g["SHOW"].unique()) * 22 + 100),
                xaxis_title="", yaxis_title="", legend_title="Function",
                plot_bgcolor="#f8fafc", paper_bgcolor="white",
                font=dict(size=12), margin=dict(l=10, r=10, t=40, b=20),
            )
            fig.update_yaxes(autorange="reversed")
            st.plotly_chart(fig, use_container_width=True)

    with v_grid:
        if df_filt.empty:
            st.warning("No shifts match the current filters.")
        else:
            df_g = df_filt.copy()
            df_g["DayOfMonth"] = df_g["DATE"].dt.day
            df_g["MonthYear"]  = df_g["DATE"].dt.strftime("%b %Y")
            month_order = (
                df_g["DATE"].dt.to_period("M").drop_duplicates().sort_values()
                    .apply(lambda p: p.strftime("%b %Y")).tolist()
            )
            pivot = (
                df_g.groupby(["MonthYear", "DayOfMonth"]).size()
                    .reset_index(name="Shifts")
                    .pivot(index="MonthYear", columns="DayOfMonth", values="Shifts")
                    .fillna(0)
                    .reindex([m for m in month_order if m in df_g["MonthYear"].values])
            )
            fig_grid = px.imshow(
                pivot,
                labels=dict(x="Day", y="Month", color="Shifts"),
                color_continuous_scale=[[0, "#f0f9ff"], [0.3, "#93c5fd"], [1, "#1d4ed8"]],
                title="Shift Density — darker = more shifts that day",
                aspect="auto",
            )
            fig_grid.update_layout(
                height=max(260, len(pivot) * 42 + 100),
                margin=dict(l=10, r=10, t=40, b=20),
            )
            fig_grid.update_xaxes(side="top", tickmode="linear", tick0=1, dtick=1)
            st.plotly_chart(fig_grid, use_container_width=True)

            st.markdown("**Breakdown by function:**")
            fn_monthly = (
                df_g.groupby([df_g["DATE"].dt.strftime("%b %Y"), "FUNCTION"])
                    .size().reset_index(name="Shifts")
            )
            fn_monthly.columns = ["Month", "Function", "Shifts"]
            fig_fn = px.bar(fn_monthly, x="Month", y="Shifts", color="Function",
                            color_discrete_map=FUNCTION_COLORS, barmode="group",
                            category_orders={"Month": month_order})
            fig_fn.update_layout(height=300, margin=dict(l=10, r=10, t=20, b=20))
            st.plotly_chart(fig_fn, use_container_width=True)

    with v_table:
        if df_filt.empty:
            st.warning("No shifts match the current filters.")
        else:
            d = df_filt[["DATE","SHOW","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW","NOTES"]].copy()
            d["DATE"] = d["DATE"].dt.strftime("%Y-%m-%d")
            st.dataframe(d.sort_values("DATE"), use_container_width=True, hide_index=True)

    with v_monthly:
        if df_filt.empty:
            st.warning("No shifts match the current filters.")
        else:
            df_m = df_filt.copy()
            df_m["Month"] = df_m["DATE"].dt.to_period("M").astype(str)
            monthly = df_m.groupby(["Month", "FUNCTION"]).size().reset_index(name="Shifts")
            fig2 = px.bar(monthly, x="Month", y="Shifts", color="FUNCTION",
                          color_discrete_map=FUNCTION_COLORS, barmode="stack",
                          title="Shifts per Month by Function")
            fig2.update_layout(height=380, margin=dict(t=40, b=20, l=10, r=10))
            st.plotly_chart(fig2, use_container_width=True)

    # ── Shift Detail ─────────────────────────────────────────────────
    if not df_filt.empty:
        st.divider()
        st.markdown('<div class="section-title">Shift Detail</div>', unsafe_allow_html=True)
        sel = st.selectbox("Select show", sorted(df_filt["SHOW"].unique().tolist()), key="cal_show")
        for _, r in df_filt[df_filt["SHOW"] == sel].sort_values("DATE").iterrows():
            fn  = r["FUNCTION"]
            cls = {"SET": "fn-set", "STRIKE": "fn-strike", "STANDBY": "fn-standby"}.get(fn, "fn-set")
            end_label = (f" → {r['END_DATE'].strftime('%b %d')}"
                         if pd.notna(r["END_DATE"]) else "")
            st.markdown(f"""
**{r['DATE'].strftime('%a %b %d, %Y')}**{end_label} &nbsp;
<span class="fn-badge {cls}">{fn}</span> &nbsp; {r['TIME']} &nbsp; {r['LOCATION']}

Foreman: {r['FOREMAN']} &nbsp; {r['NOTES']}

---""", unsafe_allow_html=True)

    # ── Exports ───────────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-title">Export Calendar</div>', unsafe_allow_html=True)
    st.caption(
        "Each .ics file imports as a separate calendar in Outlook. "
        "Importing by function type keeps SET, STRIKE, and STANDBY events visually distinct."
    )

    e1, e2, e3, e4, e5 = st.columns(5)
    with e1:
        st.download_button("All Shifts (.ics)",
            data=_build_ics(df_filt, "All"), file_name="ges_all.ics",
            mime="text/calendar", use_container_width=True, disabled=df_filt.empty)
    for fn_type, col_w in zip(["SET", "STRIKE", "STANDBY"], [e2, e3, e4]):
        df_fn = df_filt[df_filt["FUNCTION"] == fn_type]
        with col_w:
            st.download_button(f"{fn_type} (.ics)",
                data=_build_ics(df_fn, fn_type), file_name=f"ges_{fn_type.lower()}.ics",
                mime="text/calendar", use_container_width=True, disabled=df_fn.empty)
    with e5:
        csv_out = df_filt[["DATE","SHOW","FUNCTION","LOCATION","TIME","FOREMAN","TOTAL_CREW","NOTES"]].copy()
        csv_out["DATE"] = csv_out["DATE"].dt.strftime("%Y-%m-%d")
        st.download_button("Export (.csv)",
            data=csv_out.to_csv(index=False).encode(), file_name="ges_schedule.csv",
            mime="text/csv", use_container_width=True, disabled=df_filt.empty)


# ═══════════════════════════════════════════════════════════════════
# PAGE: CREW LOOKUP
# ═══════════════════════════════════════════════════════════════════
elif "Crew" in page:
    st.markdown('<div class="section-title">Crew Lookup</div>', unsafe_allow_html=True)

    has_crew = df["CREW_LIST"].apply(lambda x: len(x) > 0).any()
    if not has_crew:
        st.info(
            "Crew Lookup requires individual crew member columns in your Excel file "
            "(columns named 1, 2, 3 ... for each crew member slot). "
            "Your current file does not have those columns."
        )
        st.stop()

    all_crew = sorted({
        p.replace(" (Foreman)", "").strip()
        for _, row in df.iterrows()
        for p in row["CREW_LIST"]
        if p.replace(" (Foreman)", "").strip() not in ("TBD", "nan", "")
    })

    if not all_crew:
        st.warning("No named crew members found.")
        st.stop()

    selected = st.selectbox("Select electrician", all_crew)
    person_rows = []
    for _, row in df.iterrows():
        names = [p.replace(" (Foreman)", "").strip() for p in row["CREW_LIST"]]
        if selected in names:
            role = "Foreman" if any(selected in p and "(Foreman)" in p for p in row["CREW_LIST"]) else "Crew"
            person_rows.append({
                "Date": row["DATE"].strftime("%Y-%m-%d"),
                "Show": row["SHOW"], "Function": row["FUNCTION"],
                "Location": row["LOCATION"], "Time": row["TIME"],
                "Role": role, "Notes": row["NOTES"],
            })

    if person_rows:
        df_p = pd.DataFrame(person_rows).sort_values("Date")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Shifts", len(df_p))
        c2.metric("As Foreman",   int((df_p["Role"] == "Foreman").sum()))
        c3.metric("As Crew",      int((df_p["Role"] == "Crew").sum()))
        st.dataframe(df_p, use_container_width=True, hide_index=True)

        crew_df = df[df["CREW_LIST"].apply(lambda lst: any(selected in p for p in lst))]
        st.download_button(
            f"Download {selected}'s schedule (.ics)",
            data=_build_ics(crew_df, f"Crew {selected}"),
            file_name=f"ges_{selected.replace(' ', '_')}.ics",
            mime="text/calendar",
        )
        st.info("Double-click the .ics file to import directly into Outlook.")
    else:
        st.info(f"No shifts found for **{selected}**.")


# ═══════════════════════════════════════════════════════════════════
# PAGE: OVERTIME / DOUBLE-TIME
# ═══════════════════════════════════════════════════════════════════
elif "Overtime" in page:
    st.markdown('<div class="section-title">Overtime</div>', unsafe_allow_html=True)
    st.caption(f"Source: **{view_label}** · rules and holidays come from the admin workbook (OT Rules and Holidays sheets)")

    ot_all = compute_overtime(df)  # on the whole file, so daily totals see every shift before filtering
    if ot_all.empty:
        st.info("No crew names or +N crew in this file, so there is nothing to count.")
        st.stop()

    # ── Filters ──────────────────────────────────────────────────────
    f1, f2 = st.columns([2, 3])
    with f1:
        d_min, d_max = df["DATE"].min().date(), df["DATE"].max().date()
        d_rng = st.date_input("Dates", (d_min, d_max), min_value=d_min, max_value=d_max, key="ot_dates")
    with f2:
        fac_opts = sorted(ot_all["FACILITY"].unique().tolist())
        fac_sel  = st.multiselect("Facility", fac_opts, default=fac_opts, key="ot_fac",
                                  format_func=lambda c: f"{c} - {FACILITIES[c]['name']}" if c in FACILITIES
                                                        else (c or "(no facility)"))
    d_from, d_to = (d_rng[0], d_rng[-1]) if d_rng else (d_min, d_max)
    ot = ot_all[ot_all["DATE"].dt.date.between(d_from, d_to) & ot_all["FACILITY"].isin(fac_sel)]

    n_norule = int((ot["STATUS"] == "no OT rules").sum())
    n_notime = int((ot["STATUS"] == "no TIME").sum())
    if n_norule:
        st.warning(f"{n_norule} crew shift(s) are on calendars with no OT rules yet (admin workbook); "
                   f"their hours are counted as regular.")
    if n_notime:
        st.caption(f"{n_notime} crew shift(s) have no TIME, so their hours are unknown (counted as 0).")

    t_team, t_ind, t_dt_team, t_dt_ind = st.tabs(
        ["OT - Team", "OT - Individual", "Double-Time - Team", "Double-Time - Individual"])

    with t_team:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Crew Shifts", len(ot))
        c2.metric("Total Hours", f"{ot['HOURS'].sum():,.2f}")
        c3.metric("OT Shifts",   int(ot["OT_SHIFT"].sum()))
        c4.metric("OT Hours",    f"{ot['OT_HOURS'].sum():,.2f}")
        st.markdown("**By electrician**")
        st.dataframe(ot_summary(ot, "PERSON").round(2), hide_index=True, use_container_width=True)
        st.caption("Unnamed crew = all +N slots combined; each is counted per shift (they can't be tracked "
                   "across shifts).")
        st.markdown("**By facility**")
        st.dataframe(ot_summary(ot, "FACILITY").round(2), hide_index=True, use_container_width=True)
        st.download_button("Export OT detail (.csv)", ot.to_csv(index=False).encode(),
                           "ges_overtime.csv", "text/csv")

    with t_ind:
        names = sorted(ot.loc[ot["NAMED"], "PERSON"].unique().tolist())
        if not names:
            st.info("No named electricians match the filters.")
        else:
            who = st.selectbox("Select electrician", names, key="ot_person")
            mine = ot[ot["PERSON"] == who]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Shifts",      len(mine))
            c2.metric("Total Hours", f"{mine['HOURS'].sum():,.2f}")
            c3.metric("OT Shifts",   int(mine["OT_SHIFT"].sum()))
            c4.metric("OT Hours",    f"{mine['OT_HOURS'].sum():,.2f}")
            d = mine[["DATE", "SHOW", "FACILITY", "FUNCTION", "TIME", "HOURS", "REG_HOURS", "OT_HOURS",
                      "OT_REASON", "STATUS"]].copy()
            d.insert(1, "DAY", d["DATE"].dt.strftime("%a"))
            d["DATE"] = d["DATE"].dt.strftime("%Y-%m-%d")
            st.dataframe(d.round(2), hide_index=True, use_container_width=True)

    for tab in (t_dt_team, t_dt_ind):
        with tab:
            if not DT_RULES:
                st.info("Double-Time rules have not been provided yet, so double-time is not calculated. "
                        "This view will show double-time shifts and hours once the rules are added "
                        "(DT_RULES in schedule_config.py).")

    with st.expander("OT rules and holidays in use"):
        for region, rules in OT_RULES.items():
            cals = ", ".join(k for k, c in CALENDARS.items() if c["ot_region"] == region)
            st.markdown(f"**{region}** ({cals}): over {rules['daily_hours']} hrs on a normal workday"
                        + (", Saturday" if rules.get("saturday") else "")
                        + (", Sunday" if rules.get("sunday") else "")
                        + (", holidays" if rules.get("holidays") else "") + " = OT")
            if rules.get("holidays"):
                st.caption("Holidays: " + ", ".join(pd.to_datetime(HOLIDAYS.get(region, []))
                                                    .strftime("%a %b %d, %Y")))
