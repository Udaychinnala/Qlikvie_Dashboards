"""Qlik Dashboard Monitor — reads the workbook produced by
dashboard_status_pipeline0909_updated.py and presents every sheet as a
dedicated, color-coded tab.

Run:  streamlit run app.py
"""
from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

import config
import data
from style import (badge, badge_html_col, color_for, empty_state, fmt_dt,
                    inject_css, render_pills)

st.set_page_config(page_title="Qlik Dashboard Monitor", page_icon="📊", layout="wide")
inject_css()
px.defaults.template = "plotly_white"

# ============================================================ auto-refresh
@st.fragment(run_every=60)
def _watch_for_new_data():
    """Runs every 60s in the background; reruns the whole app the moment
    the pipeline writes a fresh workbook, so viewers never have to hit F5."""
    if st.session_state.get("sig") not in (None, data.signature()):
        st.cache_data.clear()
        st.rerun()


if st.session_state.get("autorefresh", True):
    _watch_for_new_data()

# ============================================================ load data
d, info = data.get_data()
st.session_state["sig"] = data.signature()

st.markdown('<p class="app-title">📊 Qlik Dashboard Monitor</p>', unsafe_allow_html=True)
st.markdown('<p class="app-sub">Live status across dashboards, alerts, shift timings and QMC jobs</p>',
            unsafe_allow_html=True)

if d is None:
    empty_state(f"No workbook found yet. Expected at `{config.DATA_FILE}` — "
                f"it will appear here automatically once the pipeline finishes its first run.")
    st.stop()

# ---- sidebar --------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Status")
    if info["sample"]:
        st.warning("Showing **sample data** — your pipeline hasn't written "
                    f"`{config.DATA_FILE.name}` yet.", icon="⚠️")

    age_hours = (data.now_ist() - info["as_of"]).total_seconds() / 3600
    st.metric("Data as of", info["as_of"].strftime("%d %b, %I:%M %p"))
    if age_hours > config.STALE_AFTER_HOURS:
        st.error(f"⏱️ {age_hours:.1f}h old — newer than expected for a 6-hourly run. "
                  "Check the scheduler.", icon="🚨")
    else:
        st.caption(f"🟢 Fresh — {age_hours:.1f}h ago")
    st.caption(f"⏭️ Next scheduled run: **{data.next_run().strftime('%d %b, %I:%M %p')}**")

    run_status = data.read_run_status()
    if run_status:
        icon = "✅" if run_status.get("ok") else "❌"
        st.caption(f"{icon} Last pipeline run: {run_status.get('finished_at', '—')} "
                    f"({run_status.get('duration_sec', '?')}s)")
        if not run_status.get("ok"):
            with st.expander("Last run error"):
                st.code(run_status.get("error", "unknown error"))

    st.toggle("Auto-refresh (every 60s)", value=True, key="autorefresh")
    if st.button("🔄 Reload now", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()
    st.header("🗺️ Navigate")
    st.caption("Overview · Dashboard Status · Alerts · Shift Timings · "
               "Job Monitor · Priority Jobs · Dashboards Timing")
    if d["missing"]:
        st.caption(f"Sheets not found in workbook: {', '.join(d['missing'])}")

# ============================================================ tabs
tab_overview, tab_status, tab_alerts, tab_shift, tab_jobs, tab_priority, tab_timing = st.tabs(
    ["🏠 Overview", "🖥️ Dashboard Status", "🚨 Alerts", "🕒 Shift Timings",
     "⚙️ Job Monitor", "⭐ Priority BI Jobs", "📅 Dashboards Timing"]
)

# ------------------------------------------------------------------ Overview
with tab_overview:
    s = d["summary"]
    kv = s["kv"]
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Dashboards", kv.get("Total Dashboards", len(d["status"])))
    c2.metric("Up to date", kv.get("Up To Date (Business Date = Today)", "—"))
    c3.metric("Stale (2+ days)", kv.get("Genuinely Stale (2+ days old)", "—"))
    c4.metric("Total Jobs", kv.get("Total Jobs", len(d["jobs"])))
    c5.metric("Open Alerts", len(d["alerts"]))

    left, right = st.columns([1, 1])
    with left:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("QMC Job Status")
        js = s["job_status"]
        if js.empty:
            empty_state("No job-status summary in this workbook.")
        else:
            fig = go.Figure(go.Pie(
                labels=js["Status"], values=js["Count"], hole=0.55,
                marker_colors=[color_for(x) for x in js["Status"]], sort=False,
            ))
            fig.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=10),
                               legend=dict(orientation="h", y=-0.1))
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Dashboard Status by Team")
        st_df = d["status"]
        if st_df.empty:
            empty_state("No dashboard-status data.")
        else:
            ct = st_df.groupby(["Team", "Status"]).size().reset_index(name="Count")
            fig = px.bar(ct, x="Team", y="Count", color="Status", barmode="stack",
                         color_discrete_map=color_for and {s_: color_for(s_) for s_ in ct["Status"].unique()})
            fig.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=10),
                               legend=dict(orientation="h", y=-0.25))
            st.plotly_chart(fig, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("⚠️ Needs attention")
    a1, a2 = st.columns(2)
    with a1:
        st.caption("Failed dashboards")
        if s["failed"].empty:
            st.success("No failed dashboards 🎉")
        else:
            st.dataframe(s["failed"], use_container_width=True, hide_index=True)
    with a2:
        st.caption(f"Stale dashboards ({len(s['stale'])})")
        if s["stale"].empty:
            st.success("Nothing stale 🎉")
        else:
            st.dataframe(s["stale"], use_container_width=True, hide_index=True, height=240)
    for note in s["notes"]:
        st.info(note, icon="📝")
    st.markdown("</div>", unsafe_allow_html=True)

# ------------------------------------------------------------------ Dashboard Status
with tab_status:
    df = d["status"]
    if df.empty:
        empty_state("`Dashboard_status` sheet is empty or missing.")
    else:
        f1, f2, f3, f4 = st.columns(4)
        teams = f1.multiselect("Team", sorted(df["Team"].unique()))
        statuses = f2.multiselect("Status", sorted(df["Status"].unique()))
        periods = f3.multiselect("Periodicity", sorted(df["Periodicity"].unique()))
        search = f4.text_input("Search dashboard name")

        view = df.copy()
        if teams:
            view = view[view["Team"].isin(teams)]
        if statuses:
            view = view[view["Status"].isin(statuses)]
        if periods:
            view = view[view["Periodicity"].isin(periods)]
        if search:
            view = view[view["Dashboard"].str.contains(search, case=False, na=False)]

        render_pills(df["Status"].value_counts().to_dict())
        st.caption(f"Showing {len(view)} of {len(df)} dashboards")

        show = view[["Dashboard", "Team", "Status", "Schedule (IST)", "Refresh Time (IST)",
                     "QVW Jobs", "Periodicity", "Business Date", "Validation Reason"]].copy()
        show["Refresh Time (IST)"] = show["Refresh Time (IST)"].map(fmt_dt)
        show["Business Date"] = show["Business Date"].map(lambda v: fmt_dt(v, "%d %b %Y"))
        show["Status"] = badge_html_col(show, "Status")
        st.write(
            show.to_html(escape=False, index=False, classes="styled-table"),
            unsafe_allow_html=True,
        )

        with st.expander("🔍 Inspect QVW jobs behind a dashboard"):
            pick = st.selectbox("Dashboard", sorted(df["Dashboard"].unique()))
            qvw = d["qvw"][d["qvw"]["Dashboard"] == pick].copy()
            if qvw.empty:
                st.caption("No individual QVW jobs listed for this dashboard.")
            else:
                qvw["QVW Refresh (IST)"] = qvw["QVW Refresh (IST)"].map(fmt_dt)
                st.dataframe(qvw.drop(columns="Dashboard"), use_container_width=True, hide_index=True)

# ------------------------------------------------------------------ Alerts
with tab_alerts:
    al = d["alerts"]
    if al.empty:
        st.success("No active alerts 🎉", icon="✅")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Alerts", len(al))
        c2.metric("High Severity", int((al["Severity"] == "High").sum()))
        c3.metric("Medium Severity", int((al["Severity"] == "Medium").sum()))
        c4.metric("Not Notified", int((al["Notified (Y/N)"] != "Y").sum()))

        fsev = st.multiselect("Severity", sorted(al["Severity"].unique()), key="al_sev")
        view = al[al["Severity"].isin(fsev)] if fsev else al

        left, right = st.columns([2, 1])
        with left:
            show = view[["Dashboard Name", "Severity", "Status", "Failure Reason",
                         "Owner", "Alert Timestamp", "Notified (Y/N)"]].copy()
            show["Alert Timestamp"] = show["Alert Timestamp"].map(fmt_dt)
            show["Severity"] = badge_html_col(show, "Severity")
            st.write(show.to_html(escape=False, index=False, classes="styled-table"),
                     unsafe_allow_html=True)
        with right:
            fig = px.pie(view, names="Severity", hole=0.5,
                         color="Severity", color_discrete_map={
                             k: color_for(k) for k in view["Severity"].unique()})
            fig.update_layout(height=320, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------------ Shift Timings
with tab_shift:
    sh = d["shifts"]
    if sh.empty:
        empty_state("`DB Timings` sheet is empty or missing.")
    else:
        shifts = [s for s in ["Morning", "Afternoon", "Night"] if s in sh["Shift"].unique()] \
            or sorted(sh["Shift"].unique())
        shift_tabs = st.tabs([f"🕘 {s}" for s in shifts])
        for name, stab in zip(shifts, shift_tabs):
            with stab:
                sub = sh[sh["Shift"] == name]
                counts = sub["Outcome"].value_counts().to_dict()
                render_pills(counts, ["Updated", "Not updated", "N/A"])
                cL, cR = st.columns([2, 1])
                with cL:
                    show = sub[["Dashboard", "Trigger", "Present Time", "Outcome"]].copy()
                    show["Present Time"] = show["Present Time"].map(fmt_dt)
                    show["Outcome"] = badge_html_col(show, "Outcome")
                    st.write(show.to_html(escape=False, index=False, classes="styled-table"),
                             unsafe_allow_html=True)
                with cR:
                    fig = px.pie(sub, names="Outcome", hole=0.5,
                                 color="Outcome", color_discrete_map={
                                     k: color_for(k) for k in sub["Outcome"].unique()})
                    fig.update_layout(height=300, margin=dict(t=10, b=10, l=10, r=10))
                    st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------------ Job Monitor
with tab_jobs:
    sub_names = ["All Jobs", "Disabled", "Waiting", "Running", "Failed"]
    keys = ["jobs", "disabled", "waiting", "running", "failed"]
    subtabs = st.tabs(sub_names)
    for name, key, stab in zip(sub_names, keys, subtabs):
        with stab:
            jdf = d[key]
            if jdf.empty:
                empty_state(f"No rows in the `{data.SHEET[key]}` sheet.")
                continue
            c1, c2, c3 = st.columns(3)
            c1.metric("Jobs", len(jdf))
            c2.metric("Projects", jdf["Project"].nunique())
            c3.metric("Distribution groups", (jdf["Distribution Group"] != "—").sum())

            if key == "jobs":
                fig = px.bar(jdf["Schedule Type"].value_counts().reset_index(),
                             x="Schedule Type", y="count", color="Schedule Type",
                             color_discrete_sequence=px.colors.qualitative.Set2)
                fig.update_layout(height=280, showlegend=False, margin=dict(t=10, b=10, l=10, r=10))
                st.plotly_chart(fig, use_container_width=True)

            proj = st.multiselect("Project", sorted(jdf["Project"].unique()), key=f"proj_{key}")
            view = jdf[jdf["Project"].isin(proj)] if proj else jdf
            show = view[["Project", "Task", "Status", "Schedule Type", "Distribution Group",
                         "Last Execution"]].head(1000).copy()
            st.dataframe(show, use_container_width=True, hide_index=True, height=420)
            if len(view) > 1000:
                st.caption(f"Showing first 1,000 of {len(view)} rows — narrow with the Project filter.")

# ------------------------------------------------------------------ Priority BI Jobs
with tab_priority:
    pr = d["priority"]
    if pr.empty:
        empty_state("`Priority BI Jobs` sheet is empty or missing.")
    else:
        render_pills(pr["Status"].value_counts().to_dict())
        c1, c2 = st.columns([2, 1])
        with c1:
            show = pr.copy()
            show["Last Execution"] = show["Last Execution"].map(fmt_dt)
            show["Status"] = badge_html_col(show, "Status")
            st.write(show.to_html(escape=False, index=False, classes="styled-table"),
                     unsafe_allow_html=True)
        with c2:
            fig = px.pie(pr, names="Status", hole=0.5,
                         color="Status", color_discrete_map={k: color_for(k) for k in pr["Status"].unique()})
            fig.update_layout(height=320, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig, use_container_width=True)

# ------------------------------------------------------------------ Dashboards Timing
with tab_timing:
    tm = d["timing"]
    if tm.empty:
        empty_state("`Dashboards timing` sheet is empty or missing.")
    else:
        cat = st.multiselect("Category", sorted(tm["Category"].unique()))
        view = tm[tm["Category"].isin(cat)] if cat else tm

        fig = px.bar(tm["Category"].value_counts().reset_index(),
                     x="Category", y="count", color="Category",
                     color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_layout(height=280, showlegend=False, margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig, use_container_width=True)

        show = view.copy()
        show["Last Updated (IST)"] = show["Last Updated (IST)"].map(fmt_dt)
        st.dataframe(show, use_container_width=True, hide_index=True, height=420)

# ------------------------------------------------------------------ table styling (once, at bottom)
st.markdown("""
<style>
.styled-table { width: 100%; border-collapse: collapse; font-size: 0.88rem; }
.styled-table th { background:#F4F6FB; text-align:left; padding:8px 10px; border-bottom:2px solid #E7E9F3;}
.styled-table td { padding:7px 10px; border-bottom:1px solid #F0F1F6; }
.styled-table tr:hover td { background:#FAFBFF; }
</style>
""", unsafe_allow_html=True)
