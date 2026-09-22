"""Workbook loading and cleaning. Every sheet -> a tidy DataFrame."""
from __future__ import annotations

import json
import re
from datetime import datetime, time as dtime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import config

IST = ZoneInfo(config.TIMEZONE)
ET = ZoneInfo("America/New_York")
US_FMT = "%m/%d/%Y %I:%M:%S %p"

SHEET = {
    "alerts": "Dashboard_Alerts_PA", "status": "Dashboard_status", "shifts": "DB Timings",
    "priority": "Priority BI Jobs", "timing": "Dashboards timing", "jobs": "Jobs",
    "summary": "Summary", "disabled": "Disabled", "waiting": "Waiting",
    "running": "Running", "failed": "Failed",
}


def now_ist() -> datetime:
    return datetime.now(IST).replace(tzinfo=None)


def now_et() -> datetime:
    return datetime.now(ET).replace(tzinfo=None)


def next_run(now: datetime | None = None) -> datetime:
    now = now or datetime.now(IST)
    for day in (0, 1):
        for h in config.RUN_HOURS:
            cand = (now + timedelta(days=day)).replace(hour=h, minute=0, second=0, microsecond=0)
            if cand > now:
                return cand
    raise RuntimeError("unreachable")


# ------------------------------------------------------------------ helpers
def _txt(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, str):
        v = v.replace("\xa0", " ").strip()
        return v or None
    return v


def _dt(s, fmt: str | None = None) -> pd.Series:
    s = pd.Series(s)
    if fmt:
        out = pd.to_datetime(s, format=fmt, errors="coerce")
        miss = out.isna() & s.notna()
        if miss.any():
            out[miss] = pd.to_datetime(s[miss], errors="coerce", format="mixed")
        return out
    return pd.to_datetime(s, errors="coerce", format="mixed")


def _find_header(raw: pd.DataFrame, needle: str, limit: int = 15):
    for i in range(min(limit, len(raw))):
        for v in raw.iloc[i].tolist():
            if isinstance(v, str) and needle in v.lower():
                return i
    return None


def _table(raw: pd.DataFrame, header_row: int = 0) -> pd.DataFrame:
    header = [str(_txt(h) or "") for h in raw.iloc[header_row].tolist()]
    body = raw.iloc[header_row + 1:].reset_index(drop=True)
    body.columns = range(body.shape[1])
    keep = [i for i, h in enumerate(header) if h and not h.lower().startswith("count:")]
    body = body[keep]
    body.columns = [header[i] for i in keep]
    body = body.map(_txt)
    return body.dropna(how="all").reset_index(drop=True)


def _canon(df: pd.DataFrame, mapping: list[tuple[str, str]]) -> pd.DataFrame:
    cols = []
    for c in df.columns:
        name = c
        for pat, target in mapping:
            if re.search(pat, c.strip(), re.I):
                name = target
                break
        cols.append(name)
    df.columns = cols
    return df


def _ensure(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    for c in cols:
        if c not in df.columns:
            df[c] = None
    return df


def _fmt_time_cell(v):
    if isinstance(v, (dtime, datetime)):
        return v.strftime("%I:%M %p").lstrip("0")
    return _txt(v)


# ------------------------------------------------------------------ parsers
ALERT_COLS = ["Alert ID", "Dashboard Name", "Last Successful Refresh Time", "Expected Refresh Time",
              "Business Date", "Status", "Alert Severity", "Failure Reason", "Owner", "Alert Message",
              "Alert Timestamp", "Alert Key", "Notified (Y/N)", "Alert_Sent_Time"]


def parse_alerts(raw):
    if raw is None:
        return pd.DataFrame(columns=ALERT_COLS + ["Severity"])
    h = _find_header(raw, "alert id")
    if h is None:
        return pd.DataFrame(columns=ALERT_COLS + ["Severity"])
    df = _ensure(_table(raw, h), ALERT_COLS)
    df = df[df["Alert ID"].notna()].copy()
    for c in ["Last Successful Refresh Time", "Alert Timestamp", "Alert_Sent_Time", "Business Date"]:
        df[c] = _dt(df[c])
    df["Severity"] = df["Alert Severity"].fillna("Unknown").astype(str).str.title()
    df["Notified (Y/N)"] = df["Notified (Y/N)"].fillna("N").astype(str).str.upper().str[:1]
    return df.reset_index(drop=True)


STATUS_MAP = [
    (r"^dashboard name$", "Dashboard"), (r"qmc timings", "QMC Timings (ET)"),
    (r"^schedule$", "Schedule (IST)"), (r"^dashboard refresh", "Refresh Time (IST)"),
    (r"qvw refresh", "QVW Refresh (IST)"), (r"qvw files?$", "QVW File"),
    (r"bi\s*/\s*dna", "Team"), (r"job status", "QVW Job Status"),
    (r"per[io]+dicity", "Periodicity"), (r"^dashboard status$", "Status"),
    (r"validation", "Validation Reason"), (r"business date", "Business Date"),
]
STATUS_COLS = ["Dashboard", "QMC Timings (ET)", "Schedule (IST)", "Refresh Time (IST)", "QVW File",
               "QVW Refresh (IST)", "Team", "QVW Job Status", "Periodicity", "Status",
               "Validation Reason", "Business Date"]


def parse_status(raw):
    empty_main = pd.DataFrame(columns=STATUS_COLS + ["QVW Jobs"])
    empty_qvw = pd.DataFrame(columns=["Dashboard", "QVW File", "QVW Refresh (IST)", "QVW Job Status"])
    if raw is None or raw.empty:
        return empty_main, empty_qvw
    df = _ensure(_canon(_table(raw, 0), STATUS_MAP), STATUS_COLS)
    df["_parent"] = df["Dashboard"].ffill()          # grouped rows: children have blank name
    qvw = df[df["_parent"].notna() & df["QVW File"].notna()][
        ["_parent", "QVW File", "QVW Refresh (IST)", "QVW Job Status"]].rename(columns={"_parent": "Dashboard"})
    qvw["QVW Refresh (IST)"] = _dt(qvw["QVW Refresh (IST)"])
    main = df[df["Dashboard"].notna()].drop(columns="_parent").copy()
    main["QVW Jobs"] = main["Dashboard"].map(qvw.groupby("Dashboard").size()).fillna(0).astype(int)
    main["Refresh Time (IST)"] = _dt(main["Refresh Time (IST)"])
    main["Business Date"] = _dt(main["Business Date"])
    main["Team"] = main["Team"].fillna("Unassigned")
    main["Status"] = main["Status"].fillna("Unknown")
    main["Periodicity"] = main["Periodicity"].fillna("—")
    main["QMC Timings (ET)"] = main["QMC Timings (ET)"].map(_fmt_time_cell)
    return main.reset_index(drop=True), qvw.reset_index(drop=True)


def _trigger_minutes(v):
    if v is None:
        return None
    if isinstance(v, (dtime, datetime)):
        return v.hour * 60 + v.minute
    m = re.search(r"(\d{1,2}):(\d{2})(?::\d{2})?\s*([AP]M)?", str(v), re.I)
    if not m:
        return None
    h, mi, ap = int(m[1]), int(m[2]), (m[3] or "").upper()
    if ap == "PM" and h < 12:
        h += 12
    if ap == "AM" and h == 12:
        h = 0
    return h * 60 + mi


def parse_shifts(raw):
    cols = ["Shift", "Dashboard", "Trigger", "Trigger Min", "Present Time", "Status", "Outcome"]
    if raw is None:
        return pd.DataFrame(columns=cols)
    h = _find_header(raw, "trigger time")
    if h is None:
        return pd.DataFrame(columns=cols)
    header = [str(_txt(x) or "") for x in raw.iloc[h].tolist()]
    rows = []
    for j, label in enumerate(header):
        if "shift" not in label.lower():
            continue
        shift = re.sub(r"\s*shift.*$", "", label, flags=re.I).strip().title() or label
        for r in raw.iloc[h + 1:, j:j + 4].itertuples(index=False):
            name = _txt(r[0])
            if not name:
                continue
            tm = _trigger_minutes(_txt(r[1]))
            rows.append({"Shift": shift, "Dashboard": name, "Trigger Min": tm,
                         "Trigger": None if tm is None else f"{tm // 60:02d}:{tm % 60:02d}",
                         "Present Time": _txt(r[2]), "Status": _txt(r[3]) or "N/A"})
    df = pd.DataFrame(rows, columns=[c for c in cols if c != "Outcome"])
    if df.empty:
        return pd.DataFrame(columns=cols)
    df["Present Time"] = _dt(df["Present Time"])
    df["Outcome"] = df["Status"].map(lambda s: "Updated" if s.lower().startswith("updated")
                                     else ("Not updated" if "not" in s.lower() else "N/A"))
    return df.sort_values(["Shift", "Trigger Min"], kind="stable").reset_index(drop=True)


def parse_priority(raw):
    cols = ["Name", "Executed On", "Status", "Distribution Group", "Last Execution",
            "Started/Scheduled", "Periodicity"]
    if raw is None:
        return pd.DataFrame(columns=cols)
    h = _find_header(raw, "executed on")
    if h is None:
        return pd.DataFrame(columns=cols)
    df = _ensure(_canon(_table(raw, h), [(r"per[io]+dicity", "Periodicity")]), cols)
    df = df[df["Name"].notna()].copy()
    df["Last Execution"] = _dt(df["Last Execution"], US_FMT)
    return df.reset_index(drop=True)


def parse_timing(raw):
    cols = ["Dashboard", "Category", "Last Updated (IST)"]
    if raw is None:
        return pd.DataFrame(columns=cols)
    df = _canon(_table(raw, 0), [(r"^dashboard name$", "Dashboard"), (r"^category$", "Category"),
                                 (r"last updated", "Last Updated (IST)")])
    df = _ensure(df, cols)[cols]
    df = df[df["Dashboard"].notna()].copy()
    df["Category"] = df["Category"].fillna("Uncategorised")
    df["Last Updated (IST)"] = _dt(df["Last Updated (IST)"])
    return df.reset_index(drop=True)


JOB_COLS = ["Name", "Executed On", "Status", "Distribution Group", "Last Execution", "Started/Scheduled"]


def _split_name(name):
    parts = [p for p in str(name).split("\\") if p]
    if parts and parts[0].upper().startswith("QDS@"):
        parts = parts[1:]
    project = parts[0] if len(parts) > 1 else "—"
    task = re.sub(r"\s*\(work disabled\)\s*$", "", parts[-1] if parts else str(name))
    return project, task


def _sched_type(v):
    if v is None:
        return "Unknown"
    t = str(v).lower()
    if t == "disabled":
        return "Disabled"
    if t.startswith("not scheduled"):
        return "Not scheduled"
    if t.startswith("multiple"):
        return "Multiple triggers"
    if re.match(r"\d{1,2}/\d{1,2}/\d{4}", t):
        return "Time-scheduled"
    if t.startswith("when") or "succeed" in t or "upon" in t or "events" in t:
        return "Event-triggered"
    return "Other"


def parse_jobs(raw):
    cols = JOB_COLS + ["Project", "Task", "Schedule Type", "Last Run (ET)", "Next/Started (ET)"]
    if raw is None or raw.empty:
        return pd.DataFrame(columns=cols)
    df = _ensure(_table(raw, 0), JOB_COLS)
    df = df[df["Name"].notna()].copy()
    pt = df["Name"].map(_split_name)
    df["Project"] = pt.map(lambda x: x[0])
    df["Task"] = pt.map(lambda x: x[1])
    df["Schedule Type"] = df["Started/Scheduled"].map(_sched_type)
    df["Last Run (ET)"] = _dt(df["Last Execution"], US_FMT)
    df["Next/Started (ET)"] = _dt(df["Started/Scheduled"].where(df["Schedule Type"] == "Time-scheduled"), US_FMT)
    df["Distribution Group"] = df["Distribution Group"].fillna("—")
    return df.reset_index(drop=True)


def parse_summary(raw):
    out = {"generated": None, "job_status": pd.DataFrame(columns=["Status", "Count"]),
           "kv": {}, "failed": pd.DataFrame(), "stale": pd.DataFrame(), "notes": []}
    if raw is None:
        return out
    rows = [[_txt(v) for v in r] for r in raw.values.tolist()]
    section, i = None, 0
    while i < len(rows):
        r = rows[i]
        a = r[0]
        filled = [v for v in r if v is not None]
        if not filled:
            i += 1
            continue
        if a and a.endswith(":") and len(r) > 1 and r[1] is not None:
            key = a.rstrip(":")
            out["kv"][key] = r[1]
            if key == "Generated":
                out["generated"] = str(r[1])
        elif a == "Status" and len(r) > 1 and r[1] == "Count":
            recs, i = [], i + 1
            while i < len(rows) and rows[i][0] and rows[i][1] is not None:
                recs.append({"Status": rows[i][0], "Count": int(rows[i][1])})
                i += 1
            out["job_status"] = pd.DataFrame(recs, columns=["Status", "Count"])
            continue
        elif a == "Dashboard Name":
            header = [v for v in r if v is not None]
            recs, i = [], i + 1
            while i < len(rows) and len([v for v in rows[i] if v is not None]) > 1:
                recs.append(rows[i][:len(header)])
                i += 1
            tbl = pd.DataFrame(recs, columns=header)
            key = "failed" if section and "Failed" in section else "stale"
            out[key] = tbl
            continue
        elif len(filled) == 1 and a:
            if a.lower().startswith("note:"):
                out["notes"].append(a)
            elif len(a) < 70 and not a.startswith("None"):
                section = a
        i += 1
    return out


# ------------------------------------------------------------------ public API
def resolve_file():
    if config.DATA_FILE.exists():
        return config.DATA_FILE, False
    if config.SAMPLE_FILE.exists():
        return config.SAMPLE_FILE, True
    return None, False


def signature():
    """(workbook mtime, run-status mtime) - changes whenever a new run lands."""
    path, _ = resolve_file()
    m1 = path.stat().st_mtime if path else 0
    m2 = config.STATUS_FILE.stat().st_mtime if config.STATUS_FILE.exists() else 0
    return (m1, m2)


def read_run_status():
    try:
        return json.loads(config.STATUS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


@st.cache_data(show_spinner="Loading workbook…")
def _load(path: str, mtime: float) -> dict:
    raw = pd.read_excel(path, sheet_name=None, header=None, dtype=object, engine="openpyxl")
    g = lambda k: raw.get(SHEET[k])
    status, qvw = parse_status(g("status"))
    d = {
        "alerts": parse_alerts(g("alerts")), "status": status, "qvw": qvw,
        "shifts": parse_shifts(g("shifts")), "priority": parse_priority(g("priority")),
        "timing": parse_timing(g("timing")), "summary": parse_summary(g("summary")),
        "missing": [n for k, n in SHEET.items() if n not in raw],
    }
    for k in ("jobs", "disabled", "waiting", "running", "failed"):
        d[k] = parse_jobs(g(k))
    return d


def get_data():
    """Returns (data_dict, info) or (None, None) when no workbook exists."""
    path, sample = resolve_file()
    if path is None:
        return None, None
    mtime = path.stat().st_mtime
    d = _load(str(path), mtime)
    gen = d["summary"]["generated"]
    as_of = pd.to_datetime(gen, errors="coerce") if gen else pd.NaT
    if pd.isna(as_of):
        as_of = datetime.fromtimestamp(mtime, IST).replace(tzinfo=None)
    return d, {"path": path, "sample": sample, "mtime": mtime, "as_of": as_of.to_pydatetime()}
