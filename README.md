# Qlik Dashboard Monitor

## Setup
```
pip install -r requirements.txt
```
Copy `dashboard_status_pipeline0909_updated.py` into this folder (already referenced by `config.py`).

## Configure (config.py)
- `DATA_FILE` — the workbook path the pipeline writes to and the dashboard reads (default `Expected_sheet.xlsx`).
- `PIPELINE_MODE` — `"live"` (scrapes QMC) or `"offline"`.
- `TIMEZONE` — default `Asia/Kolkata`.
All of these can also be set as environment variables (`DASHBOARD_XLSX`, `PIPELINE_MODE`, `DASHBOARD_TZ`, ...) instead of editing the file — handy for services/containers.

## Run the dashboard
```
streamlit run app.py
```
Until the pipeline has produced `DATA_FILE` for the first time, the dashboard shows the bundled sample workbook with a banner.

## Run the scheduler (12 AM / 6 AM / 12 PM / 6 PM)
```
python run_pipeline.py
```
Keep this process running continuously (systemd unit / NSSM or Task Scheduler on Windows / tmux). It sleeps between slots and only wakes at 00:00, 06:00, 12:00 and 18:00 IST to invoke:
```
python dashboard_status_pipeline0909_updated.py --mode live --excel Expected_sheet.xlsx --output Expected_sheet.xlsx
```
Each run writes `run_status.json` (success/failure, duration, last error) which the dashboard sidebar reads.

To trigger one run immediately without waiting for the next slot:
```
python run_pipeline.py --now
```

### Alternative: OS-level scheduler instead of run_pipeline.py
**Linux/Mac cron** (`crontab -e`):
```
0 0,6,12,18 * * * cd /path/to/qlik_monitor && /path/to/python dashboard_status_pipeline0909_updated.py --mode live --excel Expected_sheet.xlsx --output Expected_sheet.xlsx >> logs/cron.log 2>&1
```
**Windows Task Scheduler:** Daily trigger, "repeat every 6 hours", starting 12:00 AM, action = the same command above.
With either of these, `run_status.json` won't be created automatically — the sidebar will just show data freshness from the file's last-modified time, which still works fine.

## How data flows
`dashboard_status_pipeline...py` → writes `Expected_sheet.xlsx` → `app.py` detects the new file's mtime (cache key), reloads automatically within 60s (toggle in sidebar), and redraws every tab.
