"""Central settings - edit here (or override with environment variables)."""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# Workbook produced by your pipeline (the same file you pass as --excel).
DATA_FILE = Path(os.getenv("DASHBOARD_XLSX", BASE_DIR / "Expected_sheet.xlsx"))
# Shown (with a banner) only when DATA_FILE does not exist yet.
SAMPLE_FILE = BASE_DIR / "sample_data" / "Sample_Output_Dashboard_Alerts.xlsx"

# Pipeline runner
PIPELINE_SCRIPT = Path(os.getenv("PIPELINE_SCRIPT", BASE_DIR / "dashboard_status_pipeline0909_updated.py"))
PIPELINE_MODE = os.getenv("PIPELINE_MODE", "live")        # "live" (scrape) or "offline"
PIPELINE_TIMEOUT_MIN = int(os.getenv("PIPELINE_TIMEOUT_MIN", "90"))

# Schedule: 12 AM, 6 AM, 12 PM, 6 PM in this timezone
RUN_HOURS = (0, 6, 12, 18)
TIMEZONE = os.getenv("DASHBOARD_TZ", "Asia/Kolkata")

# Written by run_pipeline.py, shown in the dashboard sidebar
STATUS_FILE = DATA_FILE.with_name("run_status.json")
LOG_DIR = BASE_DIR / "logs"

# Data older than this (hours) is flagged as "stale" in the sidebar
STALE_AFTER_HOURS = 7
