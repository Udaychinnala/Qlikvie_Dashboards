"""Runs dashboard_status_pipeline0909_updated.py at 00:00, 06:00, 12:00, 18:00
and writes run_status.json so the Streamlit sidebar can show last-run health.

Keep this process alive (systemd service, Task Scheduler at logon, tmux, etc.)
- it does the waiting; it is not itself a cron entry.
"""
from __future__ import annotations

import json
import logging
import subprocess
import sys
import time
from datetime import datetime

from apscheduler.schedulers.blocking import BlockingScheduler
from apscheduler.triggers.cron import CronTrigger

import config

config.LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[
        logging.FileHandler(config.LOG_DIR / "scheduler.log"),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("scheduler")


def _write_status(ok: bool, started: datetime, finished: datetime, error: str = ""):
    payload = {
        "ok": ok,
        "started_at": started.strftime("%d %b, %I:%M %p"),
        "finished_at": finished.strftime("%d %b, %I:%M %p"),
        "duration_sec": round((finished - started).total_seconds(), 1),
        "error": error[-4000:],
    }
    config.STATUS_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def run_job():
    started = datetime.now()
    log.info("Pipeline run starting (mode=%s) -> %s", config.PIPELINE_MODE, config.DATA_FILE)
    cmd = [
        sys.executable, str(config.PIPELINE_SCRIPT),
        "--mode", config.PIPELINE_MODE,
        "--excel", str(config.DATA_FILE),
        "--output", str(config.DATA_FILE),  # overwrite in place; dashboard picks up the new mtime
    ]
    try:
        result = subprocess.run(
            cmd, capture_output=True, text=True,
            timeout=config.PIPELINE_TIMEOUT_MIN * 60,
        )
        finished = datetime.now()
        if result.returncode == 0:
            log.info("Pipeline finished OK in %.1fs", (finished - started).total_seconds())
            _write_status(True, started, finished)
        else:
            log.error("Pipeline failed (exit %s):\n%s", result.returncode, result.stderr[-4000:])
            _write_status(False, started, finished, result.stderr)
    except subprocess.TimeoutExpired:
        finished = datetime.now()
        log.error("Pipeline timed out after %s min", config.PIPELINE_TIMEOUT_MIN)
        _write_status(False, started, finished, f"Timed out after {config.PIPELINE_TIMEOUT_MIN} min")
    except Exception as exc:  # noqa: BLE001
        finished = datetime.now()
        log.exception("Pipeline crashed")
        _write_status(False, started, finished, str(exc))


def run_once_now():
    """`python run_pipeline.py --now` — run immediately instead of waiting for a slot."""
    run_job()


if __name__ == "__main__":
    if "--now" in sys.argv:
        run_once_now()
        raise SystemExit(0)

    scheduler = BlockingScheduler(timezone=config.TIMEZONE)
    scheduler.add_job(
        run_job,
        CronTrigger(hour=",".join(map(str, config.RUN_HOURS)), minute=0, timezone=config.TIMEZONE),
        misfire_grace_time=900,   # still fire if the machine was asleep/busy for up to 15 min
        coalesce=True,
    )
    log.info("Scheduler started. Runs at %s (%s). Writing to %s",
              ", ".join(f"{h:02d}:00" for h in config.RUN_HOURS), config.TIMEZONE, config.DATA_FILE)
    scheduler.start()
