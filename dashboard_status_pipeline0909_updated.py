#!/usr/bin/env python3
"""
dashboard_status_pipeline.py
=============================================================================
Merged pipeline: QlikView Dashboard scraper + QMC/QVW job scraper
                  + Dashboard vs. QVW validation & status engine.

This file combines the logic of:
  1) Dashboards_Dog.py  -> scrapes QlikView AccessPoint dashboard list
                            ("Last Updated" timestamps, IST).
  2) QVW_jobs.py         -> scrapes QMC task/job list
                            ("Last Execution" timestamps, EST/EDT).

...and adds a new validation layer that:
  - reads the Dashboard <-> QVW job mapping already maintained in the
    workbook ("Dashboard_status" sheet, grouped rows),
  - matches every mapped QVW job to its latest scraped status/timestamp,
  - converts QVW completion times from US/Eastern (EST/EDT) to IST,
  - decides a per-dashboard Status using the business rules below,
  - writes the result straight back into the Dashboard_status sheet,
    overwriting previous run's values.

MANUALLY-MAINTAINED "Periodicity" COLUMN
-----------------------------------------
"Dashboard_status" has a manually-maintained column (header contains
"periodicity", e.g. "Perodicity") that records each dashboard's expected
trigger timing for reference. The pipeline auto-detects this column by
header name (whatever it's actually called/spelled in the sheet), reads its
values BEFORE rebuilding the sheet, and writes them back unchanged in the
same position afterwards. The script never computes or overwrites this
column — it is pure passthrough.

HIGH-PRIORITY DAILY JOB MONITORING
------------------------------------
Jobs listed in the "Priority BI Jobs" sheet (Name + Perodicity/schedule
description) are treated as high-priority and are checked, on every run,
against their expected schedule:
  - "Delayed"  -> job is more than DELAY_THRESHOLD_MINUTES late against its
                  most recent expected run time and hasn't completed yet.
  - "Missed"   -> job never completed for a scheduled slot before the next
                  scheduled slot arrived (or before end of day, for
                  once-a-day jobs).
  - Any Failed/Aborted/Disabled QMC status on a priority job is also
    surfaced immediately, regardless of timing.
Results are logged (checked/flagged counts) on every run; they are NOT
written into the Dashboard_status sheet — that sheet no longer has a
"High Priority Job Alert" column (removed by request).

BUSINESS RULES
--------------
For each dashboard:
  * All mapped QVW jobs must have Status == "Success".
  * Dashboard refresh time (IST) must be at/after the latest QVW completion
    time (IST) — OR within a 3-minute tolerance window if it's slightly
    earlier (TIMESTAMP_TOLERANCE_MINUTES). Scrape/logging jitter between the
    AccessPoint and QMC timestamps means a refresh that genuinely picked up
    the new data can be recorded a couple of minutes before the QMC "Last
    Execution" entry for that same reload — without the tolerance this
    showed up as a false "Dashboard Pending Refresh".
  * Dashboard and all its QVW jobs must fall on the same business date
    (post timezone conversion).
  * A dashboard is only "Updated" once ALL of its related QVW jobs pass.

Resulting Dashboard Status values:
  - "Updated"                    -> all jobs succeeded & refresh is current
  - "QMC Running"                -> at least one dependent job still running
                                     (or waiting/queued/not yet confirmed)
  - "QMC Failed"                 -> at least one dependent job failed/aborted
  - "Dashboard Pending Refresh"  -> jobs all succeeded but dashboard hasn't
                                     picked up the new data yet
  - "Dashboard Refresh Missing"  -> no "Last Updated" timestamp found for
                                     the dashboard at all

QMC TIMINGS -> SCHEDULE (IST), PERIODICITY, REFRESH WINDOW, TEAMS ALERTS
-----------------------------------------------------------------------
"Dashboard_status" now has two more columns right after "Dashboard Name":
  * "QMC Timings" -- MANUALLY maintained, US/Eastern clock times (one or many,
    e.g. "7:30 AM, 4:30 PM"). Pure passthrough, exactly like Periodicity.
  * "Schedule"    -- COMPUTED every run: those times converted to IST with a
    timezone-aware (America/New_York -> Asia/Kolkata) conversion, so DST is
    handled automatically (EST: ET + 10:30, EDT: ET + 9:30) for the run date.

For every dashboard that is EXPECTED TO RUN TODAY (per its Periodicity), the
latest dashboard refresh (IST) is compared with each scheduled IST run time
using a +/-3 hour window. Dashboards not scheduled today are ignored.
Everything that needs attention is written to "Dashboard_Alerts_PA"
(cleared and rebuilt on every run; headers/formatting untouched) for Power
Automate to turn into Teams cards. See SECTION 6b / 7c below.

RUN MODES
---------
--mode offline (default)
    Does NOT open a browser. Reads dashboard timestamps from the
    "Dashboards timing" sheet and job data from the "Jobs" sheet of the
    SAME workbook (i.e. the output of a previous scrape). Useful for
    testing / re-running the validation logic on already-scraped data.

--mode live
    Runs the two Selenium scrapers first (dashboard AccessPoint + QMC),
    refreshes "Dashboards timing" and "Jobs" with fresh data, and then
    runs the same validation logic on the freshly scraped data.

Usage:
    python dashboard_status_pipeline.py --mode offline --excel Expected_sheet.xlsx
    python dashboard_status_pipeline.py --mode live     --excel Expected_sheet.xlsx
    python dashboard_status_pipeline.py --mode offline  --excel Expected_sheet.xlsx --now "2026-11-03 10:00"
        (--now overrides "current time" [IST] - handy for testing DST/periodicity)

Credentials (live mode only) are read from environment variables -- do NOT
hardcode them in this file:
    QLIK_USERNAME, QLIK_PASSWORD, QLIK_SECURITY_ANSWER   (AccessPoint login)
    QMC_USERNAME, QMC_PASSWORD                           (QMC basic-auth
                                                            fallback, only
                                                            used if the QMC
                                                            login prompt
                                                            appears instead
                                                            of NTLM pass-through)
=============================================================================
"""

from __future__ import annotations

import argparse
import calendar
import hashlib
import logging
import os
import re
import sys
import time
import difflib
import urllib.parse
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime, date, time as dtime, timedelta
from typing import Optional

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except ImportError:  # pragma: no cover
    from backports.zoneinfo import ZoneInfo  # type: ignore

import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# =============================================================================
# CONFIG
# =============================================================================

DASHBOARD_URL = "https://qlik.corp.wabtec.com/qlikview/index.htm"
DASHBOARD_BASE_URL = "https://qlik.corp.wabtec.com"
QMC_URL = "http://qlikfc.corp.wabtec.com/QMC/default.htm#"

# Credentials for the automated AccessPoint login (--mode live only).
# Read ONLY from environment variables — never hardcode a real username or
# password in this file. Set them once per machine, e.g.:
#   setx QLIK_USERNAME "your-username"     (Windows, persists across terminals)
#   setx QLIK_PASSWORD "your-password"
#   setx QLIK_SECURITY_ANSWER "your-answer"
# (macOS/Linux: use `export VAR=value` in your shell profile instead.)
QLIK_USERNAME = os.environ.get("QLIK_USERNAME", "uday.chinnala")
QLIK_PASSWORD = os.environ.get("QLIK_PASSWORD", "Wissen@1226@@")
QLIK_SECURITY_ANSWER = os.environ.get("QLIK_SECURITY_ANSWER", "udaykumar")

# Credentials for the automated QMC basic-auth login (--mode live only).
# QMC normally authenticates via Windows-integrated auth (NTLM) with no
# credential prompt at all — see scrape_qmc_jobs(). These are only used as a
# fallback when NTLM pass-through isn't available and a login prompt shows
# up instead. Leave unset to keep the previous "log in manually if prompted"
# behaviour. Reused from QLIK_USERNAME/QLIK_PASSWORD if QMC uses the same
# corporate account; set QMC_USERNAME/QMC_PASSWORD separately if not.
QMC_USERNAME = os.environ.get("QMC_USERNAME", QLIK_USERNAME)
QMC_PASSWORD = os.environ.get("QMC_PASSWORD", QLIK_PASSWORD)

# Sheet names inside the workbook (matches the actual workbook layout).
SHEET_DASHBOARD_TIMING = "Dashboards timing"   # dashboard "Last Updated" (IST)
SHEET_JOBS = "Jobs"                            # all QMC jobs (EST/EDT)
SHEET_MAPPING_AND_OUTPUT = "Dashboard_status"  # mapping + where results go
SHEET_PRIORITY_JOBS = "Priority BI Jobs"       # high-priority daily jobs to monitor

# --- High-priority daily job monitoring ------------------------------------
# A completed-but-late run past this many minutes after its scheduled time is
# flagged "Delayed". Past the NEXT scheduled slot (or end of day, for jobs
# that only run once a day) with still no successful completion, it becomes
# "Missed" instead.
DELAY_THRESHOLD_MINUTES = 30
HIGH_PRIORITY_ALERT_HEADER = "High Priority Job Alert"

# --- Dashboard vs. QVW completion timing tolerance --------------------------
# A dashboard refresh that lands up to this many minutes BEFORE the QVW job's
# recorded completion time is still treated as "at/after" it — scrape/logging
# jitter between the AccessPoint and QMC timestamps means a refresh that
# genuinely picked up the new data can be timestamped a couple of minutes
# earlier than the QMC "Last Execution" entry for the same reload.
TIMESTAMP_TOLERANCE_MINUTES = 3
TIMESTAMP_TOLERANCE = timedelta(minutes=TIMESTAMP_TOLERANCE_MINUTES)

EST_TZ = ZoneInfo("America/New_York")   # handles EST/EDT automatically
IST_TZ = ZoneInfo("Asia/Kolkata")
ET_TZ = EST_TZ   # alias: "ET" = US/Eastern; zoneinfo picks EST (UTC-5) / EDT (UTC-4) per date

# --- QMC schedule validation & Teams alert sheet ---------------------------------
SHEET_ALERTS = "Dashboard_Alerts_PA"          # consumed by Power Automate
# A dashboard counts as refreshed for a scheduled run if its refresh timestamp
# lies within scheduled IST time +/- this tolerance.
REFRESH_WINDOW_TOLERANCE = timedelta(hours=3)
# Business rule: dashboards NOT scheduled today are ignored entirely. Set True
# to still raise "Qlik reload failure" alerts for a failed QMC job on days the
# dashboard isn't scheduled (e.g. a weekly job that failed on Saturday and is
# still Failed on Tuesday).
ALERT_ON_FAILURE_WHEN_NOT_SCHEDULED = False
# Power Automate flips "Notified (Y/N)" to "Y" after posting. Because the alert
# sheet is rebuilt on every run, carry that "Y" forward for alerts whose
# "Alert Key" is unchanged - otherwise every run would re-post every alert.
PRESERVE_NOTIFIED_FLAG = True

DEFAULT_EXCEL_PATH = "Expected_sheet.xlsx"

# --- QMC status semantics -------------------------------------------------
# QMC tasks spend almost all of their life in "Waiting" (idle, waiting for
# the next scheduled trigger) -- that is the NORMAL state right after a
# successful run, not a sign of trouble. "Success" is only shown fleetingly
# right after a reload finishes, so a snapshot will rarely contain it.
# What actually tells us whether the job's data is current is:
#   - its Status is not Failed/Aborted/Running/Disabled, AND
#   - it HAS a real "Last Execution" timestamp (i.e. it has completed at
#     least once — "Never" means it hasn't, so its data can't be trusted).
FAILURE_STATUSES = {"failed", "aborted"}
RUNNING_STATUSES = {"running"}
DISABLED_STATUSES = {"disabled"}
# Statuses that represent "idle after a completed run" -- only counted as a
# genuine success if a Last Execution timestamp was actually parsed.
IDLE_COMPLETED_STATUSES = {"waiting", "success"}

# =============================================================================
# LOGGING
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("dashboard_status_pipeline.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("dashboard_status_pipeline")


# =============================================================================
# SECTION 1 — DASHBOARD SCRAPER  (merged from Dashboards_Dog.py)
# =============================================================================

def scrape_dashboards() -> list[dict]:
    """
    Logs into the QlikView AccessPoint and scrapes the dashboard list
    (Name, URL, Category, Last Updated). Returns a list of dicts.
    Only used in --mode live.
    """
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait, Select
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import TimeoutException, NoSuchElementException
    from bs4 import BeautifulSoup

    def init_browser():
        opts = webdriver.ChromeOptions()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-blink-features=AutomationControlled")
        opts.add_experimental_option("excludeSwitches", ["enable-automation"])
        opts.add_experimental_option("prefs", {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
        })
        return webdriver.Chrome(options=opts)

    def find_field(driver, *selectors, timeout=10):
        end = time.time() + timeout
        while time.time() < end:
            for sel in selectors:
                try:
                    el = driver.find_element(By.CSS_SELECTOR, sel)
                    if el.is_displayed():
                        return el
                except NoSuchElementException:
                    pass
            time.sleep(0.5)
        return None

    def try_fill(driver, value, *selectors):
        el = find_field(driver, *selectors)
        if el:
            el.clear()
            el.send_keys(value)
            return True
        return False

    def try_click(driver, *selectors, timeout=10):
        end = time.time() + timeout
        while time.time() < end:
            for sel in selectors:
                try:
                    el = driver.find_element(By.CSS_SELECTOR, sel)
                    if el.is_displayed() and el.is_enabled():
                        el.click()
                        return True
                except Exception:
                    pass
            time.sleep(0.5)
        return False

    USERNAME_SELECTORS = [
        "input[name='identifier']", "input[id='okta-signin-username']",
        "input[name='UserName']", "input[name='username']", "input[name='user']",
        "input[id='username']", "input[id='UserName']",
        "input[type='text']", "input[type='email']",
    ]
    PASSWORD_SELECTORS = [
        "input[name='credentials.passcode']", "input[id='okta-signin-password']",
        "input[name='Password']", "input[name='password']",
        "input[id='password']", "input[id='Password']", "input[type='password']",
    ]
    SUBMIT_SELECTORS = [
        "input[type='submit']", "button[type='submit']",
        "input[id='submitButton']", "button[id='submitButton']",
        "#submitButton", "#okta-signin-submit", "button.btn-primary",
        "input.button-primary", "button.button-primary",
        "input[value='Sign in']", "input[value='Log in']", "input[value='Login']",
        "input[value='Next']", "input[value='Verify']",
    ]

    def submit_current_step(driver):
        clicked = try_click(driver, *SUBMIT_SELECTORS, timeout=5)
        if not clicked:
            from selenium.webdriver.common.keys import Keys
            el = driver.switch_to.active_element
            try:
                el.send_keys(Keys.RETURN)
                clicked = True
            except Exception:
                pass
        return clicked

    def select_security_question_factor(driver):
        """
        On Okta's 'Verify it's you with a security method' screen, click the
        'Select' button for the Security Question row (not Phone). Tries an
        exact selector first (confirmed via DevTools), then falls back to a
        shadow-DOM-aware JS search, then a plain-DOM XPath search.
        """
        exact = [
            "div[data-se='security_question'] a[data-se='button']",
            "a[aria-label='Select Security Question.']",
            "div[data-se='security_question'] a.select-factor",
        ]
        if try_click(driver, *exact, timeout=8):
            return True

        js = r"""
        function collectAll(root, out) {
            var children = root.children ? Array.prototype.slice.call(root.children) : [];
            for (var i = 0; i < children.length; i++) {
                var node = children[i];
                out.push(node);
                if (node.shadowRoot) { collectAll(node.shadowRoot, out); }
                collectAll(node, out);
            }
        }
        function getRealParent(node) {
            var p = node.parentNode;
            if (!p) return null;
            if (p.nodeType === 11) { return p.host || null; }
            return p;
        }
        function ownText(el) { return (el.textContent || '').trim(); }
        var all = [];
        collectAll(document.body, all);
        var selectBtns = all.filter(function(el) {
            return (el.tagName === 'A' || el.tagName === 'BUTTON')
                && ownText(el) === 'Select' && el.offsetParent !== null;
        });
        var secQLabels = all.filter(function(el) {
            return ownText(el).indexOf('Security Question') !== -1 && el.children.length === 0;
        });
        function containsNode(container, target) {
            try { return container.contains && container.contains(target); }
            catch (e) { return false; }
        }
        var resultBtn = null;
        outer:
        for (var li = 0; li < secQLabels.length; li++) {
            var cur = secQLabels[li];
            for (var i = 0; i < 12; i++) {
                cur = getRealParent(cur);
                if (!cur) break;
                var matches = selectBtns.filter(function(b) { return containsNode(cur, b); });
                if (matches.length === 1) { resultBtn = matches[0]; break outer; }
                if (matches.length > 1) { break; }
            }
        }
        if (!resultBtn && selectBtns.length >= 2) { resultBtn = selectBtns[1]; }
        if (!resultBtn && selectBtns.length === 1) { resultBtn = selectBtns[0]; }
        if (resultBtn) { resultBtn.click(); return true; }
        return false;
        """
        try:
            if driver.execute_script(js):
                return True
        except Exception as e:
            log.warning("Shadow-DOM MFA click failed: %s", e)

        try:
            label_nodes = driver.find_elements(By.XPATH, "//*[contains(text(),'Security Question')]")
            for node in label_nodes:
                current = node
                for _ in range(8):
                    try:
                        current = current.find_element(By.XPATH, "..")
                    except Exception:
                        break
                    try:
                        selects = current.find_elements(
                            By.XPATH,
                            ".//a[normalize-space()='Select'] | .//button[normalize-space()='Select']")
                    except Exception:
                        selects = []
                    visible = [s for s in selects if s.is_displayed()]
                    if len(visible) == 1:
                        visible[0].click()
                        return True
                    elif len(visible) > 1:
                        break
        except Exception as e:
            log.warning("XPath MFA fallback failed: %s", e)
        return False

    def automated_login(driver) -> bool:
        log.info("Attempting automated dashboard login...")
        if not (QLIK_USERNAME and QLIK_PASSWORD and QLIK_SECURITY_ANSWER):
            log.warning(
                "QLIK_USERNAME / QLIK_PASSWORD / QLIK_SECURITY_ANSWER env vars "
                "not fully set — falling back to manual login in the browser."
            )
        try_fill(driver, QLIK_USERNAME, *USERNAME_SELECTORS)
        submit_current_step(driver)

        pwd_el = find_field(driver, *PASSWORD_SELECTORS, timeout=15)
        if pwd_el:
            pwd_el.clear()
            pwd_el.send_keys(QLIK_PASSWORD)
            submit_current_step(driver)

        end = time.time() + 15
        match = None
        stable_password_hits = 0
        while time.time() < end:
            if driver.find_elements(By.CSS_SELECTOR, "div[data-se='security_question']"):
                match = "mfa_select"
                break
            if driver.find_elements(By.ID, "listArea"):
                match = "listarea"
                break
            body_text = ""
            try:
                body_text = driver.find_element(By.TAG_NAME, "body").text
            except Exception:
                pass
            if "Verify with your Security Question" in body_text:
                match = "security_question"
                break
            if "Verify with your password" in body_text:
                stable_password_hits += 1
                if stable_password_hits >= 4:
                    match = "password_stuck"
                    break
            else:
                stable_password_hits = 0
            time.sleep(0.5)

        if match == "mfa_select":
            select_security_question_factor(driver)
            match = "security_question"

        if match == "security_question":
            sec_el = find_field(driver, *PASSWORD_SELECTORS, timeout=10)
            if sec_el:
                sec_el.clear()
                sec_el.send_keys(QLIK_SECURITY_ANSWER)
                submit_current_step(driver)

        log.info("Waiting for QlikView app list to confirm login...")
        try:
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "listArea")))
            log.info("Dashboard login successful.")
            return True
        except TimeoutException:
            log.warning("App list did not appear — waiting up to 60s for manual login.")
            for _ in range(60):
                time.sleep(1)
                try:
                    driver.find_element(By.ID, "listArea")
                    log.info("App list detected after manual intervention.")
                    return True
                except NoSuchElementException:
                    pass
            return False

    def ensure_list_view(driver):
        try:
            btn = driver.find_element(By.ID, "listView")
            if "selected" not in (btn.get_attribute("class") or ""):
                btn.click()
                time.sleep(1)
        except NoSuchElementException:
            pass

    def try_show_all(driver) -> bool:
        try:
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "Pagesize")))
            Select(driver.find_element(By.ID, "Pagesize")).select_by_value("0")
            time.sleep(5)
            return True
        except Exception:
            return False

    def scroll_to_load_all(driver):
        def count_rows():
            return driver.execute_script(
                "return (document.querySelectorAll('#appList li[id^=\"appList\"]') || []).length;"
            )

        def scroll_down():
            driver.execute_script("""
                var el = document.getElementById('listArea') || document.getElementById('appList');
                if (el && el.scrollHeight > el.clientHeight) { el.scrollTop = el.scrollHeight; }
                else { window.scrollTo(0, document.body.scrollHeight); }
            """)

        prev_count, stall_rounds, max_stall = 0, 0, 5
        while stall_rounds < max_stall:
            scroll_down()
            time.sleep(1.5)
            current_count = count_rows()
            if current_count > prev_count:
                prev_count, stall_rounds = current_count, 0
            else:
                stall_rounds += 1

    def get_listarea_html(driver) -> str:
        try:
            return driver.find_element(By.ID, "listArea").get_attribute("innerHTML")
        except NoSuchElementException:
            return ""

    def parse_apps(html: str) -> list[dict]:
        soup = BeautifulSoup(html, "html.parser")
        items = soup.select("ul#appList > li[id^='appList']")
        apps = []
        for li in items:
            name_tag = li.select_one("span.gridInfo span.docRow a.name") or li.select_one("a.name")
            name = name_tag.get_text(strip=True) if name_tag else ""
            href = name_tag.get("href", "") if name_tag else ""
            url = f"{DASHBOARD_BASE_URL}{href}" if href.startswith("/") else href
            cat_tag = li.select_one("span.attr2")
            category = cat_tag.get_text(strip=True) if cat_tag else ""
            upd_tag = li.select_one("span.attr4")
            if upd_tag:
                for img in upd_tag.find_all("img"):
                    img.decompose()
                last_updated = upd_tag.get_text(strip=True)
            else:
                last_updated = ""
            if name:
                apps.append({"Name": name, "URL": url, "Category": category, "Last Updated": last_updated})
        return apps

    driver = init_browser()
    try:
        driver.get(DASHBOARD_URL)
        if not automated_login(driver):
            raise RuntimeError("Dashboard login failed")
        time.sleep(2)
        ensure_list_view(driver)
        if not try_show_all(driver):
            scroll_to_load_all(driver)
        else:
            scroll_to_load_all(driver)  # safety net, matches original script
        html = get_listarea_html(driver)
        log.info("Fetched #listArea innerHTML (%d chars).", len(html))
        apps = parse_apps(html)
        log.info("Scraped %d dashboards from AccessPoint.", len(apps))
        if not apps:
            dump_path = os.path.abspath("dashboard_listarea_dump.html")
            try:
                with open(dump_path, "w", encoding="utf-8") as f:
                    f.write(html or "<!-- #listArea was not found / empty -->")
                log.error(
                    "0 dashboards parsed from AccessPoint. Dumped the raw #listArea HTML to "
                    "'%s' for inspection — if that file is empty/tiny, #listArea likely never "
                    "loaded (login may not have actually reached the app list, or List View "
                    "wasn't active). If it has content but 0 apps were parsed, AccessPoint's "
                    "markup ('ul#appList > li[id^=\"appList\"]') may have changed.", dump_path)
            except Exception as e:
                log.error("0 dashboards parsed from AccessPoint, and could not write debug dump: %s", e)
        return apps
    finally:
        driver.quit()


# =============================================================================
# SECTION 2 — QMC / QVW JOB SCRAPER  (merged from QVW_jobs.py)
# =============================================================================

def scrape_qmc_jobs() -> list[dict]:
    """
    Opens the QMC task overview (Windows-integrated auth — you log in once
    in the browser, matching the original script's behaviour) and scrapes
    every individual job row (Name, Executed On, Status, Distribution Group,
    Last Execution, Started/Scheduled). Returns a list of dicts.

    IMPORTANT: the QMC task tree is collapsible. Folder/group nodes show a
    "+" (Expand) icon and, until expanded, only render a rollup summary
    (e.g. "21 failed, 18 running") in the DOM rather than the real per-job
    rows underneath them. We expand every node before extracting, and skip
    any row that's still a collapsed folder, so we only capture real jobs.
    Only used in --mode live.
    """
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    options = Options()
    options.add_argument("--start-maximized")
    options.add_argument("--auth-server-whitelist=*.corp.wabtec.com,corp.wabtec.com")
    options.add_argument("--auth-negotiate-delegate-whitelist=*.corp.wabtec.com,corp.wabtec.com")

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 180)

    try:
        if QMC_USERNAME and QMC_PASSWORD:
            # The QMC "Sign in" prompt you get when NTLM pass-through isn't
            # active is a NATIVE Chrome basic-auth dialog, not part of the
            # page — it has no DOM elements, so it can't be automated with
            # find_element/send_keys the way the AccessPoint login form is.
            # Embedding credentials in the URL (https://user:pass@host/...)
            # is the standard Selenium/Chrome way to answer that dialog
            # automatically. Values are percent-encoded so special
            # characters (e.g. '@' in the password) don't get misparsed as
            # part of the host.
            # NOTE: a CDP-based Authorization-header injection was tried
            # here first but broke DNS resolution for this host on this
            # network (enabling CDP's Network domain appears to conflict
            # with how Chrome applies the corporate proxy/DNS here) — so
            # this URL-embedding approach is used instead.
            parsed = urllib.parse.urlsplit(QMC_URL)
            authed_netloc = "{}:{}@{}".format(
                urllib.parse.quote(QMC_USERNAME, safe=""),
                urllib.parse.quote(QMC_PASSWORD, safe=""),
                parsed.netloc,
            )
            authed_url = urllib.parse.urlunsplit(
                (parsed.scheme, authed_netloc, parsed.path, parsed.query, parsed.fragment)
            )
            log.info("Opening QMC URL with automated basic-auth credentials.")
            driver.get(authed_url)
        else:
            driver.get(QMC_URL)
            log.info("QMC URL opened — please log in manually in the browser window if prompted "
                      "(set QMC_USERNAME / QMC_PASSWORD env vars to automate this).")

        wait.until(EC.presence_of_element_located((By.ID, "StatusFilterDropDown")))
        log.info("QMC login/session confirmed.")

        # Disable auto refresh so the row set doesn't shift under us.
        try:
            for _ in range(3):
                driver.execute_script("""
                    var cb = document.getElementById('refreshcheckbox');
                    if (cb && cb.checked) cb.click();
                """)
                time.sleep(0.5)
        except Exception as e:
            log.warning("Auto refresh toggle skipped: %s", e)

        # Apply "All statuses" filter.
        try:
            status_btn = wait.until(EC.element_to_be_clickable((By.ID, "StatusFilterDropDown")))
            driver.execute_script("arguments[0].click();", status_btn)
            wait.until(EC.presence_of_element_located(
                (By.XPATH, "//input[contains(@id,'TaskStatusOverview.FilterByStatus')]")))
            driver.execute_script("""
                var boxes = document.querySelectorAll("input[id*='TaskStatusOverview.FilterByStatus']");
                boxes.forEach(function(cb) { if (!cb.checked) cb.click(); });
            """)
            ok_btn = wait.until(EC.element_to_be_clickable((By.ID, "FilterByStatus.Ok")))
            driver.execute_script("arguments[0].click();", ok_btn)
            wait.until(EC.invisibility_of_element_located((By.XPATH, "//*[contains(text(),'Processing')]")))
            time.sleep(2)
        except Exception:
            log.warning("Could not confirm 'all statuses' filter applied — continuing anyway.")

        log.info("Waiting for TreeRow elements...")
        wait.until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "div.TreeRow")) > 0)

        # ---- Expand every collapsed folder/group node -------------------------
        log.info("Expanding all tree nodes...")
        expand_script = """
            var imgs = document.querySelectorAll('.TreeImage[alt="Expand"]');
            imgs.forEach(function(img) { img.click(); });
            return imgs.length;
        """
        for expand_attempt in range(50):
            clicked = driver.execute_script(expand_script)
            if clicked == 0:
                log.info("No more collapsed nodes after %d expand pass(es).", expand_attempt)
                break
            time.sleep(0.7)
        else:
            log.warning("Stopped after 50 expand passes — some nodes may still be collapsed.")

        # ---- Scroll to load every row (also re-expands newly revealed nodes) --
        def get_row_count():
            return driver.execute_script("return document.querySelectorAll('div.TreeRow').length;")

        scroll_script = """
            var containers = [
                document.querySelector('#TaskStatusOverviewScrollContainer'),
                document.querySelector('[id*="ScrollContainer"]'),
                document.querySelector('[id*="InnerPage"]'),
                document.querySelector('[class*="scroll"]'),
                document.body
            ];
            var el = containers.find(c => c !== null);
            if (el) el.scrollTop += 600;
            return el ? el.id || el.className : 'body';
        """
        last_count, no_change_streak = 0, 0
        for scroll_attempt in range(200):
            driver.execute_script(scroll_script)
            driver.execute_script(expand_script)  # expand any newly-revealed folder nodes
            time.sleep(0.5)
            current_count = get_row_count()
            if scroll_attempt % 10 == 0:
                log.info("  Scroll #%d | Rows loaded: %d", scroll_attempt, current_count)
            if current_count == last_count:
                no_change_streak += 1
            else:
                no_change_streak, last_count = 0, current_count
            if no_change_streak >= 8:
                break
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        log.info("Final row count before extraction: %d", get_row_count())

        # ---- Extract every real (non-folder) job row ---------------------------
        # NOTE: QMC suffixes these element IDs differently depending on tree
        # depth/context (observed "TK" and "XS" so far). Match on "contains"
        # rather than a fixed prefix so it keeps working if the suffix changes.
        records = driver.execute_script(r"""
            var rows = document.querySelectorAll('div.TreeRow');
            var results = [];
            rows.forEach(function(row) {
                try {
                    var nameEl = row.querySelector('[id*="TextNode"]');
                    var name = nameEl ? (nameEl.getAttribute('title') || nameEl.innerText || '').trim() : '';

                    var nr5cells = row.querySelectorAll('[id*="Nr5TableCell"]');
                    var executedOn = '', status = '';
                    var distEl = row.querySelector('.DistributionGroupCol');
                    var distGroup = distEl ? distEl.innerText.trim() : '';
                    var nr5texts = [];
                    nr5cells.forEach(function(c) { nr5texts.push(c.innerText.trim()); });
                    nr5texts.forEach(function(t) {
                        var statusWords = ['Running','Waiting','Success','Failed','Aborted','Never','Warning','Disabled'];
                        var isStatus = statusWords.some(function(sw) { return t.indexOf(sw) !== -1; });
                        if (isStatus) { status = t; } else if (t.length > 0) { executedOn = t; }
                    });

                    var nr6El = row.querySelector('[id*="Nr6TableCell"]');
                    var lastExec = nr6El ? (nr6El.getAttribute('title') || nr6El.innerText || '').trim() : '';
                    var nr7El = row.querySelector('[id*="Nr7TableCell"]');
                    var scheduled = nr7El ? (nr7El.getAttribute('title') || nr7El.innerText || '').trim() : '';

                    if (!status) {
                        var iconEl = row.querySelector('[id*="NodeIconTreeImage"]');
                        status = iconEl ? iconEl.getAttribute('alt') || '' : '';
                    }

                    // Skip collapsed folder/group rows (still show a "+" expand
                    // icon and only hold a rollup summary, not a real job).
                    var expandImg = row.querySelector('.TreeImage[alt="Expand"]');
                    var isCollapsedFolder = !!expandImg;

                    if (name && !isCollapsedFolder) {
                        results.push([name, executedOn, status, distGroup, lastExec, scheduled]);
                    }
                } catch (e) {}
            });
            return results;
        """)

        headers = ["Name", "Executed On", "Status", "Distribution Group", "Last Execution", "Started/Scheduled"]
        jobs = [dict(zip(headers, r)) for r in records if r and r[0]]
        log.info("Scraped %d QMC job rows.", len(jobs))
        if not jobs:
            log.error("0 job rows extracted — the tree may still be collapsed or QMC's "
                      "internal element IDs changed again. Check the browser window.")
        return jobs
    finally:
        driver.quit()


# =============================================================================
# SECTION 3 — TIMEZONE / PARSING HELPERS
# =============================================================================

def parse_dashboard_timestamp(value) -> Optional[datetime]:
    """Parses a dashboard 'Last Updated' value (assumed already IST, naive)."""
    if value in (None, "", "Never"):
        return None
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y %H:%M", "%m/%d/%Y %I:%M:%S %p"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def parse_qvw_timestamp(value) -> Optional[datetime]:
    """Parses a QMC 'Last Execution' value (assumed naive US/Eastern)."""
    if value in (None, "", "Never", "Disabled", "Not scheduled"):
        return None
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    for fmt in ("%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def est_to_ist(dt: Optional[datetime]) -> Optional[datetime]:
    """Localizes a naive US/Eastern datetime and converts it to naive IST."""
    if dt is None:
        return None
    aware_est = dt.replace(tzinfo=EST_TZ)
    aware_ist = aware_est.astimezone(IST_TZ)
    return aware_ist.replace(tzinfo=None)


# --- QMC Timings (US/Eastern clock times) -> IST -------------------------------------------
class ScheduleConversionError(ValueError):
    """The 'QMC Timings' cell could not be turned into clock times."""


# One clock time inside free text: "7:15", "07:15:00", "7:30 AM", "8:05PM",
# "13:00 PM" (24h with a stray PM), "10 pm". A bare number with neither a
# colon nor an AM/PM suffix is NOT a time and is skipped.
_QMC_TIME_RE = re.compile(
    r"(?P<h>\d{1,2})(?::(?P<m>\d{2}))?(?::(?P<s>\d{2}))?\s*(?P<ap>[AaPp]\.?[Mm]\.?)?"
)
_MANUAL_ONLY_RE = re.compile(r"\b(manual|no\s+triggers?)\b", re.IGNORECASE)


def parse_qmc_timings(raw) -> list[dtime]:
    """Turns a 'QMC Timings' cell into a sorted, de-duplicated list of naive
    US/Eastern clock times.

    Accepts a datetime.time / datetime, an Excel time fraction, or free text
    with one or many times ("2:00 AM, 14:00PM", "\\xa010:00 PM", "13:00 PM").
    Returns [] for blank cells and "manual triggers only" text (nothing to
    convert). Raises ScheduleConversionError for text that has content but no
    parseable time, or an out-of-range time."""
    if raw is None:
        return []
    if isinstance(raw, datetime):
        return [raw.time().replace(microsecond=0)]
    if isinstance(raw, dtime):
        return [raw.replace(microsecond=0, tzinfo=None)]
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        if 0 <= raw < 1:  # Excel stores times as a fraction of a day
            secs = round(raw * 86400)
            return [dtime(secs // 3600 % 24, secs % 3600 // 60, secs % 60)]
        raise ScheduleConversionError(f"numeric QMC Timings value {raw!r} is not a time of day")

    text = str(raw).replace("\xa0", " ").strip()
    if not text:
        return []

    found: set[dtime] = set()
    for m in _QMC_TIME_RE.finditer(text):
        h, mi, sec, ap = m.group("h"), m.group("m"), m.group("s"), m.group("ap")
        if mi is None and ap is None:
            continue  # bare number, not a clock time
        hour, minute, second = int(h), int(mi or 0), int(sec or 0)
        if ap and 1 <= hour <= 12:
            hour = hour % 12 + (12 if ap.lower().startswith("p") else 0)
        # hour 0 or 13-23 with an AM/PM suffix is already 24h ("13:00 PM"): suffix ignored.
        if hour > 23 or minute > 59 or second > 59:
            raise ScheduleConversionError(f"invalid time {m.group(0).strip()!r} in QMC Timings {raw!r}")
        found.add(dtime(hour, minute, second))

    if found:
        return sorted(found)
    if _MANUAL_ONLY_RE.search(text):
        return []  # "Only manual triggers (No triggers as of now)"
    raise ScheduleConversionError(f"no clock time found in QMC Timings {raw!r}")


@dataclass(frozen=True)
class ExpectedRun:
    et_dt: datetime    # timezone-aware US/Eastern trigger time (EST or EDT)
    ist_dt: datetime   # the same instant as naive IST (codebase convention)


def et_to_ist(et_date: date, et_time: dtime) -> ExpectedRun:
    """DST-aware conversion of one Eastern wall-clock time on `et_date`.

    zoneinfo attaches EST (UTC-5) or EDT (UTC-4) according to `et_date`, so the
    IST result is ET+10:30 in winter and ET+9:30 in summer with no fixed
    offsets anywhere. On the two DST-change days a non-existent time (spring
    forward) or an ambiguous one (fall back) resolves with fold=0 - the
    first occurrence."""
    et_dt = datetime.combine(et_date, et_time, tzinfo=ET_TZ)
    return ExpectedRun(et_dt, et_dt.astimezone(IST_TZ).replace(tzinfo=None))


def slots_for_ist_day(et_times: list[dtime], ist_day: date) -> list[ExpectedRun]:
    """Every scheduled run that lands on calendar day `ist_day` in IST.

    IST is 9:30-10:30h ahead of ET, so an Eastern trigger belongs to the IST
    day either of the same date (early ET times) or of the NEXT date (late ET
    times, e.g. 10:00 PM ET = 07:30 IST next morning). Checking ET dates
    ist_day-1 and ist_day therefore finds each run exactly once."""
    runs = []
    for t in et_times:
        for et_date in (ist_day - timedelta(days=1), ist_day):
            run = et_to_ist(et_date, t)
            if run.ist_dt.date() == ist_day:
                runs.append(run)
    return sorted(runs, key=lambda r: r.ist_dt)


def _fmt_12h(t: dtime) -> str:
    """'4:45 PM' (no platform-specific strftime flags, so it also works on Windows)."""
    return f"{t.hour % 12 or 12}:{t.minute:02d} {'AM' if t.hour < 12 else 'PM'}"


# =============================================================================
# SECTION 4 — QVW <-> JOB MATCHING
# =============================================================================

_SUFFIX_RE = re.compile(r"\s*\((?:[^()]|\([^()]*\))*\.qvw\)\s*(\(work disabled\))?\s*$", re.IGNORECASE)
_PAREN_RE = re.compile(r"\(((?:[^()]|\([^()]*\))*\.qvw)\)\s*(\(work disabled\))?\s*$", re.IGNORECASE)


@dataclass
class JobRecord:
    raw_name: str
    no_suffix: str          # job path with the trailing "(dashboard.qvw)" stripped
    dashboard_hint: Optional[str]
    status: str
    last_execution: Optional[datetime]  # naive EST/EDT
    raw: dict = field(default_factory=dict)  # full original row (Name, Executed On,
                                              # Status, Distribution Group, Last
                                              # Execution, Started/Scheduled, ...) —
                                              # kept so any raw column can be copied
                                              # straight through (e.g. into "Priority
                                              # BI Jobs") without re-deriving it.


def build_job_index(jobs_raw: list[dict]) -> list[JobRecord]:
    records = []
    for row in jobs_raw:
        name = row.get("Name")
        if not name:
            continue
        no_suffix = _SUFFIX_RE.sub("", name).strip()
        m = _PAREN_RE.search(name)
        dash_hint = m.group(1).strip() if m else None
        records.append(JobRecord(
            raw_name=name,
            no_suffix=no_suffix,
            dashboard_hint=dash_hint,
            status=(row.get("Status") or "").strip(),
            last_execution=parse_qvw_timestamp(row.get("Last Execution")),
            raw=row,
        ))
    return records


def match_job(qvw_file: str, dashboard_name: str, job_index: list[JobRecord]) -> Optional[JobRecord]:
    """
    Matches a ".QVW files" mapping entry (from Dashboard_status sheet) to the
    scraped QMC job whose full path ends with that entry. Falls back to a
    fuzzy match if nothing matches exactly.
    """
    if not qvw_file:
        return None
    target = qvw_file.strip().lower()

    candidates = [j for j in job_index if j.no_suffix.lower().endswith(target)]
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        # Disambiguate using the dashboard name in the job's own "(...)" suffix.
        exact = [c for c in candidates if c.dashboard_hint and
                 c.dashboard_hint.lower() == str(dashboard_name).strip().lower()]
        if len(exact) == 1:
            return exact[0]
        log.warning("Ambiguous QVW match for '%s' (dashboard '%s') — %d candidates, using first.",
                    qvw_file, dashboard_name, len(candidates))
        return candidates[0]

    # Fuzzy fallback.
    all_paths = [j.no_suffix.lower() for j in job_index]
    close = difflib.get_close_matches(target, all_paths, n=1, cutoff=0.85)
    if close:
        for j in job_index:
            if j.no_suffix.lower() == close[0]:
                log.info("Fuzzy-matched QVW file '%s' -> '%s'", qvw_file, j.raw_name)
                return j
    return None


# =============================================================================
# SECTION 4b — HIGH-PRIORITY DAILY JOB MONITORING (Priority BI Jobs sheet)
# =============================================================================
# The "Priority BI Jobs" sheet lists high-priority jobs and a free-text
# description of when they're expected to run, e.g.:
#   "1 time a day, [ 10:30AM ]"
#   "2 times a day, [   03:45AM  &&   05:00AM   ]"
#   "Once every hour, starting at  8:05:00 AM"
#   "1 time a day, Upon success of \"X\" [05:45 Am roughfly]"
# We don't try to fully model dependency chains ("Upon success of ...") —
# any bracketed/plain clock time found in the text is treated as an expected
# daily anchor time, which is good enough to catch real delays/misses.

_TIME_TOKEN_RE = re.compile(r"(\d{1,2})[:.](\d{2})(?::(\d{2}))?\s*([AaPp][Mm])")
_HOURLY_RE = re.compile(r"every\s+hour", re.IGNORECASE)


def parse_schedule_times(periodicity_text: Optional[str], ref_date: date) -> list[datetime]:
    """Parses a free-text schedule description into a sorted list of naive
    IST datetimes for `ref_date`. Returns [] if nothing could be parsed
    (e.g. "Never", "Disabled", or unrecognized text)."""
    if not periodicity_text:
        return []
    text = str(periodicity_text).strip()
    if not text or text.lower() in ("never", "disabled", "n/a"):
        return []

    times: list[dtime] = []
    for h, m, s, ampm in _TIME_TOKEN_RE.findall(text):
        hour = int(h) % 12
        if ampm.lower() == "pm":
            hour += 12
        try:
            times.append(dtime(hour, int(m), int(s) if s else 0))
        except ValueError:
            continue

    if not times:
        return []

    if _HOURLY_RE.search(text):
        # "Once every hour, starting at <time>" -> expand hourly across the day.
        start = min(times)
        out = []
        t = datetime.combine(ref_date, start)
        end = datetime.combine(ref_date, dtime(23, 59, 59))
        while t <= end:
            out.append(t)
            t += timedelta(hours=1)
        return out

    return sorted(datetime.combine(ref_date, t) for t in times)


def evaluate_priority_job(
    job_name: str,
    periodicity_text: Optional[str],
    matched_job: Optional["JobRecord"],
    now_ist: datetime,
) -> Optional[str]:
    """Returns an alert string if a high-priority job is delayed, missed, or
    currently in a failed/aborted/disabled state; None if it's on time, not
    yet due, or its schedule couldn't be parsed."""
    status = (matched_job.status or "").strip().lower() if matched_job else ""
    if status in FAILURE_STATUSES | DISABLED_STATUSES:
        return (f"High-priority job {matched_job.status!r} — last run did not "
                f"complete successfully; verify immediately.")

    schedule = parse_schedule_times(periodicity_text, now_ist.date())
    if not schedule:
        return None  # unparseable/undated schedule — nothing we can check

    due_slots = [t for t in schedule if t <= now_ist]
    if not due_slots:
        return None  # first scheduled run of the day hasn't arrived yet

    most_recent = max(due_slots)
    later_slots = [t for t in schedule if t > most_recent]
    missed_cutoff = later_slots[0] if later_slots else datetime.combine(now_ist.date(), dtime(23, 59, 59))

    actual = est_to_ist(matched_job.last_execution) if matched_job else None
    same_day_actual = actual if (actual and actual.date() == now_ist.date()) else None

    if same_day_actual is None or same_day_actual < most_recent:
        # Nothing has completed for this slot yet.
        if now_ist >= missed_cutoff:
            return (f"Missed run — expected ~{most_recent:%H:%M} IST, still not completed "
                    f"and the next scheduled run ({missed_cutoff:%H:%M} IST) has passed.")
        delay_min = int((now_ist - most_recent).total_seconds() // 60)
        if delay_min > DELAY_THRESHOLD_MINUTES:
            return (f"Delayed — expected ~{most_recent:%H:%M} IST, not completed yet "
                    f"({delay_min} min late so far).")
        return None  # within grace period

    delay_min = int((same_day_actual - most_recent).total_seconds() // 60)
    if delay_min > DELAY_THRESHOLD_MINUTES:
        if same_day_actual >= missed_cutoff:
            return (f"Missed run — expected ~{most_recent:%H:%M} IST, only completed at "
                    f"{same_day_actual:%H:%M} IST, after the next scheduled run.")
        return (f"Delayed — expected ~{most_recent:%H:%M} IST, completed {same_day_actual:%H:%M} "
                f"IST ({delay_min} min late).")
    return None  # on time


def build_priority_alerts(
    ws_priority: Worksheet, job_index: list["JobRecord"], now_ist: Optional[datetime] = None,
) -> dict[str, str]:
    """Reads the 'Priority BI Jobs' sheet and returns {job_name_lower: alert
    message} for every monitored job currently showing a delay, miss, or
    failure. Jobs that are on time / not due yet are omitted entirely."""
    now_ist = now_ist or datetime.now(IST_TZ).replace(tzinfo=None)

    header_row = None
    for row_idx in range(1, ws_priority.max_row + 1):
        if _normalize_header(ws_priority.cell(row=row_idx, column=1).value) == "name":
            header_row = row_idx
            break
    if header_row is None:
        log.warning("'%s' sheet has no recognizable 'Name' header — skipping priority job monitoring.",
                    SHEET_PRIORITY_JOBS)
        return {}

    col_name = 1
    col_period = _find_header_col(ws_priority, "periodicity", "perodicity", header_row=header_row) or 2

    alerts: dict[str, str] = {}
    checked = 0
    for row_idx in range(header_row + 1, ws_priority.max_row + 1):
        name = ws_priority.cell(row=row_idx, column=col_name).value
        if not name:
            continue
        name = str(name).strip()
        periodicity = ws_priority.cell(row=row_idx, column=col_period).value
        matched = match_job(name, "", job_index)
        checked += 1
        msg = evaluate_priority_job(name, periodicity, matched, now_ist)
        if msg:
            alerts[name.strip().lower()] = msg

    log.info("Checked %d high-priority job(s); %d currently flagged (delayed/missed/failed).",
              checked, len(alerts))
    return alerts


# --- Auto-populate the "Priority BI Jobs" sheet from the QMC job index -----
# The sheet only has to list job Names (+ optionally a Perodicity schedule
# description, maintained by hand). Everything else — Executed On, Status,
# Distribution Group, Last Execution, Started/Scheduled — is looked up here
# from the freshly scraped/loaded "Jobs" data on every run, using the same
# match_job() logic already used for QVW <-> dashboard matching. Add or
# remove a row under "Name" and re-run the pipeline to pick it up.
_PRIORITY_DETAIL_COLUMNS = [
    ("Executed On", "executedon"),
    ("Status", "status"),
    ("Distribution Group", "distributiongroup"),
    ("Last Execution", "lastexecution"),
    ("Started/Scheduled", "startedscheduled"),
]


def populate_priority_jobs_sheet(ws_priority: Worksheet, job_index: list["JobRecord"]) -> int:
    """Fills in Executed On / Status / Distribution Group / Last Execution /
    Started-Scheduled for every job Name listed in the 'Priority BI Jobs'
    sheet, by matching it against the scraped QMC job list. Jobs with no
    match get "Not Found in QMC" written into Status (and the other detail
    columns cleared) so a typo'd/removed job name is obvious at a glance.
    Returns the number of configured jobs successfully matched."""
    header_row = None
    for row_idx in range(1, ws_priority.max_row + 1):
        if _normalize_header(ws_priority.cell(row=row_idx, column=1).value) == "name":
            header_row = row_idx
            break
    if header_row is None:
        log.warning("'%s' sheet has no recognizable 'Name' header — skipping auto-populate.",
                    SHEET_PRIORITY_JOBS)
        return 0

    col_name = 1
    detail_cols = {}
    for label, norm in _PRIORITY_DETAIL_COLUMNS:
        col = _find_header_col(ws_priority, norm, header_row=header_row)
        if col is None:
            log.warning("'%s' sheet has no '%s' column — skipping that field.",
                        SHEET_PRIORITY_JOBS, label)
        detail_cols[label] = col

    matched = 0
    checked = 0
    for row_idx in range(header_row + 1, ws_priority.max_row + 1):
        name = ws_priority.cell(row=row_idx, column=col_name).value
        if not name:
            continue
        name = str(name).strip()
        checked += 1
        job = match_job(name, "", job_index)

        if job is None:
            log.warning("Priority job '%s' not found in current QMC job data.", name)
            for label in ("Executed On", "Distribution Group", "Last Execution", "Started/Scheduled"):
                col = detail_cols.get(label)
                if col:
                    ws_priority.cell(row=row_idx, column=col, value=None)
            col = detail_cols.get("Status")
            if col:
                ws_priority.cell(row=row_idx, column=col, value="Not Found in QMC")
            continue

        matched += 1
        for label, _ in _PRIORITY_DETAIL_COLUMNS:
            col = detail_cols.get(label)
            if col:
                ws_priority.cell(row=row_idx, column=col, value=job.raw.get(label))

    log.info("'%s' sheet: matched %d/%d configured job(s) against QMC data.",
              SHEET_PRIORITY_JOBS, matched, checked)
    return matched


# =============================================================================
# SECTION 5 — READ DASHBOARD <-> QVW MAPPING (grouped rows in Dashboard_status)
# =============================================================================

@dataclass
class DashboardGroup:
    sl_no: Optional[int]
    dashboard_name: str
    bi_dna: Optional[str]
    qvw_files: list[str] = field(default_factory=list)
    first_row: int = 0     # 1-based row index of the group's header row in the sheet
    row_count: int = 1     # number of rows this group occupies
    periodicity: Optional[object] = None   # manually-maintained value, passed through as-is
    qmc_timings: Optional[object] = None   # manually-maintained US/Eastern time(s): datetime.time, text, or None
    qmc_timings_fmt: str = "General"       # its Excel number format, so times still display as e.g. "7:15 AM"


# --- Header auto-detection ---------------------------------------------
# The sheet's actual column layout has drifted from a fixed A/B/D/F scheme
# (columns get manually inserted/removed, e.g. the Periodicity column), so
# every read/write below locates columns by header text instead of a
# hardcoded index. This also means the pipeline is no longer broken by
# future manual column insertions.
def _normalize_header(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").strip().lower())


def _find_header_col(ws: Worksheet, *candidates: str, header_row: int = 1) -> Optional[int]:
    """Finds a column by header text. `candidates` are normalized (lowercased,
    punctuation/space-stripped) substrings to match against each header cell,
    e.g. "periodicity" also matches the sheet's actual "Perodicity" typo."""
    norm_candidates = [_normalize_header(c) for c in candidates]
    for col in range(1, ws.max_column + 1):
        header_norm = _normalize_header(ws.cell(row=header_row, column=col).value)
        if not header_norm:
            continue
        for cand in norm_candidates:
            if cand in header_norm or header_norm in cand:
                return col
    return None


def read_mapping(ws_map: Worksheet) -> tuple[list[DashboardGroup], Optional[str]]:
    """
    Reads the grouped mapping in Dashboard_status. Columns are located by
    header text (see _find_header_col) rather than fixed positions:
        "Dashboard Name"  (only on first row of a group)
        ".QVW files"      (one job per row)
        "BI/DnA"          (only on first row of a group)
        "Sl.NO"           (optional, only on first row of a group)
        "Periodicity"/"Perodicity" (optional, manually maintained,
                                     only on first row of a group)
        "QMC Timings"     (optional, manually maintained US/Eastern trigger
                           time(s), first row of a group; passed through
                           unchanged and used to compute "Schedule")

    Returns (groups, periodicity_header_text). periodicity_header_text is
    the EXACT header string found in the sheet (so it can be written back
    unchanged, typo and all), or None if no such column exists.
    """
    col_slno = _find_header_col(ws_map, "sl no", "slno")
    col_name = _find_header_col(ws_map, "dashboard name")
    col_qvw = _find_header_col(ws_map, "qvw files")
    col_bidna = _find_header_col(ws_map, "bi/dna", "bidna")
    col_period = _find_header_col(ws_map, "periodicity", "perodicity")
    col_qmc = _find_header_col(ws_map, "qmc timings", "qmctimings")

    if col_name is None:
        raise KeyError("Could not find a 'Dashboard Name' column header in the mapping sheet.")
    if col_qvw is None:
        raise KeyError("Could not find a '.QVW files' column header in the mapping sheet.")

    periodicity_header_text = None
    if col_period is not None:
        periodicity_header_text = str(ws_map.cell(row=1, column=col_period).value).strip()

    groups: list[DashboardGroup] = []
    current: Optional[DashboardGroup] = None

    for row_idx in range(2, ws_map.max_row + 1):
        sl_no = ws_map.cell(row=row_idx, column=col_slno).value if col_slno else None
        dash_name = ws_map.cell(row=row_idx, column=col_name).value
        qvw_file = ws_map.cell(row=row_idx, column=col_qvw).value
        bi_dna = ws_map.cell(row=row_idx, column=col_bidna).value if col_bidna else None
        period_val = ws_map.cell(row=row_idx, column=col_period).value if col_period else None
        qmc_cell = ws_map.cell(row=row_idx, column=col_qmc) if col_qmc else None

        if dash_name:  # start of a new group
            if current is not None:
                groups.append(current)
            current = DashboardGroup(
                sl_no=sl_no, dashboard_name=str(dash_name).strip(),
                bi_dna=bi_dna, first_row=row_idx, periodicity=period_val,
                qmc_timings=qmc_cell.value if qmc_cell is not None else None,
                qmc_timings_fmt=qmc_cell.number_format if qmc_cell is not None else "General",
            )
        if current is None:
            continue  # stray row before any dashboard header — skip

        if qvw_file:
            current.qvw_files.append(str(qvw_file).strip())
        current.row_count = row_idx - current.first_row + 1

    if current is not None:
        groups.append(current)

    log.info("Loaded %d dashboard groups from mapping sheet.", len(groups))
    if periodicity_header_text:
        log.info("Detected manually-maintained periodicity column: '%s' (will be preserved unchanged).",
                  periodicity_header_text)
    else:
        log.info("No periodicity column found in the mapping sheet — nothing to preserve.")
    if col_qmc is None:
        log.warning("No 'QMC Timings' column found in the mapping sheet — schedule/refresh-window "
                    "validation cannot run, and dashboards scheduled today will be flagged "
                    "'Schedule conversion error'.")
    return groups, periodicity_header_text


def read_dashboard_timings(ws_timing: Worksheet) -> dict[str, Optional[datetime]]:
    """Reads {Dashboard Name -> Last Updated datetime (IST)} from the timing sheet."""
    out = {}
    for row in ws_timing.iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        name = str(row[1]).strip()
        out[name.lower()] = parse_dashboard_timestamp(row[3] if len(row) > 3 else None)
    return out


# =============================================================================
# SECTION 5b — "Db updates" SHIFT-TIMING SHEET (ported from the standalone
# QlikView AccessPoint Scraper v4's fill_status() logic)
# =============================================================================
# The "Db updates" sheet is laid out as three side-by-side shift blocks:
#   Morning   -> columns A (Name) B (Trigger Time) C (Present Dashboard time) D (Status)
#   Afternoon -> columns F G H I                     (same 4 roles)
#   Night     -> columns K L M N                      (same 4 roles)
# with column E/J left blank as visual spacers, a merged title on row 1,
# headers on row 2, and data starting row 3. For every listed dashboard we
# look up its actual "Last Updated" timestamp (already scraped into the
# 'Dashboards timing' sheet by this same pipeline run) and classify it
# against that row's own Trigger Time:
#   "Updated"        -> within EARLY_LATE_THRESHOLD of the trigger time
#   "Updated (Early)"/"Updated (Late)" -> updated, but >2h off from trigger
#   "Not updated"     -> no update in the relevant window
#   "N/A"             -> row is restricted to a specific weekday via
#                        "(only <Weekday> trigger)" in the name, and today
#                        isn't that weekday
#   "No match"        -> the dashboard name couldn't be matched to any
#                        scraped "Dashboards timing" row at all

SHEET_DB_UPDATES = "DB Timings"
DB_UPDATES_DATA_START_ROW = 3

DB_UPDATES_SHIFT_BLOCKS = [
    # (name_col, trigger_col, present_time_col, status_col, shift_key)
    ("A", "B", "C", "D", "morning"),
    ("F", "G", "H", "I", "afternoon"),
    ("K", "L", "M", "N", "night"),
]

DB_UPDATED_FILL     = PatternFill("solid", fgColor="C6E0B4")   # light green — on-time
DB_NOT_UPDATED_FILL = PatternFill("solid", fgColor="FF0000")   # red
DB_NA_FILL          = PatternFill("solid", fgColor="FFF2CC")   # light yellow
DB_EARLY_FILL       = PatternFill("solid", fgColor="ADD8E6")   # blue — updated >2h early
DB_LATE_FILL        = PatternFill("solid", fgColor="006400")   # dark green — updated >2h late
DB_NOT_UPDATED_FONT = Font(color="FFFFFF", bold=True)
DB_LATE_FONT        = Font(color="FFFFFF", bold=True)

DB_EARLY_LATE_THRESHOLD = timedelta(hours=2)  # flag only if off by more than this

DB_WEEKDAY_MAP = {
    "mon": 0, "monday": 0, "tue": 1, "tues": 1, "tuesday": 1,
    "wed": 2, "wednesday": 2, "thu": 3, "thur": 3, "thurs": 3, "thursday": 3,
    "fri": 4, "friday": 4, "sat": 5, "saturday": 5, "sun": 6, "sunday": 6,
}

_DB_ONLY_DAY_RE  = re.compile(r"\(\s*only\s+([a-zA-Z]+)\s+trigger\s*\)", re.IGNORECASE)
_DB_PAREN_RE     = re.compile(r"\([^)]*\)")
_DB_QVW_PAREN_RE = re.compile(r"\(\s*([^()]*?\.qvw)\s*\)", re.IGNORECASE)


def _db_normalize_name(name: str) -> str:
    if not name:
        return ""
    name = _DB_PAREN_RE.sub("", name)
    name = name.strip().lower()
    name = re.sub(r"\s+", " ", name)
    return name


def _db_extract_weekday_restriction(raw_name: str):
    m = _DB_ONLY_DAY_RE.search(raw_name)
    if not m:
        return raw_name, None
    day_token = m.group(1).strip().lower()
    weekday = DB_WEEKDAY_MAP.get(day_token)
    clean_name = _DB_ONLY_DAY_RE.sub("", raw_name).strip()
    return clean_name, weekday


def _db_resolve_actual_dashboard_name(raw_name: str) -> str:
    """The 'Db updates' sheet often lists a QMC-style job name with the
    actual published AccessPoint dashboard name in trailing parens, e.g.
    'NTF Turbo Labseal.qvw (NTF Application.QVW)' — the real dashboard to
    look up is 'NTF Application.QVW'. Falls back to the raw name if there's
    no such trailing '(...qvw)' group."""
    m = _DB_QVW_PAREN_RE.search(raw_name)
    if m:
        return m.group(1).strip()
    return raw_name


def _db_parse_trigger_time(value) -> Optional[dtime]:
    """Accepts a datetime.time, datetime.datetime, or a messy string like
    '7:30 AM IST' / '16:30 PM IST' / '22:30:00 PM' and returns a
    datetime.time. Used only to flag early/late deviation."""
    if value is None:
        return None
    if isinstance(value, dtime):
        return value
    if isinstance(value, datetime):
        return value.time()

    s = str(value).strip()
    m = re.search(r"(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?", s)
    if not m:
        return None
    hour = int(m.group(1))
    minute = int(m.group(2))
    second = int(m.group(3) or 0)
    ampm = (m.group(4) or "").upper()

    if hour <= 12 and ampm:
        if ampm == "AM":
            hour = 0 if hour == 12 else hour
        elif ampm == "PM":
            hour = hour if hour == 12 else hour + 12
    hour = hour % 24
    return dtime(hour=hour, minute=minute, second=second)


def _db_build_last_updated_map(ws_timing: Worksheet) -> dict[str, tuple[str, Optional[datetime]]]:
    """Builds {normalized_dashboard_name: (original_name, last_updated_IST)}
    straight from this same run's 'Dashboards timing' sheet (S.No | Dashboard
    Name | Category | Last Updated) — i.e. whatever this pipeline run just
    scraped (--mode live) or loaded (--mode offline), so 'Db updates' always
    reflects the same data as the rest of the workbook."""
    mapping: dict[str, tuple[str, Optional[datetime]]] = {}
    for row in ws_timing.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 2 or not row[1]:
            continue
        name = str(row[1]).strip()
        last_updated = parse_dashboard_timestamp(row[3] if len(row) > 3 else None)
        mapping[_db_normalize_name(name)] = (name, last_updated)
    return mapping


def _db_find_match(norm_name: str, last_updated_map: dict, cutoff: float = 0.82):
    if norm_name in last_updated_map:
        return last_updated_map[norm_name], "exact"
    candidates = difflib.get_close_matches(norm_name, last_updated_map.keys(), n=1, cutoff=cutoff)
    if candidates:
        return last_updated_map[candidates[0]], "fuzzy"
    return None, None


def _db_get_shift_window(shift_key: str, now: datetime):
    """Returns (window_start, window_end, reference_date_for_weekday_check)."""
    today = now.date()
    if shift_key == "morning":
        return datetime.combine(today, dtime(7, 0)), datetime.combine(today, dtime(16, 0)), today
    if shift_key == "afternoon":
        return datetime.combine(today, dtime(16, 0)), datetime.combine(today, dtime(22, 0)), today
    if shift_key == "night":
        start_date = today - timedelta(days=1) if now.time() < dtime(7, 0) else today
        return (datetime.combine(start_date, dtime(22, 0)),
                datetime.combine(start_date + timedelta(days=1), dtime(7, 0)),
                start_date)
    raise ValueError(f"Unknown shift_key: {shift_key}")


def populate_db_updates_sheet(ws_db: Worksheet, ws_timing: Worksheet,
                               now: Optional[datetime] = None) -> dict:
    """Fills 'Present Dashboard time' + 'Status' for every row of every shift
    block in the 'Db updates' sheet, classifying each dashboard's scraped
    Last Updated timestamp against that row's own Trigger Time. Returns a
    stats dict; also logs a summary and any per-row match warnings."""
    now = now or datetime.now(IST_TZ).replace(tzinfo=None)
    last_updated_map = _db_build_last_updated_map(ws_timing)

    warnings = []
    stats = {"updated": 0, "early": 0, "late": 0, "not_updated": 0, "na": 0, "unmatched": 0}

    for name_col, trigger_col, present_col, status_col, shift_key in DB_UPDATES_SHIFT_BLOCKS:
        window_start, window_end, ref_date = _db_get_shift_window(shift_key, now)
        weekday_for_check = ref_date.weekday()

        row = DB_UPDATES_DATA_START_ROW
        while True:
            name_cell = ws_db[f"{name_col}{row}"]
            if name_cell.value is None:
                # Tolerate a single blank row before treating the block as done.
                next_name = ws_db[f"{name_col}{row + 1}"].value
                if next_name is None:
                    break
                row += 1
                continue

            raw_name = str(name_cell.value)
            present_cell = ws_db[f"{present_col}{row}"]
            status_cell = ws_db[f"{status_col}{row}"]

            clean_name, weekday_restriction = _db_extract_weekday_restriction(raw_name)

            if weekday_restriction is not None and weekday_restriction != weekday_for_check:
                status_cell.value, status_cell.fill = "N/A", DB_NA_FILL
                present_cell.value = ""
                stats["na"] += 1
                row += 1
                continue

            actual_name = _db_resolve_actual_dashboard_name(clean_name)
            norm_name = _db_normalize_name(actual_name)
            match, match_type = _db_find_match(norm_name, last_updated_map)

            if match is None:
                warnings.append(f"Row {row} ({name_col}, {shift_key}): no match found for '{clean_name}'")
                present_cell.value = "No data"
                status_cell.value, status_cell.fill = "No match", DB_NA_FILL
                stats["unmatched"] += 1
                row += 1
                continue

            matched_display_name, last_updated_dt = match
            if match_type == "fuzzy":
                warnings.append(f"Row {row} ({name_col}, {shift_key}): fuzzy-matched "
                                 f"'{clean_name}' -> '{matched_display_name}'")

            present_cell.value = last_updated_dt.strftime("%Y-%m-%d %H:%M") if last_updated_dt else "No data"

            trigger_time = _db_parse_trigger_time(ws_db[f"{trigger_col}{row}"].value)
            expected_dt = None
            if trigger_time is not None:
                if shift_key == "night":
                    # Night shift crosses midnight: trigger hours >= 12 (e.g.
                    # 10:30 PM) belong to the shift's start date; hours < 12
                    # (e.g. 1:30 AM) belong to the next day.
                    trigger_date = ref_date if trigger_time.hour >= 12 else ref_date + timedelta(days=1)
                else:
                    trigger_date = ref_date
                expected_dt = datetime.combine(trigger_date, trigger_time)

            if expected_dt is None:
                # Can't parse this row's trigger — fall back to the broad
                # shift-window check as a safety net (no early/late tagging).
                if last_updated_dt is not None and window_start <= last_updated_dt < window_end:
                    status_cell.value, status_cell.fill = "Updated", DB_UPDATED_FILL
                    stats["updated"] += 1
                else:
                    status_cell.value, status_cell.fill, status_cell.font = \
                        "Not updated", DB_NOT_UPDATED_FILL, DB_NOT_UPDATED_FONT
                    stats["not_updated"] += 1
                row += 1
                continue

            if last_updated_dt is None or last_updated_dt >= window_end:
                status_cell.value, status_cell.fill, status_cell.font = \
                    "Not updated", DB_NOT_UPDATED_FILL, DB_NOT_UPDATED_FONT
                stats["not_updated"] += 1
                row += 1
                continue

            lower_bound = expected_dt - DB_EARLY_LATE_THRESHOLD
            if last_updated_dt < lower_bound:
                status_cell.value, status_cell.fill, status_cell.font = \
                    "Not updated", DB_NOT_UPDATED_FILL, DB_NOT_UPDATED_FONT
                stats["not_updated"] += 1
            elif last_updated_dt < expected_dt:
                status_cell.value, status_cell.fill = "Updated (Early)", DB_EARLY_FILL
                stats["updated"] += 1
                stats["early"] += 1
            elif last_updated_dt <= expected_dt + DB_EARLY_LATE_THRESHOLD:
                status_cell.value, status_cell.fill = "Updated", DB_UPDATED_FILL
                stats["updated"] += 1
            else:
                status_cell.value, status_cell.fill, status_cell.font = \
                    "Updated (Late)", DB_LATE_FILL, DB_LATE_FONT
                stats["updated"] += 1
                stats["late"] += 1

            row += 1

    log.info("'%s' sheet: Updated (on-time)=%d Early=%d Late=%d Not updated=%d N/A=%d Unmatched=%d",
              SHEET_DB_UPDATES,
              stats["updated"] - stats["early"] - stats["late"],
              stats["early"], stats["late"], stats["not_updated"], stats["na"], stats["unmatched"])
    for w in warnings:
        log.warning("%s", w)

    return stats


# =============================================================================
# SECTION 6 — VALIDATION / STATUS ENGINE
# =============================================================================

@dataclass
class QvwResult:
    file_name: str
    matched: bool
    status: str
    completion_est: Optional[datetime]
    completion_ist: Optional[datetime]


@dataclass
class DashboardResult:
    dashboard_name: str
    refresh_ist: Optional[datetime]
    latest_qvw_completion_ist: Optional[datetime]
    qvw_results: list[QvwResult]
    status: str
    reason: str
    business_date: Optional[date]


_DAILY_PERIODICITY_TEXTS = {"everyday", "every day", "daily", "every-day"}


def _is_daily_periodicity(periodicity_text: Optional[object]) -> bool:
    """True if the manually-maintained Periodicity value means 'expected to
    refresh every day' (e.g. 'Everyday'). Free-text schedules tied to
    specific dates (monthly, 'Never', etc.) are NOT daily — those dashboards
    are only expected to refresh on their own scheduled days, so comparing
    them against *today* would wrongly flag correct, up-to-date dashboards
    as stale on the days they're not scheduled to run."""
    if not periodicity_text:
        return False
    return str(periodicity_text).strip().lower() in _DAILY_PERIODICITY_TEXTS


def evaluate_dashboard(
    group: DashboardGroup,
    refresh_ist: Optional[datetime],
    job_index: list[JobRecord],
    today_ist: Optional[date] = None,
) -> DashboardResult:
    today_ist = today_ist or datetime.now(IST_TZ).replace(tzinfo=None).date()
    qvw_results: list[QvwResult] = []
    for qvw_file in group.qvw_files:
        job = match_job(qvw_file, group.dashboard_name, job_index)
        if job is None:
            qvw_results.append(QvwResult(qvw_file, False, "Not Found", None, None))
            continue
        completion_ist = est_to_ist(job.last_execution)
        qvw_results.append(QvwResult(qvw_file, True, job.status or "Unknown",
                                      job.last_execution, completion_ist))

    completions = [r.completion_ist for r in qvw_results if r.completion_ist is not None]
    latest_completion = max(completions) if completions else None

    statuses_lower = {r.status.strip().lower() for r in qvw_results}

    # ---- Dashboard refresh missing -----------------------------------------
    if refresh_ist is None:
        reason = "No dashboard refresh timestamp found in AccessPoint scrape."
        return DashboardResult(group.dashboard_name, None, latest_completion,
                                qvw_results, "Dashboard Refresh Missing", reason,
                                latest_completion.date() if latest_completion else None)

    # ---- No mapped jobs at all ----------------------------------------------
    if not qvw_results:
        reason = "No QVW jobs mapped to this dashboard — cannot validate."
        return DashboardResult(group.dashboard_name, refresh_ist, None, qvw_results,
                                "Dashboard Refresh Missing", reason, refresh_ist.date())

    # ---- Any job failed / aborted, or disabled (can't ever refresh) -----------
    failed = [r for r in qvw_results
              if r.status.strip().lower() in FAILURE_STATUSES | DISABLED_STATUSES]
    if failed:
        names = ", ".join(f"{r.file_name} [{r.status}]" for r in failed)
        reason = f"{len(failed)} of {len(qvw_results)} QVW job(s) failed/aborted/disabled: {names}."
        return DashboardResult(group.dashboard_name, refresh_ist, latest_completion,
                                qvw_results, "QMC Failed", reason,
                                latest_completion.date() if latest_completion else refresh_ist.date())

    # ---- Any job currently running, or idle but never actually completed ------
    # ("Waiting"/"Success" only count once they have a real Last Execution
    # timestamp — "Waiting" + "Never" means it has not completed even once.)
    not_success = [
        r for r in qvw_results
        if r.status.strip().lower() in RUNNING_STATUSES
        or r.completion_ist is None
    ]
    if not_success:
        names = ", ".join(f"{r.file_name} [{r.status}]" for r in not_success)
        reason = f"{len(not_success)} of {len(qvw_results)} QVW job(s) not yet successfully completed: {names}."
        return DashboardResult(group.dashboard_name, refresh_ist, latest_completion,
                                qvw_results, "QMC Running", reason,
                                latest_completion.date() if latest_completion else refresh_ist.date())

    # ---- All jobs succeeded — check timing & business date --------------------
    if latest_completion is None:
        reason = "All QVW jobs report Success but no completion timestamp could be parsed."
        return DashboardResult(group.dashboard_name, refresh_ist, None, qvw_results,
                                "Dashboard Pending Refresh", reason, refresh_ist.date())

    same_business_date = refresh_ist.date() == latest_completion.date()
    time_gap = latest_completion - refresh_ist  # positive => dashboard timestamp is earlier

    # Dashboard timestamp may legitimately land a little before the QMC
    # "Last Execution" timestamp for the very same reload (scrape/logging
    # jitter between the two systems) — allow a small tolerance window
    # instead of requiring an exact refresh_ist >= latest_completion match.
    within_tolerance = timedelta(0) <= time_gap <= TIMESTAMP_TOLERANCE

    # ---- Daily dashboards must have actually refreshed TODAY ------------------
    # `same_business_date` above only checks that the dashboard and its QVW
    # jobs agree with EACH OTHER — it says nothing about whether that shared
    # date is today. A dashboard whose last successful reload was yesterday
    # (QMC job sitting idle in "Waiting" since then, not yet re-triggered
    # today) would otherwise pass every check above and be reported
    # "Updated" indefinitely, even though today's refresh hasn't happened
    # yet. For dashboards whose Periodicity says "Everyday", require the
    # latest completion to actually be dated today before calling it
    # Updated. Non-daily cadences (monthly/specific-date/"Never") are
    # exempt — they're only expected to refresh on their own scheduled
    # days, not every day.
    is_daily = _is_daily_periodicity(group.periodicity)
    stale_for_today = is_daily and latest_completion.date() < today_ist

    if same_business_date and not stale_for_today and (refresh_ist >= latest_completion or within_tolerance):
        if refresh_ist >= latest_completion:
            reason = (f"All {len(qvw_results)} QVW job(s) succeeded; dashboard refreshed "
                      f"{refresh_ist:%Y-%m-%d %H:%M} IST, at/after latest QVW completion "
                      f"{latest_completion:%Y-%m-%d %H:%M} IST; same business date.")
        else:
            gap_sec = int(time_gap.total_seconds())
            reason = (f"All {len(qvw_results)} QVW job(s) succeeded; dashboard refreshed "
                      f"{refresh_ist:%Y-%m-%d %H:%M} IST, within the {TIMESTAMP_TOLERANCE_MINUTES}-minute "
                      f"tolerance of latest QVW completion {latest_completion:%Y-%m-%d %H:%M} IST "
                      f"(gap {gap_sec}s); same business date.")
        return DashboardResult(group.dashboard_name, refresh_ist, latest_completion,
                                qvw_results, "Updated", reason, refresh_ist.date())

    reason_bits = []
    if stale_for_today:
        reason_bits.append(
            f"Periodicity is 'Everyday' but the latest successful QVW completion "
            f"({latest_completion:%Y-%m-%d %H:%M} IST) is dated {latest_completion.date()}, "
            f"not today ({today_ist}) — today's refresh hasn't happened yet")
    if refresh_ist < latest_completion and not within_tolerance:
        reason_bits.append(
            f"dashboard refresh ({refresh_ist:%Y-%m-%d %H:%M} IST) is more than "
            f"{TIMESTAMP_TOLERANCE_MINUTES} min older than latest QVW completion "
            f"({latest_completion:%Y-%m-%d %H:%M} IST)")
    if not same_business_date:
        reason_bits.append(
            f"business dates differ (dashboard {refresh_ist.date()} vs QVW {latest_completion.date()})")
    reason = "All QVW jobs succeeded, but " + "; ".join(reason_bits) + "."
    return DashboardResult(group.dashboard_name, refresh_ist, latest_completion,
                            qvw_results, "Dashboard Pending Refresh", reason,
                            latest_completion.date())


# =============================================================================
# SECTION 6b — SCHEDULE / PERIODICITY / REFRESH-WINDOW VALIDATION + ALERTS
# =============================================================================
# Pipeline per dashboard:
#   QMC Timings (ET) --parse--> clock times --DST-aware--> IST runs for today
#   Periodicity      --------> is the dashboard expected to run today?
#   Refresh (IST)    --------> inside [run - 3h, run + 3h] of the latest due run?
#   + QMC job status ("QMC Failed") from evaluate_dashboard()
#   ==> zero or one alert row per dashboard for the Dashboard_Alerts_PA sheet.

# ---- Periodicity ---------------------------------------------------------------
_DAY_TOKEN = (r"(?:monday|tuesday|wednesday|thursday|friday|saturday|sunday|"
              r"mon|tues?|wed|thur?s?|fri|sat|sun)")
_DAY_RANGE_RE = re.compile(rf"\b({_DAY_TOKEN})\b\s*(?:to|through|thru|-|–)\s*\b({_DAY_TOKEN})\b", re.IGNORECASE)
_DAY_RE = re.compile(rf"\b({_DAY_TOKEN})s?\b", re.IGNORECASE)
_NEVER_RE = re.compile(r"\b(never|none|disabled|not\s+scheduled|manual(?:ly)?|n/?a)\b", re.IGNORECASE)
_DAILY_RE = re.compile(r"\b(every\s*day|everyday|daily)\b", re.IGNORECASE)
_WEEKDAYS_RE = re.compile(r"\b(week\s?days?|business\s+days?)\b", re.IGNORECASE)
_WEEKENDS_RE = re.compile(r"\bweek\s?ends?\b", re.IGNORECASE)
_ORDINAL_RE = re.compile(r"\b(\d{1,2})(?:st|nd|rd|th)\b", re.IGNORECASE)
_LAST_DAY_RE = re.compile(r"\blast\s+day\b", re.IGNORECASE)
_WEEKDAY_INDEX = {"mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5, "sun": 6}


def _weekday_no(token: str) -> int:
    return _WEEKDAY_INDEX[token.lower()[:3]]


def is_expected_on(periodicity, day: date) -> Optional[bool]:
    """Should a dashboard with this free-text Periodicity run on `day`?

    True  -> expected on `day`         False -> not expected on `day`
    None  -> can't tell (blank, or text this parser doesn't understand:
             'Weekly Once' with no weekday, 'Monthly' with no dates, ...).
             Such dashboards are NOT validated; they are reported in the log.

    Understands: Everyday/Daily, Never/None/Disabled/Job Disabled/Not
    Scheduled/Manual, Monday to Friday / weekdays / weekends, named weekdays
    ("Weekly once on saturday's", "Mon, Wed, Fri"), and day-of-month lists
    ("At 3:30 PM on the 3rd, 7th and 12th of every month", "last day")."""
    if periodicity is None:
        return None
    text = re.sub(r"\s+", " ", str(periodicity).replace("\xa0", " ")).strip()
    if not text:
        return None
    if _NEVER_RE.search(text):
        return False
    if _DAILY_RE.search(text):
        return True

    weekdays: set[int] = set()
    if _WEEKDAYS_RE.search(text):
        weekdays |= {0, 1, 2, 3, 4}
    if _WEEKENDS_RE.search(text):
        weekdays |= {5, 6}
    for m in _DAY_RANGE_RE.finditer(text):           # "Monday to Friday" (cyclic)
        i, end = _weekday_no(m.group(1)), _weekday_no(m.group(2))
        while True:
            weekdays.add(i)
            if i == end:
                break
            i = (i + 1) % 7
    weekdays |= {_weekday_no(m.group(1)) for m in _DAY_RE.finditer(text)}

    month_days = {int(m.group(1)) for m in _ORDINAL_RE.finditer(text)}
    last_day = bool(_LAST_DAY_RE.search(text))

    if weekdays and (month_days or last_day):
        return None                                   # e.g. "every 2nd Monday" - ambiguous
    if weekdays:
        return day.weekday() in weekdays
    if month_days or last_day:
        return (day.day in month_days) or (last_day and day.day == calendar.monthrange(day.year, day.month)[1])
    return None


# ---- Schedule (QMC Timings -> IST for the run date) ---------------------------------------
@dataclass
class ScheduleInfo:
    display: str = ""                                     # text for the "Schedule" column
    runs: list[ExpectedRun] = field(default_factory=list)  # runs expected on the evaluation IST day
    expected_today: Optional[bool] = None                 # None = periodicity undetermined
    error: Optional[str] = None                           # set when the schedule couldn't be derived


def compute_schedule_info(qmc_timings, periodicity, ist_day: date) -> ScheduleInfo:
    """Converts one dashboard's QMC Timings to IST for `ist_day`, applies its
    Periodicity, and never raises: any problem is captured in `.error`."""
    info = ScheduleInfo()
    slots: list[dtime] = []
    try:
        slots = parse_qmc_timings(qmc_timings)
    except ScheduleConversionError as e:
        info.error, info.display = str(e), "Conversion error"
    except Exception as e:  # defensive: one odd cell must not abort the whole run
        log.exception("Unexpected error parsing QMC Timings %r", qmc_timings)
        info.error, info.display = f"{type(e).__name__}: {e}", "Conversion error"

    converted = slots_for_ist_day(slots, ist_day) if slots else []
    if converted:
        info.display = ", ".join(_fmt_12h(r.ist_dt.time()) for r in converted)

    if converted:
        # Periodicity (weekday / day-of-month) is defined in the QMC server's own
        # clock, so it is judged on each trigger's EASTERN date, not the IST date.
        verdicts = [is_expected_on(periodicity, r.et_dt.date()) for r in converted]
        if any(v is None for v in verdicts):
            info.expected_today = None
        else:
            info.runs = [r for r, v in zip(converted, verdicts) if v]
            info.expected_today = bool(info.runs)
    else:
        info.expected_today = is_expected_on(periodicity, ist_day)
        if info.expected_today and info.error is None:
            info.error = (f"QMC Timings has no schedulable time (value: {qmc_timings!r}) although "
                          f"Periodicity says this dashboard runs today - cannot derive the IST schedule")
    return info


# ---- ±3h refresh window --------------------------------------------------------------
@dataclass
class WindowCheck:
    verdict: str                       # "ok" | "pending" | "not_refreshed" | "delayed"
    slot: Optional[datetime] = None    # the scheduled run that was judged
    window: str = ""
    detail: str = ""


def _fmt_delta(td: timedelta) -> str:
    mins = int(td.total_seconds() // 60)
    d, rem = divmod(mins, 1440)
    h, m = divmod(rem, 60)
    return " ".join(p for p in (f"{d}d" if d else "", f"{h}h" if h else "", f"{m}m" if m or not (d or h) else "") if p)


def _fmt_window(lo: datetime, hi: datetime) -> str:
    if lo.date() == hi.date():
        return f"{lo:%H:%M}–{hi:%H:%M} IST"
    return f"{lo:%Y-%m-%d %H:%M} – {hi:%Y-%m-%d %H:%M} IST"


def check_refresh_window(run_times: list[datetime], refresh_ist: datetime, now_ist: datetime,
                         tol: timedelta = REFRESH_WINDOW_TOLERANCE) -> WindowCheck:
    """Judges the dashboard's latest refresh against today's scheduled runs.

    Only runs whose window has fully CLOSED (run + tol <= now) can be failed -
    before that the refresh may still be on its way ("pending", no alert). The
    most recent such run is the one judged:
      refresh inside [run-tol, run+tol]            -> ok
      refresh after run+tol                        -> delayed (unless it is simply
                                                      the NEXT run's refresh -> ok)
      refresh before run-tol (i.e. none for it)    -> not_refreshed
    We only hold the latest refresh timestamp, so an earlier run's refresh that
    was superseded by a later one is treated as fine, not as a miss."""
    due = [s for s in run_times if s + tol <= now_ist]
    if not due:
        return WindowCheck("pending")
    slot = max(due)
    lo, hi = slot - tol, slot + tol
    window = _fmt_window(lo, hi)
    where = f"expected {slot:%Y-%m-%d %H:%M} IST (window {window})"
    if lo <= refresh_ist <= hi:
        return WindowCheck("ok", slot, window)
    if refresh_ist > hi:
        if any(s - tol <= refresh_ist <= s + tol for s in run_times if s > slot):
            return WindowCheck("ok", slot, window)
        return WindowCheck("delayed", slot, window,
                           f"{where}; last refresh {refresh_ist:%Y-%m-%d %H:%M} IST is "
                           f"{_fmt_delta(refresh_ist - hi)} after the window closed")
    return WindowCheck("not_refreshed", slot, window,
                       f"{where}; last refresh {refresh_ist:%Y-%m-%d %H:%M} IST is "
                       f"{_fmt_delta(lo - refresh_ist)} before the window opened, so nothing refreshed for this run")


# ---- Alert generation ----------------------------------------------------------------
# Order = priority: the first category that applies becomes the row's Status/Severity;
# every other finding for that dashboard is appended to its Failure Reason.
ALERT_CATEGORIES = {
    "reload_failure":   {"status": "Qlik Reload Failure",            "headline": "Qlik reload failure",
                         "severity": "Critical", "emoji": "🔴"},
    "refresh_missing":  {"status": "Refresh Time Missing",           "headline": "Refresh time missing",
                         "severity": "High",     "emoji": "🟠"},
    "not_refreshed":    {"status": "Not Refreshed On Scheduled Day", "headline": "Dashboard not refreshed on scheduled day",
                         "severity": "High",     "emoji": "🟠"},
    "refresh_delayed":  {"status": "Refresh Delayed",                "headline": "Refresh delayed beyond threshold",
                         "severity": "Medium",   "emoji": "🟡"},
    "schedule_error":   {"status": "Schedule Conversion Error",      "headline": "Schedule conversion error",
                         "severity": "Medium",   "emoji": "🟡"},
    "validation_error": {"status": "Validation Error",               "headline": "Validation error",
                         "severity": "Medium",   "emoji": "🟡"},
}
_SEVERITY_ORDER = {"Critical": 0, "High": 1, "Medium": 2}
_UNASSIGNED_OWNER = "Unassigned – needs owner mapping"


@dataclass
class DashboardAlert:
    dashboard: str
    category: str
    status: str
    severity: str
    reason: str            # -> "Failure Reason"
    expected: str          # -> "Expected Refresh Time"
    last_refresh: str      # -> "Last Successful Refresh Time"
    business_date: str
    owner: str
    team: str
    application: str
    alert_key: str


def _alert_key(dashboard: str, category: str, ist_day: date) -> str:
    """Stable 10-hex id. Same dashboard + same problem + same IST day -> same
    key, so a 'Y' Power Automate already set survives the rebuild; a new day
    or a different problem produces a new key and is notified again."""
    return hashlib.sha1(f"{dashboard.strip().lower()}|{category}|{ist_day.isoformat()}".encode("utf-8")).hexdigest()[:10]


def _owner_and_team(bi_dna) -> tuple[str, str]:
    v = str(bi_dna or "").strip().lower()
    if v == "bi":
        return "BI Team", "BI"
    if v == "dna":
        return "DnA Team", "DnA"
    return _UNASSIGNED_OWNER, "Unassigned"


def _expected_text(info: ScheduleInfo, window_check: Optional[WindowCheck], ist_day: date) -> str:
    if window_check is not None and window_check.slot is not None:
        return f"{window_check.slot:%Y-%m-%d %H:%M} IST (window {window_check.window})"
    if info.runs:
        return f"{ist_day:%Y-%m-%d} " + " / ".join(f"{r.ist_dt:%H:%M}" for r in info.runs) + " IST"
    if info.error:
        return "N/A – schedule could not be derived from QMC Timings"
    return "N/A – no QMC Timings"


def build_dashboard_alerts(
    groups: list[DashboardGroup], results: list[DashboardResult], schedules: list[ScheduleInfo],
    categories: dict[str, str], now_ist: datetime,
) -> tuple[list[DashboardAlert], dict]:
    """Applies the alert rules to every dashboard and returns (alerts, stats).

    A dashboard is validated only if its Periodicity says it runs today
    (a dashboard whose periodicity can't be interpreted is skipped and listed
    in the log). Findings collected per validated dashboard:
      * QMC job status 'QMC Failed'                    -> Qlik reload failure
      * no refresh timestamp                           -> Refresh time missing
      * latest closed run: refresh before its window   -> Dashboard not refreshed on scheduled day
      * latest closed run: refresh after its window    -> Refresh delayed beyond threshold
      * QMC Timings blank/unparseable, dashboard due   -> Schedule conversion error
      * unexpected exception while evaluating          -> Validation error
    One row per dashboard: the highest-priority finding sets Status/Severity."""
    ist_day = now_ist.date()
    stats = {"total": len(groups), "validated": 0, "not_scheduled_today": 0, "undetermined": [],
             "pending": 0, "schedule_errors": []}
    alerts: list[DashboardAlert] = []

    for group, result, info in zip(groups, results, schedules):
        name = group.dashboard_name
        try:
            if info.expected_today is None:
                stats["undetermined"].append(f"{name} [Periodicity={group.periodicity!r}]")
                continue
            failure_only = False
            if not info.expected_today:
                stats["not_scheduled_today"] += 1
                if not (ALERT_ON_FAILURE_WHEN_NOT_SCHEDULED and result.status == "QMC Failed"):
                    continue
                failure_only = True
            else:
                stats["validated"] += 1

            findings: list[tuple[str, str]] = []
            window_check: Optional[WindowCheck] = None

            if result.status == "QMC Failed":
                findings.append(("reload_failure", result.reason))
            if not failure_only:
                if info.error:
                    findings.append(("schedule_error", info.error))
                    stats["schedule_errors"].append(f"{name}: {info.error}")
                if result.refresh_ist is None:
                    findings.append(("refresh_missing", "No dashboard refresh timestamp found in AccessPoint scrape."))
                elif info.runs:
                    window_check = check_refresh_window([r.ist_dt for r in info.runs], result.refresh_ist, now_ist)
                    if window_check.verdict == "pending":
                        stats["pending"] += 1
                    elif window_check.verdict in ("not_refreshed", "delayed"):
                        detail = window_check.detail
                        if result.status == "QMC Running":   # useful context: a reload is queued/stuck
                            detail += f" (QMC: {result.reason})"
                        key = "not_refreshed" if window_check.verdict == "not_refreshed" else "refresh_delayed"
                        findings.append((key, detail))

            if not findings:
                continue
            findings.sort(key=lambda f: list(ALERT_CATEGORIES).index(f[0]))
            cat = findings[0][0]
            meta = ALERT_CATEGORIES[cat]
            reason = " | ".join(f"{ALERT_CATEGORIES[k]['headline']}: {d}" for k, d in findings)
            owner, team = _owner_and_team(group.bi_dna)
            alerts.append(DashboardAlert(
                dashboard=name, category=cat, status=meta["status"], severity=meta["severity"], reason=reason,
                expected=_expected_text(info, window_check, ist_day),
                last_refresh=(f"{result.refresh_ist:%Y-%m-%d %H:%M}" if result.refresh_ist
                              else "N/A – no refresh timestamp"),
                business_date=(result.business_date or ist_day).strftime("%Y-%m-%d"),
                owner=owner, team=team, application=categories.get(name.strip().lower(), "Unknown"),
                alert_key=_alert_key(name, cat, ist_day),
            ))
        except Exception as e:
            log.exception("Alert evaluation failed for dashboard '%s': %s", name, e)
            owner, team = _owner_and_team(group.bi_dna)
            meta = ALERT_CATEGORIES["validation_error"]
            alerts.append(DashboardAlert(
                dashboard=name, category="validation_error", status=meta["status"], severity=meta["severity"],
                reason=f"{meta['headline']}: could not evaluate this dashboard ({type(e).__name__}: {e})",
                expected="N/A – evaluation error", last_refresh="N/A", business_date=ist_day.strftime("%Y-%m-%d"),
                owner=owner, team=team, application=categories.get(name.strip().lower(), "Unknown"),
                alert_key=_alert_key(name, "validation_error", ist_day),
            ))

    alerts.sort(key=lambda a: _SEVERITY_ORDER[a.severity])   # stable: keeps sheet order within a severity
    return alerts, stats


def read_dashboard_categories(ws_timing: Worksheet) -> dict[str, str]:
    """{dashboard name (lowercase) -> Category} from 'Dashboards timing' - used as the
    'Application/Project' line of the Teams card."""
    out: dict[str, str] = {}
    for row in ws_timing.iter_rows(min_row=2, values_only=True):
        if row and len(row) > 2 and row[1]:
            out[str(row[1]).strip().lower()] = str(row[2] or "").strip() or "Unknown"
    return out


# =============================================================================
# SECTION 7 — WRITE RESULTS BACK TO Dashboard_status
# =============================================================================

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
STATUS_FILLS = {
    "Updated": PatternFill("solid", fgColor="C6EFCE"),
    "QMC Running": PatternFill("solid", fgColor="BDD7EE"),
    "QMC Failed": PatternFill("solid", fgColor="FFC7CE"),
    "Dashboard Pending Refresh": PatternFill("solid", fgColor="FFEB9C"),
    "Dashboard Refresh Missing": PatternFill("solid", fgColor="D9D9D9"),
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Base headers that are always present. "Sl.NO" and the periodicity column
# are inserted dynamically in _build_output_headers() only if they exist in
# the source sheet, and the high-priority alert column is always appended
# last. Column *positions* are therefore computed at write time, not fixed.
BASE_HEADERS = [
    "Dashboard Name", "QMC Timings", "Schedule", "Dashboard Refresh timings (IST)", ".QVW files",
    "QVW Refresh timings (IST)", "BI/DnA", "QVW Job Status", "Dashboard Status",
    "Validation Reason", "Business Date",
]
# Column widths keyed by header text (falls back to a default for unknown headers).
_DEFAULT_WIDTHS = {
    "Sl.NO": 7, "Dashboard Name": 42, "QMC Timings": 22, "Schedule": 22,
    "Dashboard Refresh timings (IST)": 20,
    ".QVW files": 55, "QVW Refresh timings (IST)": 20, "BI/DnA": 10,
    "QVW Job Status": 16, "Dashboard Status": 22, "Validation Reason": 55,
    "Business Date": 14,
}
_DEFAULT_WIDTH_FALLBACK = 20
# Header-name -> which "level" the column is (one value per dashboard block,
# vs one value per QVW job row).
_DASHBOARD_LEVEL_HEADERS = {
    "Sl.NO", "Dashboard Name", "QMC Timings", "Schedule", "Dashboard Refresh timings (IST)", "BI/DnA",
    "Dashboard Status", "Validation Reason", "Business Date",
}
_JOB_LEVEL_HEADERS = {".QVW files", "QVW Refresh timings (IST)", "QVW Job Status"}


def _build_output_headers(include_slno: bool, periodicity_header: Optional[str]) -> list[str]:
    headers = list(BASE_HEADERS)
    if include_slno:
        headers.insert(0, "Sl.NO")
    if periodicity_header:
        # Preserve its original position: right after "QVW Job Status",
        # matching where it currently sits in the sheet.
        insert_at = headers.index("QVW Job Status") + 1
        headers.insert(insert_at, periodicity_header)
    return headers


def write_results(wb: openpyxl.Workbook, results: list[DashboardResult],
                   groups: list[DashboardGroup], sheet_name: str = SHEET_MAPPING_AND_OUTPUT,
                   periodicity_header: Optional[str] = None,
                   include_slno: Optional[bool] = None,
                   schedules: Optional[list[str]] = None):
    """Overwrites the Dashboard_status sheet with fresh results. Each
    dashboard becomes a block of rows (one row per QVW job); dashboard-level
    columns (Name, Refresh time, BI/DnA, Status, Reason, Business Date, and
    the manually-maintained Periodicity column if present) are vertically
    merged and centered across the whole block; job-level columns (.QVW
    files, QVW Refresh timings, QVW Job Status) get one value per row.
    Detail rows are collapsible (Excel row grouping), matching the
    reference layout.

    `periodicity_header`: exact header text of the manually-maintained
    periodicity column (if the source sheet had one) — its values are
    carried over from `group.periodicity` UNCHANGED; this function never
    computes or edits them.

    "QMC Timings" is passthrough (raw value + number format, exactly as
    entered); "Schedule" is computed (`schedules[i]` = IST text for group i,
    see compute_schedule_info) and is blank if `schedules` isn't supplied.

    High-priority job alerts are NOT written into this sheet — they live
    only in the "Priority BI Jobs" sheet (see populate_priority_jobs_sheet).
    """
    if include_slno is None:
        include_slno = any(g.sl_no is not None for g in groups)

    output_headers = _build_output_headers(include_slno, periodicity_header)
    col_of = {h: i for i, h in enumerate(output_headers, 1)}
    dashboard_level_cols = {col_of[h] for h in output_headers if h in _DASHBOARD_LEVEL_HEADERS
                             or h == periodicity_header}
    n_cols = len(output_headers)

    # Re-create the sheet at the position it already had. (Previously always index 0,
    # which would push a sheet placed before it - e.g. Dashboard_Alerts_PA - to second.)
    sheet_index = wb.sheetnames.index(sheet_name) if sheet_name in wb.sheetnames else 0
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, sheet_index)
    # Detail rows collapse "up" into the dashboard's header row (which sits
    # above them), so the +/- toggle appears above the group, not below.
    ws.sheet_properties.outlinePr.summaryBelow = False

    for col, header in enumerate(output_headers, 1):
        c = ws.cell(row=1, column=col, value=header)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[1].height = 30

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    row_idx = 2
    for g_i, (group, result) in enumerate(zip(groups, results)):
        group_start = row_idx
        status_fill = STATUS_FILLS.get(result.status, PatternFill())
        job_rows = result.qvw_results if result.qvw_results else [None]  # at least 1 row per dashboard

        for i, qvw in enumerate(job_rows):
            if qvw is not None:
                ws.cell(row=row_idx, column=col_of[".QVW files"], value=qvw.file_name)
                ws.cell(row=row_idx, column=col_of["QVW Refresh timings (IST)"],
                        value=qvw.completion_ist.strftime("%Y-%m-%d %H:%M:%S") if qvw.completion_ist else "")
                ws.cell(row=row_idx, column=col_of["QVW Job Status"],
                        value=qvw.status if qvw.matched else "Not Found in QMC")
            if i > 0:
                # Detail rows beyond the first are what collapses/expands.
                ws.row_dimensions[row_idx].outlineLevel = 1
            row_idx += 1
        group_end = row_idx - 1

        if include_slno:
            ws.cell(row=group_start, column=col_of["Sl.NO"], value=group.sl_no)
        ws.cell(row=group_start, column=col_of["Dashboard Name"], value=group.dashboard_name)
        qmc_cell = ws.cell(row=group_start, column=col_of["QMC Timings"], value=group.qmc_timings)
        qmc_cell.number_format = group.qmc_timings_fmt or "General"   # keeps "7:15 AM" style for time cells
        ws.cell(row=group_start, column=col_of["Schedule"],
                value=(schedules[g_i] if schedules and g_i < len(schedules) else None) or None)
        ws.cell(row=group_start, column=col_of["Dashboard Refresh timings (IST)"],
                value=result.refresh_ist.strftime("%Y-%m-%d %H:%M:%S") if result.refresh_ist else "")
        ws.cell(row=group_start, column=col_of["BI/DnA"], value=group.bi_dna)
        if periodicity_header:
            # Passthrough only — never derived/overwritten by this pipeline.
            ws.cell(row=group_start, column=col_of[periodicity_header], value=group.periodicity)
        ws.cell(row=group_start, column=col_of["Dashboard Status"], value=result.status)
        ws.cell(row=group_start, column=col_of["Validation Reason"], value=result.reason)
        ws.cell(row=group_start, column=col_of["Business Date"],
                value=result.business_date.strftime("%Y-%m-%d") if result.business_date else "")

        # Merge the dashboard-level columns down across the whole block.
        if group_end > group_start:
            for col in dashboard_level_cols:
                ws.merge_cells(start_row=group_start, start_column=col,
                                end_row=group_end, end_column=col)

        for r in range(group_start, group_end + 1):
            for col in range(1, n_cols + 1):
                cell = ws.cell(row=r, column=col)
                cell.border = BORDER
                cell.font = Font(name="Arial", size=9)
                if col in dashboard_level_cols:
                    cell.alignment = center if col not in (col_of["Dashboard Name"], col_of["Validation Reason"]) else left_wrap
                else:
                    cell.alignment = Alignment(vertical="center", wrap_text=(col == col_of[".QVW files"]))

        top_left = ws.cell(row=group_start, column=col_of["Dashboard Status"])
        top_left.fill = status_fill
        top_left.font = Font(name="Arial", size=9, bold=True)

    for header, col in col_of.items():
        ws.column_dimensions[get_column_letter(col)].width = _DEFAULT_WIDTHS.get(header, _DEFAULT_WIDTH_FALLBACK)
    ws.freeze_panes = "A2"
    if row_idx > 2:
        ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{row_idx - 1}"

    log.info("Wrote %d dashboard result group(s) / %d total rows to '%s' (%d columns, periodicity=%s).",
              len(results), row_idx - 2, sheet_name, n_cols, bool(periodicity_header))


# =============================================================================
# SECTION 7a — WRITE Dashboard_status ROLL-UP INTO Summary
#              (BI vs DnA counts, QMC Failed watch-list, refresh-currency
#              check). Runs on EVERY invocation (offline and live), right
#              after write_results(), so the Summary tab never goes stale
#              even between live scrapes.
# =============================================================================

_SUMMARY_DASH_START_ROW = 12  # everything from here down belongs to this
                               # section; rows 1-10 (QMC Jobs Report) are
                               # left untouched (only written in live mode).

_CURRENCY_GREEN = PatternFill("solid", fgColor="C6EFCE")
_CURRENCY_YELLOW = PatternFill("solid", fgColor="FFEB9C")
_CURRENCY_RED = PatternFill("solid", fgColor="FFC7CE")
_CURRENCY_BLUE = PatternFill("solid", fgColor="BDD7EE")
_CURRENCY_GREY = PatternFill("solid", fgColor="D9D9D9")
_STALE_FONT_RED = Font(name="Arial", size=10, bold=True, color="C00000")
_MUTED_ITALIC = Font(name="Arial", size=9, italic=True, color="7F7F7F")


def _clear_from_row(ws: Worksheet, start_row: int):
    if ws.max_row >= start_row:
        ws.delete_rows(start_row, ws.max_row - start_row + 1)
    # also clear any conditional formatting rules from a previous run so
    # they don't pile up / point at stale ranges.
    ws.conditional_formatting._cf_rules.clear()


def write_dashboard_status_summary(wb: openpyxl.Workbook, results: list[DashboardResult],
                                    groups: list[DashboardGroup],
                                    sheet_name: str = "Summary",
                                    today_ist: Optional[date] = None):
    """Appends a live 'Dashboard_status Report' block to the Summary sheet:
      A) BI vs DnA dashboard counts
      B) Currently QMC Failed dashboards (sorted worst-first by staleness)
      C) Refresh-currency check for every dashboard: Business Date vs today,
         sorted worst-first, color-coded (green=today, yellow=1 day,
         red=2+ days).
    Pure Python values (not formulas) — consistent with how the rest of this
    pipeline writes Dashboard_status, and correct by construction since the
    whole block is rebuilt fresh on every run.
    """
    today_ist = today_ist or datetime.now(ZoneInfo("Asia/Kolkata")).date()

    if sheet_name not in wb.sheetnames:
        wb.create_sheet(sheet_name)
    ws = wb[sheet_name]
    _clear_from_row(ws, _SUMMARY_DASH_START_ROW)

    bi_dna_by_name = {g.dashboard_name: (g.bi_dna or "").strip() for g in groups}
    periodicity_by_name = {g.dashboard_name: ("" if g.periodicity is None else str(g.periodicity).strip())
                            for g in groups}

    def days_stale(bdate: Optional[date]) -> Optional[int]:
        return (today_ist - bdate).days if bdate else None

    rows_data = []
    for r in results:
        rows_data.append({
            "name": r.dashboard_name,
            "bi_dna": bi_dna_by_name.get(r.dashboard_name, ""),
            "periodicity": periodicity_by_name.get(r.dashboard_name, ""),
            "status": r.status,
            "reason": r.reason,
            "business_date": r.business_date,
            "days_stale": days_stale(r.business_date),
        })

    total = len(rows_data)
    bi_count = sum(1 for d in rows_data if d["bi_dna"].lower() == "bi")
    dna_count = sum(1 for d in rows_data if d["bi_dna"].lower() == "dna")
    unspecified = total - bi_count - dna_count

    failed = sorted(
        [d for d in rows_data if d["status"] == "QMC Failed"],
        key=lambda d: -(d["days_stale"] if d["days_stale"] is not None else -1),
    )
    all_sorted = sorted(
        rows_data,
        key=lambda d: -(d["days_stale"] if d["days_stale"] is not None else -1),
    )
    up_to_date = sum(1 for d in rows_data if d["days_stale"] == 0)
    not_up_to_date = total - up_to_date
    genuinely_stale = sum(1 for d in rows_data if (d["days_stale"] or 0) >= 2)

    row = _SUMMARY_DASH_START_ROW

    def title(text):
        nonlocal row
        c = ws.cell(row=row, column=1, value=text)
        c.font = Font(name="Arial", size=13, bold=True, color="1F3864")
        row += 2

    def note(text):
        nonlocal row
        ws.cell(row=row, column=1, value=text).font = _MUTED_ITALIC
        row += 2

    def label_value(label, value, number_format=None):
        nonlocal row
        a = ws.cell(row=row, column=1, value=label)
        a.font = Font(name="Arial", size=10, bold=True)
        b = ws.cell(row=row, column=2, value=value)
        b.font = Font(name="Arial", size=10, bold=True)
        if number_format:
            b.number_format = number_format
        row += 1

    def table_header(headers):
        nonlocal row
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=row, column=i, value=h)
            c.font = HEADER_FONT
            c.fill = HEADER_FILL
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER
        row += 1

    def currency_fill(days):
        if days is None:
            return _CURRENCY_GREY
        if days == 0:
            return _CURRENCY_GREEN
        if days == 1:
            return _CURRENCY_YELLOW
        return _CURRENCY_RED

    # --- Section A: BI vs DnA counts ---------------------------------------
    title("Dashboard_status Report")
    label_value("As Of:", today_ist.strftime("%Y-%m-%d"))
    label_value("Total Dashboards:", total)
    label_value("BI Dashboards:", bi_count)
    label_value("DnA Dashboards:", dna_count)
    label_value("Not Specified (BI/DnA blank):", unspecified)
    row += 1
    if unspecified:
        note(f"Note: {unspecified} dashboard(s) have no BI/DnA value set in Dashboard_status — "
             f"fill these in there so they're counted correctly.")

    # --- Section B: QMC Failed watch-list -----------------------------------
    title("QMC Failed Dashboards (Action Needed)")
    label_value("Total Currently QMC Failed:", len(failed))
    row += 1
    table_header(["Dashboard Name", "BI/DnA", "Periodicity", "Last Business Date", "Days Since Update", "Failure Reason"])
    for d in failed:
        ws.cell(row=row, column=1, value=d["name"])
        ws.cell(row=row, column=2, value=d["bi_dna"])
        ws.cell(row=row, column=3, value=d["periodicity"])
        bd_cell = ws.cell(row=row, column=4,
                           value=d["business_date"].strftime("%Y-%m-%d") if d["business_date"] else "")
        ws.cell(row=row, column=5, value=d["days_stale"] if d["days_stale"] is not None else "")
        ws.cell(row=row, column=6, value=d["reason"])
        for col in range(1, 7):
            c = ws.cell(row=row, column=col)
            c.font = Font(name="Arial", size=9)
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=(col == 6))
        ws.cell(row=row, column=5).fill = currency_fill(d["days_stale"])
        row += 1
    if not failed:
        ws.cell(row=row, column=1, value="None — no dashboards currently in QMC Failed status.").font = \
            Font(name="Arial", size=10, italic=True, color="006100")
        row += 1
    row += 1

    # --- Section C: refresh currency check ----------------------------------
    title("Refresh Currency Check (Business Date vs Today)")
    note("0 days = refreshed today. 1 day = normal for daily jobs still scheduled later today. "
         "2+ days = genuinely stale and needs follow-up.")
    label_value("Up To Date (Business Date = Today):", up_to_date)
    label_value("Not Up To Date (Business Date <> Today):", not_up_to_date)
    label_value("Genuinely Stale (2+ days old):", genuinely_stale)
    row += 1
    table_header(["Dashboard Name", "BI/DnA", "Periodicity", "Dashboard Status", "Last Business Date", "Days Since Update"])
    for d in all_sorted:
        c1 = ws.cell(row=row, column=1, value=d["name"])
        ws.cell(row=row, column=2, value=d["bi_dna"])
        ws.cell(row=row, column=3, value=d["periodicity"])
        ws.cell(row=row, column=4, value=d["status"])
        ws.cell(row=row, column=5,
                value=d["business_date"].strftime("%Y-%m-%d") if d["business_date"] else "")
        days_cell = ws.cell(row=row, column=6, value=d["days_stale"] if d["days_stale"] is not None else "")
        for col in range(1, 7):
            c = ws.cell(row=row, column=col)
            c.font = Font(name="Arial", size=9)
            c.border = BORDER
            c.alignment = Alignment(vertical="center", wrap_text=(col in (1, 4)))
        fill = currency_fill(d["days_stale"])
        c1.fill = fill
        days_cell.fill = fill
        if (d["days_stale"] or 0) >= 2:
            c1.font = _STALE_FONT_RED
        row += 1

    widths = {"A": 42, "B": 10, "C": 30, "D": 22, "E": 18, "F": 55}
    for col, w in widths.items():
        cur = ws.column_dimensions[col].width or 0
        if w > cur:
            ws.column_dimensions[col].width = w

    log.info("Wrote Dashboard_status Summary roll-up: %d dashboards (BI=%d, DnA=%d), "
              "%d QMC Failed, %d not up to date (%d genuinely stale >=2d).",
              total, bi_count, dna_count, len(failed), not_up_to_date, genuinely_stale)


# =============================================================================
# SECTION 7c — WRITE Dashboard_Alerts_PA  (Power Automate -> Teams trigger sheet)
# =============================================================================
# The sheet's title (rows 1-2), header row, column widths, merged title cells,
# freeze panes and column order are owned by the workbook and are NEVER touched.
# Only the data rows beneath the header are cleared and rewritten. Columns are
# located by header text, so re-ordering/inserting columns in the workbook
# doesn't break the writer.

ALERT_HEADERS = [
    "Alert ID", "Dashboard Name", "Last Successful Refresh Time", "Expected Refresh Time", "Business Date",
    "Status", "Alert Severity", "Failure Reason", "Owner", "Alert Message", "Teams Card Title",
    "Teams Card Description", "Alert Timestamp", "Alert Key", "Notified (Y/N)",
]
_ALERT_SEVERITY_FILLS = {"Critical": "F8CBCB", "High": "FCE4B5", "Medium": "FFF2CC"}   # as in the workbook
_ALERT_WRAP_HEADERS = {"failurereason", "alertmessage", "teamscarddescription"}
_ALERT_CHARS_PER_LINE, _ALERT_LINE_HEIGHT, _ALERT_MIN_HEIGHT = 55, 12.5, 87.5   # fits the existing rows


def _alert_row_height(*texts) -> float:
    lines = max(sum(max(1, -(-len(p) // _ALERT_CHARS_PER_LINE)) for p in str(t).split("\n")) for t in texts)
    return max(_ALERT_MIN_HEIGHT, _ALERT_LINE_HEIGHT * lines)


def _alert_values(a: DashboardAlert, alert_id: int, ts: str, notified: str) -> dict[str, object]:
    """One dict per row, keyed by NORMALIZED header text."""
    meta = ALERT_CATEGORIES[a.category]
    message = f"{a.dashboard} — {a.reason}"
    if a.category == "schedule_error":
        title = f"{meta['emoji']} {a.severity.upper()}: {a.dashboard} schedule could not be validated"
    else:
        title = f"{meta['emoji']} {a.severity.upper()}: {a.dashboard} not updated"
    description = (f"**Status:** {a.status}  \n**Application/Project:** {a.application}  \n**Team:** {a.team}  \n"
                   f"**Last Successful Refresh:** {a.last_refresh}  \n**Business Date:** {a.business_date}  \n"
                   f"**Reason:** {a.reason}")
    values = [alert_id, a.dashboard, a.last_refresh, a.expected, a.business_date, a.status, a.severity,
              a.reason, a.owner, message, title, description, ts, a.alert_key, notified]
    out = {_normalize_header(h): v for h, v in zip(ALERT_HEADERS, values)}
    out["applicationproject"] = a.application     # optional column: written only if the sheet has it
    return out


def write_alerts_sheet(wb: openpyxl.Workbook, alerts: list[DashboardAlert], now_ist: datetime,
                       sheet_name: str = SHEET_ALERTS) -> dict:
    """Clears the data rows of `Dashboard_Alerts_PA` and writes the latest alerts.

    * Header row / title / widths / freeze panes / merges: untouched.
    * Data-row formatting is inherited from the FIRST data row of the sheet
      (font, border, alignment, number format); only the Alert Severity fill
      varies per row. That first row is kept (empty but styled) even when there
      are 0 alerts, so formatting survives quiet runs and the next run.
    * Columns this script does not manage (e.g. helper/lookup formulas such as
      "Send Teams Alert?") are preserved: a formula found in the first data row
      is re-created, row-adjusted, on every alert row instead of being wiped.
    * `Notified (Y/N)` is 'N' for new alerts; if PRESERVE_NOTIFIED_FLAG and the
      same Alert Key was already 'Y' (Power Automate posted it), 'Y' is kept.
    * Any existing Excel Table on the sheet is resized; the AutoFilter range
      is re-fitted to the data.
    Raises KeyError (sheet left untouched) if an expected header is missing."""
    if sheet_name not in wb.sheetnames:
        log.error("Sheet '%s' not found in the workbook — alerts were NOT written. Add the sheet "
                  "(with its header row) and re-run.", sheet_name)
        return {"written": 0, "carried_forward": 0}
    ws = wb[sheet_name]

    header_row = next((r for r in range(1, 21) if _normalize_header(ws.cell(row=r, column=1).value) == "alertid"), None)
    if header_row is None:
        raise KeyError(f"'{sheet_name}': could not find the header row (a first-column 'Alert ID' cell in rows 1-20).")
    ncols = ws.max_column
    col_of = {_normalize_header(ws.cell(row=header_row, column=c).value): c for c in range(1, ncols + 1)
              if ws.cell(row=header_row, column=c).value not in (None, "")}
    missing = [h for h in ALERT_HEADERS if _normalize_header(h) not in col_of]
    if missing:
        raise KeyError(f"'{sheet_name}' is missing expected header(s): {missing}. Nothing was written.")

    first = header_row + 1
    old_last = max(ws.max_row, first)

    # -- 1. remember previously-notified alerts (before anything is cleared) ---------
    notified_before: dict[str, str] = {}
    key_col, flag_col = col_of["alertkey"], col_of["notifiedyn"]
    for r in range(first, old_last + 1):
        key = ws.cell(row=r, column=key_col).value
        if key:
            notified_before[str(key).strip()] = str(ws.cell(row=r, column=flag_col).value or "N").strip().upper()

    # -- 2. capture the existing data-row style (col -> StyleArray) -------------------
    template = {c: (copy(ws.cell(row=first, column=c)._style) if ws.cell(row=first, column=c).has_style else None)
                for c in range(1, ncols + 1)}
    sev_col, wrap_cols = col_of["alertseverity"], {col_of[h] for h in _ALERT_WRAP_HEADERS}
    thin = Side(style="thin", color="D9D9D9")

    def style_cell(cell, c):
        if template[c] is not None:
            cell._style = copy(template[c])
        else:   # sheet had no formatted data row: fall back to the workbook's look
            cell.font = Font(name="Arial", size=10, bold=(c == sev_col))
            cell.alignment = Alignment(vertical="top", wrap_text=(c in wrap_cols))
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    managed = {_normalize_header(h) for h in ALERT_HEADERS} | {"applicationproject"}
    col_norm = {c: n for n, c in col_of.items()}
    formula_cols = {c: ws.cell(row=first, column=c).value for c in range(1, ncols + 1)
                    if col_norm.get(c) not in managed
                    and isinstance(ws.cell(row=first, column=c).value, str)
                    and ws.cell(row=first, column=c).value.startswith("=")}

    # -- 3. clear old alert rows (values only on row `first`; surplus rows removed) ------
    for r in range(first, old_last + 1):
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).value = None
    new_last = first + max(len(alerts), 1) - 1
    if old_last > new_last:
        ws.delete_rows(new_last + 1, old_last - new_last)
    for r in range(new_last + 1, old_last + 1):
        if r in ws.row_dimensions:
            del ws.row_dimensions[r]

    # -- 4. write the fresh list ----------------------------------------------------------
    ts = now_ist.strftime("%Y-%m-%d %H:%M:%S")
    carried = 0
    for i, a in enumerate(alerts):
        r = first + i
        flag = "N"
        if PRESERVE_NOTIFIED_FLAG and notified_before.get(a.alert_key) == "Y":
            flag, carried = "Y", carried + 1
        vals = _alert_values(a, i + 1, ts, flag)
        for c in range(1, ncols + 1):
            style_cell(ws.cell(row=r, column=c), c)
        for norm, v in vals.items():
            if norm in col_of:
                ws.cell(row=r, column=col_of[norm]).value = v
        ws.cell(row=r, column=sev_col).fill = PatternFill("solid", fgColor=_ALERT_SEVERITY_FILLS[a.severity])
        ws.row_dimensions[r].height = _alert_row_height(vals["failurereason"], vals["alertmessage"],
                                                        vals["teamscarddescription"])
    if not alerts:   # keep one empty, formatted row so the style template survives
        for c in range(1, ncols + 1):
            style_cell(ws.cell(row=first, column=c), c)
        ws.cell(row=first, column=sev_col).fill = PatternFill(fill_type=None)

    for c, formula in formula_cols.items():
        letter = get_column_letter(c)
        for r in range(first, new_last + 1):
            ws.cell(row=r, column=c).value = Translator(formula, origin=f"{letter}{first}").translate_formula(f"{letter}{r}")

    # -- 5. re-fit filter / table ranges ----------------------------------------------------
    ref = f"A{header_row}:{get_column_letter(ncols)}{new_last}"
    if ws.tables:
        for tbl in ws.tables.values():
            tbl.ref = ref
            if tbl.autoFilter is not None:
                tbl.autoFilter.ref = ref
    else:
        ws.auto_filter.ref = ref

    log.info("'%s': wrote %d alert row(s) (%d Notified='Y' carried forward).", sheet_name, len(alerts), carried)
    return {"written": len(alerts), "carried_forward": carried}


# =============================================================================
# SECTION 7b — WRITE RAW SCRAPE SHEETS (Jobs / Summary / status sheets /
#              Dashboards timing) — merged from 2705.py + Dashboards_Dog.py.
#              Only used in --mode live, so re-running the pipeline reproduces
#              the full tab set (Dashboard_status, Jobs, Summary, Disabled,
#              Waiting, Running, Failed, Dashboards timing) in one file.
# =============================================================================

_JOB_HEADERS = ["Name", "Executed On", "Status", "Distribution Group", "Last Execution", "Started/Scheduled"]
_JOB_STATUS_COLORS = {
    "success": "C6EFCE", "failed": "FFC7CE", "waiting": "FFEB9C",
    "running": "BDD7EE", "aborted": "F4CCCC", "warning": "FCE4D6", "disabled": "E2EFDA",
}
_JOB_TAB_COLORS = {
    "Jobs": "1F3864", "Summary": "595959", "Disabled": "70AD47",
    "Waiting": "FFC000", "Running": "2E75B6", "Failed": "FF0000",
}
_JOB_STATUS_SHEETS = ["Disabled", "Waiting", "Running", "Failed"]


def _write_job_sheet(wb: openpyxl.Workbook, title: str, rows: list[list], tab_color: str = None):
    if title in wb.sheetnames:
        del wb[title]
    ws = wb.create_sheet(title)
    if tab_color:
        ws.sheet_properties.tabColor = tab_color

    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    alt_fill = PatternFill("solid", fgColor="F5F5F5")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    thin = Side(style="thin", color="CCCCCC")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    col_widths = [65, 18, 12, 20, 22, 22]

    for col_idx, h in enumerate(_JOB_HEADERS, 1):
        c = ws.cell(row=1, column=col_idx, value=h)
        c.font, c.fill = hdr_font, hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = bdr
    ws.row_dimensions[1].height = 28

    for row_idx, row_data in enumerate(rows, 2):
        base_fill = alt_fill if row_idx % 2 == 0 else white_fill
        for col_idx, value in enumerate(row_data, 1):
            c = ws.cell(row=row_idx, column=col_idx, value=value)
            c.font = Font(name="Arial", size=9)
            c.alignment = Alignment(vertical="center", wrap_text=(col_idx == 1))
            c.border = bdr
            c.fill = base_fill
        status_val = str(row_data[2] if len(row_data) > 2 else "").lower()
        for keyword, color in _JOB_STATUS_COLORS.items():
            if keyword in status_val:
                sc = ws.cell(row=row_idx, column=3)
                sc.fill = PatternFill("solid", fgColor=color)
                sc.font = Font(name="Arial", size=9, bold=True)
                break

    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = ws.dimensions

    count_cell = ws.cell(row=1, column=len(_JOB_HEADERS) + 1, value=f"Count: {len(rows)}")
    count_cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=9)
    count_cell.fill = hdr_fill
    count_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.column_dimensions[get_column_letter(len(_JOB_HEADERS) + 1)].width = 12
    return ws


def write_qmc_raw_sheets(wb: openpyxl.Workbook, jobs_raw: list[dict]):
    """Writes Jobs / Summary / Disabled / Waiting / Running / Failed sheets
    from the freshly-scraped QMC job list (mirrors 2705.py's output)."""
    df = pd.DataFrame(jobs_raw, columns=_JOB_HEADERS).fillna("")
    all_rows = df.values.tolist()

    _write_job_sheet(wb, "Jobs", all_rows, tab_color=_JOB_TAB_COLORS["Jobs"])

    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    ws_sum = wb.create_sheet("Summary")
    ws_sum.sheet_properties.tabColor = _JOB_TAB_COLORS["Summary"]
    ws_sum["A1"] = "QMC Jobs Report"
    ws_sum["A1"].font = Font(name="Arial", bold=True, size=16, color="1F3864")
    ws_sum.merge_cells("A1:D1")
    ws_sum["A3"], ws_sum["B3"] = "Generated:", time.strftime("%Y-%m-%d %H:%M:%S")
    ws_sum["A4"], ws_sum["B4"] = "Total Jobs:", len(df)
    ws_sum["A3"].font = ws_sum["A4"].font = Font(name="Arial", bold=True, size=10)
    ws_sum["B3"].font = Font(name="Arial", size=10)
    ws_sum["B4"].font = Font(name="Arial", size=10, bold=True)

    hdr_fill = PatternFill("solid", fgColor="1F3864")
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="CCCCCC")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws_sum["A6"], ws_sum["B6"], ws_sum["C6"] = "Status", "Count", "Sheet"
    for col in ["A6", "B6", "C6"]:
        ws_sum[col].font, ws_sum[col].fill = hdr_font, hdr_fill
        ws_sum[col].alignment = Alignment(horizontal="center", vertical="center")
        ws_sum[col].border = bdr

    status_fill_map = {k: PatternFill("solid", fgColor=v) for k, v in _JOB_STATUS_COLORS.items()}
    row_off = 7
    keyword_map = {"Disabled": "disabled", "Waiting": "waiting", "Running": "running", "Failed": "failed"}
    for status_name in _JOB_STATUS_SHEETS:
        keyword = keyword_map[status_name]
        count = len(df[df["Status"].str.lower().str.contains(keyword, na=False)])
        fill = status_fill_map.get(keyword, PatternFill("solid", fgColor="FFFFFF"))
        for col_idx, val in enumerate([status_name, count, f"-> {status_name} sheet"], 1):
            c = ws_sum.cell(row=row_off, column=col_idx, value=val)
            c.font = Font(name="Arial", size=10, bold=(count > 0))
            c.fill = fill
            c.border = bdr
            c.alignment = Alignment(horizontal="center", vertical="center")
        row_off += 1

    ws_sum.column_dimensions["A"].width = 18
    ws_sum.column_dimensions["B"].width = 10
    ws_sum.column_dimensions["C"].width = 20
    ws_sum.row_dimensions[1].height = 32

    for status_name in _JOB_STATUS_SHEETS:
        keyword = keyword_map[status_name]
        filtered_rows = df[df["Status"].str.lower().str.contains(keyword, na=False)].values.tolist()
        _write_job_sheet(wb, status_name, filtered_rows, tab_color=_JOB_TAB_COLORS[status_name])

    log.info("Wrote raw QMC sheets: Jobs (%d), Disabled/Waiting/Running/Failed, Summary.", len(all_rows))


def write_dashboard_timing_sheet(wb: openpyxl.Workbook, dashboards: list[dict],
                                  sheet_name: str = SHEET_DASHBOARD_TIMING):
    """Writes the S.No / Dashboard Name (hyperlinked) / Category / Last Updated
    sheet from the freshly-scraped AccessPoint app list (mirrors
    Dashboards_Dog.py's save_excel), preserving all other sheets."""
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alt_fill = PatternFill("solid", fgColor="D6E4F0")
    white_fill = PatternFill("solid", fgColor="FFFFFF")
    link_font = Font(name="Arial", color="0563C1", underline="single", size=10)
    norm_font = Font(name="Arial", size=10)
    c_align = Alignment(horizontal="center", vertical="center")
    l_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="BDD7EE")
    bdr = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = ["S.No", "Dashboard Name", "Category", "Last Updated"]
    col_widths = [8, 55, 30, 20]
    for col, (h, w) in enumerate(zip(headers, col_widths), 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font, c.fill, c.alignment, c.border = hdr_font, hdr_fill, hdr_align, bdr
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 30

    for i, app in enumerate(dashboards, 1):
        row = i + 1
        fill = alt_fill if i % 2 == 0 else white_fill
        c = ws.cell(row=row, column=1, value=i)
        c.font, c.fill, c.alignment, c.border = norm_font, fill, c_align, bdr
        c = ws.cell(row=row, column=2, value=app.get("Name", ""))
        if app.get("URL"):
            c.hyperlink = app["URL"]
        c.font, c.fill, c.alignment, c.border = link_font, fill, l_align, bdr
        c = ws.cell(row=row, column=3, value=app.get("Category", ""))
        c.font, c.fill, c.alignment, c.border = norm_font, fill, c_align, bdr
        c = ws.cell(row=row, column=4, value=app.get("Last Updated", ""))
        c.font, c.fill, c.alignment, c.border = norm_font, fill, c_align, bdr
        ws.row_dimensions[row].height = 22

    ws.freeze_panes = "A2"
    if dashboards:
        ws.auto_filter.ref = f"A1:D{len(dashboards) + 1}"
    log.info("Wrote '%s' sheet (%d dashboards).", sheet_name, len(dashboards))


# =============================================================================
# SECTION 8 — ORCHESTRATION
# =============================================================================

def run_offline(excel_path: str) -> tuple[dict, list[dict], list[DashboardGroup], Optional[str], Worksheet]:
    """Reads dashboard timings, jobs and mapping straight out of the workbook."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    if SHEET_DASHBOARD_TIMING not in wb.sheetnames:
        raise KeyError(f"Sheet '{SHEET_DASHBOARD_TIMING}' not found in {excel_path}")
    if SHEET_JOBS not in wb.sheetnames:
        raise KeyError(f"Sheet '{SHEET_JOBS}' not found in {excel_path}")
    if SHEET_MAPPING_AND_OUTPUT not in wb.sheetnames:
        raise KeyError(f"Sheet '{SHEET_MAPPING_AND_OUTPUT}' not found in {excel_path}")

    timings = read_dashboard_timings(wb[SHEET_DASHBOARD_TIMING])

    jobs_raw = []
    headers = [c.value for c in next(wb[SHEET_JOBS].iter_rows(min_row=1, max_row=1))]
    for row in wb[SHEET_JOBS].iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        jobs_raw.append(dict(zip(headers, row)))

    groups, periodicity_header = read_mapping(wb[SHEET_MAPPING_AND_OUTPUT])
    priority_ws = wb[SHEET_PRIORITY_JOBS] if SHEET_PRIORITY_JOBS in wb.sheetnames else None
    return timings, jobs_raw, groups, periodicity_header, priority_ws


def run_live() -> tuple[dict, list[dict], list[dict], list[dict]]:
    """Runs both Selenium scrapers live. Mapping must still come from an
    existing workbook, so this returns None for groups — caller merges it
    with a mapping loaded separately via --excel. Returns
    (timings, jobs_raw, dashboards_raw)."""
    dashboards = scrape_dashboards()
    timings = {}
    for d in dashboards:
        timings[d["Name"].strip().lower()] = parse_dashboard_timestamp(d.get("Last Updated"))
    jobs_raw = scrape_qmc_jobs()
    return timings, jobs_raw, dashboards


def main():
    parser = argparse.ArgumentParser(description="Dashboard vs QVW/QMC status validation pipeline.")
    parser.add_argument("--mode", choices=["offline", "live"], default="offline",
                         help="offline = reuse existing sheet data; live = run both Selenium scrapers first.")
    parser.add_argument("--excel", default=DEFAULT_EXCEL_PATH,
                         help="Path to the workbook containing the mapping sheet "
                              "(and, in offline mode, the timing/jobs sheets too).")
    parser.add_argument("--output", default=None,
                         help="Output path. Defaults to overwriting --excel in place.")
    parser.add_argument("--now", default=None, metavar="'YYYY-MM-DD HH:MM'",
                         help="Override the current time (IST) used for periodicity, DST offset and "
                              "refresh-window checks. For testing/back-fills; default = real current time.")
    args = parser.parse_args()

    output_path = args.output or args.excel
    if args.now:
        try:
            now_ist = datetime.strptime(args.now, "%Y-%m-%d %H:%M")
        except ValueError:
            parser.error("--now must look like '2026-09-15 21:30'")
    else:
        now_ist = datetime.now(IST_TZ).replace(tzinfo=None)
    _et_now = now_ist.replace(tzinfo=IST_TZ).astimezone(ET_TZ)
    _off = now_ist.replace(tzinfo=IST_TZ).utcoffset() - _et_now.utcoffset()
    log.info("Evaluation time: %s IST%s = %s %s -> IST = ET + %d:%02d on this date.",
             now_ist.strftime("%Y-%m-%d %H:%M (%A)"), " [--now override]" if args.now else "",
             _et_now.strftime("%Y-%m-%d %H:%M"), _et_now.tzname(),
             _off.seconds // 3600, _off.seconds % 3600 // 60)

    try:
        dashboards_raw = None
        priority_ws = None
        if args.mode == "offline":
            log.info("Running in OFFLINE mode against '%s'.", args.excel)
            timings, jobs_raw, groups, periodicity_header, priority_ws = run_offline(args.excel)
        else:
            log.info("Running in LIVE mode (Selenium scrapers).")
            timings, jobs_raw, dashboards_raw = run_live()
            # Mapping structure still comes from the workbook, even in live mode.
            wb_map = openpyxl.load_workbook(args.excel, data_only=True)
            groups, periodicity_header = read_mapping(wb_map[SHEET_MAPPING_AND_OUTPUT])
            if SHEET_PRIORITY_JOBS in wb_map.sheetnames:
                priority_ws = wb_map[SHEET_PRIORITY_JOBS]

        job_index = build_job_index(jobs_raw)
        log.info("Indexed %d QMC job records.", len(job_index))

        if not job_index:
            raise RuntimeError(
                "No QMC job data was loaded at all (0 rows) — every dashboard would come "
                "back as 'Not Found in QMC', which is almost certainly wrong rather than "
                "true. In --mode live this means the QMC scrape didn't actually reach the "
                "task list (check VPN/login/the 'StatusFilterDropDown' wait in the log). "
                "In --mode offline this means the 'Jobs' sheet in --excel is empty. "
                "Aborting instead of overwriting Dashboard_status with bad data."
            )
        if args.mode == "live" and not dashboards_raw:
            raise RuntimeError(
                "0 dashboards were scraped from the AccessPoint (Dashboards_Dog side) — "
                "every 'Dashboards timing' row and 'Dashboard Refresh timings (IST)' cell "
                "would come back blank, which is almost certainly wrong rather than true. "
                "Check 'dashboard_listarea_dump.html' (written next to this script) and the "
                "log lines around 'Fetched #listArea innerHTML' / 'Scraped %d dashboards' to "
                "see whether login actually reached the app list. Aborting instead of "
                "overwriting Dashboard_status/Dashboards timing with blank data."
            )
        if not timings:
            log.warning(
                "No dashboard refresh timestamps were loaded at all — every dashboard will "
                "come back as 'Dashboard Refresh Missing'. Check the AccessPoint scrape "
                "(live mode) or the 'Dashboards timing' sheet (offline mode)."
            )

        results = []
        for group in groups:
            refresh_ist = timings.get(group.dashboard_name.strip().lower())
            try:
                result = evaluate_dashboard(group, refresh_ist, job_index, today_ist=now_ist.date())
            except Exception as e:
                log.exception("Failed to evaluate dashboard '%s': %s", group.dashboard_name, e)
                result = DashboardResult(group.dashboard_name, refresh_ist, None, [],
                                          "Dashboard Refresh Missing", f"Internal error: {e}", None)
            results.append(result)

        status_counts = pd.Series([r.status for r in results]).value_counts().to_dict()
        log.info("Status summary: %s", status_counts)

        # QMC Timings (ET) -> IST "Schedule" + periodicity for today. compute_schedule_info
        # never raises; a bad cell just becomes a "Schedule conversion error" for that dashboard.
        schedules = [compute_schedule_info(g.qmc_timings, g.periodicity, now_ist.date()) for g in groups]
        # Circuit breaker: if NO dashboard has any QMC Timings the column is missing/unfilled - a set-up
        # problem, not 60+ dashboard problems. Log once and don't flood Teams with per-dashboard
        # "Schedule conversion error" alerts (failures / missing refresh are still alerted).
        if groups and all(g.qmc_timings in (None, "") for g in groups):
            log.error("QMC Timings is empty for ALL dashboards (column missing or not filled in). Schedule "
                      "and refresh-window checks are DISABLED this run and 'Schedule conversion error' "
                      "alerts are suppressed. Fill in 'QMC Timings' on Dashboard_status.")
            for sch in schedules:
                sch.error = None

        # Still computed for its log line ("Checked N high-priority job(s); M
        # flagged") and to keep the monitoring logic exercised on every run —
        # just no longer written into the Dashboard_status sheet.
        if priority_ws is not None:
            build_priority_alerts(priority_ws, job_index)
        else:
            log.info("No '%s' sheet found — skipping high-priority job monitoring.", SHEET_PRIORITY_JOBS)

        wb_out = openpyxl.load_workbook(args.excel)
        if dashboards_raw is not None:
            write_dashboard_timing_sheet(wb_out, dashboards_raw)
            write_qmc_raw_sheets(wb_out, jobs_raw)
        write_results(wb_out, results, groups,
                      periodicity_header=periodicity_header,
                      schedules=[sch.display for sch in schedules])
        write_dashboard_status_summary(wb_out, results, groups)

        if SHEET_PRIORITY_JOBS in wb_out.sheetnames:
            populate_priority_jobs_sheet(wb_out[SHEET_PRIORITY_JOBS], job_index)
        else:
            log.info("No '%s' sheet found in output workbook — skipping auto-populate.", SHEET_PRIORITY_JOBS)

        if SHEET_DB_UPDATES in wb_out.sheetnames and SHEET_DASHBOARD_TIMING in wb_out.sheetnames:
            populate_db_updates_sheet(wb_out[SHEET_DB_UPDATES], wb_out[SHEET_DASHBOARD_TIMING])
        else:
            log.info("No '%s' sheet found in output workbook — skipping shift-timing auto-populate.",
                      SHEET_DB_UPDATES)

        # Alerts are isolated: if this step fails the rest of the workbook is still saved, but the
        # alert sheet is left untouched (never half-written), the failure is logged loudly and the
        # process exits non-zero so a scheduler/monitor notices.
        alerts_failed = False
        try:
            categories = (read_dashboard_categories(wb_out[SHEET_DASHBOARD_TIMING])
                          if SHEET_DASHBOARD_TIMING in wb_out.sheetnames else {})
            alerts, astats = build_dashboard_alerts(groups, results, schedules, categories, now_ist)
            by_sev = pd.Series([a.severity for a in alerts]).value_counts().to_dict()
            log.info("Alert evaluation: %d dashboards | %d validated (scheduled today) | %d ignored "
                     "(not scheduled today) | %d pending (run window not closed yet) | %d periodicity "
                     "undetermined | %d ALERTS %s",
                     astats["total"], astats["validated"], astats["not_scheduled_today"], astats["pending"],
                     len(astats["undetermined"]), len(alerts), by_sev or "")
            for item in astats["undetermined"]:
                log.warning("Periodicity not understood -> NOT validated: %s", item)
            for item in astats["schedule_errors"]:
                log.warning("Schedule conversion problem: %s", item)
            for a in alerts:
                log.info("  ALERT [%s] %s — %s", a.severity, a.dashboard, a.status)
            write_alerts_sheet(wb_out, alerts, now_ist)
        except Exception as e:
            alerts_failed = True
            log.exception("Could not update '%s' — the sheet was left as it was: %s", SHEET_ALERTS, e)

        wb_out.save(output_path)
        log.info("Saved results to '%s'.", output_path)
        if alerts_failed:
            sys.exit(2)

    except Exception as e:
        log.exception("Pipeline failed: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    main()