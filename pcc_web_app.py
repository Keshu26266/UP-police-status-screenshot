"""
PCC Verification Screenshot Tool — Cloud & Local Web App
-----------------------------------------------------------
Deployable to Railway, Docker, or local machine.
Upload your Excel file (columns: "PCC APPLICATION NUMBER", "RENAME ID"),
click Start, monitor real-time progress, and download a ZIP of all screenshots + audit log.
"""

import os
import io
import time
import random
import shutil
import zipfile
import threading
import webbrowser
import pandas as pd
from datetime import datetime
from flask import Flask, request, jsonify, send_file, Response
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ============ CONFIGURATION & SETTINGS ============
APP_NUMBER_COLUMN = os.environ.get("APP_NUMBER_COLUMN", "PCC APPLICATION NUMBER")
RENAME_ID_COLUMN = os.environ.get("RENAME_ID_COLUMN", "RENAME ID")
SERVICE_TYPE_TEXT = "CHARACTER CERTIFICATE"
PORTAL_URL = "https://cctnsup.gov.in/citizenportal/CitizenVerification.aspx"

WORK_DIR = os.environ.get("WORK_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "run_data"))
OUTPUT_DIR = os.path.join(WORK_DIR, "screenshots")
LOG_FILE = os.path.join(WORK_DIR, "results_log.csv")
UPLOAD_PATH = os.path.join(WORK_DIR, "uploaded_input.xlsx")

MIN_DELAY_SECONDS = float(os.environ.get("MIN_DELAY_SECONDS", 3))
MAX_DELAY_SECONDS = float(os.environ.get("MAX_DELAY_SECONDS", 6))
PAGE_LOAD_TIMEOUT_MS = int(os.environ.get("PAGE_LOAD_TIMEOUT_MS", 25000))
RESULT_WAIT_TIMEOUT_MS = int(os.environ.get("RESULT_WAIT_TIMEOUT_MS", 15000))
MAX_SCREENSHOT_BYTES = 200 * 1024
HEADLESS = os.environ.get("HEADLESS", "True").lower() in ("true", "1", "yes")

HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", 5000))
# ==================================================

app = Flask(__name__)

STATE = {
    "running": False,
    "total": 0,
    "done": 0,
    "success_count": 0,
    "not_found_count": 0,
    "failed_count": 0,
    "current": "",
    "log": [],       # list of dicts: rename_id, application_number, status, message, timestamp
    "finished": False,
    "error": None,
}
STATE_LOCK = threading.Lock()


# ---------------- Automation Logic ----------------

def save_screenshot_under_size(page, out_path, max_bytes=MAX_SCREENSHOT_BYTES):
    quality_levels = [90, 85, 80, 75, 70, 65, 60, 55, 50, 45, 40]
    best_bytes = None
    for q in quality_levels:
        img_bytes = page.screenshot(full_page=True, type="jpeg", quality=q)
        if len(img_bytes) <= max_bytes:
            with open(out_path, "wb") as f:
                f.write(img_bytes)
            return True
        best_bytes = img_bytes
    if best_bytes is not None:
        with open(out_path, "wb") as f:
            f.write(best_bytes)
    return False


def hide_footer(page):
    try:
        page.evaluate("""
            () => {
                const markers = ['CAS CITIZEN', 'National Crime Records Bureau', 'Terms and Conditions'];
                const all = document.querySelectorAll('body *');
                for (const el of all) {
                    const txt = (el.textContent || '').trim();
                    if (txt.length === 0 || txt.length > 400) continue;
                    if (!markers.some(m => txt.includes(m))) continue;
                    let target = el;
                    while (target.parentElement && target.parentElement.tagName !== 'BODY') {
                        const parentTxt = (target.parentElement.textContent || '').trim();
                        if (parentTxt.length > 400) break;
                        target = target.parentElement;
                    }
                    const rect = target.getBoundingClientRect();
                    if (rect.height > 0 && rect.height < 150) {
                        target.style.display = 'none';
                    }
                }
            }
        """)
    except Exception:
        pass


def select_service_type(page):
    dropdown = page.locator("select").filter(has_text="CHARACTER CERTIFICATE")
    if dropdown.count() > 0:
        dropdown.first.select_option(label=SERVICE_TYPE_TEXT)
        return True
    label_locator = page.get_by_text("Service Request Type", exact=False)
    if label_locator.count() > 0:
        select_el = label_locator.first.locator("xpath=following::select[1]")
        select_el.select_option(label=SERVICE_TYPE_TEXT)
        return True
    return False


def fill_application_number(page, app_number):
    label_locator = page.get_by_text("Complaint/Service Request No", exact=False)
    if label_locator.count() == 0:
        label_locator = page.get_by_text("Service Request No", exact=False)
    if label_locator.count() > 0:
        input_el = label_locator.first.locator("xpath=following::input[1]")
        input_el.fill("")
        input_el.fill(str(app_number))
        return True
    return False


def click_search(page):
    btn = page.get_by_role("button", name="Search", exact=False)
    if btn.count() > 0:
        btn.first.click()
        return True
    btn = page.locator("input[type=submit]").filter(has_text="Search")
    if btn.count() > 0:
        btn.first.click()
        return True
    btn = page.locator("input[value='Search']")
    if btn.count() > 0:
        btn.first.click()
        return True
    return False


def process_one(page, app_number, rename_id):
    out_path = os.path.join(OUTPUT_DIR, f"{rename_id}.jpg")

    if not select_service_type(page):
        return "failed", "Could not find/select Service Request Type dropdown"
    if not fill_application_number(page, app_number):
        return "failed", "Could not find application number input box"
    if not click_search(page):
        return "failed", "Could not find/click Search button"

    try:
        page.wait_for_timeout(1500)
        page.wait_for_load_state("networkidle", timeout=RESULT_WAIT_TIMEOUT_MS)
    except PWTimeout:
        pass

    page_text = page.inner_text("body")
    hide_footer(page)

    if "no record" in page_text.lower() or "not found" in page_text.lower():
        save_screenshot_under_size(page, out_path)
        return "not_found", "Portal reported no record found (screenshot saved anyway)"

    save_screenshot_under_size(page, out_path)
    return "success", "Screenshot saved successfully"


def run_batch():
    with STATE_LOCK:
        STATE["running"] = True
        STATE["done"] = 0
        STATE["success_count"] = 0
        STATE["not_found_count"] = 0
        STATE["failed_count"] = 0
        STATE["log"] = []
        STATE["finished"] = False
        STATE["error"] = None

    try:
        # Reset output folders
        if os.path.exists(OUTPUT_DIR):
            shutil.rmtree(OUTPUT_DIR, ignore_errors=True)
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        if os.path.exists(LOG_FILE):
            try:
                os.remove(LOG_FILE)
            except OSError:
                pass

        df = pd.read_excel(UPLOAD_PATH)
        df = df.dropna(subset=[APP_NUMBER_COLUMN, RENAME_ID_COLUMN])

        with STATE_LOCK:
            STATE["total"] = len(df)

        # Launch Playwright Chromium with cloud & container hardening flags
        launch_kwargs = {
            "headless": HEADLESS,
            "args": [
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--no-first-run",
                "--no-default-browser-check"
            ]
        }

        # Optional proxy configuration (supports HTTP/HTTPS/SOCKS5 if needed for geo-unblocking)
        proxy_url = os.environ.get("PROXY_URL") or os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
        if proxy_url:
            launch_kwargs["proxy"] = {"server": proxy_url}

        with sync_playwright() as p:
            browser = p.chromium.launch(**launch_kwargs)
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                viewport={"width": 1280, "height": 900}
            )
            page = context.new_page()

            for idx, row in df.iterrows():
                app_number = str(row[APP_NUMBER_COLUMN]).strip()
                rename_id = str(row[RENAME_ID_COLUMN]).strip()

                with STATE_LOCK:
                    STATE["current"] = f"{app_number} → {rename_id}.jpg"

                try:
                    page.goto(PORTAL_URL, timeout=PAGE_LOAD_TIMEOUT_MS)
                    status, message = process_one(page, app_number, rename_id)
                except Exception as e:
                    status, message = "failed", f"Exception: {str(e)[:180]}"

                entry = {
                    "rename_id": rename_id,
                    "application_number": app_number,
                    "status": status,
                    "message": message,
                    "timestamp": datetime.now().strftime("%H:%M:%S"),
                }
                with STATE_LOCK:
                    STATE["log"].append(entry)
                    STATE["done"] += 1
                    if status == "success":
                        STATE["success_count"] += 1
                    elif status == "not_found":
                        STATE["not_found_count"] += 1
                    else:
                        STATE["failed_count"] += 1

                time.sleep(random.uniform(MIN_DELAY_SECONDS, MAX_DELAY_SECONDS))

            browser.close()

        pd.DataFrame(STATE["log"]).to_csv(LOG_FILE, index=False)

    except Exception as e:
        with STATE_LOCK:
            STATE["error"] = str(e)
    finally:
        with STATE_LOCK:
            STATE["running"] = False
            STATE["finished"] = True


# ---------------- Web Interface & Routes ----------------

PAGE_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>PCC Verification Screenshot Automator</title>
<meta name="description" content="Automated UP Police CCTNS Character Certificate / PCC Verification and Screenshot Tool">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
  :root {
    --bg-base: #090d16;
    --bg-card: rgba(17, 24, 39, 0.75);
    --bg-card-hover: rgba(24, 34, 53, 0.85);
    --border-card: rgba(255, 255, 255, 0.08);
    --border-card-hover: rgba(99, 102, 241, 0.4);
    --primary: #6366f1;
    --primary-gradient: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #d946ef 100%);
    --primary-hover: #4f46e5;
    --text-main: #f3f4f6;
    --text-muted: #9ca3af;
    --success: #10b981;
    --success-bg: rgba(16, 185, 129, 0.12);
    --warning: #f59e0b;
    --warning-bg: rgba(245, 158, 11, 0.12);
    --danger: #f43f5e;
    --danger-bg: rgba(244, 63, 94, 0.12);
    --radius-lg: 16px;
    --radius-md: 10px;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    background-color: var(--bg-base);
    background-image: 
      radial-gradient(at 0% 0%, rgba(99, 102, 241, 0.15) 0px, transparent 50%),
      radial-gradient(at 100% 100%, rgba(217, 70, 239, 0.1) 0px, transparent 50%);
    color: var(--text-main);
    min-height: 100vh;
    padding: 32px 20px 60px;
    display: flex;
    flex-direction: column;
    align-items: center;
  }

  .container {
    width: 100%;
    max-width: 860px;
    display: flex;
    flex-direction: column;
    gap: 24px;
  }

  /* Header */
  .header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 16px;
    padding-bottom: 8px;
  }
  .brand-group {
    display: flex;
    align-items: center;
    gap: 14px;
  }
  .brand-icon {
    width: 46px;
    height: 46px;
    border-radius: 12px;
    background: var(--primary-gradient);
    display: flex;
    align-items: center;
    justify-content: center;
    box-shadow: 0 8px 20px -4px rgba(99, 102, 241, 0.5);
  }
  .brand-icon svg { width: 26px; height: 26px; fill: white; }
  .header h1 {
    font-size: 22px;
    font-weight: 700;
    letter-spacing: -0.02em;
    color: #ffffff;
  }
  .header p {
    font-size: 13px;
    color: var(--text-muted);
  }
  .badge-online {
    display: inline-flex;
    align-items: center;
    gap: 7px;
    padding: 6px 13px;
    background: rgba(16, 185, 129, 0.12);
    border: 1px solid rgba(16, 185, 129, 0.28);
    border-radius: 9999px;
    font-size: 12px;
    font-weight: 600;
    color: #34d399;
  }
  .pulse-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: #10b981;
    box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7);
    animation: pulse 2s infinite;
  }
  @keyframes pulse {
    0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
    70% { transform: scale(1); box-shadow: 0 0 0 8px rgba(16, 185, 129, 0); }
    100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
  }

  /* Glass Cards */
  .card {
    background: var(--bg-card);
    backdrop-filter: blur(16px);
    -webkit-backdrop-filter: blur(16px);
    border: 1px solid var(--border-card);
    border-radius: var(--radius-lg);
    padding: 28px;
    box-shadow: 0 20px 40px -15px rgba(0, 0, 0, 0.6);
    transition: border-color 0.2s, box-shadow 0.2s;
  }

  /* Upload Area */
  .upload-zone {
    border: 2px dashed rgba(255, 255, 255, 0.16);
    border-radius: var(--radius-md);
    padding: 36px 20px;
    text-align: center;
    cursor: pointer;
    transition: all 0.25s ease;
    background: rgba(255, 255, 255, 0.02);
  }
  .upload-zone:hover, .upload-zone.dragover {
    border-color: var(--primary);
    background: rgba(99, 102, 241, 0.05);
  }
  .upload-icon {
    width: 52px;
    height: 52px;
    margin: 0 auto 14px;
    background: rgba(255, 255, 255, 0.05);
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #a5b4fc;
  }
  .upload-zone h3 { font-size: 16px; font-weight: 600; margin-bottom: 6px; }
  .upload-zone p { font-size: 13px; color: var(--text-muted); }
  .columns-hint {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    margin-top: 14px;
    flex-wrap: wrap;
  }
  .tag {
    font-size: 11px;
    font-weight: 600;
    padding: 4px 9px;
    background: rgba(255, 255, 255, 0.06);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 6px;
    color: #e0e7ff;
    font-family: monospace;
  }
  #fileInput { display: none; }
  .file-selected-info {
    margin-top: 14px;
    font-size: 13px;
    color: #818cf8;
    font-weight: 600;
  }

  /* Buttons */
  .btn-primary {
    width: 100%;
    margin-top: 20px;
    padding: 14px 24px;
    border: none;
    border-radius: var(--radius-md);
    background: var(--primary-gradient);
    color: #ffffff;
    font-size: 15px;
    font-weight: 600;
    cursor: pointer;
    box-shadow: 0 10px 25px -5px rgba(99, 102, 241, 0.4);
    transition: transform 0.15s ease, box-shadow 0.15s ease, opacity 0.15s;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
  }
  .btn-primary:hover:not(:disabled) {
    transform: translateY(-1px);
    box-shadow: 0 14px 28px -5px rgba(99, 102, 241, 0.55);
  }
  .btn-primary:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  /* Stats Grid */
  .stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
    gap: 12px;
    margin-bottom: 20px;
  }
  .stat-card {
    background: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.06);
    border-radius: var(--radius-md);
    padding: 14px;
    text-align: center;
  }
  .stat-val {
    font-size: 24px;
    font-weight: 700;
    letter-spacing: -0.02em;
    margin-bottom: 2px;
  }
  .stat-label {
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--text-muted);
    font-weight: 600;
  }
  .stat-card.total .stat-val { color: #818cf8; }
  .stat-card.success .stat-val { color: #34d399; }
  .stat-card.warning .stat-val { color: #fbbf24; }
  .stat-card.failed .stat-val { color: #fb7185; }

  /* Progress Bar */
  .progress-wrapper {
    margin-bottom: 16px;
  }
  .progress-header {
    display: flex;
    justify-content: space-between;
    font-size: 13px;
    color: var(--text-muted);
    margin-bottom: 8px;
  }
  .progress-bar-bg {
    width: 100%;
    height: 10px;
    background: rgba(255, 255, 255, 0.07);
    border-radius: 9999px;
    overflow: hidden;
  }
  .progress-bar-fill {
    height: 100%;
    width: 0%;
    background: var(--primary-gradient);
    border-radius: 9999px;
    transition: width 0.3s ease;
  }
  .current-ticker {
    margin-top: 10px;
    font-size: 12px;
    font-family: monospace;
    color: #c7d2fe;
    background: rgba(99, 102, 241, 0.08);
    padding: 8px 12px;
    border-radius: 6px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  /* Log Table */
  .table-container {
    max-height: 280px;
    overflow-y: auto;
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: var(--radius-md);
    margin-top: 16px;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
    text-align: left;
  }
  th {
    background: #131b2e;
    padding: 10px 14px;
    color: #94a3b8;
    font-weight: 600;
    position: sticky;
    top: 0;
    z-index: 2;
  }
  td {
    padding: 9px 14px;
    border-top: 1px solid rgba(255, 255, 255, 0.05);
    color: #e2e8f0;
  }
  tr:hover td { background: rgba(255, 255, 255, 0.02); }

  /* Badges */
  .status-badge {
    display: inline-block;
    padding: 3px 8px;
    border-radius: 4px;
    font-size: 11px;
    font-weight: 600;
    text-transform: capitalize;
  }
  .status-badge.success { background: var(--success-bg); color: #34d399; }
  .status-badge.not_found { background: var(--warning-bg); color: #fbbf24; }
  .status-badge.failed { background: var(--danger-bg); color: #fb7185; }

  /* Completion Action */
  .completion-box {
    margin-top: 20px;
    padding: 18px;
    border-radius: var(--radius-md);
    background: rgba(16, 185, 129, 0.08);
    border: 1px solid rgba(16, 185, 129, 0.25);
    display: flex;
    align-items: center;
    justify-content: space-between;
    flex-wrap: wrap;
    gap: 12px;
  }
  .btn-download {
    padding: 10px 20px;
    border: none;
    border-radius: var(--radius-md);
    background: #10b981;
    color: white;
    font-weight: 600;
    font-size: 14px;
    cursor: pointer;
    text-decoration: none;
    display: inline-flex;
    align-items: center;
    gap: 8px;
    box-shadow: 0 8px 20px -4px rgba(16, 185, 129, 0.4);
    transition: background 0.15s;
  }
  .btn-download:hover { background: #059669; }
  .btn-reset {
    padding: 10px 16px;
    border: 1px solid rgba(255, 255, 255, 0.15);
    border-radius: var(--radius-md);
    background: transparent;
    color: var(--text-muted);
    font-size: 13px;
    cursor: pointer;
  }
  .btn-reset:hover { color: white; border-color: rgba(255, 255, 255, 0.3); }

  /* Footer */
  .footer {
    text-align: center;
    font-size: 12px;
    color: #64748b;
    margin-top: 10px;
  }
</style>
</head>
<body>

<div class="container">
  <!-- Brand Header -->
  <header class="header">
    <div class="brand-group">
      <div class="brand-icon">
        <svg viewBox="0 0 24 24"><path d="M12 2L4 5v6.09c0 5.05 3.41 9.76 8 10.91 4.59-1.15 8-5.86 8-10.91V5l-8-3zm-1 15l-4-4 1.41-1.41L11 14.17l6.59-6.59L19 9l-8 8z"/></svg>
      </div>
      <div>
        <h1 id="appTitle">PCC Verification Automator</h1>
        <p>UP Police CCTNS Citizen Portal Automated Verification & Screenshot Capture</p>
      </div>
    </div>
    <div class="badge-online">
      <span class="pulse-dot"></span>
      <span>System Ready</span>
    </div>
  </header>

  <!-- Upload Section -->
  <section class="card" id="uploadSection">
    <form id="uploadForm">
      <div class="upload-zone" id="dropZone" onclick="document.getElementById('fileInput').click()">
        <div class="upload-icon">
          <svg width="28" height="28" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path></svg>
        </div>
        <h3>Select or Drop Excel Workbook</h3>
        <p>Choose an <code>.xlsx</code> file containing applicant records</p>
        <div class="columns-hint">
          <span>Required Columns:</span>
          <span class="tag">PCC APPLICATION NUMBER</span>
          <span class="tag">RENAME ID</span>
        </div>
        <div class="file-selected-info" id="selectedFileName"></div>
        <input type="file" id="fileInput" name="file" accept=".xlsx">
      </div>

      <button type="submit" id="startBtn" class="btn-primary" disabled>
        <svg width="18" height="18" fill="currentColor" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
        Start Verification Process
      </button>
    </form>
  </section>

  <!-- Progress & Execution Section -->
  <section class="card" id="progressSection" style="display: none;">
    <!-- Stat Counters -->
    <div class="stats-grid">
      <div class="stat-card total">
        <div class="stat-val" id="statTotal">0</div>
        <div class="stat-label">Total Records</div>
      </div>
      <div class="stat-card">
        <div class="stat-val" id="statDone">0</div>
        <div class="stat-label">Processed</div>
      </div>
      <div class="stat-card success">
        <div class="stat-val" id="statSuccess">0</div>
        <div class="stat-label">Verified (OK)</div>
      </div>
      <div class="stat-card warning">
        <div class="stat-val" id="statNotFound">0</div>
        <div class="stat-label">Not Found</div>
      </div>
      <div class="stat-card failed">
        <div class="stat-val" id="statFailed">0</div>
        <div class="stat-label">Failed</div>
      </div>
    </div>

    <!-- Progress Meter -->
    <div class="progress-wrapper">
      <div class="progress-header">
        <span id="progressStatusText">Running automation...</span>
        <span id="progressPercentage">0%</span>
      </div>
      <div class="progress-bar-bg">
        <div class="progress-bar-fill" id="progressBarFill"></div>
      </div>
      <div class="current-ticker" id="currentTicker">Waiting for first record...</div>
    </div>

    <!-- Live Audit Log Table -->
    <div class="table-container">
      <table id="logTable">
        <thead>
          <tr>
            <th>Time</th>
            <th>Rename ID</th>
            <th>Application No</th>
            <th>Status</th>
            <th>Message</th>
          </tr>
        </thead>
        <tbody id="logTableBody"></tbody>
      </table>
    </div>

    <!-- Completion Action -->
    <div class="completion-box" id="completionBox" style="display: none;">
      <div>
        <strong style="color: #34d399; font-size: 15px;">✓ Batch Processing Complete!</strong>
        <p style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">All screenshots and CSV results have been bundled into a single ZIP archive.</p>
      </div>
      <div style="display: flex; gap: 8px;">
        <button type="button" class="btn-reset" onclick="resetApp()">New Batch</button>
        <a id="downloadBtn" href="/download" class="btn-download">
          <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM17 13l-5 5-5-5h3V9h4v4h3z"/></svg>
          Download Package (.ZIP)
        </a>
      </div>
    </div>
  </section>

  <footer class="footer">
    <span>PCC Verification Tool &bull; Cloud & Local Container Ready</span>
  </footer>
</div>

<script>
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const startBtn = document.getElementById('startBtn');
const selectedFileName = document.getElementById('selectedFileName');
const uploadForm = document.getElementById('uploadForm');
const uploadSection = document.getElementById('uploadSection');
const progressSection = document.getElementById('progressSection');
const completionBox = document.getElementById('completionBox');

// Drag & drop handlers
['dragenter', 'dragover'].forEach(name => {
  dropZone.addEventListener(name, (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
});
['dragleave', 'drop'].forEach(name => {
  dropZone.addEventListener(name, (e) => { e.preventDefault(); dropZone.classList.remove('dragover'); });
});
dropZone.addEventListener('drop', (e) => {
  if (e.dataTransfer.files.length) {
    fileInput.files = e.dataTransfer.files;
    handleFileSelected();
  }
});
fileInput.addEventListener('change', handleFileSelected);

function handleFileSelected() {
  if (fileInput.files.length > 0) {
    const file = fileInput.files[0];
    selectedFileName.textContent = `Selected: ${file.name} (${Math.round(file.size / 1024)} KB)`;
    startBtn.disabled = false;
  } else {
    selectedFileName.textContent = '';
    startBtn.disabled = true;
  }
}

uploadForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  if (!fileInput.files.length) return;

  startBtn.disabled = true;
  startBtn.innerHTML = 'Uploading & Initializing...';

  const formData = new FormData();
  formData.append('file', fileInput.files[0]);

  try {
    const res = await fetch('/upload', { method: 'POST', body: formData });
    const data = await res.json();
    if (data.error) {
      alert(data.error);
      startBtn.disabled = false;
      startBtn.innerHTML = 'Start Verification Process';
      return;
    }
    uploadSection.style.display = 'none';
    progressSection.style.display = 'block';
    poll();
  } catch (err) {
    alert('Upload failed: ' + err.message);
    startBtn.disabled = false;
    startBtn.innerHTML = 'Start Verification Process';
  }
});

async function poll() {
  try {
    const res = await fetch('/status');
    const data = await res.json();

    const total = data.total || 0;
    const done = data.done || 0;
    const pct = total ? Math.round((done / total) * 100) : 0;

    document.getElementById('statTotal').textContent = total;
    document.getElementById('statDone').textContent = done;
    document.getElementById('statSuccess').textContent = data.success_count || 0;
    document.getElementById('statNotFound').textContent = data.not_found_count || 0;
    document.getElementById('statFailed').textContent = data.failed_count || 0;

    document.getElementById('progressBarFill').style.width = pct + '%';
    document.getElementById('progressPercentage').textContent = pct + '%';

    if (data.current) {
      document.getElementById('currentTicker').textContent = 'Processing: ' + data.current;
    }

    if (data.error) {
      document.getElementById('progressStatusText').textContent = 'Error: ' + data.error;
      document.getElementById('progressStatusText').style.color = '#fb7185';
    }

    // Populate log table
    const tbody = document.getElementById('logTableBody');
    tbody.innerHTML = '';
    (data.log || []).slice().reverse().forEach(row => {
      const tr = document.createElement('tr');
      const badgeClass = row.status === 'success' ? 'success' : (row.status === 'not_found' ? 'not_found' : 'failed');
      tr.innerHTML = `
        <td style="color: #64748b;">${row.timestamp || ''}</td>
        <td><strong>${row.rename_id}</strong></td>
        <td>${row.application_number}</td>
        <td><span class="status-badge ${badgeClass}">${row.status.replace('_', ' ')}</span></td>
        <td>${row.message}</td>
      `;
      tbody.appendChild(tr);
    });

    if (data.finished) {
      document.getElementById('currentTicker').textContent = 'Completed ' + done + ' of ' + total + ' records.';
      completionBox.style.display = 'flex';
    } else {
      setTimeout(poll, 1500);
    }
  } catch (err) {
    setTimeout(poll, 3000);
  }
}

function resetApp() {
  uploadSection.style.display = 'block';
  progressSection.style.display = 'none';
  completionBox.style.display = 'none';
  fileInput.value = '';
  selectedFileName.textContent = '';
  startBtn.disabled = true;
  startBtn.innerHTML = '<svg width="18" height="18" fill="currentColor" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg> Start Verification Process';
}
</script>
</body>
</html>
"""


@app.route("/")
def index():
    return Response(PAGE_HTML, mimetype="text/html")


@app.route("/health")
def health():
    return jsonify({
        "status": "healthy",
        "service": "pcc-verification-tool",
        "running": STATE["running"]
    })


@app.route("/upload", methods=["POST"])
def upload():
    if STATE["running"]:
        return jsonify({"error": "A batch is already running. Please wait for it to finish."}), 400

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return jsonify({"error": "Please upload a valid .xlsx Excel workbook."}), 400

    os.makedirs(WORK_DIR, exist_ok=True)
    f.save(UPLOAD_PATH)

    thread = threading.Thread(target=run_batch, daemon=True)
    thread.start()
    return jsonify({"started": True})


@app.route("/status")
def status():
    with STATE_LOCK:
        return jsonify(STATE)


@app.route("/download")
def download():
    mem_zip = io.BytesIO()
    with zipfile.ZipFile(mem_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        if os.path.exists(OUTPUT_DIR):
            for fname in os.listdir(OUTPUT_DIR):
                zf.write(os.path.join(OUTPUT_DIR, fname), arcname=f"screenshots/{fname}")
        if os.path.exists(LOG_FILE):
            zf.write(LOG_FILE, arcname="results_log.csv")
    mem_zip.seek(0)
    return send_file(mem_zip, mimetype="application/zip", as_attachment=True,
                      download_name=f"pcc_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip")


if __name__ == "__main__":
    is_cloud = os.environ.get("RAILWAY_ENVIRONMENT") or os.environ.get("PORT")
    if not is_cloud:
        threading.Timer(1.2, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}")).start()
    print(f"Server starting on http://{HOST}:{PORT}")
    app.run(host=HOST, port=PORT, debug=False)
