#!/usr/bin/env python3
"""
RFID Attendance System - Live Dashboard & Excel Sync
Communicates with Arduino Mega 2560 over Serial, tracks attendance,
auto-generates styled Excel (.xlsx) and CSV files, and serves a live web dashboard.
"""

import os
import sys
import time
import json
import csv
import glob
import threading
from datetime import datetime

import serial
from bottle import Bottle, request, response, static_file, redirect
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Configuration
PORT = "/dev/cu.usbmodem11201"
BAUD = 9600
WEB_PORT = 5050
DATA_DIR = os.path.dirname(os.path.abspath(__file__))
USERS_FILE = os.path.join(DATA_DIR, "users.json")
EXCEL_FILE = os.path.join(DATA_DIR, "attendance.xlsx")
CSV_FILE = os.path.join(DATA_DIR, "attendance.csv")

# Default hardcoded users
DEFAULT_USERS = [
    {"uid": "2a472f3",  "roll": "HST - 01", "name": "Vivek"},
    {"uid": "29c224b7", "roll": "HST - 02", "name": "Ritik"},
    {"uid": "d33ff0",   "roll": "HST - 03", "name": "Ainul"},
]

# Thread safety lock
lock = threading.Lock()

# State
app_state = {
    "serial_connected": False,
    "serial_port": PORT,
    "last_seen_serial": None,
    "users": {},           # uid -> {uid, roll, name, status, time_in, time_out, is_inside}
    "unknown_scans": [],   # [{timestamp, uid, status}]
    "activity_log": [],    # [{timestamp, serial_num, uid, roll, name, time_in, time_out, status}]
}

def load_users():
    """Load users from users.json or initialize with defaults."""
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, "r") as f:
                data = json.load(f)
                users = {}
                for u in data:
                    uid = u["uid"].lower()
                    users[uid] = {
                        "uid": uid,
                        "roll": u.get("roll", "N/A"),
                        "name": u.get("name", "Unknown"),
                        "status": "ABSENT",
                        "time_in": "-",
                        "time_out": "-",
                        "is_inside": False
                    }
                return users
        except Exception as e:
            print(f"[Warning] Could not read {USERS_FILE}: {e}")

    # Fallback to default users
    users = {}
    for u in DEFAULT_USERS:
        uid = u["uid"].lower()
        users[uid] = {
            "uid": uid,
            "roll": u["roll"],
            "name": u["name"],
            "status": "ABSENT",
            "time_in": "-",
            "time_out": "-",
            "is_inside": False
        }
    save_users_config(users)
    return users

def save_users_config(users):
    """Save user directory to JSON."""
    data = [{"uid": u["uid"], "roll": u["roll"], "name": u["name"]} for u in users.values()]
    with open(USERS_FILE, "w") as f:
        json.dump(data, f, indent=2)

def save_csv():
    """Export current daily attendance to CSV."""
    with lock:
        try:
            today_str = datetime.now().strftime("%Y-%m-%d")
            with open(CSV_FILE, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Date", "Roll Number", "Name", "RFID Card UID", "Time In", "Time Out", "Status"])
                for u in app_state["users"].values():
                    writer.writerow([today_str, u["roll"], u["name"], u["uid"], u["time_in"], u["time_out"], u["status"]])
                
                # Unknown scans section
                if app_state["unknown_scans"]:
                    writer.writerow([])
                    writer.writerow(["--- UNREGISTERED / UNKNOWN SCANS ---"])
                    writer.writerow(["Date", "Scan Time", "Detected Card UID", "Status"])
                    for unk in app_state["unknown_scans"]:
                        writer.writerow([today_str, unk["timestamp"], unk["uid"], "UNKNOWN / NOT REGISTERED"])
        except Exception as e:
            print(f"[Error] Failed to save CSV: {e}")

def save_excel():
    """Export styled Excel sheet with formatting."""
    with lock:
        try:
            wb = openpyxl.Workbook()
            today_str = datetime.now().strftime("%Y-%m-%d")

            # Sheet 1: Daily Attendance
            ws = wb.active
            ws.title = "Daily Attendance"
            ws.views.sheetView[0].showGridLines = True

            # Styles
            title_font = Font(name="Calibri", size=16, bold=True, color="1F497D")
            subtitle_font = Font(name="Calibri", size=11, italic=True, color="595959")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
            
            fill_present = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
            font_present = Font(name="Calibri", size=11, bold=True, color="155724")

            fill_exited = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
            font_exited = Font(name="Calibri", size=11, bold=True, color="004085")

            fill_absent = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
            font_absent = Font(name="Calibri", size=11, color="721C24")

            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

            # Titles
            ws["A1"] = "RFID ATTENDANCE DASHBOARD"
            ws["A1"].font = title_font
            ws["A2"] = f"Date: {today_str} | Generated: {datetime.now().strftime('%H:%M:%S')}"
            ws["A2"].font = subtitle_font

            # Table Headers
            headers = ["Roll Number", "Student / Employee Name", "RFID Card UID", "Time In", "Time Out", "Attendance Status"]
            start_row = 4
            for col_num, h in enumerate(headers, 1):
                cell = ws.cell(row=start_row, column=col_num)
                cell.value = h
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")

            current_row = start_row + 1
            for u in app_state["users"].values():
                row_data = [u["roll"], u["name"], u["uid"], u["time_in"], u["time_out"], u["status"]]
                for col_num, val in enumerate(row_data, 1):
                    cell = ws.cell(row=current_row, column=col_num)
                    cell.value = val
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center" if col_num != 2 else "left", vertical="center")
                    
                    if col_num == 6: # Status
                        if u["status"] in ("PRESENT", "STUDENT ENTERED"):
                            cell.fill = fill_present
                            cell.font = font_present
                        elif u["status"] in ("STUDENT EXITED", "EXITED"):
                            cell.fill = fill_exited
                            cell.font = font_exited
                        else:
                            cell.fill = fill_absent
                            cell.font = font_absent
                current_row += 1

            # Sheet 2: Unregistered Scans
            ws2 = wb.create_sheet(title="Unregistered Cards")
            ws2.views.sheetView[0].showGridLines = True
            ws2["A1"] = "UNREGISTERED / UNKNOWN RFID DETECTIONS"
            ws2["A1"].font = title_font
            ws2["A2"] = f"Date: {today_str}"
            ws2["A2"].font = subtitle_font

            unk_headers = ["Scan #", "Timestamp", "Detected Card UID", "Status"]
            unk_header_fill = PatternFill(start_color="D9534F", end_color="D9534F", fill_type="solid")
            for col_num, h in enumerate(unk_headers, 1):
                cell = ws2.cell(row=4, column=col_num)
                cell.value = h
                cell.font = header_font
                cell.fill = unk_header_fill
                cell.alignment = Alignment(horizontal="center")

            row_unk = 5
            for idx, unk in enumerate(app_state["unknown_scans"], 1):
                ws2.cell(row=row_unk, column=1, value=idx).alignment = Alignment(horizontal="center")
                ws2.cell(row=row_unk, column=2, value=unk["timestamp"]).alignment = Alignment(horizontal="center")
                ws2.cell(row=row_unk, column=3, value=unk["uid"]).alignment = Alignment(horizontal="center")
                ws2.cell(row=row_unk, column=4, value="UNREGISTERED CARD").alignment = Alignment(horizontal="center")
                row_unk += 1

            # Auto-fit column widths
            for sheet in (ws, ws2):
                for col in sheet.columns:
                    max_len = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.value:
                            max_len = max(max_len, len(str(cell.value)))
                    sheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(EXCEL_FILE)
        except Exception as e:
            print(f"[Error] Failed to save Excel: {e}")

def find_arduino_port():
    """Detect Arduino Mega serial port."""
    if os.path.exists(PORT):
        return PORT
    candidates = glob.glob("/dev/cu.usbmodem*") + glob.glob("/dev/tty.usbmodem*")
    return candidates[0] if candidates else PORT

def process_serial_line(line):
    """Parse comma-separated serial data from Arduino."""
    line = line.strip()
    if not line:
        return

    print(f"[Arduino Monitor] {line}")
    timestamp_now = datetime.now().strftime("%H:%M:%S")

    # Header or info lines
    if "RFID ATTENDANCE" in line or "MFRC522" in line or line.startswith("Serial Number,"):
        return

    parts = [p.strip() for p in line.split(",")]
    if len(parts) >= 7:
        serial_num = parts[0]
        card_uid = parts[1].lower()
        roll = parts[2]
        name = parts[3]
        time_in = parts[4]
        time_out = parts[5]
        status = parts[6].upper()

        with lock:
            if card_uid in app_state["users"]:
                u = app_state["users"][card_uid]
                if "ENTERED" in status:
                    u["status"] = "PRESENT"
                    u["time_in"] = timestamp_now if time_in == "-" else time_in
                    u["is_inside"] = True
                elif "EXITED" in status:
                    u["status"] = "EXITED"
                    u["time_out"] = timestamp_now if time_out == "-" else time_out
                    u["is_inside"] = False
                
                app_state["activity_log"].insert(0, {
                    "timestamp": timestamp_now,
                    "serial_num": serial_num,
                    "uid": card_uid,
                    "roll": u["roll"],
                    "name": u["name"],
                    "status": u["status"]
                })
            else:
                # Unknown Card
                unk_entry = {
                    "timestamp": timestamp_now,
                    "uid": card_uid,
                    "status": "UNKNOWN CARD DETECTED"
                }
                app_state["unknown_scans"].insert(0, unk_entry)
                app_state["activity_log"].insert(0, {
                    "timestamp": timestamp_now,
                    "serial_num": serial_num,
                    "uid": card_uid,
                    "roll": "Unknown",
                    "name": "Unregistered Card",
                    "status": "UNKNOWN"
                })

        # Persist updates to Excel and CSV
        save_csv()
        save_excel()

def serial_worker():
    """Background thread continuously reading Arduino serial stream."""
    while True:
        port = find_arduino_port()
        try:
            ser = serial.Serial(port, BAUD, timeout=1)
            with lock:
                app_state["serial_connected"] = True
                app_state["serial_port"] = port
            print(f"[Dashboard] Successfully connected to Arduino on {port}")

            while True:
                if ser.in_waiting > 0:
                    raw = ser.readline().decode("utf-8", errors="ignore")
                    if raw:
                        with lock:
                            app_state["last_seen_serial"] = datetime.now().strftime("%H:%M:%S")
                        process_serial_line(raw)
                time.sleep(0.02)
        except Exception as e:
            with lock:
                app_state["serial_connected"] = False
            time.sleep(2)

# --- Web Server (Bottle) ---
app = Bottle()

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>RFID Attendance Live Dashboard</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.2/css/bootstrap.min.css">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    body { background-color: #f4f6f9; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; }
    .card { border-radius: 10px; border: none; box-shadow: 0 2px 12px rgba(0,0,0,0.06); }
    .stat-card { border-left: 5px solid; transition: transform 0.2s; }
    .stat-card:hover { transform: translateY(-3px); }
    .stat-total { border-color: #0d6efd; }
    .stat-present { border-color: #198754; }
    .stat-absent { border-color: #dc3545; }
    .stat-unknown { border-color: #fd7e14; }
    .badge-present { background-color: #d1e7dd; color: #0f5132; font-weight: 600; padding: 6px 12px; }
    .badge-absent { background-color: #f8d7da; color: #842029; font-weight: 600; padding: 6px 12px; }
    .badge-exited { background-color: #cff4fc; color: #055160; font-weight: 600; padding: 6px 12px; }
    .badge-unknown { background-color: #ffe69c; color: #664d03; font-weight: 600; padding: 6px 12px; }
    .table-excel { background: white; border-radius: 8px; overflow: hidden; }
    .table-excel th { background: #1f497d; color: white; font-weight: 600; }
    .table-excel td { vertical-align: middle; }
    .pulse-dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; background: #198754; animation: pulse 1.5s infinite; }
    @keyframes pulse { 0% { opacity: 1; transform: scale(1); } 50% { opacity: 0.4; transform: scale(1.3); } 100% { opacity: 1; transform: scale(1); } }
  </style>
</head>
<body class="py-4">
<div class="container-fluid px-4">
  <!-- Header -->
  <div class="d-flex justify-content-between align-items-center mb-4">
    <div>
      <h2 class="fw-bold mb-1"><i class="fa-solid fa-id-card-clip text-primary me-2"></i>RFID Attendance Dashboard</h2>
      <p class="text-muted mb-0">Live attendance logging & Excel spreadsheet synchronization</p>
    </div>
    <div class="d-flex align-items-center gap-2">
      <div id="serial-badge" class="badge bg-secondary p-2 d-flex align-items-center gap-2">
        <span class="pulse-dot"></span> <span id="serial-text">Connecting...</span>
      </div>
      <a href="/download/excel" class="btn btn-success"><i class="fa-solid fa-file-excel me-1"></i> Download Excel (.xlsx)</a>
      <a href="/download/csv" class="btn btn-outline-secondary"><i class="fa-solid fa-file-csv me-1"></i> Download CSV</a>
    </div>
  </div>

  <!-- Stats Row -->
  <div class="row g-3 mb-4">
    <div class="col-md-3">
      <div class="card p-3 stat-card stat-total">
        <div class="text-muted small fw-semibold">TOTAL REGISTERED</div>
        <div class="h3 fw-bold my-1" id="stat-total">0</div>
        <div class="small text-muted">Enrolled Students / Staff</div>
      </div>
    </div>
    <div class="col-md-3">
      <div class="card p-3 stat-card stat-present">
        <div class="text-success small fw-semibold">PRESENT TODAY</div>
        <div class="h3 fw-bold text-success my-1" id="stat-present">0</div>
        <div class="small text-muted">Scanned & Marked Present</div>
      </div>
    </div>
    <div class="col-md-3">
      <div class="card p-3 stat-card stat-absent">
        <div class="text-danger small fw-semibold">ABSENT</div>
        <div class="h3 fw-bold text-danger my-1" id="stat-absent">0</div>
        <div class="small text-muted">Not Yet Detected</div>
      </div>
    </div>
    <div class="col-md-3">
      <div class="card p-3 stat-card stat-unknown">
        <div class="text-warning small fw-semibold">UNREGISTERED SCANS</div>
        <div class="h3 fw-bold text-warning my-1" id="stat-unknown">0</div>
        <div class="small text-muted">Unknown / New Cards</div>
      </div>
    </div>
  </div>

  <div class="row g-4">
    <!-- Main Excel Attendance Sheet -->
    <div class="col-lg-8">
      <div class="card p-4">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h5 class="fw-bold mb-0"><i class="fa-solid fa-table me-2 text-primary"></i>Daily Attendance Sheet (Excel Sync)</h5>
          <span class="badge bg-light text-dark border" id="clock-text">--:--:--</span>
        </div>
        <div class="table-responsive">
          <table class="table table-bordered table-excel mb-0">
            <thead>
              <tr>
                <th style="width: 15%;">Roll Number</th>
                <th style="width: 25%;">Student / Staff Name</th>
                <th style="width: 20%;">RFID Card UID</th>
                <th style="width: 13%;">Time In</th>
                <th style="width: 13%;">Time Out</th>
                <th style="width: 14%;">Status</th>
              </tr>
            </thead>
            <tbody id="attendance-tbody">
              <tr><td colspan="6" class="text-center text-muted py-4">Loading attendance records...</td></tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Unregistered Scans Card -->
      <div class="card p-4 mt-4">
        <div class="d-flex justify-content-between align-items-center mb-3">
          <h5 class="fw-bold text-warning mb-0"><i class="fa-solid fa-triangle-exclamation me-2"></i>Unregistered / Unknown RFID Detections</h5>
          <span class="small text-muted">Scans from cards not yet assigned to an employee</span>
        </div>
        <div class="table-responsive">
          <table class="table table-sm table-hover align-middle">
            <thead class="table-light">
              <tr>
                <th>Timestamp</th>
                <th>Detected Card UID</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody id="unknown-tbody">
              <tr><td colspan="3" class="text-center text-muted py-3">No unregistered cards detected today.</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- Right Column: Live Activity Stream & Quick Register -->
    <div class="col-lg-4">
      <!-- Quick Register Card -->
      <div class="card p-4 mb-4">
        <h5 class="fw-bold mb-3"><i class="fa-solid fa-user-plus me-2 text-success"></i>Register New Card</h5>
        <form id="register-form" onsubmit="registerCard(event)">
          <div class="mb-2">
            <label class="form-label small fw-semibold">Card UID</label>
            <input type="text" id="reg-uid" class="form-control form-control-sm" placeholder="e.g. 46e10a2957980" required>
          </div>
          <div class="mb-2">
            <label class="form-label small fw-semibold">Roll / ID Number</label>
            <input type="text" id="reg-roll" class="form-control form-control-sm" placeholder="e.g. HST - 04" required>
          </div>
          <div class="mb-3">
            <label class="form-label small fw-semibold">Student / Staff Name</label>
            <input type="text" id="reg-name" class="form-control form-control-sm" placeholder="e.g. Alex Sharma" required>
          </div>
          <button type="submit" class="btn btn-sm btn-primary w-100"><i class="fa-solid fa-save me-1"></i> Save to System & Excel</button>
        </form>
      </div>

      <!-- Live Stream Log -->
      <div class="card p-4">
        <h5 class="fw-bold mb-3"><i class="fa-solid fa-clock-rotate-left me-2 text-secondary"></i>Live Activity Feed</h5>
        <div id="activity-feed" style="max-height: 400px; overflow-y: auto;">
          <p class="text-muted small text-center py-4">Waiting for RFID card scans...</p>
        </div>
      </div>
    </div>
  </div>
</div>

<script>
  function renderBadge(status) {
    if (status === 'PRESENT' || status === 'STUDENT ENTERED') return '<span class="badge badge-present"><i class="fa-solid fa-check me-1"></i>PRESENT</span>';
    if (status === 'EXITED' || status === 'STUDENT EXITED') return '<span class="badge badge-exited"><i class="fa-solid fa-arrow-right-from-bracket me-1"></i>EXITED</span>';
    if (status === 'ABSENT') return '<span class="badge badge-absent"><i class="fa-solid fa-xmark me-1"></i>ABSENT</span>';
    return '<span class="badge badge-unknown"><i class="fa-solid fa-question me-1"></i>UNKNOWN</span>';
  }

  function quickRegister(uid) {
    document.getElementById('reg-uid').value = uid;
    document.getElementById('reg-name').focus();
  }

  async function registerCard(e) {
    e.preventDefault();
    const uid = document.getElementById('reg-uid').value.trim();
    const roll = document.getElementById('reg-roll').value.trim();
    const name = document.getElementById('reg-name').value.trim();

    try {
      const res = await fetch('/api/register', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({uid, roll, name})
      });
      if (res.ok) {
        document.getElementById('register-form').reset();
        fetchData();
      }
    } catch(err) {
      alert('Error registering card: ' + err);
    }
  }

  async function fetchData() {
    try {
      const res = await fetch('/api/status');
      const data = await res.json();

      // Serial status
      const serialBadge = document.getElementById('serial-badge');
      const serialText = document.getElementById('serial-text');
      if (data.serial_connected) {
        serialBadge.className = 'badge bg-success p-2 d-flex align-items-center gap-2';
        serialText.innerText = 'Connected: ' + data.serial_port.replace('/dev/cu.', '');
      } else {
        serialBadge.className = 'badge bg-danger p-2 d-flex align-items-center gap-2';
        serialText.innerText = 'Disconnected';
      }

      // Stats
      const users = Object.values(data.users);
      const total = users.length;
      const present = users.filter(u => u.status === 'PRESENT' || u.status === 'EXITED').length;
      const absent = total - present;
      const unknownCount = data.unknown_scans.length;

      document.getElementById('stat-total').innerText = total;
      document.getElementById('stat-present').innerText = present;
      document.getElementById('stat-absent').innerText = absent;
      document.getElementById('stat-unknown').innerText = unknownCount;

      // Attendance Table
      const tbody = document.getElementById('attendance-tbody');
      tbody.innerHTML = users.map(u => `
        <tr>
          <td class="fw-bold">${u.roll}</td>
          <td>${u.name}</td>
          <td><code>${u.uid}</code></td>
          <td class="text-center">${u.time_in}</td>
          <td class="text-center">${u.time_out}</td>
          <td class="text-center">${renderBadge(u.status)}</td>
        </tr>
      `).join('');

      // Unknown Scans
      const unkTbody = document.getElementById('unknown-tbody');
      if (data.unknown_scans.length === 0) {
        unkTbody.innerHTML = '<tr><td colspan="3" class="text-center text-muted py-3">No unregistered cards detected today.</td></tr>';
      } else {
        unkTbody.innerHTML = data.unknown_scans.map(unk => `
          <tr>
            <td>${unk.timestamp}</td>
            <td><code class="text-danger fw-bold">${unk.uid}</code></td>
            <td>
              <button class="btn btn-xs btn-outline-primary py-0 px-2" onclick="quickRegister('${unk.uid}')">
                <i class="fa-solid fa-plus me-1"></i>Assign to Student
              </button>
            </td>
          </tr>
        `).join('');
      }

      // Activity Feed
      const feed = document.getElementById('activity-feed');
      if (data.activity_log.length === 0) {
        feed.innerHTML = '<p class="text-muted small text-center py-4">Waiting for RFID card scans...</p>';
      } else {
        feed.innerHTML = data.activity_log.slice(0, 15).map(item => `
          <div class="d-flex justify-content-between align-items-center border-bottom py-2">
            <div>
              <div class="fw-semibold small">${item.name} <span class="text-muted">(${item.roll})</span></div>
              <div class="small text-muted">${item.timestamp} &bull; <code>${item.uid}</code></div>
            </div>
            <div>${renderBadge(item.status)}</div>
          </div>
        `).join('');
      }

      document.getElementById('clock-text').innerText = new Date().toLocaleTimeString();
    } catch(err) {
      console.error(err);
    }
  }

  setInterval(fetchData, 1000);
  fetchData();
</script>
</body>
</html>
"""

@app.route("/")
def index():
    return HTML_TEMPLATE

@app.route("/api/status")
def api_status():
    response.content_type = "application/json"
    with lock:
        return json.dumps({
            "serial_connected": app_state["serial_connected"],
            "serial_port": app_state["serial_port"],
            "users": app_state["users"],
            "unknown_scans": app_state["unknown_scans"],
            "activity_log": app_state["activity_log"]
        })

@app.route("/api/register", method="POST")
def api_register():
    data = request.json
    uid = data.get("uid", "").strip().lower()
    roll = data.get("roll", "").strip()
    name = data.get("name", "").strip()

    if not uid or not name:
        response.status = 400
        return json.dumps({"error": "Missing uid or name"})

    with lock:
        app_state["users"][uid] = {
            "uid": uid,
            "roll": roll,
            "name": name,
            "status": "ABSENT",
            "time_in": "-",
            "time_out": "-",
            "is_inside": False
        }
        # Remove from unknown list if present
        app_state["unknown_scans"] = [u for u in app_state["unknown_scans"] if u["uid"] != uid]
        save_users_config(app_state["users"])

    save_csv()
    save_excel()
    return json.dumps({"success": True})

@app.route("/download/excel")
def download_excel():
    save_excel()
    return static_file("attendance.xlsx", root=DATA_DIR, download="attendance.xlsx")

@app.route("/download/csv")
def download_csv():
    save_csv()
    return static_file("attendance.csv", root=DATA_DIR, download="attendance.csv")

def main():
    print("=" * 60)
    print("  RFID Attendance System - Live Dashboard & Excel Sync")
    print("=" * 60)

    # Initialize users and files
    app_state["users"] = load_users()
    save_csv()
    save_excel()

    print(f"[Info] Enrolled Users: {len(app_state['users'])}")
    for u in app_state["users"].values():
        print(f"  - {u['roll']}: {u['name']} (UID: {u['uid']})")
    print(f"[Info] Excel File: {EXCEL_FILE}")
    print(f"[Info] CSV File:   {CSV_FILE}")
    print(f"[Info] Dashboard running at: http://localhost:{WEB_PORT}")
    print("=" * 60)

    # Start serial worker thread
    t = threading.Thread(target=serial_worker, daemon=True)
    t.start()

    # Start web server
    app.run(host="0.0.0.0", port=WEB_PORT, quiet=True)

if __name__ == "__main__":
    main()

