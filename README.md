# ZKTeco Attendance System

A full attendance management solution for ZKTeco biometric devices. It includes a real-time logger that streams punch events to an API, a web dashboard for viewing attendance and reports, bulk sync utilities, and Telegram notifications.

---

## Components

| File | Purpose |
|------|---------|
| `main.py` | Real-time attendance logger — streams live punch events from all devices to the API endpoint |
| `server.py` | Flask web dashboard — view today's/historical attendance, employee list, reports, device management |
| `sync_all.py` | Bulk sync CLI — pull all logs from devices and push to API (supports date range filtering) |
| `sync_db.py` | Database sync CLI — imports employees from MDB, pulls device logs into local SQLite |
| `boot_sync_60d.py` | Startup sync — syncs the last 60 days of logs on system boot via `sync_all.py` |
| `mdb_tools.py` | MDB/Access tools — export employees, generate absent reports from MDB backup |
| `telegram_notifier.py` | Telegram bot helper — sends system status notifications |

---

## Features

- Live real-time capture from multiple ZKTeco devices (self-healing reconnect)
- Buffered push to API endpoint with configurable batch size
- Web dashboard with login, role-based access, department view, calendar, reports
- Export absent reports to Excel/PDF
- Telegram notifications for startup, data push, errors
- Windows (`.bat`) and Linux (`.sh`) launchers with interactive menu
- Systemd service installer for auto-start on boot (Linux)
- Docker support

---

## Requirements

- **Python** 3.8 or later
- Dependencies listed in `requirements.txt`
- Linux dashboard: `mdbtools` system package (for MDB file reading)

---

## Quick Start

### Linux

```bash
bash install.sh       # installs packages, exports employees, syncs devices
bash attendance.sh    # interactive menu
```

### Windows

```bat
attendance.bat        # interactive menu
```

### Manual

```bash
python3 -m venv venv
source venv/bin/activate        # Linux/macOS
# venv\Scripts\activate         # Windows

pip install -r requirements.txt

# Real-time logger (pushes live events to API)
python main.py

# Web dashboard
python server.py      # open http://localhost:5000
```

---

## Configuration

### `config.json` — Real-time logger & boot sync

```json
{
  "name": "Your-System-Name",
  "log_level": "INFO",
  "endpoint": "https://your-domain.example.com/erp-api/sync/empAttSync.php",
  "buffer_limit": 10,
  "devices": [
    { "device_id": 1, "ip_address": "10.20.141.21", "port": 4370 },
    { "device_id": 2, "ip_address": "10.20.141.22", "port": 4370 }
  ],
  "telegram": {
    "bot_token": "YOUR_TELEGRAM_BOT_TOKEN",
    "chat_id": "YOUR_TELEGRAM_CHAT_ID",
    "enabled": false,
    "notifications": {
      "startup": true,
      "end_of_day": true,
      "data_push": true,
      "errors": true,
      "device_status": true
    }
  }
}
```

### `settings.ini` — Web dashboard

Edit `settings.ini` to configure device IPs, departments, session timeouts, and the dashboard port.

---

## Utilities

```bash
# Sync all logs for a date range
python sync_all.py --from 2026-01-01 --to 2026-04-01

# Sync a specific device only
python sync_all.py --device-id 2

# Import employees + pull device logs into local DB
python sync_db.py

# Export employee list from MDB
python mdb_tools.py export

# Generate today's absent report
python mdb_tools.py today
```

---

## Docker

```bash
docker build -t zkteco-attendance .
docker run --network host zkteco-attendance
```

---

## License

This project is licensed under the MIT License.

---

## 🏢 ERP Console (Web Dashboard)

A unified, high-performance web dashboard merging the ZKTeco Biometric Attendance System with an IT Helpdesk and AD Console. Built entirely with Vanilla HTML, CSS, and JavaScript — no build step required.

<img src="https://github.com/user-attachments/assets/6c291473-75cd-4f0a-b634-afcc75cbdc84">

### Overview

The ERP Console connects to two primary backends:
1. **ZKTeco Local Server** — Python/Flask API handling local biometric hardware, employee databases, and attendance logs.
2. **Google Apps Script (GAS)** — Cloud backend managing IT Helpdesk tickets, token queues, Drive storage, and admin authentication.

It features full PWA (Progressive Web App) support, offline capabilities using IndexedDB, and a built-in suite of IT networking and utility tools.

---

### 👥 Human Resources & Attendance

- **Live Dashboard:** Real-time stats for present/absent employees, offline tracking, and department attendance percentages.
- **Attendance Management:** View today's attendance, historical date-range reports, and individual employee calendars with shift and grace-period logic.
- **Remote Punching:** Allow employees to punch in/out from their devices. Includes offline support (saves to IndexedDB and syncs when back online).
- **Leave & Shift Admin:** Configure department-specific shifts, workdays, and system-wide holidays.

<img src="https://github.com/user-attachments/assets/a3e2ae7e-5724-43a8-b68b-9e6c7fc3c6b5">

---

### 📡 Hardware Management (ZKTeco)

- **Device Fleet Control:** Monitor online/offline status for all connected ZKTeco terminals.
- **Remote Actions:** Ping devices, sync clocks, reboot, and view or download enrolled users and raw attendance logs directly from the dashboard.

<img src="https://github.com/user-attachments/assets/f0d928be-4d1b-4734-8e00-04b0df7f1097">

---

### 🎫 IT Services & Collaboration

- **Helpdesk Ticketing:** Create, assign, and track IT support tickets with priority levels and due dates.
- **Token Queue:** Manage digital service counters (e.g., HR, Finance, Medical) with visual queue tracking.
- **Internal Chat & Notes:** Peer-to-peer messaging system and private, color-coded sticky notes.
- **Google Drive Integration:** Browse and upload files directly to a designated Google Drive folder.
- **Linode Terminal:** Integrated web shell connecting to your server via `ttyd` and an nginx proxy.

<img src="https://github.com/user-attachments/assets/33247b6e-dec4-4f83-9b63-8d9d577368ae">
<img src="https://github.com/user-attachments/assets/5ee74551-b9e5-44d4-906d-88317393e6c4">

---

### 🧰 Built-in IT Toolkit

A comprehensive suite of client-side utilities built right into the app:

- **Email Batch Manager:** Parse `.xlsx` files, extract, deduplicate, and batch emails for mass communication.
- **Generators:** QR codes, Code128 Barcodes, and secure Passwords.
- **PDF Tools:** Merge, split, compress, and watermark PDFs entirely in the browser using `pdf-lib`.
- **Networking:** IP/DNS Lookup, Subnet Calculator, MAC Vendor lookup, and live Ping Monitor charts.
- **Data Formatting:** Base64 encoding/decoding, Hex/Dec/Bin converters, and a Text Diff checker.

<img src="https://github.com/user-attachments/assets/06b9f97f-e4b6-424f-9723-44fbeb9f35b8">
<img src="https://github.com/user-attachments/assets/2cf96f77-d44d-49b8-b436-f5771df9ec70">
<img src="https://github.com/user-attachments/assets/1badf3c7-753e-4fdd-91a5-4cfb99e131c8">

---

### 🛠️ Tech Stack

| Layer | Details |
|-------|---------|
| **Frontend** | HTML5, CSS3 (Native CSS Variables), Vanilla JavaScript (ES6+) |
| **Offline Storage** | `localStorage` for settings; `IndexedDB` & Service Workers for offline punch caching |
| **External Libraries (CDN)** | `JsBarcode`, `qrcode.js`, `pdf-lib`, `SheetJS / xlsx`, `Chart.js` (custom Canvas charting) |

<img src="https://github.com/user-attachments/assets/2e95a033-1456-4a70-ba1b-cba16c7ebd58">

---

### 🎮 Demo Mode

Don't have the backends set up yet? Click **"🎮 Enter Demo Mode"** on the login screen to explore the UI with mock employees, tickets, devices, and attendance data.

---

### 🔒 Security Note

The ERP Console relies on its connected backend services for true authorization. The built-in Role-Based Access Control (RBAC) hides UI elements based on the user's role, but backend APIs (Flask/GAS) must independently validate session tokens for all sensitive read/write operations.

---

**Author:** Hashiq V H

