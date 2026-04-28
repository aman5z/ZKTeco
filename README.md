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

**Author:** Hashiq V H

