#!/bin/bash
# ============================================================
#  ZKTeco Attendance Dashboard  v2.2  --  Linux Launcher
#  Usage: bash attendance.sh
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# Activate venv if present, else use system python3
if [ -f "venv/bin/activate" ]; then
    source venv/bin/activate
    PYTHON="python3"
else
    PYTHON="python3"
fi

# ── Find local IP ─────────────────────────────────────────────
get_ip() {
    hostname -I 2>/dev/null | awk '{print $1}' || echo "127.0.0.1"
}

menu() {
    clear
    echo ""
    echo "  ============================================================"
    echo "   ZKTeco Attendance Dashboard  v2.2"
    echo "  ============================================================"
    echo ""
    echo "    [1]  Open Dashboard   (start server)"
    echo "    [2]  Sync Database    (full: employees + MDB + devices)"
    echo "    [3]  Sync Devices Only"
    echo "    [4]  Today's Absent Report  (Excel)"
    echo "    [5]  History Absent Report  (Excel, from MDB)"
    echo "    [6]  Export Employees from MDB"
    echo "    [7]  Scan MDB         (diagnostics)"
    echo "    [8]  Backup Database"
    echo "    [9]  Install/Update packages"
    echo "    [S]  Install as systemd service (auto-start on boot)"
    echo "    [X]  Remove systemd service"
    echo "    [0]  Exit"
    echo ""
    read -rp "  Choice: " CHOICE
    echo ""

    case "$CHOICE" in
        1) do_dashboard ;;
        2) do_sync_full ;;
        3) do_sync_devices ;;
        4) do_today ;;
        5) do_history ;;
        6) do_export ;;
        7) do_scan ;;
        8) do_backup ;;
        9) do_setup ;;
        [Ss]) do_install_service ;;
        [Xx]) do_remove_service ;;
        0) echo "  Goodbye."; exit 0 ;;
        *) echo "  Invalid choice." ; sleep 1 ; menu ;;
    esac
}

after_action() {
    echo ""
    echo "  ────────────────────────────────────────"
    read -rp "  Press Enter to return to menu..." _
    menu
}

# ── Actions ───────────────────────────────────────────────────

do_dashboard() {
    echo "  Starting server..."
    echo "  URL : http://localhost:5000/d"
    echo "  Keep this window open. Ctrl+C to stop."
    echo "  ============================================================"
    # Open browser in background (best-effort — may not work on headless servers)
    (sleep 3 && xdg-open "http://localhost:5000/d" 2>/dev/null || true) &
    $PYTHON server.py
    after_action
}

do_sync_full() {
    echo "  Full sync (employees + MDB history + devices)..."
    echo ""
    $PYTHON sync_db.py
    after_action
}

do_sync_devices() {
    echo "  Device-only sync..."
    echo ""
    $PYTHON sync_db.py devices-only
    after_action
}

do_today() {
    echo "  Today's Absent Report..."
    echo ""
    $PYTHON mdb_tools.py today
    after_action
}

do_history() {
    echo ""
    # Auto-detect MDB
    MDB_FILE=$(find "$SCRIPT_DIR" -maxdepth 1 \( -name "*.mdb" -o -name "*.accdb" \) 2>/dev/null | head -1)
    if [ -n "$MDB_FILE" ]; then
        echo "  MDB detected: $MDB_FILE"
    else
        read -rp "  Path to .mdb file: " MDB_FILE
    fi
    read -rp "  From (DD/MM/YYYY): " DATE_FROM
    read -rp "    To (DD/MM/YYYY): " DATE_TO
    echo ""
    $PYTHON mdb_tools.py history "$MDB_FILE" "$DATE_FROM" "$DATE_TO"
    after_action
}

do_export() {
    echo "  Exporting employees from MDB..."
    echo ""
    $PYTHON mdb_tools.py export
    after_action
}

do_scan() {
    echo ""
    MDB_FILE=$(find "$SCRIPT_DIR" -maxdepth 1 \( -name "*.mdb" -o -name "*.accdb" \) 2>/dev/null | head -1)
    if [ -n "$MDB_FILE" ]; then
        echo "  Scanning: $MDB_FILE"
        $PYTHON mdb_tools.py scan "$MDB_FILE"
    else
        read -rp "  Path to .mdb file: " MDB_PATH
        $PYTHON mdb_tools.py scan "$MDB_PATH"
    fi
    after_action
}

do_backup() {
    echo "  Backing up attendance.db..."
    $PYTHON sync_db.py backup
    after_action
}

do_setup() {
    echo "  Installing/updating packages..."
    echo ""
    sudo apt-get install -y mdbtools --quiet
    pip install --upgrade flask pandas openpyxl pyzk werkzeug --quiet
    echo "  OK: packages updated"
    after_action
}

do_install_service() {
    echo "  Installing systemd service..."
    local SERVICE_FILE="/etc/systemd/system/attendance-dashboard.service"
    local VENV_PYTHON="$SCRIPT_DIR/venv/bin/python3"
    local RUN_PYTHON
    RUN_PYTHON=$([ -f "$VENV_PYTHON" ] && echo "$VENV_PYTHON" || echo "$(which python3)")

    sudo tee "$SERVICE_FILE" > /dev/null << EOF
[Unit]
Description=ZKTeco Attendance Dashboard
After=network.target

[Service]
Type=simple
User=$(whoami)
WorkingDirectory=$SCRIPT_DIR
ExecStart=$RUN_PYTHON $SCRIPT_DIR/server.py
Restart=on-failure
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

    sudo systemctl daemon-reload
    sudo systemctl enable attendance-dashboard
    sudo systemctl restart attendance-dashboard

    echo ""
    echo "  Service installed and started!"
    echo "  Dashboard: http://$(get_ip):5000"
    echo ""
    echo "  Commands:"
    echo "    sudo systemctl status attendance-dashboard"
    echo "    sudo systemctl stop attendance-dashboard"
    echo "    sudo journalctl -u attendance-dashboard -f   (live logs)"
    after_action
}

do_remove_service() {
    echo "  Removing systemd service..."
    sudo systemctl stop attendance-dashboard 2>/dev/null || true
    sudo systemctl disable attendance-dashboard 2>/dev/null || true
    sudo rm -f /etc/systemd/system/attendance-dashboard.service
    sudo systemctl daemon-reload
    echo "  OK: service removed"
    after_action
}

# ── Entry point ───────────────────────────────────────────────
menu
