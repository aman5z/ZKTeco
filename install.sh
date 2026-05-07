#!/bin/bash
# ============================================================
#  ZKTeco Attendance Dashboard  --  Installer
#  Supports: Ubuntu/Debian, macOS (Homebrew), Termux (Android)
#  Run once: bash install.sh
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# ── Detect platform ───────────────────────────────────────────
detect_platform() {
    if [ -n "${TERMUX_VERSION:-}" ] || echo "${PREFIX:-}" | grep -q "com.termux" 2>/dev/null; then
        echo "termux"
    elif [ "$(uname -s 2>/dev/null)" = "Darwin" ]; then
        echo "macos"
    else
        echo "linux"
    fi
}
PLATFORM=$(detect_platform)

# Python command: Termux and macOS may use 'python' or 'python3'
if command -v python3 &>/dev/null; then
    PYTHON="python3"
elif command -v python &>/dev/null; then
    PYTHON="python"
else
    PYTHON="python3"
fi

echo ""
echo "  ============================================================"
echo "   ZKTeco Attendance Dashboard  --  Setup"
echo "   Platform : $PLATFORM"
echo "  ============================================================"
echo ""

# ── System packages ──────────────────────────────────────────
echo "  [1/5] Installing system packages..."
case "$PLATFORM" in
    termux)
        pkg install -y python sqlite mdbtools 2>/dev/null || \
            pkg install -y python sqlite
        ;;
    macos)
        if command -v brew &>/dev/null; then
            brew install mdbtools sqlite 2>/dev/null || true
        else
            echo "  NOTE: Homebrew not found. Install from https://brew.sh"
            echo "        Then run: brew install mdbtools"
        fi
        ;;
    *)
        sudo apt-get update -qq
        sudo apt-get install -y python3 python3-pip python3-venv sqlite3 mdbtools
        ;;
esac
echo "  OK: system packages"

# ── Python packages ──────────────────────────────────────────
echo ""
echo "  [2/5] Installing Python packages..."
case "$PLATFORM" in
    termux)
        # Termux manages its own Python — no --break-system-packages needed
        pip install flask pandas openpyxl pyzk werkzeug reportlab \
            flask-socketio simple-websocket httpx --quiet
        ;;
    macos)
        pip3 install flask pandas openpyxl pyzk werkzeug reportlab \
            flask-socketio simple-websocket httpx --quiet 2>/dev/null || \
            pip install flask pandas openpyxl pyzk werkzeug reportlab \
            flask-socketio simple-websocket httpx --quiet
        ;;
    *)
        pip install flask pandas openpyxl pyzk werkzeug reportlab \
            flask-socketio simple-websocket httpx \
            --break-system-packages \
            --ignore-installed \
            --quiet
        ;;
esac
echo "  OK: Python packages installed"

# ── Verify ───────────────────────────────────────────────────
echo ""
echo "  [3/5] Verifying installations..."
$PYTHON -c "import flask"           && echo "    flask          OK" || echo "    flask          MISSING"
$PYTHON -c "import pandas"          && echo "    pandas         OK" || echo "    pandas         MISSING"
$PYTHON -c "import openpyxl"        && echo "    openpyxl       OK" || echo "    openpyxl       MISSING"
$PYTHON -c "import zk"              && echo "    pyzk           OK" || echo "    pyzk           MISSING"
$PYTHON -c "import werkzeug"        && echo "    werkzeug       OK" || echo "    werkzeug       MISSING"
$PYTHON -c "import reportlab"       && echo "    reportlab      OK" || echo "    reportlab      MISSING"
$PYTHON -c "import flask_socketio"  && echo "    flask-socketio OK" || echo "    flask-socketio MISSING"
$PYTHON -c "import simple_websocket" && echo "    simple-websocket OK" || echo "    simple-websocket MISSING"
$PYTHON -c "import httpx"           && echo "    httpx          OK" || echo "    httpx          MISSING"
command -v mdb-export &>/dev/null && echo "    mdbtools       OK" || echo "    mdbtools       MISSING (MDB features unavailable)"
command -v sqlite3    &>/dev/null && echo "    sqlite3        OK" || echo "    sqlite3        MISSING"

# ── Employee export (if MDB present) ─────────────────────────
echo ""
echo "  [4/5] Checking for MDB database..."
MDB_FILE=$(find "$SCRIPT_DIR" -maxdepth 1 \( -name "*.mdb" -o -name "*.accdb" \) 2>/dev/null | head -1)
if [ -n "$MDB_FILE" ]; then
    echo "  Found: $MDB_FILE"
    echo "  Exporting employees..."
    $PYTHON mdb_tools.py export && echo "  OK: employees exported" || echo "  WARN: export failed"
else
    echo "  No MDB found — skipping export"
    echo "  Place your .mdb file here and run: $PYTHON mdb_tools.py export"
fi

# ── Device sync ───────────────────────────────────────────────
echo ""
echo "  [5/5] Syncing from devices (devices-only — skips MDB import)..."
$PYTHON sync_db.py devices-only || echo "  WARN: device sync failed (check network/devices are reachable)"

# ── Get local IP for display ──────────────────────────────────
get_ip() {
    case "$PLATFORM" in
        termux)
            ip addr show wlan0 2>/dev/null | grep 'inet ' | awk '{print $2}' | cut -d/ -f1 | head -1 \
                || ip addr show wlan1 2>/dev/null | grep 'inet ' | awk '{print $2}' | cut -d/ -f1 | head -1 \
                || echo "127.0.0.1"
            ;;
        macos)
            ipconfig getifaddr en0 2>/dev/null \
                || ipconfig getifaddr en1 2>/dev/null \
                || echo "127.0.0.1"
            ;;
        *)
            hostname -I 2>/dev/null | awk '{print $1}' || echo "127.0.0.1"
            ;;
    esac
}

# ── Done ─────────────────────────────────────────────────────
echo ""
echo "  ============================================================"
echo "   Setup complete!"
echo ""
echo "   To start the dashboard:"
echo "     bash attendance.sh         (interactive menu)"
echo "     $PYTHON server.py          (direct start)"
echo ""
if [ "$PLATFORM" = "termux" ]; then
    echo "   To auto-start on Android boot:"
    echo "     Install Termux:Boot, then copy attendance.sh to"
    echo "     ~/.termux/boot/start-attendance.sh"
elif [ "$PLATFORM" = "linux" ]; then
    echo "   To auto-start on boot:"
    echo "     bash attendance.sh  →  option [S]"
fi
echo ""
echo "   Access at: http://$(get_ip):5000"
echo "  ============================================================"
echo ""
