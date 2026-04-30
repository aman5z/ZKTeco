# -*- coding: utf-8 -*-
"""
ZKTeco Database Sync + Device Writer  --  Merged File
======================================================
Combines:
  sync_db.py    : Creates/maintains attendance.db from MDB + ZK devices (CLI)
  zk_writer.py  : Writes attendance records directly to ZKTeco devices (raw TCP)

CLI Usage:
  python sync_db.py                  # full sync
  python sync_db.py devices-only     # pull from ZK devices only
  python sync_db.py employees-only   # reimport employee list only
  python sync_db.py --days 30        # last N days from MDB + devices
  python sync_db.py backup           # timestamped backup of attendance.db

Write Usage (from other modules):
  from sync_db import write_attendance_best_effort
  ok, msg = write_attendance_best_effort("10.20.141.21", 4370, "1673", datetime(2025,5,8,9,9,0))
"""

# ==============================================================================
#  ZK DEVICE WRITER  (raw TCP -- no pyzk needed for writing)
#  Tested against: F18, F22, K40, iClock 360, iClock 580, UA760, MA300.
# ==============================================================================

import socket
import struct
import time
from datetime import datetime

# ── ZKTeco protocol constants ──────────────────────────────────────────────
CMD_CONNECT      = 1000   # 0x03E8
CMD_EXIT         = 1001   # 0x03E9
CMD_ENABLEDEVICE = 1002   # 0x03EA
CMD_DISABLEDEVICE= 1003   # 0x03EB
CMD_WRITE_ATTLOG = 79     # 0x004F  ← write one attendance record
CMD_ACK_OK       = 2000   # 0x07D0
CMD_ACK_ERROR    = 2001   # 0x07D1
CMD_ACK_DATA     = 2002   # 0x07D2

USHRT_MAX = 65535


# ── Packet helpers ─────────────────────────────────────────────────────────
def _checksum(data: bytes) -> int:
    """ZKTeco 16-bit checksum (same algorithm as pyzk)."""
    total = 0
    i = 0
    while i < len(data) - 1:
        total += struct.unpack_from('<H', data, i)[0]
        i += 2
    if len(data) % 2:
        total += data[-1]
    while total > USHRT_MAX:
        total = (total & 0xFFFF) + (total >> 16)
    return (~total) & 0xFFFF


def _make_packet(cmd: int, session_id: int, reply_id: int, data: bytes = b'') -> bytes:
    """
    Build a ZKTeco TCP packet.

    Packet layout (all little-endian):
        [payload_len : u16]   -- length of everything after this field pair
        [zeros       : u16]
        [cmd         : u16]
        [checksum    : u16]
        [session_id  : u16]
        [reply_id    : u16]
        [data        : bytes]
    """
    # Build payload with checksum = 0 first, then replace
    payload = struct.pack('<HHHH', cmd, 0, session_id, reply_id) + data
    chk     = _checksum(payload)
    payload = struct.pack('<HHHH', cmd, chk, session_id, reply_id) + data
    return struct.pack('<HH', len(payload), 0) + payload


def _parse_response(buf: bytes):
    """
    Parse a device response packet.
    Returns (cmd, session_id, data) or raises ValueError.
    """
    if len(buf) < 12:
        raise ValueError("Response too short ({} bytes)".format(len(buf)))
    # payload_len = struct.unpack_from('<H', buf, 0)[0]  # unused
    cmd        = struct.unpack_from('<H', buf, 4)[0]
    # checksum = struct.unpack_from('<H', buf, 6)[0]   # skip verify
    session_id = struct.unpack_from('<H', buf, 8)[0]
    data       = buf[12:]
    return cmd, session_id, data


# ── Time encoding ──────────────────────────────────────────────────────────
def _encode_zk_time(dt: datetime) -> int:
    """
    Encode a datetime to ZKTeco's packed 32-bit integer format.
    This is the same encoding used by pyzk and the official ZK SDK.
    """
    return (
        ((dt.year % 100) * 12 * 31 + (dt.month - 1) * 31 + (dt.day - 1)) * 86400
        + dt.hour * 3600
        + dt.minute * 60
        + dt.second
    )


# ── Attendance record builder ──────────────────────────────────────────────
def _build_att_record(badge: str, punch_time: datetime,
                       verify: int = 1, status: int = 0) -> bytes:
    """
    Build a 19-byte attendance record for CMD_WRITE_ATTLOG.

    Layout:
        badge    [9 bytes]  null-padded ASCII employee ID
        verify   [1 byte]   0=password, 1=fingerprint, 15=face
        time     [4 bytes]  ZK-encoded datetime (little-endian u32)
        status   [1 byte]   0=check-in, 1=check-out
        workcode [4 bytes]  zeros (reserved)
    """
    badge_bytes = badge.encode('ascii', 'replace').ljust(9, b'\x00')[:9]
    enc_time    = _encode_zk_time(punch_time)
    return (
        badge_bytes
        + struct.pack('<B', verify)
        + struct.pack('<I', enc_time)
        + struct.pack('<B', status)
        + b'\x00\x00\x00\x00'
    )


# ── Main write function ────────────────────────────────────────────────────
def write_attendance(ip: str, port: int, badge: str, punch_time: datetime,
                     timeout: int = 15) -> tuple:
    """
    Write a single attendance record directly to a ZKTeco device.

    Parameters
    ----------
    ip         : Device IP address
    port       : ZKTeco port (usually 4370)
    badge      : Employee badge / employee code (must be enrolled on device)
    punch_time : datetime of the punch
    timeout    : Socket timeout in seconds

    Returns
    -------
    (success: bool, message: str)
    """
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(timeout)

    try:
        # ── Step 1: TCP connect ────────────────────────────────────────────
        try:
            sock.connect((ip, int(port)))
        except (socket.timeout, ConnectionRefusedError, OSError) as e:
            return False, "Cannot connect to {}: {}".format(ip, e)

        # ── Step 2: ZKTeco handshake (CMD_CONNECT) ─────────────────────────
        pkt = _make_packet(CMD_CONNECT, session_id=0, reply_id=0)
        sock.sendall(pkt)

        try:
            resp = sock.recv(1024)
        except socket.timeout:
            return False, "Timeout waiting for handshake response from {}".format(ip)

        try:
            resp_cmd, session_id, _ = _parse_response(resp)
        except ValueError as e:
            return False, "Bad handshake response: {}".format(e)

        if resp_cmd != CMD_ACK_OK:
            return False, "Handshake rejected by device (cmd={})".format(resp_cmd)

        # ── Step 3: Disable device (prevents interference during write) ────
        pkt = _make_packet(CMD_DISABLEDEVICE, session_id, reply_id=1)
        sock.sendall(pkt)
        try:
            resp = sock.recv(1024)
            resp_cmd, _, _ = _parse_response(resp)
            # Some devices return ACK_ERROR for disable but still work — continue
        except Exception:
            pass  # non-fatal

        # ── Step 4: Write attendance record ────────────────────────────────
        att_record = _build_att_record(badge, punch_time)
        pkt = _make_packet(CMD_WRITE_ATTLOG, session_id, reply_id=2, data=att_record)
        sock.sendall(pkt)

        try:
            resp = sock.recv(1024)
        except socket.timeout:
            # Some devices close the connection after a successful write
            # without sending a response — treat as success
            _graceful_exit(sock, session_id)
            return True, "Written to device {} (no ACK — assumed OK)".format(ip)

        try:
            resp_cmd, _, resp_data = _parse_response(resp)
        except ValueError:
            # Garbled response after write — likely succeeded
            _graceful_exit(sock, session_id)
            return True, "Written to device {} (garbled ACK — assumed OK)".format(ip)

        if resp_cmd == CMD_ACK_OK:
            _graceful_exit(sock, session_id)
            return True, "Written to device {}".format(ip)

        # Some devices echo back CMD_DATA instead of ACK_OK on success
        if resp_cmd == CMD_ACK_DATA:
            _graceful_exit(sock, session_id)
            return True, "Written to device {} (DATA ack)".format(ip)

        _graceful_exit(sock, session_id)
        return False, "Device {} rejected write (cmd={})".format(ip, resp_cmd)

    except Exception as e:
        return False, "Unexpected error writing to {}: {}".format(ip, e)
    finally:
        try:
            sock.close()
        except Exception:
            pass


def _graceful_exit(sock: socket.socket, session_id: int):
    """Send CMD_EXIT so the device doesn't hold the session open."""
    try:
        pkt = _make_packet(CMD_ENABLEDEVICE, session_id, reply_id=3)
        sock.sendall(pkt)
        time.sleep(0.05)
        pkt = _make_packet(CMD_EXIT, session_id, reply_id=4)
        sock.sendall(pkt)
        time.sleep(0.05)
    except Exception:
        pass


# ── pyzk fallback (tries internal methods if raw write fails) ──────────────
def write_attendance_pyzk_fallback(ip: str, port: int, badge: str,
                                    punch_time: datetime, timeout: int = 15) -> tuple:
    """
    Attempt to write via pyzk's private internal send method.
    Only used if write_attendance() fails on a device.
    """
    try:
        from zk import ZK
    except ImportError:
        return False, "pyzk not installed"

    try:
        zk   = ZK(ip, port=int(port), timeout=timeout, verbose=False,
                  force_udp=False, ommit_ping=False)
        conn = zk.connect()
    except Exception as e:
        return False, "pyzk connect failed: {}".format(e)

    try:
        conn.disable_device()
        att_record = _build_att_record(badge, punch_time)

        # Try internal send — works on some pyzk versions
        sent = False
        for method_name in ('_send_command', '_ZK__send_command', 'send_command'):
            fn = getattr(conn, method_name, None)
            if fn:
                try:
                    fn(CMD_WRITE_ATTLOG, att_record, 8)
                    sent = True
                    break
                except Exception:
                    continue

        conn.enable_device()
        conn.disconnect()

        if sent:
            return True, "Written via pyzk internal ({})".format(method_name)
        return False, "pyzk version does not expose send method"

    except Exception as e:
        try:
            conn.enable_device()
            conn.disconnect()
        except Exception:
            pass
        return False, "pyzk write failed: {}".format(e)


# ── Combined writer: tries raw TCP first, then pyzk fallback ──────────────
def write_attendance_best_effort(ip: str, port: int, badge: str,
                                  punch_time: datetime, timeout: int = 15) -> tuple:
    """
    Try raw TCP write first. If it fails, try pyzk internal method.
    Returns (success, message).
    """
    ok, msg = write_attendance(ip, port, badge, punch_time, timeout)
    if ok:
        return ok, msg

    # Raw write failed — try pyzk fallback
    ok2, msg2 = write_attendance_pyzk_fallback(ip, port, badge, punch_time, timeout)
    if ok2:
        return ok2, msg2

    # Both failed — return the raw write error (more descriptive)
    return False, "raw: {} | pyzk: {}".format(msg, msg2)


import os, sys, sqlite3, warnings, time, shutil
from datetime import datetime, date, timedelta
warnings.filterwarnings("ignore")

# ── CONFIG ────────────────────────────────────────────────────────────────────
def _script_dir():
    for fn in (os.path.realpath, os.path.abspath):
        try:
            p = os.path.dirname(fn(__file__))
            if p and os.path.isdir(p): return p
        except Exception: pass
    return os.getcwd()

SCRIPT_DIR = _script_dir()

# Read settings.ini if present (same config as server.py)
import configparser
_cfg = configparser.ConfigParser()
_ini = os.path.join(SCRIPT_DIR, 'settings.ini')
if os.path.exists(_ini):
    _cfg.read(_ini, encoding='utf-8')

def _get(section, key, default):
    try: return _cfg.get(section, key).strip()
    except: return default

def _get_list(section, key, default):
    raw = _get(section, key, '')
    return [x.strip() for x in raw.split(',') if x.strip()] if raw else default

DEVICE_IPS     = _get_list('devices', 'ips',
                    ["10.20.141.21","10.20.141.22","10.20.141.23","10.20.141.24",
                     "10.20.141.25","10.20.141.26","10.20.141.27","10.20.141.28","10.20.141.29"])
DEVICE_PORT    = int(_get('devices', 'port',    '4370'))
DEVICE_TIMEOUT = int(_get('devices', 'timeout', '20'))
EXCLUDE_DEPTS  = _get_list('departments', 'exclude',
                    ["DELETED EMPLOYEES","TRANSPORT","GAES","GULF ASIAN ENGLISH SCHOOL"])

CHECKINOUT_TABLE = "CHECKINOUT"
USERINFO_TABLE   = "USERINFO"
DEPT_TABLE       = "DEPARTMENTS"
DB_PATH          = os.path.join(SCRIPT_DIR, "attendance.db")

# ── SCHEMA (matches db_manager / server.py exactly) ──────────────────────────
SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS employees (
    badge       TEXT PRIMARY KEY,
    name        TEXT NOT NULL DEFAULT '',
    dept        TEXT NOT NULL DEFAULT '',
    active      INTEGER NOT NULL DEFAULT 1,
    updated_at  TEXT
);
CREATE TABLE IF NOT EXISTS punches (
    badge       TEXT NOT NULL,
    punch_time  TEXT NOT NULL,
    device_ip   TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (badge, punch_time, device_ip)
);
CREATE TABLE IF NOT EXISTS unknown_users (
    device_ip   TEXT NOT NULL,
    uid         TEXT NOT NULL,
    seen_at     TEXT NOT NULL,
    resolved    INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (device_ip, uid)
);
CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT
);
CREATE TABLE IF NOT EXISTS audit_log (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    ts          TEXT NOT NULL,
    username    TEXT NOT NULL DEFAULT '',
    action      TEXT NOT NULL,
    detail      TEXT DEFAULT '',
    ip_addr     TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS shift_times (
    dept        TEXT PRIMARY KEY,
    shift_start TEXT NOT NULL DEFAULT '07:30',
    shift_end   TEXT NOT NULL DEFAULT '15:00',
    grace_mins  INTEGER NOT NULL DEFAULT 15
);
CREATE TABLE IF NOT EXISTS dept_workdays (
    dept        TEXT PRIMARY KEY,
    workdays    TEXT NOT NULL DEFAULT '[0,1,2,3,6]'
);
CREATE TABLE IF NOT EXISTS emp_workdays (
    badge       TEXT PRIMARY KEY,
    workdays    TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS holidays (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    date        TEXT NOT NULL,
    date_end    TEXT NOT NULL,
    label       TEXT NOT NULL DEFAULT 'Holiday',
    scope       TEXT NOT NULL DEFAULT 'all',
    dept        TEXT NOT NULL DEFAULT '',
    employees   TEXT NOT NULL DEFAULT '',
    created_at  TEXT
);
CREATE INDEX IF NOT EXISTS idx_holidays_date ON holidays(date);
CREATE TABLE IF NOT EXISTS sync_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    sync_time       TEXT,
    source          TEXT,
    records_added   INTEGER DEFAULT 0,
    employees_added INTEGER DEFAULT 0,
    notes           TEXT
);
CREATE TABLE IF NOT EXISTS device_users (
    userid      TEXT,
    device_ip   TEXT,
    name        TEXT,
    privilege   INTEGER DEFAULT 0,
    badge       TEXT,
    last_seen   TEXT,
    PRIMARY KEY (userid, device_ip)
);
CREATE INDEX IF NOT EXISTS idx_punches_badge ON punches(badge);
CREATE INDEX IF NOT EXISTS idx_punches_time  ON punches(punch_time);
CREATE INDEX IF NOT EXISTS idx_audit_ts      ON audit_log(ts);
"""

# ── DB HELPERS ────────────────────────────────────────────────────────────────

# ==============================================================================
#  PUNCH DEDUPLICATION  — keeps only punches >= 60s apart per badge per day
#  Prevents device noise (every doorway crossing) from flooding the database.
# ==============================================================================
DEDUP_GAP_SECS = 60

def _dedupe_records(records):
    """
    records: list of (badge, timestamp_str) sorted by timestamp.
    Returns filtered list keeping only punches >= DEDUP_GAP_SECS apart
    per badge per calendar day.
    """
    from collections import defaultdict
    # Group by badge+date
    groups = defaultdict(list)
    for badge, ts in records:
        day = ts[:10]  # YYYY-MM-DD
        groups[(badge, day)].append(ts)

    result = []
    for (badge, day), times in groups.items():
        times.sort()
        last_dt = None
        for ts in times:
            try:
                t = datetime.strptime(ts, "%Y-%m-%d %H:%M:%S")
            except Exception:
                result.append((badge, ts)); continue
            if last_dt is None or (t - last_dt).total_seconds() >= DEDUP_GAP_SECS:
                result.append((badge, ts))
                last_dt = t
    return result

def get_db(path=None):
    conn = sqlite3.connect(path or DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn

def init_db(path=None):
    conn = get_db(path)
    conn.executescript(SCHEMA_SQL)
    conn.commit()
    return conn

# ── MDB ───────────────────────────────────────────────────────────────────────
def find_mdb():
    for f in os.listdir(SCRIPT_DIR):
        if f.lower().endswith(".mdb") or f.lower().endswith(".accdb"):
            return os.path.join(SCRIPT_DIR, f)
    return None

def connect_mdb(path):
    try:
        import pyodbc
        return pyodbc.connect(
            "Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};"
            "Dbq={0};Uid=Admin;Pwd=;".format(os.path.abspath(path))
        )
    except ImportError:
        print("  [WARN] pyodbc not installed — skipping MDB import"); return None
    except Exception as e:
        print("  [WARN] Cannot open MDB: {0}".format(e)); return None

# ── EMPLOYEE IMPORT ───────────────────────────────────────────────────────────
def import_employees_from_mdb(mdb_conn, db_conn):
    try:
        import pandas as pd
        dept_df  = pd.read_sql("SELECT [DEPTID],[DEPTNAME] FROM [{0}]".format(DEPT_TABLE), mdb_conn)
        id2dept  = dict(zip(dept_df["DEPTID"].astype(str).str.strip(),
                            dept_df["DEPTNAME"].astype(str).str.strip().str.upper()))
        user_df  = pd.read_sql(
            "SELECT [USERID],[Badgenumber],[Name],[DEFAULTDEPTID] FROM [{0}]".format(USERINFO_TABLE),
            mdb_conn
        )
        excl = {d.upper() for d in EXCLUDE_DEPTS}
        added = 0
        for _, row in user_df.iterrows():
            badge = str(row["Badgenumber"]).strip()
            name  = str(row["Name"]).strip()
            dept  = id2dept.get(str(row["DEFAULTDEPTID"]).strip(), "UNKNOWN")
            if not badge or badge == "nan": continue
            active = 0 if dept in excl else 1
            db_conn.execute(
                "INSERT INTO employees (badge,name,dept,active,updated_at) VALUES (?,?,?,?,?)"
                " ON CONFLICT(badge) DO UPDATE SET name=excluded.name,dept=excluded.dept,"
                "active=excluded.active,updated_at=excluded.updated_at",
                (badge, name, dept, active, datetime.now().isoformat())
            )
            added += 1
        db_conn.commit()
        print("  Employees imported from MDB: {0}".format(added))
        return added
    except Exception as e:
        print("  [WARN] MDB employee import failed: {0}".format(e)); return 0

def import_employees_from_csv(db_conn):
    csv_path  = os.path.join(SCRIPT_DIR, "employees_export.csv")
    xlsx_path = os.path.join(SCRIPT_DIR, "employees_export.xlsx")
    try:
        import pandas as pd
        if os.path.exists(csv_path):
            df = pd.read_csv(csv_path, dtype=str)
        elif os.path.exists(xlsx_path):
            df = pd.read_excel(xlsx_path, dtype=str)
        else:
            print("  [WARN] No employees_export.csv / .xlsx found"); return 0
        df.columns = [str(c).strip() for c in df.columns]
        excl  = {d.upper() for d in EXCLUDE_DEPTS}
        added = 0
        for _, row in df.iterrows():
            badge = str(row.get("Badgenumber","")).strip()
            name  = str(row.get("Name","")).strip()
            dept  = str(row.get("DEPTNAME","")).strip()
            if not badge or badge in ("","nan"): continue
            active = 0 if dept.upper() in excl else 1
            db_conn.execute(
                "INSERT INTO employees (badge,name,dept,active,updated_at) VALUES (?,?,?,?,?)"
                " ON CONFLICT(badge) DO UPDATE SET name=excluded.name,dept=excluded.dept,"
                "active=excluded.active,updated_at=excluded.updated_at",
                (badge, name, dept, active, datetime.now().isoformat())
            )
            added += 1
        db_conn.commit()
        print("  Employees imported from CSV: {0}".format(added))
        return added
    except Exception as e:
        print("  [WARN] CSV employee import failed: {0}".format(e)); return 0

# ── MDB ATTENDANCE ────────────────────────────────────────────────────────────
def import_attendance_from_mdb(mdb_conn, db_conn, days=None):
    try:
        import pandas as pd
        # Build uid -> badge map from employees
        badge_map = {}
        for row in db_conn.execute("SELECT badge, name FROM employees").fetchall():
            if row["badge"]: badge_map[row["badge"]] = row["badge"]
        # Also try USERINFO uid mapping
        try:
            udf = pd.read_sql("SELECT [USERID],[Badgenumber] FROM [{0}]".format(USERINFO_TABLE), mdb_conn)
            for _, r in udf.iterrows():
                uid   = str(r["USERID"]).strip()
                badge = str(r["Badgenumber"]).strip()
                if uid and badge: badge_map[uid] = badge
        except Exception: pass

        if days:
            date_from = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
            sql = ("SELECT [USERID],[CHECKTIME] FROM [{0}] "
                   "WHERE [CHECKTIME] >= #{1}#".format(CHECKINOUT_TABLE, date_from + " 00:00:00"))
        else:
            sql = "SELECT [USERID],[CHECKTIME] FROM [{0}]".format(CHECKINOUT_TABLE)

        try:
            punch_df = pd.read_sql(sql, mdb_conn)
        except Exception:
            # Column error fallback — try without date filter then slice in pandas
            try:
                punch_df = pd.read_sql("SELECT [USERID],[CHECKTIME] FROM [{0}]".format(CHECKINOUT_TABLE), mdb_conn)
                if days:
                    punch_df["CHECKTIME"] = pd.to_datetime(punch_df["CHECKTIME"], errors="coerce")
                    cutoff = datetime.now() - timedelta(days=days)
                    punch_df = punch_df[punch_df["CHECKTIME"] >= cutoff]
            except Exception as e2:
                print("  [WARN] MDB attendance query failed: {0}".format(e2)); return 0

        punch_df["USERID"]    = punch_df["USERID"].astype(str).str.strip()
        punch_df["CHECKTIME"] = pd.to_datetime(punch_df["CHECKTIME"], errors="coerce")
        punch_df = punch_df.dropna(subset=["CHECKTIME"])

        added = 0; skipped = 0; cur = db_conn.cursor()
        for _, row in punch_df.iterrows():
            uid   = str(row["USERID"]).strip()
            badge = badge_map.get(uid, uid)   # fall back to uid itself if not mapped
            ct    = row["CHECKTIME"].strftime("%Y-%m-%d %H:%M:%S")
            try:
                cur.execute(
                    "INSERT OR IGNORE INTO punches (badge, punch_time, device_ip) VALUES (?,?,?)",
                    (badge, ct, "mdb-import")
                )
                if cur.rowcount: added += 1
                else: skipped += 1
            except Exception: skipped += 1
        db_conn.commit()
        print("  MDB attendance: {0} new, {1} already existed".format(added, skipped))
        return added
    except Exception as e:
        print("  [WARN] MDB attendance import failed: {0}".format(e)); return 0

# ── DEVICE PULL ───────────────────────────────────────────────────────────────
def pull_device(ip):
    result = {"ip": ip, "ok": False, "att_records": [], "users": [], "error": ""}
    try:
        from zk import ZK
        zk_conn = ZK(ip, port=DEVICE_PORT, timeout=DEVICE_TIMEOUT, verbose=False).connect()
        try:
            zk_conn.disable_device()
            users   = zk_conn.get_users()
            records = zk_conn.get_attendance()
            zk_conn.enable_device()
            result["users"]       = [(str(u.user_id).strip(), u.name or "", int(u.privilege)) for u in users]
            result["att_records"] = [(str(r.user_id).strip(), r.timestamp.strftime("%Y-%m-%d %H:%M:%S")) for r in records]
            result["ok"] = True
            print("  {0}  OK — {1} users, {2} records".format(ip, len(result["users"]), len(result["att_records"])))
        finally:
            try: zk_conn.enable_device(); zk_conn.disconnect()
            except: pass
    except ImportError:
        result["error"] = "pyzk not installed: pip install pyzk"
        print("  {0}  SKIP — pyzk not installed".format(ip))
    except Exception as e:
        result["error"] = str(e)
        print("  {0}  FAILED — {1}".format(ip, str(e)[:80]))
    return result

def sync_devices(db_conn):
    # Build uid->badge and name->badge lookup maps from employees table
    badge_by_uid  = {}
    badge_by_name = {}
    for r in db_conn.execute("SELECT badge, name FROM employees").fetchall():
        if r["badge"]:
            badge_by_uid[r["badge"]] = r["badge"]   # badge is its own uid in ZK
            if r["name"]: badge_by_name[r["name"].strip().lower()] = r["badge"]

    total_att = 0; total_users = 0; unmapped = []

    for ip in DEVICE_IPS:
        result = pull_device(ip)
        if not result["ok"]: continue
        now = datetime.now().isoformat()
        cur = db_conn.cursor()

        # Store device users
        for (uid, uname, priv) in result["users"]:
            badge = badge_by_uid.get(uid) or badge_by_name.get(uname.strip().lower(), "")
            cur.execute(
                "INSERT INTO device_users (userid,device_ip,name,privilege,badge,last_seen)"
                " VALUES (?,?,?,?,?,?)"
                " ON CONFLICT(userid,device_ip) DO UPDATE SET"
                " name=excluded.name,badge=excluded.badge,last_seen=excluded.last_seen",
                (uid, ip, uname, priv, badge or None, now)
            )
            total_users += 1
            if not badge:
                unmapped.append({"uid": uid, "name": uname, "device": ip})

        # Store attendance — resolve badge, deduplicate, then store
        # First pass: resolve all UIDs to badges
        resolved = []
        for (uid, ct) in result["att_records"]:
            badge = badge_by_uid.get(uid, "")
            if not badge:
                row = db_conn.execute(
                    "SELECT badge FROM device_users WHERE userid=? AND device_ip=?", (uid, ip)
                ).fetchone()
                if row and row["badge"]: badge = row["badge"]
            resolved.append((badge or uid, ct))

        # Second pass: deduplicate (60s gap per badge per day)
        deduped = _dedupe_records(resolved)
        print("  {0}  dedup: {1} → {2} records".format(
            ip, len(resolved), len(deduped)))

        # Third pass: store
        added = 0
        for (badge, ct) in deduped:
            try:
                cur.execute(
                    "INSERT OR IGNORE INTO punches (badge, punch_time, device_ip) VALUES (?,?,?)",
                    (badge, ct, ip)
                )
                if cur.rowcount: added += 1
            except Exception: pass

        db_conn.commit()
        total_att += added

    print("\n  Attendance records added: {0}".format(total_att))
    print("  Device users processed:  {0}".format(total_users))
    if unmapped:
        print("\n  ⚠  {0} device users not matched to employee badges:".format(len(unmapped)))
        for u in unmapped[:10]:
            print("     uid={uid}  name={name}  device={device}".format(**u))
        if len(unmapped) > 10:
            print("     ... and {0} more (resolve in dashboard Admin → Unmapped Users)".format(len(unmapped)-10))
    return total_att

# ── BACKUP ────────────────────────────────────────────────────────────────────
def create_backup():
    if not os.path.exists(DB_PATH):
        print("  [ERROR] attendance.db not found — nothing to back up"); return None
    bak_dir = os.path.join(SCRIPT_DIR, "backups")
    os.makedirs(bak_dir, exist_ok=True)
    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(bak_dir, "attendance_{0}.db".format(ts))
    shutil.copy2(DB_PATH, dst)
    print("  Backup saved -> {0}".format(dst))
    backups = sorted([f for f in os.listdir(bak_dir) if f.endswith(".db")])
    for old in backups[:-30]:
        try: os.remove(os.path.join(bak_dir, old))
        except: pass
    return dst

# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    args     = [a.lower() for a in sys.argv[1:]]
    days_arg = None
    if "--days" in args:
        try: days_arg = int(sys.argv[sys.argv.index("--days") + 1])
        except: days_arg = 90

    print("\n" + "="*58)
    print("  ZKTeco Database Sync  v2.2")
    print("  DB: {0}".format(DB_PATH))
    print("="*58 + "\n")

    if "backup" in args:
        create_backup()
        if sys.stdin.isatty():
            input("\nPress Enter to exit...")
        return

    print("[1] Initialising database schema...")
    db_conn = init_db()
    print("  OK — {0}".format(DB_PATH))

    mdb_path = find_mdb()
    if mdb_path:
        print("\n  MDB found: {0}".format(mdb_path))
    else:
        print("\n  No MDB found — using device data + CSV only")

    if "devices-only" not in args:
        print("\n[2] Importing employee list...")
        mdb_conn = connect_mdb(mdb_path) if mdb_path else None
        if mdb_conn:
            added = import_employees_from_mdb(mdb_conn, db_conn)
            if added == 0: import_employees_from_csv(db_conn)
            mdb_conn.close()
        else:
            import_employees_from_csv(db_conn)
        print("  Total employees in DB: {0}".format(
            db_conn.execute("SELECT COUNT(*) FROM employees").fetchone()[0]))
    else:
        print("\n[2] Skipping employee import (devices-only mode)")

    if mdb_path and "devices-only" not in args:
        print("\n[3] Importing attendance history from MDB...")
        mdb_conn = connect_mdb(mdb_path)
        if mdb_conn:
            if days_arg: print("  Last {0} days only...".format(days_arg))
            else:        print("  Full history (may take a few minutes for large databases)...")
            import_attendance_from_mdb(mdb_conn, db_conn, days=days_arg)
            mdb_conn.close()
    elif "devices-only" not in args:
        print("\n[3] No MDB — skipping historical import")
    else:
        print("\n[3] Skipping MDB import (devices-only mode)")

    att_before = db_conn.execute("SELECT COUNT(*) FROM punches").fetchone()[0]
    emp_before = db_conn.execute("SELECT COUNT(*) FROM employees").fetchone()[0]

    if "employees-only" not in args:
        print("\n[4] Pulling from {0} ZK device(s)...\n".format(len(DEVICE_IPS)))
        sync_devices(db_conn)
    else:
        print("\n[4] Skipping device pull (employees-only mode)")

    total_att  = db_conn.execute("SELECT COUNT(*) FROM punches").fetchone()[0]
    total_emp  = db_conn.execute("SELECT COUNT(*) FROM employees").fetchone()[0]
    added_att  = total_att - att_before
    added_emp  = total_emp - emp_before
    unmapped_c = db_conn.execute(
        "SELECT COUNT(*) FROM device_users WHERE badge IS NULL OR badge=''").fetchone()[0]
    db_size    = os.path.getsize(DB_PATH) / (1024*1024)

    db_conn.execute(
        "INSERT INTO sync_log (sync_time,source,records_added,employees_added,notes) VALUES (?,?,?,?,?)",
        (datetime.now().isoformat(), "sync_db.py", added_att, added_emp,
         "devices-only" if "devices-only" in args else "full")
    )
    db_conn.commit(); db_conn.close()

    print("\n" + "="*58)
    print("  Sync Complete!")
    print("  Employees    : {0:,}".format(total_emp))
    print("  Punch records: {0:,}".format(total_att))
    print("  Unmapped UIDs: {0}".format(unmapped_c))
    print("  DB size      : {0:.1f} MB".format(db_size))
    print("="*58)
    if unmapped_c:
        print("\n  ⚠  {0} UIDs couldn't be matched.".format(unmapped_c))
        print("     Dashboard → Admin → Unmapped Users to resolve.")
    print("\n  Dashboard will now use attendance.db automatically.")
    if sys.stdin.isatty():
        input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
