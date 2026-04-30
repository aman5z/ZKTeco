# -*- coding: utf-8 -*-
"""
Telegram Notifier for ZKTeco Attendance Dashboard
==================================================
Features:
  1. Device online / offline alerts
  2. Per-punch notifications (IP, time, employee code)
  3. Daily 08:10 absent report (XLSX document, grouped by dept category)

Configuration (settings.ini [telegram] section):
  bot_token   = <your bot token>
  chat_id     = <target chat / group id>
  enabled     = 1
  notify_device_status = 1
  notify_punches       = 1
  notify_daily_report  = 1
  daily_report_hour    = 8
  daily_report_minute  = 10
"""

import io
import logging
import threading
import time
from datetime import datetime
from typing import Optional, Dict

import httpx

logger = logging.getLogger("ZKTeco.Telegram")


# ---------------------------------------------------------------------------
#  Low-level HTTP helpers (fully synchronous, no asyncio dependency)
# ---------------------------------------------------------------------------

def _post(url: str, **kwargs) -> Optional[dict]:
    """POST to Telegram API; return parsed JSON or None on error."""
    try:
        with httpx.Client(timeout=15) as client:
            resp = client.post(url, **kwargs)
        if resp.status_code == 200:
            return resp.json()
        logger.warning("Telegram API %s: status %s  %s", url, resp.status_code, resp.text[:200])
    except Exception as exc:
        logger.warning("Telegram send error: %s", exc)
    return None


# ---------------------------------------------------------------------------
#  TelegramNotifier
# ---------------------------------------------------------------------------

class TelegramNotifier:
    """
    Telegram bot notifier for the ZKTeco attendance dashboard.
    All send_* methods are synchronous and thread-safe.
    """

    def __init__(
        self,
        bot_token: str,
        chat_id: str,
        enabled: bool = True,
        notify_device_status: bool = True,
        notify_punches: bool = True,
        notify_daily_report: bool = True,
        system_name: str = "Attendance",
        notification_settings: dict = None,
    ):
        # notification_settings can override individual notify_* flags when provided
        if notification_settings:
            notify_device_status = notification_settings.get("device_status", notify_device_status)
            notify_punches       = notification_settings.get("punches",       notify_punches)
            notify_daily_report  = notification_settings.get("daily_report",  notify_daily_report)
        self.bot_token = bot_token.strip() if bot_token else ""
        self.chat_id = str(chat_id).strip() if chat_id else ""
        self.enabled = enabled
        self.notify_device_status = notify_device_status
        self.notify_punches = notify_punches
        self.notify_daily_report = notify_daily_report
        self.system_name = system_name
        self._base = "https://api.telegram.org/bot{0}".format(self.bot_token)
        self._lock = threading.Lock()

    # ------------------------------------------------------------------ #
    #  Internal plumbing                                                   #
    # ------------------------------------------------------------------ #

    def _ok(self) -> bool:
        return bool(self.enabled and self.bot_token and self.chat_id)

    def _send_message(self, text: str, parse_mode: str = "HTML") -> bool:
        """Send a plain text message."""
        if not self._ok():
            return False
        # Telegram HTML has a 4096-char limit; truncate gracefully
        if len(text) > 4000:
            text = text[:3990] + "\n…"
        with self._lock:
            result = _post(
                self._base + "/sendMessage",
                json={"chat_id": self.chat_id, "text": text, "parse_mode": parse_mode},
            )
        return result is not None and result.get("ok", False)

    def _send_document(self, file_bytes: bytes, filename: str, caption: str = "") -> bool:
        """Upload a binary file as a Telegram document."""
        if not self._ok():
            return False
        with self._lock:
            result = _post(
                self._base + "/sendDocument",
                data={"chat_id": self.chat_id, "caption": caption[:1024]},
                files={"document": (filename, file_bytes,
                                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            )
        return result is not None and result.get("ok", False)

    # ------------------------------------------------------------------ #
    #  Feature 1: Device online / offline                                  #
    # ------------------------------------------------------------------ #

    def notify_device_online(self, ip: str, name: str = "") -> bool:
        if not (self._ok() and self.notify_device_status):
            return False
        label = "{0} ({1})".format(name, ip) if name else ip
        msg = (
            "✅ <b>{sys}</b> — Device ONLINE\n"
            "📡 <b>{label}</b>\n"
            "🕐 {ts}"
        ).format(sys=self.system_name, label=label,
                 ts=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        return self._send_message(msg)

    def notify_device_offline(self, ip: str, name: str = "", error: str = "") -> bool:
        if not (self._ok() and self.notify_device_status):
            return False
        label = "{0} ({1})".format(name, ip) if name else ip
        err_line = "\n⚠️ {0}".format(error) if error else ""
        msg = (
            "🔴 <b>{sys}</b> — Device OFFLINE\n"
            "📡 <b>{label}</b>\n"
            "🕐 {ts}{err}"
        ).format(sys=self.system_name, label=label,
                 ts=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 err=err_line)
        return self._send_message(msg)

    # ------------------------------------------------------------------ #
    #  Feature 2: Per-punch notification                                   #
    # ------------------------------------------------------------------ #

    def notify_punch(self, badge: str, name: str, ip: str, ts: str) -> bool:
        if not (self._ok() and self.notify_punches):
            return False
        msg = (
            "👆 <b>Punch recorded</b>\n"
            "🪪 <b>{badge}</b> — {name}\n"
            "📡 Device: <code>{ip}</code>\n"
            "🕐 {ts}"
        ).format(badge=badge, name=name or "Unknown", ip=ip, ts=ts)
        return self._send_message(msg)

    # ------------------------------------------------------------------ #
    #  Feature 3 & 4: Daily absent report at 08:10                        #
    # ------------------------------------------------------------------ #

    def send_daily_absent_report(
        self,
        absent: list,
        present_count: int,
        total: int,
        date_str: str,
        dept_order: list,
    ) -> bool:
        """
        Send the daily absent report.
        Sends a short summary text message AND an XLSX attachment
        grouped by dept category:
          Teachers | Admin+Support | Drivers+Conductors | Cleaners | Others
        """
        if not (self._ok() and self.notify_daily_report):
            return False

        absent_count = len(absent)

        # ---------- 1. Text summary ----------
        # Group absent by dept for the inline summary
        dept_buckets: Dict[str, list] = {}
        for emp in absent:
            dept_buckets.setdefault(emp.get("dept", "OTHER"), []).append(emp)

        lines = [
            "📋 <b>{sys} — Daily Absent Report</b>".format(sys=self.system_name),
            "📅 {date}".format(date=date_str),
            "❌ Absent: <b>{a}</b>  ✅ Present: <b>{p}</b>  👥 Total: <b>{t}</b>".format(
                a=absent_count, p=present_count, t=total),
            "",
        ]

        # Use configured dept order for display priority
        ordered_depts = [d for d in (dept_order or []) if d in dept_buckets]
        ordered_depts += sorted(k for k in dept_buckets if k not in ordered_depts)
        for dept in ordered_depts:
            emps = dept_buckets[dept]
            lines.append("<b>{dept}</b> ({n} absent)".format(dept=dept, n=len(emps)))
            for emp in sorted(emps, key=lambda e: e.get("name", "")):
                lines.append("  · {code}  {name}".format(
                    code=emp.get("code", ""), name=emp.get("name", "")))
            lines.append("")

        msg_ok = self._send_message("\n".join(lines))

        # ---------- 2. XLSX attachment ----------
        try:
            xlsx_bytes = _build_absent_xlsx(absent, date_str)
            fname = "absent_{0}.xlsx".format(
                datetime.now().strftime("%Y%m%d"))
            caption = "Absent report {0} — {1} absent / {2} total".format(
                date_str, absent_count, total)
            self._send_document(xlsx_bytes, fname, caption)
        except Exception as exc:
            logger.warning("Could not build XLSX for Telegram: %s", exc)

        return msg_ok

    # ------------------------------------------------------------------ #
    #  Utility                                                             #
    # ------------------------------------------------------------------ #

    def test_connection(self) -> bool:
        """Send a test message to verify the bot is working."""
        msg = (
            "🧪 <b>{sys} — Test</b>\n"
            "📅 {ts}\n"
            "✅ Telegram bot is connected and working."
        ).format(sys=self.system_name, ts=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        return self._send_message(msg)

    # Kept for backward compatibility with old code
    def send_message_sync(self, message: str, parse_mode: str = "HTML") -> bool:
        return self._send_message(message, parse_mode)


# ---------------------------------------------------------------------------
#  XLSX builder (no server imports — self-contained)
# ---------------------------------------------------------------------------

_CATEGORY_MAP = {
    "TEACHING":       "Teachers",
    "ADMIN":          "Admin & Support",
    "SUPPORT":        "Admin & Support",
    "DRIVER":         "Drivers & Conductors",
    "CONDUCTOR":      "Drivers & Conductors",
    "CLEANING STAFF": "Cleaners",
}
_CATEGORY_ORDER = ["Teachers", "Admin & Support", "Drivers & Conductors", "Cleaners"]
_CAT_COLORS = {
    "Teachers":           "1F4E79",
    "Admin & Support":    "375623",
    "Drivers & Conductors": "7B3F00",
    "Cleaners":           "4A235A",
}


def _build_absent_xlsx(absent: list, date_str: str) -> bytes:
    """
    Build an XLSX workbook for the daily absent report.
    Columns: No. | Name | Date | Timetable
    Rows grouped by dept category (Teachers / Admin+Support / etc.)
    Returns raw bytes.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    def _thin():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    wb = Workbook()

    # ---------- single combined sheet ----------
    ws = wb.active
    ws.title = "Absent"
    ws.sheet_view.showGridLines = False

    # Title row
    ws.merge_cells("A1:D1")
    c = ws.cell(row=1, column=1, value="Daily Absent Report — {0}".format(date_str))
    c.font = Font(name="Arial", bold=True, size=13, color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Sub-title
    ws.cell(row=2, column=1,
            value="Generated: {0}".format(datetime.now().strftime("%d %b %Y %H:%M"))
            ).font = Font(name="Arial", size=9, italic=True, color="888888")

    # Header row
    headers = ["No.", "Name", "Date", "Timetable"]
    for col, hdr in enumerate(headers, 1):
        c = ws.cell(row=4, column=col, value=hdr)
        c.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", start_color="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = _thin()
    ws.row_dimensions[4].height = 18

    # Group by category
    buckets: Dict[str, list] = {}
    for emp in absent:
        dept = (emp.get("dept") or "").upper()
        cat = _CATEGORY_MAP.get(dept, "Others")
        buckets.setdefault(cat, []).append(emp)

    ordered_cats = [c for c in _CATEGORY_ORDER if c in buckets]
    ordered_cats += sorted(k for k in buckets if k not in ordered_cats)

    row = 5
    for cat in ordered_cats:
        emps = sorted(buckets[cat], key=lambda e: e.get("name", ""))
        cat_color = _CAT_COLORS.get(cat, "444444")

        # Category separator row
        ws.merge_cells("A{r}:D{r}".format(r=row))
        c = ws.cell(row=row, column=1, value="{cat}  ({n} absent)".format(cat=cat, n=len(emps)))
        c.font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
        c.fill = PatternFill("solid", start_color=cat_color)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        c.border = _thin()
        ws.row_dimensions[row].height = 20
        row += 1

        for i, emp in enumerate(emps):
            bg = "EBF3FB" if i % 2 == 0 else "FFFFFF"
            vals = [
                emp.get("code", ""),
                emp.get("name", ""),
                date_str,
                emp.get("dept", ""),
            ]
            for col, val in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=val)
                c.font = Font(name="Arial", size=10)
                c.fill = PatternFill("solid", start_color=bg)
                c.border = _thin()
                c.alignment = Alignment(
                    horizontal="center" if col in (1, 3) else "left",
                    vertical="center",
                )
            row += 1

    # Column widths
    for col, w in zip("ABCD", [12, 34, 14, 22]):
        ws.column_dimensions[col].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()