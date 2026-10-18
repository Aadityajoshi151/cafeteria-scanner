#!/usr/bin/env python3
"""
cafeteria scanner - prototype v2
Reader + Veg/Non-Veg/Complete buttons + buzzer + OLED + local storage + Excel export
+ email on Complete (settings in settings.ini).
"""

import configparser
import csv
import os
import queue
import smtplib
import sqlite3
import subprocess
import sys
import termios
import threading
import time
import tty
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import serial
from gpiozero import Button, Buzzer
from luma.core.interface.serial import spi
from luma.core.render import canvas
from luma.oled.device import ssd1306
from openpyxl import Workbook
from openpyxl.styles import Font
from PIL import ImageFont

# ---------------- Settings (these will move to a config file later) ----------------
BASE_DIR = Path(__file__).resolve().parent
EMPLOYEES_CSV = BASE_DIR / "employees.csv"   # columns: card_id,emp_id,name
DB_PATH = BASE_DIR / "cafeteria.db"
OUTPUT_DIR = BASE_DIR / "output"
SETTINGS_INI = BASE_DIR / "settings.ini"     # email settings (read each time Complete is pressed)

SERIAL_PORT = "/dev/serial0"
PIN_VEG, PIN_NONVEG, PIN_COMPLETE, PIN_BUZZER = 17, 27, 22, 18
OLED_DC, OLED_RST = 24, 25

CHOICE_TIMEOUT = 10   # seconds to press Veg/Non-Veg after tapping a card
MESSAGE_TIME = 3      # seconds a result message stays on screen
REPEAT_IGNORE = 2     # ignore the same card for this long (module repeats while a card is held)
QUIT_KEY = "0"        # press this key in the terminal to stop the script
IDLE_REFRESH = 5      # seconds between idle-screen updates (date, counts, clock warning)
CLOCK_WARNING = True  # show "Clock not set" until the time is synced from the internet
                      # (set to False once an RTC module keeps the time without network)

FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

VEG, NONVEG, COMPLETE = "VEG", "NON-VEG", "COMPLETE"

events = queue.Queue()   # ("card", n), ("button", name) or ("email", (ok, text)); used by the main loop


# ---------------- Card reader ----------------
def parse_frame(buf):
    """RDM6300 frame body: 10 hex chars of data + 2 hex chars of checksum. Returns card number or None."""
    try:
        frame = buf.decode("ascii")
        if len(frame) != 12:
            return None
        data, checksum = frame[:10], int(frame[10:], 16)
        x = 0
        for i in range(0, 10, 2):
            x ^= int(data[i:i + 2], 16)
        return int(data[2:], 16) if x == checksum else None
    except (UnicodeDecodeError, ValueError):
        return None


def reader_thread():
    ser = serial.Serial(SERIAL_PORT, 9600, timeout=1)
    buf, last_id, last_time = b"", None, 0.0
    while True:
        b = ser.read(1)
        if not b:
            continue
        if b == b"\x02":                 # start of frame
            buf = b""
        elif b == b"\x03":               # end of frame
            n = parse_frame(buf)
            buf = b""
            if n is None:
                continue
            now = time.monotonic()
            if n == last_id and now - last_time < REPEAT_IGNORE:
                last_time = now          # card still being held: ignore
                continue
            last_id, last_time = n, now
            events.put(("card", n))
        else:
            buf += b
            if len(buf) > 20:            # junk guard
                buf = b""


def card_formats(n):
    """All the ways a card number may be written in the employee list."""
    return {f"{n:010d}", str(n), f"{(n >> 16) & 0xFF:03d},{n & 0xFFFF:05d}"}


# ---------------- Employee list ----------------
class Employees:
    """Reads employees.csv and reloads it automatically whenever the file changes."""

    def __init__(self, path):
        self.path, self.mtime, self.by_card = path, None, {}

    def _reload_if_changed(self):
        try:
            mtime = self.path.stat().st_mtime
        except FileNotFoundError:
            if self.by_card:
                print("employees.csv not found - all cards will be unmatched")
            self.by_card, self.mtime = {}, None
            return
        if mtime == self.mtime:
            return
        self.mtime = mtime
        table = {}
        with open(self.path, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                card_id = (row.get("card_id") or "").strip()
                if card_id:
                    table[card_id] = ((row.get("emp_id") or "").strip(),
                                      (row.get("name") or "").strip())
        self.by_card = table
        print(f"Loaded {len(table)} employees")

    def lookup(self, n):
        self._reload_if_changed()
        for key in card_formats(n):
            if key in self.by_card:
                return self.by_card[key]
        return None


# ---------------- Storage ----------------
def open_db():
    db = sqlite3.connect(DB_PATH)
    db.execute("PRAGMA synchronous=FULL")    # every scan is safely on disk before we beep OK
    db.execute("""CREATE TABLE IF NOT EXISTS lunch(
                    day TEXT, card_id TEXT, emp_id TEXT, name TEXT, meal TEXT, time TEXT,
                    PRIMARY KEY (day, card_id))""")
    db.execute("""CREATE TABLE IF NOT EXISTS log(
                    ts TEXT, card_id TEXT, meal TEXT, result TEXT)""")
    db.commit()
    return db


def record(db, n, emp, meal):
    """One row per card per day. Returns 'new', 'same' (already recorded) or 'changed'."""
    now = datetime.now()
    day, t = now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S")
    card_id = f"{n:010d}"
    emp_id, name = emp if emp else ("", "")
    row = db.execute("SELECT meal FROM lunch WHERE day=? AND card_id=?", (day, card_id)).fetchone()
    result = "new" if row is None else ("same" if row[0] == meal else "changed")
    if result != "same":
        db.execute("INSERT OR REPLACE INTO lunch VALUES (?,?,?,?,?,?)",
                   (day, card_id, emp_id, name, meal, t))
    db.execute("INSERT INTO log VALUES (?,?,?,?)",
               (now.isoformat(timespec="seconds"), card_id, meal, result))
    db.commit()
    return result


def today_counts(db):
    day = datetime.now().strftime("%Y-%m-%d")
    rows = dict(db.execute("SELECT meal, COUNT(*) FROM lunch WHERE day=? GROUP BY meal",
                           (day,)).fetchall())
    return rows.get(VEG, 0), rows.get(NONVEG, 0)


def export_today(db):
    """Writes output/lunch_YYYY-MM-DD.xlsx with a 'Lunch' sheet and an 'Unmatched' sheet."""
    day = datetime.now().strftime("%Y-%m-%d")
    rows = db.execute("SELECT emp_id, name, meal, time, card_id FROM lunch WHERE day=? ORDER BY time",
                      (day,)).fetchall()
    matched = [r for r in rows if r[0]]
    unmatched = [r for r in rows if not r[0]]
    bold = Font(bold=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Lunch"
    ws.append(["Emp ID", "Name", "Veg", "Non-Veg", "Time"])
    for emp_id, name, meal, t, _ in matched:
        ws.append([emp_id, name, "✓" if meal == VEG else "", "✓" if meal == NONVEG else "", t])
    ws.append([])
    ws.append(["", "Total",
               sum(1 for r in matched if r[2] == VEG),
               sum(1 for r in matched if r[2] == NONVEG), ""])
    for cell in ws[1] + ws[ws.max_row]:
        cell.font = bold
    for col, width in zip("ABCDE", (12, 28, 8, 10, 10)):
        ws.column_dimensions[col].width = width

    ws2 = wb.create_sheet("Unmatched")
    ws2.append(["Card ID", "Wiegand", "Meal", "Time", "Who (fill from register)"])
    for _, _, meal, t, card_id in unmatched:
        n = int(card_id)
        ws2.append([card_id, f"{(n >> 16) & 0xFF:03d},{n & 0xFFFF:05d}",