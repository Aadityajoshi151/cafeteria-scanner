#!/usr/bin/env python3
"""
Adds cards to employees.csv: tap a card, type the employee ID and name, repeat.

    python3 enroll.py

Press 0 while it is waiting for a card to quit. If the cafeteria service is running it is
stopped while you enroll (both need the card reader) and started again when you quit.
cafeteria.py reloads employees.csv by itself, so new cards work straight away.
"""

import csv
import select
import subprocess
import sys
import termios
import time
import tty
from pathlib import Path

import serial

BASE_DIR = Path(__file__).resolve().parent
EMPLOYEES_CSV = BASE_DIR / "employees.csv"
SERIAL_PORT = "/dev/serial0"
SERVICE = "cafeteria"
FIELDS = ["card_id", "emp_id", "name"]


# ---------------- Card reader ----------------
def parse_frame(buf):
    """RDM6300 frame body: 10 hex data chars + 2 hex checksum chars. Returns card number or None."""
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


def wait_for_card(ser):
    """Returns a card number, or None if 0 was pressed. Single keys, no Enter needed."""
    fd = sys.stdin.fileno()
    old = termios.tcgetattr(fd)
    tty.setcbreak(fd)
    buf = b""
    try:
        while True:
            ready, _, _ = select.select([ser, sys.stdin], [], [])
            if sys.stdin in ready and sys.stdin.read(1) in ("0", ""):
                return None
            if ser not in ready:
                continue
            for i in ser.read(ser.in_waiting or 1):
                b = bytes([i])
                if b == b"\x02":
                    buf = b""
                elif b == b"\x03":
                    n = parse_frame(buf)
                    buf = b""
                    if n is not None:
                        return n
                else:
                    buf += b
    finally:
        termios.tcsetattr(fd, termios.TCSADRAIN, old)   # normal typing for input()


def drain(ser, quiet_for=0.5):
    """The reader repeats a card while it is held: throw those repeats away."""
    end = time.monotonic() + quiet_for
    while time.monotonic() < end:
        if ser.read(ser.in_waiting or 1):
            end = time.monotonic() + quiet_for       # still sending: keep waiting


# ---------------- employees.csv ----------------
def load_rows():
    if not EMPLOYEES_CSV.exists():
        return []
    with open(EMPLOYEES_CSV, newline="", encoding="utf-8-sig") as f:
        return [{k: (row.get(k) or "").strip() for k in FIELDS} for row in csv.DictReader(f)]


def card_keys(n):
    """The ways the same card may already be written in the file."""
    return {f"{n:010d}", str(n), f"{(n >> 16) & 0xFF:03d},{n & 0xFFFF:05d}"}


def append_row(row):
    new_file = not EMPLOYEES_CSV.exists() or EMPLOYEES_CSV.stat().st_size == 0
    if not new_file:
        with open(EMPLOYEES_CSV, "rb") as f:          # make sure we start on a fresh line
            f.seek(-1, 2)
            needs_newline = f.read(1) not in (b"\n", b"\r")
    with open(EMPLOYEES_CSV, "a", newline="", encoding="utf-8") as f:
        if not new_file and needs_newline:
            f.write("\n")
        w = csv.DictWriter(f, fieldnames=FIELDS)
        if new_file:
            w.writeheader()
        w.writerow(row)


def rewrite_all(rows):
    tmp = EMPLOYEES_CSV.with_suffix(".csv.tmp")
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)
    tmp.replace(EMPLOYEES_CSV)                         # all-or-nothing swap


def ask(prompt):
    try:
        return input(prompt).strip()
    except EOFError:
        return ""


# ---------------- Service ----------------
def run(cmd):
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except (OSError, subprocess.SubprocessError):
        return None


def service_active():
    r = run(["systemctl", "is-active", "--quiet", SERVICE])
    return r is not None and r.returncode == 0


# ---------------- Main ----------------
def main():
    restart = False
    if service_active():
        print(f"Stopping the {SERVICE} service while enrolling...")
        r = run(["sudo", "-n", "systemctl", "stop", SERVICE])
        if r is None or r.returncode != 0:
            print("Could not stop it. Run: sudo systemctl stop cafeteria  then try again.")
            return 1
        restart = True
    else:
        r = run(["pgrep", "-f", r"python[0-9.]* (-u )?\S*cafeteria\.py"])
        if r is not None and r.returncode == 0:
            print("cafeteria.py is running in a terminal. Quit it (press 0) and try again.")
            return 1

    added = updated = 0
    try:
        ser = serial.Serial(SERIAL_PORT, 9600, timeout=0)
        print(f"Adding cards to {EMPLOYEES_CSV.name}. Press 0 to quit.\n")
        while True:
            print("Tap a card...")
            n = wait_for_card(ser)
            if n is None:
                break
            card_id = f"{n:010d}"
            print(f"  Card ID: {card_id}  (Wiegand {(n >> 16) & 0xFF:03d},{n & 0xFFFF:05d})")

            rows = load_rows()
            existing = next((r for r in rows if r["card_id"] in card_keys(n)), None)
            if existing:
                print(f"  Already in the list: {existing['emp_id']}, {existing['name']}")
                if ask("  Replace it? (y/N): ").lower() != "y":
                    print("  Kept as it is.\n")
                    drain(ser)
                    continue

            emp_id = ask("  Employee ID (Enter to skip this card): ")
            if not emp_id:
                print("  Skipped.\n")
                drain(ser)
                continue
            name = ask("  Name: ")

            other = next((r for r in rows if r["emp_id"].lower() == emp_id.lower()
                          and r is not existing), None)
            if other:
                print(f"  Note: {emp_id} is also on card {other['card_id']} ({other['name']}).")

            row = {"card_id": card_id, "emp_id": emp_id, "name": name}
            if existing:
                rows[rows.index(existing)] = row
                rewrite_all(rows)
                updated += 1
                print(f"  Updated: {card_id}, {emp_id}, {name}\n")
            else:
                append_row(row)
                added += 1
                print(f"  Added: {card_id}, {emp_id}, {name}\n")
            drain(ser)                                 # ignore the card if it is still on the reader
        ser.close()
    except KeyboardInterrupt:
        print()
    finally:
        print(f"Done: {added} added, {updated} updated.")
        if restart:
            print(f"Starting the {SERVICE} service again...")
            r = run(["sudo", "-n", "systemctl", "start", SERVICE])
            if r is None or r.returncode != 0:
                print("Could not start it. Run: sudo systemctl start cafeteria")
    return 0


if __name__ == "__main__":
    sys.exit(main())