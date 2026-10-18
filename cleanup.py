#!/usr/bin/env python3
"""
Deletes all scan records (cafeteria.db) and the generated files in output/.
Keeps settings.ini and employees.csv.

    python3 cleanup.py        # shows what will be deleted, asks for confirmation
    python3 cleanup.py -y     # no confirmation

If the cafeteria service is running, it is stopped first and started again afterwards:
deleting the database while cafeteria.py has it open would silently lose new scans.
"""

import subprocess
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_FILES = [BASE_DIR / f"cafeteria.db{suffix}" for suffix in ("", "-journal", "-wal", "-shm")]
OUTPUT_DIR = BASE_DIR / "output"
SERVICE = "cafeteria"


def run(cmd):
    try:
        return subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    except (OSError, subprocess.SubprocessError):
        return None


def service_active():
    r = run(["systemctl", "is-active", "--quiet", SERVICE])
    return r is not None and r.returncode == 0


def manual_run_active():
    """cafeteria.py started by hand in a terminal (not by the service)."""
    r = run(["pgrep", "-f", "cafeteria.py"])
    return r is not None and r.returncode == 0


def main():
    targets = [f for f in DB_FILES if f.exists()]
    if OUTPUT_DIR.is_dir():
        targets += sorted(p for p in OUTPUT_DIR.iterdir() if p.is_file())

    if not targets:
        print("Nothing to clean: no database and no output files.")
        return 0

    print("Will delete:")
    for f in targets:
        print(f"  {f.relative_to(BASE_DIR)}")
    print("Kept: settings.ini, employees.csv")

    if "-y" not in sys.argv[1:]:
        if input("\nType 'yes' to delete: ").strip().lower() != "yes":
            print("Cancelled. Nothing deleted.")
            return 1

    restart = False
    if service_active():
        print(f"Stopping the {SERVICE} service...")
        r = run(["sudo", "-n", "systemctl", "stop", SERVICE])
        if r is None or r.returncode != 0:
            print("Could not stop the service. Run: sudo systemctl stop cafeteria  then try again.")
            return 1
        restart = True
    elif manual_run_active():
        print("cafeteria.py is running in a terminal. Quit it (press 0) and run cleanup again.")
        return 1

    deleted = 0
    for f in targets:
        try:
            f.unlink()
            deleted += 1
        except FileNotFoundError:
            pass
        except OSError as e:
            print(f"  Could not delete {f.name}: {e}")
    print(f"Deleted {deleted} file(s).")

    if restart:
        print(f"Starting the {SERVICE} service again...")
        r = run(["sudo", "-n", "systemctl", "start", SERVICE])
        if r is None or r.returncode != 0:
            print("Could not start the service. Run: sudo systemctl start cafeteria")
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())