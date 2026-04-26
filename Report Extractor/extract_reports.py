#!/usr/bin/env python3
"""
Report Text Extractor
Bear Engineering — Foundation & Drainage Inspection Reports

Walks a folder tree, finds all matching Keynote reports, extracts
slide titles and body text via AppleScript, and writes one .txt file
per report to a timestamped folder on the Desktop.

Usage: Double-click Extract Reports.command to run.
"""

import os
import sys
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

# ── Config ────────────────────────────────────────────────────────────────────

REPORT_KEYWORD = "Foundation & Drainage Inspection Report"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
APPLESCRIPT_PATH = os.path.join(SCRIPT_DIR, "extract_report.applescript")

# ── Helpers ───────────────────────────────────────────────────────────────────

def choose_folder():
    """Open a folder picker dialog and return the chosen path, or None."""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(
        title="Select the folder containing inspection reports"
    )
    root.destroy()
    return folder if folder else None


def find_keynote_files(root_folder):
    """Recursively find all .key files whose name contains the report keyword."""
    matches = []
    for dirpath, dirnames, filenames in os.walk(root_folder):
        # Skip hidden directories (e.g. .git, .Trash)
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for filename in filenames:
            if filename.endswith(".key") and REPORT_KEYWORD in filename:
                matches.append(os.path.join(dirpath, filename))
    return sorted(matches)


def extract_address_from_filename(filepath):
    """
    Pull the address out of the filename.
    Filename pattern: Foundation & Drainage Inspection Report - ADDRESS.key
    Returns the address string, or the full stem if pattern not matched.
    """
    stem = Path(filepath).stem  # filename without .key
    prefix = REPORT_KEYWORD + " - "
    if stem.startswith(prefix):
        return stem[len(prefix):]
    return stem


def is_file_local(path):
    """
    Check if a file is fully downloaded locally (not a cloud placeholder).
    Uses xattr to check for com.apple.icloud.itemName which indicates
    an iCloud placeholder that hasn't been downloaded yet.
    """
    try:
        result = subprocess.run(
            ["xattr", "-l", path],
            capture_output=True, text=True
        )
        # If the file has iCloud download pending attributes, it needs downloading
        return "com.apple.icloud.downloadRequested" not in result.stdout
    except Exception:
        return True  # Assume local if we can't check


def pre_open_file(key_path):
    """
    Open a Keynote file and immediately close it.
    This forces macOS to download and fully initialize cloud-synced files
    before the extraction script tries to read them.
    Uses a shorter delay for already-local files.
    Returns (success: bool, message: str)
    """
    # Cloud files need longer to download; local files just need a brief open
    delay = 1 if is_file_local(key_path) else 4

    script = f'''
        tell application "Keynote"
            set theDoc to open POSIX file "{key_path}"
            delay {delay}
            close theDoc saving no
        end tell
        return "OK"
    '''
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        err = result.stderr.strip() or result.stdout.strip()
        return False, f"Could not pre-open file: {err}"
    return True, "OK"


def run_applescript(key_path, temp_output_path):
    """
    Run the AppleScript extractor on one Keynote file.
    Returns (success: bool, message: str)
    """
    result = subprocess.run(
        [
            "osascript",
            APPLESCRIPT_PATH,
            key_path,
            temp_output_path,
        ],
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        err = result.stderr.strip() or result.stdout.strip()
        return False, f"AppleScript error: {err}"

    stdout = result.stdout.strip()
    if stdout.startswith("ERROR"):
        return False, stdout

    return True, "OK"


def build_output_text(address, extracted_text):
    """Wrap the extracted slide content with a report header."""
    header = (
        f"=== REPORT: {address} ===\n"
        f"{'=' * (len(address) + 12)}\n\n"
    )
    return header + extracted_text.strip() + "\n"


def make_safe_filename(address):
    """Turn an address string into a safe filename."""
    # Replace characters that are problematic in filenames
    safe = address.replace("/", "-").replace(":", "-").replace("\\", "-")
    return safe + ".txt"


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    # 1. Pick a folder
    print("Opening folder picker...")
    source_folder = choose_folder()
    if not source_folder:
        print("No folder selected. Exiting.")
        sys.exit(0)

    print(f"Searching: {source_folder}")

    # 2. Find all matching Keynote files
    keynote_files = find_keynote_files(source_folder)

    if not keynote_files:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo(
            "No Reports Found",
            f"No files matching\n\"{REPORT_KEYWORD}\"\nwere found in the selected folder."
        )
        root.destroy()
        sys.exit(0)

    print(f"Found {len(keynote_files)} report(s).")

    # 3. Create output folder on the Desktop
    desktop = os.path.expanduser("~/Desktop")
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    output_folder_name = f"Extracted_Inspection_Reports_{timestamp}"
    output_folder = os.path.join(desktop, output_folder_name)
    os.makedirs(output_folder, exist_ok=True)
    print(f"Output folder: {output_folder}\n")

    # 4. Process each file
    success_count = 0
    fail_count = 0
    errors = []

    for idx, key_path in enumerate(keynote_files, start=1):
        address = extract_address_from_filename(key_path)
        print(f"[{idx}/{len(keynote_files)}] {address} ...", end=" ", flush=True)

        # Use a temp file for AppleScript to write into
        with tempfile.NamedTemporaryFile(
            mode="r", suffix=".txt", delete=False
        ) as tmp:
            tmp_path = tmp.name

        try:
            # Pre-open the file to force cloud download/initialization.
            # Required for iCloud-synced files that haven't been opened locally yet.
            pre_ok, pre_msg = pre_open_file(key_path)
            if not pre_ok:
                print(f"FAILED (pre-open)")
                errors.append(f"{address}: {pre_msg}")
                fail_count += 1
                continue

            success, message = run_applescript(key_path, tmp_path)

            if not success:
                print(f"FAILED")
                errors.append(f"{address}: {message}")
                fail_count += 1
                continue

            # Read what AppleScript wrote
            with open(tmp_path, "r", encoding="mac_roman") as f:
                extracted = f.read()

            if not extracted.strip():
                print(f"SKIPPED (no text extracted)")
                errors.append(f"{address}: No text was extracted")
                fail_count += 1
                continue

            # Write the final .txt
            output_text = build_output_text(address, extracted)
            out_filename = make_safe_filename(address)
            out_path = os.path.join(output_folder, out_filename)

            with open(out_path, "w", encoding="utf-8") as f:
                f.write(output_text)

            print(f"OK")
            success_count += 1

        finally:
            # Clean up temp file
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    # 5. Summary
    print(f"\n{'─' * 50}")
    print(f"Done.  {success_count} succeeded,  {fail_count} failed.")
    print(f"Output: {output_folder}")

    if errors:
        print("\nErrors:")
        for e in errors:
            print(f"  • {e}")

    # 6. Show a finish dialog and open the output folder
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    if fail_count == 0:
        msg = (
            f"Extraction complete.\n\n"
            f"{success_count} report(s) extracted successfully.\n\n"
            f"Output folder:\n{output_folder_name}"
        )
    else:
        msg = (
            f"Extraction complete.\n\n"
            f"{success_count} succeeded,  {fail_count} failed.\n\n"
            f"Check the terminal window for error details.\n\n"
            f"Output folder:\n{output_folder_name}"
        )

    messagebox.showinfo("Report Extractor", msg)
    root.destroy()

    # Open the output folder in Finder
    subprocess.run(["open", output_folder])


if __name__ == "__main__":
    main()
