#!/usr/bin/env python3
"""
AYSO PlayMetrics Player Import Tool
====================================
Converts Sports Affinity player data into PlayMetrics Player Import CSV format.

Built for AYSO regions migrating from SportsConnect/Sports Affinity to PlayMetrics.
Designed to be shared across pilot regions — no dependencies beyond pandas and openpyxl.

USAGE:
    python playmetrics_import.py

    The tool will prompt you for file locations. Place your downloaded reports
    in the same folder as this script, or provide full paths when prompted.

REQUIRED INPUT FILES:
    1. "Player Detail | upload format" from Sports Affinity (playerUpload.xlsx)
       - This is your PRIMARY data source — has player names, DOB, gender,
         parent contacts, addresses, and team assignments.
       - Found under: Players/Admins → Player Lookup → Report dropdown

    2. "Player Photo BC Info" from Sports Affinity (Player_Photo_BC_Info.xlsx)
       - Shows birth certificate upload dates and verification dates per player.
       - Players with a BC Uploaded date OR a BC Verified date have a BC on file.
       - Found under: Reports → Player Photo BC Info

HOW TO DOWNLOAD THE REPORTS:
    Report 1 — "Player Detail | upload format":
    1. Log in to Sports Connect → click "Change Login" → select the Association site
    2. Navigate to Players/Admins → Player Lookup
    3. Set the season dropdown (e.g., "2025-2026 MY2025")
    4. Set Region to your region
    5. Set Age Group to "Multiple" (select all)
    6. From the Report dropdown, select "Player Detail | upload format" → Print
    7. In the report viewer, select "Excel" format → Export
    8. Save the file (downloads as playerUpload.xlsx)

    Report 2 — "Player Photo BC Info":
    1. In Sports Affinity, navigate to Reports (top menu)
    2. From the report dropdown, select "Player Photo BC Info"
    3. Set Area and Region filters
    4. Leave Date Range blank (gets all players)
    5. Click "Generate Report"
    6. In the report viewer, select "Excel" format → Export
    7. Save the file (downloads as Player_Photo_BC_Info.xlsx)

    For multi-year imports: change the season/date range and repeat for each year

OUTPUT:
    Two CSV files ready for PlayMetrics:
    - playmetrics_bc_verified_YYYYMMDD.csv   → Send to PlayMetrics (success@playmetrics.com)
    - playmetrics_non_verified_YYYYMMDD.csv  → Import yourself in PlayMetrics admin

Author: Steve Davis, Registrar — AYSO Region 58
        Built with assistance from Claude (Anthropic)
Version: 1.0 — May 2026
"""

import os
import re
import csv
import sys
import glob
import argparse
import io
import subprocess
import platform
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple, Callable

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required. Install it with:")
    print("  pip install pandas openpyxl")
    sys.exit(1)

# =========================================================
#  CONSTANTS
# =========================================================

# PlayMetrics Player Import CSV columns (must match their template exactly)
PM_COLUMNS = [
    "team",
    "season_id",
    "season",
    "player_first_name",
    "player_last_name",
    "gender",
    "birth_date",
    "age_group",
    "position",
    "number",
    "Foot",
    "parent1_email",
    "parent1_first_name",
    "parent1_last_name",
    "parent1_mobile_number",
    "parent2_email",
    "parent2_first_name",
    "parent2_last_name",
    "parent2_mobile_number",
    "street",
    "city",
    "state",
    "zip",
]

# Oldest eligible birth date for Fall 2026 (Aug 1 cutoff, 19U max)
# Adjust this for your season
OLDEST_ELIGIBLE_DOB = "2007-08-01"

# =========================================================
#  DATA LOADING
# =========================================================


def find_file(pattern: str, search_dir: str = ".") -> Optional[str]:
    """Find the most recent file matching a glob pattern."""
    matches = glob.glob(os.path.join(search_dir, pattern))
    if matches:
        return max(matches, key=os.path.getmtime)
    return None


def load_player_upload(filepath: str) -> pd.DataFrame:
    """
    Load the "Player Detail | upload format" report from Sports Affinity.

    This is the primary data source with pre-split parent names, M/F gender,
    addresses, and team assignments.
    """
    print(f"  Loading player upload: {filepath}")
    df = pd.read_excel(filepath)
    print(f"  → {len(df)} players loaded")

    # Validate expected columns
    required = [
        "PlayerFirstName",
        "PlayerLastName",
        "Gender",
        "DOB",
        "FatherEmailAddress",
        "FatherFirstName",
        "FatherLastName",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"  ⚠️  Missing expected columns: {missing}")
        print(f"  Available columns: {list(df.columns)}")
        print(
            f"  This may not be the correct report. Expected 'Player Detail | upload format'."
        )
        sys.exit(1)

    return df


def load_bc_info(filepath: str) -> pd.DataFrame:
    """
    Load the "Player Photo BC Info" report from Sports Affinity.

    This report has per-player birth certificate status:
    - Birth Certificate Uploaded (date) — file was uploaded
    - Birth Certificate Verified (date) — someone reviewed and verified it

    If either date is present, the player has a BC on file.
    """
    print(f"  Loading BC info: {filepath}")
    df = pd.read_excel(filepath, header=1)
    print(f"  → {len(df)} players loaded")

    # Standardize column names (report has long names)
    col_map = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if cl == "first name":
            col_map[c] = "First Name"
        elif cl == "last name":
            col_map[c] = "Last Name"
        elif cl == "dob":
            col_map[c] = "DOB"
        elif cl == "birth certificate uploaded":
            col_map[c] = "BC_Uploaded"
        elif cl == "birth certificate verified by district/state":
            col_map[c] = "BC_Verified_District"
        elif cl == "birth certificate verified by league/club":
            col_map[c] = "BC_Verified_League"
        elif cl == "birth certificate verified by":
            col_map[c] = "BC_Verified_By"
        elif cl == "birth certificate verified":
            col_map[c] = "BC_Verified"
    df.rename(columns=col_map, inplace=True)

    # Count BC statuses
    has_uploaded = df["BC_Uploaded"].notna().sum() if "BC_Uploaded" in df.columns else 0
    has_verified = df["BC_Verified"].notna().sum() if "BC_Verified" in df.columns else 0
    has_either = (
        (df.get("BC_Uploaded", pd.Series(dtype="object")).notna())
        | (df.get("BC_Verified", pd.Series(dtype="object")).notna())
    ).sum()
    has_neither = len(df) - has_either

    print(f"  → BC Verified: {has_verified}, BC Uploaded: {has_uploaded}")
    print(f"  → Total with BC on file: {has_either}, No BC: {has_neither}")

    return df


# =========================================================
#  DATA TRANSFORMATION
# =========================================================


def format_phone(phone) -> str:
    """
    Normalize phone to XXX-XXX-XXXX format.
    Handles (XXX) XXX-XXXX, XXX-XXX-XXXX, XXXXXXXXXX, 1XXXXXXXXXX.
    Validates area code per NANP rules (must start with 2-9).
    Returns empty string for invalid numbers so PM doesn't reject the row.
    """
    if pd.isna(phone) or not phone:
        return ""
    phone = str(phone).strip()
    digits = re.sub(r"\D", "", phone)
    # Strip leading country code
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        area_code = digits[0]
        # NANP: area codes cannot start with 0 or 1
        if area_code in ("0", "1"):
            return ""  # Invalid — blank it rather than fail import
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    # Non-standard length — return empty rather than import garbage
    return ""


def format_zip(zipcode) -> str:
    """Truncate ZIP+4 to 5-digit ZIP."""
    if pd.isna(zipcode) or not zipcode:
        return ""
    z = str(zipcode).strip()
    # Handle ZIP+4 format (91607-3715 → 91607)
    if "-" in z:
        z = z.split("-")[0]
    # Handle float conversion (91607.0 → 91607)
    if "." in z:
        z = z.split(".")[0]
    return z


def format_birth_date(dob) -> str:
    """Ensure birth date is in MM/DD/YYYY format for PlayMetrics."""
    if pd.isna(dob) or not dob:
        return ""
    dob_str = str(dob).strip()
    # Already in MM/DD/YYYY?
    if re.match(r"^\d{2}/\d{2}/\d{4}$", dob_str):
        return dob_str
    # Try parsing various formats
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%y", "%d/%m/%Y"):
        try:
            dt = datetime.strptime(dob_str, fmt)
            return dt.strftime("%m/%d/%Y")
        except ValueError:
            continue
    # Try pandas
    try:
        ts = pd.Timestamp(dob_str)
        if not pd.isna(ts):
            return ts.strftime("%m/%d/%Y")
    except Exception:
        pass
    return dob_str


def safe_str(val) -> str:
    """Convert a value to a clean string, handling NaN and None."""
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()


# =========================================================
#  CORE PROCESSING
# =========================================================


def build_import_data(
    player_upload: pd.DataFrame,
    player_apps: pd.DataFrame,
    season_name: str = "",
    exclude_aged_out: bool = True,
    bc_strict: bool = False,
) -> Tuple[List[Dict], List[Dict], Dict, List]:
    """
    Build PlayMetrics import rows from Sports Affinity data.

    bc_strict controls how birth certificate status is determined:
      False (default): BC Uploaded OR BC Verified = has BC on file.
            Use this if your region considers an uploaded BC sufficient
            even if nobody manually reviewed it.
      True:  Only BC Verified counts. Players who uploaded a BC but were
            never formally verified go in the non-verified file.

    Returns:
        (bc_verified_rows, non_verified_rows, stats_dict, phone_corrections)
    """
    stats = {
        "total_input": len(player_upload),
        "aged_out": 0,
        "bc_verified": 0,
        "non_verified": 0,
        "missing_email": 0,
        "missing_dob": 0,
        "phones_blanked": 0,
        "bc_mode": (
            "strict (verified only)" if bc_strict else "standard (uploaded or verified)"
        ),
    }
    phone_corrections = []

    # Build BC lookup from BC Info report
    bc_lookup = set()
    if not player_apps.empty:
        for _, row in player_apps.iterrows():
            bc_uploaded = row.get("BC_Uploaded")
            bc_verified = row.get("BC_Verified")
            if bc_strict:
                has_bc = pd.notna(bc_verified)
            else:
                has_bc = pd.notna(bc_uploaded) or pd.notna(bc_verified)
            if has_bc:
                key = (
                    safe_str(row.get("First Name", "")).lower()
                    + "|"
                    + safe_str(row.get("Last Name", "")).lower()
                    + "|"
                    + safe_str(row.get("DOB", ""))
                )
                bc_lookup.add(key)

    bc_verified = []
    non_verified = []

    for _, record in player_upload.iterrows():
        first = safe_str(record.get("PlayerFirstName"))
        last = safe_str(record.get("PlayerLastName"))
        dob = safe_str(record.get("DOB"))
        gender = safe_str(record.get("Gender"))

        # Skip if missing critical fields
        if not first or not last:
            continue
        if not dob:
            stats["missing_dob"] += 1
            continue

        # Age eligibility check
        if exclude_aged_out:
            try:
                dob_dt = pd.to_datetime(dob, format="%m/%d/%Y", errors="coerce")
                if pd.notna(dob_dt) and dob_dt < pd.Timestamp(OLDEST_ELIGIBLE_DOB):
                    stats["aged_out"] += 1
                    continue
            except Exception:
                pass

        # Check parent email
        parent1_email = safe_str(record.get("FatherEmailAddress"))
        if not parent1_email:
            stats["missing_email"] += 1

        # Format phone numbers and track invalid ones
        raw_p1_phone = safe_str(record.get("FatherCellPhone"))
        raw_p2_phone = safe_str(record.get("MotherCellPhone"))
        p1_phone = format_phone(raw_p1_phone)
        p2_phone = format_phone(raw_p2_phone)

        # Track phones that were blanked due to invalid area codes
        if raw_p1_phone and not p1_phone:
            stats["phones_blanked"] += 1
            phone_corrections.append(
                f"  {first} {last}: parent1 '{raw_p1_phone}' → blanked (invalid area code)"
            )
        if raw_p2_phone and not p2_phone:
            stats["phones_blanked"] += 1
            phone_corrections.append(
                f"  {first} {last}: parent2 '{raw_p2_phone}' → blanked (invalid area code)"
            )

        # Build the PM row
        # Team is always blank — SA data has prior season teams that don't
        # exist in PM. Teams get built fresh after registration and draft.
        row = {
            "team": "",
            "season_id": "",
            "season": season_name,
            "player_first_name": first,
            "player_last_name": last,
            "gender": gender,  # Already M/F from SA
            "birth_date": format_birth_date(dob),
            "age_group": "",
            "position": "",
            "number": "",
            "Foot": "",
            "parent1_email": parent1_email,
            "parent1_first_name": safe_str(record.get("FatherFirstName")),
            "parent1_last_name": safe_str(record.get("FatherLastName")),
            "parent1_mobile_number": p1_phone,
            "parent2_email": safe_str(record.get("MotherEmailAddress")),
            "parent2_first_name": safe_str(record.get("MotherFirstName")),
            "parent2_last_name": safe_str(record.get("MotherLastName")),
            "parent2_mobile_number": p2_phone,
            "street": safe_str(record.get("Address1")),
            "city": safe_str(record.get("City")),
            "state": safe_str(record.get("State")),
            "zip": format_zip(record.get("ZIPCode")),
        }

        # Check BC status
        player_key = first.lower() + "|" + last.lower() + "|" + dob
        if player_key in bc_lookup:
            bc_verified.append(row)
            stats["bc_verified"] += 1
        else:
            non_verified.append(row)
            stats["non_verified"] += 1

    # Sort both lists by last name, first name
    bc_verified.sort(
        key=lambda r: (r["player_last_name"].lower(), r["player_first_name"].lower())
    )
    non_verified.sort(
        key=lambda r: (r["player_last_name"].lower(), r["player_first_name"].lower())
    )

    return bc_verified, non_verified, stats, phone_corrections


def merge_seasons(dataframes: List[pd.DataFrame]) -> pd.DataFrame:
    """
    Merge player data from multiple seasons.
    Deduplicates by player name + DOB, keeping the most recent record's contact info.
    """
    if len(dataframes) == 1:
        return dataframes[0]

    combined = pd.concat(dataframes, ignore_index=True)
    print(f"  Combined {len(combined)} records from {len(dataframes)} seasons")

    # Create dedup key
    combined["_dedup_key"] = (
        combined["PlayerFirstName"].str.strip().str.lower()
        + "|"
        + combined["PlayerLastName"].str.strip().str.lower()
        + "|"
        + combined["DOB"].astype(str)
    )

    # Keep the first occurrence (most recent if files were loaded newest-first)
    before = len(combined)
    combined = combined.drop_duplicates(subset="_dedup_key", keep="first")
    dupes = before - len(combined)
    if dupes > 0:
        print(f"  → Removed {dupes} duplicate players across seasons")
    print(f"  → {len(combined)} unique players after merge")

    return combined


# =========================================================
#  FILE OUTPUT
# =========================================================


def write_csv(filepath: str, rows: List[Dict]):
    """Write rows to PlayMetrics CSV format."""
    Path(filepath).parent.mkdir(parents=True, exist_ok=True)
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=PM_COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def print_stats(
    stats: Dict, bc_verified: List, non_verified: List, phone_corrections: List = None
):
    """Print a summary report."""
    print()
    print("=" * 60)
    print("  PLAYMETRICS IMPORT SUMMARY")
    print("=" * 60)
    print(f"  BC Mode:                         {stats.get('bc_mode', 'standard')}")
    print(f"  Total players in source file:    {stats['total_input']}")
    if stats["aged_out"]:
        print(f"  Removed (aged out):              {stats['aged_out']}")
    if stats["missing_dob"]:
        print(f"  Skipped (no DOB):                {stats['missing_dob']}")
    if stats["missing_email"]:
        print(f"  ⚠️  Missing parent email:         {stats['missing_email']}")
    if stats["phones_blanked"]:
        print(f"  ⚠️  Invalid phones blanked:        {stats['phones_blanked']}")
    print(f"  ─────────────────────────────────")
    print(f"  BC Verified (send to PM):        {len(bc_verified)}")
    print(f"  Non-verified (import yourself):  {len(non_verified)}")
    print(f"  Total for import:                {len(bc_verified) + len(non_verified)}")
    print()
    if phone_corrections:
        print("  PHONE CORRECTIONS (blanked — invalid area codes):")
        print("  " + "─" * 56)
        for line in phone_corrections:
            print(line)
        print()
        print(
            f"  These {len(phone_corrections)} phone number(s) had area codes starting"
        )
        print(f"  with 0 or 1 (invalid per NANP). They've been blanked in")
        print(f"  the CSV so PlayMetrics won't reject the import. Parents")
        print(f"  will update their phone when they accept the invite.")
        print()


# =========================================================
#  CORE PROCESSING (called by CLI and GUI)
# =========================================================


def process_and_write(
    upload_files: List[str],
    bc_files: List[str],
    output_dir: str = "playmetrics_output",
    bc_strict: bool = False,
) -> Dict:
    """
    Core processing: load files, merge, build import CSVs, write output.
    Returns a results dict with file paths and stats.
    """
    run_time = datetime.now()
    log_lines = []
    log_lines.append("=" * 60)
    log_lines.append("  PLAYMETRICS IMPORT — RUN LOG")
    log_lines.append("=" * 60)
    log_lines.append(f"  Date/Time: {run_time.strftime('%Y-%m-%d %H:%M:%S')}")
    bc_mode_label = (
        "Strict (verified only)" if bc_strict else "Standard (uploaded or verified)"
    )
    log_lines.append(f"  BC Mode: {bc_mode_label}")
    log_lines.append("")

    # ── Log input files ──
    log_lines.append("  INPUT FILES")
    log_lines.append("  " + "─" * 56)
    log_lines.append(f"  Player Upload files ({len(upload_files)}):")
    for f in upload_files:
        log_lines.append(f"    • {os.path.abspath(f)}")
    log_lines.append(f"  BC Info files ({len(bc_files)}):")
    for f in bc_files:
        log_lines.append(f"    • {os.path.abspath(f)}")
    if not bc_files:
        log_lines.append(f"    (none — all players will be non-verified)")
    log_lines.append("")

    # ── Load player upload files ──
    upload_dfs = []
    for f in upload_files:
        df = load_player_upload(f)
        upload_dfs.append(df)
        log_lines.append(f"  Loaded {len(df)} players from {os.path.basename(f)}")

    bc_dfs = []
    for f in bc_files:
        df = load_bc_info(f)
        bc_dfs.append(df)
        bc_uploaded = (
            df["BC_Uploaded"].notna().sum() if "BC_Uploaded" in df.columns else 0
        )
        bc_verified_count = (
            df["BC_Verified"].notna().sum() if "BC_Verified" in df.columns else 0
        )
        log_lines.append(
            f"  Loaded {len(df)} BC records from {os.path.basename(f)}"
            f" (verified: {bc_verified_count}, uploaded: {bc_uploaded})"
        )
    log_lines.append("")

    # ── Merge seasons ──
    if len(upload_dfs) > 1:
        print()
        print("  Merging seasons...")
        combined_count = sum(len(df) for df in upload_dfs)
        player_data = merge_seasons(upload_dfs)
        dupes_removed = combined_count - len(player_data)
        log_lines.append("  SEASON MERGE")
        log_lines.append("  " + "─" * 56)
        log_lines.append(
            f"  Combined records: {combined_count} from {len(upload_dfs)} seasons"
        )
        log_lines.append(f"  Duplicates removed: {dupes_removed}")
        log_lines.append(f"  Unique players after merge: {len(player_data)}")
        log_lines.append("")
    else:
        player_data = upload_dfs[0]

    # ── Merge BC info data ──
    if bc_dfs:
        if len(bc_dfs) > 1:
            bc_data = pd.concat(bc_dfs, ignore_index=True)
            bc_data["_key"] = (
                bc_data["First Name"].str.strip().str.lower()
                + "|"
                + bc_data["Last Name"].str.strip().str.lower()
                + "|"
                + bc_data["DOB"].astype(str)
            )
            bc_data["_has_verified"] = bc_data.get(
                "BC_Verified", pd.Series(dtype="object")
            ).notna()
            bc_data["_has_uploaded"] = bc_data.get(
                "BC_Uploaded", pd.Series(dtype="object")
            ).notna()
            bc_data = bc_data.sort_values(
                ["_has_verified", "_has_uploaded"], ascending=False
            )
            bc_before = len(bc_data)
            bc_data = bc_data.drop_duplicates(subset="_key", keep="first")
            log_lines.append(f"  BC info merged: {bc_before} → {len(bc_data)} unique")
            log_lines.append("")
        else:
            bc_data = bc_dfs[0]
    else:
        bc_data = pd.DataFrame()

    print()
    print("  Building PlayMetrics import files...")
    print("─" * 40)

    bc_verified, non_verified, stats, phone_corrections = build_import_data(
        player_data, bc_data, bc_strict=bc_strict
    )

    print_stats(stats, bc_verified, non_verified, phone_corrections)

    # ── Log processing results ──
    log_lines.append("  PROCESSING RESULTS")
    log_lines.append("  " + "─" * 56)
    log_lines.append(f"  Total players in source: {stats['total_input']}")
    if stats["aged_out"]:
        log_lines.append(f"  Aged out (removed): {stats['aged_out']}")
    if stats["missing_dob"]:
        log_lines.append(f"  Missing DOB (skipped): {stats['missing_dob']}")
    if stats["missing_email"]:
        log_lines.append(f"  Missing parent email: {stats['missing_email']}")
    log_lines.append(f"  BC Verified: {len(bc_verified)}")
    log_lines.append(f"  Non-verified: {len(non_verified)}")
    log_lines.append(f"  Total for import: {len(bc_verified) + len(non_verified)}")
    log_lines.append("")

    # ── Log phone corrections ──
    if phone_corrections:
        log_lines.append("  PHONE CORRECTIONS (blanked — invalid area codes)")
        log_lines.append("  " + "─" * 56)
        for line in phone_corrections:
            log_lines.append(line)
        log_lines.append("")

    # ── Write CSVs ──
    print("  Writing CSV files...")
    print("─" * 40)

    timestamp = run_time.strftime("%Y%m%d")
    os.makedirs(output_dir, exist_ok=True)
    results = {"output_dir": output_dir, "files": []}

    log_lines.append("  OUTPUT FILES")
    log_lines.append("  " + "─" * 56)

    if bc_verified:
        verified_file = os.path.join(
            output_dir, f"playmetrics_bc_verified_{timestamp}.csv"
        )
        write_csv(verified_file, bc_verified)
        results["files"].append(verified_file)
        print(f"  ✅ BC Verified:    {verified_file} ({len(bc_verified)} players)")
        print(f"     → Send this file to success@playmetrics.com")
        log_lines.append(
            f"  BC Verified:   {verified_file} ({len(bc_verified)} players)"
        )

    if non_verified:
        non_verified_file = os.path.join(
            output_dir, f"playmetrics_non_verified_{timestamp}.csv"
        )
        write_csv(non_verified_file, non_verified)
        results["files"].append(non_verified_file)
        print(f"  ✅ Non-verified:   {non_verified_file} ({len(non_verified)} players)")
        print(f"     → Import this file yourself in PlayMetrics admin.")
        log_lines.append(
            f"  Non-verified:  {non_verified_file} ({len(non_verified)} players)"
        )

    all_players = bc_verified + non_verified
    if all_players:
        combined_file = os.path.join(
            output_dir, f"playmetrics_all_players_{timestamp}.csv"
        )
        write_csv(combined_file, all_players)
        results["files"].append(combined_file)
        print(f"  📋 Combined:      {combined_file} ({len(all_players)} players)")
        print(f"     → Reference copy. Do not import.")
        log_lines.append(
            f"  Combined:      {combined_file} ({len(all_players)} players)"
        )

    # ── Write log file ──
    log_file = os.path.join(output_dir, f"import_log_{timestamp}.txt")
    log_lines.append("")
    log_lines.append("=" * 60)
    with open(log_file, "w", encoding="utf-8") as f:
        f.write("\n".join(log_lines) + "\n")
    results["files"].append(log_file)
    print(f"  📋 Log:           {log_file}")

    print()
    results["stats"] = stats
    results["bc_verified_count"] = len(bc_verified)
    results["non_verified_count"] = len(non_verified)
    return results


# =========================================================
#  GUI
# =========================================================


def launch_gui():
    """Launch the tkinter GUI with optional drag-and-drop support."""
    import tkinter as tk
    from tkinter import ttk, filedialog, scrolledtext

    # Try to load drag-and-drop support
    has_dnd = False
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES

        has_dnd = True
    except ImportError:
        pass

    class ImportApp:
        def __init__(self, root):
            self.root = root
            self.root.title("AYSO PlayMetrics Player Import Tool")
            self.root.minsize(700, 600)

            # File lists
            self.upload_files = []
            self.bc_files = []

            self._build_ui()

        def _build_ui(self):
            # ── Header ──
            header = tk.Frame(self.root, bg="#0a2351", padx=20, pady=15)
            header.pack(fill="x")
            tk.Label(
                header,
                text="AYSO PlayMetrics Player Import Tool",
                font=("Helvetica", 16, "bold"),
                fg="white",
                bg="#0a2351",
            ).pack(anchor="w")
            tk.Label(
                header,
                text="Convert Sports Affinity reports → PlayMetrics CSV",
                font=("Helvetica", 10),
                fg="#93c5fd",
                bg="#0a2351",
            ).pack(anchor="w")

            # ── Main content ──
            content = ttk.Frame(self.root, padding=20)
            content.pack(fill="both", expand=True)

            # Player Upload files
            ttk.Label(
                content, text="Player Upload Files", font=("Helvetica", 11, "bold")
            ).grid(row=0, column=0, sticky="w", pady=(0, 5))
            ttk.Label(
                content,
                text='SA report: "Player Detail | upload format"',
                font=("Helvetica", 9),
            ).grid(row=1, column=0, sticky="w")

            upload_frame = tk.Frame(content, bg="#f0f4f8", relief="groove", bd=2)
            upload_frame.grid(row=2, column=0, sticky="ew", pady=(2, 5))

            self.upload_listbox = tk.Listbox(
                upload_frame,
                height=3,
                width=70,
                font=("Courier", 9),
                bg="#f0f4f8",
                relief="flat",
                highlightthickness=0,
            )
            self.upload_listbox.pack(fill="both", expand=True, padx=5, pady=5)

            drop_hint = "Drag files here" if has_dnd else ""
            self.upload_hint = tk.Label(
                upload_frame,
                text=f"Drop .xlsx files here or use Add Files →" if has_dnd else "",
                font=("Helvetica", 9, "italic"),
                fg="#999",
                bg="#f0f4f8",
            )
            if has_dnd:
                self.upload_hint.pack(pady=(0, 5))

            upload_btns = ttk.Frame(content)
            upload_btns.grid(row=2, column=1, padx=(10, 0))
            ttk.Button(
                upload_btns, text="Add Files...", command=self._add_upload_files
            ).pack(fill="x", pady=1)
            ttk.Button(
                upload_btns, text="Clear", command=self._clear_upload_files
            ).pack(fill="x", pady=1)

            # BC Info files
            ttk.Label(
                content,
                text="Birth Certificate Info Files",
                font=("Helvetica", 11, "bold"),
            ).grid(row=3, column=0, sticky="w", pady=(15, 5))
            ttk.Label(
                content, text='SA report: "Player Photo BC Info"', font=("Helvetica", 9)
            ).grid(row=4, column=0, sticky="w")

            bc_frame = tk.Frame(content, bg="#f0f4f8", relief="groove", bd=2)
            bc_frame.grid(row=5, column=0, sticky="ew", pady=(2, 5))

            self.bc_listbox = tk.Listbox(
                bc_frame,
                height=3,
                width=70,
                font=("Courier", 9),
                bg="#f0f4f8",
                relief="flat",
                highlightthickness=0,
            )
            self.bc_listbox.pack(fill="both", expand=True, padx=5, pady=5)

            self.bc_hint = tk.Label(
                bc_frame,
                text=f"Drop .xlsx files here or use Add Files →" if has_dnd else "",
                font=("Helvetica", 9, "italic"),
                fg="#999",
                bg="#f0f4f8",
            )
            if has_dnd:
                self.bc_hint.pack(pady=(0, 5))

            bc_btns = ttk.Frame(content)
            bc_btns.grid(row=5, column=1, padx=(10, 0))
            ttk.Button(bc_btns, text="Add Files...", command=self._add_bc_files).pack(
                fill="x", pady=1
            )
            ttk.Button(bc_btns, text="Clear", command=self._clear_bc_files).pack(
                fill="x", pady=1
            )

            # Register drag-and-drop targets
            if has_dnd:
                upload_frame.drop_target_register(DND_FILES)
                upload_frame.dnd_bind("<<Drop>>", self._drop_upload)
                upload_frame.dnd_bind(
                    "<<DragEnter>>", lambda e: upload_frame.configure(bg="#dbeafe")
                )
                upload_frame.dnd_bind(
                    "<<DragLeave>>", lambda e: upload_frame.configure(bg="#f0f4f8")
                )

                bc_frame.drop_target_register(DND_FILES)
                bc_frame.dnd_bind("<<Drop>>", self._drop_bc)
                bc_frame.dnd_bind(
                    "<<DragEnter>>", lambda e: bc_frame.configure(bg="#dbeafe")
                )
                bc_frame.dnd_bind(
                    "<<DragLeave>>", lambda e: bc_frame.configure(bg="#f0f4f8")
                )

            # Tip
            tip_text = (
                "Tip: Select multiple files at once for multi-season imports. "
                "The tool deduplicates automatically."
            )
            tip = ttk.Label(
                content, text=tip_text, font=("Helvetica", 9), foreground="#666"
            )
            tip.grid(row=6, column=0, columnspan=2, sticky="w", pady=(5, 10))

            # BC verification mode
            bc_frame = ttk.Frame(content)
            bc_frame.grid(row=7, column=0, columnspan=2, sticky="w", pady=(0, 5))

            self.bc_strict_var = tk.BooleanVar(value=False)
            self.bc_check = ttk.Checkbutton(
                bc_frame,
                text="Strict BC mode — only count verified birth certificates (not just uploaded)",
                variable=self.bc_strict_var,
            )
            self.bc_check.pack(anchor="w")
            ttk.Label(
                bc_frame,
                text="Default (unchecked): uploaded or verified = BC on file.  Strict: only verified by a reviewer counts.",
                font=("Helvetica", 8),
                foreground="#999",
            ).pack(anchor="w", padx=(22, 0))

            # Run button
            self.run_btn = ttk.Button(
                content, text="▶  Run Import", command=self._run_import
            )
            self.run_btn.grid(row=8, column=0, columnspan=2, pady=(5, 10), sticky="ew")

            # Output log
            ttk.Label(content, text="Output", font=("Helvetica", 11, "bold")).grid(
                row=9, column=0, sticky="w", pady=(5, 5)
            )

            self.log = scrolledtext.ScrolledText(
                content,
                height=15,
                width=80,
                font=("Courier", 9),
                state="disabled",
                wrap="word",
                bg="#1e1e1e",
                fg="#d4d4d4",
                insertbackground="white",
            )
            self.log.grid(row=10, column=0, columnspan=2, sticky="nsew", pady=(0, 10))

            # Open folder button (hidden until output exists)
            self.open_btn = ttk.Button(
                content, text="Open Output Folder", command=self._open_output_folder
            )
            self.open_btn.grid(row=11, column=0, columnspan=2, sticky="ew")
            self.open_btn.grid_remove()

            # Grid weights
            content.columnconfigure(0, weight=1)
            content.rowconfigure(10, weight=1)

            self.output_dir = None

        def _parse_drop_data(self, data):
            """Parse dropped file paths from tkdnd event data."""
            files = []
            # tkdnd wraps paths with spaces in {braces} on Windows
            # and separates multiple files with spaces
            current = ""
            in_braces = False
            for char in data:
                if char == "{":
                    in_braces = True
                elif char == "}":
                    in_braces = False
                    if current:
                        files.append(current)
                        current = ""
                elif char == " " and not in_braces:
                    if current:
                        files.append(current)
                        current = ""
                else:
                    current += char
            if current:
                files.append(current)
            # Filter to .xlsx only
            return [f for f in files if f.lower().endswith(".xlsx")]

        def _drop_upload(self, event):
            files = self._parse_drop_data(event.data)
            for f in files:
                if f not in self.upload_files:
                    self.upload_files.append(f)
                    self.upload_listbox.insert("end", os.path.basename(f))
            if self.upload_files and has_dnd:
                self.upload_hint.pack_forget()
            event.widget.configure(bg="#f0f4f8")

        def _drop_bc(self, event):
            files = self._parse_drop_data(event.data)
            for f in files:
                if f not in self.bc_files:
                    self.bc_files.append(f)
                    self.bc_listbox.insert("end", os.path.basename(f))
            if self.bc_files and has_dnd:
                self.bc_hint.pack_forget()
            event.widget.configure(bg="#f0f4f8")

        def _add_upload_files(self):
            files = filedialog.askopenfilenames(
                title="Select Player Upload files",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            )
            for f in files:
                if f not in self.upload_files:
                    self.upload_files.append(f)
                    self.upload_listbox.insert("end", os.path.basename(f))
            if self.upload_files and has_dnd:
                self.upload_hint.pack_forget()

        def _clear_upload_files(self):
            self.upload_files.clear()
            self.upload_listbox.delete(0, "end")
            if has_dnd:
                self.upload_hint.pack(pady=(0, 5))

        def _add_bc_files(self):
            files = filedialog.askopenfilenames(
                title="Select Player Photo BC Info files",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            )
            for f in files:
                if f not in self.bc_files:
                    self.bc_files.append(f)
                    self.bc_listbox.insert("end", os.path.basename(f))
            if self.bc_files and has_dnd:
                self.bc_hint.pack_forget()

        def _clear_bc_files(self):
            self.bc_files.clear()
            self.bc_listbox.delete(0, "end")
            if has_dnd:
                self.bc_hint.pack(pady=(0, 5))

        def _log_write(self, text):
            self.log.configure(state="normal")
            self.log.insert("end", text + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
            self.root.update_idletasks()

        def _run_import(self):
            if not self.upload_files:
                self._log_write("ERROR: No Player Upload files selected.")
                return

            # Clear log
            self.log.configure(state="normal")
            self.log.delete("1.0", "end")
            self.log.configure(state="disabled")
            self.open_btn.grid_remove()

            self.run_btn.configure(state="disabled")
            self.root.update_idletasks()

            # Redirect stdout to capture print output
            old_stdout = sys.stdout
            sys.stdout = io.StringIO()

            try:
                results = process_and_write(
                    self.upload_files, self.bc_files, bc_strict=self.bc_strict_var.get()
                )
                output = sys.stdout.getvalue()
                self.output_dir = results["output_dir"]
            except Exception as e:
                output = sys.stdout.getvalue()
                output += f"\n\nERROR: {e}"
                import traceback

                output += "\n" + traceback.format_exc()
            finally:
                sys.stdout = old_stdout

            self._log_write(output)

            if self.output_dir:
                self.open_btn.grid()
                self._log_write("─" * 50)
                self._log_write("  Done! Click 'Open Output Folder' to see your files.")

            self.run_btn.configure(state="normal")

        def _open_output_folder(self):
            if not self.output_dir:
                return
            folder = os.path.abspath(self.output_dir)
            if platform.system() == "Windows":
                os.startfile(folder)
            elif platform.system() == "Darwin":
                subprocess.run(["open", folder])
            else:
                subprocess.run(["xdg-open", folder])

    # Create root window — use TkinterDnD if available for drag-and-drop
    if has_dnd:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    app = ImportApp(root)
    root.mainloop()


# =========================================================
#  MAIN
# =========================================================


def main():
    # ── Parse command-line arguments ──
    parser = argparse.ArgumentParser(
        description="AYSO PlayMetrics Player Import Tool",
        epilog="Example: playmetrics_import.py --dir C:\\Users\\sdavis\\sa_reports",
    )
    parser.add_argument(
        "--dir",
        metavar="FOLDER",
        help="Folder containing SA report files. Auto-loads all "
        "playerUpload*.xlsx and Player*Photo*BC*.xlsx files "
        "found, merging multiple seasons automatically.",
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Force interactive command-line mode (skip GUI).",
    )
    parser.add_argument(
        "--strict-bc",
        action="store_true",
        help="Strict BC mode: only count verified birth certificates, "
        "not just uploaded. Default: uploaded or verified both count.",
    )
    args = parser.parse_args()

    # ── Batch mode: --dir ──
    if args.dir:
        folder = args.dir.strip().strip('"')
        if not os.path.isdir(folder):
            print(f"  ERROR: Folder not found: {folder}")
            sys.exit(1)

        print()
        print("╔══════════════════════════════════════════════════════╗")
        print("║  AYSO PlayMetrics Player Import Tool                ║")
        print("╚══════════════════════════════════════════════════════╝")
        print()
        print(f"  Scanning folder: {folder}")
        print("─" * 40)

        upload_files = sorted(glob.glob(os.path.join(folder, "playerUpload*.xlsx")))
        if not upload_files:
            upload_files = sorted(
                glob.glob(os.path.join(folder, "PlayerDetail*upload*.xlsx"))
            )

        bc_files = sorted(glob.glob(os.path.join(folder, "Player_Photo_BC*.xlsx")))
        if not bc_files:
            bc_files = sorted(glob.glob(os.path.join(folder, "Player Photo BC*.xlsx")))
        if not bc_files:
            bc_files = sorted(glob.glob(os.path.join(folder, "PlayerPhotoBC*.xlsx")))

        if not upload_files:
            print("  ERROR: No playerUpload*.xlsx files found in folder")
            sys.exit(1)

        print(f"  Found {len(upload_files)} player upload file(s):")
        for f in upload_files:
            print(f"    • {os.path.basename(f)}")

        if bc_files:
            print(f"  Found {len(bc_files)} BC info file(s):")
            for f in bc_files:
                print(f"    • {os.path.basename(f)}")
        else:
            print(
                "  ⚠️  No Player Photo BC files found — all players will be non-verified"
            )

        print()
        if args.strict_bc:
            print("  BC Mode: STRICT (only verified birth certificates count)")
        print()
        process_and_write(upload_files, bc_files, bc_strict=args.strict_bc)

        print("─" * 60)
        print("  NEXT STEPS:")
        print("  1. Spot-check 10-15 records against your source data")
        print("  2. Send the BC-verified file to success@playmetrics.com")
        print("  3. Import the non-verified file in PlayMetrics (Players → Import)")
        print("─" * 60)
        print()
        return

    # ── GUI mode (default) ──
    if not args.cli:
        try:
            launch_gui()
            return
        except Exception:
            # tkinter not available (headless server) — fall through to CLI
            pass

    # ── Interactive CLI mode ──
    print()
    print("╔══════════════════════════════════════════════════════╗")
    print("║  AYSO PlayMetrics Player Import Tool                ║")
    print("║  Converts Sports Affinity data → PlayMetrics CSV    ║")
    print("╚══════════════════════════════════════════════════════╝")
    print()
    print("  Tip: Use --dir <folder> to auto-load all files from a folder")
    print()

    print("STEP 1: Locate source files")
    print("─" * 40)

    upload_file = find_file("playerUpload*.xlsx") or find_file(
        "PlayerDetail*upload*.xlsx"
    )
    bc_file = (
        find_file("Player_Photo_BC*.xlsx")
        or find_file("Player Photo BC*.xlsx")
        or find_file("PlayerPhotoBC*.xlsx")
    )

    if upload_file:
        print(f"  Found player upload: {upload_file}")
        resp = input(f"  Use this file? (Y/n): ").strip().lower()
        if resp == "n":
            upload_file = None

    if not upload_file:
        upload_file = (
            input("  Path to 'Player Detail | upload format' Excel file: ")
            .strip()
            .strip('"')
        )
        if not os.path.exists(upload_file):
            print(f"  ERROR: File not found: {upload_file}")
            sys.exit(1)

    if bc_file:
        print(f"  Found BC info: {bc_file}")
        resp = input(f"  Use this file? (Y/n): ").strip().lower()
        if resp == "n":
            bc_file = None

    if not bc_file:
        bc_file = (
            input("  Path to 'Player Photo BC Info' Excel file (or Enter to skip): ")
            .strip()
            .strip('"')
        )
        if bc_file and not os.path.exists(bc_file):
            print(f"  ERROR: File not found: {bc_file}")
            bc_file = None

    print()

    print("STEP 2: Additional seasons (optional)")
    print("─" * 40)
    print("  If you downloaded reports for prior years, you can include them.")
    print()

    upload_files = [upload_file]
    bc_files = [bc_file] if bc_file else []

    while True:
        resp = (
            input("  Add another season's playerUpload file? (y/N): ").strip().lower()
        )
        if resp != "y":
            break
        extra = input("  Path to additional playerUpload file: ").strip().strip('"')
        if os.path.exists(extra):
            upload_files.append(extra)
            extra_bc = (
                input("  Corresponding Player Photo BC Info file (or Enter to skip): ")
                .strip()
                .strip('"')
            )
            if extra_bc and os.path.exists(extra_bc):
                bc_files.append(extra_bc)
        else:
            print(f"  File not found: {extra}")

    print()
    process_and_write(upload_files, bc_files, bc_strict=args.strict_bc)

    print("─" * 60)
    print("  NEXT STEPS:")
    print("  1. Spot-check 10-15 records against your source data")
    print("  2. Send the BC-verified file to success@playmetrics.com")
    print("  3. Import the non-verified file in PlayMetrics (Players → Import)")
    print("─" * 60)
    print()


if __name__ == "__main__":
    main()
