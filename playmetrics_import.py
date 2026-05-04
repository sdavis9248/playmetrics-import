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

    2. "All Player Applications Detail" from Sports Affinity (playerApplications.xlsx)
       - Used to determine birth certificate status via the "Media" column.
       - Media=B means a birth certificate was used for age verification.

HOW TO DOWNLOAD THE REPORTS:
    1. Log in to Sports Connect → click "Change Login" → select the Association site
    2. Navigate to Players/Admins → Player Lookup
    3. Set the season dropdown (e.g., "2025-2026 MY2025")
    4. Set Region to your region
    5. Set Age Group to "Multiple" (select all)
    6. From the Report dropdown, select "Player Detail | upload format" → Print
    7. Save the Excel file
    8. Repeat with "All Player Applications Detail" → Print → Save
    9. For multi-year imports: change the season dropdown and repeat for each year

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
from datetime import datetime
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

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
    'team', 'season_id', 'season', 'player_first_name', 'player_last_name',
    'gender', 'birth_date', 'age_group', 'position', 'number', 'Foot',
    'parent1_email', 'parent1_first_name', 'parent1_last_name', 'parent1_mobile_number',
    'parent2_email', 'parent2_first_name', 'parent2_last_name', 'parent2_mobile_number',
    'street', 'city', 'state', 'zip'
]

# Oldest eligible birth date for Fall 2026 (Aug 1 cutoff, 19U max)
# Adjust this for your season
OLDEST_ELIGIBLE_DOB = '2007-08-01'

# =========================================================
#  DATA LOADING
# =========================================================

def find_file(pattern: str, search_dir: str = '.') -> Optional[str]:
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
    required = ['PlayerFirstName', 'PlayerLastName', 'Gender', 'DOB',
                'FatherEmailAddress', 'FatherFirstName', 'FatherLastName']
    missing = [c for c in required if c not in df.columns]
    if missing:
        print(f"  ⚠️  Missing expected columns: {missing}")
        print(f"  Available columns: {list(df.columns)}")
        print(f"  This may not be the correct report. Expected 'Player Detail | upload format'.")
        sys.exit(1)
    
    return df


def load_player_applications(filepath: str) -> pd.DataFrame:
    """
    Load the "All Player Applications Detail" report from Sports Affinity.
    
    Used to extract birth certificate status from the "Media" column.
    Media=B means a birth certificate was used for age verification.
    """
    print(f"  Loading player applications: {filepath}")
    df = pd.read_excel(filepath)
    print(f"  → {len(df)} application records loaded")
    
    if 'Media' not in df.columns:
        print(f"  ⚠️  'Media' column not found — cannot determine BC status")
        print(f"  Available columns: {list(df.columns)}")
        return pd.DataFrame()
    
    # Deduplicate to one record per player (most recent application)
    if 'Appl. Date' in df.columns:
        df = df.sort_values('Appl. Date', ascending=False)
    
    df['_key'] = (df['First Name'].str.strip().str.lower() + '|' + 
                  df['Last Name'].str.strip().str.lower() + '|' + 
                  df['DOB'].astype(str))
    df = df.drop_duplicates(subset='_key', keep='first')
    
    bc_count = (df['Media'].str.upper().str.startswith('B', na=False)).sum()
    print(f"  → {len(df)} unique players, {bc_count} with BC on file (Media=B)")
    
    return df


# =========================================================
#  DATA TRANSFORMATION
# =========================================================

def format_phone(phone) -> str:
    """
    Normalize phone to XXX-XXX-XXXX format.
    Handles (XXX) XXX-XXXX, XXX-XXX-XXXX, XXXXXXXXXX, 1XXXXXXXXXX.
    """
    if pd.isna(phone) or not phone:
        return ''
    phone = str(phone).strip()
    digits = re.sub(r'\D', '', phone)
    if len(digits) == 11 and digits.startswith('1'):
        digits = digits[1:]
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:6]}-{digits[6:]}"
    return phone


def format_zip(zipcode) -> str:
    """Truncate ZIP+4 to 5-digit ZIP."""
    if pd.isna(zipcode) or not zipcode:
        return ''
    z = str(zipcode).strip()
    # Handle ZIP+4 format (91607-3715 → 91607)
    if '-' in z:
        z = z.split('-')[0]
    # Handle float conversion (91607.0 → 91607)
    if '.' in z:
        z = z.split('.')[0]
    return z


def format_birth_date(dob) -> str:
    """Ensure birth date is in MM/DD/YYYY format for PlayMetrics."""
    if pd.isna(dob) or not dob:
        return ''
    dob_str = str(dob).strip()
    # Already in MM/DD/YYYY?
    if re.match(r'^\d{2}/\d{2}/\d{4}$', dob_str):
        return dob_str
    # Try parsing various formats
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%m/%d/%y', '%d/%m/%Y'):
        try:
            dt = datetime.strptime(dob_str, fmt)
            return dt.strftime('%m/%d/%Y')
        except ValueError:
            continue
    # Try pandas
    try:
        ts = pd.Timestamp(dob_str)
        if not pd.isna(ts):
            return ts.strftime('%m/%d/%Y')
    except Exception:
        pass
    return dob_str


def safe_str(val) -> str:
    """Convert a value to a clean string, handling NaN and None."""
    if pd.isna(val) or val is None:
        return ''
    return str(val).strip()


# =========================================================
#  CORE PROCESSING
# =========================================================

def build_import_data(player_upload: pd.DataFrame,
                      player_apps: pd.DataFrame,
                      season_name: str = '',
                      exclude_aged_out: bool = True) -> Tuple[List[Dict], List[Dict], Dict]:
    """
    Build PlayMetrics import rows from Sports Affinity data.
    
    All age-eligible players are included. Players on Jamboree teams have their
    team name cleared (those teams won't exist in PM) but are still imported
    so returning families don't have to re-upload their birth certificate.
    
    Returns:
        (bc_verified_rows, non_verified_rows, stats_dict)
    """
    stats = {
        'total_input': len(player_upload),
        'aged_out': 0,
        'bc_verified': 0,
        'non_verified': 0,
        'missing_email': 0,
        'missing_dob': 0,
    }
    
    # Build BC lookup from applications
    bc_lookup = set()
    if not player_apps.empty and 'Media' in player_apps.columns:
        for _, row in player_apps.iterrows():
            media = safe_str(row.get('Media', '')).upper()
            if media.startswith('B'):
                key = (safe_str(row.get('First Name', '')).lower() + '|' + 
                       safe_str(row.get('Last Name', '')).lower() + '|' + 
                       safe_str(row.get('DOB', '')))
                bc_lookup.add(key)
    
    bc_verified = []
    non_verified = []
    
    for _, record in player_upload.iterrows():
        first = safe_str(record.get('PlayerFirstName'))
        last = safe_str(record.get('PlayerLastName'))
        dob = safe_str(record.get('DOB'))
        gender = safe_str(record.get('Gender'))
        
        # Skip if missing critical fields
        if not first or not last:
            continue
        if not dob:
            stats['missing_dob'] += 1
            continue
        
        # Age eligibility check
        if exclude_aged_out:
            try:
                dob_dt = pd.to_datetime(dob, format='%m/%d/%Y', errors='coerce')
                if pd.notna(dob_dt) and dob_dt < pd.Timestamp(OLDEST_ELIGIBLE_DOB):
                    stats['aged_out'] += 1
                    continue
            except Exception:
                pass
        
        # Check parent email
        parent1_email = safe_str(record.get('FatherEmailAddress'))
        if not parent1_email:
            stats['missing_email'] += 1
        
        # Build the PM row
        # Team is always blank — SA data has prior season teams that don't
        # exist in PM. Teams get built fresh after registration and draft.
        row = {
            'team': '',
            'season_id': '',
            'season': season_name,
            'player_first_name': first,
            'player_last_name': last,
            'gender': gender,  # Already M/F from SA
            'birth_date': format_birth_date(dob),
            'age_group': '',
            'position': '',
            'number': '',
            'Foot': '',
            'parent1_email': parent1_email,
            'parent1_first_name': safe_str(record.get('FatherFirstName')),
            'parent1_last_name': safe_str(record.get('FatherLastName')),
            'parent1_mobile_number': format_phone(record.get('FatherCellPhone')),
            'parent2_email': safe_str(record.get('MotherEmailAddress')),
            'parent2_first_name': safe_str(record.get('MotherFirstName')),
            'parent2_last_name': safe_str(record.get('MotherLastName')),
            'parent2_mobile_number': format_phone(record.get('MotherCellPhone')),
            'street': safe_str(record.get('Address1')),
            'city': safe_str(record.get('City')),
            'state': safe_str(record.get('State')),
            'zip': format_zip(record.get('ZIPCode')),
        }
        
        # Check BC status
        player_key = first.lower() + '|' + last.lower() + '|' + dob
        if player_key in bc_lookup:
            bc_verified.append(row)
            stats['bc_verified'] += 1
        else:
            non_verified.append(row)
            stats['non_verified'] += 1
    
    # Sort both lists by last name, first name
    bc_verified.sort(key=lambda r: (r['player_last_name'].lower(), r['player_first_name'].lower()))
    non_verified.sort(key=lambda r: (r['player_last_name'].lower(), r['player_first_name'].lower()))
    
    return bc_verified, non_verified, stats


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
    combined['_dedup_key'] = (
        combined['PlayerFirstName'].str.strip().str.lower() + '|' +
        combined['PlayerLastName'].str.strip().str.lower() + '|' +
        combined['DOB'].astype(str)
    )
    
    # Keep the first occurrence (most recent if files were loaded newest-first)
    before = len(combined)
    combined = combined.drop_duplicates(subset='_dedup_key', keep='first')
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
    with open(filepath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=PM_COLUMNS, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(rows)


def print_stats(stats: Dict, bc_verified: List, non_verified: List):
    """Print a summary report."""
    print()
    print("=" * 60)
    print("  PLAYMETRICS IMPORT SUMMARY")
    print("=" * 60)
    print(f"  Total players in source file:    {stats['total_input']}")
    if stats['aged_out']:
        print(f"  Removed (aged out):              {stats['aged_out']}")
    if stats['missing_dob']:
        print(f"  Skipped (no DOB):                {stats['missing_dob']}")
    if stats['missing_email']:
        print(f"  ⚠️  Missing parent email:         {stats['missing_email']}")
    print(f"  ─────────────────────────────────")
    print(f"  BC Verified (send to PM):        {len(bc_verified)}")
    print(f"  Non-verified (import yourself):  {len(non_verified)}")
    print(f"  Total for import:                {len(bc_verified) + len(non_verified)}")
    print()


# =========================================================
#  MAIN
# =========================================================

def main():
    print()
    print("╔══════════════════════════════════════════════════════╗")
    print("║  AYSO PlayMetrics Player Import Tool                ║")
    print("║  Converts Sports Affinity data → PlayMetrics CSV    ║")
    print("╚══════════════════════════════════════════════════════╝")
    print()
    
    # ── Step 1: Locate files ──
    print("STEP 1: Locate source files")
    print("─" * 40)
    
    # Try auto-detection first
    upload_file = find_file('playerUpload*.xlsx') or find_file('PlayerDetail*upload*.xlsx')
    apps_file = find_file('playerApplication*.xlsx') or find_file('AllPlayerApplications*.xlsx')
    
    if upload_file:
        print(f"  Found player upload: {upload_file}")
        resp = input(f"  Use this file? (Y/n): ").strip().lower()
        if resp == 'n':
            upload_file = None
    
    if not upload_file:
        upload_file = input("  Path to 'Player Detail | upload format' Excel file: ").strip().strip('"')
        if not os.path.exists(upload_file):
            print(f"  ERROR: File not found: {upload_file}")
            sys.exit(1)
    
    if apps_file:
        print(f"  Found player applications: {apps_file}")
        resp = input(f"  Use this file? (Y/n): ").strip().lower()
        if resp == 'n':
            apps_file = None
    
    if not apps_file:
        apps_file = input("  Path to 'All Player Applications Detail' Excel file (or Enter to skip): ").strip().strip('"')
        if apps_file and not os.path.exists(apps_file):
            print(f"  ERROR: File not found: {apps_file}")
            apps_file = None
    
    print()
    
    # ── Step 2: Multi-season? ──
    print("STEP 2: Additional seasons (optional)")
    print("─" * 40)
    print("  If you downloaded reports for prior years, you can include them")
    print("  to capture returning players. The tool will deduplicate and keep")
    print("  the most recent contact info.")
    print()
    
    additional_uploads = []
    additional_apps = []
    
    while True:
        resp = input("  Add another season's playerUpload file? (y/N): ").strip().lower()
        if resp != 'y':
            break
        extra = input("  Path to additional playerUpload file: ").strip().strip('"')
        if os.path.exists(extra):
            additional_uploads.append(extra)
            extra_apps = input("  Corresponding playerApplications file (or Enter to skip): ").strip().strip('"')
            if extra_apps and os.path.exists(extra_apps):
                additional_apps.append(extra_apps)
        else:
            print(f"  File not found: {extra}")
    
    print()
    
    # ── Step 3: Season name ──
    print("STEP 3: Configuration")
    print("─" * 40)
    season = input("  Season name for PM import (e.g., '2026 Fall Core', or Enter for blank): ").strip()
    print()
    
    # ── Step 4: Load and process ──
    print("STEP 4: Loading data")
    print("─" * 40)
    
    # Load primary files
    upload_dfs = [load_player_upload(upload_file)]
    for extra in additional_uploads:
        upload_dfs.append(load_player_upload(extra))
    
    apps_dfs = []
    if apps_file:
        apps_dfs.append(load_player_applications(apps_file))
    for extra in additional_apps:
        apps_dfs.append(load_player_applications(extra))
    
    # Merge seasons if needed
    if len(upload_dfs) > 1:
        print()
        print("  Merging seasons...")
        player_data = merge_seasons(upload_dfs)
    else:
        player_data = upload_dfs[0]
    
    # Merge application data
    if apps_dfs:
        if len(apps_dfs) > 1:
            apps_data = pd.concat(apps_dfs, ignore_index=True)
            # Deduplicate applications across seasons
            apps_data['_key'] = (apps_data['First Name'].str.strip().str.lower() + '|' +
                                 apps_data['Last Name'].str.strip().str.lower() + '|' +
                                 apps_data['DOB'].astype(str))
            if 'Appl. Date' in apps_data.columns:
                apps_data = apps_data.sort_values('Appl. Date', ascending=False)
            apps_data = apps_data.drop_duplicates(subset='_key', keep='first')
        else:
            apps_data = apps_dfs[0]
    else:
        apps_data = pd.DataFrame()
    
    print()
    
    # ── Step 5: Build import files ──
    print("STEP 5: Building PlayMetrics import files")
    print("─" * 40)
    
    bc_verified, non_verified, stats = build_import_data(
        player_data, apps_data, season_name=season
    )
    
    print_stats(stats, bc_verified, non_verified)
    
    # ── Step 6: Write output ──
    print("STEP 6: Writing CSV files")
    print("─" * 40)
    
    timestamp = datetime.now().strftime('%Y%m%d')
    output_dir = 'playmetrics_output'
    os.makedirs(output_dir, exist_ok=True)
    
    if bc_verified:
        verified_file = os.path.join(output_dir, f'playmetrics_bc_verified_{timestamp}.csv')
        write_csv(verified_file, bc_verified)
        print(f"  ✅ BC Verified:    {verified_file} ({len(bc_verified)} players)")
        print(f"     → Send this file to success@playmetrics.com")
        print(f"       They will import and flag these players as BC-verified.")
    
    if non_verified:
        non_verified_file = os.path.join(output_dir, f'playmetrics_non_verified_{timestamp}.csv')
        write_csv(non_verified_file, non_verified)
        print(f"  ✅ Non-verified:   {non_verified_file} ({len(non_verified)} players)")
        print(f"     → Import this file yourself in PlayMetrics admin.")
        print(f"       These players will be prompted to upload a BC during registration.")
    
    # Also write a combined file for reference
    all_players = bc_verified + non_verified
    if all_players:
        combined_file = os.path.join(output_dir, f'playmetrics_all_players_{timestamp}.csv')
        write_csv(combined_file, all_players)
        print(f"  📋 Combined:      {combined_file} ({len(all_players)} players)")
        print(f"     → Reference copy. Do not import this — use the split files above.")
    
    print()
    print("─" * 60)
    print("  NEXT STEPS:")
    print(f"  1. Review the CSV files in the '{output_dir}/' folder")
    print(f"  2. Spot-check 10-15 records against your source data")
    print(f"  3. Send the BC-verified file to success@playmetrics.com")
    print(f"  4. Import the non-verified file in PlayMetrics (Players → Import)")
    print(f"  5. Send invites to imported families")
    print(f"  6. THEN open registration")
    print("─" * 60)
    print()


if __name__ == '__main__':
    main()
