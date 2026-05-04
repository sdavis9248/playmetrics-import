================================================================================
  AYSO PlayMetrics Player Import Tool
  Version 1.0 — May 2026
================================================================================

  Converts Sports Affinity player data into PlayMetrics Player Import CSV
  format for AYSO regions migrating from SportsConnect to PlayMetrics.

  Built by Steve Davis, Registrar — AYSO Region 58
  With assistance from Claude (Anthropic)

================================================================================
  WHAT THIS TOOL DOES
================================================================================

  Takes two Excel reports you download from Sports Affinity and produces
  ready-to-import CSV files for PlayMetrics. It handles:

    - Player names, DOB, gender (already in M/F format)
    - Parent 1 and Parent 2 contact info (email, phone, name)
    - Mailing addresses
    - Team assignments
    - Birth certificate status (from the Applications report)
    - Phone number formatting (strips to XXX-XXX-XXXX)
    - ZIP code cleanup (truncates ZIP+4 to 5 digits)
    - Age eligibility filtering (removes aged-out players)
    - Multi-season merge and deduplication (optional)

  The tool splits the output into two files:

    1. BC-VERIFIED players   → Send to PlayMetrics for import
    2. NON-VERIFIED players  → Import yourself in PlayMetrics admin

  PlayMetrics flags the BC-verified players so returning families skip the
  birth certificate upload during registration. New families are prompted
  to upload automatically.

================================================================================
  WHAT YOU NEED
================================================================================

  1. This tool (PlayMetricsImport.exe)
  2. Two Excel files from Sports Affinity (instructions below)

  That's it. No Python installation or other software needed.

================================================================================
  HOW TO DOWNLOAD THE REPORTS FROM SPORTS AFFINITY
================================================================================

  Both reports come from the same page. You just pick a different report
  from the dropdown each time.

  STEP 1: Navigate to Sports Affinity
  ------------------------------------
  a) Log in to Sports Connect at your region's URL
  b) Click "Change Login" in the top-right menu bar
  c) Select the Association site (this switches you from the Registration
     site to the Sports Affinity / Association platform)
  d) Click "Players / Admins" in the top menu
  e) Click "Player Lookup"

  STEP 2: Set Your Filters
  -------------------------
  a) Season dropdown (left of "PLAYER LOOKUP"): Select the current
     membership year (e.g., "2025-2026 MY2025")
  b) Select Region: Choose your region
  c) Select Age Group: Set to "Multiple" and check ALL age groups
  d) Leave everything else at defaults

  STEP 3: Download Report 1 — "Player Detail | upload format"
  -------------------------------------------------------------
  This is your PRIMARY data source (player names, parents, addresses).

  a) Click the "Report" dropdown (top right, above the player list)
  b) Select "Player Detail | upload format"
  c) Click "Print" — a new window opens with the report
  d) In the report window, find the format dropdown (top left)
  e) Select "Excel" from the format dropdown
  f) Click "Export"
  g) Save the file — it downloads as "playerUpload.xlsx"

  STEP 4: Download Report 2 — "All Player Applications Detail"
  --------------------------------------------------------------
  This report tells us which players have a birth certificate on file.

  a) Go back to the Player Lookup page (close the report window)
  b) Keep the same filters (same season, same region, all age groups)
  c) Click the "Report" dropdown again
  d) Select "All Player Applications Detail"
  e) Click "Print" → select "Excel" → click "Export"
  f) Save the file — it downloads as "playerApplications.xlsx"

================================================================================
  HOW TO RUN THE TOOL
================================================================================

  STEP 1: Put Everything in One Folder
  --------------------------------------
  Create a folder and put these files in it:

    PlayMetricsImport.exe
    playerUpload.xlsx
    playerApplications.xlsx

  STEP 2: Run the Tool
  ----------------------
  Double-click PlayMetricsImport.exe (or run from command line).

  The tool will:
    - Auto-detect the Excel files in the folder
    - Ask you to confirm each file
    - Ask if you have additional seasons to merge (optional)
    - Ask for a season name (optional — can leave blank)
    - Process the data and show a summary
    - Create output files in a "playmetrics_output" subfolder

  STEP 3: Review the Output
  ---------------------------
  Open the CSV files in Excel and spot-check a few records:

    - Player names look correct?
    - Birth dates are MM/DD/YYYY?
    - Phone numbers are XXX-XXX-XXXX?
    - ZIP codes are 5 digits?
    - Parent emails are populated?

================================================================================
  WHAT TO DO WITH THE OUTPUT FILES
================================================================================

  The tool creates files in a "playmetrics_output" folder:

  FILE 1: playmetrics_bc_verified_YYYYMMDD.csv
  -----------------------------------------------
  Players with a birth certificate on file.

  → Email this file to: success@playmetrics.com
    Subject: "BC-Verified Player Import — [Your Region Name]"
    They will import these players and flag them as BC-verified.
    Wait for confirmation before proceeding.

  FILE 2: playmetrics_non_verified_YYYYMMDD.csv
  ------------------------------------------------
  Players WITHOUT a birth certificate on file.

  → Import this file yourself in PlayMetrics admin:
    Players → More Actions → Import Players → upload the CSV
    These players will be prompted to upload a BC during registration.

  NOTE: If all your players have BC on file (which is common — Region 58
  had 100% coverage), the tool won't create a non-verified file. That's
  fine — just send the one file to PlayMetrics and you're done.

  FILE 3: playmetrics_all_players_YYYYMMDD.csv
  -----------------------------------------------
  Combined reference copy. Do NOT import this — it's for your records.

================================================================================
  IMPORTING MULTIPLE SEASONS (OPTIONAL)
================================================================================

  Want to capture players who sat out last season but might come back?
  Download the reports for prior years too.

  HOW:
    1. Go back to Player Lookup on Sports Affinity
    2. Change the Season dropdown to the prior year
       (e.g., "2024-2025 MY2024")
    3. Download both reports again
    4. Save them with different names so you can tell them apart
       (e.g., playerUpload_2024.xlsx, playerApplications_2024.xlsx)

  When you run the tool, it will ask if you want to add additional
  seasons. Say "y" and point it to the extra files. The tool will:

    - Combine all seasons
    - Remove duplicate players (by name + date of birth)
    - Keep the most recent contact information
    - Any player with a BC in ANY season gets flagged as verified

  RECOMMENDATION: Current year + 1 prior year is a good balance.
  Going back 3+ years has diminishing returns.

================================================================================
  IMPORTANT NOTES
================================================================================

  ABOUT "FATHER" AND "MOTHER" FIELDS:
  The Sports Affinity columns say "Father" and "Mother" but they are NOT
  gendered — they are simply Parent 1 and Parent 2. Many "Father" entries
  have female names and vice versa. The tool maps them as parent1 and
  parent2 in the PlayMetrics import.

  ABOUT BIRTH CERTIFICATE STATUS:
  The tool uses the "Media" column from the Applications report. Media=B
  means a birth certificate was used as the age verification document.
  We treat "BC on file" as verified — even if nobody manually clicked a
  verify button in Sports Connect. This is a deliberate choice to maximize
  the number of returning families who don't have to re-upload.

  ABOUT JAMBOREE/4U PLAYERS:
  Players on Jamboree teams (typically 4U) are included in the import.
  They may be aging up to 5U and should keep their BC verification
  status. Like all imported players, the team column is blank — prior
  season teams don't carry over to PlayMetrics. Teams are built fresh
  after registration and draft.

  ABOUT AGED-OUT PLAYERS:
  Players born before August 1, 2007 are automatically excluded (they're
  too old for the 19U division in Fall 2026). Adjust the cutoff date in
  the script if your season has different age boundaries.

  ABOUT INVITES:
  After PlayMetrics imports the file, the players sit in the system
  waiting. You choose when to send invitations. Imported families receive
  an invite to verify their PlayMetrics account. You can assign coaches
  to teams and organize everything BEFORE families accept their invites —
  you don't need to wait for them.

  SEQUENCING IS CRITICAL:
  Complete ALL imports BEFORE opening registration. If you open
  registration first and then import, you'll create account conflicts
  when a parent who already registered overlaps with an imported record.

  The correct sequence is:
    1. PlayMetrics reviews and approves your program
    2. Import players (this tool)
    3. Send invites to imported families
    4. THEN open registration

================================================================================
  WINDOWS SMARTSCREEN WARNING
================================================================================

  The first time you run PlayMetricsImport.exe, Windows Defender
  SmartScreen may show a warning:

    "Windows protected your PC"
    "Microsoft Defender SmartScreen prevented an unrecognized app
     from starting."

  This is normal for any unsigned .exe file. It does NOT mean the file
  is harmful.

  To proceed:
    1. Click "More info"
    2. Click "Run anyway"

  This only happens once — Windows remembers your choice.

================================================================================
  TROUBLESHOOTING
================================================================================

  "No Enrollment_Details file found" or "No playerUpload file found"
    → Make sure the Excel files are in the same folder as the .exe
    → The tool looks for files starting with "playerUpload" or
      "PlayerDetail" — don't rename them to something unrecognizable

  "pandas is required"
    → This should not happen with the .exe version. If you're running
      the Python script directly, install dependencies:
      pip install pandas openpyxl

  The tool shows 0 players
    → Check that you selected the correct Region in Sports Affinity
    → Check that you selected "Multiple" for Age Group with all checked
    → Check the Season dropdown — is it the right membership year?

  Phone numbers look wrong
    → The tool strips to 10 digits and formats as XXX-XXX-XXXX
    → International numbers or extensions may not format correctly
    → These can be fixed manually in the CSV before import

  ZIP codes have decimals (91607.0)
    → The tool handles this automatically. If you see it in the source
      Excel file, that's an Excel formatting issue — the CSV output
      will be correct.

================================================================================
  QUESTIONS OR ISSUES
================================================================================

  Contact: Steve Davis, Registrar — AYSO Region 58
  Email:   registrar@ayso58.org

  For PlayMetrics import support:
  Email:   success@playmetrics.com

================================================================================
