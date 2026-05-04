# AYSO PlayMetrics Player Import Tool

Converts Sports Affinity player data into PlayMetrics Player Import CSV format for AYSO regions migrating from SportsConnect to PlayMetrics.

Built for the AYSO PlayMetrics pilot program (MY2026). Designed to be shared across regions — no configuration needed beyond downloading two reports and running the tool.

## What It Does

- Reads two standard reports from the Sports Affinity (Association) site
- Maps player names, DOB, gender, parent contacts, addresses, and team assignments to the PlayMetrics Player Import CSV format
- Determines birth certificate status from the Applications report (`Media=B` = BC on file)
- Splits output into **BC-verified** (send to PlayMetrics) and **non-verified** (import yourself)
- Handles phone formatting, ZIP code truncation, age eligibility filtering, and Jamboree team cleanup
- Supports **multi-season merge** — include prior years to capture returning players who sat out a season

## Quick Start

### Requirements

- Python 3.8+ with `pandas` and `openpyxl`
- Two Excel reports from Sports Affinity (see [Downloading Reports](#downloading-reports))

```bash
pip install pandas openpyxl
```

### Usage

```bash
# Place the SA reports in the same folder as the script
python playmetrics_import.py
```

The tool auto-detects the Excel files, prompts you through any options, and produces output CSVs in a `playmetrics_output/` folder.

### For Non-Python Users

A standalone Windows `.exe` is available — no Python installation needed. See [Releases](../../releases) or contact your region's registrar.

## Downloading Reports

Both reports come from the **Sports Affinity (Association) site**, not the SportsConnect Registration site.

### Navigate to Sports Affinity

1. Log in to **Sports Connect**
2. Click **"Change Login"** → select the **Association site**
3. Go to **Players / Admins** → **Player Lookup**
4. Set **Season** to your membership year (e.g., "2025-2026 MY2025")
5. Set **Region** to your region
6. Set **Age Group** to "Multiple" (select all)

### Report 1: "Player Detail | upload format"

This is the **primary data source** — player names, DOB, gender, parent contacts, addresses, teams.

1. Click the **Report** dropdown → select **"Player Detail | upload format"**
2. Click **Print** → in the report window, select **"Excel"** → click **Export**
3. Save as `playerUpload.xlsx`

### Report 2: "All Player Applications Detail"

Provides birth certificate status via the **Media** column (`B` = birth certificate on file).

1. Click the **Report** dropdown → select **"All Player Applications Detail"**
2. Click **Print** → select **"Excel"** → click **Export**
3. Save as `playerApplications.xlsx`

### Multi-Season (Optional)

To capture returning players from prior years, change the **Season dropdown** to each prior year and download both reports again. The tool merges them automatically, deduplicates by player name + DOB, and keeps the most recent contact info.

## Output Files

| File | What It Is | What To Do |
|------|-----------|------------|
| `playmetrics_bc_verified_YYYYMMDD.csv` | Players with BC on file | Email to `success@playmetrics.com` — they import and flag as verified |
| `playmetrics_non_verified_YYYYMMDD.csv` | Players without BC | Import yourself in PlayMetrics admin (Players → Import) |
| `playmetrics_all_players_YYYYMMDD.csv` | Combined reference | Do not import — for your records only |

## Field Mapping

| PlayMetrics Column | Sports Affinity Source | Notes |
|---|---|---|
| `player_first_name` | `PlayerFirstName` | Already split |
| `player_last_name` | `PlayerLastName` | Already split |
| `gender` | `Gender` | Already M/F |
| `birth_date` | `DOB` | Already MM/DD/YYYY |
| `parent1_email` | `FatherEmailAddress` | "Father" = Parent1, not gendered |
| `parent1_first_name` | `FatherFirstName` | Already split |
| `parent1_last_name` | `FatherLastName` | Already split |
| `parent1_mobile_number` | `FatherCellPhone` | Formatted to XXX-XXX-XXXX |
| `parent2_email` | `MotherEmailAddress` | "Mother" = Parent2 |
| `parent2_first_name` | `MotherFirstName` | Already split |
| `parent2_last_name` | `MotherLastName` | Already split |
| `parent2_mobile_number` | `MotherCellPhone` | Formatted to XXX-XXX-XXXX |
| `street` | `Address1` | — |
| `city` | `City` | — |
| `state` | `State` | — |
| `zip` | `ZIPCode` | Truncated from ZIP+4 to 5 digits |
| `team` | `TeamName` | Cleared for Jamboree teams |

## Important Notes

- **"Father"/"Mother" = Parent1/Parent2.** The SA column names are not gendered — many "Father" entries have female names. The tool maps them as parent1 and parent2.

- **Birth certificate status** uses the `Media` column from the Applications report. `Media=B` means a BC was used for age verification. We treat "BC on file" as verified regardless of whether someone manually clicked a verify button in SportsConnect.

- **Jamboree/4U players** are included with their team name cleared. They may age up to 5U and should keep their BC verification status.

- **Aged-out players** (born before Aug 1, 2007 for Fall 2026) are automatically excluded. Adjust `OLDEST_ELIGIBLE_DOB` in the script for different seasons.

- **Import sequencing is critical.** Complete all imports BEFORE opening registration. Importing after registration opens creates account conflicts.

## Import Sequence

```
1. PlayMetrics reviews and approves your program
2. Run this tool → produce CSVs
3. Send BC-verified CSV to PlayMetrics (they import)
4. Import non-verified CSV yourself
5. Verify imported data (spot-check 10-15 records)
6. Send invites to imported families
7. THEN open registration
```

## Distribution

For regions without Python, build a standalone `.exe`:

```bash
pip install pyinstaller
pyinstaller --onefile --name PlayMetricsImport playmetrics_import.py
```

The `.exe` goes in `dist/`. Distribute it with `README.txt` — recipients just drop their Excel files in the same folder and double-click.

## License

MIT — see [LICENSE](LICENSE)

## Author

Steve Davis, Registrar — AYSO Region 58
Built with assistance from Claude (Anthropic)

## Links

- [AYSO Region 58 Migration Resources](https://sdavis9248.github.io/playmetrics-migration-region58/)
- [PlayMetrics Help Center](https://help.playmetrics.com)
- [PlayMetrics Support](mailto:success@playmetrics.com)
