# APP Billing System

## Project Overview

Anesthesiologist billing data entry system for the Anesthesia Practice Plan (APP) at Royal Columbian Hospital (RCH) and Eagle Ridge Hospital (ERH) in British Columbia, Canada. Built in Microsoft Excel with VBA, designed for 15+ concurrent users via per-user daily files on a network share.

## Tech Stack

- **Language**: VBA (Visual Basic for Applications)
- **Platform**: Microsoft Excel (.xlsm macro-enabled workbook)
- **Storage**: Network file share with per-user daily Excel files
- **Target OS**: Windows 10/11

## Repository Structure

```
APP_Billing/
  APPBilling.xlsm          # Main workbook (UI, sheets, data)
  SETUP_INSTRUCTIONS.txt    # Full import/deployment guide
  .env                      # Local config (gitignored)
  .env.example              # Config template for new deployments
  src/                      # VBA source files (import into workbook)
    Module1.bas             # Main logic: Submit, Reset, Sync, Setup
    modConfig.bas           # Settings, network path, folder management
    modNetworkIO.bas        # Per-user daily file read/write, sync
    modPDFReport.bas        # ORReportingForm population, PDF export
    modConsolidate.bas      # Daily/range data consolidation
    modSecurity.bas         # Superuser auth (Windows user + password)
    modHelpers.bas           # Utility subs
    frmSaveData.frm         # Data entry form
    frmPrntData.frm         # PDF report generation form
    frmSuperUser.frm        # Superuser admin panel
    frmDailyExport.frm      # Daily export form
```

## Architecture

### Multi-User Strategy
Each user gets their own copy of `APPBilling.xlsm`. Data is saved to **per-user daily files** on the network share (`UserName_YYYYMMDD.xlsx`), eliminating file lock conflicts. Records are always saved locally first, then synced to the network.

### Network Folder Layout
```
\\server\APP_Billing\
  Data\YYYY-MM\UserName_YYYYMMDD.xlsx    # Per-user daily data
  DailyExports\AllUsers_YYYYMMDD.xlsx    # Consolidated daily reports
  PDFReports\UserName_YYYYMMDD.pdf       # Individual PDF reports
  Config\SuperUsers.xlsx                  # Authorized superuser list
```

### Security Model
Superuser access requires **both** Windows username verification (checked against `Config\SuperUsers.xlsx`) **and** a shared password. Access levels: `Admin` (full control) or `ReadOnly`.

## DailyDatabase Schema (28 columns, A-AB)

| Col | Field | Col | Field |
|-----|-------|-----|-------|
| A | S # (serial) | O | Fee Modifier 3 |
| B | Anesthesiologist | P | Resuscitation |
| C | Site (RCH/ERH) | Q | Obstetrics |
| D | Date of Service | R | Acute Pain |
| E | Shift Name | S | Diagnostic/Chronic Pain |
| F | On Call | T | Miscellaneous Fee Items |
| G | Shift Type (OR/Out of OR) | U | WCB Number |
| H | Surgical Procedure Code | V | Side |
| I | Procedure Start Time | W | Diagnostic Code |
| J | Procedure Finish Time | X | Injury Code |
| K | Maximum IC Level | Y | Date of Injury |
| L | Consults | Z | Submitted By |
| M | Fee Modifier 1 | AA | Submitted On |
| N | Fee Modifier 2 | AB | Sync Status |

## Coding Conventions

- **Date format**: DD/MM/YYYY throughout all code and UI
- **Error handling**: All public subs/functions use `On Error GoTo ErrHandler` pattern
- **Network operations**: Always use retry logic (3 attempts, 2-second delay)
- **Local-first**: Save to local `DailyDatabase` sheet before attempting network sync
- **Column constants**: Use named constants from `modNetworkIO` (e.g., `COL_ANESTH`, `COL_DATE`) — never hardcode column numbers
- **Form controls**: VBA code references controls by name (e.g., `lstAnesth`, `txtDteOfSer`); visual layout must be created in VBA Editor form designer
- **File naming**: Network files use `UserName_YYYYMMDD.xlsx` pattern; sanitize names with `GetUserDailyFileName()`

## Key Module Responsibilities

| Module | Responsibility | Key Functions |
|--------|---------------|---------------|
| `Module1` | Form launchers, Submit/Reset, Sync | `Submit()`, `Reset()`, `SyncNow()`, `InitialSetup()` |
| `modConfig` | Settings read/write, folder creation | `GetNetworkPath()`, `EnsureNetworkFolders()`, `GetCurrentUser()` |
| `modNetworkIO` | Network file I/O, sync tracking | `SaveToNetwork()`, `ReadUserDailyData()`, `SyncPendingRecords()` |
| `modPDFReport` | PDF generation from ORReportingForm | `GenerateDailyPDF()`, `PopulateORForm()`, `ExportToPDF()` |
| `modConsolidate` | Merge user files into one workbook | `ConsolidateDailyData()`, `ConsolidateDateRange()` |
| `modSecurity` | Authentication and access control | `AuthenticateSuperUser()`, `IsSuperUser()`, `IsAdmin()` |

## Development Workflow

1. **VBA source files** in `src/` are the source of truth for code
2. Changes are made to `.bas`/`.frm` files, then imported into the workbook via VBA Editor (Alt+F11 > File > Import)
3. The `.xlsm` workbook should be re-saved after importing updated modules
4. `.frm` files contain code only — form visual layout (control positions, sizes) must be configured in the VBA Editor form designer

## Git Workflow

- **Commit and push after every major milestone or feature addition**
- Branch: `master`
- Remote: `origin` → `git@github.com:jramsden58/APP_Billing.git`
- `.env` is gitignored; `.env.example` is the committed template
- Never commit Excel temp files (`~$*.xls*`)

## Testing

1. Set `NETWORK_SHARE_PATH` in `.env` to a local test folder for offline testing
2. Run `InitialSetup` macro to create folder structure and Settings sheet
3. Test concurrent access by opening two instances of the workbook
4. Verify PDF output matches `ORReportingForm` layout (6 procedures per page)
5. Test network failure: disconnect, save records, reconnect, run `SyncNow`
