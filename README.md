# BSE Puller

[Download the latest installer](https://github.com/mcjeston/bse-puller/releases/download/v2026.03.31.2/BsePullerSetup.exe)

[View the latest release](https://github.com/mcjeston/bse-puller/releases/tag/v2026.03.31.2)

Windows desktop app for pulling approved BILL Spend and Expense transactions into the accounting CSV layout used by the team.

See `LEDGER.md` for a plain-English reference on how the program is built, installed, and used.

## What the app does

- Pull Transactions (API-based)
  - Calls `GET /v3/spend/transactions`
  - Uses this BILL filter in the API request:
    - `type:ne:DECLINE,syncStatus:eq:NOT_SYNCED,complete:eq:true`
  - Applies these local checks after the API response:
    - `accountingIntegrationTransactions` must be empty
    - `reviewers` must include `ADMIN` with `APPROVED`
    - duplicate `id` values are removed
    - conflicting GL account values are logged and excluded from any later sync-eligible set
  - Converts the results into the team's accounting CSV column layout
  - Saves exports into the local `CSV exports` folder beside the app
  - Keeps the newest 4 previous CSV exports as backups
  - Copies exported data rows (without the header row) to the clipboard for Sage import
  - Shows an import dialog with `Copy Again` and `Done`, then a summary step with `Back` and `Done`
  - Saves the CSV in `CSV exports` without opening the spreadsheet automatically
  - If a pull has no exportable rows, no CSV file is saved and previous backup exports remain unchanged
  - Shows a final reminder with transaction count and amount total so the user can mark those transactions as synced manually in BILL Spend and Expense
- Pull Reimbursements
  - Coming soon (button is disabled while the flow is reworked)
- Checks for GitHub updates automatically once every 24 hours when running from an installed copy

## Modules

- Pull Transactions module
  - Owns the BILL API pull, CSV export, and clipboard/import dialog flow.
- Pull Reimbursements module (disabled)
  - Placeholder module for the future reimbursement export flow.
- Previous Exports module
  - Owns exports folder access and backup retention.
- Settings module
  - Owns update checks, API key reset, uninstall, and log downloads.

## Installer and user settings

- Installer file:
  - `dist\BsePullerSetup.exe`
- Install location:
  - `%LocalAppData%\Programs\BsePuller`
- Per-user settings location:
  - `%LocalAppData%\BsePuller\settings.json`
- The installer asks for the BILL API token during setup
- If the token is not already saved, the app will ask for it on first pull

## In-app actions

- `Pull Transactions` runs the export
- `Pull Reimbursements` is disabled (Coming Soon)
- `Previous Exports` opens the local `CSV exports` folder
- The top-right settings gear includes:
  - `Check for Updates`
  - `Reset API Key`
  - `Download Log File` (saves a text log under `%LocalAppData%\BsePuller\Logs`, keeps the newest 20, and opens the folder)
  - `Uninstall`

## Build locally

Use the local .NET SDK included in this folder:

```powershell
.\.dotnet\dotnet.exe build .\BsePuller.csproj -c Release
.\.dotnet\dotnet.exe publish .\BsePuller.csproj -c Release -o .\bin\Release\net8.0-windows\publish
```

Published app:

```text
bin\Release\net8.0-windows\publish\BsePuller.exe
```

Build installer:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\Build-Installer.ps1 -SkipPublish
```

## Snapshots

Working snapshots are stored in:

```text
snapshots
```
