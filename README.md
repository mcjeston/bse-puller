# BSE Puller

[Download the latest installer](https://github.com/mcjeston/bse-puller/raw/main/dist/BsePullerSetup.exe)

Windows desktop app for pulling approved BILL Spend and Expense transactions into the accounting CSV layout used by the team.

## What the app does

- Calls `GET /v3/spend/transactions`
- Uses this BILL filter in the API request:
  - `type:ne:DECLINE,syncStatus:eq:NOT_SYNCED,complete:eq:true`
- Applies these local checks after the API response:
  - `accountingIntegrationTransactions` must be empty
  - `reviewers` must include `ADMIN` with `APPROVED`
  - duplicate `id` values are removed
  - conflicting GL account values are logged and excluded from any later sync-eligible set
- Converts the results into the team’s accounting CSV column layout
- Saves exports into the local `CSV exports` folder beside the app
- Keeps the newest 4 previous CSV exports as backups
- Opens the CSV automatically after saving
- Shows a final reminder with transaction count and amount total so the user can mark those transactions as synced manually in BILL Spend and Expense

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
- `Previous Exports` opens the local `CSV exports` folder
- `Reset API Key` removes the saved BILL API token for the current Windows user
- `Uninstall` removes the installed app, Start Menu shortcut, saved user settings, and installed export folder

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
