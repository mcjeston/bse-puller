# BSE Puller

Windows desktop app for exporting filtered BILL Spend and Expense transactions into the accounting CSV layout used by the team.

## Download

- Direct installer download:
  - [BsePullerSetup.exe](https://github.com/mcjeston/bse-puller/raw/main/dist/BsePullerSetup.exe)
- GitHub repository:
  - [mcjeston/bse-puller](https://github.com/mcjeston/bse-puller)

## Current behavior

- Opens as a Windows GUI.
- Calls `GET /v3/spend/transactions`.
- Uses this API-side filter:
  - `type:ne:DECLINE,syncStatus:eq:NOT_SYNCED,complete:eq:true`
- Applies these local checks after the API response:
  - `accountingIntegrationTransactions` must be empty
  - `reviewers` must include `ADMIN` with `APPROVED`
  - duplicate `id` values are skipped
  - GL account merge conflicts are logged and excluded from later sync steps
- Saves accounting-formatted CSV files to the local `CSV exports` folder beside the app.
- Keeps the newest 4 older CSV exports as backups.
- Opens the CSV automatically after saving.
- Shows a reminder dialog with the transaction count and charge total so the user can mark those transactions as synced manually in BILL.

## API token

- The API token is no longer stored in source code.
- Each Windows user has their own token stored here:

```text
%LocalAppData%\BsePuller\settings.json
```

- The installer prompts for the token during install.
- If the app is launched without a saved token, it prompts the user on first pull and saves it for that Windows user.

## Build the app

Use the local SDK in this folder:

```powershell
.\.dotnet\dotnet.exe build BsePuller.csproj -c Release
.\.dotnet\dotnet.exe publish BsePuller.csproj -c Release --no-restore
```

Published app location:

```text
bin\Release\net8.0-windows\publish\BsePuller.exe
```

## Build the installer

This project includes a simple per-user installer based on Windows `IExpress`.

Build it with:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\Build-Installer.ps1
```

Installer output:

```text
dist\BsePullerSetup.exe
```

Installer behavior:

- prompts for the BILL API token during install
- installs the app to `%LocalAppData%\Programs\BsePuller`
- writes the token to `%LocalAppData%\BsePuller\settings.json`
- creates a Start Menu shortcut for `BSE Puller`

## Snapshots

Saved snapshots of working builds are in:

```text
snapshots
```
