# BSE Puller Ledger

This file is a plain-English reference for how BSE Puller is built, installed, updated, and used.

## What This Program Does

BSE Puller is a Windows desktop app that connects to BILL Spend and Expense, pulls approved transactions that have not been synced yet, converts them into the accounting CSV format the team needs, saves that CSV beside the app, and copies the export rows to the clipboard for Sage import.

## Main Parts Of The Project

- `Program.cs`
  - Starts the Windows desktop app.
- `MainForm.cs`
  - Builds the app window and the buttons the user clicks.
  - Handles pull, open exports, update checks, reset API key, and uninstall actions.
- `BseClient.cs`
  - Calls the BILL API and collects transaction data.
- `UpdateService.cs`
  - Checks GitHub latest release and downloads installer updates.
- `AccountingCsvFormatter.cs`
  - Converts BILL transactions into the accounting CSV columns.
- `RawCsvWriter.cs`
  - Writes the final CSV file.
- `BseSettings.cs`
  - Controls local folders, saved API token location, and install-related paths.
- `Assets\favicon.ico`
  - The icon used for the app window and executable.
- `installer\Build-Installer.ps1`
  - Builds the installer EXE.
- `installer\Install-BsePuller.ps1`
  - Runs during installation and copies the app into the user's Windows profile.

## How To Build The Program

This project uses the local .NET SDK stored in the project folder.

Build the Release version:

```powershell
.\.dotnet\dotnet.exe build .\BsePuller.csproj -c Release
```

This compiles the app and produces the main executable files under:

```text
bin\Release\net8.0-windows
```

Publish the app files:

```powershell
.\.dotnet\dotnet.exe publish .\BsePuller.csproj -c Release -o .\bin\Release\net8.0-windows\publish
```

This creates the folder that is copied into snapshots and used to build the installer:

```text
bin\Release\net8.0-windows\publish
```

Build the installer:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\Build-Installer.ps1 -SkipPublish
```

The installer output file is:

```text
dist\BsePullerSetup.exe
```

## How Installation Works Today

When the installer runs:

1. It checks whether this Windows user already has a saved BILL API token.
2. It prompts for a token only when one is not already saved.
3. It saves the token for that Windows user.
4. It copies the published app files into the user's local Programs folder.
5. It creates a Start menu shortcut.

Install location:

```text
%LocalAppData%\Programs\BsePuller
```

Saved settings location:

```text
%LocalAppData%\BsePuller\settings.json
```

Start menu shortcut folder:

```text
%AppData%\Microsoft\Windows\Start Menu\Programs\BSE Puller
```

## How The User Uses The Program

Normal user flow:

1. Open BSE Puller from the Start menu.
2. Click `Pull Transactions`.
3. Wait for the app to pull approved not-synced BILL transactions.
4. In the first dialog step, confirm the data was copied to clipboard, then paste it in Sage 100 Contractor screen `4-7-7` using card issuer account `21010 - Bill Spend & Expense`.
5. If needed, click `Copy Again` to restore clipboard content.
6. Click `Done` to move to the summary step. If clicked by mistake, click `Back` to return to the copy step.
7. In the summary step, click `Done` and then mark those transactions as synced manually in BILL Spend and Expense.

Other buttons:

- `Previous Exports`
  - Opens the folder that contains earlier CSV exports.
- `Check for Updates`
  - Runs an immediate GitHub release check and offers install when a newer version exists.
- `Reset API Key`
  - Deletes the saved BILL API token for the current Windows user.
- `Uninstall`
  - Removes the installed app, Start menu shortcut, saved settings, and local export folder for that user.

CSV export folder:

```text
CSV exports
```

This folder lives beside the installed EXE.

## Export and Clipboard Behavior

- The app still writes a CSV file to the local `CSV exports` folder on every export.
- The app does not auto-open the CSV file.
- Clipboard content is row data only (no header row), in tab-separated grid format so paste works directly into Sage cells.
- If clipboard data is lost, the user can click `Copy Again` in the import dialog.
- If no exportable transactions are returned, the app shows a notice and does not change clipboard contents.

## Current Update Model

Installed copies automatically check for updates at startup once every 24 hours.

When a newer release exists, the app can download `BsePullerSetup.exe` from GitHub and launch it.

Users can also click `Check for Updates` in the app for an immediate manual check.

To ship a new version:

1. Change the source files.
2. Build and publish the Release version.
3. Build a new installer.
4. Publish a GitHub release with `BsePullerSetup.exe` attached.
5. Installed copies detect and offer that update.

## Snapshot Pattern

Snapshots are stored in:

```text
snapshots
```

Each snapshot should include:

- `source`
  - The project files needed to rebuild the app later.
- `publish`
  - The published app output.
- `SNAPSHOT_INFO.txt`
  - A short note about what that snapshot represents.
