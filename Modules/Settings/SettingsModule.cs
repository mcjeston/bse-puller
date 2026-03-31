using System.Diagnostics;
using System.Linq;
using System.Text;

namespace BsePuller.Modules.Settings;

internal sealed class SettingsModule
{
    private readonly Form _owner;
    private readonly Action<string> _log;
    private readonly Action<string> _status;
    private readonly Action _closeAction;
    private readonly Button _pullButton;
    private readonly Button _pullReimbursementsButton;
    private readonly Button _openExportsButton;
    private readonly ToolStripMenuItem _checkUpdatesMenuItem;
    private readonly ToolStripMenuItem _resetApiKeyMenuItem;
    private readonly ToolStripMenuItem _downloadLogMenuItem;
    private readonly ToolStripMenuItem _uninstallMenuItem;
    private bool _isUpdateCheckRunning;
    private bool _startupUpdateCheckStarted;

    public SettingsModule(
        Form owner,
        Action<string> log,
        Action<string> status,
        Action closeAction,
        Button pullButton,
        Button pullReimbursementsButton,
        Button openExportsButton,
        ToolStripMenuItem checkUpdatesMenuItem,
        ToolStripMenuItem resetApiKeyMenuItem,
        ToolStripMenuItem downloadLogMenuItem,
        ToolStripMenuItem uninstallMenuItem)
    {
        _owner = owner;
        _log = log;
        _status = status;
        _closeAction = closeAction;
        _pullButton = pullButton;
        _pullReimbursementsButton = pullReimbursementsButton;
        _openExportsButton = openExportsButton;
        _checkUpdatesMenuItem = checkUpdatesMenuItem;
        _resetApiKeyMenuItem = resetApiKeyMenuItem;
        _downloadLogMenuItem = downloadLogMenuItem;
        _uninstallMenuItem = uninstallMenuItem;
    }

    public bool IsUpdateCheckRunning => _isUpdateCheckRunning;

    public async Task RunStartupUpdateCheckIfNeededAsync()
    {
        if (_startupUpdateCheckStarted)
        {
            return;
        }

        _startupUpdateCheckStarted = true;

        if (!BseSettings.IsRunningInstalledCopy())
        {
            _log("Skipping automatic update check because this is not the installed copy.");
            return;
        }

        var lastCheckedUtc = BseSettings.GetLastUpdateCheckUtc();
        if (lastCheckedUtc is not null &&
            DateTimeOffset.UtcNow - lastCheckedUtc.Value < TimeSpan.FromHours(24))
        {
            _log($"Skipping automatic update check. Last checked at {lastCheckedUtc.Value.ToLocalTime():g}.");
            return;
        }

        await CheckForUpdatesAsync(isManual: false);
    }

    public async Task CheckForUpdatesAsync(bool isManual)
    {
        if (_isUpdateCheckRunning)
        {
            if (isManual)
            {
                MessageBox.Show(
                    _owner,
                    "An update check is already in progress.",
                    "Update check in progress",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            return;
        }

        var originalCheckMenuText = _checkUpdatesMenuItem.Text;
        var pullWasEnabled = _pullButton.Enabled;
        var pullReimbursementsWasEnabled = _pullReimbursementsButton.Enabled;
        var openExportsWasEnabled = _openExportsButton.Enabled;
        var checkMenuWasEnabled = _checkUpdatesMenuItem.Enabled;
        var resetWasEnabled = _resetApiKeyMenuItem.Enabled;
        var downloadLogWasEnabled = _downloadLogMenuItem.Enabled;
        var uninstallWasEnabled = _uninstallMenuItem.Enabled;

        _isUpdateCheckRunning = true;
        _pullButton.Enabled = false;
        _pullReimbursementsButton.Enabled = false;
        _openExportsButton.Enabled = false;
        _checkUpdatesMenuItem.Enabled = false;
        _checkUpdatesMenuItem.Text = "Checking...";
        _resetApiKeyMenuItem.Enabled = false;
        _downloadLogMenuItem.Enabled = false;
        _uninstallMenuItem.Enabled = false;

        try
        {
            _log(isManual ? "Checking for updates..." : "Running automatic update check...");
            _status("Checking for updates...");

            var result = await UpdateService.CheckForUpdatesAsync(CancellationToken.None);
            if (!result.CheckedSuccessfully)
            {
                var error = result.ErrorMessage ?? "Unknown error.";
                _log($"Update check failed. {error}");
                _status("Update check failed.");
                if (isManual)
                {
                    MessageBox.Show(
                        _owner,
                        $"Update check failed.{Environment.NewLine}{Environment.NewLine}{error}",
                        "Update check failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                return;
            }

            var latestTag = result.LatestTag ?? "(unknown)";
            if (!result.IsUpdateAvailable)
            {
                _log($"No updates available. Current version: {result.CurrentTag}.");
                _status("No updates available.");
                if (isManual)
                {
                    MessageBox.Show(
                        _owner,
                        $"No updates available. Current version: {result.CurrentTag}.",
                        "Up to date",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                return;
            }

            _log($"Update available. Current: {result.CurrentTag}. Latest: {latestTag}.");
            if (!ShowUpdateAvailableDialog(result.CurrentTag ?? "(unknown)", latestTag))
            {
                _log("Update postponed.");
                _status("Update postponed.");
                return;
            }

            if (string.IsNullOrWhiteSpace(result.DownloadUrl))
            {
                const string missingAssetMessage = "A newer release was found, but it does not include BsePullerSetup.exe.";
                _log(missingAssetMessage);
                _status("Installer missing.");
                MessageBox.Show(
                    _owner,
                    missingAssetMessage,
                    "Update unavailable",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            _status("Downloading update...");
            var downloadResult = await UpdateService.DownloadInstallerAsync(result.DownloadUrl, CancellationToken.None);
            if (!downloadResult.Success)
            {
                var error = downloadResult.ErrorMessage ?? "Unknown error.";
                _log($"Update download failed. {error}");
                _status("Update download failed.");
                MessageBox.Show(
                    _owner,
                    $"Update download failed.{Environment.NewLine}{Environment.NewLine}{error}",
                    "Update download failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(downloadResult.FilePath) || !File.Exists(downloadResult.FilePath))
            {
                const string missingFileMessage = "The update was downloaded, but the installer file could not be found.";
                _log(missingFileMessage);
                _status("Installer missing.");
                MessageBox.Show(
                    _owner,
                    missingFileMessage,
                    "Update unavailable",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = downloadResult.FilePath,
                UseShellExecute = true
            });

            _status("Update installer started. Closing...");
            _log("Update installer started. Closing BSE Puller.");
            _closeAction();
        }
        catch (Exception ex)
        {
            _log($"Update check failed. {ex.Message}");
            _status("Update check failed.");
            if (isManual)
            {
                MessageBox.Show(
                    _owner,
                    $"Update check failed.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                    "Update check failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        finally
        {
            if (BseSettings.IsRunningInstalledCopy())
            {
                try
                {
                    BseSettings.SaveLastUpdateCheckUtc(DateTimeOffset.UtcNow);
                }
                catch (Exception ex)
                {
                    _log($"Warning: could not save update-check timestamp. {ex.Message}");
                }
            }

            _isUpdateCheckRunning = false;
            _pullButton.Enabled = pullWasEnabled;
            _pullReimbursementsButton.Enabled = pullReimbursementsWasEnabled;
            _openExportsButton.Enabled = openExportsWasEnabled;
            _checkUpdatesMenuItem.Enabled = checkMenuWasEnabled;
            _checkUpdatesMenuItem.Text = originalCheckMenuText;
            _resetApiKeyMenuItem.Enabled = resetWasEnabled;
            _downloadLogMenuItem.Enabled = downloadLogWasEnabled;
            _uninstallMenuItem.Enabled = uninstallWasEnabled;
        }
    }

    public void ResetApiKey()
    {
        if (MessageBox.Show(
                _owner,
                "Remove the saved BILL API token for this user?",
                "Reset API Key",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning) != DialogResult.Yes)
        {
            return;
        }

        BseSettings.ClearApiToken();
        _log("Removed the saved BILL API token for this user.");
        _status("API token cleared.");
        MessageBox.Show(
            _owner,
            "The BILL API token has been cleared for this user.",
            "API token cleared",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    public void DownloadLog(string? logContents)
    {
        var folder = BseSettings.GetLogFolder();
        Directory.CreateDirectory(folder);

        var fileName = $"BSE-log-{DateTime.Now:yyyyMMdd-HHmmss}.txt";
        var filePath = Path.Combine(folder, fileName);
        File.WriteAllText(filePath, logContents ?? string.Empty, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        TrimOldLogs(folder, keepCount: 20);

        _log($"Saved log to: {filePath}");
        _status("Log saved.");
        Process.Start(new ProcessStartInfo
        {
            FileName = folder,
            UseShellExecute = true
        });
    }

    public void StartUninstall()
    {
        if (!BseSettings.IsRunningInstalledCopy())
        {
            MessageBox.Show(
                _owner,
                $"Uninstall is only available from the installed copy in:{Environment.NewLine}{BseSettings.GetInstalledAppFolder()}",
                "Uninstall unavailable",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        var confirm = MessageBox.Show(
            _owner,
            "This will remove the installed BSE Puller files, settings, and exports for this Windows user. Continue?",
            "Uninstall BSE Puller",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (confirm != DialogResult.Yes)
        {
            return;
        }

        try
        {
            var helperPath = Path.Combine(
                Path.GetTempPath(),
                $"BsePuller-Uninstall-{Guid.NewGuid():N}.ps1");

            File.WriteAllText(helperPath, BuildUninstallScript(helperPath), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

            Process.Start(new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"{helperPath}\"",
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });

            _log("Started the uninstall helper. Closing BSE Puller.");
            _status("Closing for uninstall...");
            _closeAction();
        }
        catch (Exception ex)
        {
            _log($"Error: could not start uninstall helper. {ex.Message}");
            MessageBox.Show(
                _owner,
                $"Could not start the uninstall helper.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                "Uninstall failed",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    public bool EnsureApiTokenConfigured()
    {
        if (BseSettings.IsConfigured)
        {
            return true;
        }

        using var dialog = new Form
        {
            Text = "Enter BILL API Token",
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            ClientSize = new Size(520, 210),
            Font = _owner.Font
        };

        var instructionLabel = new Label
        {
            Text = "This copy of BSE Puller needs your BILL Spend and Expense API token before it can pull transactions.",
            AutoSize = false,
            Location = new Point(18, 18),
            Size = new Size(484, 46)
        };

        var tokenLabel = new Label
        {
            Text = "API token",
            AutoSize = true,
            Location = new Point(18, 78)
        };

        var tokenBox = new TextBox
        {
            Location = new Point(18, 100),
            Size = new Size(484, 28),
            Text = BseSettings.ApiToken
        };

        var noteLabel = new Label
        {
            Text = "The token will be saved for this Windows user and reused on future launches.",
            AutoSize = false,
            Location = new Point(18, 136),
            Size = new Size(484, 28),
            ForeColor = Color.FromArgb(96, 100, 108)
        };

        var saveButton = new Button
        {
            Text = "Save",
            DialogResult = DialogResult.OK,
            Size = new Size(90, 32),
            Location = new Point(316, 170)
        };

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Size = new Size(90, 32),
            Location = new Point(412, 170)
        };

        dialog.AcceptButton = saveButton;
        dialog.CancelButton = cancelButton;
        dialog.Controls.Add(instructionLabel);
        dialog.Controls.Add(tokenLabel);
        dialog.Controls.Add(tokenBox);
        dialog.Controls.Add(noteLabel);
        dialog.Controls.Add(saveButton);
        dialog.Controls.Add(cancelButton);

        while (true)
        {
            var result = dialog.ShowDialog(_owner);
            if (result != DialogResult.OK)
            {
                _log("Pull canceled because the BILL API token was not provided.");
                _status("API token required.");
                return false;
            }

            var token = tokenBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(token))
            {
                BseSettings.SaveApiToken(token);
                _log($"Saved the BILL API token for this user to: {BseSettings.GetUserSettingsPath()}");
                _status("API token saved.");
                return true;
            }

            MessageBox.Show(
                dialog,
                "Enter a BILL API token before continuing.",
                "API token required",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
    }

    private bool ShowUpdateAvailableDialog(string currentTag, string latestTag)
    {
        var dialog = new Form
        {
            Text = "Update Available",
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            ClientSize = new Size(520, 220),
            Font = _owner.Font
        };

        var messageLabel = new Label
        {
            AutoSize = false,
            Location = new Point(18, 18),
            Size = new Size(484, 120),
            Text = $"A newer version of BSE Puller is available.{Environment.NewLine}{Environment.NewLine}" +
                   $"Current: {currentTag}{Environment.NewLine}" +
                   $"Latest: {latestTag}{Environment.NewLine}{Environment.NewLine}" +
                   "Click Install Update to download and run the latest installer now."
        };

        var installButton = new Button
        {
            Text = "Install Update",
            DialogResult = DialogResult.OK,
            Size = new Size(130, 34),
            Location = new Point(268, 165)
        };

        var cancelButton = new Button
        {
            Text = "Not now",
            DialogResult = DialogResult.Cancel,
            Size = new Size(98, 34),
            Location = new Point(404, 165)
        };

        dialog.AcceptButton = installButton;
        dialog.CancelButton = cancelButton;
        dialog.Controls.Add(messageLabel);
        dialog.Controls.Add(installButton);
        dialog.Controls.Add(cancelButton);

        return dialog.ShowDialog(_owner) == DialogResult.OK;
    }

    private string BuildUninstallScript(string helperPath)
    {
        var installDir = BseSettings.GetInstalledAppFolder();
        var shortcutFolder = BseSettings.GetStartMenuShortcutFolder();
        var settingsFolder = BseSettings.GetUserSettingsFolder();
        var currentProcessId = Environment.ProcessId;
        var lines = new[]
        {
            "$ErrorActionPreference = 'SilentlyContinue'",
            "Add-Type -AssemblyName System.Windows.Forms",
            string.Empty,
            $"$processId = {currentProcessId}",
            $"$installDir = '{EscapePowerShellString(installDir)}'",
            $"$shortcutFolder = '{EscapePowerShellString(shortcutFolder)}'",
            $"$settingsFolder = '{EscapePowerShellString(settingsFolder)}'",
            $"$helperPath = '{EscapePowerShellString(helperPath)}'",
            string.Empty,
            "try {",
            "    Wait-Process -Id $processId",
            "} catch {",
            "}",
            string.Empty,
            "Start-Sleep -Seconds 1",
            string.Empty,
            "foreach($path in @($installDir, $shortcutFolder, $settingsFolder)) {",
            "    if (Test-Path $path) {",
            "        Remove-Item -Path $path -Recurse -Force",
            "    }",
            "}",
            string.Empty,
            "[System.Windows.Forms.MessageBox]::Show(",
            "    'BSE Puller was uninstalled for this Windows user.',",
            "    'Uninstall complete',",
            "    [System.Windows.Forms.MessageBoxButtons]::OK,",
            "    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null",
            string.Empty,
            "Start-Process -FilePath 'cmd.exe' -ArgumentList \"/c ping 127.0.0.1 -n 2 > nul & del `\"$helperPath`\"\" -WindowStyle Hidden"
        };

        return string.Join(Environment.NewLine, lines);
    }

    private void TrimOldLogs(string folder, int keepCount)
    {
        try
        {
            var files = Directory.GetFiles(folder, "*.txt")
                .Select(path => new FileInfo(path))
                .OrderByDescending(info => info.LastWriteTimeUtc)
                .ToList();

            if (files.Count <= keepCount)
            {
                return;
            }

            foreach (var file in files.Skip(keepCount))
            {
                try
                {
                    file.Delete();
                }
                catch (Exception ex)
                {
                    _log($"Warning: could not delete old log file {file.FullName}. {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            _log($"Warning: could not prune old log files. {ex.Message}");
        }
    }

    private static string EscapePowerShellString(string value)
    {
        return value.Replace("'", "''");
    }
}
