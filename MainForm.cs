using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text;

namespace BsePuller;

internal sealed class MainForm : Form
{
    private readonly Button _pullButton;
    private readonly Button _pullReimbursementsButton;
    private readonly Button _openExportsButton;
    private readonly Button _settingsButton;
    private readonly ContextMenuStrip _settingsMenu;
    private readonly ToolStripMenuItem _checkUpdatesMenuItem;
    private readonly ToolStripMenuItem _resetApiKeyMenuItem;
    private readonly ToolStripMenuItem _uninstallMenuItem;
    private readonly TextBox _logBox;
    private readonly Label _statusLabel;
    private Form? _reimbursementBrowserForm;
    private bool _isUpdateCheckRunning;
    private bool _startupUpdateCheckStarted;
    private const string TransactionsIssuerAccount = "21010 - Bill Spend & Expense";
    private const string ReimbursementsIssuerAccount = "21011 - BSE Reimbursements";
    private const string TransactionsCreditCardLabel = "1 - Bill Spend & Expense";
    private const string ReimbursementsCreditCardLabel = "1 - BSE Reimbursements";

    public MainForm()
    {
        Text = "BSE Puller";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(752, 555);
        MinimumSize = new Size(752, 555);
        Font = new Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = Color.FromArgb(244, 246, 248);
        ShowIcon = true;
        TrySetWindowIcon();

        _pullButton = new Button
        {
            Text = "PULL TRANSACTIONS",
            Size = new Size(158, 45),
            BackColor = Color.FromArgb(255, 122, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
            Cursor = Cursors.Hand
        };
        _pullButton.FlatAppearance.BorderSize = 0;
        _pullButton.Click += async (_, _) => await PullTransactionsAsync();

        _pullReimbursementsButton = new Button
        {
            Text = "PULL REIMBURSEMENTS",
            Size = new Size(158, 45),
            BackColor = Color.FromArgb(124, 130, 137),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
            Cursor = Cursors.Default,
            Enabled = false
        };
        _pullReimbursementsButton.FlatAppearance.BorderSize = 0;

        _openExportsButton = new Button
        {
            Text = "PREVIOUS EXPORTS",
            Size = new Size(148, 45),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(50, 50, 50),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
            Cursor = Cursors.Hand
        };
        _openExportsButton.FlatAppearance.BorderColor = Color.FromArgb(210, 214, 220);
        _openExportsButton.FlatAppearance.BorderSize = 1;
        _openExportsButton.Click += (_, _) => OpenExportsFolder();

        _checkUpdatesMenuItem = new ToolStripMenuItem("Check for Updates");
        _checkUpdatesMenuItem.Click += async (_, _) => await CheckForUpdatesAsync(isManual: true);

        _resetApiKeyMenuItem = new ToolStripMenuItem("Reset API Key");
        _resetApiKeyMenuItem.Click += (_, _) => ResetApiKey();

        _uninstallMenuItem = new ToolStripMenuItem("Uninstall");
        _uninstallMenuItem.Click += (_, _) => StartUninstall();

        _settingsMenu = new ContextMenuStrip();
        _settingsMenu.Items.AddRange(new ToolStripItem[]
        {
            _checkUpdatesMenuItem,
            _resetApiKeyMenuItem,
            _uninstallMenuItem
        });

        var settingsHoverColor = Color.FromArgb(230, 234, 240);
        var settingsDownColor = Color.FromArgb(214, 218, 224);

        _settingsButton = new Button
        {
            Text = "\uE713",
            Size = new Size(36, 36),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(90, 94, 102),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe MDL2 Assets", 12F, FontStyle.Regular, GraphicsUnit.Point),
            Cursor = Cursors.Hand,
            TabStop = false
        };
        _settingsButton.UseVisualStyleBackColor = false;
        _settingsButton.FlatAppearance.BorderSize = 0;
        _settingsButton.FlatAppearance.MouseOverBackColor = settingsHoverColor;
        _settingsButton.FlatAppearance.MouseDownBackColor = settingsDownColor;
        _settingsButton.AccessibleName = "Settings";
        _settingsButton.MouseEnter += (_, _) => _settingsButton.BackColor = settingsHoverColor;
        _settingsButton.MouseLeave += (_, _) => _settingsButton.BackColor = Color.White;
        _settingsButton.MouseDown += (_, _) => _settingsButton.BackColor = settingsDownColor;
        _settingsButton.MouseUp += (_, _) =>
        {
            var inside = _settingsButton.ClientRectangle.Contains(_settingsButton.PointToClient(Cursor.Position));
            _settingsButton.BackColor = inside ? settingsHoverColor : Color.White;
        };
        _settingsButton.Click += (_, _) =>
        {
            _settingsMenu.Show(_settingsButton, new Point(0, _settingsButton.Height));
        };

        _statusLabel = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 9.25F, FontStyle.Regular, GraphicsUnit.Point),
            ForeColor = Color.FromArgb(70, 74, 82),
            Text = "Progress, warnings, and export details appear here.",
            Margin = new Padding(12, 4, 0, 0)
        };

        _logBox = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            BorderStyle = BorderStyle.None,
            Font = new Font("Consolas", 9.25F, FontStyle.Regular, GraphicsUnit.Point),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(32, 32, 32)
        };

        BuildLayout();
        Shown += async (_, _) => await RunStartupUpdateCheckIfNeededAsync();

        AppendLog("Ready.");
        if (!BseSettings.IsConfigured)
        {
            AppendLog("Warning: the API token has not been set for this user yet.");
            _statusLabel.Text = "Install with your BILL API token, or enter it when prompted on first pull.";
        }
    }

    private void BuildLayout()
    {
        SuspendLayout();

        var rootLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20),
            ColumnCount = 1,
            RowCount = 3,
            BackColor = Color.Transparent
        };
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 82F));
        rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var headerCard = CreateCardPanel();
        headerCard.Padding = new Padding(16, 12, 16, 12);

        var accentBar = new Panel
        {
            Dock = DockStyle.Left,
            Width = 6,
            BackColor = Color.FromArgb(255, 122, 0),
            Margin = new Padding(0)
        };

        var title = new Label
        {
            AutoSize = true,
            Text = "BILL Spend and Expense Puller",
            Font = new Font("Segoe UI Semibold", 19F, FontStyle.Bold, GraphicsUnit.Point),
            ForeColor = Color.FromArgb(28, 28, 30)
        };

        var headerContent = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            BackColor = Color.Transparent,
            Margin = new Padding(0)
        };
        headerContent.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        headerContent.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 56F));
        headerContent.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        title.Margin = new Padding(16, 2, 0, 0);
        _settingsButton.Margin = new Padding(0);
        _settingsButton.Anchor = AnchorStyles.None;

        var settingsHost = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.Transparent,
            Margin = new Padding(0)
        };
        settingsHost.Controls.Add(_settingsButton);
        settingsHost.Resize += (_, _) =>
        {
            _settingsButton.Location = new Point(
                Math.Max(0, settingsHost.Width - _settingsButton.Width),
                Math.Max(0, (settingsHost.Height - _settingsButton.Height) / 2));
        };

        headerContent.Controls.Add(title, 0, 0);
        headerContent.Controls.Add(settingsHost, 1, 0);

        headerCard.Controls.Add(headerContent);
        headerCard.Controls.Add(accentBar);

        var actionCard = CreateCardPanel();
        actionCard.Padding = new Padding(18, 16, 18, 16);
        actionCard.AutoSize = true;
        actionCard.AutoSizeMode = AutoSizeMode.GrowAndShrink;

        var actionLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 3,
            BackColor = Color.Transparent
        };
        actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        actionLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));
        actionLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 18F));
        actionLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45F));

        _pullButton.Dock = DockStyle.Fill;
        _pullButton.Margin = new Padding(0, 0, 8, 0);
        _pullReimbursementsButton.Dock = DockStyle.Fill;
        _pullReimbursementsButton.Margin = new Padding(8, 0, 0, 0);
        _openExportsButton.Dock = DockStyle.Fill;
        _openExportsButton.Margin = new Padding(0);

        var reimbursementsComingSoonLabel = new Label
        {
            Text = "Coming Soon",
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.TopCenter,
            Font = new Font("Segoe UI", 8F, FontStyle.Regular, GraphicsUnit.Point),
            ForeColor = Color.FromArgb(120, 124, 130),
            Margin = new Padding(8, 2, 0, 0)
        };

        actionLayout.Controls.Add(_pullButton, 0, 0);
        actionLayout.Controls.Add(_pullReimbursementsButton, 1, 0);
        actionLayout.Controls.Add(reimbursementsComingSoonLabel, 1, 1);
        actionLayout.Controls.Add(_openExportsButton, 0, 2);
        actionLayout.SetColumnSpan(_openExportsButton, 2);
        actionCard.Controls.Add(actionLayout);

        actionLayout.AutoSize = true;
        actionLayout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        actionLayout.Dock = DockStyle.Top;

        var activityCard = CreateCardPanel();
        activityCard.Padding = new Padding(18, 16, 18, 18);

        var activityLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.Transparent
        };
        activityLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        activityLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var activityHeaderPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Height = 30,
            BackColor = Color.Transparent,
            Margin = new Padding(0, 0, 0, 12)
        };

        var activityTitle = new Label
        {
            AutoSize = true,
            Text = "Activity",
            Font = new Font("Segoe UI Semibold", 10.5F, FontStyle.Bold, GraphicsUnit.Point),
            ForeColor = Color.FromArgb(28, 28, 30),
            Location = new Point(0, 3)
        };

        var activityHeaderFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            BackColor = Color.Transparent,
            Margin = new Padding(0)
        };
        activityTitle.Margin = new Padding(0, 3, 0, 0);

        activityHeaderFlow.Controls.Add(activityTitle);
        activityHeaderFlow.Controls.Add(_statusLabel);
        activityHeaderPanel.Controls.Add(activityHeaderFlow);

        var logHost = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            Padding = new Padding(10)
        };
        logHost.Paint += (_, e) => DrawRoundedBorder(e.Graphics, logHost.ClientRectangle, 10, Color.FromArgb(226, 230, 235), Color.White);
        logHost.Controls.Add(_logBox);

        activityLayout.Controls.Add(activityHeaderPanel, 0, 0);
        activityLayout.Controls.Add(logHost, 0, 1);
        activityCard.Controls.Add(activityLayout);

        rootLayout.Controls.Add(headerCard, 0, 0);
        rootLayout.Controls.Add(actionCard, 0, 1);
        rootLayout.Controls.Add(activityCard, 0, 2);

        Controls.Add(rootLayout);
        ResumeLayout();
    }

    private async Task PullTransactionsAsync()
    {
        _pullButton.Enabled = false;
        _pullReimbursementsButton.Enabled = false;
        _openExportsButton.Enabled = false;
        _checkUpdatesMenuItem.Enabled = false;
        _resetApiKeyMenuItem.Enabled = false;
        _uninstallMenuItem.Enabled = false;
        var originalText = _pullButton.Text;
        _pullButton.Text = "Pulling...";
        _statusLabel.Text = "Contacting BILL and preparing the export.";

        try
        {
            if (!EnsureApiTokenConfigured())
            {
                return;
            }

            AppendLog("Starting pull...");
            var exportsFolder = BseSettings.GetExportsFolder();
            Directory.CreateDirectory(exportsFolder);

            using var client = new BseClient();
            var progress = new Progress<string>(message =>
            {
                AppendLog(message);
                _statusLabel.Text = message;
            });

            var result = await client.GetFilteredTransactionsAsync(progress, CancellationToken.None);
            AppendLog($"Received {result.Rows.Count} transaction row(s).");
            _statusLabel.Text = $"Prepared {result.Rows.Count} row(s).";

            var exportRows = AccountingCsvFormatter.BuildRows(result.Rows, TransactionsCreditCardLabel);
            if (exportRows.Count == 0)
            {
                AppendLog("No exportable transactions were returned. Clipboard was not changed.");
                _statusLabel.Text = "No transactions exported.";
                ShowNoExportableItemsDialog("transactions");
                return;
            }

            TrimPreviousExportFiles(exportsFolder, "BSE-export-*.csv", "export");

            var fileName = $"BSE-export-{DateTime.Now:yyyyMMdd-HHmmss}.csv";
            var filePath = Path.Combine(exportsFolder, fileName);
            RawCsvWriter.Write(filePath, AccountingCsvFormatter.Headers, exportRows);
            AppendLog($"Saved CSV to: {filePath}");

            var clipboardText = BuildClipboardRowsText(exportRows);
            var copiedToClipboard = TryCopyTextToClipboard(clipboardText, out var clipboardError);
            if (copiedToClipboard)
            {
                AppendLog("Copied exported rows to the clipboard.");
                _statusLabel.Text = "Rows copied to clipboard.";
            }
            else
            {
                AppendLog($"Warning: could not copy exported rows to the clipboard. {clipboardError}");
                _statusLabel.Text = "Could not copy rows to clipboard.";
            }

            if (result.SyncExcludedTransactionIds.Count > 0)
            {
                AppendLog($"Excluded {result.SyncExcludedTransactionIds.Count} exported transaction(s) from sync updates because of Sage General Ledger Account merge conflicts.");
            }

            if (result.ExportedTransactionIds.Count == 0)
            {
                AppendLog("No exported transactions were available for sync updates.");
            }

            var exportSummary = BuildExportSummary(exportRows);
            var summaryLine = BuildSummaryLine(exportSummary.TransactionCount, exportSummary.TotalAmount, "transaction(s)");
            ShowClipboardAndCompletionDialog(
                clipboardText,
                BuildCompletionMessage(exportSummary.TransactionCount, exportSummary.TotalAmount, "transaction(s)"),
                copiedToClipboard,
                TransactionsIssuerAccount,
                summaryLine);

            AppendLog($"Export summary: {exportSummary.TransactionCount} transaction(s), total charge amount {exportSummary.TotalAmount.ToString("C2", CultureInfo.GetCultureInfo("en-US"))}.");
            AppendLog("Reminder shown to mark the exported transactions as synced manually in BILL Spend and Expense.");
            _statusLabel.Text = "Export complete.";
        }
        catch (Exception ex)
        {
            AppendLog("Error: " + ex.Message);
            _statusLabel.Text = "Error";
            MessageBox.Show(this, ex.ToString(), "BSE Puller error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _pullButton.Text = originalText;
            _pullButton.Enabled = true;
            _pullReimbursementsButton.Enabled = false;
            _openExportsButton.Enabled = true;
            _checkUpdatesMenuItem.Enabled = !_isUpdateCheckRunning;
            _resetApiKeyMenuItem.Enabled = true;
            _uninstallMenuItem.Enabled = true;
        }
    }

    private void OpenExportsFolder()
    {
        var exportsFolder = BseSettings.GetExportsFolder();
        Directory.CreateDirectory(exportsFolder);

        Process.Start(new ProcessStartInfo
        {
            FileName = exportsFolder,
            UseShellExecute = true
        });

        AppendLog($"Opened exports folder: {exportsFolder}");
    }

    private async Task RunStartupUpdateCheckIfNeededAsync()
    {
        if (_startupUpdateCheckStarted)
        {
            return;
        }

        _startupUpdateCheckStarted = true;

        if (!BseSettings.IsRunningInstalledCopy())
        {
            AppendLog("Skipping automatic update check because this is not the installed copy.");
            return;
        }

        var lastCheckedUtc = BseSettings.GetLastUpdateCheckUtc();
        if (lastCheckedUtc is not null &&
            DateTimeOffset.UtcNow - lastCheckedUtc.Value < TimeSpan.FromHours(24))
        {
            AppendLog($"Skipping automatic update check. Last checked at {lastCheckedUtc.Value.ToLocalTime():g}.");
            return;
        }

        await CheckForUpdatesAsync(isManual: false);
    }

    private async Task CheckForUpdatesAsync(bool isManual)
    {
        if (_isUpdateCheckRunning)
        {
            if (isManual)
            {
                MessageBox.Show(
                    this,
                    "An update check is already in progress.",
                    "Update check in progress",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }

            return;
        }

        var originalStatusText = _statusLabel.Text;
        var originalCheckMenuText = _checkUpdatesMenuItem.Text;
        var pullWasEnabled = _pullButton.Enabled;
        var pullReimbursementsWasEnabled = _pullReimbursementsButton.Enabled;
        var openExportsWasEnabled = _openExportsButton.Enabled;
        var checkMenuWasEnabled = _checkUpdatesMenuItem.Enabled;
        var resetWasEnabled = _resetApiKeyMenuItem.Enabled;
        var uninstallWasEnabled = _uninstallMenuItem.Enabled;

        _isUpdateCheckRunning = true;
        _pullButton.Enabled = false;
        _pullReimbursementsButton.Enabled = false;
        _openExportsButton.Enabled = false;
        _checkUpdatesMenuItem.Enabled = false;
        _checkUpdatesMenuItem.Text = "Checking...";
        _resetApiKeyMenuItem.Enabled = false;
        _uninstallMenuItem.Enabled = false;

        try
        {
            AppendLog(isManual ? "Checking for updates..." : "Running automatic update check...");
            _statusLabel.Text = "Checking for updates...";

            var result = await UpdateService.CheckForUpdatesAsync(CancellationToken.None);
            if (!result.CheckedSuccessfully)
            {
                var error = string.IsNullOrWhiteSpace(result.ErrorMessage) ? "Unknown error." : result.ErrorMessage;
                AppendLog($"Update check failed. {error}");

                if (isManual)
                {
                    MessageBox.Show(
                        this,
                        $"Could not check for updates.{Environment.NewLine}{Environment.NewLine}{error}",
                        "Update check failed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }

                return;
            }

            if (!result.IsUpdateAvailable)
            {
                AppendLog($"No updates available. Current version: {result.CurrentTag}.");

                if (isManual)
                {
                    MessageBox.Show(
                        this,
                        $"You already have the latest version ({result.CurrentTag}).",
                        "No updates available",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }

                return;
            }

            var latestTag = result.LatestTag ?? "unknown";
            AppendLog($"Update available. Current: {result.CurrentTag}. Latest: {latestTag}.");

            if (!ShowUpdateAvailableDialog(result.CurrentTag, latestTag))
            {
                AppendLog("Update postponed.");
                return;
            }

            if (string.IsNullOrWhiteSpace(result.DownloadUrl))
            {
                const string missingAssetMessage = "A newer release was found, but it does not include BsePullerSetup.exe.";
                AppendLog(missingAssetMessage);
                MessageBox.Show(
                    this,
                    missingAssetMessage,
                    "Update unavailable",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            AppendLog("Downloading update installer...");
            _statusLabel.Text = "Downloading update installer...";

            var downloadResult = await UpdateService.DownloadInstallerAsync(result.DownloadUrl, CancellationToken.None);
            if (!downloadResult.Success || string.IsNullOrWhiteSpace(downloadResult.FilePath))
            {
                var error = string.IsNullOrWhiteSpace(downloadResult.ErrorMessage) ? "Unknown error." : downloadResult.ErrorMessage;
                AppendLog($"Update download failed. {error}");
                MessageBox.Show(
                    this,
                    $"Could not download the update installer.{Environment.NewLine}{Environment.NewLine}{error}",
                    "Update download failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            AppendLog($"Downloaded update installer to: {downloadResult.FilePath}");
            Process.Start(new ProcessStartInfo
            {
                FileName = downloadResult.FilePath,
                UseShellExecute = true
            });

            AppendLog("Launched update installer. Closing BSE Puller.");
            _statusLabel.Text = "Update installer started. Closing...";
            Close();
        }
        catch (Exception ex)
        {
            AppendLog($"Update check failed. {ex.Message}");

            if (isManual)
            {
                MessageBox.Show(
                    this,
                    $"Could not complete the update check.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                    "Update check failed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }
        finally
        {
            try
            {
                BseSettings.SaveLastUpdateCheckUtc(DateTimeOffset.UtcNow);
            }
            catch (Exception ex)
            {
                AppendLog($"Warning: could not save update-check timestamp. {ex.Message}");
            }

            _isUpdateCheckRunning = false;

            if (!IsDisposed)
            {
                _pullButton.Enabled = pullWasEnabled;
                _pullReimbursementsButton.Enabled = pullReimbursementsWasEnabled;
                _openExportsButton.Enabled = openExportsWasEnabled;
                _checkUpdatesMenuItem.Enabled = checkMenuWasEnabled;
                _checkUpdatesMenuItem.Text = originalCheckMenuText;
                _resetApiKeyMenuItem.Enabled = resetWasEnabled;
                _uninstallMenuItem.Enabled = uninstallWasEnabled;
                _statusLabel.Text = originalStatusText;
            }
        }
    }

    private bool ShowUpdateAvailableDialog(string currentTag, string latestTag)
    {
        using var dialog = new Form
        {
            Text = "Update Available",
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            ClientSize = new Size(560, 230),
            Font = Font
        };

        var messageLabel = new Label
        {
            AutoSize = false,
            Location = new Point(18, 18),
            Size = new Size(524, 138),
            Text =
                "A newer version of BSE Puller is available." +
                Environment.NewLine +
                Environment.NewLine +
                $"Current version: {currentTag}" +
                Environment.NewLine +
                $"Latest version: {latestTag}" +
                Environment.NewLine +
                Environment.NewLine +
                "Click Install Update to download and run the latest installer now."
        };

        var installButton = new Button
        {
            Text = "Install Update",
            DialogResult = DialogResult.OK,
            Size = new Size(118, 32),
            Location = new Point(332, 180)
        };

        var laterButton = new Button
        {
            Text = "Later",
            DialogResult = DialogResult.Cancel,
            Size = new Size(92, 32),
            Location = new Point(450, 180)
        };

        dialog.AcceptButton = laterButton;
        dialog.CancelButton = laterButton;
        dialog.Controls.Add(messageLabel);
        dialog.Controls.Add(installButton);
        dialog.Controls.Add(laterButton);

        return dialog.ShowDialog(this) == DialogResult.OK;
    }

    private void ResetApiKey()
    {
        var confirm = MessageBox.Show(
            this,
            "This will remove the saved BILL API token for this Windows user. The next pull will prompt for a new token.",
            "Reset API key?",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question,
            MessageBoxDefaultButton.Button2);

        if (confirm != DialogResult.Yes)
        {
            return;
        }

        BseSettings.ClearApiToken();
        AppendLog("Removed the saved BILL API token for this Windows user.");
        _statusLabel.Text = "API token removed. The next pull will prompt for a new token.";

        MessageBox.Show(
            this,
            "The saved BILL API token was removed. The next pull will ask for a new token.",
            "API token reset",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void StartUninstall()
    {
        if (!BseSettings.IsRunningInstalledCopy())
        {
            MessageBox.Show(
                this,
                $"Uninstall is only available from the installed copy in:{Environment.NewLine}{BseSettings.GetInstalledAppFolder()}",
                "Uninstall unavailable",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return;
        }

        var confirm = MessageBox.Show(
            this,
            "This will close BSE Puller and remove the installed app, the Start menu shortcut, the saved API token for this Windows user, and the local CSV exports folder. Continue?",
            "Uninstall BSE Puller?",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);

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

            AppendLog("Started the uninstall helper. Closing BSE Puller.");
            _statusLabel.Text = "Closing for uninstall...";
            Close();
        }
        catch (Exception ex)
        {
            AppendLog($"Error: could not start uninstall helper. {ex.Message}");
            MessageBox.Show(
                this,
                $"Could not start the uninstall helper.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                "Uninstall failed",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private bool EnsureApiTokenConfigured()
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
            Font = Font
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
            var result = dialog.ShowDialog(this);
            if (result != DialogResult.OK)
            {
                AppendLog("Pull canceled because the BILL API token was not provided.");
                _statusLabel.Text = "API token required.";
                return false;
            }

            var token = tokenBox.Text.Trim();
            if (!string.IsNullOrWhiteSpace(token))
            {
                BseSettings.SaveApiToken(token);
                AppendLog($"Saved the BILL API token for this user to: {BseSettings.GetUserSettingsPath()}");
                _statusLabel.Text = "API token saved.";
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

    private static (int TransactionCount, decimal TotalAmount) BuildExportSummary(IReadOnlyList<Dictionary<string, string?>> exportRows)
    {
        decimal totalAmount = 0;

        foreach (var row in exportRows)
        {
            if (row.TryGetValue("Charge Amount", out var amountText) &&
                decimal.TryParse(amountText, NumberStyles.Any, CultureInfo.InvariantCulture, out var amount))
            {
                totalAmount += amount;
            }
        }

        return (exportRows.Count, totalAmount);
    }

    private static (int RowCount, decimal TotalAmount) BuildRawAmountSummary(IReadOnlyList<Dictionary<string, string?>> rows)
    {
        decimal totalAmount = 0;

        foreach (var row in rows)
        {
            var amountText = GetFirstRawAmountValue(row);
            if (!string.IsNullOrWhiteSpace(amountText) &&
                decimal.TryParse(amountText, NumberStyles.Any, CultureInfo.InvariantCulture, out var amount))
            {
                totalAmount += amount;
            }
        }

        return (rows.Count, totalAmount);
    }

    private static string? GetFirstRawAmountValue(IReadOnlyDictionary<string, string?> row)
    {
        if (row.TryGetValue("amount", out var direct) && !string.IsNullOrWhiteSpace(direct))
        {
            return direct;
        }

        if (row.TryGetValue("amount.value", out var nested) && !string.IsNullOrWhiteSpace(nested))
        {
            return nested;
        }

        return null;
    }

    private static string BuildCompletionMessage(int itemCount, decimal totalAmount, string itemLabel)
    {
        var summaryLine = BuildSummaryLine(itemCount, totalAmount, itemLabel);
        return $"The CSV export is ready.{Environment.NewLine}{Environment.NewLine}{summaryLine}{Environment.NewLine}{Environment.NewLine}Then mark these {itemLabel} as synced manually in BILL Spend and Expense.";
    }

    private static string BuildSummaryLine(int itemCount, decimal totalAmount, string itemLabel)
    {
        var amountText = totalAmount.ToString("C2", CultureInfo.GetCultureInfo("en-US"));
        return $"Verify {itemCount} {itemLabel} with a total charge amount of {amountText}.";
    }

    private static string BuildClipboardRowsText(IReadOnlyList<Dictionary<string, string?>> exportRows)
    {
        var builder = new StringBuilder();

        for (var rowIndex = 0; rowIndex < exportRows.Count; rowIndex++)
        {
            var row = exportRows[rowIndex];

            for (var columnIndex = 0; columnIndex < AccountingCsvFormatter.Headers.Count; columnIndex++)
            {
                if (columnIndex > 0)
                {
                    builder.Append('\t');
                }

                var header = AccountingCsvFormatter.Headers[columnIndex];
                row.TryGetValue(header, out var value);
                builder.Append(SanitizeClipboardCell(value));
            }

            if (rowIndex < exportRows.Count - 1)
            {
                builder.AppendLine();
            }
        }

        return builder.ToString();
    }

    private static string BuildClipboardRowsText(
        IReadOnlyList<Dictionary<string, string?>> exportRows,
        IReadOnlyList<string> headers)
    {
        var builder = new StringBuilder();

        for (var rowIndex = 0; rowIndex < exportRows.Count; rowIndex++)
        {
            var row = exportRows[rowIndex];

            for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
            {
                if (columnIndex > 0)
                {
                    builder.Append('\t');
                }

                var header = headers[columnIndex];
                row.TryGetValue(header, out var value);
                builder.Append(SanitizeClipboardCell(value));
            }

            if (rowIndex < exportRows.Count - 1)
            {
                builder.AppendLine();
            }
        }

        return builder.ToString();
    }

    private static string SanitizeClipboardCell(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        return value
            .Replace("\r\n", " ", StringComparison.Ordinal)
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Replace('\t', ' ');
    }

    private static bool TryCopyTextToClipboard(string text, out string errorMessage)
    {
        for (var attempt = 1; attempt <= 5; attempt++)
        {
            try
            {
                Clipboard.SetText(text);
                errorMessage = string.Empty;
                return true;
            }
            catch (Exception) when (attempt < 5)
            {
                Thread.Sleep(80);
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        errorMessage = "Unknown clipboard error.";
        return false;
    }

    private static string BuildClipboardInstructionMessage(string issuerAccountLabel, string summaryLine)
    {
        return "The information has been copied to the clipboard." +
               Environment.NewLine +
               Environment.NewLine +
               summaryLine +
               Environment.NewLine +
               Environment.NewLine +
               "1. Open Sage 100 Contractor screen 4-7-7 (Import Credit Card Transactions)" +
               Environment.NewLine +
               $"2. Select card issuer account {issuerAccountLabel}" +
               Environment.NewLine +
               "3. Paste into the first cell";
    }

    private static string BuildClipboardUnavailableMessage(string issuerAccountLabel, string summaryLine)
    {
        return "The CSV export was saved, but the information could not be copied to the clipboard automatically." +
               Environment.NewLine +
               Environment.NewLine +
               summaryLine +
               Environment.NewLine +
               Environment.NewLine +
               "1. Open Sage 100 Contractor screen 4-7-7 (Import Credit Card Transactions)" +
               Environment.NewLine +
               $"2. Select card issuer account {issuerAccountLabel}" +
               Environment.NewLine +
               "3. Paste into the first cell";
    }

    private void ShowNoExportableItemsDialog(string itemLabel)
    {
        MessageBox.Show(
            this,
            $"No exportable {itemLabel} were returned, so nothing was copied to the clipboard.",
            $"No {itemLabel} to export",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void ShowClipboardAndCompletionDialog(
        string clipboardText,
        string completionMessage,
        bool copiedToClipboard,
        string issuerAccountLabel,
        string summaryLine,
        Func<Task<TagOperationResult>>? tagAction = null)
    {
        using var dialog = new Form
        {
            Text = "Sage Import Instructions",
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            ClientSize = new Size(700, 270),
            Font = Font
        };

        var messageLabel = new Label
        {
            AutoSize = false,
            Location = new Point(18, 18),
            Size = new Size(664, 176)
        };
        messageLabel.UseMnemonic = false;

        var copyAgainButton = new Button
        {
            Text = "Copy Again",
            Size = new Size(102, 32),
            Location = new Point(480, 220)
        };

        var tagButton = new Button
        {
            Text = "Add Synced Tag",
            Size = new Size(130, 32),
            Location = new Point(340, 220),
            Visible = false,
            Enabled = false
        };

        var backButton = new Button
        {
            Text = "Back",
            Size = new Size(102, 32),
            Location = new Point(480, 220),
            Visible = false,
            Enabled = false
        };

        var doneButton = new Button
        {
            Text = "Done",
            Size = new Size(92, 32),
            Location = new Point(590, 220)
        };

        var showingSummary = false;

        void ShowInstructionState()
        {
            showingSummary = false;
            dialog.Text = "Sage Import Instructions";
            messageLabel.Text = copiedToClipboard
                ? BuildClipboardInstructionMessage(issuerAccountLabel, summaryLine)
                : BuildClipboardUnavailableMessage(issuerAccountLabel, summaryLine);
            copyAgainButton.Visible = true;
            copyAgainButton.Enabled = true;
            tagButton.Visible = false;
            tagButton.Enabled = false;
            backButton.Visible = false;
            backButton.Enabled = false;
        }

        void ShowSummaryState()
        {
            showingSummary = true;
            dialog.Text = "Manual Sync Reminder";
            messageLabel.Text = completionMessage;
            copyAgainButton.Visible = false;
            copyAgainButton.Enabled = false;
            tagButton.Visible = tagAction is not null;
            tagButton.Enabled = tagAction is not null;
            backButton.Visible = true;
            backButton.Enabled = true;
        }

        copyAgainButton.Click += (_, _) =>
        {
            if (TryCopyTextToClipboard(clipboardText, out var errorMessage))
            {
                copiedToClipboard = true;
                AppendLog("Copied exported rows to the clipboard again.");
                _statusLabel.Text = "Rows copied to clipboard.";
            }
            else
            {
                copiedToClipboard = false;
                AppendLog($"Warning: could not copy exported rows to the clipboard. {errorMessage}");
                _statusLabel.Text = "Could not copy rows to clipboard.";
                MessageBox.Show(
                    dialog,
                    $"Could not copy the information to the clipboard.{Environment.NewLine}{Environment.NewLine}{errorMessage}",
                    "Clipboard unavailable",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }

            ShowInstructionState();
        };

        backButton.Click += (_, _) => ShowInstructionState();

        tagButton.Click += async (_, _) =>
        {
            if (tagAction is null)
            {
                return;
            }

            tagButton.Enabled = false;
            backButton.Enabled = false;
            doneButton.Enabled = false;
            _statusLabel.Text = "Adding bse_synced tag...";
            AppendLog("Adding bse_synced tag to exported transactions...");

            try
            {
                var result = await tagAction();
                if (result.AttemptedCount == 0)
                {
                    AppendLog("No transactions were available to tag.");
                    _statusLabel.Text = "No transactions to tag.";
                }
                else
                {
                    AppendLog($"Tagging complete: {result.SuccessfulCount} succeeded, {result.FailedCount} failed.");
                    _statusLabel.Text = $"Tagging complete ({result.SuccessfulCount} ok, {result.FailedCount} failed).";
                }
            }
            catch (Exception ex)
            {
                AppendLog($"Tagging failed. {ex.Message}");
                _statusLabel.Text = "Tagging failed.";
            }
            finally
            {
                tagButton.Enabled = true;
                backButton.Enabled = true;
                doneButton.Enabled = true;
            }
        };

        doneButton.Click += (_, _) =>
        {
            if (!showingSummary)
            {
                ShowSummaryState();
                return;
            }

            dialog.DialogResult = DialogResult.OK;
            dialog.Close();
        };

        dialog.AcceptButton = doneButton;
        dialog.Controls.Add(messageLabel);
        dialog.Controls.Add(copyAgainButton);
        dialog.Controls.Add(tagButton);
        dialog.Controls.Add(backButton);
        dialog.Controls.Add(doneButton);

        ShowInstructionState();
        dialog.ShowDialog(this);
    }

    private async Task EnsureReimbursementsBrowserOpenAsync()
    {
        if (_reimbursementBrowserForm is not null && !_reimbursementBrowserForm.IsDisposed)
        {
            return;
        }

        AppendLog("Reopening reimbursements page for manual sync.");
        _statusLabel.Text = "Reopening reimbursements page for manual sync...";
        _reimbursementBrowserForm = await ReimbursementWebExporter.OpenReimbursementsWindowAsync(this, AppendLog, message => _statusLabel.Text = message);
        if (_reimbursementBrowserForm is not null)
        {
            _reimbursementBrowserForm.FormClosed += (_, _) => _reimbursementBrowserForm = null;
        }
    }

    private void TrimPreviousExportFiles(string exportsFolder, string searchPattern, string label)
    {
        var existingFiles = new DirectoryInfo(exportsFolder)
            .GetFiles(searchPattern, SearchOption.TopDirectoryOnly)
            .OrderByDescending(file => file.CreationTimeUtc)
            .ThenByDescending(file => file.LastWriteTimeUtc)
            .ThenByDescending(file => file.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (existingFiles.Count == 0)
        {
            AppendLog($"No previous {label} file(s) were found.");
            return;
        }

        if (existingFiles.Count <= 4)
        {
            AppendLog($"Found {existingFiles.Count} previous {label} file(s). Keeping them as backups.");
            return;
        }

        var filesToDelete = existingFiles.Skip(4).ToList();
        var deletedCount = 0;
        var failedCount = 0;

        foreach (var file in filesToDelete)
        {
            try
            {
                file.Delete();
                deletedCount++;
            }
            catch (Exception ex)
            {
                failedCount++;
                AppendLog($"Warning: could not delete old export {file.Name}. {ex.Message}");
            }
        }

        AppendLog($"Found {existingFiles.Count} previous {label} file(s). Kept the newest 4 backup file(s), deleted {deletedCount}, and left {failedCount} undeleted because they were unavailable.");
    }

    private void AppendLog(string message)
    {
        if (_logBox.TextLength > 0)
        {
            _logBox.AppendText(Environment.NewLine);
        }

        _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}");
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

    private static string EscapePowerShellString(string value)
    {
        return value.Replace("'", "''");
    }

    private void TrySetWindowIcon()
    {
        try
        {
            Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        }
        catch
        {
            // Keep the default icon if the executable icon cannot be read.
        }
    }

    private static RoundedPanel CreateCardPanel()
    {
        return new RoundedPanel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            BorderColor = Color.FromArgb(228, 232, 236),
            CornerRadius = 14,
            ClipToRoundedRegion = true,
            Margin = new Padding(0, 0, 0, 14)
        };
    }

    private static void DrawRoundedBorder(Graphics graphics, Rectangle bounds, int radius, Color borderColor, Color fillColor)
    {
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        var rect = Rectangle.Inflate(bounds, -1, -1);
        using var path = CreateRoundedRectanglePath(rect, radius);
        using var brush = new SolidBrush(fillColor);
        using var pen = new Pen(borderColor);
        graphics.FillPath(brush, path);
        graphics.DrawPath(pen, path);
    }

    private static GraphicsPath CreateRoundedRectanglePath(Rectangle bounds, int radius)
    {
        var diameter = radius * 2;
        var path = new GraphicsPath();

        path.AddArc(bounds.X, bounds.Y, diameter, diameter, 180, 90);
        path.AddArc(bounds.Right - diameter, bounds.Y, diameter, diameter, 270, 90);
        path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
        path.AddArc(bounds.X, bounds.Bottom - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();

        return path;
    }

    private sealed class RoundedPanel : Panel
    {
        public int CornerRadius { get; set; } = 16;
        public Color BorderColor { get; set; } = Color.FromArgb(226, 230, 235);
        public bool ClipToRoundedRegion { get; set; }

        public RoundedPanel()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);

            if (!ClipToRoundedRegion || Width <= 0 || Height <= 0)
            {
                Region = null;
                return;
            }

            using var path = CreateRoundedRectanglePath(new Rectangle(0, 0, Width, Height), CornerRadius);
            Region?.Dispose();
            Region = new Region(path);
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            base.OnPaintBackground(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var path = CreateRoundedRectanglePath(new Rectangle(0, 0, Width - 1, Height - 1), CornerRadius);
            using var pen = new Pen(BorderColor);
            e.Graphics.DrawPath(pen, path);
        }
    }
}
