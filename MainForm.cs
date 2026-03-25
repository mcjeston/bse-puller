using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Globalization;

namespace BsePuller;

internal sealed class MainForm : Form
{
    private readonly Button _pullButton;
    private readonly Button _openExportsButton;
    private readonly TextBox _logBox;
    private readonly Label _statusLabel;
    private readonly Label _subtitleLabel;

    public MainForm()
    {
        Text = "BSE Puller";
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(780, 540);
        MinimumSize = new Size(740, 520);
        Font = new Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = Color.FromArgb(244, 246, 248);

        _pullButton = new Button
        {
            Text = "Pull Transactions",
            Size = new Size(158, 38),
            BackColor = Color.FromArgb(255, 122, 0),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold, GraphicsUnit.Point),
            Cursor = Cursors.Hand
        };
        _pullButton.FlatAppearance.BorderSize = 0;
        _pullButton.Click += async (_, _) => await PullTransactionsAsync();

        _openExportsButton = new Button
        {
            Text = "Previous Exports",
            Size = new Size(148, 38),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(50, 50, 50),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
            Cursor = Cursors.Hand
        };
        _openExportsButton.FlatAppearance.BorderColor = Color.FromArgb(210, 214, 220);
        _openExportsButton.Click += (_, _) => OpenExportsFolder();

        _statusLabel = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI", 9.25F, FontStyle.Regular, GraphicsUnit.Point),
            ForeColor = Color.FromArgb(70, 74, 82),
            Text = "Waiting for you to start a pull.",
            Margin = new Padding(16, 10, 0, 0)
        };

        _subtitleLabel = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Fill,
            Text = "Pull approved not-synced transactions, format the accounting CSV, and save it in the local CSV exports folder.",
            Font = new Font("Segoe UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point),
            ForeColor = Color.FromArgb(96, 100, 108),
            Margin = new Padding(0, 8, 0, 0),
            TextAlign = ContentAlignment.TopLeft
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
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 116F));
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 94F));
        rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var headerCard = CreateCardPanel();
        headerCard.Padding = new Padding(22, 18, 22, 18);

        var accentBar = new Panel
        {
            Dock = DockStyle.Left,
            Width = 6,
            BackColor = Color.FromArgb(255, 122, 0),
            Margin = new Padding(0)
        };

        var headerTextPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.Transparent,
            Margin = new Padding(16, 0, 0, 0)
        };
        headerTextPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        headerTextPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var title = new Label
        {
            AutoSize = true,
            Text = "BILL Spend and Expense Puller",
            Font = new Font("Segoe UI Semibold", 19F, FontStyle.Bold, GraphicsUnit.Point),
            ForeColor = Color.FromArgb(28, 28, 30)
        };

        headerTextPanel.Controls.Add(title, 0, 0);
        headerTextPanel.Controls.Add(_subtitleLabel, 0, 1);
        headerCard.Controls.Add(headerTextPanel);
        headerCard.Controls.Add(accentBar);

        var actionCard = CreateCardPanel();
        actionCard.Padding = new Padding(18, 16, 18, 16);

        var actionLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            BackColor = Color.Transparent
        };
        actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        actionLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        var buttonFlow = new FlowLayoutPanel
        {
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Dock = DockStyle.Left,
            BackColor = Color.Transparent,
            Margin = new Padding(0)
        };
        buttonFlow.Controls.Add(_pullButton);
        buttonFlow.Controls.Add(_openExportsButton);

        actionLayout.Controls.Add(buttonFlow, 0, 0);
        actionLayout.Controls.Add(_statusLabel, 1, 0);
        actionCard.Controls.Add(actionLayout);

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

        var activityHint = new Label
        {
            AutoSize = true,
            Text = "Progress, warnings, and export details appear here.",
            Font = new Font("Segoe UI", 8.75F, FontStyle.Regular, GraphicsUnit.Point),
            ForeColor = Color.FromArgb(114, 118, 125),
            Location = new Point(82, 4)
        };

        activityHeaderPanel.Controls.Add(activityTitle);
        activityHeaderPanel.Controls.Add(activityHint);

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
        _openExportsButton.Enabled = false;
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
            TrimPreviousExportFiles(exportsFolder);

            using var client = new BseClient();
            var progress = new Progress<string>(message =>
            {
                AppendLog(message);
                _statusLabel.Text = message;
            });

            var result = await client.GetFilteredTransactionsAsync(progress, CancellationToken.None);
            AppendLog($"Received {result.Rows.Count} transaction row(s).");
            _statusLabel.Text = $"Prepared {result.Rows.Count} row(s).";

            var fileName = $"BSE-export-{DateTime.Now:yyyyMMdd-HHmmss}.csv";
            var filePath = Path.Combine(exportsFolder, fileName);

            var exportRows = AccountingCsvFormatter.BuildRows(result.Rows);
            RawCsvWriter.Write(filePath, AccountingCsvFormatter.Headers, exportRows);
            AppendLog($"Saved CSV to: {filePath}");

            Process.Start(new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            });
            AppendLog("Opened CSV.");

            if (result.SyncExcludedTransactionIds.Count > 0)
            {
                AppendLog($"Excluded {result.SyncExcludedTransactionIds.Count} exported transaction(s) from sync updates because of Sage General Ledger Account merge conflicts.");
            }

            if (result.ExportedTransactionIds.Count == 0)
            {
                AppendLog("No exported transactions were available for sync updates.");
                _statusLabel.Text = "Export complete.";
                return;
            }

            var exportSummary = BuildExportSummary(exportRows);
            ShowCompletionDialog(BuildCompletionMessage(exportSummary.TransactionCount, exportSummary.TotalAmount));

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
            _openExportsButton.Enabled = true;
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

    private static string BuildCompletionMessage(int transactionCount, decimal totalAmount)
    {
        var amountText = totalAmount.ToString("C2", CultureInfo.GetCultureInfo("en-US"));
        return $"The CSV export is ready.{Environment.NewLine}{Environment.NewLine}Verify {transactionCount} transaction(s) with a total charge amount of {amountText}.{Environment.NewLine}{Environment.NewLine}Then mark these transactions as synced manually in BILL Spend and Expense.";
    }

    private void ShowCompletionDialog(string message)
    {
        using var dialog = new Form
        {
            Text = "Manual Sync Reminder",
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            ClientSize = new Size(470, 210),
            Font = Font
        };

        var messageLabel = new Label
        {
            Text = message,
            AutoSize = false,
            Location = new Point(18, 18),
            Size = new Size(434, 130)
        };

        var doneButton = new Button
        {
            Text = "Done",
            DialogResult = DialogResult.OK,
            Size = new Size(92, 32),
            Location = new Point(360, 160)
        };

        dialog.AcceptButton = doneButton;
        dialog.CancelButton = doneButton;
        dialog.Controls.Add(messageLabel);
        dialog.Controls.Add(doneButton);
        dialog.ShowDialog(this);
    }

    private void TrimPreviousExportFiles(string exportsFolder)
    {
        var existingFiles = new DirectoryInfo(exportsFolder)
            .GetFiles("BSE-export-*.csv", SearchOption.TopDirectoryOnly)
            .OrderByDescending(file => file.CreationTimeUtc)
            .ThenByDescending(file => file.LastWriteTimeUtc)
            .ThenByDescending(file => file.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (existingFiles.Count == 0)
        {
            AppendLog("No previous export files were found.");
            return;
        }

        if (existingFiles.Count <= 4)
        {
            AppendLog($"Found {existingFiles.Count} previous export file(s). Keeping them as backups.");
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

        AppendLog($"Found {existingFiles.Count} previous export file(s). Kept the newest 4 backup file(s), deleted {deletedCount}, and left {failedCount} undeleted because they were unavailable.");
    }

    private void AppendLog(string message)
    {
        if (_logBox.TextLength > 0)
        {
            _logBox.AppendText(Environment.NewLine);
        }

        _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}");
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
