using System.Drawing.Drawing2D;
using BsePuller.Modules.Exports;
using BsePuller.Modules.Reimbursements;
using BsePuller.Modules.Settings;
using BsePuller.Modules.Transactions;

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
    private readonly ToolStripMenuItem _downloadLogMenuItem;
    private readonly ToolStripMenuItem _uninstallMenuItem;
    private readonly TextBox _logBox;
    private readonly Label _statusLabel;
    private readonly ExportsModule _exportsModule;
    private readonly SettingsModule _settingsModule;
    private readonly TransactionsModule _transactionsModule;
    private readonly ReimbursementsModule _reimbursementsModule;

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

        _checkUpdatesMenuItem = new ToolStripMenuItem("Check for Updates");

        _resetApiKeyMenuItem = new ToolStripMenuItem("Reset API Key");

        _downloadLogMenuItem = new ToolStripMenuItem("Download Log File");

        _uninstallMenuItem = new ToolStripMenuItem("Uninstall");

        _settingsMenu = new ContextMenuStrip();
        _settingsMenu.Items.AddRange(new ToolStripItem[]
        {
            _checkUpdatesMenuItem,
            _resetApiKeyMenuItem,
            _downloadLogMenuItem,
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

        _exportsModule = new ExportsModule(AppendLog);
        _settingsModule = new SettingsModule(
            this,
            AppendLog,
            message => _statusLabel.Text = message,
            Close,
            _pullButton,
            _pullReimbursementsButton,
            _openExportsButton,
            _checkUpdatesMenuItem,
            _resetApiKeyMenuItem,
            _downloadLogMenuItem,
            _uninstallMenuItem);
        _transactionsModule = new TransactionsModule(
            this,
            AppendLog,
            message => _statusLabel.Text = message,
            _settingsModule,
            _exportsModule);
        _reimbursementsModule = new ReimbursementsModule(
            this,
            AppendLog,
            message => _statusLabel.Text = message);

        _openExportsButton.Click += (_, _) => _exportsModule.OpenExportsFolder();
        _checkUpdatesMenuItem.Click += async (_, _) => await _settingsModule.CheckForUpdatesAsync(isManual: true);
        _resetApiKeyMenuItem.Click += (_, _) => _settingsModule.ResetApiKey();
        _downloadLogMenuItem.Click += (_, _) => _settingsModule.DownloadLog(_logBox.Text);
        _uninstallMenuItem.Click += (_, _) => _settingsModule.StartUninstall();

        BuildLayout();
        Shown += async (_, _) => await _settingsModule.RunStartupUpdateCheckIfNeededAsync();

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
        _downloadLogMenuItem.Enabled = false;
        _uninstallMenuItem.Enabled = false;
        var originalText = _pullButton.Text;
        _pullButton.Text = "Pulling...";
        _statusLabel.Text = "Contacting BILL and preparing the export.";

        try
        {
            await _transactionsModule.RunAsync();
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
            _checkUpdatesMenuItem.Enabled = !_settingsModule.IsUpdateCheckRunning;
            _resetApiKeyMenuItem.Enabled = true;
            _downloadLogMenuItem.Enabled = true;
            _uninstallMenuItem.Enabled = true;
        }
    }

    private void AppendLog(string message)
    {
        if (_logBox.TextLength > 0)
        {
            _logBox.AppendText(Environment.NewLine);
        }

        _logBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}");
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

            if (!ClipToRoundedRegion)
            {
                return;
            }

            var rect = ClientRectangle;
            if (rect.Width <= 0 || rect.Height <= 0)
            {
                return;
            }

            using var path = CreateRoundedRectanglePath(rect, CornerRadius);
            Region = new Region(path);
        }
    }
}
