using System.Globalization;
using System.Text;
using System.Threading;
using BsePuller.Infrastructure;
using BsePuller.Modules.Exports;
using BsePuller.Modules.Settings;

namespace BsePuller.Modules.Transactions;

internal sealed class TransactionsModule
{
    private const string TransactionsIssuerAccount = "21010 - Bill Spend & Expense";
    private const string TransactionsCreditCardLabel = "1 - Bill Spend & Expense";

    private readonly Form _owner;
    private readonly Action<string> _log;
    private readonly Action<string> _status;
    private readonly SettingsModule _settingsModule;
    private readonly ExportsModule _exportsModule;

    public TransactionsModule(
        Form owner,
        Action<string> log,
        Action<string> status,
        SettingsModule settingsModule,
        ExportsModule exportsModule)
    {
        _owner = owner;
        _log = log;
        _status = status;
        _settingsModule = settingsModule;
        _exportsModule = exportsModule;
    }

    public async Task RunAsync()
    {
        if (!_settingsModule.EnsureApiTokenConfigured())
        {
            return;
        }

        _log("Starting pull...");
        var exportsFolder = _exportsModule.EnsureExportsFolder();

        using var client = new BillApiClient();
        var service = new TransactionsService(client);
        var progress = new Progress<string>(message =>
        {
            _log(message);
            _status(message);
        });

        var result = await service.GetFilteredTransactionsAsync(progress, CancellationToken.None);
        _log($"Received {result.Rows.Count} transaction row(s).");
        _status($"Prepared {result.Rows.Count} row(s).");

        var exportRows = AccountingCsvFormatter.BuildRows(result.Rows, TransactionsCreditCardLabel);
        if (exportRows.Count == 0)
        {
            _log("No exportable transactions were returned. Clipboard was not changed.");
            _status("No transactions exported.");
            ShowNoExportableItemsDialog("transactions");
            return;
        }

        _exportsModule.TrimPreviousExportFiles(exportsFolder, "BSE-export-*.csv", "export");

        var fileName = $"BSE-export-{DateTime.Now:yyyyMMdd-HHmmss}.csv";
        var filePath = Path.Combine(exportsFolder, fileName);
        RawCsvWriter.Write(filePath, AccountingCsvFormatter.Headers, exportRows);
        _log($"Saved CSV to: {filePath}");

        var clipboardText = BuildClipboardRowsText(exportRows);
        var copiedToClipboard = TryCopyTextToClipboard(clipboardText, out var clipboardError);
        if (copiedToClipboard)
        {
            _log("Copied exported rows to the clipboard.");
            _status("Rows copied to clipboard.");
        }
        else
        {
            _log($"Warning: could not copy exported rows to the clipboard. {clipboardError}");
            _status("Could not copy rows to clipboard.");
        }

        if (result.SyncExcludedTransactionIds.Count > 0)
        {
            _log($"Excluded {result.SyncExcludedTransactionIds.Count} exported transaction(s) from sync updates because of Sage General Ledger Account merge conflicts.");
        }

        if (result.ExportedTransactionIds.Count == 0)
        {
            _log("No exported transactions were available for sync updates.");
        }

        var exportSummary = BuildExportSummary(exportRows);
        var summaryLine = BuildSummaryLine(exportSummary.TransactionCount, exportSummary.TotalAmount, "transaction(s)");
        ShowClipboardAndCompletionDialog(
            clipboardText,
            BuildCompletionMessage(exportSummary.TransactionCount, exportSummary.TotalAmount, "transaction(s)"),
            copiedToClipboard,
            TransactionsIssuerAccount,
            summaryLine);

        _log($"Export summary: {exportSummary.TransactionCount} transaction(s), total charge amount {exportSummary.TotalAmount.ToString("C2", CultureInfo.GetCultureInfo("en-US"))}.");
        _log("Reminder shown to mark the exported transactions as synced manually in BILL Spend and Expense.");
        _status("Export complete.");
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
            _owner,
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
            Font = _owner.Font
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
                _log("Copied exported rows to the clipboard again.");
                _status("Rows copied to clipboard.");
            }
            else
            {
                copiedToClipboard = false;
                _log($"Warning: could not copy exported rows to the clipboard. {errorMessage}");
                _status("Could not copy rows to clipboard.");
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
            _status("Adding bse_synced tag...");
            _log("Adding bse_synced tag to exported transactions...");

            try
            {
                var result = await tagAction();
                if (result.AttemptedCount == 0)
                {
                    _log("No transactions were available to tag.");
                    _status("No transactions to tag.");
                }
                else
                {
                    _log($"Tagging complete: {result.SuccessfulCount} succeeded, {result.FailedCount} failed.");
                    _status($"Tagging complete ({result.SuccessfulCount} ok, {result.FailedCount} failed).");
                }
            }
            catch (Exception ex)
            {
                _log($"Tagging failed. {ex.Message}");
                _status("Tagging failed.");
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
        dialog.ShowDialog(_owner);
    }
}
