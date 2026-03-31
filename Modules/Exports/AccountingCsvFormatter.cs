using System.Globalization;
using System.Text.RegularExpressions;

namespace BsePuller.Modules.Exports;

internal static class AccountingCsvFormatter
{
    public static readonly IReadOnlyList<string> Headers =
    [
        "* Credit Card",
        "Include",
        "* Transaction#",
        "* Description",
        "* Payee",
        "Charge Amount",
        "Credit Amount",
        "* Posted Date",
        "Notes",
        "* Account",
        "Subaccount",
        "Job",
        "Phase",
        "Job Cost Code",
        "Job Cost Type",
        "Equipment",
        "Equipment Cost Code",
        "Equipment Cost Type"
    ];

    private static readonly Regex TransactionNumberPattern = new(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}", RegexOptions.Compiled);
    private static readonly Regex JobPattern = new(@"^\s*\d+\s*-\s*.+$", RegexOptions.Compiled);

    public static IReadOnlyList<Dictionary<string, string?>> BuildRows(IReadOnlyList<Dictionary<string, string?>> sourceRows)
    {
        return BuildRows(sourceRows, "1 - Bill Spend & Expense");
    }

    public static IReadOnlyList<Dictionary<string, string?>> BuildRows(
        IReadOnlyList<Dictionary<string, string?>> sourceRows,
        string creditCardLabel)
    {
        var formattedRows = new List<Dictionary<string, string?>>(sourceRows.Count);

        foreach (var sourceRow in sourceRows)
        {
            var row = CreateEmptyRow();

            var notesSource = GetValue(sourceRow, "Notes");
            row["* Transaction#"] = NormalizeTransactionNumber(GetFirstValue(sourceRow, "updatedTime", "submittedTime", "occurredTime", "occurredDate"));
            row["* Description"] = BuildDescription(GetValue(sourceRow, "userName"), notesSource);
            row["* Payee"] = GetValue(sourceRow, "merchantName");
            row["Charge Amount"] = NormalizeAmount(GetValue(sourceRow, "amount"));
            row["Credit Amount"] = string.Empty;
            var postedDate = FormatShortDate(GetFirstValue(sourceRow, "authorizedTime", "occurredTime", "occurredDate", "submittedTime"));
            if (string.IsNullOrWhiteSpace(postedDate))
            {
                postedDate = FormatShortDate(GetValue(sourceRow, "Cleared Time in Statement"));
            }

            row["* Posted Date"] = postedDate;
            row["Notes"] = string.Empty;
            row["* Account"] = GetValue(sourceRow, "Sage General Ledger Account");
            row["Subaccount"] = GetValue(sourceRow, "Sage Vehicle and Equipment List");
            row["Job"] = NormalizeJob(GetValue(sourceRow, "budgetName"));
            row["Phase"] = string.Empty;
            row["Job Cost Code"] = GetValue(sourceRow, "2 - Sage Job Cost Codes");
            row["Job Cost Type"] = GetValue(sourceRow, "3 - Sage Cost Types");
            row["Equipment"] = string.Empty;
            row["Equipment Cost Code"] = string.Empty;
            row["Equipment Cost Type"] = string.Empty;

            if (HasAnyNonCreditCardValue(row))
            {
                row["* Credit Card"] = creditCardLabel;
                row["Include"] = "Include";
            }

            formattedRows.Add(row);
        }

        return formattedRows;
    }

    private static Dictionary<string, string?> CreateEmptyRow()
    {
        var row = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var header in Headers)
        {
            row[header] = string.Empty;
        }

        return row;
    }

    private static string? GetValue(IReadOnlyDictionary<string, string?> row, string key)
    {
        return row.TryGetValue(key, out var value) ? value?.Trim() : null;
    }

    private static string? GetFirstValue(IReadOnlyDictionary<string, string?> row, params string[] keys)
    {
        foreach (var key in keys)
        {
            var value = GetValue(row, key);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private static string NormalizeTransactionNumber(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var match = TransactionNumberPattern.Match(value);
        return match.Success ? match.Value : value.Trim();
    }

    private static string BuildDescription(string? userName, string? notes)
    {
        var cleanUserName = userName?.Trim();
        var cleanNotes = notes?.Trim();

        if (!string.IsNullOrWhiteSpace(cleanUserName) && !string.IsNullOrWhiteSpace(cleanNotes))
        {
            return $"{cleanUserName} - {cleanNotes}";
        }

        return cleanUserName ?? cleanNotes ?? string.Empty;
    }

    private static string NormalizeAmount(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var amount))
        {
            return amount.ToString("0.##", CultureInfo.InvariantCulture);
        }

        return value.Trim();
    }

    private static string FormatShortDate(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var timestamp))
        {
            return timestamp.ToString("MM/dd/yy", CultureInfo.InvariantCulture);
        }

        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dateTime))
        {
            return dateTime.ToString("MM/dd/yy", CultureInfo.InvariantCulture);
        }

        return value.Trim();
    }

    private static string NormalizeJob(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var trimmed = value.Trim();
        return JobPattern.IsMatch(trimmed) ? trimmed : string.Empty;
    }

    private static bool HasAnyNonCreditCardValue(IReadOnlyDictionary<string, string?> row)
    {
        foreach (var header in Headers)
        {
            if (string.Equals(header, "* Credit Card", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (row.TryGetValue(header, out var value) && !string.IsNullOrWhiteSpace(value))
            {
                return true;
            }
        }

        return false;
    }
}
