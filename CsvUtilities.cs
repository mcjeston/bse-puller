using System.Globalization;
using System.Text;

namespace BsePuller;

internal static class CsvUtilities
{
    public static IReadOnlyList<string[]> ParseCsv(string content)
    {
        var rows = new List<string[]>();
        var currentRow = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < content.Length; i++)
        {
            var c = content[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < content.Length && content[i + 1] == '"')
                    {
                        field.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    field.Append(c);
                }

                continue;
            }

            switch (c)
            {
                case '"':
                    inQuotes = true;
                    break;
                case ',':
                    currentRow.Add(field.ToString());
                    field.Clear();
                    break;
                case '\r':
                    if (i + 1 < content.Length && content[i + 1] == '\n')
                    {
                        i++;
                    }

                    currentRow.Add(field.ToString());
                    field.Clear();
                    rows.Add(currentRow.ToArray());
                    currentRow.Clear();
                    break;
                case '\n':
                    currentRow.Add(field.ToString());
                    field.Clear();
                    rows.Add(currentRow.ToArray());
                    currentRow.Clear();
                    break;
                default:
                    field.Append(c);
                    break;
            }
        }

        if (inQuotes)
        {
            inQuotes = false;
        }

        if (field.Length > 0 || currentRow.Count > 0)
        {
            currentRow.Add(field.ToString());
            rows.Add(currentRow.ToArray());
        }

        return rows;
    }

    public static IReadOnlyList<Dictionary<string, string?>> BuildRowDictionaries(IReadOnlyList<string[]> rows)
    {
        if (rows.Count == 0)
        {
            return Array.Empty<Dictionary<string, string?>>();
        }

        var headers = rows[0];
        var result = new List<Dictionary<string, string?>>(Math.Max(0, rows.Count - 1));

        for (var rowIndex = 1; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];
            var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

            for (var columnIndex = 0; columnIndex < headers.Length; columnIndex++)
            {
                var header = headers[columnIndex];
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                var value = columnIndex < row.Length ? row[columnIndex] : string.Empty;
                values[header] = value;
            }

            if (values.Count > 0)
            {
                result.Add(values);
            }
        }

        return result;
    }

    public static IReadOnlyList<Dictionary<string, string?>> NormalizeReimbursementRows(IReadOnlyList<Dictionary<string, string?>> rows)
    {
        var normalized = new List<Dictionary<string, string?>>(rows.Count);

        foreach (var row in rows)
        {
            var clone = new Dictionary<string, string?>(row, StringComparer.OrdinalIgnoreCase);

            ApplyAlias(clone, "amount", AmountExactKeys, AmountTokens);
            ApplyAlias(clone, "userName", UserExactKeys, UserTokens);
            ApplyAlias(clone, "merchantName", MerchantExactKeys, MerchantTokens);
            ApplyAlias(clone, "Notes", NotesExactKeys, NotesTokens);

            ApplyAlias(clone, "authorizedTime", AuthorizedExactKeys, AuthorizedTokens);
            ApplyAlias(clone, "submittedTime", SubmittedExactKeys, SubmittedTokens);
            ApplyAlias(clone, "occurredDate", OccurredExactKeys, OccurredTokens);

            if (string.IsNullOrWhiteSpace(GetValue(clone, "merchantName")))
            {
                var fallbackPayee = GetValue(clone, "userName");
                if (!string.IsNullOrWhiteSpace(fallbackPayee))
                {
                    clone["merchantName"] = fallbackPayee;
                }
            }

            normalized.Add(clone);
        }

        return normalized;
    }

    public static string BuildClipboardTextFromCsv(IReadOnlyList<string[]> rows)
    {
        if (rows.Count <= 1)
        {
            return string.Empty;
        }

        var headers = rows[0];
        var columnCount = headers.Length;
        var builder = new StringBuilder();

        for (var rowIndex = 1; rowIndex < rows.Count; rowIndex++)
        {
            var row = rows[rowIndex];

            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                if (columnIndex > 0)
                {
                    builder.Append('\t');
                }

                var value = columnIndex < row.Length ? row[columnIndex] : string.Empty;
                builder.Append(SanitizeClipboardCell(value));
            }

            if (rowIndex < rows.Count - 1)
            {
                builder.AppendLine();
            }
        }

        return builder.ToString();
    }

    public static (int RowCount, decimal TotalAmount) BuildAmountSummary(IReadOnlyList<string[]> rows)
    {
        if (rows.Count <= 1)
        {
            return (0, 0m);
        }

        var headers = rows[0];
        var amountIndex = Array.FindIndex(headers, header =>
            string.Equals(header, "amount", StringComparison.OrdinalIgnoreCase));

        var total = 0m;
        var rowCount = 0;

        for (var i = 1; i < rows.Count; i++)
        {
            var row = rows[i];
            if (row.Length == 0)
            {
                continue;
            }

            rowCount++;

            if (amountIndex < 0 || amountIndex >= row.Length)
            {
                continue;
            }

            if (decimal.TryParse(row[amountIndex], NumberStyles.Any, CultureInfo.InvariantCulture, out var amount))
            {
                total += amount;
            }
        }

        return (rowCount, total);
    }

    private static readonly string[] AmountExactKeys =
    {
        "amount",
        "reimbursement amount",
        "total amount",
        "amount (usd)",
        "amount (us dollar)",
        "amount (us$)"
    };

    private static readonly string[] AmountTokens = { "amount", "total" };

    private static readonly string[] UserExactKeys =
    {
        "employee",
        "employee name",
        "requester",
        "requestor",
        "submitted by",
        "submitted by name",
        "user",
        "user name"
    };

    private static readonly string[] UserTokens =
    {
        "employee",
        "requester",
        "requestor",
        "submitted by",
        "user name",
        "user"
    };

    private static readonly string[] MerchantExactKeys =
    {
        "merchant",
        "merchant name",
        "vendor",
        "vendor name",
        "payee",
        "store",
        "supplier",
        "business"
    };

    private static readonly string[] MerchantTokens =
    {
        "merchant",
        "vendor",
        "payee",
        "store",
        "supplier",
        "business"
    };

    private static readonly string[] NotesExactKeys =
    {
        "memo",
        "notes",
        "note",
        "description",
        "comment",
        "purpose",
        "reason"
    };

    private static readonly string[] NotesTokens =
    {
        "memo",
        "notes",
        "note",
        "description",
        "comment",
        "purpose",
        "reason"
    };

    private static readonly string[] AuthorizedExactKeys =
    {
        "paid date",
        "paid at",
        "reimbursed date",
        "reimbursed at",
        "payment date",
        "payment time",
        "approved date",
        "approved at",
        "authorized date",
        "authorized at"
    };

    private static readonly string[] AuthorizedTokens =
    {
        "paid",
        "reimbursed",
        "payment",
        "approved",
        "authorized"
    };

    private static readonly string[] SubmittedExactKeys =
    {
        "submitted date",
        "submitted at",
        "submitted time",
        "created date",
        "created at",
        "created time"
    };

    private static readonly string[] SubmittedTokens =
    {
        "submitted",
        "created"
    };

    private static readonly string[] OccurredExactKeys =
    {
        "expense date",
        "transaction date",
        "purchase date",
        "occurred date",
        "occurred time",
        "date"
    };

    private static readonly string[] OccurredTokens =
    {
        "expense date",
        "transaction date",
        "purchase date",
        "occurred"
    };

    private static void ApplyAlias(
        IDictionary<string, string?> row,
        string targetKey,
        IReadOnlyList<string> exactKeys,
        IReadOnlyList<string> tokens)
    {
        if (!string.IsNullOrWhiteSpace(GetValue(row, targetKey)))
        {
            return;
        }

        var value = GetFirstValue(row, exactKeys) ?? FindValueByTokens(row, tokens);
        if (!string.IsNullOrWhiteSpace(value))
        {
            row[targetKey] = value;
        }
    }

    private static string? GetFirstValue(IDictionary<string, string?> row, IReadOnlyList<string> keys)
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

    private static string? FindValueByTokens(IDictionary<string, string?> row, IReadOnlyList<string> tokens)
    {
        if (tokens.Count == 0)
        {
            return null;
        }

        foreach (var pair in row)
        {
            if (string.IsNullOrWhiteSpace(pair.Key) || string.IsNullOrWhiteSpace(pair.Value))
            {
                continue;
            }

            foreach (var token in tokens)
            {
                if (pair.Key.Contains(token, StringComparison.OrdinalIgnoreCase))
                {
                    return pair.Value?.Trim();
                }
            }
        }

        return null;
    }

    private static string? GetValue(IDictionary<string, string?> row, string key)
    {
        return row.TryGetValue(key, out var value) ? value?.Trim() : null;
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
}
