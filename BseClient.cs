using System.Text;
using System.Text.Json;

namespace BsePuller;

internal sealed record TransactionPullResult(
    IReadOnlyList<Dictionary<string, string?>> Rows,
    IReadOnlyList<string> ExportedTransactionIds,
    IReadOnlyList<string> SyncEligibleTransactionIds,
    IReadOnlyList<string> SyncExcludedTransactionIds);

internal sealed record TransactionSyncResult(
    int AttemptedCount,
    int SuccessfulCount,
    IReadOnlyList<string> FailureMessages);

internal sealed class BseClient : IDisposable
{
    private readonly HttpClient _httpClient;
    private const int MaxPages = 200;

    public BseClient()
    {
        _httpClient = new HttpClient
        {
            BaseAddress = BseSettings.BaseUri,
            Timeout = TimeSpan.FromMinutes(2)
        };

        _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("apiToken", BseSettings.ApiToken);
        _httpClient.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
    }

    public async Task<TransactionPullResult> GetFilteredTransactionsAsync(
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        var rows = new List<Dictionary<string, string?>>();
        var seenIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var exportedTransactionIds = new List<string>();
        var syncEligibleTransactionIds = new List<string>();
        var syncExcludedTransactionIds = new List<string>();
        var skippedAccountingIntegration = 0;
        var skippedUnapprovedAdmin = 0;
        var skippedDeclined = 0;
        var skippedDuplicates = 0;
        string? pageToken = null;
        var pageNumber = 0;
        var seenPages = new HashSet<string>(StringComparer.Ordinal);

        while (true)
        {
            if (pageNumber >= MaxPages)
            {
                throw new InvalidOperationException($"Stopped after {MaxPages} pages to avoid an endless pagination loop.");
            }

            pageNumber++;
            progress?.Report(pageNumber == 1
                ? "Requesting first page from BILL..."
                : $"Requesting page {pageNumber} from BILL...");

            var uri = BuildTransactionsUri(pageToken);
            using var response = await _httpClient.GetAsync(uri, cancellationToken);
            var body = await response.Content.ReadAsStringAsync(cancellationToken);

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(
                    $"BILL API request failed with {(int)response.StatusCode} {response.ReasonPhrase}.{Environment.NewLine}{body}");
            }

            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;

            progress?.Report($"Processing page {pageNumber}...");

            foreach (var transaction in ExtractTransactionElements(root))
            {
                if (HasAccountingIntegrationTransactions(transaction))
                {
                    skippedAccountingIntegration++;
                    continue;
                }

                if (!HasApprovedAdminReviewer(transaction))
                {
                    skippedUnapprovedAdmin++;
                    continue;
                }

                if (IsDeclined(transaction))
                {
                    skippedDeclined++;
                    continue;
                }

                var id = GetObjectString(transaction, "id");
                if (!string.IsNullOrWhiteSpace(id) && !seenIds.Add(id))
                {
                    skippedDuplicates++;
                    continue;
                }

                var row = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
                BuildTransactionRow(transaction, row);
                var mergeResult = MergeGeneralLedgerAccountColumns(row, id);

                if (!string.IsNullOrWhiteSpace(id))
                {
                    exportedTransactionIds.Add(id);

                    if (mergeResult.HasConflict)
                    {
                        syncExcludedTransactionIds.Add(id);
                    }
                    else
                    {
                        syncEligibleTransactionIds.Add(id);
                    }
                }

                if (!string.IsNullOrWhiteSpace(mergeResult.ErrorMessage))
                {
                    progress?.Report(mergeResult.ErrorMessage);
                }

                rows.Add(row);
            }

            pageToken = GetNextPageToken(root);
            if (string.IsNullOrWhiteSpace(pageToken))
            {
                progress?.Report("No more pages returned by BILL.");
                break;
            }

            if (!seenPages.Add(pageToken))
            {
                progress?.Report("BILL returned the same nextPage token more than once. Stopping pagination.");
                break;
            }
        }

        if (skippedAccountingIntegration > 0)
        {
            progress?.Report($"Skipped {skippedAccountingIntegration} transaction(s) that already had accountingIntegrationTransactions data.");
        }

        if (skippedUnapprovedAdmin > 0)
        {
            progress?.Report($"Skipped {skippedUnapprovedAdmin} transaction(s) without an ADMIN reviewer in APPROVED status.");
        }

        if (skippedDeclined > 0)
        {
            progress?.Report($"Skipped {skippedDeclined} declined transaction(s).");
        }

        if (skippedDuplicates > 0)
        {
            progress?.Report($"Skipped {skippedDuplicates} duplicate transaction(s) by id.");
        }

        return new TransactionPullResult(
            rows,
            exportedTransactionIds,
            syncEligibleTransactionIds,
            syncExcludedTransactionIds);
    }

    public async Task<TransactionSyncResult> MarkTransactionsAsSyncedAsync(
        IReadOnlyCollection<string> transactionIds,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        if (transactionIds.Count == 0)
        {
            return new TransactionSyncResult(0, 0, Array.Empty<string>());
        }

        var failureMessages = new List<string>();
        var successfulCount = 0;

        progress?.Report($"Attempting to mark {transactionIds.Count} transaction(s) as synced in BILL...");

        foreach (var transactionId in transactionIds)
        {
            try
            {
                var requestUri = new Uri(BseSettings.BaseUri, $"{BseSettings.TransactionsPath}/{Uri.EscapeDataString(transactionId)}");
                using var request = new HttpRequestMessage(HttpMethod.Put, requestUri)
                {
                    Content = new StringContent(
                        """{"syncStatus":"MANUAL_SYNCED"}""",
                        Encoding.UTF8,
                        "application/json")
                };

                using var response = await _httpClient.SendAsync(request, cancellationToken);
                var body = await response.Content.ReadAsStringAsync(cancellationToken);

                if (response.IsSuccessStatusCode)
                {
                    successfulCount++;
                    progress?.Report($"Marked transaction {transactionId} as synced.");
                    continue;
                }

                var failureMessage =
                    $"Could not mark transaction {transactionId} as synced. BILL returned {(int)response.StatusCode} {response.ReasonPhrase}.";

                if (!string.IsNullOrWhiteSpace(body))
                {
                    failureMessage += $" Response: {body}";
                }

                failureMessages.Add(failureMessage);
                progress?.Report(failureMessage);
            }
            catch (Exception ex)
            {
                var failureMessage = $"Could not mark transaction {transactionId} as synced. {ex.Message}";
                failureMessages.Add(failureMessage);
                progress?.Report(failureMessage);
            }
        }

        return new TransactionSyncResult(transactionIds.Count, successfulCount, failureMessages);
    }

    private static Uri BuildTransactionsUri(string? pageToken)
    {
        var query = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(pageToken))
        {
            AppendQuery(query, "nextPage", pageToken);
        }
        else
        {
            AppendQuery(query, "filters", BseSettings.TransactionsApiFilter);
        }

        AppendQuery(query, "max", BseSettings.PageSize.ToString(System.Globalization.CultureInfo.InvariantCulture));

        var builder = new UriBuilder(new Uri(BseSettings.BaseUri, BseSettings.TransactionsPath))
        {
            Query = query.ToString().TrimStart('?')
        };

        return builder.Uri;
    }

    private static void AppendQuery(StringBuilder builder, string name, string value)
    {
        if (builder.Length > 0)
        {
            builder.Append('&');
        }

        builder.Append(Uri.EscapeDataString(name));
        builder.Append('=');
        builder.Append(Uri.EscapeDataString(value));
    }

    private static IEnumerable<JsonElement> ExtractTransactionElements(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in root.EnumerateArray())
            {
                yield return item;
            }

            yield break;
        }

        if (root.ValueKind == JsonValueKind.Object)
        {
            foreach (var propertyName in new[] { "results", "transactions", "items", "data", "records" })
            {
                if (root.TryGetProperty(propertyName, out var array) && array.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in array.EnumerateArray())
                    {
                        yield return item;
                    }

                    yield break;
                }
            }

            foreach (var property in root.EnumerateObject())
            {
                if (property.Value.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in property.Value.EnumerateArray())
                    {
                        yield return item;
                    }

                    yield break;
                }
            }

            yield return root;
        }
    }

    private static string? GetNextPageToken(JsonElement root)
    {
        if (root.ValueKind != JsonValueKind.Object)
        {
            return null;
        }

        if (root.TryGetProperty("nextPage", out var nextPage) && nextPage.ValueKind == JsonValueKind.String)
        {
            return nextPage.GetString();
        }

        if (root.TryGetProperty("nextPageToken", out var nextPageToken) && nextPageToken.ValueKind == JsonValueKind.String)
        {
            return nextPageToken.GetString();
        }

        return null;
    }

    private static void BuildTransactionRow(JsonElement transaction, IDictionary<string, string?> row)
    {
        if (transaction.ValueKind != JsonValueKind.Object)
        {
            return;
        }

        foreach (var property in transaction.EnumerateObject())
        {
            if (string.Equals(property.Name, "tags", StringComparison.OrdinalIgnoreCase))
            {
                AddTagColumns(property.Value, row);
                continue;
            }

            if (string.Equals(property.Name, "customFields", StringComparison.OrdinalIgnoreCase))
            {
                AddCustomFieldColumns(property.Value, row);
                continue;
            }

            FlattenJson(property.Value, property.Name, row);
        }
    }

    private static bool HasAccountingIntegrationTransactions(JsonElement transaction)
    {
        if (transaction.ValueKind != JsonValueKind.Object)
        {
            return true;
        }

        if (!transaction.TryGetProperty("accountingIntegrationTransactions", out var value))
        {
            return false;
        }

        return value.ValueKind switch
        {
            JsonValueKind.Array => value.GetArrayLength() > 0,
            JsonValueKind.Object => value.EnumerateObject().Any(),
            JsonValueKind.Null or JsonValueKind.Undefined => false,
            JsonValueKind.String => !string.IsNullOrWhiteSpace(value.GetString()),
            _ => true
        };
    }

    private static bool HasApprovedAdminReviewer(JsonElement transaction)
    {
        if (transaction.ValueKind != JsonValueKind.Object ||
            !transaction.TryGetProperty("reviewers", out var reviewers) ||
            reviewers.ValueKind != JsonValueKind.Array)
        {
            return false;
        }

        foreach (var reviewer in reviewers.EnumerateArray())
        {
            if (reviewer.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var approverType = GetObjectString(reviewer, "approverType");
            var status = GetObjectString(reviewer, "status");

            if (string.Equals(approverType, "ADMIN", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(status, "APPROVED", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsDeclined(JsonElement transaction)
    {
        if (transaction.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        var transactionType = GetObjectString(transaction, "transactionType");
        return string.Equals(transactionType, "DECLINE", StringComparison.OrdinalIgnoreCase);
    }

    private static string? GetObjectString(JsonElement element, params string[] path)
    {
        var current = element;

        foreach (var segment in path)
        {
            if (current.ValueKind != JsonValueKind.Object || !current.TryGetProperty(segment, out current))
            {
                return null;
            }
        }

        return current.ValueKind switch
        {
            JsonValueKind.String => current.GetString(),
            JsonValueKind.Number => current.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => null
        };
    }

    private static void AddTagColumns(JsonElement tags, IDictionary<string, string?> row)
    {
        if (tags.ValueKind != JsonValueKind.Array)
        {
            return;
        }

        foreach (var tag in tags.EnumerateArray())
        {
            if (tag.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var tagName = GetObjectString(tag, "tagType", "name")
                ?? GetObjectString(tag, "name");

            if (string.IsNullOrWhiteSpace(tagName))
            {
                continue;
            }

            var selectedValues = new List<string>();
            if (tag.TryGetProperty("selectedTagValues", out var selectedTagValues) &&
                selectedTagValues.ValueKind == JsonValueKind.Array)
            {
                foreach (var value in selectedTagValues.EnumerateArray())
                {
                    var text = value.ValueKind switch
                    {
                        JsonValueKind.String => value.GetString(),
                        JsonValueKind.Number => value.GetRawText(),
                        JsonValueKind.True => "true",
                        JsonValueKind.False => "false",
                        JsonValueKind.Object => GetObjectString(value, "value") ?? GetObjectString(value, "name"),
                        _ => null
                    };

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        selectedValues.Add(text);
                    }
                }
            }

            row[tagName] = string.Join("; ", selectedValues);
        }
    }

    private static void AddCustomFieldColumns(JsonElement customFields, IDictionary<string, string?> row)
    {
        if (customFields.ValueKind != JsonValueKind.Array)
        {
            return;
        }

        foreach (var field in customFields.EnumerateArray())
        {
            if (field.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var fieldName = GetObjectString(field, "name");
            if (string.IsNullOrWhiteSpace(fieldName))
            {
                continue;
            }

            var values = new List<string>();

            if (field.TryGetProperty("selectedValues", out var selectedValues) &&
                selectedValues.ValueKind == JsonValueKind.Array)
            {
                foreach (var value in selectedValues.EnumerateArray())
                {
                    var text = ExtractDisplayValue(value);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        values.Add(text);
                    }
                }
            }
            else
            {
                var singleValue =
                    GetObjectString(field, "value") ??
                    GetObjectString(field, "selectedValue", "value") ??
                    GetObjectString(field, "selectedValue", "name");

                if (!string.IsNullOrWhiteSpace(singleValue))
                {
                    values.Add(singleValue);
                }
            }

            row[fieldName] = string.Join("; ", values);
        }
    }

    private static string? ExtractDisplayValue(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number => value.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Object => GetObjectString(value, "value") ?? GetObjectString(value, "name"),
            _ => null
        };
    }

    private static GeneralLedgerMergeResult MergeGeneralLedgerAccountColumns(
        IDictionary<string, string?> row,
        string? transactionId)
    {
        var populatedSources = new List<string>();
        var distinctValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        CollectGeneralLedgerValues(populatedSources, distinctValues, "Sage General Ledger Account", GetRowValue(row, "Sage General Ledger Account"));
        CollectGeneralLedgerValues(populatedSources, distinctValues, "1 - Sage Direct Expense GL Accounts", GetRowValue(row, "1 - Sage Direct Expense GL Accounts"));
        CollectGeneralLedgerValues(populatedSources, distinctValues, "Sage Vehicle and Equipment GL Accounts", GetRowValue(row, "Sage Vehicle and Equipment GL Accounts"));

        GeneralLedgerMergeResult result;
        if (distinctValues.Count == 1)
        {
            row["Sage General Ledger Account"] = distinctValues.First();
            result = GeneralLedgerMergeResult.NoConflict;
        }
        else if (distinctValues.Count > 1)
        {
            row["Sage General Ledger Account"] = string.Empty;
            var transactionLabel = string.IsNullOrWhiteSpace(transactionId) ? "(missing id)" : transactionId;
            var details = string.Join("; ", populatedSources);
            result = new GeneralLedgerMergeResult(
                true,
                $"Error: transaction {transactionLabel} has multiple values across the Sage General Ledger Account merge fields. Details: {details}. This transaction will be excluded from sync updates.");
        }
        else
        {
            row.Remove("Sage General Ledger Account");
            result = GeneralLedgerMergeResult.NoConflict;
        }

        row.Remove("1 - Sage Direct Expense GL Accounts");
        row.Remove("Sage Vehicle and Equipment GL Accounts");

        return result;
    }

    private static string? GetRowValue(IDictionary<string, string?> row, string key)
    {
        return row.TryGetValue(key, out var value) ? value : null;
    }

    private static void CollectGeneralLedgerValues(
        ICollection<string> populatedSources,
        ISet<string> distinctValues,
        string sourceName,
        string? sourceValue)
    {
        var values = SplitMultiValueCell(sourceValue).ToArray();
        if (values.Length == 0)
        {
            return;
        }

        populatedSources.Add($"{sourceName} = {string.Join(" | ", values)}");
        foreach (var value in values)
        {
            distinctValues.Add(value);
        }
    }

    private static IEnumerable<string> SplitMultiValueCell(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            yield break;
        }

        foreach (var item in value.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!string.IsNullOrWhiteSpace(item))
            {
                yield return item;
            }
        }
    }

    private static void FlattenJson(JsonElement element, string prefix, IDictionary<string, string?> values)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in element.EnumerateObject())
                {
                    var nextPrefix = string.IsNullOrEmpty(prefix)
                        ? property.Name
                        : $"{prefix}.{property.Name}";

                    FlattenJson(property.Value, nextPrefix, values);
                }
                break;

            case JsonValueKind.Array:
                values[prefix] = JoinArrayValues(element);
                break;

            case JsonValueKind.String:
                values[prefix] = element.GetString();
                break;

            case JsonValueKind.Number:
                values[prefix] = element.GetRawText();
                break;

            case JsonValueKind.True:
            case JsonValueKind.False:
                values[prefix] = element.GetBoolean().ToString();
                break;

            case JsonValueKind.Null:
            case JsonValueKind.Undefined:
                values[prefix] = string.Empty;
                break;

            default:
                values[prefix] = element.GetRawText();
                break;
        }
    }

    private static string JoinArrayValues(JsonElement element)
    {
        var values = new List<string>();

        foreach (var item in element.EnumerateArray())
        {
            var text = item.ValueKind switch
            {
                JsonValueKind.String => item.GetString(),
                JsonValueKind.Number => item.GetRawText(),
                JsonValueKind.True or JsonValueKind.False => item.GetBoolean().ToString(),
                JsonValueKind.Object when item.TryGetProperty("value", out var valueProperty) && valueProperty.ValueKind == JsonValueKind.String => valueProperty.GetString(),
                JsonValueKind.Object when item.TryGetProperty("name", out var nameProperty) && nameProperty.ValueKind == JsonValueKind.String => nameProperty.GetString(),
                _ => item.GetRawText()
            };

            if (!string.IsNullOrWhiteSpace(text))
            {
                values.Add(text);
            }
        }

        return string.Join("; ", values);
    }

    public void Dispose()
    {
        _httpClient.Dispose();
    }

    private readonly record struct GeneralLedgerMergeResult(bool HasConflict, string? ErrorMessage)
    {
        public static readonly GeneralLedgerMergeResult NoConflict = new(false, null);
    }
}
