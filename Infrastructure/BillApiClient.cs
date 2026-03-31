using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;

using BsePuller.Modules.Settings;

namespace BsePuller.Infrastructure;

internal sealed record TransactionPullResult(
    IReadOnlyList<Dictionary<string, string?>> Rows,
    IReadOnlyList<string> ExportedTransactionIds,
    IReadOnlyList<string> SyncEligibleTransactionIds,
    IReadOnlyList<string> SyncExcludedTransactionIds);

internal sealed record ReimbursementPullResult(
    IReadOnlyList<Dictionary<string, string?>> Rows,
    IReadOnlyList<string> ExportedReimbursementIds,
    IReadOnlyList<string> DuplicateReimbursementIds,
    int MissingIdCount);

internal sealed record TransactionSyncResult(
    int AttemptedCount,
    int SuccessfulCount,
    IReadOnlyList<string> FailureMessages);

internal sealed record TagOperationResult(
    int AttemptedCount,
    int SuccessfulCount,
    int FailedCount);

internal sealed class BillApiClient : IDisposable
{
    private readonly HttpClient _httpClient;
    private const int MaxPages = 200;
    private const string SyncedTagName = "bse_synced";

    public BillApiClient()
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
            var body = await GetResponseBodyWithRetriesAsync(uri, progress, cancellationToken);

            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;

            progress?.Report($"Processing page {pageNumber}...");

            foreach (var transaction in ExtractListElements(root))
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
                BuildRowFromPayload(transaction, row);

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

    public async Task<ReimbursementPullResult> GetFilteredReimbursementsAsync(
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        var rows = new List<Dictionary<string, string?>>();
        var seenIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var exportedReimbursementIds = new List<string>();
        var duplicateIds = new List<string>();
        var missingIdCount = 0;
        var skippedStatusOther = 0;
        var includedPaid = 0;
        var includedPaymentPending = 0;
        var exportedCount = 0;
        var skippedRetired = 0;
        var skippedSynced = 0;
        var skippedSyncUnknown = 0;
        var skippedTagged = 0;
        var scannedItems = 0;
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
                ? "Requesting first reimbursement page from BILL..."
                : $"Requesting reimbursement page {pageNumber} from BILL...");

            var uri = BuildReimbursementsUri(pageToken);
            var body = await GetResponseBodyWithRetriesAsync(uri, progress, cancellationToken);

            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;

            var pageItems = ExtractListElements(root).ToList();
            scannedItems += pageItems.Count;
            progress?.Report($"Processing reimbursement page {pageNumber} ({pageItems.Count} item(s)).");

            if (pageItems.Count == 0)
            {
                progress?.Report("BILL returned an empty reimbursement page. Stopping pagination.");
                break;
            }

            foreach (var reimbursement in pageItems)
            {
                if (IsReimbursementRetired(reimbursement))
                {
                    skippedRetired++;
                    continue;
                }

                var status = GetObjectString(reimbursement, "status");
                var isPaid = string.Equals(status, "PAID", StringComparison.OrdinalIgnoreCase);
                var isPaymentPending = IsPaymentPendingStatus(status);
                if (!isPaid && !isPaymentPending)
                {
                    skippedStatusOther++;
                    continue;
                }

                if (isPaid)
                {
                    includedPaid++;
                }
                else
                {
                    includedPaymentPending++;
                }

                if (!IsNotSyncedReimbursement(reimbursement, out var syncStatus))
                {
                    skippedSynced++;
                    continue;
                }

                if (string.IsNullOrWhiteSpace(syncStatus))
                {
                    skippedSyncUnknown++;
                }

                var id = GetReimbursementId(reimbursement);
                if (string.IsNullOrWhiteSpace(id))
                {
                    missingIdCount++;
                }
                else if (!seenIds.Add(id))
                {
                    duplicateIds.Add(id);
                    continue;
                }

                var row = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
                BuildRowFromPayload(reimbursement, row);

                if (HasSyncedTag(row))
                {
                    skippedTagged++;
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(id))
                {
                    exportedReimbursementIds.Add(id);
                }

                rows.Add(row);
                exportedCount++;
            }

            pageToken = GetNextPageToken(root);
            if (string.IsNullOrWhiteSpace(pageToken))
            {
                progress?.Report("No more reimbursement pages returned by BILL.");
                break;
            }

            if (!seenPages.Add(pageToken))
            {
                progress?.Report("BILL returned the same nextPage token more than once. Stopping pagination.");
                break;
            }
        }

        progress?.Report($"Scanned {scannedItems} reimbursement item(s) across {pageNumber} page(s).");
        progress?.Report($"Exported {exportedCount} reimbursement item(s) after filters.");

        if (includedPaid > 0)
        {
            progress?.Report($"Included {includedPaid} paid reimbursement(s).");
        }

        if (includedPaymentPending > 0)
        {
            progress?.Report($"Included {includedPaymentPending} payment pending reimbursement(s).");
        }

        if (skippedStatusOther > 0)
        {
            progress?.Report($"Skipped {skippedStatusOther} reimbursement(s) that were not PAID or PAYMENT_PENDING.");
        }

        if (skippedRetired > 0)
        {
            progress?.Report($"Skipped {skippedRetired} retired reimbursement(s).");
        }

        if (skippedSynced > 0)
        {
            progress?.Report($"Skipped {skippedSynced} reimbursement(s) already synced in BILL.");
        }

        if (skippedSyncUnknown > 0)
        {
            progress?.Report($"Note: {skippedSyncUnknown} reimbursement(s) did not report a sync status.");
        }

        if (skippedTagged > 0)
        {
            progress?.Report($"Skipped {skippedTagged} reimbursement(s) already tagged with {SyncedTagName}.");
        }

        if (duplicateIds.Count > 0)
        {
            progress?.Report($"Skipped {duplicateIds.Count} duplicate reimbursement(s) by id.");
        }

        return new ReimbursementPullResult(rows, exportedReimbursementIds, duplicateIds, missingIdCount);
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

    public async Task<TagOperationResult> TagReimbursementsWithSyncedTagAsync(
        IReadOnlyCollection<string> reimbursementIds,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        if (reimbursementIds.Count == 0)
        {
            return new TagOperationResult(0, 0, 0);
        }

        progress?.Report($"Preparing {SyncedTagName} tag for {reimbursementIds.Count} reimbursement(s)...");

        var (customFieldId, customFieldValueId) = await EnsureCustomFieldWithValueAsync(
            SyncedTagName,
            SyncedTagName,
            progress,
            cancellationToken);

        var successCount = 0;
        var failedCount = 0;

        foreach (var reimbursementId in reimbursementIds)
        {
            try
            {
                var requestUri = new Uri(BseSettings.BaseUri,
                    $"{BseSettings.ReimbursementsPath}/{Uri.EscapeDataString(reimbursementId)}/custom-fields");
                var payload = JsonSerializer.Serialize(new
                {
                    customFields = new[]
                    {
                        new
                        {
                            customFieldUuid = customFieldId,
                            selectedValues = new[] { customFieldValueId }
                        }
                    }
                });

                await SendJsonWithRetriesAsync(HttpMethod.Put, requestUri, payload, progress, cancellationToken);
                successCount++;
                progress?.Report($"Added {SyncedTagName} tag to reimbursement {reimbursementId}.");
            }
            catch (Exception ex)
            {
                failedCount++;
                progress?.Report($"Could not tag reimbursement {reimbursementId}. {ex.Message}");
            }
        }

        return new TagOperationResult(reimbursementIds.Count, successCount, failedCount);
    }

    private async Task<(string CustomFieldId, string CustomFieldValueId)> EnsureCustomFieldWithValueAsync(
        string fieldName,
        string valueName,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        var customFieldId = await FindCustomFieldIdByNameAsync(fieldName, progress, cancellationToken);
        if (string.IsNullOrWhiteSpace(customFieldId))
        {
            progress?.Report($"Creating custom field {fieldName} in BILL...");
            customFieldId = await CreateCustomFieldAsync(fieldName, valueName, progress, cancellationToken);
        }

        if (string.IsNullOrWhiteSpace(customFieldId))
        {
            throw new InvalidOperationException($"Could not resolve custom field {fieldName} in BILL.");
        }

        var valueId = await FindCustomFieldValueIdByNameAsync(customFieldId, valueName, progress, cancellationToken);
        if (string.IsNullOrWhiteSpace(valueId))
        {
            progress?.Report($"Creating custom field value {valueName}...");
            await CreateCustomFieldValueAsync(customFieldId, valueName, progress, cancellationToken);
            valueId = await FindCustomFieldValueIdByNameAsync(customFieldId, valueName, progress, cancellationToken);
        }

        if (string.IsNullOrWhiteSpace(valueId))
        {
            throw new InvalidOperationException($"Could not resolve custom field value {valueName} in BILL.");
        }

        return (customFieldId, valueId);
    }

    private async Task<string?> FindCustomFieldIdByNameAsync(
        string fieldName,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        string? pageToken = null;
        var seenPages = new HashSet<string>(StringComparer.Ordinal);

        while (true)
        {
            var uri = BuildCustomFieldsUri(pageToken);
            var body = await GetResponseBodyWithRetriesAsync(uri, progress, cancellationToken);
            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;

            foreach (var field in ExtractListElements(root))
            {
                var name = GetObjectString(field, "name");
                if (string.Equals(name, fieldName, StringComparison.OrdinalIgnoreCase))
                {
                    return GetObjectString(field, "uuid") ?? GetObjectString(field, "id");
                }
            }

            pageToken = GetNextPageToken(root);
            if (string.IsNullOrWhiteSpace(pageToken) || !seenPages.Add(pageToken))
            {
                return null;
            }
        }
    }

    private async Task<string?> FindCustomFieldValueIdByNameAsync(
        string customFieldId,
        string valueName,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        string? pageToken = null;
        var seenPages = new HashSet<string>(StringComparer.Ordinal);

        while (true)
        {
            var uri = BuildCustomFieldValuesUri(customFieldId, pageToken);
            var body = await GetResponseBodyWithRetriesAsync(uri, progress, cancellationToken);
            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;

            foreach (var value in ExtractListElements(root))
            {
                var name = GetObjectString(value, "value") ?? GetObjectString(value, "name");
                if (string.Equals(name, valueName, StringComparison.OrdinalIgnoreCase))
                {
                    return GetObjectString(value, "uuid") ?? GetObjectString(value, "id");
                }
            }

            pageToken = GetNextPageToken(root);
            if (string.IsNullOrWhiteSpace(pageToken) || !seenPages.Add(pageToken))
            {
                return null;
            }
        }
    }

    private async Task<string> CreateCustomFieldAsync(
        string fieldName,
        string valueName,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        var uri = new Uri(BseSettings.BaseUri, BseSettings.CustomFieldsPath);
        var payload = JsonSerializer.Serialize(new
        {
            name = fieldName,
            allowCustomValues = false,
            global = true,
            values = new[] { valueName }
        });

        var body = await SendJsonWithRetriesAsync(HttpMethod.Post, uri, payload, progress, cancellationToken);
        var id = ExtractFirstIdFromBody(body);
        if (!string.IsNullOrWhiteSpace(id))
        {
            return id;
        }

        var resolved = await FindCustomFieldIdByNameAsync(fieldName, progress, cancellationToken);
        if (!string.IsNullOrWhiteSpace(resolved))
        {
            return resolved;
        }

        throw new InvalidOperationException($"Custom field {fieldName} was created but no id was returned.");
    }

    private async Task CreateCustomFieldValueAsync(
        string customFieldId,
        string valueName,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        var uri = new Uri(BseSettings.BaseUri, $"{BseSettings.CustomFieldsPath}/{Uri.EscapeDataString(customFieldId)}/values");
        var payload = JsonSerializer.Serialize(new { values = new[] { valueName } });

        try
        {
            await SendJsonWithRetriesAsync(HttpMethod.Post, uri, payload, progress, cancellationToken);
        }
        catch (Exception ex)
        {
            progress?.Report($"First attempt to create custom field value failed. {ex.Message}");
            var fallbackPayload = JsonSerializer.Serialize(new { value = valueName });
            await SendJsonWithRetriesAsync(HttpMethod.Post, uri, fallbackPayload, progress, cancellationToken);
        }
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

    private static Uri BuildCustomFieldsUri(string? pageToken)
    {
        var query = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(pageToken))
        {
            AppendQuery(query, "nextPage", pageToken);
        }

        AppendQuery(query, "max", "100");

        var builder = new UriBuilder(new Uri(BseSettings.BaseUri, BseSettings.CustomFieldsPath))
        {
            Query = query.ToString().TrimStart('?')
        };

        return builder.Uri;
    }

    private static Uri BuildCustomFieldValuesUri(string customFieldId, string? pageToken)
    {
        var query = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(pageToken))
        {
            AppendQuery(query, "nextPage", pageToken);
        }

        AppendQuery(query, "max", "100");

        var builder = new UriBuilder(new Uri(
            BseSettings.BaseUri,
            $"{BseSettings.CustomFieldsPath}/{Uri.EscapeDataString(customFieldId)}/values"))
        {
            Query = query.ToString().TrimStart('?')
        };

        return builder.Uri;
    }

    private static Uri BuildReimbursementsUri(string? pageToken)
    {
        var query = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(pageToken))
        {
            AppendQuery(query, "nextPage", pageToken);
        }
        else
        {
            AppendQuery(query, "filters", BseSettings.ReimbursementsApiFilter);
        }

        AppendQuery(query, "max", BseSettings.ReimbursementsPageSize.ToString(System.Globalization.CultureInfo.InvariantCulture));

        var builder = new UriBuilder(new Uri(BseSettings.BaseUri, BseSettings.ReimbursementsPath))
        {
            Query = query.ToString().TrimStart('?')
        };

        return builder.Uri;
    }

    private async Task<string> GetResponseBodyWithRetriesAsync(
        Uri uri,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        const int maxAttempts = 5;
        var lastStatus = string.Empty;
        var lastBody = string.Empty;

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            using var response = await _httpClient.GetAsync(uri, cancellationToken);
            var body = await response.Content.ReadAsStringAsync(cancellationToken);

            if (response.IsSuccessStatusCode)
            {
                return body;
            }

            lastStatus = $"{(int)response.StatusCode} {response.ReasonPhrase}";
            lastBody = body;

            if (response.StatusCode == HttpStatusCode.TooManyRequests)
            {
                var delay = GetRetryDelay(response) ?? TimeSpan.FromSeconds(60);
                if (delay < TimeSpan.FromSeconds(1))
                {
                    delay = TimeSpan.FromSeconds(1);
                }

                progress?.Report($"BILL rate limit reached. Waiting {Math.Ceiling(delay.TotalSeconds)} seconds before retrying...");
                await Task.Delay(delay, cancellationToken);
                continue;
            }

            break;
        }

        throw new InvalidOperationException(
            $"BILL API request failed with {lastStatus}.{Environment.NewLine}{lastBody}");
    }

    private async Task<string> SendJsonWithRetriesAsync(
        HttpMethod method,
        Uri uri,
        string json,
        IProgress<string>? progress,
        CancellationToken cancellationToken)
    {
        const int maxAttempts = 5;
        var lastStatus = string.Empty;
        var lastBody = string.Empty;

        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            using var request = new HttpRequestMessage(method, uri)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };

            using var response = await _httpClient.SendAsync(request, cancellationToken);
            var body = await response.Content.ReadAsStringAsync(cancellationToken);

            if (response.IsSuccessStatusCode)
            {
                return body;
            }

            lastStatus = $"{(int)response.StatusCode} {response.ReasonPhrase}";
            lastBody = body;

            if (response.StatusCode == HttpStatusCode.TooManyRequests)
            {
                var delay = GetRetryDelay(response) ?? TimeSpan.FromSeconds(60);
                if (delay < TimeSpan.FromSeconds(1))
                {
                    delay = TimeSpan.FromSeconds(1);
                }

                progress?.Report($"BILL rate limit reached. Waiting {Math.Ceiling(delay.TotalSeconds)} seconds before retrying...");
                await Task.Delay(delay, cancellationToken);
                continue;
            }

            break;
        }

        throw new InvalidOperationException(
            $"BILL API request failed with {lastStatus}.{Environment.NewLine}{lastBody}");
    }

    private static TimeSpan? GetRetryDelay(HttpResponseMessage response)
    {
        var retryAfter = response.Headers.RetryAfter;
        if (retryAfter is null)
        {
            return null;
        }

        if (retryAfter.Delta.HasValue)
        {
            return retryAfter.Delta.Value;
        }

        if (retryAfter.Date.HasValue)
        {
            var delay = retryAfter.Date.Value - DateTimeOffset.UtcNow;
            return delay > TimeSpan.Zero ? delay : TimeSpan.Zero;
        }

        return null;
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

    private static string? ExtractFirstIdFromBody(string body)
    {
        try
        {
            using var document = JsonDocument.Parse(body);
            var root = document.RootElement;

            if (root.ValueKind == JsonValueKind.Object)
            {
                var direct = GetObjectString(root, "uuid") ?? GetObjectString(root, "id");
                if (!string.IsNullOrWhiteSpace(direct))
                {
                    return direct;
                }
            }

            foreach (var item in ExtractListElements(root))
            {
                var id = GetObjectString(item, "uuid") ?? GetObjectString(item, "id");
                if (!string.IsNullOrWhiteSpace(id))
                {
                    return id;
                }
            }
        }
        catch
        {
        }

        return null;
    }

    private static IEnumerable<JsonElement> ExtractListElements(JsonElement root)
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
            foreach (var propertyName in new[] { "results", "reimbursements", "transactions", "items", "data", "records" })
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

    private static bool HasSyncedTag(IReadOnlyDictionary<string, string?> row)
    {
        return row.TryGetValue(SyncedTagName, out var value) && !string.IsNullOrWhiteSpace(value);
    }

    private static void BuildRowFromPayload(JsonElement transaction, IDictionary<string, string?> row)
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

    internal static bool HasAccountingIntegrationTransactions(JsonElement transaction)
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

    internal static bool HasApprovedAdminReviewer(JsonElement transaction)
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

    internal static bool IsDeclined(JsonElement transaction)
    {
        if (transaction.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        var transactionType = GetObjectString(transaction, "transactionType");
        return string.Equals(transactionType, "DECLINE", StringComparison.OrdinalIgnoreCase);
    }

    internal static int CountDuplicateTransactionIds(IEnumerable<JsonElement> transactions)
    {
        var seenIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var duplicates = 0;

        foreach (var transaction in transactions)
        {
            var id = GetObjectString(transaction, "id");
            if (string.IsNullOrWhiteSpace(id))
            {
                continue;
            }

            if (!seenIds.Add(id))
            {
                duplicates++;
            }
        }

        return duplicates;
    }

    private static bool IsReimbursementRetired(JsonElement reimbursement)
    {
        if (reimbursement.ValueKind != JsonValueKind.Object)
        {
            return true;
        }

        if (!reimbursement.TryGetProperty("retired", out var retired))
        {
            return false;
        }

        return retired.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String => string.Equals(retired.GetString(), "true", StringComparison.OrdinalIgnoreCase),
            _ => false
        };
    }

    private static string? GetReimbursementId(JsonElement reimbursement)
    {
        return GetObjectString(reimbursement, "id") ?? GetObjectString(reimbursement, "uuid");
    }

    private static bool IsPaymentPendingStatus(string? status)
    {
        if (string.IsNullOrWhiteSpace(status))
        {
            return false;
        }

        var normalized = status.Trim().Replace(' ', '_');
        if (normalized.StartsWith("PAYMENT_PENDING", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return normalized.StartsWith("PENDING_PAYMENT", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNotSyncedReimbursement(JsonElement reimbursement, out string? syncStatus)
    {
        syncStatus = GetObjectString(reimbursement, "syncStatus")
            ?? GetObjectString(reimbursement, "sync_status");

        if (!string.IsNullOrWhiteSpace(syncStatus))
        {
            return string.Equals(syncStatus, "NOT_SYNCED", StringComparison.OrdinalIgnoreCase);
        }

        if (HasAccountingIntegrationTransactions(reimbursement))
        {
            return false;
        }

        return true;
    }

    private static bool TryGetPaidTimestamp(JsonElement reimbursement, out DateTimeOffset paidAtUtc)
    {
        paidAtUtc = default;

        if (!reimbursement.TryGetProperty("statusHistory", out var history))
        {
            return false;
        }

        var candidates = new List<DateTimeOffset>();

        switch (history.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in history.EnumerateObject())
                {
                    if (!string.Equals(property.Name, "PAID", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    CollectOccurredTimes(property.Value, candidates);
                }
                break;
            case JsonValueKind.Array:
                foreach (var entry in history.EnumerateArray())
                {
                    if (entry.ValueKind != JsonValueKind.Object)
                    {
                        continue;
                    }

                    var status = GetObjectString(entry, "status");
                    if (string.Equals(status, "PAID", StringComparison.OrdinalIgnoreCase))
                    {
                        CollectOccurredTimes(entry, candidates);
                        continue;
                    }

                    if (entry.TryGetProperty("PAID", out var paidNode))
                    {
                        CollectOccurredTimes(paidNode, candidates);
                    }
                }
                break;
        }

        if (candidates.Count == 0)
        {
            return false;
        }

        paidAtUtc = candidates.Max();
        return true;
    }

    private static bool TryParseTimestamp(string? value, out DateTimeOffset parsed)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            parsed = default;
            return false;
        }

        if (DateTimeOffset.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out parsed))
        {
            return true;
        }

        return DateTimeOffset.TryParse(value, out parsed);
    }

    private static void CollectOccurredTimes(JsonElement element, List<DateTimeOffset> candidates)
    {
        var stack = new Stack<JsonElement>();
        stack.Push(element);

        while (stack.Count > 0)
        {
            var current = stack.Pop();

            switch (current.ValueKind)
            {
                case JsonValueKind.Object:
                    foreach (var property in current.EnumerateObject())
                    {
                        if (string.Equals(property.Name, "occurredTime", StringComparison.OrdinalIgnoreCase))
                        {
                            var value = property.Value.ValueKind == JsonValueKind.String
                                ? property.Value.GetString()
                                : property.Value.GetRawText();
                            if (TryParseTimestamp(value, out var parsed))
                            {
                                candidates.Add(parsed);
                            }
                        }

                        if (property.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                        {
                            stack.Push(property.Value);
                        }
                    }
                    break;
                case JsonValueKind.Array:
                    foreach (var item in current.EnumerateArray())
                    {
                        if (item.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                        {
                            stack.Push(item);
                        }
                    }
                    break;
            }
        }
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
