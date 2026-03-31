using System.Diagnostics;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

using BsePuller.Modules.Settings;

namespace BsePuller.Modules.Reimbursements;

internal sealed record ReimbursementWebExportResult(bool Success, string? FilePath, string? ErrorMessage, Form? BrowserForm);

internal sealed class ReimbursementWebExporter
{
    private const string ReimbursementsUrl = "https://spend.bill.com/companies/Q29tcGFueTo0ODMyNg==/reimbursements";
    private readonly Form _owner;
    private readonly Action<string> _log;
    private readonly Action<string> _status;

    public ReimbursementWebExporter(Form owner, Action<string> log, Action<string> status)
    {
        _owner = owner;
        _log = log;
        _status = status;
    }

    public async Task<ReimbursementWebExportResult> RunAsync(CancellationToken cancellationToken)
    {
        var profilePath = BseSettings.GetWebView2UserDataFolder();
        var deleteProfileOnExit = false;
        var resetProfile = false;
        var keepBrowserOpen = false;
        var closeReason = string.Empty;
        Form? browserForm = null;
        WebView2? webView = null;
        TaskCompletionSource<bool>? browserClosedTcs = null;
        CancellationTokenSource? linkedCts = null;
        var tempDownloadFolder = BseSettings.GetTempFolder();
        string exportFile;

        try
        {
            Directory.CreateDirectory(tempDownloadFolder);
        }
        catch
        {
            tempDownloadFolder = Path.GetTempPath();
        }

        exportFile = Path.Combine(tempDownloadFolder, $"BSE-reimbursement-download-{DateTime.Now:yyyyMMdd-HHmmss}.csv");

        try
        {
            Directory.CreateDirectory(profilePath);
            var lastLogin = BseSettings.GetLastReimbursementWebLoginUtc();
            DateTimeOffset? lastActivity = lastLogin;
            if (!lastActivity.HasValue)
            {
                var lastWrite = Directory.GetLastWriteTimeUtc(profilePath);
                if (lastWrite != DateTime.MinValue)
                {
                    lastActivity = new DateTimeOffset(lastWrite, TimeSpan.Zero);
                }
            }

            var needsReset = !lastActivity.HasValue ||
                             (DateTimeOffset.UtcNow - lastActivity.Value) > TimeSpan.FromDays(30);

            if (needsReset && Directory.Exists(profilePath))
            {
                try
                {
                    Directory.Delete(profilePath, recursive: true);
                    resetProfile = true;
                }
                catch
                {
                }

                Directory.CreateDirectory(profilePath);
            }
        }
        catch (Exception)
        {
            profilePath = Path.Combine(Path.GetTempPath(), $"BsePuller-WebView2-{Guid.NewGuid():N}");
            try
            {
                Directory.CreateDirectory(profilePath);
                deleteProfileOnExit = true;
            }
            catch (Exception fallbackEx)
            {
                return new ReimbursementWebExportResult(false, null, $"Could not create browser profile. {fallbackEx.Message}", null);
            }
        }

        browserForm = new Form
        {
            Text = "BSE Reimbursements",
            StartPosition = FormStartPosition.CenterScreen,
            WindowState = FormWindowState.Normal,
            ShowInTaskbar = true,
            Width = 1200,
            Height = 900
        };

        webView = new WebView2
        {
            Dock = DockStyle.Fill
        };

        browserForm.Controls.Add(webView);
        browserForm.Show(_owner);

        try
        {
            browserClosedTcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            browserForm.FormClosed += (_, _) =>
            {
                closeReason = "Browser window was closed before the export finished.";
                _log("Browser window closed. Canceling reimbursement export.");
                _status("Reimbursement export canceled.");
                browserClosedTcs.TrySetResult(true);
                var cts = linkedCts;
                if (cts is not null)
                {
                    try
                    {
                        cts.Cancel();
                    }
                    catch (ObjectDisposedException)
                    {
                    }
                }
            };

            _status("Browser opened. Log in if needed.");
            _log("Opening BSE reimbursements in a browser window. Log in to continue.");
            if (deleteProfileOnExit)
            {
                _log("Using a temporary browser session. You will need to log in this time.");
            }
            else if (resetProfile)
            {
                _log("Cleared the saved login because it was older than 30 days.");
            }
            else
            {
                _log("Using the saved login session when available.");
            }

            var environment = await CoreWebView2Environment.CreateAsync(null, profilePath);
            await webView.EnsureCoreWebView2Async(environment);

            var downloadTcs = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            webView.CoreWebView2.DownloadStarting += (_, args) =>
            {
                args.ResultFilePath = exportFile;
                var operation = args.DownloadOperation;
                operation.StateChanged += (_, _) =>
                {
                    if (operation.State == CoreWebView2DownloadState.Completed)
                    {
                        downloadTcs.TrySetResult(exportFile);
                    }
                    else if (operation.State == CoreWebView2DownloadState.Interrupted)
                    {
                        downloadTcs.TrySetException(new InvalidOperationException("The CSV download was interrupted."));
                    }
                };
            };

            webView.CoreWebView2.Navigate(ReimbursementsUrl);

            _status("Complete login/MFA. Waiting for the reimbursements page to load...");
            _log("Waiting for reimbursements page to load.");

            if (!await WaitForReimbursementsPageAsync(webView, TimeSpan.FromMinutes(10), linkedCts.Token))
            {
                if (browserClosedTcs.Task.IsCompleted && !string.IsNullOrWhiteSpace(closeReason))
                {
                    return new ReimbursementWebExportResult(false, null, closeReason, null);
                }

                return new ReimbursementWebExportResult(false, null, "Timed out waiting for the reimbursements page to load after login.", null);
            }

            _status("Apply filters, then Export -> CSV -> Export -> Download.");
            _log("Apply your filters, then Export -> CSV -> Export -> Download. Waiting for the download...");

            _log("Waiting for reimbursement CSV download...");

            var downloadTask = downloadTcs.Task;
            var timeoutTask = Task.Delay(TimeSpan.FromMinutes(15), linkedCts.Token);
            var completed = await Task.WhenAny(downloadTask, browserClosedTcs.Task, timeoutTask);
            if (completed == browserClosedTcs.Task || timeoutTask.IsCanceled)
            {
                return new ReimbursementWebExportResult(false, null, closeReason, null);
            }

            if (completed != downloadTask)
            {
                return new ReimbursementWebExportResult(false, null, "Timed out waiting for the reimbursement CSV download. Make sure you click Export, choose CSV, and then Download.", null);
            }

            var path = await downloadTask;
            BseSettings.SaveLastReimbursementWebLoginUtc(DateTimeOffset.UtcNow);
            keepBrowserOpen = true;
            return new ReimbursementWebExportResult(true, path, null, browserForm);
        }
        catch (OperationCanceledException) when (browserClosedTcs is not null && browserClosedTcs.Task.IsCompleted)
        {
            return new ReimbursementWebExportResult(false, null, closeReason, null);
        }
        catch (Exception ex)
        {
            return new ReimbursementWebExportResult(false, null, ex.Message, null);
        }
        finally
        {
            var ctsToDispose = linkedCts;
            linkedCts = null;
            ctsToDispose?.Dispose();
            try
            {
                if (!keepBrowserOpen && browserForm is not null)
                {
                    browserForm.Close();
                    browserForm.Dispose();
                }
            }
            catch
            {
            }

            try
            {
                if (!keepBrowserOpen && deleteProfileOnExit && Directory.Exists(profilePath))
                {
                    Directory.Delete(profilePath, recursive: true);
                }
            }
            catch
            {
            }
        }
    }

    public static async Task<Form?> OpenReimbursementsWindowAsync(Form owner, Action<string> log, Action<string> status)
    {
        var profilePath = BseSettings.GetWebView2UserDataFolder();
        try
        {
            Directory.CreateDirectory(profilePath);
        }
        catch
        {
        }

        var browserForm = new Form
        {
            Text = "BSE Reimbursements",
            StartPosition = FormStartPosition.CenterScreen,
            WindowState = FormWindowState.Normal,
            ShowInTaskbar = true,
            Width = 1200,
            Height = 900
        };

        var webView = new WebView2
        {
            Dock = DockStyle.Fill
        };

        browserForm.Controls.Add(webView);
        browserForm.Show(owner);

        try
        {
            var environment = await CoreWebView2Environment.CreateAsync(null, profilePath);
            await webView.EnsureCoreWebView2Async(environment);
            webView.CoreWebView2.Navigate(ReimbursementsUrl);
            log("Reopened reimbursements page for manual sync.");
            status("Reopened reimbursements page for manual sync.");
            return browserForm;
        }
        catch (Exception ex)
        {
            log($"Could not reopen reimbursements page. {ex.Message}");
            status("Could not reopen reimbursements page.");
            try
            {
                browserForm.Close();
                browserForm.Dispose();
            }
            catch
            {
            }

            return null;
        }
    }

    private async Task<bool> ApplyFiltersAndExportAsync(WebView2 webView, CancellationToken cancellationToken)
    {
        if (!await OpenFilterPanelAsync(webView, TimeSpan.FromMinutes(3), cancellationToken))
        {
            return false;
        }

        if (!await WaitForAndClickAsync(webView, "Status", TimeSpan.FromSeconds(30), cancellationToken))
        {
            return false;
        }

        if (!await WaitForAndClickAsync(webView, "Payment pending", TimeSpan.FromSeconds(30), cancellationToken, containsMatch: true))
        {
            return false;
        }

        if (!await WaitForAndClickAsync(webView, "Paid", TimeSpan.FromSeconds(30), cancellationToken))
        {
            return false;
        }

        if (!await WaitForAndClickAsync(webView, "Sync", TimeSpan.FromSeconds(30), cancellationToken))
        {
            return false;
        }

        if (!await WaitForAndClickAsync(webView, "Not synced", TimeSpan.FromSeconds(30), cancellationToken, containsMatch: true))
        {
            return false;
        }

        if (!await WaitForAndClickAsync(webView, "Done", TimeSpan.FromSeconds(30), cancellationToken))
        {
            return false;
        }

        if (!await WaitForSelectAllAsync(webView, TimeSpan.FromSeconds(30), cancellationToken))
        {
            return false;
        }

        if (!await WaitForAndClickByAttributeAsync(webView, "Export", TimeSpan.FromSeconds(30), cancellationToken))
        {
            if (!await WaitForAndClickAsync(webView, "Export", TimeSpan.FromSeconds(30), cancellationToken))
            {
                return false;
            }
        }

        if (!await WaitForAndClickAsync(webView, "CSV", TimeSpan.FromSeconds(30), cancellationToken))
        {
            return false;
        }

        if (!await WaitForAndClickAsync(webView, "Export", TimeSpan.FromSeconds(30), cancellationToken))
        {
            return false;
        }

        if (!await WaitForAndClickAsync(webView, "Download", TimeSpan.FromMinutes(3), cancellationToken))
        {
            return false;
        }

        return true;
    }

    private async Task<bool> WaitForReimbursementsPageAsync(WebView2 webView, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();
        var notifiedSlow = false;
        var nextDismissAttempt = DateTime.UtcNow;

        while (sw.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var currentUrl = webView.Source?.ToString() ?? string.Empty;
            var onPage = currentUrl.Contains("/reimbursements", StringComparison.OrdinalIgnoreCase);

            if (onPage)
            {
                var readyScript = "(function() {"
                    + "const candidates = ['+ Add filter','Add filter','Add filters','Filters','Export','Reimbursements'];"
                    + "const selectors = ['button','[role=\"button\"]','a','span','div','label'];"
                    + "const normalize = (value) => (value || '').replace(/\\s+/g,' ').trim();"
                    + "const attrSelectors = ["
                    + "  '[aria-label*=\"Add filter\" i]',"
                    + "  '[aria-label*=\"Filter\" i]',"
                    + "  '[title*=\"Filter\" i]',"
                    + "  '[data-testid*=\"filter\" i]',"
                    + "  '[data-testid*=\"filters\" i]',"
                    + "  '[aria-label*=\"Export\" i]',"
                    + "  '[title*=\"Export\" i]',"
                    + "  '[data-testid*=\"export\" i]'"
                    + "];"
                    + "const rootHasMatches = (root) => {"
                    + "  for (const selector of attrSelectors) {"
                    + "    if (root.querySelector(selector)) return true;"
                    + "  }"
                    + "  const elements = selectors.flatMap(s => Array.from(root.querySelectorAll(s)));"
                    + "  for (const el of elements) {"
                    + "    const text = normalize(el.innerText || el.textContent || el.getAttribute('aria-label') || el.getAttribute('title') || '');"
                    + "    if (!text) continue;"
                    + "    for (const target of candidates) {"
                    + "      if (text === target || text.includes(target)) return true;"
                    + "    }"
                    + "  }"
                    + "  if (root.querySelector('table, [role=\"grid\"], [role=\"table\"]')) return true;"
                    + "  if (root.querySelector('input[type=\"checkbox\"], [role=\"checkbox\"]')) return true;"
                    + "  return false;"
                    + "};"
                    + "const roots = [document];"
                    + "for (const el of Array.from(document.querySelectorAll('*'))) {"
                    + "  if (el.shadowRoot) roots.push(el.shadowRoot);"
                    + "}"
                    + "for (const root of roots) {"
                    + "  try { if (rootHasMatches(root)) return true; } catch (err) { }"
                    + "}"
                    + "return false;"
                    + "})();";

                var result = await webView.CoreWebView2.ExecuteScriptAsync(readyScript);
                if (IsScriptTrue(result))
                {
                    return true;
                }
            }

            if (!notifiedSlow && sw.Elapsed > TimeSpan.FromSeconds(30))
            {
                notifiedSlow = true;
                _log("Still waiting for the reimbursements page to finish loading after login.");
                _status("Waiting for the reimbursements page to finish loading...");
            }

            if (DateTime.UtcNow >= nextDismissAttempt)
            {
                nextDismissAttempt = DateTime.UtcNow.AddSeconds(5);
                await TryDismissOverlayAsync(webView, cancellationToken);
            }

            await Task.Delay(800, cancellationToken);
        }

        return false;
    }

    private async Task<bool> OpenFilterPanelAsync(WebView2 webView, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();

        while (sw.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var panelReadyScript = "(function() {"
                + "const normalize = (value) => (value || '').replace(/\\s+/g,' ').trim();"
                + "const elements = Array.from(document.querySelectorAll('button,[role=\"button\"],span,div,label'));"
                + "for (const el of elements) {"
                + "  const text = normalize(el.innerText || el.textContent || el.getAttribute('aria-label') || el.getAttribute('title') || '');"
                + "  if (text === 'Status' || text.includes('Status')) return true;"
                + "}"
                + "return false;"
                + "})();";

            var panelReady = await webView.CoreWebView2.ExecuteScriptAsync(panelReadyScript);
            if (IsScriptTrue(panelReady))
            {
                return true;
            }

            if (await TryClickAddFilterAsync(webView, cancellationToken))
            {
                await Task.Delay(800, cancellationToken);
            }
            else
            {
                await Task.Delay(600, cancellationToken);
            }
        }

        _log("Could not open the filter panel.");
        return false;
    }

    private async Task<bool> TryClickAddFilterAsync(WebView2 webView, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var script = "(function() {"
            + "const candidates = ['+ Add filter','Add filter','Add filters','Filters'];"
            + "const selectors = ['button','[role=\"button\"]','a','span','div','label'];"
            + "const normalize = (value) => (value || '').replace(/\\s+/g,' ').trim();"
            + "const elements = selectors.flatMap(s => Array.from(document.querySelectorAll(s)));"
            + "for (const el of elements) {"
            + "  const text = normalize(el.innerText || el.textContent || '');"
            + "  if (!text) continue;"
            + "  for (const target of candidates) {"
            + "    if (text === target || text.includes(target)) {"
            + "      el.click();"
            + "      return true;"
            + "    }"
            + "  }"
            + "}"
            + "const aria = document.querySelector('[aria-label*=\"Add filter\" i]')"
            + "  || document.querySelector('[aria-label*=\"Filter\" i]')"
            + "  || document.querySelector('[title*=\"Filter\" i]')"
            + "  || document.querySelector('[data-testid*=\"filter\" i]')"
            + "  || document.querySelector('[data-testid*=\"filters\" i]');"
            + "if (aria) { aria.click(); return true; }"
            + "return false;"
            + "})();";

        var result = await webView.CoreWebView2.ExecuteScriptAsync(script);
        if (IsScriptTrue(result))
        {
            _log("Opened filter panel.");
            return true;
        }

        return false;
    }

    private async Task<bool> WaitForAndClickAsync(
        WebView2 webView,
        string text,
        TimeSpan timeout,
        CancellationToken cancellationToken,
        bool containsMatch = false)
    {
        var sw = Stopwatch.StartNew();

        while (sw.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var script = "(function() {"
                + "const normalize = (value) => (value || '').replace(/\\s+/g,' ').trim();"
                + "const target = " + EscapeForScript(text) + ";"
                + "const contains = " + (containsMatch ? "true" : "false") + ";"
                + "const selectors = ['button','[role=\"button\"]','a','span','div','label','li'];"
                + "for (const selector of selectors) {"
                + "  const elements = Array.from(document.querySelectorAll(selector));"
                + "  for (const el of elements) {"
                + "    const text = normalize(el.innerText || el.textContent || el.getAttribute('aria-label') || el.getAttribute('title') || '');"
                + "    if (!text) continue;"
                + "    if ((contains && text.includes(target)) || text === target) {"
                + "      el.click();"
                + "      return true;"
                + "    }"
                + "  }"
                + "}"
                + "return false;"
                + "})();";

            var result = await webView.CoreWebView2.ExecuteScriptAsync(script);
            if (IsScriptTrue(result))
            {
                _log($"Clicked \"{text}\".");
                return true;
            }

            await Task.Delay(600, cancellationToken);
        }

        _log($"Could not find \"{text}\" within {timeout.TotalSeconds:n0} seconds.");
        return false;
    }

    private async Task<bool> WaitForAndClickByAttributeAsync(
        WebView2 webView,
        string value,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();

        while (sw.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var script = "(function() {"
                + "const target = " + EscapeForScript(value) + ";"
                + "const selectors = ["
                + "  `[aria-label=\"${target}\"]`,"
                + "  `[title=\"${target}\"]`,"
                + "  `[data-testid=\"${target}\"]`,"
                + "  `[aria-label*=\"${target}\" i]`,"
                + "  `[title*=\"${target}\" i]`,"
                + "  `[data-testid*=\"${target}\" i]`"
                + "];"
                + "for (const selector of selectors) {"
                + "  const el = document.querySelector(selector);"
                + "  if (el) {"
                + "    el.click();"
                + "    return true;"
                + "  }"
                + "}"
                + "return false;"
                + "})();";

            var result = await webView.CoreWebView2.ExecuteScriptAsync(script);
            if (IsScriptTrue(result))
            {
                _log($"Clicked \"{value}\" button.");
                return true;
            }

            await Task.Delay(600, cancellationToken);
        }

        return false;
    }

    private async Task<bool> TryDismissOverlayAsync(WebView2 webView, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var script = "(function() {"
            + "const normalize = (value) => (value || '').replace(/\\s+/g,' ').trim();"
            + "const candidates = ['Got it','Got it!','Close','Dismiss','Not now','Skip','No thanks','Okay','Ok'];"
            + "const selectors = ['button','[role=\"button\"]','a','span','div'];"
            + "for (const selector of selectors) {"
            + "  const elements = Array.from(document.querySelectorAll(selector));"
            + "  for (const el of elements) {"
            + "    const text = normalize(el.innerText || el.textContent || el.getAttribute('aria-label') || el.getAttribute('title') || '');"
            + "    if (!text) continue;"
            + "    for (const target of candidates) {"
            + "      if (text === target) {"
            + "        el.click();"
            + "        return true;"
            + "      }"
            + "    }"
            + "  }"
            + "}"
            + "const close = document.querySelector('[aria-label=\"Close\"],[aria-label*=\"Close\" i]');"
            + "if (close) { close.click(); return true; }"
            + "return false;"
            + "})();";

        var result = await webView.CoreWebView2.ExecuteScriptAsync(script);
        return IsScriptTrue(result);
    }

    private async Task<bool> WaitForSelectAllAsync(WebView2 webView, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();

        while (sw.Elapsed < timeout)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var script = """
                (function() {
                    const header = document.querySelector('thead') || document.querySelector('[role="rowgroup"]');
                    let checkbox = null;
                    if (header) {
                        checkbox = header.querySelector('input[type="checkbox"]') || header.querySelector('[role="checkbox"]');
                    }
                    if (!checkbox) {
                        checkbox = document.querySelector('input[type="checkbox"][aria-label*="Select all"]');
                    }
                    if (checkbox) {
                        checkbox.click();
                        return true;
                    }
                    return false;
                })();
                """;

            var result = await webView.CoreWebView2.ExecuteScriptAsync(script);
            if (IsScriptTrue(result))
            {
                _log("Selected all reimbursements on the page.");
                return true;
            }

            await Task.Delay(600, cancellationToken);
        }

        return false;
    }

    private static bool IsScriptTrue(string? scriptResult)
    {
        return string.Equals(scriptResult?.Trim(), "true", StringComparison.OrdinalIgnoreCase);
    }

    private static string EscapeForScript(string value)
    {
        return System.Text.Json.JsonSerializer.Serialize(value);
    }
}
