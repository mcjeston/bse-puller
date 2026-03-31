namespace BsePuller.Modules.Reimbursements;

internal sealed class ReimbursementsModule
{
    private readonly Form _owner;
    private readonly Action<string> _log;
    private readonly Action<string> _status;
    private Form? _reimbursementBrowserForm;

    public ReimbursementsModule(Form owner, Action<string> log, Action<string> status)
    {
        _owner = owner;
        _log = log;
        _status = status;
    }

    public async Task EnsureBrowserOpenAsync()
    {
        if (_reimbursementBrowserForm is not null && !_reimbursementBrowserForm.IsDisposed)
        {
            return;
        }

        _log("Reopening reimbursements page for manual sync.");
        _status("Reopening reimbursements page for manual sync...");
        _reimbursementBrowserForm = await ReimbursementWebExporter.OpenReimbursementsWindowAsync(_owner, _log, _status);
        if (_reimbursementBrowserForm is not null)
        {
            _reimbursementBrowserForm.FormClosed += (_, _) => _reimbursementBrowserForm = null;
        }
    }
}
