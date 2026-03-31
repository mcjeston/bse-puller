using BsePuller.Infrastructure;

namespace BsePuller.Modules.Transactions;

internal sealed class TransactionsService
{
    private readonly BillApiClient _client;

    public TransactionsService(BillApiClient client)
    {
        _client = client;
    }

    public Task<TransactionPullResult> GetFilteredTransactionsAsync(IProgress<string>? progress, CancellationToken cancellationToken)
    {
        return _client.GetFilteredTransactionsAsync(progress, cancellationToken);
    }
}
