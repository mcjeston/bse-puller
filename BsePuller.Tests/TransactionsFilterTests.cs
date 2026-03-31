using System.Text.Json;
using BsePuller.Infrastructure;
using Xunit;

namespace BsePuller.Tests;

public class TransactionsFilterTests
{
    [Fact]
    public void HasApprovedAdminReviewer_ReturnsTrue_WhenAdminApprovedPresent()
    {
        using var doc = JsonDocument.Parse("""
            {"reviewers":[{"approverType":"ADMIN","status":"APPROVED"}]}
            """);
        Assert.True(BillApiClient.HasApprovedAdminReviewer(doc.RootElement));
    }

    [Fact]
    public void HasApprovedAdminReviewer_ReturnsFalse_WhenNoAdminApproval()
    {
        using var doc = JsonDocument.Parse("""
            {"reviewers":[{"approverType":"MANAGER","status":"APPROVED"}]}
            """);
        Assert.False(BillApiClient.HasApprovedAdminReviewer(doc.RootElement));
    }

    [Fact]
    public void IsDeclined_ReturnsTrue_WhenTransactionTypeDecline()
    {
        using var doc = JsonDocument.Parse("""{"transactionType":"DECLINE"}""");
        Assert.True(BillApiClient.IsDeclined(doc.RootElement));
    }

    [Fact]
    public void HasAccountingIntegrationTransactions_ReturnsTrue_WhenArrayHasItems()
    {
        using var doc = JsonDocument.Parse("""
            {"accountingIntegrationTransactions":[{"id":"1"}]}
            """);
        Assert.True(BillApiClient.HasAccountingIntegrationTransactions(doc.RootElement));
    }

    [Fact]
    public void HasAccountingIntegrationTransactions_ReturnsFalse_WhenMissing()
    {
        using var doc = JsonDocument.Parse("""{"id":"123"}""");
        Assert.False(BillApiClient.HasAccountingIntegrationTransactions(doc.RootElement));
    }

    [Fact]
    public void CountDuplicateTransactionIds_ReturnsDuplicateCount()
    {
        using var doc1 = JsonDocument.Parse("""{"id":"1"}""");
        using var doc2 = JsonDocument.Parse("""{"id":"1"}""");
        using var doc3 = JsonDocument.Parse("""{"id":"2"}""");

        var duplicates = BillApiClient.CountDuplicateTransactionIds(new[]
        {
            doc1.RootElement,
            doc2.RootElement,
            doc3.RootElement
        });

        Assert.Equal(1, duplicates);
    }
}
