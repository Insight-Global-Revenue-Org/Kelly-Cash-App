namespace KellyCashApp.Models
{
    internal record KenvueVmsMatch(
        string ParentInvoiceId,
        string WorkerName,
        string FeeDescription,
        DateTime? FeeMonth
    );
}