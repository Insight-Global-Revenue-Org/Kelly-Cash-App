namespace KellyCashApp.Models
{
    public record MicrosoftVmsMatch(
        string VmsIdentifier,
        string WorkerName,
        decimal AggregateInvoicedNet,
        decimal Hours,
        decimal RtRate,
        decimal OtRate,
        decimal DtRate
    );
}