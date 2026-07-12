namespace KellyCashApp.Models
{
    // Records for Microsoft VMS data and dictionary (VmsIdentifier, WorkerName) to match with OIR data - will refactor from imported memory to a pull from static path > OpenXML in the future
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