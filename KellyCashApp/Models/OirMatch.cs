namespace KellyCashApp.Models
{
    // Records for OIR data and dictionary (DocumentNumber, RemainingAmount, ClientProjects)
    public record OirMatch(
        string DocumentNumber,
        decimal RemainingAmount,
        string ClientProjects = ""
    );
}