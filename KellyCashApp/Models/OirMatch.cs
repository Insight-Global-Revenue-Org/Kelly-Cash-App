namespace KellyCashApp.Models
{
    public record OirMatch(
        string DocumentNumber,
        decimal RemainingAmount,
        string ClientProjects = ""
    );
}