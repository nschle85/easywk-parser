namespace easywk_parser_dotnet;

public class SchwimmerVerein
{
    public readonly string Schwimmer;
    public readonly string Verein;

    public SchwimmerVerein(string schwimmer, string verein)
    {
        this.Schwimmer = schwimmer;
        this.Verein = verein;
    }
}