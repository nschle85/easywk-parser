namespace easywk_parser_dotnet;

public class StarterLine
{
    public readonly string Wettkampf;
    public readonly string Lauf;
    public readonly string Bahn;
    public readonly string Teilnehmer;
    public readonly string Verein;
    public readonly string Meldezeit;

    public StarterLine(string wettkampf, string lauf, string bahn, string teilnehmer, string verein, string meldezeit)
    {
        Wettkampf = wettkampf;
        Lauf = lauf;
        Bahn = bahn;
        Teilnehmer = teilnehmer;
        Verein = verein;
        Meldezeit = meldezeit;
    }

    public override string ToString()
    {
        return $"{Wettkampf}\t{Lauf}\t{Bahn}\t{Teilnehmer}\t{Verein}\t{Meldezeit}";
    }
}