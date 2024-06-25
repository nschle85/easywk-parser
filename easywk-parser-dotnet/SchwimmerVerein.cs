using DocumentFormat.OpenXml.Spreadsheet;

namespace easywk_parser_dotnet;

public class SchwimmerVerein
{
    public readonly string _schwimmer;
    public readonly string _verein;

    public SchwimmerVerein(string schwimmer, string verein)
    {
        this._schwimmer = schwimmer;
        this._verein = verein;
    }
}