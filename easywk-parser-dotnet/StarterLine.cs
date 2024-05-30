namespace easywk_parser_dotnet;

public class StarterLine
{
    private readonly string _wettkampf;
    private readonly string _lauf;
    private readonly string _bahn;
    private readonly string _teilnehmer;
    private readonly string _verein;
    private readonly string _meldezeit;


    public StarterLine(string wettkampf, string lauf, string bahn, string teilnehmer, string verein, string meldezeit) {
        this._wettkampf = wettkampf;
        this._lauf = lauf;
        this._bahn = bahn;
        this._teilnehmer = teilnehmer;
        this._verein = verein;
        this._meldezeit = meldezeit;
    }

    public string toString() {
        return this._wettkampf+"\t"+this._lauf+"\t"+this._bahn+"\t"+this._teilnehmer+"\t"+this._verein+"\t"+this._meldezeit;
    }
}