using DocumentFormat.OpenXml.Spreadsheet;

namespace easywk_parser_dotnet;

public class StarterLine
{
    private readonly string _wettkampf;
    private readonly string _lauf;
    private readonly string _bahn;
    private readonly string _teilnehmer;
    private readonly string _verein;
    private readonly string _meldezeit;


    public StarterLine(string wettkampf, string lauf, string bahn, string teilnehmer, string verein, string meldezeit)
    {
        this._wettkampf = wettkampf;
        this._lauf = lauf;
        this._bahn = bahn;
        this._teilnehmer = teilnehmer;
        this._verein = verein;
        this._meldezeit = meldezeit;
    }

    public string toString()
    {
        return this._wettkampf + "\t" + this._lauf + "\t" + this._bahn + "\t" + this._teilnehmer + "\t" + this._verein +
               "\t" + this._meldezeit;
    }

    public string toHtmlRow()
    {
        return $$"""
                     <tr height=17 style='height:12.75pt'>
                      <td height=17 class=xlGeneral style='height:12.75pt'>{{this._wettkampf}}</td>
                      <td height=17 class=xlGeneral style='height:12.75pt'>{{this._lauf}}</td>
                      <td height=17 class=xlGeneral style='height:12.75pt'>{{this._bahn}}</td>
                      <td height=17 class=xlGeneral style='height:12.75pt'>{{this._teilnehmer}}</td>
                      <td height=17 class=xlGeneral style='height:12.75pt'>{{this._verein}}</td>
                      <td height=17 class=xlGeneral style='height:12.75pt'>{{this._meldezeit}}</td>
                    </tr>
                 """;
    }

    public Row toXLSXRowWettkampf()
    {
        var result = new Row();
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._wettkampf)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._lauf)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(this._bahn)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._teilnehmer)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._verein)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._meldezeit)
        });
        return result;
    }
    
    public Row toXLSXRowTeilnehmer()
    {
        var result = new Row();
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._teilnehmer)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._wettkampf)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._lauf)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(this._bahn)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._verein)
        });
        result.Append(new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(this._meldezeit)
        });
        return result;
    }
}