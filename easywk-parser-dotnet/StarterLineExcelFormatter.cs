using DocumentFormat.OpenXml.Spreadsheet;

namespace easywk_parser_dotnet;

// Separate Klasse für OpenXML-Generierung
public static class StarterLineExcelFormatter
{
    public static Row ToXLSXRowWettkampf(StarterLine starterLine)
    {
        var result = new Row();
        result.Append(CreateCell(starterLine.Wettkampf));
        result.Append(CreateCell(starterLine.Lauf));
        result.Append(CreateCell(starterLine.Bahn));
        result.Append(CreateCell(starterLine.Teilnehmer));
        result.Append(CreateCell(starterLine.Verein));
        result.Append(CreateCell(starterLine.Meldezeit));
        return result;
    }

    public static Row ToXLSXRowTeilnehmer(StarterLine starterLine)
    {
        var result = new Row();
        result.Append(CreateCell(starterLine.Teilnehmer));
        result.Append(CreateCell(starterLine.Wettkampf));
        result.Append(CreateCell(starterLine.Lauf));
        result.Append(CreateCell(starterLine.Bahn));
        result.Append(CreateCell(starterLine.Verein));
        result.Append(CreateCell(starterLine.Meldezeit));
        return result;
    }

    private static Cell CreateCell(string value)
    {
        return new Cell
        {
            DataType = CellValues.String,
            CellValue = new CellValue(value)
        };
    }
}