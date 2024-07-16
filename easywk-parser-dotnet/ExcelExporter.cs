using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace easywk_parser_dotnet;

public static class ExcelExporter
{
    private static String ColNameWettkampf="Wettkampf";
    private static String ColNameLauf="Lauf";
    private static String ColNameBahn="Bahn";
    private static String ColNameTeilnehmer="Teilnehmer";
    private static String ColNameVerein="Verein";
    private static String ColNameMeldezeit="Meldezeit";
    private static String ColNameKommentar="Kommentar";
    
    
    
    public static void ExportXLSX(List<StarterLine> result, SpreadsheetDocument doc)
    {
        //GeneralPart
        WorkbookPart workbookPart = doc.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
            
        //Sheet all
        WorksheetPart worksheetPartAll = workbookPart.AddNewPart<WorksheetPart>();
            
        worksheetPartAll.Worksheet = new Worksheet(new SheetData());
        Sheet sheetAll = new Sheet()
            { Id = doc.WorkbookPart.GetIdOfPart(worksheetPartAll), SheetId = 1, Name = "Sheet All" };
        sheets.Append(sheetAll);
        // workbookPart.Workbook.Save();

        SheetData sheetDataAll = worksheetPartAll.Worksheet.Elements<SheetData>().First();
            
        var headerRowAll = new Row();
        headerRowAll.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameWettkampf) });
        headerRowAll.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameLauf) });
        headerRowAll.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameBahn) });
        headerRowAll.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameTeilnehmer) });
        headerRowAll.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameVerein) });
        headerRowAll.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameMeldezeit) });
        sheetDataAll.Append(headerRowAll);
            
        foreach (var starterLine in result)
        {
            //Console.WriteLine(starterLine.toString());
            sheetDataAll.Append(starterLine.toXLSXRowWettkampf());
        }
            
        //Sheet Wettkampf
        WorksheetPart worksheetPartWettkampf = workbookPart.AddNewPart<WorksheetPart>();
            
        worksheetPartWettkampf.Worksheet = new Worksheet(new SheetData());
        Sheet sheetWettkampf = new Sheet()
            { Id = doc.WorkbookPart.GetIdOfPart(worksheetPartWettkampf), SheetId = 2, Name = "Sheet Wettkampf" };
        sheets.Append(sheetWettkampf);

        SheetData sheetDataWettkampf = worksheetPartWettkampf.Worksheet.Elements<SheetData>().First();
            
        var headerRowWettkampf = new Row();
        headerRowWettkampf.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameWettkampf) });
        headerRowWettkampf.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameLauf) });
        headerRowWettkampf.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameBahn) });
        headerRowWettkampf.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameTeilnehmer) });
        headerRowWettkampf.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameVerein) });
        headerRowWettkampf.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameMeldezeit) });
        headerRowWettkampf.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameKommentar) });
        sheetDataWettkampf.Append(headerRowWettkampf);
            
        foreach (var starterLine in result)
        {
            //Console.WriteLine(starterLine.toString());
            sheetDataWettkampf.Append(starterLine.toXLSXRowWettkampf());
        }
            
            
        SheetProperties spWettkampf = new SheetProperties(new PageSetupProperties());
        worksheetPartWettkampf.Worksheet.SheetProperties = spWettkampf;
        worksheetPartWettkampf.Worksheet.SheetProperties.PageSetupProperties.FitToPage = true;
            
        // Set the page orientation to landscape (horizontal)
        PageSetup pageSetupWettkampf = GetPageSetup();
        
        worksheetPartWettkampf.Worksheet.Append(pageSetupWettkampf);
            
        TableColumns tableColumnsWettkampf = new TableColumns() { Count = 6 };
        tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)1, Name = ColNameWettkampf });
        tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)2, Name = ColNameLauf });
        tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)3, Name = ColNameBahn });
        tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)4, Name = ColNameTeilnehmer });
        tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)5, Name = ColNameVerein });
        tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)6, Name = ColNameMeldezeit });
        tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)7, Name = ColNameKommentar });
        AddTable(worksheetPartWettkampf,tableColumnsWettkampf,1, GetTsvErdingFilterColumn(4));
            
            
        //Sheet Teilnehmer
        WorksheetPart worksheetPartTeilnehmer = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPartTeilnehmer.Worksheet = new Worksheet(new SheetData());
            
        Sheet sheetTeilnehmer = new Sheet()
            { Id = doc.WorkbookPart.GetIdOfPart(worksheetPartTeilnehmer), SheetId = 3, Name = "Sheet Teilnehmer" };
        sheets.Append(sheetTeilnehmer);
        // workbookPart.Workbook.Save();

        SheetData sheetDataTeilnehmer = worksheetPartTeilnehmer.Worksheet.Elements<SheetData>().First();
            
        var headerRowTeilnehmer = new Row();
        headerRowTeilnehmer.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameTeilnehmer) });
        headerRowTeilnehmer.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameWettkampf) });
        headerRowTeilnehmer.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameLauf) });
        headerRowTeilnehmer.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameBahn) });
        headerRowTeilnehmer.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameVerein) });
        headerRowTeilnehmer.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameMeldezeit) });
        headerRowTeilnehmer.Append(new Cell() { DataType = CellValues.String, CellValue = new CellValue(ColNameKommentar) });
        sheetDataTeilnehmer.Append(headerRowTeilnehmer);
        
        foreach (var starterLine in result)
        {
            //Console.WriteLine(starterLine.toString());
            sheetDataTeilnehmer.Append(starterLine.toXLSXRowTeilnehmer());
        }
            
        SheetProperties spTeilnehmer = new SheetProperties(new PageSetupProperties());
        worksheetPartTeilnehmer.Worksheet.SheetProperties = spTeilnehmer;
        worksheetPartTeilnehmer.Worksheet.SheetProperties.PageSetupProperties.FitToPage = true;

        PageSetup pageSetupTeilnehmer = GetPageSetup();
        
        worksheetPartTeilnehmer.Worksheet.Append(pageSetupTeilnehmer);
            
        // Spalten definieren
        TableColumns tableColumnsTeilnehmer = new TableColumns() { Count = 6 };
        tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)4, Name = ColNameTeilnehmer });
        tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)1, Name = ColNameWettkampf });
        tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)2, Name = ColNameLauf });
        tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)3, Name = ColNameBahn });
        tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)5, Name = ColNameVerein });
        tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)6, Name = ColNameMeldezeit });
        tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)7, Name = ColNameKommentar });
        AddTable(worksheetPartTeilnehmer,tableColumnsTeilnehmer, 2, GetTsvErdingFilterColumn(4));
        
        doc.Save();
    }

    private static PageSetup GetPageSetup()
    {
        PageSetup pageSetup = new PageSetup()
        {
            Orientation = OrientationValues.Landscape,
            PaperSize = 9, //A4 https://learn.microsoft.com/de-de/dotnet/api/documentformat.openxml.spreadsheet.pagesetup?view=openxml-3.0.1
            FitToWidth = 1,
            FitToHeight = 0
        };
        return pageSetup;
    }
    
    private static FilterColumn GetTsvErdingFilterColumn(int id)
    {
        FilterColumn filterColumn = new FilterColumn() { ColumnId = (UInt32)id }; // id starts from 0
        Filters filters = new Filters();
        Filter filter = new Filter() { Val = "TSV Erding" };
        filters.Append(filter);
        filterColumn.Append(filters);
        return filterColumn;
    }
    
    private static void AddTable(WorksheetPart tableSheetPart, TableColumns tableColumns, int tableNo, FilterColumn tsvErdingFilterColumn)
    {
        // Tabellen-Definition
        TableDefinitionPart tableDefinitionPart = tableSheetPart.AddNewPart<TableDefinitionPart>();
        // int tableNo = tableSheetPart.TableDefinitionParts.Count() + 1;
        Console.WriteLine(tableNo);

        // Bereich der Tabelle
        string reference = "A1:G2007"; // Zellenbereich für die Tabelle
        
        // Tabelle erstellen
        Table table = new Table() 
        { 
            Id = (UInt32)tableNo, 
            Name = "Table" + tableNo, 
            DisplayName = "Table" + tableNo, 
            Reference = reference, 
            TotalsRowShown = false 
        };
        
        AutoFilter autoFilter = new AutoFilter() { Reference = reference };
        

        // Tabellenstil definieren
        TableStyleInfo tableStyleInfo = new TableStyleInfo() 
        { 
            Name = "TableStyleMedium9", 
            ShowFirstColumn = false, 
            ShowLastColumn = false, 
            ShowRowStripes = true, 
            ShowColumnStripes = false 
        };

        table.Append(autoFilter);
        table.Append(tableColumns);
        table.Append(tableStyleInfo);

        tableDefinitionPart.Table = table;

        // Tabelle zur Arbeitsblatt hinzufügen
        TableParts tableParts = new TableParts() { Count = (UInt32)1 };
        TablePart tablePart = new TablePart()
        {
            Id = tableSheetPart.GetIdOfPart(tableDefinitionPart)
        };

        tableParts.Append(tablePart);
        tableSheetPart.Worksheet.Append(tableParts);
        
        autoFilter.Append(tsvErdingFilterColumn);
        
    }
}