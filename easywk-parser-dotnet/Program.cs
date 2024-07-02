// See https://aka.ms/new-console-template for more information

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using UglyToad.PdfPig;

namespace easywk_parser_dotnet
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Display the number of command line arguments.
            Console.WriteLine("Arguments count: " + args.Length);
            Console.WriteLine("Hello, World!");
            //var filePath = @"/Users/nschle85/IdeaProjects/MeldelisteParser/meldeliste.pdf";
            var filePath = args[0];
            using (PdfDocument document = PdfDocument.Open(filePath))
            {
                var parser = new ParserImpl(document);
                var result = parser.Parse();
                
                /*
                using (StreamWriter outputFile = new StreamWriter(filePath + "-fromdotnet_"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+".csv"))
                {
                    ExportCSV(result, outputFile);
                }
                
                using (StreamWriter outputFile = new StreamWriter(filePath + "-fromdotnet_"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+".xls"))
                {
                    ExportXLS(result, outputFile);
                }
                */
                
                using (SpreadsheetDocument outputFile = SpreadsheetDocument.Create(filePath + "-fromdotnet_"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+".xlsx", SpreadsheetDocumentType.Workbook))
                {
                    ExportXLSX(result, outputFile);
                }

                using (SpreadsheetDocument inputFile =
                       SpreadsheetDocument.Open(@"/Users/nschle85/IdeaProjects/MeldelisteParser/meldeergebnis_veranstaltungkm24.pdf-fromdotnet_20240628-201918_V01.xlsx", false))
                {
                    foreach (var otto in inputFile.WorkbookPart.WorksheetParts)
                    {
                        foreach (var tabledef in otto.TableDefinitionParts)
                        {
                            foreach (var tableColumn in tabledef.Parts)
                            {
                                Console.WriteLine();
                            }
                        }
                    }
                }
            }
        }

        private static void ExportCSV(List<StarterLine> result, StreamWriter outputFile)
        {
            outputFile.WriteLine("Wettkampf\tLauf\tBahn\tSchwimmer\tVerein\tMeldezeit");
            foreach (var starterLine in result)
            {
                Console.WriteLine(starterLine.toString());
                outputFile.WriteLine(starterLine.toString());
            }
        }
        
        private static void ExportXLS(List<StarterLine> result, StreamWriter outputFile)
        { 
            String topPart ="""
                <html            
                    xmlns:o="urn:schemas-microsoft-com:office:office">
                    xmlns:x="urn:schemas-microsoft-com:office:excel"
                    xmlns="http://www.w3.org/TR/REC-html40">
                        <head>
                            <meta http-equiv=Content-Type content="text/html; charset=utf-8">
                            <meta name=ProgId content=Excel.Sheet>
                            <meta name=Generator content="Microsoft Excel 11">
                            <style id="STI_5961_Styles">
                        </head>
                <body>
                    <div id="STI_5961" align=center x:publishsource="Excel">
                     <table x:str border=0 cellpadding=0 cellspacing=0 width=2020 style='border-collapse: 
                       collapse;table-layout:fixed;width:1518pt'>
                        <col width=147 style='mso-width-source:userset;mso-width-alt:5376;width:110pt'>
                        <col width=55 style='mso-width-source:userset;mso-width-alt:2011;width:41pt'>
                        <col width=55 style='mso-width-source:userset;mso-width-alt:2011;width:41pt'>
                        <col width=55 style='mso-width-source:userset;mso-width-alt:2011;width:41pt'>
                        <col width=55 style='mso-width-source:userset;mso-width-alt:2011;width:41pt'>
                        <col width=55 style='mso-width-source:userset;mso-width-alt:2011;width:41pt'>
                        
                        <tr class=xlGeneralBold height=51 style='mso-height-source:userset;height:38.25pt'>
                            <td height=51 class=xlGeneralBold width=147 style='height:38.25pt;width:110pt'>Wettkampf</td>
                            <td class=xlGeneralBold width=55 style='border-left:none;width:41pt'>Lauf</td>
                            <td class=xlGeneralBold width=55 style='border-left:none;width:41pt'>Bahn</td>
                            <td class=xlGeneralBold width=55 style='border-left:none;width:41pt'>Schwimmer</td>
                            <td class=xlGeneralBold width=55 style='border-left:none;width:41pt'>Verein</td>
                            <td class=xlGeneralBold width=55 style='border-left:none;width:41pt'>Meldezeit</td>
                       </tr>
                """; 
            
            String bottomPart ="""
                       </table>
                      </div>
                    </body>
                </html>
                """;
            
            //outputFile.WriteLine("Wettkampf\tLauf\tBahn\tSchwimmer\tVerein\tMeldezeit");
            outputFile.WriteLine(topPart);
            foreach (var starterLine in result)
            {
                Console.WriteLine(starterLine.toHtmlRow());
                outputFile.WriteLine(starterLine.toHtmlRow());
            }
            outputFile.WriteLine(bottomPart);
        }
        
        private static void ExportXLSX(List<StarterLine> result, SpreadsheetDocument doc)
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
            headerRowAll.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Wettkampf")
            });
            headerRowAll.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Lauf")
            });
            headerRowAll.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Bahn")
            });
            headerRowAll.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Teilnehmer")
            });
            headerRowAll.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Verein")
            });
            headerRowAll.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Meldezeit")
            });
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
            headerRowWettkampf.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Wettkampf")
            });
            headerRowWettkampf.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Lauf")
            });
            headerRowWettkampf.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Bahn")
            });
            headerRowWettkampf.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Teilnehmer")
            });
            headerRowWettkampf.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Verein")
            });
            headerRowWettkampf.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Meldezeit")
            });
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
            PageSetup pageSetupWettkampf = new PageSetup()
            {
                Orientation = OrientationValues.Landscape,
                PaperSize = 9, //A4 https://learn.microsoft.com/de-de/dotnet/api/documentformat.openxml.spreadsheet.pagesetup?view=openxml-3.0.1
                FitToWidth = 1,
                FitToHeight = 0
                
            };
            worksheetPartWettkampf.Worksheet.Append(pageSetupWettkampf);
            
            TableColumns tableColumnsWettkampf = new TableColumns() { Count = 6 };
            tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)1, Name = "Wettkampf" });
            tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)2, Name = "Lauf" });
            tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)3, Name = "Bahn" });
            tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)4, Name = "Teilnehmer" });
            tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)5, Name = "Verein" });
            tableColumnsWettkampf.Append(new TableColumn() { Id = (UInt32)6, Name = "Meldezeit" });
            AddTable(worksheetPartWettkampf,tableColumnsWettkampf,1);
            
            
            //Sheet Teilnehmer
            WorksheetPart worksheetPartTeilnehmer = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPartTeilnehmer.Worksheet = new Worksheet(new SheetData());
            
            Sheet sheetTeilnehmer = new Sheet()
                { Id = doc.WorkbookPart.GetIdOfPart(worksheetPartTeilnehmer), SheetId = 3, Name = "Sheet Teilnehmer" };
            sheets.Append(sheetTeilnehmer);
            // workbookPart.Workbook.Save();

            SheetData sheetDataTeilnehmer = worksheetPartTeilnehmer.Worksheet.Elements<SheetData>().First();
            
            var headerRowTeilnehmer = new Row();
            headerRowTeilnehmer.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Teilnehmer")
            });
            headerRowTeilnehmer.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Wettkampf")
            });
            headerRowTeilnehmer.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Lauf")
            });
            headerRowTeilnehmer.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Bahn")
            });
            headerRowTeilnehmer.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Verein")
            });
            headerRowTeilnehmer.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Meldezeit")
            });
            sheetDataTeilnehmer.Append(headerRowTeilnehmer);
            
            foreach (var starterLine in result)
            {
                //Console.WriteLine(starterLine.toString());
                sheetDataTeilnehmer.Append(starterLine.toXLSXRowTeilnehmer());
            }
            
            SheetProperties spTeilnehmer = new SheetProperties(new PageSetupProperties());
            worksheetPartTeilnehmer.Worksheet.SheetProperties = spTeilnehmer;
            worksheetPartTeilnehmer.Worksheet.SheetProperties.PageSetupProperties.FitToPage = true;
            
            PageSetup pageSetupTeilnehmer = new PageSetup()
            {
                Orientation = OrientationValues.Landscape,
                PaperSize = 9, //A4 https://learn.microsoft.com/de-de/dotnet/api/documentformat.openxml.spreadsheet.pagesetup?view=openxml-3.0.1
                FitToWidth = 1,
                FitToHeight = 0
            };
            worksheetPartTeilnehmer.Worksheet.Append(pageSetupTeilnehmer);
            
            // Spalten definieren
            TableColumns tableColumnsTeilnehmer = new TableColumns() { Count = 6 };
            tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)4, Name = "Teilnehmer" });
            tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)1, Name = "Wettkampf" });
            tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)2, Name = "Lauf" });
            tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)3, Name = "Bahn" });
            tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)5, Name = "Verein" });
            tableColumnsTeilnehmer.Append(new TableColumn() { Id = (UInt32)6, Name = "Meldezeit" });
            AddTable(worksheetPartTeilnehmer,tableColumnsTeilnehmer, 2);
            
        }
        
        private static void AddTable(WorksheetPart tableSheetPart, TableColumns tableColumns, int tableNo)
        {
            // Tabellen-Definition
            TableDefinitionPart tableDefinitionPart = tableSheetPart.AddNewPart<TableDefinitionPart>();
            // int tableNo = tableSheetPart.TableDefinitionParts.Count() + 1;
            Console.WriteLine(tableNo);

            // Bereich der Tabelle
            string reference = "A1:F1269"; // Zellenbereich für die Tabelle

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
        }
        private static void AddTeilnehmerTable(WorksheetPart tableSheetPart, TableColumns tableColumns)
        {
            // Tabellen-Definition
            TableDefinitionPart tableDefinitionPart = tableSheetPart.AddNewPart<TableDefinitionPart>();
            int tableNo = tableSheetPart.TableDefinitionParts.Count() + 2;
            Console.WriteLine(tableNo);

            // Bereich der Tabelle
            string reference = "A1:F1269"; // Zellenbereich für die Tabelle

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
                Name = "TableStyleLight1", 
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
        }
    }
    
}


