// See https://aka.ms/new-console-template for more information

using DocumentFormat.OpenXml;
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
                using (StreamWriter outputFile = new StreamWriter(filePath + "-fromdotnet.csv"))
                {
                    ExportCSV(result, outputFile);
                }
                
                using (StreamWriter outputFile = new StreamWriter(filePath + "-fromdotnet.xls"))
                {
                    ExportXLS(result, outputFile);
                }
                
                using (SpreadsheetDocument outputFile = SpreadsheetDocument.Create(filePath + "-fromdotnet.xlsx", SpreadsheetDocumentType.Workbook))
                {
                    ExportXLSX(result, outputFile);
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
            WorkbookPart workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet()
                { Id = doc.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
            sheets.Append(sheet);
            // workbookPart.Workbook.Save();

            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            
            var headerRow = new Row();
            headerRow.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Wettkampf")
            });
            headerRow.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Lauf")
            });
            headerRow.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Bahn")
            });
            headerRow.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Teilnehmer")
            });
            headerRow.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Verein")
            });
            headerRow.Append(new Cell()
            {
                DataType = CellValues.String,
                CellValue = new CellValue("Meldezeit")
            });
            sheetData.Append(headerRow);
            
            foreach (var starterLine in result)
            {
                Console.WriteLine(starterLine.toString());
                sheetData.Append(starterLine.toXLSXRow());
            }
            
            workbookPart.Workbook.Save();
            
            
            
        }
    }
    
}


