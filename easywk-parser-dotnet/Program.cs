// See https://aka.ms/new-console-template for more information

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
                var excelExporter = new ExcelExporter(filePath, result);
                excelExporter.ExportXLSX();
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
        
        private static void ExportHTML(List<StarterLine> result, StreamWriter outputFile)
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
    }
    
}


