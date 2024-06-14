using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.Util;

namespace easywk_parser_dotnet;

public class ParserImpl
{
    private readonly PdfDocument _document;
    private string _wettkampf = "";
    private string _lauf = "";
    
    private readonly Regex _wettkampfRegex = new Regex(@"(Wettkampf) (\d+) (.*)");
    private readonly Regex _laufRegex = new Regex(@"(Lauf) (\d+)/(\d+) (.*)");
    // match (Bahn), (BahnNummer), (schwimmer: Teilnehmer, Jahrgang, Verein etc), (Meldezeit)
    private readonly Regex _bahnRegex = new Regex(@"(Bahn) (\d+) (.+?) (\d+:\d+,\d+)");
    
    //Splash Meet Manager, 11.79888
    private readonly Regex _wettkampfRegexNew = new Regex(@"(Wettkampf) (\d+),(.*)");
    private readonly Regex _laufRegexNew = new Regex(@"(Lauf) (\d+) von (\d+)(.*)");
    private readonly Regex _bahnRegexNew = new Regex(@"^(\d+) (.*?)");

    
    public ParserImpl(PdfDocument document)
        {
            this._document = document;
        }

        public List<StarterLine> Parse()
        {
            List<StarterLine> result = []; 
            foreach (var page in _document.GetPages())
            {
                var pageText = page.Text;
                var words = DefaultWordExtractor.Instance.GetWords(page.Letters);
                var blocks = DefaultPageSegmenter.Instance.GetBlocks(words);
                foreach (var textBlock in blocks)
                {
                    foreach (var textBlockTextLine in textBlock.TextLines)
                    {
                        var line = ParseTextLine(textBlockTextLine);
                        if (line!=null)
                        {
                            result.Add(line);
                        }
                    }
                }
            }

            return result;
        }

        private StarterLine? ParseTextLine(TextLine textLine)
        {
            StarterLine? result = null;
            var line = textLine.Text;
            //Console.WriteLine(line);
            
            if (line.StartsWith("Wettkampf"))
            {
                var  wettkampfMatcher = _wettkampfRegex.Match(line);
                var wettkampfMAtcherNew = _wettkampfRegexNew.Match(line);
                if (wettkampfMatcher.Success)
                {
                    _wettkampf = wettkampfMatcher.Groups[1].Value + " " + ToLeadingZeroString(wettkampfMatcher.Groups[2].Value) + " " + wettkampfMatcher.Groups[3].Value;
                    Console.WriteLine("Found parsed :"+_wettkampf);
                }
                else
                {
                    Console.Error.WriteLine("No matching Wettkampf:" + line);
                }
            }
            else if (line.StartsWith("Lauf"))
            {
                // Matcher laufMatcher = laufPattern.matcher(line);
                var laufMatcher = _laufRegex.Match(line);
                var laufMatcherNew = _laufRegexNew.Match(line);
                if (laufMatcher.Success)
                {
                    _lauf = laufMatcher.Groups[1].Value + " " + ToLeadingZeroString(laufMatcher.Groups[2].Value) + 
                            "/" + ToLeadingZeroString(laufMatcher.Groups[3].Value) + " " + laufMatcher.Groups[4].Value;
                }
                else if (laufMatcherNew.Success)
                {
                    _lauf = laufMatcherNew.Groups[1].Value + " " + ToLeadingZeroString(laufMatcherNew.Groups[2].Value) +
                            "/" + ToLeadingZeroString(laufMatcherNew.Groups[3].Value);
                }
                else
                {
                    Console.Error.WriteLine("No matching Lauf:" + line);
                }
            }
            else if (line.StartsWith("Bahn"))
            {
                // match (Bahn), (BahnNummer), (schwimmer: Teilnehmer, Jahrgang, Verein etc), (Meldezeit)
                var r = _bahnRegex;
                // Now create matcher object.
                var m = r.Match(line);
                if (m.Success)
                {
                    // var bahn = m.group(1);
                    var bahnNr = m.Groups[2].Value;
                    var schwimmer = m.Groups[3].Value;

                    //extract verein if possible and update schwimmer and verein
                    string verein = "";

                    // Phdksfh, Fdsds  2013/ TSV Neuburg
                    var schwimmerVereinPattern = new Regex(@"(.+?)\/ (.*)");
                    var schwimmerVereinMatcher = schwimmerVereinPattern.Match(schwimmer);

                    // Rdsds, Sfdfd  1978/AK 45	SV Lohhof
                    var schwimmerAkVereinPattern = new Regex(@"(.+?\/AK \d+) (.*)");
                    var schwimmerAkVereinMatcher = schwimmerAkVereinPattern.Match(schwimmer);

                    //Rdsds, Sfdfd  1978 45	SV Lohhof
                    var schwimmerJgVereinPattern = new Regex(@"(.+? \d+) (.*)");
                    var schwimmerJgVereinMatcher = schwimmerJgVereinPattern.Match(schwimmer);

                    //Rdsds, Sfdfd  Offen 45 SV Lohhof
                    var schwimmerOffenVereinPattern = new Regex(@"(.+? Offen) (.*)");
                    var schwimmerOffenVereinMatcher = schwimmerOffenVereinPattern.Match(schwimmer);

                    if (schwimmerVereinMatcher.Success) {
                             schwimmer = schwimmerVereinMatcher.Groups[1].Value;
                             verein = schwimmerVereinMatcher.Groups[2].Value;
                    }
                    else if (schwimmerAkVereinMatcher.Success) {
                             schwimmer = schwimmerAkVereinMatcher.Groups[1].Value;
                             verein = schwimmerAkVereinMatcher.Groups[2].Value;
                    }
                    else if (schwimmerJgVereinMatcher.Success) {
                             schwimmer = schwimmerJgVereinMatcher.Groups[1].Value;
                             verein = schwimmerJgVereinMatcher.Groups[2].Value;
                    }
                    else if (schwimmerOffenVereinMatcher.Success) {
                             schwimmer = schwimmerOffenVereinMatcher.Groups[1].Value;
                             verein = schwimmerOffenVereinMatcher.Groups[2].Value;
                    }

                    else {
                            Console.Error.WriteLine("Could not parse: " + schwimmer);
                    }
                    
                    var meldezeit = m.Groups[4].Value;
                    
                    result = new StarterLine(_wettkampf, _lauf, "Bahn "+bahnNr, schwimmer, verein, meldezeit);
                }
            }
            else if (_bahnRegexNew.Match(line).Success)
            {
                // match (BahnNummer), (schwimmer: Teilnehmer, Jahrgang, Verein etc), (Meldezeit)
                //var r = new Regex(@"^(\d+) (.*)");
                var r = new Regex(@"^(\d+) (.+?) ((\d+:\d+\.\d+)|(NT)|(\d+\.\d+))");
                // Now create matcher object.
                var m = r.Match(line);
                if (m.Success)
                {
                    // var bahn = m.group(1);
                    var bahnNr = m.Groups[1].Value;
                    var schwimmer = m.Groups[2].Value;
                    
                    string verein = "";
                    
                    // Schwimmer GER Verein
                    var schwimmerVereinPattern = new Regex(@"(.+?) GER (.*)");
                    var schwimmerVereinMatcher = schwimmerVereinPattern.Match(schwimmer);
                    
                    // Rdsds, Sfdfd  1978/AK 45	SV Lohhof                                     
                    var schwimmerAkVereinPattern = new Regex(@"(.+?\/AK \d+) (.*)");          
                    var schwimmerAkVereinMatcher = schwimmerAkVereinPattern.Match(schwimmer); 
                    
                    if (schwimmerVereinMatcher.Success) {
                        schwimmer = schwimmerVereinMatcher.Groups[1].Value;
                        verein = schwimmerVereinMatcher.Groups[2].Value;
                    }
                    else if (schwimmerAkVereinMatcher.Success) {
                        schwimmer = schwimmerAkVereinMatcher.Groups[1].Value;
                        verein = schwimmerAkVereinMatcher.Groups[2].Value;
                    }
                    else
                    {
                        Console.WriteLine("Cannot parse Schwimmer: "+schwimmer);
                    }
                    var meldezeit = m.Groups[3].Value;
                    
                    if (_wettkampf.Length>0 && _lauf.Length >0)
                    {
                        result = new StarterLine(_wettkampf, _lauf, bahnNr, schwimmer, verein, meldezeit);
                        //Console.WriteLine("Matched new bahn" + line);
                    }
                    
                }
                else
                {
                    Console.WriteLine("No Matched bahn "+ line);
                }
            }
            else
            {
                Console.Error.WriteLine("NO MATCH "+ line);
            }

            return result;
        }
        
        private string ToLeadingZeroString(string value)
        {
            return Int32.Parse(value).ToString("00");
        }
}