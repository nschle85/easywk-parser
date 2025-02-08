using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.PageSegmenter;
using UglyToad.PdfPig.Util;

namespace easywk_parser_dotnet;

public class MeldeListenPdfReader
{
    private string _inputFilePath;

    public MeldeListenPdfReader(string inputFilePath)
    {
        _inputFilePath = inputFilePath;
    }

    public List<string> GetTextLines()
    {
        var result = new List<string>();
        using (PdfDocument document = PdfDocument.Open(_inputFilePath))
        {
            foreach (var page in document.GetPages())
            {
                var words = DefaultWordExtractor.Instance.GetWords(page.Letters);
                var blocks = DefaultPageSegmenter.Instance.GetBlocks(words);
                foreach (var textBlock in blocks)
                {
                    foreach (var textBlockTextLine in textBlock.TextLines)
                    {
                        result.Add(textBlockTextLine.Text);
                    }
                }
            }
        }
        return result;
    }
}