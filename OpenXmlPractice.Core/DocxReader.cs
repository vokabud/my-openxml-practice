using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlPractice.Core;

public static class DocxReader
{
    public static string ReadControllValue(string filePath, string tagName)
    {
        using var wordDoc = WordprocessingDocument.Open(filePath, false);

        // Get the main document part
        var mainPart = wordDoc.MainDocumentPart;

        if (mainPart == null)
        {
            Console.WriteLine("No main part found in the document.");
            return string.Empty;
        }

        var controls = mainPart
            .Document?
            .Body?
            .Descendants<SdtElement>();

        if (controls == null)
        {
            Console.WriteLine("No controls found in the document.");
            return string.Empty;
        }

        var control = controls
            .FirstOrDefault(_ => _
                .SdtProperties?
                .GetFirstChild<SdtAlias>()?
                .Val?
                .Value == tagName);

        if (control == null)
        {
            Console.WriteLine($"Control with tag name '{tagName}' not found.");
            return string.Empty;
        }

        return control
            .Descendants<Text>()
            .Select(t => t.Text)
            .First();
    }
}
