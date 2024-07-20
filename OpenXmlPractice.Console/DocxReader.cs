using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlPractice.Console;

internal static class DocxReader
{
    public static void Read(string filePath,  string fileName)
    {
        using (var wordDoc = WordprocessingDocument.Open($@"{filePath}\{fileName}", false))
        {
            // Get the main document part
            var mainPart = wordDoc.MainDocumentPart;

            if (mainPart == null)
            {
                return;
            }

            var controls = mainPart
                .Document
                .Body
                .Descendants<SdtElement>();

            foreach (var control in controls)
            {
                string controlName = control.SdtProperties.GetFirstChild<SdtAlias>()?.Val?.Value;
                string controlValue = control.Descendants<Text>().Select(t => t.Text).FirstOrDefault();

                System.Console.WriteLine($"Control Name: {controlName}");
                System.Console.WriteLine($"Control Value: {controlValue}");
                System.Console.WriteLine("------------------------------");
            }
        }
    }
}
