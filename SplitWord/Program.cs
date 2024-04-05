using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Globalization;

class Program {
    public static void SezionaFileWord(string filePath)
    {
        // Open the Word document
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            // Get the main document part
            MainDocumentPart mainPart = doc.MainDocumentPart;

            // Get the document body
            Body body = mainPart.Document.Body;

            // Iterate over paragraphs
            foreach (Paragraph paragraph in body.Elements<Paragraph>())
            {
                // Check for chapter style using paragraph properties
                if (paragraph.ParagraphProperties != null)
                {
                    // Access paragraph style information
                    string styleName = null; // Initialize a variable to hold style name

                    // Try accessing style name using Val property (if available in your OpenXML version)
                    if (paragraph.ParagraphProperties.ParagraphStyleId != null)
                    {
                        try
                        {
                            styleName = paragraph.ParagraphProperties.ParagraphStyleId.Val; // Might throw exception in OpenXML 3.0.2
                        }
                        catch (NullReferenceException)
                        {
                            // Handle null ParagraphStyleId.Val gracefully (optional)
                            Console.WriteLine("Paragraph {0} has no style", paragraph.GetHashCode());
                        }
                    }

                    // Alternative approach (if Val property is unavailable)
                    if (styleName == null) // Check if Val property access failed
                    {
                        // Consider using a different property or method to access style name based on your OpenXML version
                        // (Consult OpenXML documentation for alternative ways to access style information)
                    }

                    // Check for chapter style
                    if (styleName != null && styleName == "TitoloCapitolo") // Replace with your chapter style name
                    {
                        Console.WriteLine("Capitolo: " + paragraph.InnerText);
                    }
                    else // Check for paragraphs not containing "Sottotitolo" in style name
                    {
                        if (styleName != null && !styleName.ToLower().Contains("sottotitolo")) // Use the retrieved style name
                        {
                            Console.WriteLine(paragraph.InnerText);
                        }
                    }
                }
            }
        }
    }

    static void Main(string[] args)
    {
        // Inserire il percorso del file Word
        string filePath = @"C:\Users\RMD.Cataldi\source\repos\SplitWord\SplitWord\DocDaSplittare.docx";

        SezionaFileWord(filePath);
    }
}

