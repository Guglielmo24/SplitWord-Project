using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Linq;

public class WordDocumentSplitter
{
    public void SplitDocument(string sourceFilePath)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(sourceFilePath, true))
        {
            var doc = wordDoc.MainDocumentPart.Document;

            var titles = doc.Body.Elements<Paragraph>().Where(p => IsHeading2(p)).ToList();
            Console.WriteLine("Titoli trovati nel documento:");
            foreach (var title in titles)
            {
                Console.WriteLine(title.InnerText);
            }
            int sectionIndex = 0;

            string outputDirectory = "C:\\Users\\Gu.Langella\\Documents\\split"; 

            foreach (var title in titles)
            {
                sectionIndex++;
                string newFileName = $"Section_{sectionIndex}.docx";
                string newFilePath = Path.Combine(outputDirectory, newFileName);

                while (File.Exists(newFilePath))
                {
                    newFileName = $"Section_{sectionIndex++}.docx";
                    newFilePath = Path.Combine(outputDirectory, newFileName);
                }

                using (WordprocessingDocument newDoc = WordprocessingDocument.Create(newFilePath, WordprocessingDocumentType.Document))
                {
                    newDoc.AddMainDocumentPart();
                    newDoc.MainDocumentPart.Document = new Document();
                    newDoc.MainDocumentPart.Document.Body = new Body();
                    CloneStyles(wordDoc, newDoc);

                    var current = title.NextSibling();
                    while (current != null && !IsHeading2(current as Paragraph))
                    {
                        newDoc.MainDocumentPart.Document.Body.Append(current.CloneNode(true));
                        current = current.NextSibling();
                    }

                    newDoc.MainDocumentPart.Document.Save();
                }
            }
        }
    }

    private bool IsHeading2(Paragraph paragraph)
    {
        return paragraph != null &&
               paragraph.ParagraphProperties != null &&
               paragraph.ParagraphProperties.ParagraphStyleId != null &&
               paragraph.ParagraphProperties.ParagraphStyleId.Val != null &&
               paragraph.ParagraphProperties.ParagraphStyleId.Val == "Titolo2";
    }

    private void CloneStyles(WordprocessingDocument sourceDoc, WordprocessingDocument targetDoc)
    {
        var sourcePart = sourceDoc.MainDocumentPart.StyleDefinitionsPart;
        if (sourcePart != null)
        {
            var targetPart = targetDoc.MainDocumentPart.StyleDefinitionsPart;
            if (targetPart == null)
            {
                targetPart = targetDoc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            }

            using (Stream sourceStream = sourcePart.GetStream())
            {
                targetPart.FeedData(sourceStream);
            }
        }
    }

    static void Main(string[] args)
    {
        string sourceFilePath = "C:\\Users\\Gu.Langella\\Documents\\GitHub\\SplitWord-Project\\SplitWord\\DocDaSplittare.docx";

        WordDocumentSplitter splitter = new WordDocumentSplitter();
        splitter.SplitDocument(sourceFilePath);

        Console.WriteLine("Documento diviso con successo.");
        Console.ReadLine(); 
    }
}
