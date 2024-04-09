using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

class Program {
    static void DividiWord(string fileInput, string cartellaOutput)
    {
        using (var doc = WordprocessingDocument.Open(fileInput, false))
        {
            int capitolo = 1;
            var paragrafi = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>();
            static void Main(string[] args)
            {
                // Inserire il percorso del file Word
                string filePath = @"C:\Users\RMD.Cataldi\source\repos\SplitWord\SplitWord\DocDaSplittare.docx";

                SezionaFileWord(filePath);
            }

            foreach (var paragrafo in paragrafi)
            {
                if (paragrafo.ParagraphProperties.StyleId == "Titolo")
                {
                    var nomeFile = Path.Combine(cartellaOutput, $"Capitolo{capitolo:00}.docx");
                    using (var docCapitolo = WordprocessingDocument.Create(nomeFile, WordprocessingDocumentType.Document))
                    {
                        var mainPart = docCapitolo.AddMainDocumentPart();
                        mainPart.Document = new Document();

                        // Copia il paragrafo titolo
                        var paragrafoClone = paragrafo.CloneNode(true);
                        mainPart.Document.AppendChild(paragrafoClone);

                        // Copia il corpo del capitolo
                        var paragrafiSuccessivi = paragrafo.NextSibling;
                        while (paragrafiSuccessivi != null && paragrafiSuccessivi.ParagraphProperties.StyleId != "Titolo")
                        {
                            var paragrafoSuccessivoClone = paragrafiSuccessivi.CloneNode(true);
                            mainPart.Document.AppendChild(paragrafoSuccessivoClone);
                            paragrafiSuccessivi = paragrafiSuccessivi.NextSibling;
                        }

                        // Copia tabelle e immagini
                        CopiaTabelleImmagini(doc, mainPart, paragrafo, paragrafiSuccessivi);
                    }
                    capitolo++;
                }
            }
        }
    }

    private static void SezionaFileWord(string filePath)
    {
        throw new NotImplementedException();
    }

    static void CopiaTabelleImmagini(WordprocessingDocument doc, MainDocumentPart mainPart, Body body, OpenXmlElement paragrafo)
    {
        var relazioni = doc.MainDocumentPart.GetPartsOfType<ImagePart>();
        foreach (var relazione in relazioni)
        {
            var id = relazione.RelationshipId;
            var img = body.Descendants<Drawing>().Where(x => x.Id == id).FirstOrDefault();
            if (img != null)
            {
                var imgClone = img.CloneNode(true);
                mainPart.Document.AppendChild(imgClone);
            }
        }

        var relazioniTabelle = doc.MainDocumentPart.GetPartsOfType<TablePart>();
        foreach (var relazioneTabella in relazioniTabelle)
        {
            var id = relazioneTabella.RelationshipId;
            var tabella = body.Descendants<Table>().Where(x => x.Id == id).FirstOrDefault();
            if (tabella != null)
            {
                var tabellaClone = tabella.CloneNode(true);
                mainPart.Document.AppendChild(tabellaClone);
            }
        }
    }
    static void UnisciWord(string cartellaInput, string fileOutput)
    {
        using (var doc = WordprocessingDocument.Create(fileOutput, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();

            var files = Directory.GetFiles(cartellaInput, "*.docx");
            foreach (var file in files.OrderBy(x => int.Parse(Path.GetFileNameWithoutExtension(x).Substring(8))))
            {
                using (var docCapitolo = WordprocessingDocument.Open(file, false))
                {
                    var bodyCapitolo = docCapitolo.MainDocumentPart.Document.Body;
                    foreach (var elemento in bodyCapitolo.Elements())
                    {
                        mainPart.Document.AppendChild(elemento.CloneNode(true));
                    }
                }
            }
        }
    }
}


//    static void Main(string[] args)
//    {
//        // Inserire il percorso del file Word
//        string filePath = @"C:\Users\RMD.Cataldi\source\repos\SplitWord\SplitWord\DocDaSplittare.docx";

//        SezionaFileWord(filePath);
//    }
//}

