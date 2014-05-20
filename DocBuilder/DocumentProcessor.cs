using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;

namespace DocBuilder
{
    public class DocumentProcessor : IDisposable
    {
        private const string SamplePath = @"C:\Users\brian.mcnulty\Documents\SampleDocs";

        public void MergeDocuments()
        {
            ProcessStationeryTemplate();
            DocumentBuilder.BuildDocument(Sources(), Path.Combine(SamplePath, "Output.docx"));
            PostProcessDocument();
        }

        private void PostProcessDocument()
        {
            RemoveKnownParagraph(Path.Combine(SamplePath, "Output.docx"), RemoveKnownParagraph);
        }

        private void ProcessStationeryTemplate()
        {
            File.Copy(Path.Combine(SamplePath, "Stationery.docx"), Path.Combine(SamplePath, "Temporary.docx"), true);
            RemoveKnownParagraph(Path.Combine(SamplePath, "Temporary.docx"), RemoveRedundantContent);
        }

        private List<Source> Sources()
        {
            return new List<Source>
                {
                    new Source(new WmlDocument(Path.Combine(SamplePath, "Temporary.docx")), false),
                    new Source(new WmlDocument(Path.Combine(SamplePath, "Template.docx")), false),
                };
        }

        private void RemoveKnownParagraph(string fileName, Action<WordprocessingDocument> processDocument)
        {
            using (var mem = ReadInContents(fileName))
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(mem, true))
                {
                    processDocument(wordDoc);
                }
                SaveDocument(mem, fileName);
            }
        }

        private MemoryStream ReadInContents(string fileName)
        {
            byte[] byteArray = File.ReadAllBytes(fileName);
            var mem = new MemoryStream();
            mem.Write(byteArray, 0, byteArray.Length);
            return mem;
        }

        private void RemoveRedundantContent(WordprocessingDocument wordDoc)
        {
            var body =
                wordDoc
                    .MainDocumentPart
                    .Document
                    .Body;

            var sectionProperties = body.ChildElements.First<SectionProperties>();
            var firstParagraph = body.ChildElements.First<Paragraph>();
            firstParagraph.ParagraphId = new HexBinaryValue("FFA");
            body.RemoveAllChildren();
            body.InsertAt(firstParagraph, 0);
            body.InsertAt(sectionProperties, 1);
        }

        private void SaveDocument(MemoryStream mem, string fileName)
        {
            using (var fileStream = new FileStream(fileName, FileMode.Create))
            {
                mem.WriteTo(fileStream);
            }
        }

        private void RemoveKnownParagraph(WordprocessingDocument wordDoc)
        {
            wordDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>().Single(x => x.ParagraphId == "FFA").Remove();
        }

        public void Dispose()
        {
            if (File.Exists(Path.Combine(SamplePath, "Temporary.docx")))
            {
                File.Delete(Path.Combine(SamplePath, "Temporary.docx"));
            }
        }
    }
}