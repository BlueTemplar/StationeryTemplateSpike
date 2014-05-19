using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using System.IO;

namespace DocBuilder
{
    class Program
    {
        private const string SamplePath = @"C:\Users\brian.mcnulty\Documents\SampleDocs";

        static void Main(string[] args)
        {
            MergeDocuments();
        }

        public static void MergeDocuments()
        {
            using (var nullifier = new TemplateNullifier(SamplePath, "Stationery.docx"))
            {
                var tempFileName = nullifier.RemoveBodyContents();
                var sources = new List<Source>
                    {
                        new Source(new WmlDocument(Path.Combine(SamplePath, tempFileName)), false),
                        new Source(new WmlDocument(Path.Combine(SamplePath, "Template.docx")), false),
                    };
                DocumentBuilder.BuildDocument(sources, Path.Combine(SamplePath, "Output.docx"));
            }
        }
    }

    public class TemplateNullifier : IDisposable
    {
        private readonly string temporaryFileName;
        private readonly string stationeryFileName;

        public TemplateNullifier(string samplePath, string stationeryFileName)
        {
            this.temporaryFileName = Path.Combine(samplePath, "Oyster.docx");
            this.stationeryFileName = Path.Combine(samplePath, stationeryFileName);
        }

        public string RemoveBodyContents()
        {
            using (var stream = ReadInContents())
            {
                using (var wordDoc = WordprocessingDocument.Open(stream, true))
                {
                    RemoveBodyContents(wordDoc);
                    SaveDocument(stream);
                }
            }
            return temporaryFileName;
        }

        private MemoryStream ReadInContents()
        {
            byte[] byteArray = File.ReadAllBytes(stationeryFileName);
            return new MemoryStream(byteArray);
        }

        private void SaveDocument(MemoryStream stream)
        {
            File.WriteAllBytes(temporaryFileName, stream.ToArray());
        }

        private void RemoveBodyContents(WordprocessingDocument wordDoc)
        {
            wordDoc
                .MainDocumentPart
                .Document
                .Body
                .ChildElements
                .Where(IsNotSectionProperties)
                .ToList()
                .ForEach(RemoveChild);
        }

        private bool IsNotSectionProperties(OpenXmlElement childElement)
        {
            return childElement.GetType() != typeof (SectionProperties);
        }


        private void RemoveChild(OpenXmlElement childElement)
        {
            childElement.Remove();
        }

        public void Dispose()
        {
            if (File.Exists(temporaryFileName))
            {
                File.Delete(temporaryFileName);
            }
        }
    }
}
