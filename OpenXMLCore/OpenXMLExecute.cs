using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Xml;

namespace OpenXMLCore
{
    public static class OpenXMLExecute
    {
        public static void CreateDocx(string filePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());

                // String msg contains the text, "Hello, Word!"
                run.AppendChild(new Text("Hello, Word!"));
            }
        }

        public static void SetMainTitle(Document document, string title)
        {
            
        }

        public static void AddNewPart(string document)
        {
            // Create a new word processing document.
            WordprocessingDocument wordDoc =
               WordprocessingDocument.Create(document,
               WordprocessingDocumentType.Document);

            // Add the MainDocumentPart part in the new word processing document.
            var mainDocPart = wordDoc.AddNewPart<MainDocumentPart>
        ("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml", "rId1");
            mainDocPart.Document = new Document();

            // Add Content
            Body body = mainDocPart.Document.AppendChild(new Body());
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            // String msg contains the text, "Hello, Word!"
            run.AppendChild(new Text("Hello, Word!"));

            // Add the CustomFilePropertiesPart part in the new word processing document.
            var customFilePropPart = wordDoc.AddCustomFilePropertiesPart();
            customFilePropPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

            // Add the CoreFilePropertiesPart part in the new word processing document.
            var coreFilePropPart = wordDoc.AddCoreFilePropertiesPart();
            using (XmlTextWriter writer = new
        XmlTextWriter(coreFilePropPart.GetStream(FileMode.Create), System.Text.Encoding.UTF8))
            {
                writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\r\n<cp:coreProperties xmlns:cp=\"https://schemas.openxmlformats.org/package/2006/metadata/core-properties\"></cp:coreProperties>");
                writer.Flush();
            }

            // Add the DigitalSignatureOriginPart part in the new word processing document.
            wordDoc.AddNewPart<DigitalSignatureOriginPart>("rId4");

            // Add the ExtendedFilePropertiesPart part in the new word processing document.
            var extendedFilePropPart = wordDoc.AddNewPart<ExtendedFilePropertiesPart>("rId5");
            extendedFilePropPart.Properties =
        new DocumentFormat.OpenXml.ExtendedProperties.Properties();

            // Add the ThumbnailPart part in the new word processing document.
            wordDoc.AddNewPart<ThumbnailPart>("image/jpeg", "rId6");

            wordDoc.Dispose();
        }
    }

}
