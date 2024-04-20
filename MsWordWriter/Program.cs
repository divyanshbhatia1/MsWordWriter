using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace MsWordWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new WordprocessingDocument
            string fileName = @"C:\Users\bhati\Downloads\Output\document.docx"; // Specify your desired file path

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                // Add a main document part
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph paragraph = body.AppendChild(new Paragraph());
                Run run = paragraph.AppendChild(new Run());
                run.AppendChild(new Text("Hello, this is a test document."));

                // Save changes
                mainPart.Document.Save();
            }

            Console.WriteLine("Document created successfully!");
        }
    }

}
