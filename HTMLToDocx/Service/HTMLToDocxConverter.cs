using System.IO;
using Spire.Doc;

namespace DocConverterFunctionApp
{
    public class HtmlToDocxConverter
    {
        public byte[] ConvertHtmlToDocx(string htmlContent)
        {
            using (var memoryStream = new MemoryStream())
            {
                // Create a new document
                var document = new Document();

                // Add a section and insert the HTML content
                var section = document.AddSection();
                var paragraph = section.AddParagraph();
                paragraph.AppendHTML(htmlContent);

                // Save the document to the memory stream
                document.SaveToStream(memoryStream, FileFormat.Docx);

                // Return the DOCX file as a byte array
                return memoryStream.ToArray();
            }
        }
    }
}