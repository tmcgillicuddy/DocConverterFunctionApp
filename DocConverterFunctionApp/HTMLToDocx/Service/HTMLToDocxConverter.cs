using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Spire.Doc;
using Microsoft.Extensions.Logging;
using HtmlAgilityPack;

namespace DocConverterFunctionApp
{
    public class HtmlToDocxConverter
    {
        public byte[] ConvertHtmlToDocx(string htmlContent, List<string> resourceFiles, ILogger log)
        {
            if (string.IsNullOrEmpty(htmlContent))
            {
                throw new ArgumentException("HTML content cannot be null or empty.");
            }

            if (resourceFiles == null)
            {
                log.LogWarning("Resource files list is null. Initializing to an empty list.");
                resourceFiles = new List<string>();
            }
            else
            {
                log.LogInformation($"Resource files count: {resourceFiles.Count}");
            }

            using (var memoryStream = new MemoryStream())
            {
                // Create a new document
                var document = new Document();

                // Add a section and insert the HTML content
                var section = document.AddSection();

                log.LogInformation("Validating HTML content.");
                htmlContent = Util.ValidateHtml(htmlContent, log);
                log.LogInformation("HTML content validated.");

                log.LogInformation("Simplifying HTML content.");
                htmlContent = Util.SimplifyHtml(htmlContent, log);
                log.LogInformation("HTML content simplified.");

                log.LogInformation("Processing HTML content for embedded resources.");
                htmlContent = EmbedResources(htmlContent, resourceFiles, section, log);

                // Add the processed HTML content to the document
                var paragraph = section.AddParagraph();

                log.LogInformation("Adding HTML content to the document.");
                try
                {
                    paragraph.AppendHTML(htmlContent);
                    log.LogInformation("HTML content added to the document.");
                }
                catch (Exception ex)
                {
                    log.LogError($"Error in AppendHTML: {ex.Message}");
                    throw new InvalidOperationException("Failed to append HTML content to the document.", ex);
                }

                // Save the document to the memory stream
                document.SaveToStream(memoryStream, FileFormat.Docx);
                log.LogInformation("Document saved to memory stream.");

                // Return the DOCX file as a byte array
                return memoryStream.ToArray();
            }
        }

        private string EmbedResources(string htmlContent, List<string> resourceFiles, Section section, ILogger log)
        {
            var htmlDoc = new HtmlAgilityPack.HtmlDocument();
            htmlDoc.LoadHtml(htmlContent);

            var baseDirectory = Path.GetFullPath(Util.FindCommonBaseDirectory(resourceFiles));
            log.LogInformation($"Base directory for resources: {baseDirectory}");

            log.LogInformation("Unzipping resource files...");
            foreach (var resourceFile in resourceFiles)
            {
                log.LogInformation($"Resource file path: {resourceFile}");
                if (File.Exists(resourceFile))
                {
                    log.LogInformation($"File exists: {resourceFile}");
                }
                else
                {
                    log.LogWarning($"File does not exist: {resourceFile}");
                }
            }

            var imgNodes = htmlDoc.DocumentNode.SelectNodes("//img[@src]");
            if (imgNodes != null)
            {
                foreach (var imgNode in imgNodes)
                {
                    var src = imgNode.GetAttributeValue("src", null);
                    if (string.IsNullOrEmpty(src))
                    {
                        log.LogWarning("Skipping <img> tag with missing 'src' attribute.");
                        continue;
                    }

                    var resourceFile = resourceFiles.FirstOrDefault(r =>
                        string.Equals(Path.GetFullPath(r), Path.GetFullPath(Path.Combine(baseDirectory, src)), StringComparison.OrdinalIgnoreCase));

                    if (resourceFile != null && File.Exists(resourceFile))
                    {
                        log.LogInformation($"Embedding image: {resourceFile}");
                        var image = section.AddParagraph().AppendPicture(File.ReadAllBytes(resourceFile));
                        imgNode.ParentNode.ReplaceChild(HtmlNode.CreateNode($"[Embedded Image: {src}]"), imgNode);
                    }
                    else
                    {
                        var expectedPath = Path.GetFullPath(Path.Combine(baseDirectory, src));
                        log.LogWarning($"Resource file for <img> tag not found: {src}. Expected path: {expectedPath}");
                        imgNode.Remove();
                    }
                }
            }

            log.LogInformation("Processed linked resources in HTML content.");
            return htmlDoc.DocumentNode.OuterHtml;
        }
    }
}