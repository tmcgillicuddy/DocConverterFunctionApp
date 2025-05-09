using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
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
                // Validate and clean the HTML content
                htmlContent = ValidateHtml(htmlContent, log);
                log.LogInformation("HTML content validated.");

                // Simplify the HTML content
                log.LogInformation("Simplifying HTML content.");
                htmlContent = SimplifyHtml(htmlContent, log);
                log.LogInformation("HTML content simplified.");

                // Process and embed resources
                htmlContent = EmbedResources(htmlContent, resourceFiles, section, log);

                log.LogInformation("Processing HTML content for embedded resources.");

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

            // Ensure the base directory is absolute
            var baseDirectory = Path.GetFullPath(FindCommonBaseDirectory(resourceFiles));
            log.LogInformation($"Base directory for resources: {baseDirectory}");

            // Log all resource files
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

            // Handle <img> tags
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

                        // Add the image to the document
                        var image = section.AddParagraph().AppendPicture(File.ReadAllBytes(resourceFile));

                        // Replace the <img> tag with a placeholder or remove it
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

        private string ValidateHtml(string htmlContent, ILogger log)
        {
            try
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(htmlContent);

                // Check for errors
                if (htmlDoc.ParseErrors != null && htmlDoc.ParseErrors.Any())
                {
                    log.LogWarning("HTML content contains parse errors.");
                }

                // Return the cleaned HTML
                return htmlDoc.DocumentNode.OuterHtml;
            }
            catch (Exception ex)
            {
                log.LogError($"Error validating HTML content: {ex.Message}");
                throw new InvalidOperationException("Invalid HTML content.", ex);
            }
        }

        private string SimplifyHtml(string htmlContent, ILogger log)
        {
            try
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(htmlContent);

                // Remove <script> and <style> tags
                var scriptNodes = htmlDoc.DocumentNode.SelectNodes("//script");
                var styleNodes = htmlDoc.DocumentNode.SelectNodes("//style");

                scriptNodes?.ToList().ForEach(node => node.Remove());
                styleNodes?.ToList().ForEach(node => node.Remove());

                log.LogInformation("Simplified HTML content by removing <script> and <style> tags.");
                return htmlDoc.DocumentNode.OuterHtml;
            }
            catch (Exception ex)
            {
                log.LogError($"Error simplifying HTML content: {ex.Message}");
                throw new InvalidOperationException("Failed to simplify HTML content.", ex);
            }
        }

        private string FindCommonBaseDirectory(List<string> resourceFiles)
        {
            if (resourceFiles == null || resourceFiles.Count == 0)
            {
                return string.Empty;
            }

            // Normalize all paths to absolute paths
            var normalizedPaths = resourceFiles.Select(Path.GetFullPath).ToList();

            // Split the first path into its directory components
            var commonPathParts = normalizedPaths[0].Split(Path.DirectorySeparatorChar);

            foreach (var filePath in normalizedPaths)
            {
                var pathParts = filePath.Split(Path.DirectorySeparatorChar);

                // Find the common prefix between the current common path and the current file path
                for (int i = 0; i < commonPathParts.Length; i++)
                {
                    if (i >= pathParts.Length || !string.Equals(commonPathParts[i], pathParts[i], StringComparison.OrdinalIgnoreCase))
                    {
                        // Trim the common path to the shared prefix
                        commonPathParts = commonPathParts.Take(i).ToArray();
                        break;
                    }
                }
            }

            // Combine the common path parts back into a single directory path
            return Path.Combine(commonPathParts);
        }
    }
}