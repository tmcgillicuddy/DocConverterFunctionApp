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

                    var resourceFile = resourceFiles.FirstOrDefault(r => r.EndsWith(src, StringComparison.OrdinalIgnoreCase));
                    if (resourceFile != null && File.Exists(resourceFile))
                    {
                        log.LogInformation($"Embedding image: {src}");

                        // Add the image to the document
                        var image = section.AddParagraph().AppendPicture(File.ReadAllBytes(resourceFile));

                        // Replace the <img> tag with a placeholder or remove it
                        imgNode.ParentNode.ReplaceChild(HtmlNode.CreateNode($"[Embedded Image: {src}]"), imgNode);
                    }
                    else
                    {
                        log.LogWarning($"Resource file for <img> tag not found: {src}. Removing the tag.");
                        imgNode.Remove();
                    }
                }
            }

            // Handle <link> tags for CSS
            var linkNodes = htmlDoc.DocumentNode.SelectNodes("//link[@rel='stylesheet' and @href]");
            if (linkNodes != null)
            {
                foreach (var linkNode in linkNodes)
                {
                    var href = linkNode.GetAttributeValue("href", null);
                    if (string.IsNullOrEmpty(href))
                    {
                        log.LogWarning("Skipping <link> tag with missing 'href' attribute.");
                        continue;
                    }

                    var resourceFile = resourceFiles.FirstOrDefault(r => r.EndsWith(href, StringComparison.OrdinalIgnoreCase));
                    if (resourceFile != null && File.Exists(resourceFile))
                    {
                        log.LogInformation($"Inlining CSS: {href}");
                        var cssContent = File.ReadAllText(resourceFile);

                        // Inline the CSS into a <style> tag
                        var styleNode = HtmlNode.CreateNode($"<style>{cssContent}</style>");
                        linkNode.ParentNode.ReplaceChild(styleNode, linkNode);
                    }
                    else
                    {
                        log.LogWarning($"Resource file for <link> tag not found: {href}. Removing the tag.");
                        linkNode.Remove();
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
    }
}