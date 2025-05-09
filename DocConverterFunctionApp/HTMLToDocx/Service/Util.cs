using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using Microsoft.Extensions.Logging;

namespace DocConverterFunctionApp
{
    public static class Util
    {
        public static string ValidateHtml(string htmlContent, ILogger log)
        {
            try
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(htmlContent);

                if (htmlDoc.ParseErrors != null && htmlDoc.ParseErrors.Any())
                {
                    log.LogWarning("HTML content contains parse errors.");
                }

                return htmlDoc.DocumentNode.OuterHtml;
            }
            catch (Exception ex)
            {
                log.LogError($"Error validating HTML content: {ex.Message}");
                throw new InvalidOperationException("Invalid HTML content.", ex);
            }
        }

        public static string SimplifyHtml(string htmlContent, ILogger log)
        {
            try
            {
                var htmlDoc = new HtmlAgilityPack.HtmlDocument();
                htmlDoc.LoadHtml(htmlContent);

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

        public static string FindCommonBaseDirectory(List<string> resourceFiles)
        {
            if (resourceFiles == null || resourceFiles.Count == 0)
            {
                return string.Empty;
            }

            var normalizedPaths = resourceFiles.Select(Path.GetFullPath).ToList();
            var commonPathParts = normalizedPaths[0].Split(Path.DirectorySeparatorChar);

            foreach (var filePath in normalizedPaths)
            {
                var pathParts = filePath.Split(Path.DirectorySeparatorChar);

                for (int i = 0; i < commonPathParts.Length; i++)
                {
                    if (i >= pathParts.Length || !string.Equals(commonPathParts[i], pathParts[i], StringComparison.OrdinalIgnoreCase))
                    {
                        commonPathParts = commonPathParts.Take(i).ToArray();
                        break;
                    }
                }
            }

            // Combine the common parts and ensure the leading '/' is preserved
            var commonPath = Path.Combine(commonPathParts);
            if (Path.IsPathRooted(normalizedPaths[0]))
            {
                commonPath = Path.DirectorySeparatorChar + commonPath;
            }

            return commonPath;
        }
    }
}