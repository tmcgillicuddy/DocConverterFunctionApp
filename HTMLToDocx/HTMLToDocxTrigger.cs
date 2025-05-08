using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;

namespace DocConverterFunctionApp
{
    public static class HTMLToDocxTrigger
    {
        [FunctionName("HTMLToDocxTrigger")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req,
            ILogger log)
        {
            // Check if the request contains a file
            if (!req.HasFormContentType || req.Form.Files.Count == 0)
            {
                return new BadRequestObjectResult("Please upload a ZIP file containing an HTML file and resources.");
            }

            // Get the uploaded file
            var file = req.Form.Files[0];

            // Ensure the file has a ZIP extension
            if (Path.GetExtension(file.FileName).ToLower() != ".zip")
            {
                return new BadRequestObjectResult("Only ZIP files are supported.");
            }

            // Create a temporary directory to extract the ZIP file
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);

            try
            {
                // Extract the ZIP file
                using (var zipStream = file.OpenReadStream())
                using (var archive = new ZipArchive(zipStream))
                {
                    archive.ExtractToDirectory(tempDir);
                }

                // Find the HTML file in the extracted contents
                var htmlFilePath = Directory.GetFiles(tempDir, "*.html", SearchOption.AllDirectories).FirstOrDefault();
                if (htmlFilePath == null)
                {
                    log.LogError("No HTML file found in the ZIP archive.");
                    return new BadRequestObjectResult("The ZIP file must contain an HTML file.");
                }

                // Read the content of the HTML file
                string htmlContent = await File.ReadAllTextAsync(htmlFilePath);

                // Log the received HTML file
                log.LogInformation($"Received HTML file: {Path.GetFileName(htmlFilePath)}");

                // Find any images or resources in the extracted contents
                var resourceFiles = Directory.GetFiles(tempDir, "*.*", SearchOption.AllDirectories)
                    .Where(f => f != htmlFilePath)
                    .ToList();

                if (!resourceFiles.Any())
                {
                    log.LogWarning("No additional resources (e.g., images) found in the ZIP archive.");
                }

                // Convert HTML to DOCX with resources
                HtmlToDocxConverter _converter = new HtmlToDocxConverter();
                var docxBytes = _converter.ConvertHtmlToDocx(htmlContent, resourceFiles, log);

                // Return the DOCX file as a response
                return new FileContentResult(docxBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                {
                    FileDownloadName = $"{Path.GetFileNameWithoutExtension(file.FileName)}.docx"
                };
            }
            catch (Exception ex)
            {
                log.LogError($"Error processing the ZIP file: {ex.Message}");
                return new StatusCodeResult(StatusCodes.Status500InternalServerError);
            }
            finally
            {
                // Clean up the temporary directory
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                }
            }
        }
    }
}
