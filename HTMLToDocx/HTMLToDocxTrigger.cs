using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

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
                return new BadRequestObjectResult("Please upload an HTML file.");
            }
            // Get the uploaded file
            var file = req.Form.Files[0];

            // Ensure the file has an HTML extension
            if (Path.GetExtension(file.FileName).ToLower() != ".html")
            {
                return new BadRequestObjectResult("Only HTML files are supported.");
            }
            // Read the content of the HTML file
            string htmlContent;
            using (var reader = new StreamReader(file.OpenReadStream()))
            {
                htmlContent = await reader.ReadToEndAsync();
            }

            log.LogInformation($"Received HTML file: {file.FileName}");

            HtmlToDocxConverter _converter = new HtmlToDocxConverter();

            //Convert HTML to DOCX
            // Use the HtmlToDocxConverter service to convert the HTML to DOCX
            var docxBytes = _converter.ConvertHtmlToDocx(htmlContent);


            // Test logic
            log.LogInformation("C# HTTP trigger function processed a request.");

            // Return the DOCX file as a response
            return new FileContentResult(docxBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                FileDownloadName = $"{Path.GetFileNameWithoutExtension(file.FileName)}.docx"
            };
        }
    }
}
