using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.SearchExtract;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Word;
using System.Text;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.SearchExtract {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordMetadataController : WordControllerBase {
        [HttpPost("metadata/read")]
        public async Task<IActionResult> ReadMetadata(IFormFile file, [FromForm] WordMetadataModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Generate plain text metadata output
            var output = new StringBuilder();
            var metadataText = ReadDocumentProperties(inputStream, model.ReadDocumentProperties, model.ReadCustomProperties);
            output.AppendLine(metadataText);
            using var outputStream = new MemoryStream();
            using(var writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(output.ToString());
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
        }

        #region Code
        string ReadDocumentProperties(Stream inputStream, bool readDocumentProperties, bool readCustomProperties) {
            var output = new StringBuilder();

            using var document = new WordDocument(inputStream);
            var documentProperties = document.DocumentProperties;

            if(readDocumentProperties) {
                // Core Properties
                output.AppendLine("CORE DOCUMENT PROPERTIES:");
                output.AppendLine(new string('-', 30));
                output.AppendLine($"Title: {documentProperties.Title ?? ""}");
                output.AppendLine($"Author: {documentProperties.Author ?? ""}");
                output.AppendLine($"Subject: {documentProperties.Subject ?? ""}");
                output.AppendLine($"Keywords: {documentProperties.Keywords ?? ""}");
                output.AppendLine($"Category: {documentProperties.Category ?? ""}");
                output.AppendLine($"Company: {documentProperties.Company ?? ""}");
                output.AppendLine($"Manager: {documentProperties.Manager ?? ""}");
                output.AppendLine($"Created Date: {documentProperties.Created.ToString("yyyy-MM-dd HH:mm:ss")}");
                output.AppendLine($"Last Printed: {documentProperties.Printed.ToString("yyyy-MM-dd HH:mm:ss")}");
                output.AppendLine();
            }

            if(readCustomProperties) {
                var customProperties = documentProperties.CustomProperties;

                output.AppendLine("CUSTOM PROPERTIES:");
                output.AppendLine(new string('-', 30));

                // Check if there are any custom properties
                if(customProperties.Count == 0)
                    output.AppendLine("No custom properties found.");
                else
                    foreach(var customProperty in customProperties) {
                        var propertyValue = customProperty.Value.Value;
                        var valueText = propertyValue?.ToString() ?? "";
                        var typeName = propertyValue?.GetType().Name ?? "String";
                        output.AppendLine($"{customProperty.Key}: {valueText} ({typeName})");
                    }

                output.AppendLine();
            }

            return output.ToString();
        }
        #endregion Code
    }
}