using BlazorDemo.OfficeFileApi.Models.WordProcessing.SearchExtract;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using System.Text;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.SearchExtract {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordMetadataController : WordControllerBase {
        [HttpPost("metadata/read")]
        public async Task<IActionResult> ReadMetadata(IFormFile file, [FromForm] WordMetadataModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Generate plain text metadata output
            var output = new StringBuilder();
            var metadataText = ReadDocumentPropertiest(inputStream, model.ReadDocumentProperties, model.ReadCustomProperties);
            output.AppendLine(metadataText);
            using var outputStream = new MemoryStream();
            using (StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(output.ToString());
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
        }

        #region Code
        string ReadDocumentPropertiest(Stream inputStream, bool readDocumentProperties, bool readCustomProperties) {
            var output = new StringBuilder();

            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(inputStream);

            var document = wordProcessor.Document;

            if (readDocumentProperties) {
                // Core Properties
                var coreProperties = document.DocumentProperties;
                output.AppendLine("CORE DOCUMENT PROPERTIES:");
                output.AppendLine(new string('-', 30));
                output.AppendLine($"Title: {coreProperties.Title ?? ""}");
                output.AppendLine($"Author: {coreProperties.Author ?? ""}");
                output.AppendLine($"Subject: {coreProperties.Subject ?? ""}");
                output.AppendLine($"Keywords: {coreProperties.Keywords ?? ""}");
                output.AppendLine($"Category: {coreProperties.Category ?? ""}");
                output.AppendLine($"Company: {coreProperties.Company ?? ""}");
                output.AppendLine($"Manager: {coreProperties.Manager ?? ""}");
                output.AppendLine($"Created Date: {coreProperties.Created.ToString("yyyy-MM-dd HH:mm:ss")}");
                output.AppendLine($"Last Printed: {coreProperties.LastPrinted.ToString("yyyy-MM-dd HH:mm:ss")}");
                output.AppendLine();
            }

            if (readCustomProperties) {
                var customProperties = document.CustomProperties;

                output.AppendLine("CUSTOM PROPERTIES:");
                output.AppendLine(new string('-', 30));

                // Check if there are any custom properties
                var hasCustomProperties = false;
                var customPropsList = new List<string>();

                foreach (string propName in customProperties.Names) {
                    hasCustomProperties = true;
                    var propValue = customProperties[propName];
                    var valueStr = propValue?.ToString() ?? "";
                    var typeStr = propValue?.GetType().Name ?? "String";
                    customPropsList.Add($"{propName}: {valueStr} ({typeStr})");
                }
                if (!hasCustomProperties)
                    output.AppendLine("No custom properties found.");
                else
                    foreach (var customProp in customPropsList)
                        output.AppendLine(customProp);
                output.AppendLine();
            }

            return output.ToString();
        }
        #endregion Code
    }
}
