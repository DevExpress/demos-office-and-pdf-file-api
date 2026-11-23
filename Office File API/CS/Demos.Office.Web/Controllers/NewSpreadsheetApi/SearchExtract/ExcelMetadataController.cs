#region Code
using DevExpress.Docs.Spreadsheet;
using System.Text;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.SearchExtract;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.SearchExtract {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelMetadataController : SpreadsheetControllerBase {
        [HttpPost("metadata/read")]
        public async Task<IActionResult> Read(IFormFile file, [FromForm] ExcelMetadataModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Seek(0, SeekOrigin.Begin);

            string metadataText = ReadDocumentProperties(inputStream, model.ReadDocumentProperties, model.ReadCustomProperties);
            using var outputStream = new MemoryStream();
            using(StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(metadataText);
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
        }
        #region Code
        string ReadDocumentProperties(Stream inputStream, bool readDocumentProperties, bool readCustomProperties) {
            var output = new StringBuilder();
            using var workbook = new Workbook(inputStream);
            var documentProperties = workbook.DocumentProperties;

            if(readDocumentProperties) {
                output.AppendLine("MAIN DOCUMENT PROPERTIES:");
                output.AppendLine(new string('-', 30));
                output.AppendLine($"Application: {documentProperties.Application ?? ""}");
                output.AppendLine($"Author: {documentProperties.Author ?? ""}");
                output.AppendLine($"Category: {documentProperties.Category ?? ""}");
                output.AppendLine($"Company: {documentProperties.Company ?? ""}");
                output.AppendLine($"Content Status: {documentProperties.ContentStatus ?? ""}");
                output.AppendLine($"Created Date: {documentProperties.Created.ToString("yyyy-MM-dd HH:mm:ss")}");
                output.AppendLine($"Description: {documentProperties.Description ?? ""}");
                output.AppendLine($"Document Revision: {documentProperties.DocumentRevision ?? ""}");
                output.AppendLine($"Document Version: {documentProperties.DocumentVersion ?? ""}");
                output.AppendLine($"Keywords: {documentProperties.Keywords ?? ""}");
                output.AppendLine($"Last Modified By: {documentProperties.LastModifiedBy ?? ""}");
                output.AppendLine($"Manager: {documentProperties.Manager ?? ""}");
                output.AppendLine($"Modified Date: {documentProperties.Modified.ToString("yyyy-MM-dd HH:mm:ss")}");
                output.AppendLine($"Printed Date: {documentProperties.Printed.ToString("yyyy-MM-dd HH:mm:ss")}");
                output.AppendLine($"Security: {documentProperties.Security}");
                output.AppendLine($"Subject: {documentProperties.Subject ?? ""}");
                output.AppendLine($"Title: {documentProperties.Title ?? ""}");
                output.AppendLine($"Version: {documentProperties.Version ?? ""}");
                output.AppendLine();
            }

            if(readCustomProperties) {
                var customProperties = documentProperties.CustomProperties;
                output.AppendLine("CUSTOM PROPERTIES:");
                output.AppendLine(new string('-', 30));

                if(customProperties.Count == 0)
                    output.AppendLine("No custom properties found.");
                else
                    foreach(var customProperty in customProperties)
                        output.AppendLine($"{customProperty.Key}: {customProperty.Value.Value} ({customProperty.Value.Type})");
                output.AppendLine();
            }

            return output.ToString();
        }
        #endregion Code
    }
}
