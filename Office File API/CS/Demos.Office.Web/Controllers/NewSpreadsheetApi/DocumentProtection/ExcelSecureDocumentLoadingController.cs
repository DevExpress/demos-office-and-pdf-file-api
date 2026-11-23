#region Code
using DevExpress.Docs.Spreadsheet;
using DevExpress.Docs.Office;
using DevExpress.Docs;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.DocumentProtection {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelSecureDocumentLoadingController : SpreadsheetControllerBase {
        [HttpPost("secure-document-loading/load")]
        public async Task<IActionResult> Load(IFormFile file, [FromForm] ExcelSecureDocumentLoadingModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            string loadErrorMessage = string.Empty;
            var outputStream = TryLoadDocument(
                inputStream,
                new WorkbookSecurityLoadingLimits() {
                    MaxChartCount = model.MaxChartCount,
                    MaxFileSize = 5 * 1024 * 1024,
                    MaxSheetCellCount = model.MaxSheetCellCount,
                    MaxSheetRowCount = model.MaxSheetRowCount,
                    MaxWorksheetCount = model.MaxWorksheetCount,
                    MaxXmlElementCount = 1000,
                    MaxXmlElementDepth = 10
                },
                new WorkbookSecurityLoadingOptions() {
                    RemoveActiveXContent = model.RemoveActiveXContent,
                    RemoveCustomXMLParts = model.RemoveCustomXMLParts,
                    RemoveExternalConnections = model.RemoveExternalConnections,
                    RemoveExternalWorkbooks = model.RemoveExternalWorkbooks,
                    RemoveExternalImages = model.RemoveExternalImages,
                    RemoveMacros = model.RemoveMacros,
                    RemoveOleObjects = model.RemoveOleObjects,
                    RemovePivotCaches = model.RemovePivotCaches,
                    RemoveRestrictedFormulas = model.RemoveRestrictedFormulas,
                    RestrictedHyperlinkRemovalMode = model.RestrictedHyperlinkRemovalMode
                },
                out loadErrorMessage,
                outputFormat);
            if(string.IsNullOrEmpty(loadErrorMessage)) {
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
            else {
                return BadRequest(loadErrorMessage);
            }
        }
        #region Code
        Stream TryLoadDocument(Stream inputStream, WorkbookSecurityLoadingLimits securityLoadingLimits,
            WorkbookSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
            var loadOptions = new LoadOptions() {
                SecurityLoadingHandler = new WorkbookSecurityLoadingHandler(),
                SecurityLoadingOptions = securityLoadingOptions,
                SecurityLoadingLimits = securityLoadingLimits
            };

            loadErrorMessage = "";
            try {
                using var workbook = new Workbook(inputStream, loadOptions);

                var outputStream = new MemoryStream();

                workbook.Save(outputStream, outputFormat);

                outputStream.Position = 0;
                return outputStream;
            }
            catch(SecurityLoadingLimitExceededException limitExceededException) {
                loadErrorMessage = limitExceededException.Message;
                return new MemoryStream();
            }
        }

        class WorkbookSecurityLoadingHandler : ISecurityLoadingHandler {
            public SecurityLimitAction HandleLimitExceeded(SecurityLimitExceededParams parameters) {
                return SecurityLimitAction.Throw;
            }

            public SecurityViolationAction HandleOptionsViolation(SecurityOptionsViolationParams parameters) {
                return SecurityViolationAction.Remove;
            }
        }
        #endregion Code
    }
}
