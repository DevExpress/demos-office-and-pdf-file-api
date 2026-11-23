#region Code
using DevExpress.Spreadsheet;
using DevExpress.Office;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.DocumentProtection {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelSecureDocumentLoadingController : SpreadsheetControllerBase {
        [HttpPost("secure-document-loading/load")]
        public async Task<IActionResult> Load(IFormFile file, [FromForm] ExcelSecureDocumentLoadingModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process document and remove unsafe content
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
            using var workbook = new Workbook();

            var workbookSecurityLoadingOptions = workbook.Options.SecurityLoadingOptions;
            workbookSecurityLoadingOptions.RemoveActiveXContent = securityLoadingOptions.RemoveActiveXContent;
            workbookSecurityLoadingOptions.RemoveCustomXMLParts = securityLoadingOptions.RemoveCustomXMLParts;
            workbookSecurityLoadingOptions.RemoveExternalConnections = securityLoadingOptions.RemoveExternalConnections;
            workbookSecurityLoadingOptions.RemoveExternalWorkbooks = securityLoadingOptions.RemoveExternalWorkbooks;
            workbookSecurityLoadingOptions.RemoveExternalImages = securityLoadingOptions.RemoveExternalImages;
            workbookSecurityLoadingOptions.RemoveMacros = securityLoadingOptions.RemoveMacros;
            workbookSecurityLoadingOptions.RemoveOleObjects = securityLoadingOptions.RemoveOleObjects;
            workbookSecurityLoadingOptions.RemovePivotCaches = securityLoadingOptions.RemovePivotCaches;
            workbookSecurityLoadingOptions.RemoveRestrictedFormulas = securityLoadingOptions.RemoveRestrictedFormulas;
            workbookSecurityLoadingOptions.RestrictedHyperlinkRemovalMode = securityLoadingOptions.RestrictedHyperlinkRemovalMode;

            var workbookSecurityLoadingLimits = workbook.Options.SecurityLoadingLimits;
            workbookSecurityLoadingLimits.MaxChartCount = securityLoadingLimits.MaxChartCount;
            workbookSecurityLoadingLimits.MaxFileSize = securityLoadingLimits.MaxFileSize;
            workbookSecurityLoadingLimits.MaxSheetCellCount = securityLoadingLimits.MaxSheetCellCount;
            workbookSecurityLoadingLimits.MaxSheetRowCount = securityLoadingLimits.MaxSheetRowCount;
            workbookSecurityLoadingLimits.MaxWorksheetCount = securityLoadingLimits.MaxWorksheetCount;
            workbookSecurityLoadingLimits.MaxXmlElementCount = securityLoadingLimits.MaxXmlElementCount;
            workbookSecurityLoadingLimits.MaxXmlElementDepth = securityLoadingLimits.MaxXmlElementDepth;

            workbook.SecurityLoadingLimitExceeded += (o, ea) => {
                // Use the SecurityLoadingLimitExceeded event to log information about the security loading limits
                // Set the "Handled" parameter to true to load a document even if it exceeds the specified limits
            };

            workbook.SecurityLoadingOptionsViolation += (o, ea) => {
                // Use the SecurityLoadingOptionsViolation event to log information about the security loading options violation
                // Set the "Handled" parameter to true to ignore the violation and load a document even if it contains unsafe content
            };

            loadErrorMessage = "";
            try {
                workbook.LoadDocument(inputStream);

                // Save document
                var outputStream = new MemoryStream();

                workbook.SaveDocument(outputStream, outputFormat);

                outputStream.Position = 0;
                return outputStream;
            }
            catch(SecurityLoadingLimitExceededException limitExceededException) {
                loadErrorMessage = limitExceededException.Message;
                return new MemoryStream();
            }
        }
        #endregion Code
    }
}
