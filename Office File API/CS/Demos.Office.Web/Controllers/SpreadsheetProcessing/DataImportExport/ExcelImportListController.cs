#region Code
using DevExpress.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
#endregion Code

using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.DataImportExport;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.DataImportExport {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelImportListController : SpreadsheetControllerBase {
        [HttpPost("import-export/importlist")]
        public async Task<IActionResult> Import([FromForm] ExcelImportListModel model) {
            DocumentFormat outputFormatConverted = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Import Users
            var documentStream = await ImportList(model.Users, model.SelectedFields.ToArray(), outputFormatConverted);

            // Create Output File
            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        async Task<Stream> ImportList(IList<UserModel> users, string[] selectedFields, DocumentFormat outputFormat) {

            using Workbook workbook = new Workbook();

            // Prepare first worksheet
            var firstWorksheet = workbook.Worksheets.First();
            firstWorksheet.Name = "Users";
            firstWorksheet["A1"].Value = "Import data from a list of users:";

            // Set headers
            if(selectedFields == null || !selectedFields.Any()) {
                selectedFields = new string[3];
                selectedFields[0] = nameof(UserModel.FirstName);
                selectedFields[1] = nameof(UserModel.LastName);
                selectedFields[2] = nameof(UserModel.BirthDate);
            }
            for(int i = 0; i < selectedFields.Length; i++) {
                firstWorksheet[0, i + 1].Value = selectedFields[i];
            }
            // Make header text bold
            var formatting = firstWorksheet["B1:D1"].BeginUpdateFormatting();
            formatting.Font.Bold = true;
            firstWorksheet["B1:D1"].EndUpdateFormatting(formatting);

            // Import values from specific data source fields
            // Data starts with the B1 cell.
            var importOptions = new DataSourceImportOptions() {
                PropertyNames = selectedFields
            };
            firstWorksheet.Import(users, 1, 1, importOptions);

            // AutoFit columns
            firstWorksheet.Columns.AutoFit(0, 4);

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the document
            workbook.SaveDocument(outputStream, outputFormat);
            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code
    }
}
