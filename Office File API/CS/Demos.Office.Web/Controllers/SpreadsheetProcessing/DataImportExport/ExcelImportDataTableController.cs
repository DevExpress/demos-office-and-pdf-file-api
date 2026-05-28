#region Code
using DevExpress.Drawing;
using DevExpress.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using System.Data;
#endregion Code

using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.DataImportExport;
using BlazorDemo.OfficeFileApi.Services;


namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.DataImportExport {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelImportDataTableController : SpreadsheetControllerBase {
        [HttpPost("import-export/importdatatable")]
        public async Task<IActionResult> Import([FromForm] ExcelImportDataTableModel model) {
            DocumentFormat outputFormatConverted = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Import Data Table
            var documentStream = await ImportDataTable(model.EmployeesTable, outputFormatConverted);

            // Create Output File
            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        async Task<Stream> ImportDataTable(DataTable sourceTable, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Prepare first worksheet
            var firstWorksheet = workbook.Worksheets.First();
            firstWorksheet.Name = "Employees";
            firstWorksheet["A1"].Value = "Import data from an Employees data table:";

            // Make header text bold
            var formatting = firstWorksheet["B1:G1"].BeginUpdateFormatting();
            formatting.Font.Bold = true;
            firstWorksheet["B1:G1"].EndUpdateFormatting(formatting);

            // Specify import options
            var importOptions = new DataImportOptions() {
                // Specify converter to convert Base64 images to Excel images
                Converter = new DataValueConverter()                
            };
            
            // Import data from the data table into the worksheet and insert it, starting with cell B1.
            firstWorksheet.Import(sourceTable, true, 0, 1, importOptions);

            // AutoFit columns
            firstWorksheet.Columns.AutoFit(0, 6);

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the document
            await workbook.SaveDocumentAsync(outputStream, outputFormat);
            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }

        // Use this converter to convert Base64 images to Excel images
        class DataValueConverter : IDataValueConverter {
            public bool TryConvert(object value, int index, out CellValue result) {
                if(value is string strValue) {
                    try {
                        result = DXImage.FromBase64String(strValue);
                    } catch {
                        int str2int = 0;
                        if(Int32.TryParse(strValue, out str2int))
                            result = str2int;
                        else
                            result = strValue;
                    }
                    return true;
                }
                result = DevExpress.Spreadsheet.CellValue.TryCreateFromObject(value);
                return true;
            }
        }
        #endregion Code
    }
}
