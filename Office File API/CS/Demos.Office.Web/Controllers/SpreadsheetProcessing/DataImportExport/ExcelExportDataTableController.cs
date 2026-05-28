#region Code
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Export;
using System.Data;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.DataImportExport;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.DataImportExport {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelExportDataTableController : SpreadsheetControllerBase {
        [HttpPost("import-export/exportdatatable")]
        public async Task<IActionResult> Export([FromForm] IFormFile file, [FromForm] ExcelExportDataTableModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Export data table
            DataTable exportedTable = ExportDataTable(inputStream);
            return Ok(ConvertToDTO(exportedTable));
        }
        #region Code
        DataTable ExportDataTable(Stream inputStream) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Get data table range in the document
            Worksheet worksheet = workbook.Worksheets[0];
            CellRange range = worksheet.Tables[0].Range;

            // Create data table schema
            DataTable dataTable = worksheet.CreateDataTable(range, true);

            // Change the data type of the "As Of" column to text
            dataTable.Columns["As Of"].DataType = typeof(string);

            // Create exporter object
            DataTableExporter exporter = worksheet.CreateDataTableExporter(range, dataTable, true);
            exporter.CellValueConversionError += Exporter_CellValueConversionError;

            // Create cell value converter
            EmptyCellValueConverter myconverter = new EmptyCellValueConverter();
            exporter.Options.CustomConverters.Add("As Of", myconverter);

            // Set the export value for empty cell
            myconverter.EmptyCellValue = "N/A";
            exporter.Options.ConvertEmptyCells = true;
            exporter.Options.DefaultCellValueToColumnTypeConverter.SkipErrorValues = false;

            //Export data table
            exporter.Export();

            return dataTable;
        }

        void Exporter_CellValueConversionError(object sender, CellValueConversionErrorEventArgs e) {
            e.DataTableValue = null;
            e.Action = DataTableExporterAction.Continue;
        }

        class EmptyCellValueConverter : ICellValueToColumnTypeConverter {
            public bool SkipErrorValues { get; set; }
            public CellValue EmptyCellValue { get; set; }

            public ConversionResult Convert(
                Cell readOnlyCell, CellValue cellValue,
                Type dataColumnType, out object result) {
                result = DBNull.Value;

                ConversionResult converted = ConversionResult.Success;
                if(cellValue.IsEmpty) {
                    result = EmptyCellValue;
                    return converted;
                }
                if(cellValue.IsError) {
                    // You can return an error, subsequently the exporter throws an exception if the CellValueConversionError event is unhandled.
                    //return SkipErrorValues ? ConversionResult.Success : ConversionResult.Error;
                    result = "N/A";
                    return ConversionResult.Success;
                }
                result = String.Format("{0:MMMM-yyyy}", cellValue.DateTimeValue);
                return converted;
            }
        }
        #endregion Code
        List<TradingPartnerDTO> ConvertToDTO(DataTable exportedTable) {
            List<TradingPartnerDTO> dTOs = new List<TradingPartnerDTO>();
            foreach(DataRow row in exportedTable.Rows) {
                dTOs.Add(new TradingPartnerDTO() {
                    Country = (string)row["Country"],
                    Exports = (double)row["Exports"],
                    Imports = (double)row["Imports"],
                    Balance = (double)row["Balance"],
                    Exports1YChg = (double)row["Exports 1Y Chg"],
                    Imports1YChg = (double)row["Imports 1Y Chg"],
                    Balance1YChg = (double)row["Balance 1Y Chg"],
                    AsOf = (string)row["As Of"],
                });
            }
            return dTOs;
        }
        class TradingPartnerDTO {
            public string Country { get; set; }
            public double Exports { get; set; }
            public double Imports { get; set; }
            public double Balance { get; set; }
            public double Exports1YChg { get; set; }
            public double Imports1YChg { get; set; }
            public double Balance1YChg { get; set; }
            public string AsOf { get; set; }
        }
    }
}
