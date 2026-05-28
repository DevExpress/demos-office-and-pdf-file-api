#region Code
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Formulas;
using DevExpress.XtraSpreadsheet.Services;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.Calculations;
using BlazorDemo.OfficeFileApi.Services;
using System.Text;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.Calculations {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelCircularReferencesController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("circular-references/detect")]
        public async Task<IActionResult> Detect(IFormFile file, [FromForm] ExcelCircularReferencesModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Detect circular references
            string calculationResult = await DetectWorkbookCircularReferences(inputStream, model.MaxIterationCount, targetFormat);
            using var outputStream = new MemoryStream();
            using(StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(calculationResult);
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
        }

        #region Code
        async Task<string> DetectWorkbookCircularReferences(Stream inputStream, int maxIterationCount, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            await workbook.LoadDocumentAsync(inputStream);

            // Create CustomCalculationService and replace the default service
            CustomCalculationService customCalculationService = new CustomCalculationService();
            workbook.ReplaceService<ICustomCalculationService>(customCalculationService);

            // Set calculation engine type to ChainBased
            if(workbook.Options.CalculationEngineType != CalculationEngineType.ChainBased)
                workbook.Options.CalculationEngineType = CalculationEngineType.ChainBased;

            // Circular references are calulated only if the Iterative property is set to true
            workbook.DocumentSettings.Calculation.Iterative = true;

            // Set maximum iteration count
            workbook.DocumentSettings.Calculation.MaxIterationCount = maxIterationCount;

            // Force workbook calculation
            workbook.CalculateFullRebuild();

            // Prepare calculation result
            var circularReferencedCells = customCalculationService.CircularReferencedCells;

            CalculationResult calculationResult = new CalculationResult();
            foreach(var circularReferencedCell in circularReferencedCells) {
                var worksheet = workbook.Worksheets[circularReferencedCell.SheetId - 1];
                string sheetName = worksheet.Name;

                var cell = worksheet.Cells[circularReferencedCell.RowIndex, circularReferencedCell.ColumnIndex];
                string cellAddress = cell.GetReferenceA1();
                object cellValue = cell.Value.ToObject();
                calculationResult.Cells.Add(new CellInfo(sheetName, cellAddress, cellValue));
            }

            // Return calculation result as a string
            return calculationResult.ToString();
        }

        // CustomCalculationService implementation to detect circular references
        public class CustomCalculationService : DevExpress.XtraSpreadsheet.Services.ICustomCalculationService {
            public HashSet<CellKey> CircularReferencedCells { get; set; } = new HashSet<CellKey>();
            public CustomCalculationService() { }
            void Reset() {
                // Reset service when a new calculation starts
                CircularReferencedCells.Clear();
            }

            // This method is called when the calculation process starts 
            public bool OnBeginCalculation() {
                Reset();
                // return True to perform a calculation. Return False to cancel it
                return true;
            }

            // This method is called when a cell is being calculated
            public void OnBeginCellCalculation(CellCalculationArgs args) { }

            // This method is called when a circular reference calculation starts
            public bool OnBeginCircularReferencesCalculation() {
                // If the Workbook.DocumentSettings.Calculation.Iterative property is false (default), this method is not called.
                // If a circular reference calculation starts, the OnBeginCellCalculation method is called again.            
                // True to perform a calculation. Return False to cancel it.
                return true;
            }

            // This method is called after the calculation process has finished
            public void OnEndCalculation() { }

            // This method is called after a cell calculation has finished
            public void OnEndCellCalculation(CellKey cellKey, CellValue startValue, CellValue endValue) { }

            // This method is called after a circular reference calculation has finished
            public void OnEndCircularReferencesCalculation(IList<CellKey> cellKeys) {
                // Store cell keys of cells with circular references to use them later
                CircularReferencedCells = new HashSet<DevExpress.Spreadsheet.CellKey>(cellKeys);
            }

            public bool ShouldMarkupCalculateAlwaysCells() {
                // Mark as needing calculation the "calculate always" cells, 
                // such as cells containing volatile function or referencing another "calculate always" cell.
                return true;
            }
        }

        public class CalculationResult {
            public bool HasCircularReferences { get { return Cells.Count > 0; } }
            public List<CellInfo> Cells { get; set; } = new List<CellInfo>();
            public override string ToString() {
                StringBuilder sb = new StringBuilder();
                if(HasCircularReferences) {
                    sb.AppendLine("Circular references detected.");
                    sb.AppendLine(string.Join(Environment.NewLine, Cells));
                } else {
                    sb.Append("No circular references detected.");
                }
                return sb.ToString().TrimEnd();
            }
        }

        public class CellInfo {
            public string SheetName { get; set; }
            public string A1Reference { get; set; }
            public object Value { get; set; }
            public CellInfo(string sheetName, string a1Reference, object value) {
                SheetName = sheetName;
                A1Reference = a1Reference;
                Value = value;
            }
            public override string ToString() {
                return $"Sheet = \"{SheetName}\", Cell = {A1Reference}, Value = {Value}";
            }
        }
        #endregion Code
    }
}
