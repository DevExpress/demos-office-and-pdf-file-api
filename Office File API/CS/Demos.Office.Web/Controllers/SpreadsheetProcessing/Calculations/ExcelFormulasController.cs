#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.Calculations;
using BlazorDemo.OfficeFileApi.Services;


namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.Calculations {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelFormulasController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("formulas/add")]
        public async Task<IActionResult> Add([FromForm] ExcelFormulasModel model) {
            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Create a document that contains formulas based on the specified type
            Stream outputStream = new MemoryStream();
            
            switch(model.FormulaType) {
                case FormulaType.ReferencesInFormulas:
                    // Add Formulas that use different cell refernce types
                    outputStream = AddFormulasWithCellReferences(targetFormat);
                    break;
                case FormulaType.NamedRangesInFormulas:
                    // Add a Formula that uses a named range
                    outputStream = AddFormulaWithNamedRange("myRange", targetFormat);
                    break;
                case FormulaType.CreateNamedFormulas:
                    // Create Named Formulas
                    outputStream = CreateNamedFormulas("Range_Sum", "Range_DoubleSum", targetFormat);
                    break;
                case FormulaType.FunctionsInFormulas:
                    // Use Functions and Nested Functions in Formulas
                    outputStream = AddFormulasWithFunctions(targetFormat);
                    break;
                case FormulaType.CreateSharedFormulas:
                    // Create Shared Formulas
                    outputStream = CreateSharedFormulas(targetFormat);
                    break;
                case FormulaType.CreateArrayFormulas:
                    //Create Legacy and Dynamic Array Formulas
                    outputStream = CreateArrayFormulas(targetFormat);
                    break;
            }
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AddFormulasWithCellReferences(DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Create worksheets and add data
            Worksheet dataWorksheet = workbook.Worksheets.First();
            dataWorksheet.Name = "Data";

            Worksheet formulasWorksheet = workbook.Worksheets.Add();
            formulasWorksheet.Name = "Data And Formulas";

            dataWorksheet.Cells["A1"].Value = "Value 1";
            dataWorksheet.Cells["B1"].Value = "Value 2";
            dataWorksheet.Cells["C1"].Value = "Value 3";

            dataWorksheet.Cells["A2"].Value = 10;
            dataWorksheet.Cells["B2"].Value = 20;
            dataWorksheet.Cells["C2"].Value = 30;

            formulasWorksheet.Cells["A1"].Value = "Value 1";
            formulasWorksheet.Cells["B1"].Value = "Value 2";
            formulasWorksheet.Cells["C1"].Value = "Value 3";

            formulasWorksheet.Cells["A2"].Value = 10;
            formulasWorksheet.Cells["B2"].Value = 20;
            formulasWorksheet.Cells["C2"].Value = 30;

            formulasWorksheet.Cells["F1"].Value = "Formula";
            formulasWorksheet.Cells["G1"].Value = "Value";

            // Use a relative cell reference in a formula
            formulasWorksheet.Cells["E2"].Value = "A1 Relative References";
            formulasWorksheet.Cells["F2"].Value = "=A2+B2";
            formulasWorksheet.Cells["G2"].Formula = "=A2+B2";

            // Use an absolute cell reference in a formula
            formulasWorksheet.Cells["E3"].Value = "A1 Absolute References";
            formulasWorksheet.Cells["F3"].Value = "=$A$2";
            formulasWorksheet.Cells["G3"].Formula = "=$A$2";

            // Use a mixed cell reference in a formula
            formulasWorksheet.Cells["E4"].Value = "A1 Mixed References";
            formulasWorksheet.Cells["F4"].Value = "=A$2";
            formulasWorksheet.Cells["G4"].Formula = "=A$2";

            // Sum values of cells located in different worksheets
            formulasWorksheet.Cells["E5"].Value = "Cross-Worksheet References";
            formulasWorksheet.Cells["F5"].Value = "=Data!A2 + Data!B2";
            formulasWorksheet.Cells["G5"].Formula = "=Data!A2 + Data!B2";

            // Use a 3D cell reference in a formula
            formulasWorksheet.Cells["E6"].Value = "3D References";
            formulasWorksheet.Cells["F6"].Value = "=SUM('Data:Data And Formulas'!A2:C2)";
            formulasWorksheet.Cells["G6"].Formula = "=SUM('Data:Data And Formulas'!A2:C2)";

            // Auto-fit columns
            formulasWorksheet.Columns.AutoFit("E", "G");

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddFormulaWithNamedRange(string rangeName, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Add data to the first worksheet
            Worksheet formulasWorksheet = workbook.Worksheets.First();
            formulasWorksheet.Name = "Named Range In Formula";

            formulasWorksheet["B1"].Value = $"{rangeName}:";

            // Access the "A2:C5" range of cells in the worksheet
            CellRange range = formulasWorksheet.Range["A2:C5"];

            // Set the value for each cell in the range to 2
            range.Value = 2;

            // Specify the name for the created range
            range.Name = "myRange";

            // Specify cell borders for the created range
            range.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);

            // Add worksheet data
            string formula = $"=SUM({rangeName})";
            formulasWorksheet.Cells["E1"].Value = "Formula";
            formulasWorksheet.Cells["F1"].Value = "Value";
            formulasWorksheet.Cells["E2"].Value = formula;

            // Create a formula that sums up the values of all cells included in the specified named range
            formulasWorksheet.Cells["F2"].Formula = formula;

            // Auto-fit columns
            formulasWorksheet.Columns.AutoFit("E", "F");

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream CreateNamedFormulas(string formulaName1, string formulaName2, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Create worksheets and add data
            Worksheet worksheet1 = workbook.Worksheets.First();
            worksheet1.Name = "Data";

            Worksheet worksheet2 = workbook.Worksheets.Add("Formulas");
            worksheet2.Cells["A1"].Value = "Formula";
            worksheet2.Cells["B1"].Value = "Value";

            // Access the "A1:C3" range of cells in the worksheet
            CellRange range = worksheet1.Range["A1:C3"];

            // Set the value for each cell in the range to 2
            range.Value = 2;

            // Create a name for a formula that sums up the values of all cells included in the "A1:C3" range of the "Data" worksheet. 
            // The scope of this name will be limited by the "Data" worksheet
            worksheet1.DefinedNames.Add(formulaName1, "=SUM(Data!$A$1:$C$3)");

            // Create a name for a formula that doubles the value resulting from the "Range_Sum" named formula and
            // make this name available within the entire workbook
            workbook.DefinedNames.Add(formulaName2, $"=2*Data!{formulaName1}");

            // Create formulas that use other formulas with the specified names
            worksheet2.Cells["A2"].Value = $"=Data!{formulaName1}";
            worksheet2.Cells["B2"].Formula = $"=Data!{formulaName1}";

            worksheet2.Cells["A3"].Value = $"={formulaName2}";
            worksheet2.Cells["B3"].Formula = $"={formulaName2}";

            worksheet2.Cells["A4"].Value = $"={formulaName2} + 100";
            worksheet2.Cells["B4"].Formula = $"={formulaName2} + 100";

            // Auto-fit columns
            worksheet2.Columns.AutoFit("A", "B");

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddFormulasWithFunctions(DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Add data to the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();
            worksheet.Name = "Functions In Formulas";

            worksheet.Cells["A1"].Value = "Data";
            worksheet.Cells["B1"].Value = "Formula";
            worksheet.Cells["C1"].Value = "Value";

            // If the number in cell A2 is less than 10, the formula returns "Normal" 
            // and this text is displayed in cell C2. Otherwise, cell C2 displays "Excess"
            worksheet.Cells["A2"].Value = 15;
            worksheet.Cells["B2"].Value = @"=IF(A2<10, ""Normal"", ""Excess"")";
            worksheet.Cells["C2"].Formula = @"=IF(A2<10, ""Normal"", ""Excess"")";

            // Calculate the average value for cell values within the "A2:A7" range
            worksheet.Cells["A3"].Value = 3;
            worksheet.Cells["B3"].Value = "=AVERAGE(A2:A7)";
            worksheet.Cells["C3"].Formula = "=AVERAGE(A2:A7)";

            // Add the values contained in cells A3 through A5, add the value contained in cell A6, 
            // and add 100 to that result
            worksheet.Cells["A4"].Value = 3;
            worksheet.Cells["B4"].Value = "=SUM(A3:A5,A6,100)";
            worksheet.Cells["C4"].Formula = "=SUM(A3:A5,A6,100)";

            // Use a nested function in a formula.
            // Round the sum of the values contained in cells A6 and A7 to two decimal places
            worksheet.Cells["A5"].Value = 3;
            worksheet.Cells["B5"].Value = "=ROUND(SUM(A6,A7),2)";
            worksheet.Cells["C5"].Formula = "=ROUND(SUM(A6,A7),2)";

            // Add the current date to cell C6
            worksheet.Cells["A6"].Value = 20;
            worksheet.Cells["B6"].Value = "=Today()";
            worksheet.Cells["C6"].Formula = "=Today()";
            worksheet.Cells["C6"].NumberFormat = "m/d/yy";

            // Convert the specified text to uppercase
            worksheet.Cells["A7"].Value = 15.12345;
            worksheet.Cells["B7"].Value = @"=UPPER(""formula"")";
            worksheet.Cells["C7"].Formula = @"=UPPER(""formula"")";

            // Auto-fit columns
            worksheet.Columns.AutoFit("A", "C");

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream CreateSharedFormulas(DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Add data to the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();
            worksheet.Name = "Shared Formulas";

            worksheet.Cells["A1"].Value = "Values";
            worksheet.Cells["A2"].Value = 1;

            // Use the shared formula in the "A3:A11" range of cells.
            worksheet.Range["A3:A11"].Formula = "=SUM(A2+1)";

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream CreateArrayFormulas(DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Add data to the first worksheet
            Worksheet worksheet1 = workbook.Worksheets.First();
            worksheet1.Name = "Legacy Array Formulas";

            worksheet1.Cells["A1"].Value = "Value 1";
            worksheet1.Cells["B1"].Value = "Value 2";
            worksheet1.Cells["D1"].Value = "Array Formula";
            worksheet1.Cells["E1"].Value = "Value";

            // Access the "A2:B4" range of cells in the worksheet
            CellRange range = worksheet1.Range["A2:B2"];
            // Set the value for each cell in the range to 2
            range.Value = 2;

            // Access the "A3:B3" range of cells in the worksheet
            range = worksheet1.Range["A3:B3"];
            // Set the value for each cell in the range to 3
            range.Value = 3;

            // Access the "A4:B4" range of cells in the worksheet
            range = worksheet1.Range["A4:B4"];
            // Set the value for each cell in the range to 4
            range.Value = 4;

            // Create an array formula that multiplies values contained in the cell range A2 through A4
            // by the corresponding cells in the range B2 through B4,
            // and displays the results in cells E2 through E4
            worksheet1.Range["D2"].Value = "=A2:A4*B2:B4";
            worksheet1.Range["E2:E4"].ArrayFormula = "=A2:A4*B2:B4";

            // Auto-fit columns
            worksheet1.Columns.AutoFit("A", "E");

            // Add the second worksheet and insert dynamic array formulas
            Worksheet worksheet2 = workbook.Worksheets.Add("Dynamic Array Formulas");
            // Insert dynamic array formulas
            worksheet2["A1"].DynamicArrayFormulaInvariant = "={\"Red\",\"Green\",\"Orange\",\"Blue\"}";
            worksheet2.DynamicArrayFormulas.Add(worksheet2["A2"], "=LEN(A1:D1)");

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}
