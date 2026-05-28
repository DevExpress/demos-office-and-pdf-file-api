#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelDataValidationController : SpreadsheetControllerBase {
        [HttpPost("data-validation/apply")]
        public async Task<IActionResult> ApplyValidation(IFormFile file, [FromForm] ExcelDataValidationModel model) {
            // Read file stream
            Stream inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            Stream outputStream = new MemoryStream();
            IEnumerable<DataValidationRule> selectedRules = model.SelectedRules;
            bool showInputMessage = model.ShowInputMessage;
            bool showErrorMessage = model.ShowErrorMessage;

            // Add data validation to the first worksheet
            if(selectedRules != null && selectedRules.Any()) {
                foreach(var rule in selectedRules) {
                    switch(rule) {
                        case DataValidationRule.DecimalNumberRange:
                            outputStream = AddDecimalNumberRangeValidation(inputStream, showInputMessage, showErrorMessage, outputFormat);
                            break;
                        case DataValidationRule.Formula:
                            outputStream = AddFormulaValidation(inputStream, showInputMessage, showErrorMessage, outputFormat);
                            break;
                        case DataValidationRule.TextLength:
                            outputStream = AddTextLengthValidation(inputStream, showInputMessage, showErrorMessage, outputFormat);
                            break;
                        case DataValidationRule.ListCreatedInCode:
                            outputStream = AddValidationBasedOnListCreatedInCode(inputStream, showInputMessage, showErrorMessage, outputFormat);
                            break;
                        case DataValidationRule.ExistingWorksheetValues:
                            outputStream = AddValidationBasedOnExistingWorksheetValues(inputStream, showInputMessage, showErrorMessage, outputFormat);
                            break;
                    }
                    inputStream = outputStream;
                }
            }
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AddDecimalNumberRangeValidation(Stream inputStream, bool showInputMessage, bool showErrorMessage, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Restrict data entry to a number within limits
            var validationRule = worksheet.DataValidations.Add(worksheet["F4:F11"], DataValidationType.Decimal, DataValidationOperator.Between, 10, 40);

            if(showInputMessage) {
                ShowInputMessage(validationRule, "Decimal Number", "Enter decimal numbers between 10 and 40");
            }
            if(showErrorMessage) {
                ShowErrorMessage(validationRule, "Wrong Number", "The value you entered is not valid. Enter a decimal number between 10 and 40.");
            }

            // Highlight data validation ranges
            HighlightDataValidationRanges(worksheet);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddFormulaValidation(Stream inputStream, bool showInputMessage, bool showErrorMessage, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Restrict data entry based on a formula
            var validationRule = worksheet.DataValidations.Add(worksheet["B4:B11"], DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)");

            if(showInputMessage) {
                // Show input message
                ShowInputMessage(validationRule, "Employee Id", "Enter 5-digit number");
            }
            if(showErrorMessage) {
                // Show error message
                ShowErrorMessage(validationRule, "Wrong Employee Id", "The value you entered is not valid. Use 5-digit number for the employee ID.");
            }
            // Highlight data validation ranges
            HighlightDataValidationRanges(worksheet);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddTextLengthValidation(Stream inputStream, bool showInputMessage, bool showErrorMessage, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Restrict data entry to 3 symbols
            var validationRule = worksheet.DataValidations.Add(worksheet["D4:D11"], DataValidationType.TextLength, DataValidationOperator.Equal, 3);

            if(showInputMessage) {
                ShowInputMessage(validationRule, "Alias", "Enter a text value that contains 3 symbols");
            }
            if(showErrorMessage) {
                ShowErrorMessage(validationRule, "Wrong Alias", "The value you entered is not valid. Entered value must contain 3 symbols.");
            }

            // Highlight data validation ranges
            HighlightDataValidationRanges(worksheet);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddValidationBasedOnListCreatedInCode(Stream inputStream, bool showInputMessage, bool showErrorMessage, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Restrict data entry to values in a drop-down list specified in code
            var validationRule = worksheet.DataValidations.Add(worksheet["A4:A11"], DataValidationType.List, "PASS, FAIL");

            if(showInputMessage) {
                // Show input message
                ShowInputMessage(validationRule, "Assessment", "Enter a \"PASS\" or \"FAIL\" value");
            }
            if(showErrorMessage) {
                // Show error message
                ShowErrorMessage(validationRule, "Wrong Value", "The value you entered is not valid. Valid values: \"PASS\", \"FAIL\".");
            }

            // Highlight data validation ranges
            HighlightDataValidationRanges(worksheet);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddValidationBasedOnExistingWorksheetValues(Stream inputStream, bool showInputMessage, bool showErrorMessage, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Restrict data entry to values in a drop-down list obtained from a worksheet
            var validationRule = worksheet.DataValidations.Add(worksheet["E4:E11"], DataValidationType.List, ValueObject.FromRange(worksheet["H4:H9"].GetRangeWithAbsoluteReference()));

            if(showInputMessage) {
                // Show input message
                ShowInputMessage(validationRule, "Department", "Select a Department from the drop-down list");
            }
            if(showErrorMessage) {
                // Show error message
                ShowErrorMessage(validationRule, "Wrong Department", "The value you entered is not valid. Select a valid Department value from the drop-down list.");
            }

            // Highlight data validation ranges
            HighlightDataValidationRanges(worksheet);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        void ShowInputMessage(DevExpress.Spreadsheet.DataValidation validation, string inputTitle, string inputMessage) {
            validation.ShowInputMessage = true;
            validation.InputTitle = inputTitle;
            validation.InputMessage = inputMessage;
        }

        void ShowErrorMessage(DevExpress.Spreadsheet.DataValidation validation, string errorTitle, string errorMessage) {
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = DataValidationErrorStyle.Information;
            validation.ErrorTitle = errorTitle;
            validation.ErrorMessage = errorMessage;
        }

        void HighlightDataValidationRanges(Worksheet worksheet) {
            worksheet["H4:H9"].FillColor = Color.LightGray;
            int[] myColorScheme = new int[] { 0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3, 0xFFDFC4, 0xFFDAE9 };
            for(int i = 0; i < worksheet.DataValidations.Count; i++) {
                worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(myColorScheme[i]);
            }
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
