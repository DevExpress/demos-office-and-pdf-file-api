#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

using FilterType = BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets.FilterType;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelFilterDataController : SpreadsheetControllerBase {
        [HttpPost("filter-data/filter")]
        public async Task<IActionResult> Filter(IFormFile file, [FromForm] ExcelFilterDataModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Filter data based on the specified filter type
            Stream outputStream = new MemoryStream();            
            switch(model.FilterType) {
                case FilterType.ByCellValues:
                    // Filter by cell values
                    outputStream = FilterByMultipleValues(inputStream, model.SelectedProductFilterValues, targetFormat);
                    break;
                case FilterType.ByDateAndTextValues:
                    // Filter by date values
                    outputStream = FilterMixedDataTypesByValues(inputStream, model.SelectedReportedDateFilterValue, targetFormat);
                    break;
                case FilterType.ByCustomDateRange:
                    // Filter dates by condition
                    outputStream = FilterDatesByCondition(inputStream, model.ReportedDateStartDate, model.ReportedDateEndDate, targetFormat);
                    break;
                case FilterType.ByCustomNumericRange:
                    // Filter numeric values by condition
                    outputStream = FilterNumericByCondition(inputStream, model.SalesFilterMinValue, model.SalesFilterMaxValue, targetFormat);
                    break;
                case FilterType.ByCustomText:
                    // Filter text values by condition
                    outputStream = FilterTextByCondition(inputStream, targetFormat);
                    break;
                case FilterType.ByTopValues:
                    // Filter by top values
                    outputStream = Top10FilterValue(inputStream, targetFormat);
                    break;
                case FilterType.Dynamic:
                    // Apply dynamic filter
                    outputStream = DynamicFilterValue(inputStream, targetFormat);
                    break;
                case FilterType.ByFontColor:
                    // Filter by font color
                    outputStream = FilterByFontColor(inputStream, targetFormat);
                    break;
                case FilterType.ByFillColor:
                    // Filter by fill color
                    outputStream = FilterByFillColor(inputStream, targetFormat);
                    break;
                case FilterType.ByBackgroundColor:
                    // Filter by background color
                    outputStream = FilterByBackgroundColor(inputStream, targetFormat);
                    break;
                case FilterType.SortFilteredRange:
                    // Apply filter and sort the filtered range by multiple columns
                    outputStream = FilterAndSortByMultipleColumns(inputStream, model.SortOrder == SortOrder.Descending, targetFormat);
                    break;
            }
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FilterByMultipleValues(Stream inputStream, IEnumerable<string> selectedProductFilterValues, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            //workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            CellValue[] cellValues;

            if(selectedProductFilterValues == null || !selectedProductFilterValues.Any()) {
                cellValues = ["Mozzarella di Giovanni", "Gorgonzola Telino"];
            } else {
                cellValues = new CellValue[selectedProductFilterValues.Count()];
                int index = 0;
                foreach(var cellValue in selectedProductFilterValues) {
                    cellValues[index++] = cellValue;
                }
            }

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Filter data in the "Product" column by an array of values
            AutoFilterColumn productsAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Product")];
            productsAutoFilterColumn.ApplyFilterCriteria(cellValues);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterMixedDataTypesByValues(Stream inputStream, DateTime filterDate, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Create date grouping item to filter the specified date
            IList<DateGrouping> groupings = new List<DateGrouping>();
            DateGrouping dateGrouping = new DateGrouping(filterDate, DateTimeGroupingType.Month);
            groupings.Add(dateGrouping);

            // Filter data in the "Reported Date" column to display values reported in January 2015 and specified date
            AutoFilterColumn reportedDateAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Reported Date")];
            reportedDateAutoFilterColumn.ApplyFilterCriteria("gennaio 2015", groupings);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterDatesByCondition(Stream inputStream, DateTime reportedDateStartDate, DateTime reportedDateEndDate, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Filter values in the "Reported Date" column to display dates that are between the specified start and end dates
            AutoFilterColumn reportedDateAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Reported Date")];
            reportedDateAutoFilterColumn.ApplyCustomFilter
                (reportedDateStartDate, FilterComparisonOperator.GreaterThanOrEqual,
                reportedDateEndDate, FilterComparisonOperator.LessThanOrEqual, true);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterNumericByCondition(Stream inputStream, int minValue, int maxValue, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Filter values in the "Sales" column that are in the specified range
            AutoFilterColumn salesAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Sales")];
            salesAutoFilterColumn.ApplyCustomFilter(minValue, FilterComparisonOperator.GreaterThanOrEqual, maxValue, FilterComparisonOperator.LessThanOrEqual, true);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterTextByCondition(Stream inputStream, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Filter values in the "Product" column that contain "Gi" and include empty cells
            AutoFilterColumn productsAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Product")];
            productsAutoFilterColumn.ApplyCustomFilter("*Gi*", FilterComparisonOperator.Equal, FilterValue.FilterByBlank, FilterComparisonOperator.Equal, false);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream Top10FilterValue(Stream inputStream, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Apply a filter to the "Sales" column to display the top ten values.
            AutoFilterColumn salesAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Sales")];
            salesAutoFilterColumn.ApplyTop10Filter(Top10Type.Top10Items, 10);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream DynamicFilterValue(Stream inputStream, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Apply a dynamic filter to the "Sales" column to display only values that are above the average
            AutoFilterColumn salesAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Sales")];
            salesAutoFilterColumn.ApplyDynamicFilter(DynamicFilterType.AboveAverage);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterByFontColor(Stream inputStream, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Filter values in the "Sales" column by font color
            AutoFilterColumn salesAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Sales")];
            salesAutoFilterColumn.ApplyFontColorFilter(regionalSalesWorksheet["D10"].Font.Color);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterByFillColor(Stream inputStream, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Filter values in the "Products" column by fill color
            AutoFilterColumn productsAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Product")];
            productsAutoFilterColumn.ApplyFillFilter(regionalSalesWorksheet["C10"].Fill);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterByBackgroundColor(Stream inputStream, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Filter values in the "Products" column by background color
            AutoFilterColumn productsAutoFilterColumn = regionalSalesWorksheet.AutoFilter.Columns[GetColumnIndexByName("Product")];
            productsAutoFilterColumn.ApplyFillColorFilter(regionalSalesWorksheet["C12"].FillColor);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterAndSortByMultipleColumns(Stream inputStream, bool descending, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Activate the "Regional sales" worksheet
            Worksheet regionalSalesWorksheet = workbook.Worksheets["Regional sales"];
            workbook.Worksheets.ActiveWorksheet = regionalSalesWorksheet;

            // Enable filtering
            EnableFilter(regionalSalesWorksheet);

            // Sort data in the "B2:E23" range by columns "Region" and "Sales"
            List<SortCondition> sortConditions = new List<SortCondition>();
            Color color = regionalSalesWorksheet["D12"].Font.Color;

            sortConditions.Add(new SortCondition(GetColumnIndexByName("Region"), descending));
            sortConditions.Add(new SortCondition(GetColumnIndexByName("Sales"), color, descending));
            regionalSalesWorksheet.AutoFilter.SortState.Sort(sortConditions);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        int GetColumnIndexByName(string columnName) {
            return columnName.ToLowerInvariant()
                switch {
                    "region" => 0,
                    "product" => 1,
                    "sales" => 2,
                    "reported date" => 3,
                    _ => 0
                };
        }

        void EnableFilter(Worksheet worksheet) { 
            // Enable filtering for the "B2:E23" cell range
            CellRange range = worksheet["B2:E23"];
            worksheet.AutoFilter.Apply(range);
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
