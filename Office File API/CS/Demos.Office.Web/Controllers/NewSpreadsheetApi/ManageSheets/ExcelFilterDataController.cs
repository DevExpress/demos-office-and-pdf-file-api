#region Code
using DevExpress.Docs.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.ManageSheets {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelFilterDataController : SpreadsheetControllerBase {
        [HttpPost("filter-data/filter")]
        public async Task<IActionResult> Filter(IFormFile file, [FromForm] ExcelFilterDataModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            Stream outputStream = new MemoryStream();
            switch(model.FilterType) {
                case FilterType.ByCellValues:
                    outputStream = FilterByMultipleValues(inputStream, model.SelectedProductFilterValues, targetFormat);
                    break;
                case FilterType.ByDateAndTextValues:
                    outputStream = FilterMixedDataTypesByValues(inputStream, model.SelectedReportedDateFilterValue, targetFormat);
                    break;
                case FilterType.ByCustomDateRange:
                    outputStream = FilterDatesByCondition(inputStream, model.ReportedDateStartDate, model.ReportedDateEndDate, targetFormat);
                    break;
                case FilterType.ByCustomNumericRange:
                    outputStream = FilterNumericByCondition(inputStream, model.SalesFilterMinValue, model.SalesFilterMaxValue, targetFormat);
                    break;
                case FilterType.ByCustomText:
                    outputStream = FilterTextByCondition(inputStream, targetFormat);
                    break;
                case FilterType.ByTopValues:
                    outputStream = Top10FilterValue(inputStream, targetFormat);
                    break;
                case FilterType.Dynamic:
                    outputStream = DynamicFilterValue(inputStream, targetFormat);
                    break;
                case FilterType.ByFontColor:
                    outputStream = FilterByFontColor(inputStream, targetFormat);
                    break;
                case FilterType.ByFillColor:
                    outputStream = FilterByFillColor(inputStream, targetFormat);
                    break;
                case FilterType.ByBackgroundColor:
                    outputStream = FilterByBackgroundColor(inputStream, targetFormat);
                    break;
                case FilterType.SortFilteredRange:
                    outputStream = FilterAndSortByMultipleColumns(inputStream, model.SortOrder == SortOrder.Descending, targetFormat);
                    break;
            }
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FilterByMultipleValues(Stream inputStream, IEnumerable<string> selectedProductFilterValues, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            FilterValue[] filterValues;

            if(selectedProductFilterValues == null || !selectedProductFilterValues.Any()) {
                filterValues = ["Mozzarella di Giovanni", "Gorgonzola Telino"];
            } else {
                filterValues = selectedProductFilterValues.Select(value => (FilterValue)value).ToArray();
            }

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            AutoFilterColumn productsAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Product")];
            productsAutoFilterColumn.Criteria = new ValueFilterCriteria(filterValues);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterMixedDataTypesByValues(Stream inputStream, DateTime filterDate, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            DateGroup[] dateGroups = [new DateGroup(filterDate, DateTimeGroupingType.Month)];

            FilterValue[] textValues = ["gennaio 2015"];
            AutoFilterColumn reportedDateAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Reported Date")];
            reportedDateAutoFilterColumn.Criteria = new ValueFilterCriteria(textValues, dateGroups);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterDatesByCondition(Stream inputStream, DateTime reportedDateStartDate, DateTime reportedDateEndDate, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            AutoFilterColumn reportedDateAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Reported Date")];
            reportedDateAutoFilterColumn.Criteria = new ComparisonFilterCriteria(
                reportedDateStartDate, FilterComparisonOperator.GreaterThanOrEqual,
                reportedDateEndDate, FilterComparisonOperator.LessThanOrEqual, matchAll: true);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterNumericByCondition(Stream inputStream, int minValue, int maxValue, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            AutoFilterColumn salesAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Sales")];
            salesAutoFilterColumn.Criteria = new ComparisonFilterCriteria(
                minValue, FilterComparisonOperator.GreaterThanOrEqual,
                maxValue, FilterComparisonOperator.LessThanOrEqual, matchAll: true);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterTextByCondition(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            AutoFilterColumn productsAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Product")];
            productsAutoFilterColumn.Criteria = new ComparisonFilterCriteria(
                "*Gi*", FilterComparisonOperator.Equal,
                FilterValue.Blank, FilterComparisonOperator.Equal, matchAll: false);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream Top10FilterValue(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            AutoFilterColumn salesAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Sales")];
            salesAutoFilterColumn.Criteria = new TopBottomFilterCriteria(10);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream DynamicFilterValue(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            AutoFilterColumn salesAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Sales")];
            salesAutoFilterColumn.Criteria = new DynamicFilterCriteria(DynamicFilterType.AboveAverage);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterByFontColor(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            CellColor fontColor = regionalSalesWorksheet.Cells["D10"].GetEffectiveFormat().Font.Value.Color;
            AutoFilterColumn salesAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Sales")];
            salesAutoFilterColumn.Criteria = new FontColorFilterCriteria(fontColor);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterByFillColor(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            FillFormat fill = regionalSalesWorksheet.Cells["C10"].GetEffectiveFormat().Fill;
            AutoFilterColumn productsAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Product")];
            productsAutoFilterColumn.Criteria = new FillFilterCriteria(fill);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterByBackgroundColor(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

            var solidFill = (SolidFillFormat)regionalSalesWorksheet.Cells["C12"].GetEffectiveFormat().Fill;
            AutoFilterColumn productsAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Product")];
            productsAutoFilterColumn.Criteria = new FillFilterCriteria(solidFill.Color);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FilterAndSortByMultipleColumns(Stream inputStream, bool descending, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

            EnableFilter(regionalSalesWorksheet);

            CellColor fontColor = regionalSalesWorksheet.Cells["D12"].GetEffectiveFormat().Font.Value.Color;
            List<SortField> sortFields = new List<SortField> {
                new ValueSortField(GetColumnIndexByName("Region"), descending),
                new FontColorSortField(GetColumnIndexByName("Sales"), fontColor) { Descending = descending }
            };
            regionalSalesWorksheet.Sort(regionalSalesWorksheet["B3:E23"], sortFields);

            return CreateOutputStream(workbook, outputFormat);
        }

        Worksheet GetRegionalSalesWorksheet(Workbook workbook) {
            return workbook.Worksheets.First(sheet => sheet.Name == "Regional Sales");
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

        AutoFilter EnableFilter(Worksheet worksheet) {
            worksheet.AutoFilter = new AutoFilter(worksheet["B2:E23"]);
            return worksheet.AutoFilter;
        }

        Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
