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
    public class ExcelSortDataController : SpreadsheetControllerBase {
        [HttpPost("sort-data/sort")]
        public async Task<IActionResult> Sort(IFormFile file, [FromForm] ExcelSortDataModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);
            Stream outputStream = null;
            bool ascending = model.SortOrder == SortOrder.Ascending;

            switch(model.SortMode) {
                case SortMode.SingleColumn:
                    outputStream = SortBySingleColumn(inputStream, model.SingleSortColumnName, ascending, targetFormat);
                    break;
                case SortMode.MultipleColumns:
                    outputStream = SortByMultipleColumns(inputStream, model.MultipleSortColumnNames, ascending, targetFormat);
                    break;
                case SortMode.FontColor:
                    outputStream = SortByFontColorValue(inputStream, ascending, targetFormat);
                    break;
                case SortMode.Fill:
                    outputStream = SortByFillColorValue(inputStream, ascending, targetFormat);
                    break;
                case SortMode.Custom:
                    outputStream = SortByCustomComparer(inputStream, targetFormat);
                    break;
            }
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream SortBySingleColumn(Stream inputStream, string columnName, bool ascending, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet worksheet = workbook.Worksheets[0];

            CellRange range = worksheet["A3:F22"];
            int columnIndex = GetColumnIndexByName(columnName);
            worksheet.Sort(range, columnIndex, ascending);

            string order = ascending ? "ascending" : "descending";
            worksheet.Cells["D1"].Value = $"Sort by column with index = {columnIndex} in {order} order";

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream SortByMultipleColumns(Stream inputStream, IEnumerable<string> columnNames, bool ascending, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet worksheet = workbook.Worksheets[0];

            if(columnNames == null || !columnNames.Any()) {
                columnNames = new string[] {
                    "Continent",
                    "Country" };
            }

            List<SortField> fields = new List<SortField>();

            foreach(var columnName in columnNames) {
                int columnIndex = GetColumnIndexByName(columnName);
                fields.Add(new ValueSortField(columnIndex, descending: !ascending));
            }

            CellRange range = worksheet["A3:F22"];
            worksheet.Sort(range, fields);

            string order = ascending ? "ascending" : "descending";
            worksheet.Cells["D1"].Value = $"Sort by columns: {string.Join(',', columnNames)} in {order} order";

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream SortByFontColorValue(Stream inputStream, bool ascending, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet worksheet = workbook.Worksheets[0];

            CellRange range = worksheet["A3:F22"];
            CellColor fontColor = worksheet.Cells["F12"].GetEffectiveFormat().Font.Value.Color;
            var sortField = new FontColorSortField(GetColumnIndexByName("Area"), fontColor) { Descending = !ascending };
            worksheet.Sort(range, new SortField[] { sortField });

            worksheet.Cells["D1"].Value = $"Sort the Area column by Font color";

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream SortByFillColorValue(Stream inputStream, bool ascending, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet worksheet = workbook.Worksheets[0];

            CellRange range = worksheet["A3:F22"];
            FillFormat fill = worksheet.Cells["A3"].GetEffectiveFormat().Fill;
            var sortField = new FillSortField(GetColumnIndexByName("Continent"), fill) { Descending = !ascending };
            worksheet.Sort(range, new SortField[] { sortField });

            worksheet.Cells["D1"].Value = $"Sort the Continent column by Fill color";

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream SortByCustomComparer(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet worksheet = workbook.Worksheets[0];

            CellRange range = worksheet["A3:F22"];
            var sortField = new ValueSortField(GetColumnIndexByName("Continent")) { Comparer = new SampleComparer() };
            worksheet.Sort(range, new SortField[] { sortField });

            worksheet.Cells["D1"].Value = $"Sort first column by text length";

            return CreateOutputStream(workbook, outputFormat);
        }

        int GetColumnIndexByName(string columnName) {
            return columnName.ToLowerInvariant()
                switch {
                    "continent" => 0,
                    "country" => 1,
                    "symbol" => 2,
                    "population" => 3,
                    "gdp" => 4,
                    "area" => 5,
                    _ => 0
                };
        }

        Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        class SampleComparer : IComparer<CellValue> {
            public int Compare(CellValue a, CellValue b) {
                if(a.Type != CellValueType.Text || b.Type != CellValueType.Text) return 0;
                if(a.TextValue.Length == b.TextValue.Length) return 0;
                return (a.TextValue.Length > b.TextValue.Length) ? 1 : -1;
            }
        }
        #endregion Code
    }
}
