#region Code
using DevExpress.Spreadsheet;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelSortDataController : SpreadsheetControllerBase {
        [HttpPost("sort-data/sort")]
        public async Task<IActionResult> Sort(IFormFile file, [FromForm] ExcelSortDataModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);    
            Stream outputStream = null;
            bool ascending = model.SortOrder == SortOrder.Ascending;

            // Sort data on the first worksheet based on the specified sort mode
            switch(model.SortMode) {
                case SortMode.SingleColumn:
                    // Sort by a single column
                    outputStream = SortBySingleColumn(inputStream, model.SingleSortColumnName, ascending, targetFormat);
                    break;
                case SortMode.MultipleColumns:
                    // Sort by multiple columns
                    outputStream = SortByMultipleColumns(inputStream, model.MultipleSortColumnNames, ascending, targetFormat);
                    break;
                case SortMode.FontColor:
                    // Sort by font color
                    outputStream = SortByFontColorValue(inputStream, ascending, targetFormat);
                    break;
                case SortMode.Fill:
                    // Sort by fill color
                    outputStream = SortByFillColorValue(inputStream, ascending, targetFormat);
                    break;
                case SortMode.Custom:
                    // Use Custom Comparer
                    outputStream = SortByCustomComparer(inputStream, targetFormat);
                    break;
            }
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream SortBySingleColumn(Stream inputStream, string columnName, bool ascending, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Sort the "A3:F22" range in the specified order
            CellRange range = worksheet.Range["A3:F22"];
            int columnIndex = GetColumnIndexByName(columnName);
            worksheet.Sort(range, columnIndex, ascending);

            // Add a note
            string order = ascending ? "ascending" : "descending";
            worksheet["D1"].Value = $"Sort by column with index = {columnIndex} in {order} order";

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream SortByMultipleColumns(Stream inputStream, IEnumerable<string> columnNames, bool ascending, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Sort by the Continent and Country columns if columnNames aren't specified
            if(columnNames == null || columnNames.Count() == 0) {
                columnNames = new string[] {
                    "Continent",
                    "Country" };
            }

            // Create sorting fields
            List<SortField> fields = new List<SortField>();

            foreach(var columnName in columnNames) {
                int columnIndex = GetColumnIndexByName(columnName);
                SortField sortField = new SortField();
                sortField.ColumnOffset = columnIndex;
                sortField.Comparer = ascending ? worksheet.Comparers.Ascending : worksheet.Comparers.Descending;
                fields.Add(sortField);
            }

            // Sort the "A3:F22" range by sorting fields
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, fields);

            // Add a note
            string order = ascending ? "ascending" : "descending";
            worksheet["D1"].Value = $"Sort by columns: {string.Join(',', columnNames)} in {order} order";

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream SortByFontColorValue(Stream inputStream, bool ascending, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Sort the "A3:F22" range by column "F" in the specified order
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, GetColumnIndexByName("Area"), worksheet["F12"].Font.Color, ascending);

            // Add a note         
            worksheet["D1"].Value = $"Sort the Area column by Font color";

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream SortByFillColorValue(Stream inputStream, bool ascending, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Sort the "A3:F22" range by column "A" in the specified order
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, GetColumnIndexByName("Continent"), worksheet["A3"].Fill, ascending);

            // Add a note         
            worksheet["D1"].Value = $"Sort the Continent column by Fill color";

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream SortByCustomComparer(Stream inputStream, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet worksheet = workbook.Worksheets.First();

            // Sort the "A3:F22" range values using a custom comparer
            CellRange range = worksheet.Range["A3:F22"];
            worksheet.Sort(range, GetColumnIndexByName("Continent"), new SampleComparer());

            // Add a note         
            worksheet["D1"].Value = $"Sort first column by text lenght";

            // Create output stream and save document in target format
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

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }

        //Compares cell values by text length
        class SampleComparer : IComparer<DevExpress.Spreadsheet.CellValue> {
            public int Compare(DevExpress.Spreadsheet.CellValue a, DevExpress.Spreadsheet.CellValue b) {
                if(!a.IsText || !b.IsText) return 0;
                if(a.TextValue.Length == b.TextValue.Length) return 0;
                return (a.TextValue.Length > b.TextValue.Length) ? 1 : -1;
            }
        }
        #endregion Code
    }
}
