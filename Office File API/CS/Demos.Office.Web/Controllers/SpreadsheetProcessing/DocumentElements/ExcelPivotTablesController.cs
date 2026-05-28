#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    [ApiController]
    [Route("api/spreadsheetprocessing/")]
    public class ExcelPivotTablesController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("pivot-tables/generate")]
        public async Task<IActionResult> GeneratePivotTablesWorkbook([FromForm] ExcelPivotTablesModel model) {
            // Determine target format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with a pivot table
            var outputStream = GenerateDocumentWithPivotTable(
                new PivotTableOptions() {
                    CreatePivotTableSheet = model.CreatePivotTableSheet,
                    ReportLayout = model.ReportLayout,
                    ShowRowGrandTotals = model.ShowRowGrandTotals,
                    ShowColumnGrandTotals = model.ShowColumnGrandTotals,
                    ShowSubtotals = model.ShowSubtotals,
                    ShowRowHeaders = model.ShowRowHeaders,
                    ShowColumnHeaders = model.ShowColumnHeaders,
                    ShowFieldHeaders = model.ShowFieldHeaders             
                },
                model.SalesRegion,
                outputFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream GenerateDocumentWithPivotTable(PivotTableOptions options, string salesRegion, DocumentFormat outputFormat) {
            using var workbook = new Workbook();
            var worksheet = workbook.Worksheets.First();

            // Specify pivot table data
            Worksheet dataWorksheet = workbook.Worksheets.First();
            dataWorksheet.Name = "Pivot Data";

            dataWorksheet["A1"].Value = "Product";
            dataWorksheet["B1"].Value = "Region";
            dataWorksheet["C1"].Value = "Category";
            dataWorksheet["D1"].Value = "Sales";

            string[] products = { "Laptop", "Desktop", "Tablet", "Phone" };
            string[] regions = { "North", "South", "East", "West" };
            string[] categories = { "Electronics", "Gadgets", "Devices" };

            for(int i = 0; i < 20; i++) {
                int row = i + 2;
                dataWorksheet[$"A{row}"].Value = products[i % products.Length];
                dataWorksheet[$"B{row}"].Value = regions[i % regions.Length];
                dataWorksheet[$"C{row}"].Value = categories[i % categories.Length];
                dataWorksheet[$"D{row}"].Value = 1000 + (i * 150);
                dataWorksheet[$"D{row}"].NumberFormat = "$#,##0.00";
            }

            // Apply formatting to pivot table data cells
            var headerRange = dataWorksheet["A1:D1"];
            headerRange.Font.Bold = true;
            headerRange.Fill.BackgroundColor = Color.LightGray;

            var dataRange = dataWorksheet.GetDataRange();
            dataRange.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);

            dataWorksheet.Columns.AutoFit(0, 3);

            Worksheet pivotWorksheet = dataWorksheet;

            // Create a new worksheet for the pivot table if necessary and activate it
            if(options.CreatePivotTableSheet) {
                pivotWorksheet = workbook.Worksheets.Add("Pivot Table Sheet");
                workbook.Worksheets.ActiveWorksheet = pivotWorksheet;
            }

            // Create pivot table
            PivotTable pivotTable = pivotWorksheet.PivotTables.Add(
                dataRange,
                options.CreatePivotTableSheet ? pivotWorksheet["A1"] : pivotWorksheet["F1"]
            );

            pivotTable.BeginUpdate();

            // Add fields
            var regionField = pivotTable.RowFields.Add(pivotTable.Fields["Region"]);
            pivotTable.RowFields.Add(pivotTable.Fields["Product"]);
            pivotTable.ColumnFields.Add(pivotTable.Fields["Category"]);
   
            // Add a data field and specify its number format
            var salesField = pivotTable.DataFields.Add(pivotTable.Fields["Sales"]);
            salesField.NumberFormat = "$#,##0.00";

            // Filter the "Region" field by text to display sales data for the specified region
            pivotTable.Filters.Add(regionField.Field, PivotFilterType.CaptionEqual, salesRegion);

            // Apply pivot table layout settings
            pivotTable.Layout.SetReportLayout(options.ReportLayout);
            pivotTable.Layout.ShowRowGrandTotals = options.ShowRowGrandTotals;
            pivotTable.Layout.ShowColumnGrandTotals = options.ShowColumnGrandTotals;

            if(options.ShowSubtotals)
                pivotTable.Layout.ShowAllSubtotals(false);

            // Apply pivot table appearance and view settings
            pivotTable.ShowRowHeaders = options.ShowRowHeaders;
            pivotTable.ShowColumnHeaders = options.ShowColumnHeaders;
            pivotTable.View.ShowFieldHeaders = options.ShowFieldHeaders;

            // Access the built-in pivot table style by its Id and apply it
            pivotTable.Style = workbook.TableStyles[BuiltInPivotStyleId.PivotStyleMedium2];

            pivotTable.EndUpdate();

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the document
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        public class PivotTableOptions {
            public bool CreatePivotTableSheet { get; set; } = false;
            public PivotReportLayout ReportLayout { get; set; } = PivotReportLayout.Compact;
            public bool ShowRowGrandTotals { get; set; } = true;
            public bool ShowColumnGrandTotals { get; set; } = true;
            public bool ShowSubtotals { get; set; } = true;
            public bool ShowRowHeaders { get; set; } = true;
            public bool ShowColumnHeaders { get; set; } = true;
            public bool ShowFieldHeaders { get; set; } = true;
        }
        #endregion Code
    }
}
