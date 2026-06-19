using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "OutlineGrouping_template.xlsx");
    const string OutputFileName = "Group_Document.xlsx";
    const bool GroupRows = true;
    const bool GroupColumns = false;
    const bool ShowSubtotals = true;
    const bool CollapseGroups = false;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = GroupData(inputStream, GroupRows, GroupColumns, ShowSubtotals, CollapseGroups, OutputFormat);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);
            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream GroupData(Stream inputStream, bool groupRows, bool groupColumns, bool showSubtotals, bool collapsed, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);
        Worksheet worksheet = workbook.Worksheets.First();
        if(!showSubtotals)
            worksheet.RemoveSubtotal(worksheet["B3:G35"]);
        if(groupRows) {
            worksheet.Rows.Group(3, 8, collapsed);
            worksheet.Rows.Group(showSubtotals ? 11 : 10, showSubtotals ? 16 : 15, collapsed);
            worksheet.Rows.Group(showSubtotals ? 19 : 17, showSubtotals ? 24 : 22, collapsed);
            worksheet.Rows.Group(showSubtotals ? 27 : 24, showSubtotals ? 32 : 29, collapsed);
            if(showSubtotals)
                worksheet.Rows.Group(2, 33, collapsed);
        }
        if(groupColumns)
            worksheet.Columns.Group("C", "F", collapsed);
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
