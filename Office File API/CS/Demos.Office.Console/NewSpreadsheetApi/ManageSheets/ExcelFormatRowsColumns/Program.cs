using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Modified_Document.xlsx";
    static readonly bool FormatRows = true;
    const FormatOperationType OperationType = FormatOperationType.Add;
    const int RowCount = 1;
    const int ColumnCount = 1;
    const int StartRowIndex = 1;
    const int StartColumnIndex = 1;
    const double RowHeight = 150;
    const double ColumnWidth = 20;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = Process(inputStream);
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

    static Stream Process(Stream inputStream) {
        if(FormatRows) {
            int startRowIndex = StartRowIndex - 1;
            return OperationType switch {
                FormatOperationType.Remove => RemoveRows(inputStream, RowCount, startRowIndex, OutputFormat),
                FormatOperationType.SetSize => SetSizeRows(inputStream, RowCount, startRowIndex, RowHeight, OutputFormat),
                FormatOperationType.Hide => HideRows(inputStream, RowCount, startRowIndex, OutputFormat),
                FormatOperationType.AutoFit => AutoFitRows(inputStream, RowCount, startRowIndex, OutputFormat),
                _ => AddRows(inputStream, RowCount, startRowIndex, OutputFormat)
            };
        }
        int startColumnIndex = StartColumnIndex - 1;
        return OperationType switch {
            FormatOperationType.Remove => RemoveColumns(inputStream, ColumnCount, startColumnIndex, OutputFormat),
            FormatOperationType.SetSize => SetSizeColumns(inputStream, ColumnCount, startColumnIndex, ColumnWidth, OutputFormat),
            FormatOperationType.Hide => HideColumns(inputStream, ColumnCount, startColumnIndex, OutputFormat),
            FormatOperationType.AutoFit => AutoFitColumns(inputStream, ColumnCount, startColumnIndex, OutputFormat),
            _ => AddColumns(inputStream, ColumnCount, startColumnIndex, OutputFormat)
        };
    }

    static Stream AddRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        firstWorksheet.InsertRows(startRowIndex, rowCount);

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream RemoveRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        firstWorksheet.DeleteRows(startRowIndex, rowCount);

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream SetSizeRows(Stream inputStream, int rowCount, int startRowIndex, double rowHeight, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        for(int i = startRowIndex; i < startRowIndex + rowCount; i++) {
            firstWorksheet.Rows[i].Height = rowHeight;
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream HideRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        for(int i = startRowIndex; i < startRowIndex + rowCount; i++) {
            firstWorksheet.Rows[i].Visible = false;
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AutoFitRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        for(int i = startRowIndex; i < startRowIndex + rowCount; i++) {
            firstWorksheet.Rows[i].AutoFit();
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AddColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        firstWorksheet.InsertColumns(startColumnIndex, columnCount);

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream RemoveColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        firstWorksheet.DeleteColumns(startColumnIndex, columnCount);

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream SetSizeColumns(Stream inputStream, int columnCount, int startColumnIndex, double columnWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        for(int i = startColumnIndex; i < startColumnIndex + columnCount; i++) {
            firstWorksheet.Columns[i].Width = columnWidth;
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream HideColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        for(int i = startColumnIndex; i < startColumnIndex + columnCount; i++) {
            firstWorksheet.Columns[i].Visible = false;
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AutoFitColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        for(int i = startColumnIndex; i < startColumnIndex + columnCount; i++) {
            firstWorksheet.Columns[i].AutoFit();
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}

enum FormatOperationType {
    Add,
    Remove,
    SetSize,
    Hide,
    AutoFit
}
