using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Protected_Document.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;
    const string WorksheetPassword = "123";
    static readonly int[] WorksheetIndexes = { 0, 1 };
    static readonly WorksheetProtectionPermissions[] SelectedWorksheetProtectionPermissions = {
        WorksheetProtectionPermissions.SelectLockedCells,
        WorksheetProtectionPermissions.SelectUnlockedCells
    };

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProtectWorksheetWithPassword(
                inputStream,
                OutputFormat,
                WorksheetPassword,
                WorksheetIndexes,
                SelectedWorksheetProtectionPermissions);
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

    static Stream ProtectWorksheetWithPassword(Stream inputStream, DocumentFormat outputFormat, string worksheetPassword,
        int[] worksheetIndexes, IEnumerable<WorksheetProtectionPermissions> worksheetProtectionPermissions) {
        using var workbook = new Workbook(inputStream);

        var protectionOptions = new WorksheetProtectionOptions(worksheetPassword) {
            Permissions = WorksheetProtectionPermissions.None
        };
        foreach(WorksheetProtectionPermissions value in worksheetProtectionPermissions)
            protectionOptions.Permissions |= value;

        int worksheetCount = workbook.Worksheets.Count;
        foreach(int index in worksheetIndexes) {
            if(index < worksheetCount)
                workbook.Worksheets[index].Protect(protectionOptions);
        }

        var outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}
