using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Protected_Document.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;
    const string WorkbookPassword = "123";
    const bool LockStructure = true;
    const bool LockWindows = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProtectWorkbookWithPassword(inputStream, OutputFormat, WorkbookPassword, LockStructure, LockWindows);
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

    static Stream ProtectWorkbookWithPassword(Stream inputStream, DocumentFormat outputFormat, string workbookPassword, bool lockStructure, bool lockWindows) {
        using var workbook = new Workbook(inputStream);

        if(workbook.ProtectionMode == WorkbookProtectionMode.None) {
            workbook.Protect(new WorkbookProtectionOptions(workbookPassword) {
                LockStructure = lockStructure,
                LockWindows = lockWindows,
                RequirePasswordToModify = true
            });
        }

        var outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}
