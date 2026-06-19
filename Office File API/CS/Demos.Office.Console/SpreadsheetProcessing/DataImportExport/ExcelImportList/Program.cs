using DevExpress.Spreadsheet;

class Program {
    const string OutputFileName = "Imported_Users.xlsx";
    static readonly string[] SelectedFields = [nameof(UserModel.FirstName), nameof(UserModel.LastName), nameof(UserModel.BirthDate)];

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var outputStream = ImportList(CreateUsers(), SelectedFields, DocumentFormat.Xlsx);
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

    static Stream ImportList(IList<UserModel> users, string[] selectedFields, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        var firstWorksheet = workbook.Worksheets.First();
        firstWorksheet.Name = "Users";
        firstWorksheet["A1"].Value = "Import data from a list of users:";

        if(selectedFields == null || !selectedFields.Any()) {
            selectedFields = [nameof(UserModel.FirstName), nameof(UserModel.LastName), nameof(UserModel.BirthDate)];
        }

        for(int i = 0; i < selectedFields.Length; i++) {
            firstWorksheet[0, i + 1].Value = selectedFields[i];
        }

        var formatting = firstWorksheet["B1:D1"].BeginUpdateFormatting();
        formatting.Font.Bold = true;
        firstWorksheet["B1:D1"].EndUpdateFormatting(formatting);

        var importOptions = new DataSourceImportOptions() {
            PropertyNames = selectedFields
        };
        firstWorksheet.Import(users, 1, 1, importOptions);
        firstWorksheet.Columns.AutoFit(0, 4);

        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }

    static IList<UserModel> CreateUsers() {
        return new List<UserModel> {
            new() { FirstName = "Nancy", LastName = "Davolio", BirthDate = new DateTime(1982, 3, 14) },
            new() { FirstName = "Andrew", LastName = "Fuller", BirthDate = new DateTime(1979, 7, 21) },
            new() { FirstName = "Janet", LastName = "Leverling", BirthDate = new DateTime(1991, 11, 5) }
        };
    }
}

sealed class UserModel {
    public string FirstName { get; set; } = string.Empty;
    public string LastName { get; set; } = string.Empty;
    public DateTime? BirthDate { get; set; }
}
