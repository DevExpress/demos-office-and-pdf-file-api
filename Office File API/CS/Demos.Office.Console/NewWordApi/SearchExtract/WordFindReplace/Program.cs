using DevExpress.Docs.Office;
using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Modified_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const FindReplaceActionType Operation = FindReplaceActionType.Replace;
    const string SearchText = "search text";
    const string ReplaceText = "replacement text";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            var searchOptions = new TextSearchOptions();
            using var outputStream = Operation == FindReplaceActionType.Replace
                ? FindReplace(inputStream, SearchText, searchOptions, ReplaceText, OutputFormat)
                : FindRemove(inputStream, SearchText, searchOptions, OutputFormat);
            outputStream.Position = 0;
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

    static Stream FindReplace(Stream inputStream, string findText, TextSearchOptions searchOptions, string replaceText, DocumentFormat outputFormat) {
        // Load document
        using var document = new WordDocument(inputStream);

        document.Body.ReplaceText(findText, replaceText, searchOptions);

        var outputStream = new MemoryStream();
        document.Save(outputStream, outputFormat);
        return outputStream;
    }

    static Stream FindRemove(Stream inputStream, string findText, TextSearchOptions searchOptions, DocumentFormat outputFormat) {
        // Load document
        using var document = new WordDocument(inputStream);

        // Find all occurrences of the search text
        var searchResults = document.Body.FindText(findText, searchOptions);

        // Remove found ranges
        for(int i = searchResults.Count - 1; i >= 0; i--)
            document.Body.Delete(searchResults[i]);

        var outputStream = new MemoryStream();
        document.Save(outputStream, outputFormat);
        return outputStream;
    }
}

enum FindReplaceActionType {
    Replace,
    Remove
}