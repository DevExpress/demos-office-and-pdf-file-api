using System.Text;
using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Document_Metadata.txt";
    const bool IncludeDocumentProperties = true;
    const bool IncludeCustomProperties = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            var metadataText = ReadDocumentProperties(inputStream, IncludeDocumentProperties, IncludeCustomProperties);
            File.WriteAllText(outputPath, metadataText, Encoding.UTF8);
            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static string ReadDocumentProperties(Stream inputStream, bool readDocumentProperties, bool readCustomProperties) {
        var output = new StringBuilder();

        using var document = new WordDocument(inputStream);
        var documentProperties = document.DocumentProperties;

        if(readDocumentProperties) {
            // Core Properties
            output.AppendLine("CORE DOCUMENT PROPERTIES:");
            output.AppendLine(new string('-', 30));
            output.AppendLine($"Title: {documentProperties.Title ?? ""}");
            output.AppendLine($"Author: {documentProperties.Author ?? ""}");
            output.AppendLine($"Subject: {documentProperties.Subject ?? ""}");
            output.AppendLine($"Keywords: {documentProperties.Keywords ?? ""}");
            output.AppendLine($"Category: {documentProperties.Category ?? ""}");
            output.AppendLine($"Company: {documentProperties.Company ?? ""}");
            output.AppendLine($"Manager: {documentProperties.Manager ?? ""}");
            output.AppendLine($"Created Date: {documentProperties.Created.ToString("yyyy-MM-dd HH:mm:ss")}");
            output.AppendLine($"Last Printed: {documentProperties.Printed.ToString("yyyy-MM-dd HH:mm:ss")}");
            output.AppendLine();
        }

        if(readCustomProperties) {
            var customProperties = documentProperties.CustomProperties;

            output.AppendLine("CUSTOM PROPERTIES:");
            output.AppendLine(new string('-', 30));

            // Check if there are any custom properties
            if(customProperties.Count == 0)
                output.AppendLine("No custom properties found.");
            else
                foreach(var customProperty in customProperties) {
                    var propertyValue = customProperty.Value.Value;
                    var valueText = propertyValue?.ToString() ?? "";
                    var typeName = propertyValue?.GetType().Name ?? "String";
                    output.AppendLine($"{customProperty.Key}: {valueText} ({typeName})");
                }

            output.AppendLine();
        }

        return output.ToString();
    }
}