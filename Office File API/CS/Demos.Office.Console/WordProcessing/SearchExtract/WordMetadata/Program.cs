using System.Text;
using DevExpress.XtraRichEdit;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Document_Metadata.txt";
    const bool ReadDocumentProperties = true;
    const bool ReadCustomProperties = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            var metadataText = ReadDocumentPropertiesText(inputStream, ReadDocumentProperties, ReadCustomProperties);
            File.WriteAllText(outputPath, metadataText, Encoding.UTF8);
            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static string ReadDocumentPropertiesText(Stream inputStream, bool readDocumentProperties, bool readCustomProperties) {
        var output = new StringBuilder();
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;

        if(readDocumentProperties) {
            var coreProperties = document.DocumentProperties;
            output.AppendLine("CORE DOCUMENT PROPERTIES:");
            output.AppendLine(new string('-', 30));
            output.AppendLine($"Title: {coreProperties.Title ?? string.Empty}");
            output.AppendLine($"Author: {coreProperties.Author ?? string.Empty}");
            output.AppendLine($"Subject: {coreProperties.Subject ?? string.Empty}");
            output.AppendLine($"Keywords: {coreProperties.Keywords ?? string.Empty}");
            output.AppendLine($"Category: {coreProperties.Category ?? string.Empty}");
            output.AppendLine($"Company: {coreProperties.Company ?? string.Empty}");
            output.AppendLine($"Manager: {coreProperties.Manager ?? string.Empty}");
            output.AppendLine($"Created Date: {coreProperties.Created:yyyy-MM-dd HH:mm:ss}");
            output.AppendLine($"Last Printed: {coreProperties.LastPrinted:yyyy-MM-dd HH:mm:ss}");
            output.AppendLine();
        }

        if(readCustomProperties) {
            var customProperties = document.CustomProperties;
            output.AppendLine("CUSTOM PROPERTIES:");
            output.AppendLine(new string('-', 30));
            var hasCustomProperties = false;
            foreach(string propertyName in customProperties.Names) {
                hasCustomProperties = true;
                var propertyValue = customProperties[propertyName];
                var valueString = propertyValue?.ToString() ?? string.Empty;
                var typeString = propertyValue?.GetType().Name ?? "String";
                output.AppendLine($"{propertyName}: {valueString} ({typeString})");
            }
            if(!hasCustomProperties)
                output.AppendLine("No custom properties found.");
            output.AppendLine();
        }

        return output.ToString();
    }
}
