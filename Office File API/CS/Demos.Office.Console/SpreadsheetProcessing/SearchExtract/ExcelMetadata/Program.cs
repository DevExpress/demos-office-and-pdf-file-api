using DevExpress.Spreadsheet;
using System.Text;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
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
            string metadataText = ReadDocumentPropertiesText(inputStream, ReadDocumentProperties, ReadCustomProperties);
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
        using var workbook = new Workbook();
        var documentProperties = workbook.LoadDocumentProperties(inputStream);

        if(readDocumentProperties) {
            output.AppendLine("MAIN DOCUMENT PROPERTIES:");
            output.AppendLine(new string('-', 30));
            output.AppendLine($"Application: {documentProperties.Application ?? string.Empty}");
            output.AppendLine($"Author: {documentProperties.Author ?? string.Empty}");
            output.AppendLine($"Category: {documentProperties.Category ?? string.Empty}");
            output.AppendLine($"Company: {documentProperties.Company ?? string.Empty}");
            output.AppendLine($"Content Status: {documentProperties.ContentStatus ?? string.Empty}");
            output.AppendLine($"Created Date: {documentProperties.Created:yyyy-MM-dd HH:mm:ss}");
            output.AppendLine($"Description: {documentProperties.Description ?? string.Empty}");
            output.AppendLine($"Document Revision: {documentProperties.DocumentRevision ?? string.Empty}");
            output.AppendLine($"Document Version: {documentProperties.DocumentVersion ?? string.Empty}");
            output.AppendLine($"Keywords: {documentProperties.Keywords ?? string.Empty}");
            output.AppendLine($"Last Modified By: {documentProperties.LastModifiedBy ?? string.Empty}");
            output.AppendLine($"Manager: {documentProperties.Manager ?? string.Empty}");
            output.AppendLine($"Modified Date: {documentProperties.Modified:yyyy-MM-dd HH:mm:ss}");
            output.AppendLine($"Printed Date: {documentProperties.Printed:yyyy-MM-dd HH:mm:ss}");
            output.AppendLine($"Security: {documentProperties.Security}");
            output.AppendLine($"Subject: {documentProperties.Subject ?? string.Empty}");
            output.AppendLine($"Title: {documentProperties.Title ?? string.Empty}");
            output.AppendLine($"Version: {documentProperties.Version ?? string.Empty}");
            output.AppendLine();
        }

        if(readCustomProperties) {
            var customProperties = documentProperties.Custom;
            output.AppendLine("CUSTOM PROPERTIES:");
            output.AppendLine(new string('-', 30));

            bool hasCustomProperties = false;
            foreach(string propName in customProperties.Names) {
                hasCustomProperties = true;
                var propValue = customProperties[propName];
                output.AppendLine($"{propName}: {propValue?.ToString() ?? string.Empty} ({propValue?.GetType().Name ?? "String"})");
            }

            if(!hasCustomProperties)
                output.AppendLine("No custom properties found.");

            output.AppendLine();
        }

        return output.ToString();
    }
}
