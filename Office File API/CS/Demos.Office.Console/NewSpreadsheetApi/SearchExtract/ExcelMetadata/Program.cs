using System.Text;
using DevExpress.Docs.Spreadsheet;

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
        using var workbook = new Workbook(inputStream);
        var documentProperties = workbook.DocumentProperties;

        if(readDocumentProperties) {
            output.AppendLine("MAIN DOCUMENT PROPERTIES:");
            output.AppendLine(new string('-', 30));
            output.AppendLine($"Application: {documentProperties.Application ?? ""}");
            output.AppendLine($"Author: {documentProperties.Author ?? ""}");
            output.AppendLine($"Category: {documentProperties.Category ?? ""}");
            output.AppendLine($"Company: {documentProperties.Company ?? ""}");
            output.AppendLine($"Content Status: {documentProperties.ContentStatus ?? ""}");
            output.AppendLine($"Created Date: {documentProperties.Created.ToString("yyyy-MM-dd HH:mm:ss")}");
            output.AppendLine($"Description: {documentProperties.Description ?? ""}");
            output.AppendLine($"Document Revision: {documentProperties.DocumentRevision ?? ""}");
            output.AppendLine($"Document Version: {documentProperties.DocumentVersion ?? ""}");
            output.AppendLine($"Keywords: {documentProperties.Keywords ?? ""}");
            output.AppendLine($"Last Modified By: {documentProperties.LastModifiedBy ?? ""}");
            output.AppendLine($"Manager: {documentProperties.Manager ?? ""}");
            output.AppendLine($"Modified Date: {documentProperties.Modified.ToString("yyyy-MM-dd HH:mm:ss")}");
            output.AppendLine($"Printed Date: {documentProperties.Printed.ToString("yyyy-MM-dd HH:mm:ss")}");
            output.AppendLine($"Security: {documentProperties.Security}");
            output.AppendLine($"Subject: {documentProperties.Subject ?? ""}");
            output.AppendLine($"Title: {documentProperties.Title ?? ""}");
            output.AppendLine($"Version: {documentProperties.Version ?? ""}");
            output.AppendLine();
        }

        if(readCustomProperties) {
            var customProperties = documentProperties.CustomProperties;
            output.AppendLine("CUSTOM PROPERTIES:");
            output.AppendLine(new string('-', 30));

            if(customProperties.Count == 0)
                output.AppendLine("No custom properties found.");
            else
                foreach(var customProperty in customProperties)
                    output.AppendLine($"{customProperty.Key}: {customProperty.Value.Value} ({customProperty.Value.Type})");
            output.AppendLine();
        }

        return output.ToString();
    }
}
