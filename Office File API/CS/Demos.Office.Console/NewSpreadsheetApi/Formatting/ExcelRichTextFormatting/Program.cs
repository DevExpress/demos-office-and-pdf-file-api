using System.Drawing;
using DevExpress.Docs.Spreadsheet;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Rich_Text_Formatting.xlsx";
    const string FirstTextPart = "Rich";
    const string SecondTextPart = "Text";
    const string ThirdTextPart = "Formatting";
    static readonly Color FirstPartColor = ColorTranslator.FromHtml("#C59FC9");
    static readonly Color SecondPartColor = ColorTranslator.FromHtml("#2C608E");
    static readonly Color ThirdPartColor = ColorTranslator.FromHtml("#2F244F");
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var outputStream = CreateWorkbookWithRichTextFormatting(FirstTextPart, SecondTextPart, ThirdTextPart, FirstPartColor, SecondPartColor, ThirdPartColor, OutputFormat);
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

    static Stream CreateWorkbookWithRichTextFormatting(string firstTextPart, string secondTextPart, string thirdTextPart,
        Color firstPartColor, Color secondPartColor, Color thirdPartColor, DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Name = "Rich Text Formatting";

        RichText richText = new RichText()
            .AddRun(firstTextPart + " ", new FontFormat("Arial", 14).WithColor(firstPartColor))
            .AddRun(secondTextPart + " ", new FontFormat("Tahoma", 14).WithColor(secondPartColor))
            .AddRun(thirdTextPart, new FontFormat("Castellar", 14).WithColor(thirdPartColor));

        worksheet.Cells["A1"].SetRichText(richText);

        Stream outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }
}
