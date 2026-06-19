using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
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
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
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
        using Workbook workbook = new Workbook();

        workbook.BeginUpdate();
        try {
            Worksheet worksheet = workbook.Worksheets.First();
            worksheet.Name = "Rich Text Formatting";

            RichTextString richText = new RichTextString();
            richText.AddTextRun(firstTextPart + " ", new RichTextRunFont("Arial", 14, firstPartColor));
            richText.AddTextRun(secondTextPart + " ", new RichTextRunFont("Tahoma", 14, secondPartColor));
            richText.AddTextRun(thirdTextPart, new RichTextRunFont("Castellar", 14, thirdPartColor));

            worksheet["A1"].SetRichText(richText);
        }
        finally {
            workbook.EndUpdate();
        }

        Stream outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }
}
