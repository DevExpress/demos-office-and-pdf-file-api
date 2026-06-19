using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    const string OutputFileName = "List_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const int NumberOfItems = 5;
    const float IndentSize = 0.5f;
    const float HangingIndent = 0.25f;
    const float SpaceBefore = 0f;
    const float SpaceAfter = 0f;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var outputStream = CreateBulletList(OutputFormat);
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

    static Stream CreateBulletList(DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        var document = wordProcessor.Document;
        document.Unit = DevExpress.Office.DocumentUnit.Inch;

        for(int i = 1; i <= NumberOfItems; i++)
            document.AppendText($"Item {i}\n");
        document.Delete(document.Paragraphs.Last().Range);

        var list = document.AbstractNumberingLists.Add();
        list.NumberingType = NumberingType.Bullet;
        var level = list.Levels[0];
        ApplyListLevelFormatting(level, NumberingFormat.Bullet, "●", 1, IndentSize, HangingIndent, SpaceBefore, SpaceAfter);

        var bulletedList = document.NumberingLists.Add(0);
        document.Paragraphs.AddParagraphsToList(document.Range, bulletedList, 0);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void ApplyListLevelFormatting(ListLevel level, NumberingFormat numberingFormat, string displayFormat, int startNumber, float indent, float hangingIndent, float spacingBeforePoints, float spacingAfterPoints) {
        level.NumberingFormat = numberingFormat;
        level.DisplayFormatString = displayFormat;
        level.Start = startNumber;
        level.ParagraphProperties.LeftIndent = indent + hangingIndent;
        level.ParagraphProperties.FirstLineIndentType = ParagraphFirstLineIndent.Hanging;
        level.ParagraphProperties.FirstLineIndent = hangingIndent;
        level.ParagraphProperties.SpacingBefore = spacingBeforePoints / 72;
        level.ParagraphProperties.SpacingAfter = spacingAfterPoints / 72;
    }
}
