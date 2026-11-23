using DevExpress.Docs.Word;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "List_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const int NumberOfItems = 5;
    const string BulletCharacter = "●";
    const float IndentSize = 0.5f;
    const float HangingIndent = 0.25f;
    const float SpaceBefore = 0f;
    const float SpaceAfter = 0f;

    const float DocumentUnitsPerInch = 300f;
    const float DocumentUnitsPerPoint = DocumentUnitsPerInch / 72f;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var outputStream = CreateBulletList(NumberOfItems, BulletCharacter, IndentSize, HangingIndent, SpaceBefore, SpaceAfter, OutputFormat);
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

    static Stream CreateBulletList(int itemCount, string bullet, float indent, float hangingIndent, float spacingBeforePoints, float spacingAfterPoints, DocumentFormat targetFormat) {
        using var document = new WordDocument();

        // Create a new list pattern object
        var template = new ListTemplate();
        ApplyListLevelFormatting(template.Levels[0], NumberingFormat.Bullet, bullet, 1, indent, hangingIndent, spacingBeforePoints, spacingAfterPoints);
        // Create a new list based on the specific pattern
        var bulletedList = document.NumberingLists.Add(template);

        // Add paragraphs to the list
        for(var i = 1; i <= itemCount; i++) {
            var paragraph = document.Body.AddParagraph($"Item {i}");
            paragraph.NumberingList = bulletedList;
            paragraph.ListLevelIndex = 0;
        }

        return CreateOutputStream(document, targetFormat);
    }

    static void ApplyListLevelFormatting(ListLevel level, NumberingFormat numberingFormat, string displayFormat, int startNumber, float indent,
                                         float hangingIndent, float spacingBeforePoints, float spacingAfterPoints) {
        level.Format = numberingFormat;
        level.DisplayFormat = displayFormat;
        level.Start = startNumber;

        level.ParagraphProperties.LeftIndent = (indent + hangingIndent) * DocumentUnitsPerInch;
        level.ParagraphProperties.FirstLineIndent = -hangingIndent * DocumentUnitsPerInch;

        // Convert points to document units for spacing
        level.ParagraphProperties.SpacingBefore = spacingBeforePoints * DocumentUnitsPerPoint;
        level.ParagraphProperties.SpacingAfter = spacingAfterPoints * DocumentUnitsPerPoint;
    }

    static Stream CreateOutputStream(WordDocument document, DocumentFormat targetFormat) {
        var outputStream = new MemoryStream();
        document.Save(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}