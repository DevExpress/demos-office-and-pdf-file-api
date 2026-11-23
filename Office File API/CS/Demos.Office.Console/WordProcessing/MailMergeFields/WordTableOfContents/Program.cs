using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "TableOfContents_Template.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Table_of_Contents.docx";
    const string OutputFormat = "docx";
    const bool UseHyperlinks = true;
    const bool IncludePageNumbers = true;
    const int StartLevel = 1;
    const int EndLevel = 3;
    const TOCInsertMode InsertMode = TOCInsertMode.OutlineLevels;
    const TabLeaderType TabLeaderTypeValue = TabLeaderType.Dots;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = CreateTableOfContents(inputStream, TabLeaderTypeValue, IncludePageNumbers, InsertMode, StartLevel, EndLevel, UseHyperlinks, GetDocumentFormat(OutputFormat));
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

    static MemoryStream CreateTableOfContents(Stream stream, TabLeaderType tabLeaderType, bool includePageNumbers, TOCInsertMode insertMode, int startLevel, int endLevel, bool useHyperlinks, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(stream);

        List<string> switches = new();
        if(useHyperlinks)
            switches.Add("\\h");
        if(insertMode != TOCInsertMode.TCFields)
            switches.Add($"\\o \"{startLevel}-{endLevel}\"");
        if(!includePageNumbers)
            switches.Add($"\\n \"{startLevel}-{endLevel}\"");

        switch(insertMode) {
            case TOCInsertMode.OutlineLevels:
                switches.Add("\\u");
                break;
            case TOCInsertMode.TCFields:
                switches.Add("\\f chapters");
                switches.Add($"\\l \"{startLevel}-{endLevel}\"");
                break;
        }

        string fieldCode = "TOC" + (switches.Count > 0 ? (" " + string.Join(" ", switches)) : string.Empty) + " ";
        var document = wordProcessor.Document;
        var tocParagraphRange = document.Bookmarks["toc_bookmark"].Range;
        Field tocField = document.Fields.Create(tocParagraphRange.Start, fieldCode);
        tocField.Update();
        document.UpdateAllFields();

        var output = new MemoryStream();
        wordProcessor.SaveDocument(output, targetFormat);
        output.Position = 0;
        return output;
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "doc" => DocumentFormat.Doc,
            "docx" => DocumentFormat.Docx,
            "docm" => DocumentFormat.Docm,
            "dotx" => DocumentFormat.Dotx,
            "dotm" => DocumentFormat.Dotm,
            "rtf" => DocumentFormat.Rtf,
            _ => DocumentFormat.Docx
        };
    }

    enum TOCInsertMode {
        OutlineLevels,
        ParagraphStyles,
        TCFields
    }
}
