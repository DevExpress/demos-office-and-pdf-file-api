using DevExpress.XtraRichEdit;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Comments_template.docx");
    const string OutputFileName = "Comments_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = AddNewComment(inputStream, OutputFormat);
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

    static Stream AddNewComment(Stream inputStream, DocumentFormat outputFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;
        var paragraph = document.Paragraphs[0];
        var comment = document.Comments.Create(paragraph.Range, "TestUser");
        var commentDoc = comment.BeginUpdate();
        commentDoc.InsertText(commentDoc.Range.Start, "Test Comment");
        comment.EndUpdate(commentDoc);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
