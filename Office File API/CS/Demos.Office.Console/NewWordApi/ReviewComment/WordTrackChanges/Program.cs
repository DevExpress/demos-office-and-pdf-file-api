using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "TrackChanges_template.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "TrackChanges_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const RevisionActionType RevisionAction = RevisionActionType.Accept;
    const bool DisableTrackedChanges = false;
    const string AuthorFilter = "";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProcessTrackChangesDocument(inputStream, DisableTrackedChanges, AuthorFilter,
                RevisionAction == RevisionActionType.Accept, RevisionAction == RevisionActionType.Reject, OutputFormat);
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

    static Stream ProcessTrackChangesDocument(Stream inputStream, bool disableTrackChanges,
                                              string authorFilter, bool acceptAll, bool rejectAll, DocumentFormat outputFormat) {
        // Load document
        using var document = new WordDocument(inputStream);

        if(acceptAll)
            AcceptRevisions(document, authorFilter);
        else if(rejectAll)
            RejectRevisions(document, authorFilter);

        document.TrackChanges = !disableTrackChanges;

        var outputStream = new MemoryStream();
        document.Save(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void AcceptRevisions(WordDocument document, string authorFilter) {
        if(authorFilter == null)
            document.Revisions.AcceptAll();
        else
            document.Revisions.AcceptAll(revision => revision.Author.Contains(authorFilter, StringComparison.OrdinalIgnoreCase));
    }

    static void RejectRevisions(WordDocument document, string authorFilter) {
        if(authorFilter == null)
            document.Revisions.RejectAll();
        else
            document.Revisions.RejectAll(revision => revision.Author.Contains(authorFilter, StringComparison.OrdinalIgnoreCase));
    }
}

enum RevisionActionType {
    None,
    Accept,
    Reject
}