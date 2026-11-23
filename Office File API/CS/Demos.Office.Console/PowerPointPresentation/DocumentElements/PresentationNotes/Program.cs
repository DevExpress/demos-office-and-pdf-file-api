using DevExpress.Docs.Presentation;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.pptx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Presentation_Notes.pptx";
    const string NoteText = "Sample note text";
    const int SlideIndex = 1;
    const NotesActionType ActionType = NotesActionType.AddNoteToSlide;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Pptx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ApplyNotes(inputStream, ActionType, NoteText, SlideIndex - 1, OutputFormat);
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

    static Stream ApplyNotes(Stream input, NotesActionType action, string noteText, int slideIndex, DocumentFormat format) {
        using var presentation = new Presentation(input);

        switch(action) {
            case NotesActionType.ClearAllNotes:
                ClearAllNotes(presentation);
                break;
            case NotesActionType.AddNoteToSlide:
                AddUpdateNote(presentation, slideIndex, noteText);
                break;
        }

        var outputStream = new MemoryStream();
        presentation.Save(outputStream, format);
        outputStream.Position = 0;
        return outputStream;
    }

    static void AddUpdateNote(Presentation presentation, int slideIndex, string noteText) {
        if(slideIndex >= 0 && slideIndex < presentation.Slides.Count) {
            var slide = presentation.Slides[slideIndex];
            if(presentation.NotesMaster == null)
                presentation.NotesMaster = new NotesMaster(name: "notesMaster");

            NotesSlide notes;
            if(slide.Notes == null) {
                notes = new NotesSlide();
                slide.Notes = notes;
            }
            else
                notes = slide.Notes;

            notes.TextArea.Text = noteText;
        }
    }

    static void ClearAllNotes(Presentation presentation) {
        foreach(var slide in presentation.Slides)
            slide.Notes = null;
    }

    enum NotesActionType {
        AddNoteToSlide,
        ClearAllNotes
    }
}
