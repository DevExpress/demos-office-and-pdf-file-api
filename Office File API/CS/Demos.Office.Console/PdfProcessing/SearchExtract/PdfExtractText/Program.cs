using DevExpress.Pdf;
using System.Text;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Alternative.pdf");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var outputStream = GetText(File.OpenRead(inputPath), []);
            SaveDocument(outputStream, Path.Combine(outputDirectory, "Document Text.txt"));

            Console.WriteLine($"Created {outputDirectory}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static void SaveDocument(Stream stream, string outputPath) {
        stream.Position = 0;
        using var fileStream = File.Create(outputPath);
        stream.CopyTo(fileStream);
    }

    static Stream GetText(Stream documentStream, IEnumerable<int> pageRange) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);
        var outputStream = new MemoryStream();

        string text = string.Empty;

        if(!pageRange.Any())
            text = processor.GetText();
        else
            foreach(var index in pageRange)
                if(index < 1 || index > processor.Document.Pages.Count)
                    continue;
                else
                    text += processor.GetPageText(index);

        using var writer = new StreamWriter(outputStream, Encoding.UTF8, leaveOpen: true);
        writer.Write(text);
        writer.Flush();

        outputStream.Position = 0;
        return outputStream;
    }
}
