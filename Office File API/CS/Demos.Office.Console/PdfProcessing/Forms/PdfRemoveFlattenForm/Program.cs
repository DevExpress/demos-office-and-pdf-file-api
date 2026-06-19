using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Form.pdf");
    const string OutputDirectoryName = "Output";

    static readonly string[] SampleFields = [
        "LastName",
        "FirstName",
        "Nationality",
        "Address",
        "PassportNo",
        "VisaNo",
        "Gender",
        "MM",
        "DD",
        "YYYY",
        "FlightNo"
    ];

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var inputBytes = File.ReadAllBytes(inputPath);

            SaveDocument(FlattenForm(new MemoryStream(inputBytes), flattenAll: true, SampleFields), Path.Combine(outputDirectory, "flatten_all.pdf"));
            SaveDocument(FlattenForm(new MemoryStream(inputBytes), flattenAll: false, ["LastName", "FirstName"]), Path.Combine(outputDirectory, "flatten_selected.pdf"));
            SaveDocument(RemoveForm(new MemoryStream(inputBytes), removeAll: true, SampleFields), Path.Combine(outputDirectory, "remove_all.pdf"));
            SaveDocument(RemoveForm(new MemoryStream(inputBytes), removeAll: false, ["LastName", "FirstName"]), Path.Combine(outputDirectory, "remove_selected.pdf"));

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

    static Stream FlattenForm(MemoryStream documentStream, bool flattenAll, string[] fieldsToFlatten) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        if(flattenAll)
            processor.FlattenForm();
        else
            foreach(var field in fieldsToFlatten)
                processor.FlattenFormField(field);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }

    static Stream RemoveForm(MemoryStream documentStream, bool removeAll, string[] fieldsToRemove) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        if(removeAll)
            processor.RemoveForm();
        else
            foreach(var field in fieldsToRemove)
                processor.RemoveFormField(field);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
