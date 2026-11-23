using DevExpress.Docs.Pdf;

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

    static Stream FlattenForm(Stream documentStream, bool flattenAll, string[] fieldsToFlatten) {
        using var doc = new PdfDocument(documentStream);

        if(flattenAll)
            doc.FlattenFormFields();
        else
            doc.FlattenFormFields(fieldsToFlatten);

        var outputStream = new MemoryStream();
        doc.Save(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }

    static Stream RemoveForm(Stream documentStream, bool removeAll, string[] fieldsToRemove) {
        using var doc = new PdfDocument(documentStream);

        if(removeAll)
            doc.Fields.Clear();
        else
            foreach(var fieldName in fieldsToRemove) {
                var field = doc.Fields.FindByName(fieldName);
                if(field != null)
                    doc.Fields.Remove(field);
            }

        var outputStream = new MemoryStream();
        doc.Save(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
