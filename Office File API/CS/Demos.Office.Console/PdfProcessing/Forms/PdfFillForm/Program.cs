using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Form.pdf");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var outputStream = FillForm(
                new MemoryStream(File.ReadAllBytes(inputPath)),
                lastName: "Leverling",
                firstName: "Janet",
                address: "98033, 722 Moss Bay Blvd., Kirkland, WA, USA",
                nationality: "Thailand",
                passportNumber: "31195855",
                visaNumber: "73203393",
                gender: "Female",
                day: 30,
                month: 8,
                year: 1985,
                flightNumber: "(KL) KLM 876");
            SaveDocument(outputStream, Path.Combine(outputDirectory, "Document.pdf"));

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

    static Stream FillForm(MemoryStream documentStream, string lastName, string firstName,
        string address, string nationality, string passportNumber, string visaNumber,
        string gender, int day, int month, int year, string flightNumber) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        var documentFacade = processor.DocumentFacade;
        var form = documentFacade.AcroForm;

        form.GetTextFormField("LastName").Value = lastName;
        form.GetTextFormField("FirstName").Value = firstName;
        form.GetTextFormField("Address").Value = address;
        form.GetComboBoxFormField("Nationality").Value = nationality;
        form.GetTextFormField("PassportNo").Value = passportNumber;
        form.GetTextFormField("VisaNo").Value = visaNumber;
        form.GetRadioGroupFormField("Gender").Value = gender;
        form.GetTextFormField("DD").Value = day.ToString();
        form.GetTextFormField("MM").Value = month.ToString();
        form.GetTextFormField("YYYY").Value = year.ToString();
        form.GetTextFormField("FlightNo").Value = flightNumber;

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
