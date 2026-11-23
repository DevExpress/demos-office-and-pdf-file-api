using DevExpress.Docs.Barcode;
using DevExpress.Docs.Word;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Bracode_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const BarcodeSymbology BarcodeType = BarcodeSymbology.Code128;
    const string BarcodeValue = "123456789012";

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var outputStream = GenerateDocumentWithBarcode(BarcodeValue, BarcodeType, OutputFormat);
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

    static Stream GenerateDocumentWithBarcode(string barcodeValue, BarcodeSymbology barcodeSymbology, DocumentFormat targetFormat) {
        using var document = new WordDocument();

        BarcodeOptions barcodeOptions = barcodeSymbology switch {
            BarcodeSymbology.QRCode => new QRCodeOptions(),
            BarcodeSymbology.DataBar => new DataBarOptions(),
            BarcodeSymbology.Code128 => new Code128Options(),
        };
        using var barcodeGenerator = new BarcodeGenerator(barcodeOptions);
        var picture = new PictureShape(barcodeGenerator.ExportToImage(barcodeValue));
        document.Body.AddParagraph().InlineElements.Add(picture);

        var outputStream = new MemoryStream();
        document.Save(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}

enum BarcodeSymbology {
    Code128,
    QRCode,
    DataBar
}