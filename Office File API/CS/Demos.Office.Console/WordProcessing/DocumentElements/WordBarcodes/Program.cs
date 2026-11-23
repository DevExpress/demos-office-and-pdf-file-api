using DevExpress.BarCodes;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Barcode_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const Symbology BarcodeType = Symbology.Code128;
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

    static Stream GenerateDocumentWithBarcode(string barcodeValue, Symbology barcodeSymbology, DocumentFormat outputFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        var document = wordProcessor.Document;

        var barCode = new BarCode {
            Symbology = barcodeSymbology,
            CodeText = barcodeValue
        };

        var shape = document.Shapes.InsertPicture(document.Range.End, DocumentImageSource.FromImage(barCode.BarCodeImage));
        shape.TextWrapping = TextWrappingType.InLineWithText;

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
