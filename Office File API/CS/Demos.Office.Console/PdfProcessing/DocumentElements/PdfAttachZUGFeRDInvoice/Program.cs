using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_A3b.pdf");
    static readonly string InvoiceFilePath = Path.Combine("Documents", "PDF", "Sample_Invoice.xml");
    const string OutputFileName = "Document_with_ZUGFeRD.pdf";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var invoicePath = Path.Combine(AppContext.BaseDirectory, InvoiceFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var invoiceStream = File.OpenRead(invoicePath);
            using var outputStream = AttachZugferd(documentStream, invoiceStream);
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

    static Stream AttachZugferd(Stream documentStream, Stream invoiceStream) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);
        processor.Document.AttachZugferdInvoice(invoiceStream);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
