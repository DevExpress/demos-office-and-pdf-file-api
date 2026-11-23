using System.Text.Json;
using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "MailMergeMasterDetail_Template.docx");
    static readonly string DataSourcePath = Path.Combine("Documents", "nwind_data.json");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "MailMerge_Result.pdf";
    const string OutputFormat = "pdf";
    const bool IsPdfOutput = true;
    const MergeMode DefaultMergeMode = MergeMode.NewSection;
    const RecordExportMode DefaultRecordExportMode = RecordExportMode.AllRecords;
    const DocumentExportMode DefaultDocumentExportMode = DocumentExportMode.SingleFile;
    const int ActiveRecordIndex = 0;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            using var inputStream = File.OpenRead(inputPath);

            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var outputStreams = MailMergeDocument(
                inputStream,
                DefaultMergeMode,
                DefaultRecordExportMode == RecordExportMode.ActiveRecord,
                ActiveRecordIndex,
                DefaultDocumentExportMode == DocumentExportMode.SingleFile,
                IsPdfOutput,
                OutputFormat);

            if(outputStreams.Count == 0)
                throw new InvalidOperationException("No output documents were created.");

            if(outputStreams.Count == 1) {
                var outputPath = Path.Combine(outputDirectory, OutputFileName);
                using var fileStream = File.Create(outputPath);
                using var outputStream = outputStreams[0];
                outputStream.CopyTo(fileStream);
                Console.WriteLine($"Created {outputPath}");
                return 0;
            }

            for(var index = 0; index < outputStreams.Count; index++) {
                var extension = IsPdfOutput ? "pdf" : OutputFormat;
                var outputPath = Path.Combine(outputDirectory, $"record_{index + 1}.{extension}");
                using var fileStream = File.Create(outputPath);
                using var outputStream = outputStreams[index];
                outputStream.CopyTo(fileStream);
            }

            Console.WriteLine($"Created {outputDirectory}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static List<Stream> MailMergeDocument(Stream stream, MergeMode mergeMode, bool exportActiveRecord, int activeRecordIndex,
                                          bool exportAsOneDocument, bool isPdfOutput, string outputFormat) {
        using var document = new WordDocument(stream);

        var dataSource = PrepareDataSource();

        var firstRecordIndex = exportActiveRecord ? activeRecordIndex : 0;
        var lastRecordIndex = exportActiveRecord ? activeRecordIndex : dataSource.Count - 1;

        var options = new MailMergeOptions {
            DataSource = dataSource,
            MergeMode = mergeMode,
            FirstRecordIndex = firstRecordIndex,
            LastRecordIndex = lastRecordIndex
        };

        var outputStreams = new List<Stream>();

        if(exportAsOneDocument) {
            using var mergedDocument = document.MailMerge(options);
            outputStreams.Add(SaveDocument(mergedDocument, isPdfOutput, outputFormat));
        }
        else
            document.MailMerge(options, new RecordDocumentConsumer(outputStreams, isPdfOutput, outputFormat));

        return outputStreams;
    }

    static List<NWind.Customer> PrepareDataSource() {
        try {
            var dataSourcePath = Path.Combine(AppContext.BaseDirectory, DataSourcePath);
            using var dsStream = new FileStream(dataSourcePath, FileMode.Open, FileAccess.Read);
            var nwind = JsonSerializer.Deserialize<NWind>(dsStream);
            return nwind?.Customers ?? new List<NWind.Customer>();
        }
        catch(Exception ex) {
            Console.Error.WriteLine($"Deserialization failed: {ex}");
            return new List<NWind.Customer>();
        }
    }

    static Stream SaveDocument(WordDocument document, bool isPdfOutput, string outputFormat) {
        var outputStream = new MemoryStream();

        if(isPdfOutput)
            document.ExportToPdf(outputStream);
        else
            document.Save(outputStream, GetDocumentFormat(outputFormat));

        outputStream.Position = 0;
        return outputStream;
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "doc" => DocumentFormat.Doc,
            "docx" => DocumentFormat.Docx,
            "docm" => DocumentFormat.Docm,
            "dotx" => DocumentFormat.Dotx,
            "dotm" => DocumentFormat.Dotm,
            "rtf" => DocumentFormat.Rtf,
            _ => DocumentFormat.Docx
        };
    }

    class RecordDocumentConsumer : IMailMergeDocumentConsumer {
        readonly List<Stream> outputStreams;
        readonly bool isPdfOutput;
        readonly string outputFormat;

        public RecordDocumentConsumer(List<Stream> outputStreams, bool isPdfOutput, string outputFormat) {
            this.outputStreams = outputStreams;
            this.isPdfOutput = isPdfOutput;
            this.outputFormat = outputFormat;
        }

        public void Consume(int recordIndex, WordDocument document) {
            using(document) {
                outputStreams.Add(SaveDocument(document, isPdfOutput, outputFormat));
            }
        }
    }

    class NWind {
        public List<Customer> Customers { get; set; } = new();

        public record Customer(
            string CustomerId,
            string CompanyName,
            string ContactName,
            string Country,
            string Address,
            string City,
            string Phone,
            List<Order> Orders);

        public record Order(
            string OrderID,
            DateTime? OrderDate,
            string ShipCountry,
            double Freight,
            List<OrderDetails> OrderDetails);

        public record OrderDetails(
            int ProductId,
            string ProductName,
            int Quantity,
            decimal UnitPrice,
            double Discount);
    }

    enum RecordExportMode {
        AllRecords,
        ActiveRecord
    }

    enum DocumentExportMode {
        SeparateFiles,
        SingleFile
    }
}