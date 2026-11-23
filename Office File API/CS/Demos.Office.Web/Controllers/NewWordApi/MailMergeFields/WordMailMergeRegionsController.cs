#region Code

using DevExpress.Docs.Word;
using System.Text.Json;
        #endregion Code

using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.MailMergeFields;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.MailMergeFields {
    [ApiController]
    [Route("api/newwordapi/mail-merge-regions")]
    public class WordMailMergeRegionsController : WordControllerBase {
        [HttpPost("process")]
        public async Task<IActionResult> CreateMailMergedDocument(IFormFile file, [FromForm] WordMailMergeRegionsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Process track changes on the template
            var outputStreams = MailMergeDocument(inputStream, model.MergeMode, model.RecordExportMode == RecordExportMode.ActiveRecord, model.ActiveRecordIndex, model.DocumentExportMode == DocumentExportMode.SingleFile, model.IsPdfOutput, model.OutputFormat);

            return CreateZipOutputFile(outputStreams, model.OutputFormat, "record", model.OutputName);
        }

        #region Code
        List<Stream> MailMergeDocument(Stream stream, MergeMode mergeMode, bool exportActiveRecord, int activeRecordIndex,
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

        List<NWind.Customer> PrepareDataSource() {
            try {
                using var dsStream = new FileStream(DemoDocumentPaths.NWindDataSource, FileMode.Open, FileAccess.Read);
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
                document.Save(outputStream, NewWordFormatHelper.GetDocumentFormat(outputFormat));

            outputStream.Position = 0;
            return outputStream;
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
        #endregion Code
    }
}