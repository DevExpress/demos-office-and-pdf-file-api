using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.BarCodes;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentElements {
    [ApiController]
    [Route("api/wordprocessing/barcodes")]
    public class WordBarcodesController : WordControllerBase {
        [HttpPost("generate")]
        public async Task<IActionResult> GenerateDocument([FromForm] WordBarcodesModel model) {
            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with chart
            var outputStream = GenerateDocumentWithBarcode(model.BarcodeValue, model.BarcodeType, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream GenerateDocumentWithBarcode(string barcodeValue, Symbology barcodeSymbology, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            BarCode barCode = new() {
                Symbology = barcodeSymbology,
                CodeText = barcodeValue,
            };
            var shape = document.Shapes.InsertPicture(document.Range.End, DocumentImageSource.FromImage(barCode.BarCodeImage));
            shape.TextWrapping = TextWrappingType.InLineWithText;

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
