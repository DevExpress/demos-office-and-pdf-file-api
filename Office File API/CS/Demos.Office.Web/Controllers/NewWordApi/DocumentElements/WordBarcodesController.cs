using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Barcode;
using DevExpress.Docs.Word;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentElements {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordBarcodesController : WordControllerBase {
        [HttpPost("barcodes/generate")]
        public IActionResult GenerateDocument([FromForm] WordBarcodesModel model) {
            // Determine target format
            var targetFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with chart
            var outputStream = GenerateDocumentWithBarcode(model.BarcodeValue, model.BarcodeType, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream GenerateDocumentWithBarcode(string barcodeValue, BarcodeSymbology barcodeSymbology, DocumentFormat targetFormat) {
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
        #endregion Code
    }
}