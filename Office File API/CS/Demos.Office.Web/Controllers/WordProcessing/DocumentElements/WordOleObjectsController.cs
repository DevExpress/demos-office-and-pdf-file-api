using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code


namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentElements {
    [ApiController]
    [Route("api/wordprocessing/ole-objects")]
    public class WordOleObjectsController : WordControllerBase {
        [HttpPost("generate")]
        public async Task<IActionResult> GenerateOleObjectsDocument([FromForm] WordOleObjectsModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordOleExcelTemplateFile, FileMode.Open, FileAccess.Read);
            var oleObjectStream = new MemoryStream();
            templateStream.CopyTo(oleObjectStream);
            oleObjectStream.Position = 0;

            using var imageTemplateStream = new FileStream(DemoDocumentPaths.WordOleExcelIcon, FileMode.Open, FileAccess.Read);
            var imageStream = new MemoryStream();
            imageTemplateStream.CopyTo(imageStream);
            imageStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with OLE objects
            var outputStream = await GenerateDocumentWithOleObject(oleObjectStream, OleObjectType.ExcelWorksheet, imageStream, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        async Task<Stream> GenerateDocumentWithOleObject(Stream oleObjectStream, string objectType, Stream imageStream, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            var oleShape = document.Shapes.InsertOleObjectAsIcon(document.Range.Start, oleObjectStream, objectType, DocumentImageSource.FromStream(imageStream));
            oleShape.TextWrapping = TextWrappingType.InLineWithText;

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        #endregion Code
    }
}
