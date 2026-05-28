using BlazorDemo.OfficeFileApi.Models.PdfProcessing.Forms;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.Forms {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfRemoveFlattenFormController : PdfControllerBase {
        [HttpPost("pdf-form/flatten")]
        public async Task<IActionResult> FlattenPdfForm(IFormFile file, [FromForm] PdfRemoveFlattenFormModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = FlattenForm(documentStream, model.ProcessAll, model.FieldsToProcess.Split(";"));
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FlattenForm(MemoryStream documentStream, bool flattenAll, string[] fieldsToFlatten) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            if (flattenAll)
                processor.FlattenForm();
            else
                foreach (var field in fieldsToFlatten)
                    processor.FlattenFormField(field);

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion
        [HttpPost("pdf-form/remove")]
        public async Task<IActionResult> RemovePdfForm(IFormFile file, [FromForm] PdfRemoveFlattenFormModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = RemoveForm(documentStream, model.ProcessAll, model.FieldsToProcess.Split(";"));
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream RemoveForm(MemoryStream documentStream, bool removeAll, string[] fieldsToRemove) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            if(removeAll)
                processor.RemoveForm();
            else
                foreach(var field in fieldsToRemove)
                    processor.RemoveFormField(field);

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion
    }
}
