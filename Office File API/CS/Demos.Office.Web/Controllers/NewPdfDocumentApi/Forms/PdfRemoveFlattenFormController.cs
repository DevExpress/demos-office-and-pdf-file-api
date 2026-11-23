using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.Forms;
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.Forms {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
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
        Stream FlattenForm(Stream documentStream, bool flattenAll, string[] fieldsToFlatten) {
            using var doc = new PdfDocument(documentStream);

            if(flattenAll)
                doc.FlattenFormFields();
            else
                doc.FlattenFormFields(fieldsToFlatten);

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
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
        Stream RemoveForm(Stream documentStream, bool removeAll, string[] fieldsToRemove) {
            using var doc = new PdfDocument(documentStream);

            if(removeAll)
                doc.Fields.Clear();
            else
                foreach(var fieldName in fieldsToRemove) {
                    var field = doc.Fields.FindByName(fieldName);
                    if(field != null)
                        doc.Fields.Remove(field);
                }

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
    }
}
