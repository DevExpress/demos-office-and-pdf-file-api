using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentElements {
    [ApiController]
    [Route("api/wordprocessing/content-controls")]
    public class WordContentControlsController : WordControllerBase {

        [HttpPost("process")]
        public async Task<IActionResult> ProcessContentControls(IFormFile file, [FromForm] WordContentControlsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process content controls with appointment form data
            var outputStream = await ProcessContentControlsDocument(inputStream, model.AppointmentForm, model.IsPdfOutput(), model.ExportAsAcroFormFields, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        async Task<Stream> ProcessContentControlsDocument(Stream inputStream, AppointmentForm appointmentForm,
            bool isPdfOutput, bool exportAsAcroForm, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            PopulateContentControls(document, appointmentForm);

            var outputStream = new MemoryStream();

            if (isPdfOutput) {
                DevExpress.XtraPrinting.PdfExportOptions options = new DevExpress.XtraPrinting.PdfExportOptions
                {
                    ExportEditingFieldsToAcroForms = exportAsAcroForm
                };
                wordProcessor.ExportToPdf(outputStream, options);
            }
            else
                wordProcessor.SaveDocument(outputStream, targetFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        void PopulateContentControls(Document document, AppointmentForm appointmentForm) {
            foreach (var control in document.ContentControls) {
                switch (control.Tag) {
                    case "PatientName":
                        UpdateControlText(document, (ContentControlPlainText)control, appointmentForm.PatientName);
                        break;
                    case "DateOfBirth":
                        ((ContentControlDate)control).Date = appointmentForm.DateOfBirth;
                        break;
                    case "PhoneNumber":
                        UpdateControlText(document, (ContentControlPlainText)control, appointmentForm.PhoneNumber);
                        break;
                    case "EmailAddress":
                        UpdateControlText(document, (ContentControlPlainText)control, appointmentForm.EmailAddress);
                        break;
                    case "AppointmentDate":
                        ((ContentControlDate)control).Date = appointmentForm.AppointmentDate;
                        break;
                    case "AppointmentTime":
                        UpdateControlText(document, (ContentControlPlainText)control, appointmentForm.AppointmentTime);
                        break;
                    case "DoctorName":
                        UpdateControlText(document, (ContentControlPlainText)control, appointmentForm.DoctorName);
                        break;
                    case "VisitReason":
                        ((ContentControlComboBox)control).SelectedItemIndex = (int)appointmentForm.VisitReason;
                        break;
                    case "AdditionalNotes":
                        UpdateControlText(document, (ContentControlRichText)control, appointmentForm.AdditionalNotes ?? string.Empty);
                        break;
                    case "WheelchairAccess":
                        ((ContentControlCheckbox)control).Checked = appointmentForm.WheelchairAccess;
                        break;
                    case "LanguageInterpreter":
                        ((ContentControlCheckbox)control).Checked = appointmentForm.LanguageInterpreter;
                        break;
                    case "OtherRequirements":
                        ((ContentControlCheckbox)control).Checked = appointmentForm.OtherRequirements;
                        break;
                    case "OtherRequirement":
                        UpdateControlText(document, (ContentControlPlainText)control, appointmentForm.OtherRequirement ?? string.Empty);
                        break;
                }
            }
        }

        void UpdateControlText(Document document, ContentControlBase control, string text) {
            document.Delete(document.CreateRange(control.Range.Start.ToInt() + 1, control.Range.Length - 2));
            document.InsertText(document.CreatePosition(control.Range.Start.ToInt() + 1), text);
        }
        #endregion Code
    }
}
