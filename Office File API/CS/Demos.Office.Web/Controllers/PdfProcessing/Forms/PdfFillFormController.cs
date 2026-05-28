using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.Forms;
#region Code
using DevExpress.Pdf;
using DevExpress.Drawing;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.Forms {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfFillFormController : PdfControllerBase {
        [HttpPost("pdf-form/fill")]
        public async Task<IActionResult> FillPdfForm(IFormFile file, [FromForm] PdfFillFormModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = FillForm(documentStream, model.LastName, model.FirstName, model.Address, model.Nationality, model.PassportNumber, model.VisaNumber, model.Gender, model.Day, model.Month, model.Year, model.FlightNumber);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FillForm(MemoryStream documentStream, string lastName, string firstName,
            string address, string nationality, string passportNumber, string visaNumber,
            string gender, int day, int month, int year, string flightNumber) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var documentFacade = processor.DocumentFacade;
            var form = documentFacade.AcroForm;

            form.GetTextFormField("LastName").Value = lastName;
            form.GetTextFormField("FirstName").Value = firstName;
            form.GetTextFormField("Address").Value = address;
            form.GetComboBoxFormField("Nationality").Value = nationality;
            form.GetTextFormField("PassportNo").Value = passportNumber;
            form.GetTextFormField("VisaNo").Value = visaNumber;
            form.GetRadioGroupFormField("Gender").Value = gender;
            form.GetTextFormField("DD").Value = day.ToString();
            form.GetTextFormField("MM").Value = month.ToString();
            form.GetTextFormField("YYYY").Value = year.ToString();
            form.GetTextFormField("FlightNo").Value = flightNumber;

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion
    }
}
