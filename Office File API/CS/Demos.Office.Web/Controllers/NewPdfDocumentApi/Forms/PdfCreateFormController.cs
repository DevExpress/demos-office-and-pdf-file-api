using System.Drawing;
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.Forms;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
using DevExpress.Drawing.Printing;

#region Code
using System.Collections.Generic;
using System.Linq;
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.Forms {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfCreateFormController : PdfControllerBase {
        [HttpPost("pdf-form/create")]
        public async Task<IActionResult> CreateForm([FromForm] PdfCreateFormModel model, IFormFile xmlFile) {
            Stream xmlStream = null;
            if(model.FillFromFile && xmlFile != null) {
                xmlStream = new MemoryStream();
                await xmlFile.CopyToAsync(xmlStream);
                xmlStream.Position = 0;
            }

            Stream outputStream = CreateFormDocument(model.FillFromFile, xmlStream);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream CreateFormDocument(bool fillFromFile, Stream xmlStream) {
            using var doc = new PdfDocument();
            var page = doc.Pages.Add(DXPaperKind.A4);
            float pageW = page.CropBox.Width;
            float pageH = page.CropBox.Height;
            float left = 50f, right = 545f, w = right - left;

            // ── helpers ──────────────────────────────────────────────────────
            void AddTextField(string name, RectangleF rect, bool multiline = false) {
                var field = new TextBoxField(name) { Multiline = multiline };
                var widget = new TextBoxWidgetAnnotation(field, rect) { BackgroundColor = PdfColor.White };
                page.Annotations.Add(widget);
                doc.Fields.Add(field);
            }

            void AddComboBoxField(string name, RectangleF rect, params string[] options) {
                var field = new ComboBoxField(name) {
                    Items = options.Select(o => new ChoiceFieldItem(o)).ToList()
                };
                var widget = new ComboBoxWidgetAnnotation(field, rect) { BackgroundColor = PdfColor.White };
                page.Annotations.Add(widget);
                doc.Fields.Add(field);
            }

            // ── header bar ───────────────────────────────────────────────────
            page.AddFragment(PathFragment.Rectangle(0, pageH - 62, pageW, 62, SolidFill.DarkBlue));
            page.AddFragment(new TextFragment {
                Text = "PASSENGER INFORMATION FORM",
                Location = new PointF(left, pageH - 26),
                ForegroundFill = SolidFill.White,
            });

            // ── Personal Information section ──────────────────────────────────
            page.AddFragment(PathFragment.Rectangle(left, 722, w, 18, SolidFill.LightGray));
            page.AddTextFragment("Personal Information", left + 3, 726);

            page.AddTextFragment("First Name:", left, 706);
            page.AddTextFragment("Last Name:", left + 245, 706);
            AddTextField("FirstName", new RectangleF(left, 683, 200, 18));
            AddTextField("LastName", new RectangleF(left + 245, 683, 200, 18));

            page.AddTextFragment("Date of Birth:", left, 661);
            page.AddTextFragment("Gender:", left + 245, 661);
            AddTextField("MM", new RectangleF(left, 638, 38, 18));
            page.AddTextFragment("/", left + 41, 642);
            AddTextField("DD", new RectangleF(left + 48, 638, 38, 18));
            page.AddTextFragment("/", left + 89, 642);
            AddTextField("YYYY", new RectangleF(left + 96, 638, 52, 18));

            var genderField = new RadioGroupField("Gender") { ButtonStyle = FormFieldButtonStyle.Check };
            page.Annotations.Add(new RadioGroupItemWidgetAnnotation(genderField, "Male",
                new RectangleF(left + 245, 640, 14, 14)) { ButtonStyle = genderField.ButtonStyle });
            page.AddTextFragment("Male", left + 262, 642);
            page.Annotations.Add(new RadioGroupItemWidgetAnnotation(genderField, "Female",
                new RectangleF(left + 300, 640, 14, 14)) { ButtonStyle = genderField.ButtonStyle });
            page.AddTextFragment("Female", left + 317, 642);
            doc.Fields.Add(genderField);

            page.AddTextFragment("Nationality:", left, 621);
            AddComboBoxField("Nationality", new RectangleF(left, 598, 250, 18),
                "American", "British", "Spanish", "Cypriot", "Hungarian", "Armenian");

            // ── Travel Documents section ──────────────────────────────────────
            page.AddFragment(PathFragment.Rectangle(left, 568, w, 18, SolidFill.LightGray));
            page.AddTextFragment("Travel Documents", left + 3, 572);

            page.AddTextFragment("Passport No:", left, 553);
            page.AddTextFragment("Visa No:", left + 245, 553);
            AddTextField("PassportNo", new RectangleF(left, 530, 200, 18));
            AddTextField("VisaNo", new RectangleF(left + 245, 530, 200, 18));

            page.AddTextFragment("Flight No:", left, 513);
            AddTextField("FlightNo", new RectangleF(left, 490, 250, 18));

            // ── Address section ───────────────────────────────────────────────
            page.AddFragment(PathFragment.Rectangle(left, 450, w, 18, SolidFill.LightGray));
            page.AddTextFragment("Address", left + 3, 454);

            page.AddTextFragment("Address:", left, 430);
            AddTextField("Address", new RectangleF(left, 380, w, 46), multiline: true);

            if (fillFromFile && xmlStream != null) {
                doc.ImportFormData(xmlStream, ExportDataFormat.Xml);
            }
            var stream = new MemoryStream();
            doc.Save(stream);
            return stream;
        }
        #endregion
    }
}
