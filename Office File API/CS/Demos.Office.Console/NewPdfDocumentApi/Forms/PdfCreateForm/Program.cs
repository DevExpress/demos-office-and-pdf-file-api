using System.Drawing;
using DevExpress.Docs.Pdf;
using DevExpress.Drawing.Printing;

internal static class Program {
    const string OutputFileName = "CreatedForm.pdf";
    const bool FillFromFile = true;
    static readonly string FormDataFilePath = Path.Combine("Documents", "PDF", "Sample_FormData.xml");

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var xmlStream = FillFromFile ? File.OpenRead(Path.Combine(AppContext.BaseDirectory, FormDataFilePath)) : null;
            using var outputStream = CreateFormDocument(FillFromFile, xmlStream);
            SaveDocument(outputStream, outputPath);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static void SaveDocument(Stream stream, string outputPath) {
        stream.Position = 0;
        using var fileStream = File.Create(outputPath);
        stream.CopyTo(fileStream);
    }

    static Stream CreateFormDocument(bool fillFromFile, Stream? xmlStream) {
        using var doc = new PdfDocument();
        var page = doc.Pages.Add(DXPaperKind.A4);
        float pageW = page.CropBox.Width;
        float pageH = page.CropBox.Height;
        float left = 50f, right = 545f, width = right - left;

        void AddTextField(string name, RectangleF rect, bool multiline = false) {
            var field = new TextBoxField(name) { Multiline = multiline };
            var widget = new TextBoxWidgetAnnotation(field, rect) { BackgroundColor = PdfColor.White };
            page.Annotations.Add(widget);
            doc.Fields.Add(field);
        }

        void AddComboBoxField(string name, RectangleF rect, params string[] options) {
            var field = new ComboBoxField(name) {
                Items = options.Select(option => new ChoiceFieldItem(option)).ToList()
            };
            var widget = new ComboBoxWidgetAnnotation(field, rect) { BackgroundColor = PdfColor.White };
            page.Annotations.Add(widget);
            doc.Fields.Add(field);
        }

        page.AddFragment(PathFragment.Rectangle(0, pageH - 62, pageW, 62, SolidFill.DarkBlue));
        page.AddFragment(new TextFragment {
            Text = "PASSENGER INFORMATION FORM",
            Location = new PointF(left, pageH - 26),
            ForegroundFill = SolidFill.White,
        });

        page.AddFragment(PathFragment.Rectangle(left, 722, width, 18, SolidFill.LightGray));
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
        page.Annotations.Add(new RadioGroupItemWidgetAnnotation(genderField, "Male", new RectangleF(left + 245, 640, 14, 14)) { ButtonStyle = genderField.ButtonStyle });
        page.AddTextFragment("Male", left + 262, 642);
        page.Annotations.Add(new RadioGroupItemWidgetAnnotation(genderField, "Female", new RectangleF(left + 300, 640, 14, 14)) { ButtonStyle = genderField.ButtonStyle });
        page.AddTextFragment("Female", left + 317, 642);
        doc.Fields.Add(genderField);

        page.AddTextFragment("Nationality:", left, 621);
        AddComboBoxField("Nationality", new RectangleF(left, 598, 250, 18), "American", "British", "Spanish", "Cypriot", "Hungarian", "Armenian");

        page.AddFragment(PathFragment.Rectangle(left, 568, width, 18, SolidFill.LightGray));
        page.AddTextFragment("Travel Documents", left + 3, 572);
        page.AddTextFragment("Passport No:", left, 553);
        page.AddTextFragment("Visa No:", left + 245, 553);
        AddTextField("PassportNo", new RectangleF(left, 530, 200, 18));
        AddTextField("VisaNo", new RectangleF(left + 245, 530, 200, 18));
        page.AddTextFragment("Flight No:", left, 513);
        AddTextField("FlightNo", new RectangleF(left, 490, 250, 18));

        page.AddFragment(PathFragment.Rectangle(left, 450, width, 18, SolidFill.LightGray));
        page.AddTextFragment("Address", left + 3, 454);
        page.AddTextFragment("Address:", left, 430);
        AddTextField("Address", new RectangleF(left, 380, width, 46), multiline: true);

        if(fillFromFile && xmlStream != null)
            doc.ImportFormData(xmlStream, ExportDataFormat.Xml);

        var stream = new MemoryStream();
        doc.Save(stream);
        stream.Position = 0;
        return stream;
    }
}
