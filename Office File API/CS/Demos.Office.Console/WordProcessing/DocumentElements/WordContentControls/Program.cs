using PdfExportOptions = DevExpress.XtraPrinting.PdfExportOptions;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "ContentControls_AppointmentForm.docx");
    const string OutputFileName = "AppointmentForm_Filled.pdf";
    const bool ExportAsAcroFormFields = true;

    static readonly AppointmentForm FormData = new() {
        PatientName = "John Doe",
        DateOfBirth = new DateTime(1990, 1, 15),
        PhoneNumber = "(198) 555-8888",
        EmailAddress = "johndoe@examplemail.com",
        AppointmentDate = DateTime.Today.AddDays(7),
        AppointmentTime = "10:00 AM",
        DoctorName = "Dr. Alice Smith",
        VisitReason = VisitReason.AnnualPhysicalExamination,
        AdditionalNotes = "N/A",
        WheelchairAccess = false,
        LanguageInterpreter = false,
        OtherRequirements = false,
        OtherRequirement = string.Empty
    };

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProcessContentControlsDocument(inputStream, FormData, ExportAsAcroFormFields);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);
            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream ProcessContentControlsDocument(Stream inputStream, AppointmentForm appointmentForm, bool exportAsAcroForm) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;

        PopulateContentControls(document, appointmentForm);

        var outputStream = new MemoryStream();
        var options = new PdfExportOptions {
            ExportEditingFieldsToAcroForms = exportAsAcroForm
        };
        wordProcessor.ExportToPdf(outputStream, options);
        outputStream.Position = 0;
        return outputStream;
    }

    static void PopulateContentControls(Document document, AppointmentForm appointmentForm) {
        foreach(var control in document.ContentControls) {
            switch(control.Tag) {
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

    static void UpdateControlText(Document document, ContentControlBase control, string text) {
        document.Delete(document.CreateRange(control.Range.Start.ToInt() + 1, control.Range.Length - 2));
        document.InsertText(document.CreatePosition(control.Range.Start.ToInt() + 1), text);
    }

    class AppointmentForm {
        public string PatientName { get; set; } = string.Empty;
        public DateTime DateOfBirth { get; set; }
        public string PhoneNumber { get; set; } = string.Empty;
        public string EmailAddress { get; set; } = string.Empty;
        public DateTime AppointmentDate { get; set; }
        public string AppointmentTime { get; set; } = string.Empty;
        public string DoctorName { get; set; } = string.Empty;
        public VisitReason VisitReason { get; set; }
        public string AdditionalNotes { get; set; } = string.Empty;
        public bool WheelchairAccess { get; set; }
        public bool LanguageInterpreter { get; set; }
        public bool OtherRequirements { get; set; }
        public string OtherRequirement { get; set; } = string.Empty;
    }

    enum VisitReason {
        AnnualPhysicalExamination,
        FollowUpVisit,
        SpecificConcern
    }
}
