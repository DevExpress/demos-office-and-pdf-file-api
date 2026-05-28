using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements
{
    public class WordContentControlsModel : WordFileModelBase
    {
        public AppointmentForm AppointmentForm { get; set; } = new AppointmentForm()
        {
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
            OtherRequirement = ""
        };

        [Display(Name = "Export as AcroForm Fields")]
        public bool ExportAsAcroFormFields { get; set; } = true;

        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "PDF Document (.pdf)", Value = "pdf" },
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" }
        };

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".docx", ".docm", ".dotx", ".dotm", ".pdf" };
        public override string OutputName { get; set; } = "AppointmentForm_Filled";
        public override string OutputFormat { get; set; } = "pdf";

        public bool IsPdfOutput()
        {
            return OutputFormat?.ToLowerInvariant() == "pdf";
        }

        public bool IsWordOutput()
        {
            var wordFormats = new[] { "docx", "docm", "dotx", "dotm" };
            return wordFormats.Contains(OutputFormat?.ToLowerInvariant());
        }

        internal readonly List<object> AvailableDoctors = new()
        {
            new { Text = "Dr. Alice Smith - Family Medicine", Value = "Dr. Alice Smith" },
            new { Text = "Dr. Sarah Johnson - General Practice", Value = "Dr. Sarah Johnson" },
            new { Text = "Dr. Michael Chen - Cardiology", Value = "Dr. Michael Chen" },
            new { Text = "Dr. Emily Rodriguez - Pediatrics", Value = "Dr. Emily Rodriguez" },
            new { Text = "Dr. David Thompson - Orthopedics", Value = "Dr. David Thompson" },
            new { Text = "Dr. Lisa Williams - Dermatology", Value = "Dr. Lisa Williams" },
            new { Text = "Dr. Robert Brown - Internal Medicine", Value = "Dr. Robert Brown" }
        };
        internal readonly List<object> VisitReasons = new()
        {
            new { Text = "Annual Physical Examination", Value = VisitReason.AnnualPhysicalExamination },
            new { Text = "Follow-up Visit", Value = VisitReason.FollowUpVisit },
            new { Text = "Specific Concern", Value = VisitReason.SpecificConcern }
        };
        internal readonly List<object> AppointmentTimes = GenerateAppointmentTimes();
        internal static List<object> GenerateAppointmentTimes()
        {
            var times = new List<object>();
            for (int hour = 8; hour <= 17; hour++)
            {
                for (int minute = 0; minute < 60; minute += 30)
                {
                    var time = new TimeSpan(hour, minute, 0);
                    var timeString = time.ToString(@"hh\:mm");
                    var period = hour < 12 ? "AM" : "PM";
                    var displayHour = hour == 0 ? 12 : hour > 12 ? hour - 12 : hour;
                    var displayTime = $"{displayHour:D2}:{minute:D2} {period}";

                    times.Add(new { Text = displayTime, Value = $"{displayHour:D2}:{minute:D2} {period}" });
                }
            }
            return times;
        }
    }
}
