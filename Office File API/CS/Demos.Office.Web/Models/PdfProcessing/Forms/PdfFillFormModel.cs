using System;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.Forms {
    public class PdfFillFormModel : PdfModelBase {
        [Required]
        [StringLength(30)]
        public string LastName { get; set; } = "Leverling";
        [Required]
        [StringLength(30)]
        public string FirstName { get; set; } = "Janet";
        [Required]
        [StringLength(30)]
        public string Nationality { get; set; } = "Thailand";
        [Required]
        [StringLength(50)]
        public string Address { get; set; } = "98033, 722 Moss Bay Blvd., Kirkland, WA, USA";
        [Required]
        [RegularExpression(@"^[A-Za-z0-9\-]{4,20}$", ErrorMessage = "Passport number must be 4-20 alphanumeric characters or hyphens.")]
        public string PassportNumber { get; set; } = "31195855";
        [Required]
        [RegularExpression(@"^[A-Za-z0-9\-]{4,20}$", ErrorMessage = "Visa number must be 4-20 alphanumeric characters or hyphens.")]
        public string VisaNumber { get; set; } = "73203393";
        [Required]
        [StringLength(20)]
        public string Gender { get; set; } = "Female";
        public int Day { get; set; } = 30;
        public int Month { get; set; } = 8;
        public int Year { get; set; } = 1985;
        internal DateTime DateOfBirth = new DateTime(1985, 8, 30);
        [Required]
        [StringLength(30)]
        public string FlightNumber { get; set; } = "(KL) KLM 876";
    }
}
