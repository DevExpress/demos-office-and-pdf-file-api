namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.Forms {
    public class PdfCreateFormModel : PdfModelBase {
        public bool CreateTextBox { get; set; } = true;
        public bool CreateComboBox { get; set; } = true;
        public bool CreateListBox { get; set; } = true;
        public bool CreateCheckBox { get; set; } = true;
        public bool CreateRadioGroup { get; set; } = true;
    }
}
