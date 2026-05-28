
namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    public class ExcelOleObjectsModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "OLE_Document";
        public bool DisplayAsIcon { get; set; } = true;
    }
}
