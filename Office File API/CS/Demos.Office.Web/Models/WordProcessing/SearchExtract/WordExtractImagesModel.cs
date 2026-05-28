using DevExpress.Drawing;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.SearchExtract
{
    public class WordExtractImagesModel : WordFileModelBase
    {
        public DXImageFormat ImageFormat { get; set; } = DXImageFormat.Png;
        public override string OutputName { get; set; } = "Document_Images";
        public bool ExportCharts { get; set; } = true;
        internal readonly List<FormatInfo> ImageFormats = new()
        {
            new FormatInfo { Text = "PNG (Portable Network Graphics)", Value = DXImageFormat.Png },
            new FormatInfo { Text = "JPEG (Joint Photographic Experts Group)", Value = DXImageFormat.Jpeg },
            new FormatInfo { Text = "BMP (Bitmap)", Value = DXImageFormat.Bmp }
        };
    }
}
