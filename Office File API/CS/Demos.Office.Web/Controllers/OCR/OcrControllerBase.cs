using System.Drawing;
using System.Text.Json;
using BlazorDemo.OfficeFileApi.Models.OCR.Recognition;
using DevExpress.Docs.Ocr;
using Microsoft.AspNetCore.Mvc;

namespace BlazorDemo.OfficeFileApi.Controllers.OCR {
    public abstract class OcrControllerBase : FileStreamControllerBase {
        static readonly JsonSerializerOptions JsonOptions = new() {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        public override string GetMimeContentType(string format) {
            return format.ToLowerInvariant() switch {
                "txt" => "text/plain",
                "json" => "application/json",
                "pdf" => "application/pdf",
                _ => base.GetMimeContentType(format)
            };
        }

        protected static string Serialize(object value) {
            return JsonSerializer.Serialize(value, JsonOptions);
        }

        protected static string ToText(IEnumerable<ImageTextFragment> fragments) {
            return WithFallback(string.Join(Environment.NewLine, fragments.Select(f => f.Text)));
        }

        protected static string ToText(IReadOnlyList<PdfPageRecognitionResult> pages) {
            if(pages.Count == 1)
                return WithFallback(string.Join(Environment.NewLine, pages[0].TextFragments.Select(f => f.Text)));
            return WithFallback(string.Join(Environment.NewLine + Environment.NewLine, pages.Select(p =>
                $"--- Page {p.PageIndex + 1} ---{Environment.NewLine}" +
                string.Join(Environment.NewLine, p.TextFragments.Select(f => f.Text)))));
        }

        static string WithFallback(string text) {
            return text.Length > 0 ? text : "Nothing found";
        }

        protected static IEnumerable<OcrFragment> ToFragments(IEnumerable<ImageTextFragment> fragments) {
            return fragments.Select(f => new OcrFragment(f.Text, Round(f.Confidence), ToBounds(f.Bounds),
                f.Words.Select(w => new OcrWord(w.Text, Round(w.Confidence), ToBounds(w.Bounds))).ToList()));
        }

        protected static IEnumerable<OcrFragment> ToFragments(IEnumerable<PdfTextFragment> fragments) {
            return fragments.Select(f => new OcrFragment(f.Text, Round(f.Confidence), ToBounds(f.Bounds),
                f.Words.Select(w => new OcrWord(w.Text, Round(w.Confidence), ToBounds(w.Bounds))).ToList()));
        }

        protected static IEnumerable<OcrPage> ToPages(IReadOnlyList<PdfPageRecognitionResult> pages) {
            return pages.Select(p => new OcrPage(p.PageIndex + 1, ToFragments(p.TextFragments).ToList()));
        }

        static float? Round(float? confidence) {
            return confidence.HasValue ? (float)Math.Round(confidence.Value, 3) : null;
        }

        static OcrBounds ToBounds(RectangleF bounds) {
            return new OcrBounds((float)Math.Round(bounds.X, 2), (float)Math.Round(bounds.Y, 2),
                (float)Math.Round(bounds.Width, 2), (float)Math.Round(bounds.Height, 2));
        }
    }
}
