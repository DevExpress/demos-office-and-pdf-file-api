using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.GenerateDocument;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;

#region Code
using System.Drawing;
using System.Globalization;
using DevExpress.Drawing.Printing;
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.GenerateDocument {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfGenerateInvoiceController : PdfControllerBase {
        [HttpPost("generate-document/generate-invoice")]
        public async Task<IActionResult> GenerateInvoice([FromForm] PdfGenerateInvoiceModel model, IFormFile xmlFile) {
            MemoryStream xmlStream = null;
            if(model.IncludeZugferd && xmlFile != null) {
                xmlStream = new MemoryStream();
                await xmlFile.CopyToAsync(xmlStream);
                xmlStream.Position = 0;
            }

            using(xmlStream) {
                CultureInfo documentCulture = CultureInfo.GetCultureInfo("en-US");
                int invoiceNumber = 10643;
                var outputStream = GenerateInvoiceDocument(documentCulture, invoiceNumber, model.IncludeZugferd, xmlStream);
                outputStream.Seek(0, SeekOrigin.Begin);
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
        }

        #region Code
        Stream GenerateInvoiceDocument(CultureInfo documentCulture, int invoiceNumber, bool includeZugferd, Stream xmlStream) {
            var title = $"Invoice {invoiceNumber}";
            var darkFill = new SolidFill(PdfColor.FromRgb(0x2F, 0x3B, 0x4A));
            var mutedFill = new SolidFill(PdfColor.FromRgb(0x7A, 0x87, 0x94));
            var rowDivider = Outline.Create(new SolidFill(PdfColor.FromRgb(0xE1, 0xE5, 0xEA)), 0.5f);
            var heavyBorder = Outline.Create(new SolidFill(PdfColor.FromRgb(0x9A, 0xA5, 0xB1)), 2f);

            using var doc = new PdfDocument() {
                ViewerPreferences = new ViewerPreferences() { DisplayDocTitle = true },
                LanguageCulture = documentCulture,
                Metadata = new DocumentMetadata() { Xmp = new XmpMetadata() },
                MarkInfo = new MarkInfo() { Marked = true }
            };

            // ── metadata ─────────────────────────────────────────────────────
            doc.Metadata.Xmp.XmpPdfUASchema.Part = new XmpInteger(1);
            doc.Metadata.Xmp.XmpPdfAExtensions.Register(new XmpPdfAExtensionDescriptor(
                "http://www.aiim.org/pdfua/ns/id/",
                "pdfuaid",
                "PDF/UA identification schema",
                new List<XmpPdfAExtensionProperty> {
                    new XmpPdfAExtensionProperty("part", "Indicates which part of ISO 14289 is followed.", "Integer")
                }
            ));
            doc.Metadata.Xmp.XmpPdfASchema.Part = new XmpInteger(3);
            doc.Metadata.Xmp.XmpPdfASchema.Conformance = new XmpString("B");
            doc.Metadata.Xmp.XmpDublinCoreSchema.Title.Add(documentCulture.Name, title);

            // ── new page and structure root ───────────────────────────────────────
            var root = doc.StructureTree.AddChildElement(Pdf17StructureType.Document);
            root.LanguageCulture = documentCulture;
            root.Title = title;
            var page = doc.Pages.Add(DXPaperKind.A4);

            // ── header ───────────────────────────────────────────────────────
            var headerSection = root.AddChildElement(Pdf17StructureType.Sect);
            var artifactGroupContent = new MarkedContentGroup("Artifact");
            page.AddFragment(artifactGroupContent);
            artifactGroupContent.Fragments.Add(PathFragment.Rectangle(0, page.CropBox.Height - 130, page.CropBox.Width, 130,
                new SolidFill(PdfColor.FromRgb(0xDF, 0xE3, 0xE8))));

            float cx = page.CropBox.Left + 60f;
             headerSection.AddChildElement(Pdf17StructureType.H1)
                 .AddFragment(page, new TextFragment { Text = "Northwind Traders", Location = new PointF(cx, page.CropBox.Height - 40), ForegroundFill = darkFill, FontSize = 14});
            headerSection.AddChildElement( Pdf17StructureType.P).AddFragment(page,
                new TextFragment { Text = "One Portals Way, Twin Points WA, 98156", Location = new PointF(cx, page.CropBox.Height - 58) });

            headerSection.AddChildElement(Pdf17StructureType.P)
                .AddFragment(page, new TextFragment { Text = "1-206-555-1417", Location = new PointF(cx, page.CropBox.Height - 74) });
            headerSection.AddChildElement(Pdf17StructureType.P)
                .AddFragment(page, new TextFragment { Text = "northwind@mail.com", Location = new PointF(cx, page.CropBox.Height - 90) });
            headerSection.AddChildElement(Pdf17StructureType.P)
                .AddFragment(page,  new TextFragment { Text = "www.northwind.com", Location = new PointF(cx, page.CropBox.Height - 106) });
            headerSection.AddChildElement(Pdf17StructureType.H1)
                .AddFragment(page,new TextFragment { Text = "INVOICE", Location = new PointF(393f, page.CropBox.Height - 68), ForegroundFill = darkFill, FontSize = 28});

            float metaLabelX = page.CropBox.Right  - 200f, metaValX = page.CropBox.Right - 90f;
            var metaSection = headerSection.AddChildElement(Pdf17StructureType.Sect);

            metaSection.AddChildElement(Pdf17StructureType.P)
                .AddFragment(page, new TextFragment { Text = "Invoice №:", Location = new PointF(metaLabelX, page.CropBox.Height - 90), ForegroundFill = mutedFill })
                .AddFragment(new TextFragment { Text = invoiceNumber.ToString(), Location = new PointF(metaValX, page.CropBox.Height - 90) });

            metaSection.AddChildElement(Pdf17StructureType.P)
                .AddFragment(page, new TextFragment { Text = "Invoice Date:", Location = new PointF(metaLabelX, page.CropBox.Height - 106), ForegroundFill = mutedFill })
                .AddFragment(new TextFragment { Text = "04-21-26", Location = new PointF(metaValX, page.CropBox.Height - 106) });

            // ── bill to ───────────────────────────────────────────────────────
            var billSection = root.AddChildElement(Pdf17StructureType.Sect);

            billSection.AddChildElement(Pdf17StructureType.H2)
                .AddFragment(page, new TextFragment { Text = "Bill to:", Location = new PointF(page.CropBox.Left + 50, page.CropBox.Height - 155) });

            (string label, string value)[] billRows = {
                ("Company:", "Alfreds Futterkiste"),
                ("Contact Name:", "Maria Anders"),
                ("Address:", "Obere Str. 57, Berlin, Germany, 12209"),
                ("Phone:", "030-0074321"),
                ("Mail:", "alfredsfutterkiste@mail.com"),
            };
            float billValX = page.CropBox.Left + 150f;
            float billRowY = page.CropBox.Height - 175f;
            foreach(var (label, value) in billRows) {
                billSection.AddChildElement( Pdf17StructureType.P)
                    .AddFragment(page, new TextFragment { Text = label, Location = new PointF(page.CropBox.Left + 50, billRowY), ForegroundFill = mutedFill })
                    .AddFragment(new TextFragment { Text = value, Location = new PointF(billValX, billRowY) });
                billRowY -= 16f;
            }

            // ── table ─────────────────────────────────────────────────────────
            float tableTop = billRowY - 20f;
            float rowH = 20f;
            float tableW = page.CropBox.Right  - page.CropBox.Left - 100f;

            var table = root.AddChildElement(Pdf17StructureType.Table);
            var thead= table.AddChildElement(Pdf17StructureType.THead);

            artifactGroupContent.Fragments.Add(PathFragment.Rectangle(page.CropBox.Left + 50, tableTop - rowH, tableW, rowH, SolidFill.LightGray));
            artifactGroupContent.Fragments.Add(PathFragment.Line(page.CropBox.Left + 50, tableTop - rowH, page.CropBox.Right - 50f, tableTop - rowH, heavyBorder));

            var hRow = thead.AddChildElement(Pdf17StructureType.TR);
            float[] colX = { 55f, 95f, 305f, 375f, 435f, 495f };
            string[] headers = { "#", "Product Name", "Unit Price", "Quantity", "Discount", "Total" };
            for(int i = 0; i < headers.Length; i++) {
                var th = hRow.AddChildElement(Pdf17StructureType.TH);
                th.Attributes.Add(new TableAttribute { Scope = TableScope.Column });
                th.AddFragment(page,
                    new TextFragment { Text = headers[i], Location = new PointF(colX[i], tableTop - rowH + 5), ForegroundFill = darkFill, });
            }

            (string num, string name, string price, string qty, string disc, string total)[] items = {
                ("01", "Rössle Sauerkraut", "$45.60",  "15", "$3.75", "$683.75"),
                ("02", "Chartreuse verte",  "$18.00",  "21", "$5.25", "$377.75"),
                ("03", "Spegesild",         "$12.00",   "2", "$0.50",  "$23.75"),
            };

            var tbody = table.AddChildElement( Pdf17StructureType.TBody);

            float rowY = tableTop - 2 * rowH;
            foreach(var (num, name, price, qty, disc, total) in items) {
                var tr = tbody.AddChildElement(Pdf17StructureType.TR);
                (string val, float x)[] cells = {
                    (num, colX[0]), (name, colX[1]), (price, colX[2]),
                    (qty, colX[3]), (disc, colX[4]), (total, colX[5]),
                };
                foreach(var (val, x) in cells)
                    tr.AddChildElement(Pdf17StructureType.TD)
                        .AddFragment(page, new TextFragment { Text = val, Location = new PointF(x, rowY + 5) });
                artifactGroupContent.Fragments.Add(PathFragment.Line(page.CropBox.Left + 50 , rowY, page.CropBox.Right - 50f, rowY, rowDivider));
                rowY -= rowH;
            }

            // ── totals ────────────────────────────────────────────────────────
            float totY = rowY + rowH - 25f;
            float totLabelX = colX[3], totValX = colX[5];

            var totalsSect = root.AddChildElement(Pdf17StructureType.Sect);
            artifactGroupContent.Fragments.Add(PathFragment.Line(totLabelX - 5f, totY - 22f, page.CropBox.Right - 50f, totY - 22f, heavyBorder));

             totalsSect.AddChildElement(Pdf17StructureType.P)
                .AddFragment(page, new TextFragment { Text = "Sub Total:", Location = new PointF(totLabelX, totY) })
                .AddFragment(new TextFragment { Text = "$1086.00", Location = new PointF(totValX, totY) });

            totalsSect.AddChildElement(Pdf17StructureType.P)
                .AddFragment(page, new TextFragment { Text = "Discount Total:", Location = new PointF(totLabelX, totY - 16f) })
                .AddFragment(new TextFragment { Text = "$9.50", Location = new PointF(totValX, totY - 16f) });

            float gtY = totY - 38f;
            totalsSect.AddChildElement(Pdf17StructureType.P)
                .AddFragment(page, new TextFragment { Text = "Grand Total:", Location = new PointF(totLabelX, gtY), ForegroundFill = darkFill })
                .AddFragment(new TextFragment { Text = "$1076.50", Location = new PointF(totValX, gtY), ForegroundFill = darkFill });

            if(includeZugferd && xmlStream != null) {
                doc.AttachZugferdInvoice(xmlStream);
            }

            var outputStream = new MemoryStream();
            doc.Save(outputStream, new SaveOptions() {SyncMetadata = true, UpdateCreatedAt = true, UpdateModifiedAt = true });
            return outputStream;
        }
        #endregion
    }
}
