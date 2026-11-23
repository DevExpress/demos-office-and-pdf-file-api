using System.Drawing;
using System.Globalization;
using DevExpress.Docs.Pdf;
using DevExpress.Drawing.Printing;

internal static class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "GeneratedInvoice.pdf";
    const bool IncludeZugferd = true;
    static readonly string XmlFilePath = Path.Combine("Documents", "PDF", "Northwind_Invoice.xml");
    const int InvoiceNumber = 10643;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var xmlStream = IncludeZugferd ? File.OpenRead(Path.Combine(AppContext.BaseDirectory, XmlFilePath)) : null;
            using var outputStream = GenerateInvoiceDocument(CultureInfo.GetCultureInfo("en-US"), InvoiceNumber, IncludeZugferd, xmlStream);
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

    static Stream GenerateInvoiceDocument(CultureInfo documentCulture, int invoiceNumber, bool includeZugferd, Stream? xmlStream) {
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

        doc.Metadata.Xmp.XmpPdfUASchema.Part = new XmpInteger(1);
        doc.Metadata.Xmp.XmpPdfAExtensions.Register(new XmpPdfAExtensionDescriptor(
            "http://www.aiim.org/pdfua/ns/id/",
            "pdfuaid",
            "PDF/UA identification schema",
            new List<XmpPdfAExtensionProperty> {
                new("part", "Indicates which part of ISO 14289 is followed.", "Integer")
            }
        ));
        doc.Metadata.Xmp.XmpPdfASchema.Part = new XmpInteger(3);
        doc.Metadata.Xmp.XmpPdfASchema.Conformance = new XmpString("B");
        doc.Metadata.Xmp.XmpDublinCoreSchema.Title.Add(documentCulture.Name, title);

        var root = doc.StructureTree.AddChildElement(Pdf17StructureType.Document);
        root.LanguageCulture = documentCulture;
        root.Title = title;
        var page = doc.Pages.Add(DXPaperKind.A4);

        var headerSection = root.AddChildElement(Pdf17StructureType.Sect);
        var artifactGroupContent = new MarkedContentGroup("Artifact");
        page.AddFragment(artifactGroupContent);
        artifactGroupContent.Fragments.Add(PathFragment.Rectangle(0, page.CropBox.Height - 130, page.CropBox.Width, 130, new SolidFill(PdfColor.FromRgb(0xDF, 0xE3, 0xE8))));

        float currentX = page.CropBox.Left + 60f;
        headerSection.AddChildElement(Pdf17StructureType.H1).AddFragment(page, new TextFragment { Text = "Northwind Traders", Location = new PointF(currentX, page.CropBox.Height - 40), ForegroundFill = darkFill, FontSize = 14 });
        headerSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "One Portals Way, Twin Points WA, 98156", Location = new PointF(currentX, page.CropBox.Height - 58) });
        headerSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "1-206-555-1417", Location = new PointF(currentX, page.CropBox.Height - 74) });
        headerSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "northwind@mail.com", Location = new PointF(currentX, page.CropBox.Height - 90) });
        headerSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "www.northwind.com", Location = new PointF(currentX, page.CropBox.Height - 106) });
        headerSection.AddChildElement(Pdf17StructureType.H1).AddFragment(page, new TextFragment { Text = "INVOICE", Location = new PointF(393f, page.CropBox.Height - 68), ForegroundFill = darkFill, FontSize = 28 });

        float metaLabelX = page.CropBox.Right - 200f;
        float metaValX = page.CropBox.Right - 90f;
        var metaSection = headerSection.AddChildElement(Pdf17StructureType.Sect);
        metaSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "Invoice №:", Location = new PointF(metaLabelX, page.CropBox.Height - 90), ForegroundFill = mutedFill }).AddFragment(new TextFragment { Text = invoiceNumber.ToString(), Location = new PointF(metaValX, page.CropBox.Height - 90) });
        metaSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "Invoice Date:", Location = new PointF(metaLabelX, page.CropBox.Height - 106), ForegroundFill = mutedFill }).AddFragment(new TextFragment { Text = "04-21-26", Location = new PointF(metaValX, page.CropBox.Height - 106) });

        var billSection = root.AddChildElement(Pdf17StructureType.Sect);
        billSection.AddChildElement(Pdf17StructureType.H2).AddFragment(page, new TextFragment { Text = "Bill to:", Location = new PointF(page.CropBox.Left + 50, page.CropBox.Height - 155) });

        (string Label, string Value)[] billRows = {
            ("Company:", "Alfreds Futterkiste"),
            ("Contact Name:", "Maria Anders"),
            ("Address:", "Obere Str. 57, Berlin, Germany, 12209"),
            ("Phone:", "030-0074321"),
            ("Mail:", "alfredsfutterkiste@mail.com")
        };
        float billValueX = page.CropBox.Left + 150f;
        float billRowY = page.CropBox.Height - 175f;
        foreach(var (label, value) in billRows) {
            billSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = label, Location = new PointF(page.CropBox.Left + 50, billRowY), ForegroundFill = mutedFill }).AddFragment(new TextFragment { Text = value, Location = new PointF(billValueX, billRowY) });
            billRowY -= 16f;
        }

        float tableTop = billRowY - 20f;
        float rowHeight = 20f;
        float tableWidth = page.CropBox.Right - page.CropBox.Left - 100f;
        var table = root.AddChildElement(Pdf17StructureType.Table);
        var tableHead = table.AddChildElement(Pdf17StructureType.THead);
        artifactGroupContent.Fragments.Add(PathFragment.Rectangle(page.CropBox.Left + 50, tableTop - rowHeight, tableWidth, rowHeight, SolidFill.LightGray));
        artifactGroupContent.Fragments.Add(PathFragment.Line(page.CropBox.Left + 50, tableTop - rowHeight, page.CropBox.Right - 50f, tableTop - rowHeight, heavyBorder));

        var headerRow = tableHead.AddChildElement(Pdf17StructureType.TR);
        float[] columnPositions = { 55f, 95f, 305f, 375f, 435f, 495f };
        string[] headers = { "#", "Product Name", "Unit Price", "Quantity", "Discount", "Total" };
        for(int index = 0; index < headers.Length; index++) {
            var headerCell = headerRow.AddChildElement(Pdf17StructureType.TH);
            headerCell.Attributes.Add(new TableAttribute { Scope = TableScope.Column });
            headerCell.AddFragment(page, new TextFragment { Text = headers[index], Location = new PointF(columnPositions[index], tableTop - rowHeight + 5), ForegroundFill = darkFill });
        }

        (string Num, string Name, string Price, string Quantity, string Discount, string Total)[] items = {
            ("01", "Rössle Sauerkraut", "$45.60", "15", "$3.75", "$683.75"),
            ("02", "Chartreuse verte", "$18.00", "21", "$5.25", "$377.75"),
            ("03", "Spegesild", "$12.00", "2", "$0.50", "$23.75")
        };

        var tableBody = table.AddChildElement(Pdf17StructureType.TBody);
        float rowY = tableTop - 2 * rowHeight;
        foreach(var (num, name, price, quantity, discount, total) in items) {
            var row = tableBody.AddChildElement(Pdf17StructureType.TR);
            (string Value, float X)[] cells = {
                (num, columnPositions[0]), (name, columnPositions[1]), (price, columnPositions[2]),
                (quantity, columnPositions[3]), (discount, columnPositions[4]), (total, columnPositions[5])
            };
            foreach(var (value, positionX) in cells)
                row.AddChildElement(Pdf17StructureType.TD).AddFragment(page, new TextFragment { Text = value, Location = new PointF(positionX, rowY + 5) });
            artifactGroupContent.Fragments.Add(PathFragment.Line(page.CropBox.Left + 50, rowY, page.CropBox.Right - 50f, rowY, rowDivider));
            rowY -= rowHeight;
        }

        float totalY = rowY + rowHeight - 25f;
        float totalLabelX = columnPositions[3];
        float totalValueX = columnPositions[5];
        var totalSection = root.AddChildElement(Pdf17StructureType.Sect);
        artifactGroupContent.Fragments.Add(PathFragment.Line(totalLabelX - 5f, totalY - 22f, page.CropBox.Right - 50f, totalY - 22f, heavyBorder));
        totalSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "Sub Total:", Location = new PointF(totalLabelX, totalY) }).AddFragment(new TextFragment { Text = "$1086.00", Location = new PointF(totalValueX, totalY) });
        totalSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "Discount Total:", Location = new PointF(totalLabelX, totalY - 16f) }).AddFragment(new TextFragment { Text = "$9.50", Location = new PointF(totalValueX, totalY - 16f) });
        float grandTotalY = totalY - 38f;
        totalSection.AddChildElement(Pdf17StructureType.P).AddFragment(page, new TextFragment { Text = "Grand Total:", Location = new PointF(totalLabelX, grandTotalY), ForegroundFill = darkFill }).AddFragment(new TextFragment { Text = "$1076.50", Location = new PointF(totalValueX, grandTotalY), ForegroundFill = darkFill });

        if(includeZugferd && xmlStream != null)
            doc.AttachZugferdInvoice(xmlStream);

        var outputStream = new MemoryStream();
        doc.Save(outputStream, new SaveOptions() { SyncMetadata = true, UpdateCreatedAt = true, UpdateModifiedAt = true });
        outputStream.Position = 0;
        return outputStream;
    }
}
