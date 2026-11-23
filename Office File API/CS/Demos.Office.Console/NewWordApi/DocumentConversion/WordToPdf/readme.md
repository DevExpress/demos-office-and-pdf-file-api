# Word to PDF (New Word API)

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to convert a Word document to PDF.

The sample reads `Documents/Sample.docx` and exports it to PDF with the following predefined settings:

- PDF/A conformance level: `PdfACompatibility.None`
- PDF/UA compatibility: `PdfUACompatibility.None`
- Title: `Custom Title`
- Author: `Custom Author`
- Subject and keywords: empty
- Export interactive controls as AcroForm fields: `false`

The sample writes `Output/ConvertedDocument.pdf` next to the executable.