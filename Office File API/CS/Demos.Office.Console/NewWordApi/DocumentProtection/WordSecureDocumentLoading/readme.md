# Secure Word Document Loading

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to load Word document with protection against potentially harmful content.

- Input file: `Documents/Sample.docx`
- Security limits: 10 sections, 500 paragraphs, 10 tables, 100 rows per table, 50 sub-documents, 5 MB file size, 5000 XML elements, XML depth 50
- Content to remove: ActiveX content, custom XML parts, macros, OLE objects, external images, `INCLUDEPICTURE` fields, DDE fields
- Restricted hyperlink removal mode: `RestrictedHyperlinkRemovalMode.UriOnly`
- Output format: `DocumentFormat.Docx`
- Output file: `Output/Secure_Document.docx` beside executable

Custom `ISecurityLoadingHandler` throws when document exceeds limit and removes unsafe content. When limit is exceeded, sample prints error and returns exit code 1.