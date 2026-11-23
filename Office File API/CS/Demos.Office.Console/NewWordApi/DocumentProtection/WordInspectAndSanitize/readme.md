# Inspect and Sanitize Word Document

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to inspect Word document for hidden or sensitive content, or to remove that content.

- Input file: `Documents/Sample.docx`
- Operation: `InspectOperationType.Inspect` (`Operation` constant). Set it to `InspectOperationType.Sanitize` to clean document.
- Inspect output: `Output/Inspection_Log.txt` beside executable. Log lists content types that `WordDocument.Inspect` finds, or `No issues found.`
- Sanitize output: `Output/Sanitized_Document.docx` beside executable, saved as `DocumentFormat.Docx`
- Sanitize options (all `true`): remove hidden text, invisible text, metadata, comments, custom XML parts, ActiveX content, macros, OLE objects, and reject tracked changes