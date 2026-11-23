# Protect Word Document with Password

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to protect Word document from modifications with password.

- Input file: `Documents/Sample.docx`
- Password: `demo-password` (`Password` constant)
- Protection type: `DocumentProtectionType.ReadOnly`. Other values: `AllowComments`, `TrackedChanges`, `FillInForms`.
- Output format: `DocumentFormat.Docx`
- Output file: `Output/Protected_Document.docx` beside executable

Sample removes existing protection, applies new protection with `DocumentProtectionOptions`, and saves document.