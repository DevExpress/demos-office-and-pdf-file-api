# Word Mail Merge

This sample uses the **New DevExpress Word Processing Document API (WordDocument)** to generate a personalized letter from a predefined template.

- Input: `Documents/MailMerge_Template.docx`.
- `EmployeeId`: `1` (Nancy Davolio, the sender). The sample replaces the `INCLUDEPICTURE "dbimg://..."` field with `Documents/Photo1.jpeg`.
- `CustomerId`: `ALFKI` (Alfreds Futterkiste, the recipient).
- `OutputFormat`: `pdf`. The sample exports the merged document to PDF. Set it to `docx`, `doc`, `docm`, `dotx`, `dotm`, or `rtf` (and `IsPdfOutput` to `false`) to save a Word file.
- Output: `Output/MailMerge_Result.pdf` beside the executable.