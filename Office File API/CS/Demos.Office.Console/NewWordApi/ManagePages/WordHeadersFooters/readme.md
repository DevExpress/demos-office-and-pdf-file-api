# Word Headers and Footers

This sample uses the **New DevExpress Word Processing Document API (WordDocument)** to add headers and footers to a predefined template.

- Input: `Documents/HeaderFooter_template.docx`.
- `AddHeaders`: `true`. Header text `Document Title`, center alignment.
- `AddFooters`: `true`. Footer text `Confidential`, center alignment.
- `AddPageNumbering`: `true`. The footer gets a `Page X of Y` paragraph with `PAGE` and `NUMPAGES` fields.
- `DifferentFirstPage`: `true`. The first page gets header `Cover Page` and footer `Internal Use Only`, center alignment.
- `OutputFormat`: `docx`.
- Output: `Output/Modified_Document.docx` beside the executable.