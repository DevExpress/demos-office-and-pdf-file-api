# Word Table Style

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to create custom table style and apply it to first table in template.

- Input file: `Documents/ContentControls_AppointmentForm.docx`
- Table style: `CustomTableStyle` with white inside borders, green cell background, and conditional styles for header row, total row, first and last columns, and row and column banding
- Table look: `ApplyFirstRow | DoNotApplyColumnBanding | DoNotApplyRowBanding` (`TableLook` constant)
- Output format: `DocumentFormat.Docx`
- Output file: `Output/Table_Style.docx` beside executable