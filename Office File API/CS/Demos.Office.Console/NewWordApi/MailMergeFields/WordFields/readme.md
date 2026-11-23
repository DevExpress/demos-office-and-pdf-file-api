# Word Fields

This sample uses the **New DevExpress Word Processing Document API (WordDocument)** to manage fields in a predefined template.

- Input: `Documents/Fields_Template.docx`. `Documents/DevExpress.png` is the image for the `Test` document variable.
- `Action`: `Insert` (default). The sample adds one table row per field type and updates all fields. Set it to `Update` to update field values, `Lock` to lock fields, `Unlink` to replace fields with their values, or `Delete` to remove fields.
- `FieldEntryValues`: `DATE`, `TIME`, `PAGE`, `NUMPAGES`, `DocVariable`, `AUTHOR`, `TITLE`, `DOCPROPERTY`. `Lock`, `Unlink`, and `Delete` process only fields of these types.
- `OutputFormat`: `docx`.
- Output: `Output/Fields_Result.docx` beside the executable.