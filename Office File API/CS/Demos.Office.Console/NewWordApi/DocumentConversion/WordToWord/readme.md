# Word to Word (New Word API)

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to convert a Word document to another Word-compatible format.

The sample reads `Documents/Sample.docx` and saves it in the format set by the `OutputFormat` constant (`docx`). Supported values: `docx`, `doc`, `docm`, `dotx`, `dotm`, `rtf`, `mht`, `html`, `txt`, and `odt`. The `EmbedFonts` constant is `false`. Images are embedded when the output format is HTML or MHT.

The sample writes `Output/ConvertedDocument.docx` next to the executable.