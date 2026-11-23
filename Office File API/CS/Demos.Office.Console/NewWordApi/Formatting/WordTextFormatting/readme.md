# Word Text Formatting

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to format document text.

- Input file: `Documents/Sample.docx`
- Action: `TextFormattingActionType.ApplyFormatting`. Set `ActionType` to `ClearFormatting` to reset attributes listed in `ResetAttributes` (all attributes by default).
- Range: `TextFormattingRangeType.FirstParagraph`. Other values: `FirstSection`, `WholeDocument`.
- Formatting: font `Segoe UI`, size 16, regular style, font color `#000000`, background color `#FFFFFF`, no underline
- Output format: `docx` (`pdf`, `doc`, `dotx`, `rtf` are also supported)
- Output file: `Output/Formatted_Document.docx` beside executable