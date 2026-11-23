# Encrypt Word Document with Password

This sample uses the new DevExpress Word Processing Document API (`WordDocument`) to encrypt Word document with password.

- Input file: `Documents/Sample.docx`
- Password: `demo-password` (`Password` constant)
- Encryption type: `EncryptionType.Strong` (AES 256-bit). Use `EncryptionType.Compatible` for RC4 128-bit encryption.
- Output format: `DocumentFormat.Docx`
- Output file: `Output/Encrypted_Document.docx` beside executable

Sample loads document, calls `WordDocument.Encrypt` with `EncryptionOptions`, and saves encrypted document.