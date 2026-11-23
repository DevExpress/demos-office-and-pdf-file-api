# Word Comments

This sample uses the **New DevExpress Word Processing Document API (WordDocument)** to manage comments in a Word document.

- Input: `Documents/Comments_template.docx`.
- `CommentAction`: `AddNew` (default). The sample adds a comment to the first paragraph. Set it to `Reply` to reply to the first comment, or to `Remove` to remove comments.
- `CommentText` / `CommentAuthor`: `Test Comment` / `TestUser`.
- `ReplyText` / `ReplyAuthor`: `Test Reply` / `TestUser`.
- `RemoveCommentAction`: `ByAuthor` removes comments by `AuthorFilter` (`Author1`); `AllComments` removes all comments.
- Output: `Output/Comments_Result.docx` (DOCX) beside the executable.