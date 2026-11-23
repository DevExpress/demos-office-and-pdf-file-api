using DevExpress.Docs.Word;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Comments_template.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Comments_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    static readonly CommentActionType CommentAction = CommentActionType.AddNew;
    static readonly RemoveCommentActionType RemoveCommentAction = RemoveCommentActionType.ByAuthor;
    const string AuthorFilter = "Author1";
    const string CommentText = "Test Comment";
    const string CommentAuthor = "TestUser";
    const string ReplyText = "Test Reply";
    const string ReplyAuthor = "TestUser";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProcessComments(inputStream);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);
            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream ProcessComments(Stream inputStream) {
        switch(CommentAction) {
            case CommentActionType.AddNew:
                return AddNewComment(inputStream, CommentText, CommentAuthor, OutputFormat);
            case CommentActionType.Reply:
                return ReplyToComment(inputStream, ReplyText, ReplyAuthor, OutputFormat);
            case CommentActionType.Remove:
                return RemoveComments(inputStream, RemoveCommentAction == RemoveCommentActionType.AllComments, AuthorFilter, OutputFormat);
            default:
                throw new InvalidOperationException("Unsupported comment action");
        }
    }

    static Stream AddNewComment(Stream inputStream, string commentText, string author, DocumentFormat outputFormat) {
        using var document = new WordDocument(inputStream);

        var paragraph = document.Body.OfType<Paragraph>().First();
        var comment = document.Comments.Create(document.Body.GetRange(paragraph), new CommentDefinition { Author = author });
        comment.Content.AddText(commentText);

        var outputStream = new MemoryStream();
        document.Save(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream ReplyToComment(Stream inputStream, string replyText, string author, DocumentFormat outputFormat) {
        using var document = new WordDocument(inputStream);

        // Reply to the first comment
        if(document.Comments.Count > 0) {
            var parentComment = document.Comments[0];
            var reply = parentComment.CreateReply(new CommentDefinition { Author = author });
            reply.Content.AddText(replyText);
        }

        var outputStream = new MemoryStream();
        document.Save(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream RemoveComments(Stream inputStream, bool removeAll, string authorFilter, DocumentFormat outputFormat) {
        using var document = new WordDocument(inputStream);

        if(removeAll) {
            for(var i = document.Comments.Count - 1; i >= 0; i--)
                document.Comments.RemoveAt(i);
        }
        else if(!string.IsNullOrWhiteSpace(authorFilter)) {
            for(var i = document.Comments.Count - 1; i >= 0; i--) {
                if(string.Equals(document.Comments[i].Author, authorFilter, StringComparison.OrdinalIgnoreCase)) {
                    document.Comments.RemoveAt(i);
                }
            }
        }

        var outputStream = new MemoryStream();
        document.Save(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}

enum CommentActionType {
    AddNew,
    Reply,
    Remove
}

enum RemoveCommentActionType {
    ByAuthor,
    AllComments
}