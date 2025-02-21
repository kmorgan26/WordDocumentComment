using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public static class CommentConsolidator
{
    public static void ConsolidateComments(string filePath1, string filePath2, string outputFilePath)
    {
        // Create a new Word document for consolidated comments
        using (WordprocessingDocument newDoc = WordprocessingDocument.Create(outputFilePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Add main document part
            var mainPart = newDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Add comments part
            var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments();

            // Copy content from the first document
            CopyContentFromDocument(filePath1, mainPart);

            // Copy content from the second document
            CopyContentFromDocument(filePath2, mainPart);

            // Extract comments from both documents and add them to the new document
            AddCommentsFromDocument(filePath1, commentsPart, "Document 1");
            AddCommentsFromDocument(filePath2, commentsPart, "Document 2");

            // Save the comments part
            commentsPart.Comments.Save();

            // Save the new document
            mainPart.Document.Save();
        }
    }

    private static void CopyContentFromDocument(string filePath, MainDocumentPart mainPart)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
        {
            var body = doc.MainDocumentPart.Document.Body;
            foreach (var element in body.Elements())
            {
                mainPart.Document.Body.Append(element.CloneNode(true));
            }
        }
    }

    private static void AddCommentsFromDocument(string filePath, WordprocessingCommentsPart commentsPart, string sourceDocumentName)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
        {
            if (doc.MainDocumentPart.WordprocessingCommentsPart != null)
            {
                foreach (var comment in doc.MainDocumentPart.WordprocessingCommentsPart.Comments.Elements<Comment>())
                {
                    // Clone the comment and add metadata
                    var newComment = new Comment
                    {
                        Id = comment.Id,
                        Author = comment.Author,
                        Date = comment.Date,
                        Initials = comment.Initials
                    };
                    newComment.Append(new Paragraph(new Run(new Text(comment.InnerText))));
                    commentsPart.Comments.Append(newComment);
                }
            }
        }
    }
}
