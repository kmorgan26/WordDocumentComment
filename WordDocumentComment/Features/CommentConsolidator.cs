using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

/// <summary>
/// Provides methods to consolidate comments from multiple Word documents.
/// </summary>
public static class CommentConsolidator
{
    /// <summary>
    /// Consolidates comments from two Word documents into a new document.
    /// </summary>
    /// <param name="filePath1">The file path of the first Word document.</param>
    /// <param name="filePath2">The file path of the second Word document.</param>
    /// <param name="outputFilePath">The file path for the output consolidated document.</param>
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

    /// <summary>
    /// Copies the content from a Word document to the main part of a new document.
    /// </summary>
    /// <param name="filePath">The file path of the source Word document.</param>
    /// <param name="mainPart">The main document part of the new document.</param>
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

    /// <summary>
    /// Extracts comments from a Word document and adds them to the comments part of a new document.
    /// </summary>
    /// <param name="filePath">The file path of the source Word document.</param>
    /// <param name="commentsPart">The comments part of the new document.</param>
    /// <param name="sourceDocumentName">The name of the source document to include in the comment metadata.</param>
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

    /// <summary>
    /// Compares Word documents from two directories, consolidates their content and comments into a new document, and logs the differences and comments.
    /// </summary>
    public static void CompareAndConsolidateComments()
    {
        Console.WriteLine("This program compares Word documents from two directories, consolidates their content and comments into a new document, and logs the differences and comments.");
        Console.WriteLine("Please provide the paths to the two directories containing the Word documents and the output directory for the consolidated document and log files.");

        Console.WriteLine("Enter the path to the first directory:");
        string directory1 = Console.ReadLine();

        Console.WriteLine("Enter the path to the second directory:");
        string directory2 = Console.ReadLine();

        Console.WriteLine("Enter the path to the output directory for consolidated comments:");
        string outputDirectory = Console.ReadLine();

        if (Directory.Exists(directory1) && Directory.Exists(directory2))
        {
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Get all Word documents in the first directory
            var filesInDir1 = Directory.GetFiles(directory1, "*.docx");

            foreach (var file1 in filesInDir1)
            {
                string fileName = Path.GetFileName(file1);
                string file2 = Path.Combine(directory2, fileName);

                if (File.Exists(file2))
                {
                    Console.WriteLine($"Comparing {file1} and {file2}...");

                    // Compare document content
                    var contentDifferences = DocumentComparer.CompareWordDocuments(file1, file2);
                    if (contentDifferences.Count == 0)
                    {
                        Console.WriteLine("No differences in document content found.");
                    }
                    else
                    {
                        Console.WriteLine("Differences in document content found:");
                        foreach (var diff in contentDifferences)
                        {
                            Console.WriteLine(diff);
                        }
                    }

                    // Consolidate comments into a new document
                    string outputFilePath = Path.Combine(outputDirectory, $"ConsolidatedComments_{fileName}");
                    CommentConsolidator.ConsolidateComments(file1, file2, outputFilePath);

                    Console.WriteLine($"Comments consolidated into {outputFilePath}");

                    // Create log file
                    string logFilePath = Path.Combine(outputDirectory, $"Log_{fileName}.txt");
                    Logger.CreateLogFile(logFilePath, contentDifferences, file1, file2);
                }
                else
                {
                    Console.WriteLine($"No matching file found in {directory2} for {fileName}");
                }
            }
        }
        else
        {
            Console.WriteLine("One or both directories do not exist.");
        }
    }
}
