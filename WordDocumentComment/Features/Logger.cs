using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

/// <summary>
/// Provides methods for logging differences and comments from Word documents.
/// </summary>
public static class Logger
{
    /// <summary>
    /// Extracts comments from all Word documents in a directory and writes them to a CSV file.
    /// </summary>
    public static void CommentsToCSV()
    {
        Console.WriteLine("This feature extracts comments from all Word documents in a directory and writes them to a CSV file.");
        Console.WriteLine("Enter the path to the directory containing the Word documents:");
        string directory = Console.ReadLine();

        if (Directory.Exists(directory))
        {
            string csvFilePath = Path.Combine(directory, "Comments.csv");
            ExtractCommentsToCSV(directory, csvFilePath);
            Console.WriteLine($"Comments extracted to {csvFilePath}");
        }
        else
        {
            Console.WriteLine("The directory does not exist.");
        }
    }

    /// <summary>
    /// Creates a log file with content differences and comments from two Word documents.
    /// </summary>
    /// <param name="logFilePath">The file path for the log file.</param>
    /// <param name="contentDifferences">A list of content differences between the two documents.</param>
    /// <param name="filePath1">The file path of the first Word document.</param>
    /// <param name="filePath2">The file path of the second Word document.</param>
    public static void CreateLogFile(string logFilePath, List<string> contentDifferences, string filePath1, string filePath2)
    {
        using (StreamWriter writer = new StreamWriter(logFilePath))
        {
            writer.WriteLine("Content Differences:");
            foreach (var diff in contentDifferences)
            {
                writer.WriteLine(diff);
            }

            writer.WriteLine();
            writer.WriteLine("Comments from Document 1:");
            WriteCommentsToLog(filePath1, writer);

            writer.WriteLine();
            writer.WriteLine("Comments from Document 2:");
            WriteCommentsToLog(filePath2, writer);
        }
    }

    /// <summary>
    /// Writes comments from a Word document to a log file.
    /// </summary>
    /// <param name="filePath">The file path of the Word document.</param>
    /// <param name="writer">The StreamWriter to write the comments to the log file.</param>
    private static void WriteCommentsToLog(string filePath, StreamWriter writer)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
        {
            if (doc.MainDocumentPart.WordprocessingCommentsPart != null)
            {
                foreach (var comment in doc.MainDocumentPart.WordprocessingCommentsPart.Comments.Elements<Comment>())
                {
                    writer.WriteLine($"Author: {comment.Author}");
                    writer.WriteLine($"Date: {comment.Date}");
                    writer.WriteLine($"Comment: {comment.InnerText}");
                    writer.WriteLine(); // Add an empty line between comments for readability
                }
            }
        }
    }


    /// <summary>
    /// Extracts comments from all Word documents in a directory and writes them to a CSV file.
    /// </summary>
    /// <param name="directory">The directory containing the Word documents.</param>
    /// <param name="csvFilePath">The file path for the CSV file.</param>
    private static void ExtractCommentsToCSV(string directory, string csvFilePath)
    {
        using (StreamWriter writer = new StreamWriter(csvFilePath, false, new UTF8Encoding(true)))
        {
            writer.WriteLine("FileName,Author,Date,Comment");

            var files = Directory.GetFiles(directory, "*.docx");
            foreach (var file in files)
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(file, false))
                {
                    if (doc.MainDocumentPart.WordprocessingCommentsPart != null)
                    {
                        foreach (var comment in doc.MainDocumentPart.WordprocessingCommentsPart.Comments.Elements<Comment>())
                        {
                            string commentText = comment.InnerText.Replace("\"", "\"\"");
                            writer.WriteLine($"{Path.GetFileName(file)},{comment.Author},{comment.Date},\"{commentText}\"");
                        }
                    }
                }
            }
        }
    }

}
