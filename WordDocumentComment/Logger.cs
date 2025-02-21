using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

public static class Logger
{
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
}
