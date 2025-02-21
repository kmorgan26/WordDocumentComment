using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

public static class Logger
{
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
