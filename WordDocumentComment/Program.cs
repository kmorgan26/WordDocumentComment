using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Enter the directory path containing Word documents:");
        string directoryPath = Console.ReadLine();

        if (Directory.Exists(directoryPath))
        {
            string[] files = Directory.GetFiles(directoryPath, "*.docx");

            if (files.Length < 2)
            {
                Console.WriteLine("At least two Word documents are required for comparison.");
                return;
            }

            string file1 = files[0];
            string file2 = files[1];

            Console.WriteLine($"Comparing {file1} and {file2}...");

            // Compare document content
            var contentDifferences = CompareWordDocuments(file1, file2);
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

            // Compare comments
            var commentDifferences = CompareComments(file1, file2);
            if (commentDifferences.Count == 0)
            {
                Console.WriteLine("No differences in comments found.");
            }
            else
            {
                Console.WriteLine("Differences in comments found:");
                foreach (var diff in commentDifferences)
                {
                    Console.WriteLine(diff);
                }
            }
        }
        else
        {
            Console.WriteLine("Directory does not exist.");
        }
    }

    static List<string> CompareWordDocuments(string filePath1, string filePath2)
    {
        var differences = new List<string>();

        using (WordprocessingDocument doc1 = WordprocessingDocument.Open(filePath1, false))
        using (WordprocessingDocument doc2 = WordprocessingDocument.Open(filePath2, false))
        {
            var body1 = doc1.MainDocumentPart.Document.Body;
            var body2 = doc2.MainDocumentPart.Document.Body;

            differences.AddRange(CompareElements(body1.Elements(), body2.Elements()));
        }

        return differences;
    }

    static List<string> CompareComments(string filePath1, string filePath2)
    {
        var differences = new List<string>();

        using (WordprocessingDocument doc1 = WordprocessingDocument.Open(filePath1, false))
        using (WordprocessingDocument doc2 = WordprocessingDocument.Open(filePath2, false))
        {
            var comments1 = GetComments(doc1);
            var comments2 = GetComments(doc2);

            differences.AddRange(CompareCommentLists(comments1, comments2));
        }

        return differences;
    }

    static List<Comment> GetComments(WordprocessingDocument doc)
    {
        var comments = new List<Comment>();

        if (doc.MainDocumentPart.WordprocessingCommentsPart != null)
        {
            foreach (var comment in doc.MainDocumentPart.WordprocessingCommentsPart.Comments.Elements<Comment>())
            {
                comments.Add(comment);
            }
        }

        return comments;
    }
    static List<string> CompareCommentLists(List<Comment> comments1, List<Comment> comments2)
    {
        var differences = new List<string>();

        int index = 0;
        while (index < comments1.Count && index < comments2.Count)
        {
            var comment1 = comments1[index];
            var comment2 = comments2[index];

            if (comment1.InnerText != comment2.InnerText)
            {
                differences.Add($"Difference in comment at index {index}:");
                differences.Add($"File 1: {comment1.InnerText}");
                differences.Add($"File 2: {comment2.InnerText}");
            }

            index++;
        }

        // Check if one document has more comments than the other
        if (index < comments1.Count)
        {
            differences.Add("File 1 has additional comments:");
            while (index < comments1.Count)
            {
                differences.Add(comments1[index].InnerText);
                index++;
            }
        }
        else if (index < comments2.Count)
        {
            differences.Add("File 2 has additional comments:");
            while (index < comments2.Count)
            {
                differences.Add(comments2[index].InnerText);
                index++;
            }
        }

        return differences;
    }

    static List<string> CompareElements(IEnumerable<OpenXmlElement> elements1, IEnumerable<OpenXmlElement> elements2)
    {
        var differences = new List<string>();
        var enumerator1 = elements1.GetEnumerator();
        var enumerator2 = elements2.GetEnumerator();

        int index = 0;
        while (enumerator1.MoveNext() && enumerator2.MoveNext())
        {
            var element1 = enumerator1.Current;
            var element2 = enumerator2.Current;

            if (element1.InnerText != element2.InnerText)
            {
                differences.Add($"Difference at index {index}:");
                differences.Add($"File 1: {element1.InnerText}");
                differences.Add($"File 2: {element2.InnerText}");
            }

            index++;
        }

        // Check if one document has more elements than the other
        if (enumerator1.MoveNext())
        {
            differences.Add("File 1 has additional content:");
            while (enumerator1.MoveNext())
            {
                differences.Add(enumerator1.Current.InnerText);
            }
        }
        else if (enumerator2.MoveNext())
        {
            differences.Add("File 2 has additional content:");
            while (enumerator2.MoveNext())
            {
                differences.Add(enumerator2.Current.InnerText);
            }
        }

        return differences;
    }
}