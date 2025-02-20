using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Enter the path to the first directory:");
        //string directory1 = Console.ReadLine();
        string directory1 = "D:\\WordDocs\\folder1";

        Console.WriteLine("Enter the path to the second directory:");
        //string directory2 = Console.ReadLine();
        string directory2 = "D:\\WordDocs\\folder2";

        Console.WriteLine("Enter the path to the output directory for consolidated comments:");
        //string outputDirectory = Console.ReadLine();
        string outputDirectory = "D:\\WordDocs\\folder3";

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

                    // Consolidate comments into a new document
                    string outputFilePath = Path.Combine(outputDirectory, $"ConsolidatedComments_{fileName}");
                    ConsolidateComments(file1, file2, outputFilePath);

                    Console.WriteLine($"Comments consolidated into {outputFilePath}");
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

    static void ConsolidateComments(string filePath1, string filePath2, string outputFilePath)
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
    static void CopyContentFromDocument(string filePath, MainDocumentPart mainPart)
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

    static void AddCommentsFromDocument(string filePath, WordprocessingCommentsPart commentsPart, string sourceDocumentName)
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
                    //newComment.Append(new Paragraph(new Run(new Text($"From {sourceDocumentName}: {comment.InnerText}"))));
                    newComment.Append(new Paragraph(new Run(new Text(comment.InnerText))));
                    commentsPart.Comments.Append(newComment);
                }
            }
        }
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