class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("This program compares Word documents from two directories, consolidates their content and comments into a new document, and logs the differences and comments.");
        Console.WriteLine("Please provide the paths to the two directories containing the Word documents and the output directory for the consolidated document and log files.");

        Console.WriteLine("Enter the path to the first directory:");
        string directory1 = Console.ReadLine();
        //string directory1 = "D:\\WordDocs\\folder1";

        Console.WriteLine("Enter the path to the second directory:");
        string directory2 = Console.ReadLine();
        //string directory2 = "D:\\WordDocs\\folder2";

        Console.WriteLine("Enter the path to the output directory for consolidated comments:");
        string outputDirectory = Console.ReadLine();
        //string outputDirectory = "D:\\WordDocs\\folder3";

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
