Console.WriteLine("Select a feature:");
Console.WriteLine("1. Compare and Consolidate Comments");
Console.WriteLine("2. Comments to CSV");
string choice = Console.ReadLine();

switch (choice)
{
    case "1":
        CommentConsolidator.CompareAndConsolidateComments();
        break;
    case "2":
        Logger.CommentsToCSV();
        break;
    default:
        Console.WriteLine("Invalid choice. Exiting.");
        break;
}