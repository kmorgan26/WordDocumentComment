using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

/// <summary>
/// Provides methods to compare Word documents and identify differences.
/// </summary>
public static class DocumentComparer
{
    /// <summary>
    /// Compares the content of two Word documents and returns a list of differences.
    /// </summary>
    /// <param name="filePath1">The file path of the first Word document.</param>
    /// <param name="filePath2">The file path of the second Word document.</param>
    /// <returns>A list of differences between the two documents.</returns>
    public static List<string> CompareWordDocuments(string filePath1, string filePath2)
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

    /// <summary>
    /// Compares the elements of two Word document bodies and returns a list of differences.
    /// </summary>
    /// <param name="elements1">The elements of the first document body.</param>
    /// <param name="elements2">The elements of the second document body.</param>
    /// <returns>A list of differences between the elements of the two document bodies.</returns>
    private static List<string> CompareElements(IEnumerable<OpenXmlElement> elements1, IEnumerable<OpenXmlElement> elements2)
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
