using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace ExchangeEwsMcp.Exchange;

/// <summary>Turns attachments into text the model can read: plain text formats, PDF and Word (.docx).</summary>
internal static class AttachmentText
{
    /// <summary>Returns the text, an empty string if the file has no extractable text, or null for unsupported formats.</summary>
    public static string? Extract(byte[] content, string contentType, string? name)
    {
        var ext = Path.GetExtension(name ?? "").ToLowerInvariant();
        if (ext == ".pdf" || contentType.Contains("pdf", StringComparison.OrdinalIgnoreCase))
        {
            return Pdf(content);
        }
        if (ext == ".docx" || contentType.Contains("wordprocessingml", StringComparison.OrdinalIgnoreCase))
        {
            return Docx(content);
        }
        if (contentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("json", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("xml", StringComparison.OrdinalIgnoreCase)
            || ext is ".txt" or ".csv" or ".json" or ".xml" or ".md" or ".html" or ".htm" or ".ics" or ".eml" or ".log" or ".yaml" or ".yml")
        {
            return Encoding.UTF8.GetString(content);
        }
        return null;
    }

    private static string Pdf(byte[] content)
    {
        using var document = PdfDocument.Open(content);
        var text = new StringBuilder();
        foreach (var page in document.GetPages())
        {
            text.Append("--- Page ").Append(page.Number).AppendLine(" ---");
            text.AppendLine(ContentOrderTextExtractor.GetText(page));
        }
        // Scanned PDFs contain only images: report "no text" instead of page markers.
        return text.ToString().Split('\n').All(l => l.Trim().Length == 0 || l.StartsWith("--- Page ")) ? "" : text.ToString();
    }

    private static string Docx(byte[] content)
    {
        using var zip = new ZipArchive(new MemoryStream(content), ZipArchiveMode.Read);
        var entry = zip.GetEntry("word/document.xml");
        if (entry is null) return "";
        using var stream = entry.Open();
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var paragraphs = XDocument.Load(stream).Descendants(w + "p")
            .Select(p => string.Concat(p.Descendants().Select(e => e.Name == w + "t" ? e.Value : e.Name == w + "tab" ? "\t" : "")));
        return string.Join('\n', paragraphs);
    }
}
