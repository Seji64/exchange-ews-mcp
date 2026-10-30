using System.IO.Compression;
using System.Text;
using ExchangeEwsMcp.Exchange;
using UglyToad.PdfPig.Fonts.Standard14Fonts;
using UglyToad.PdfPig.Writer;

namespace ExchangeEwsMcp.Tests;

public class AttachmentTextTests
{
    [Fact]
    public void Pdf_text_is_extracted()
    {
        var builder = new PdfDocumentBuilder();
        var font = builder.AddStandard14Font(Standard14Font.Helvetica);
        builder.AddPage(UglyToad.PdfPig.Content.PageSize.A4).AddText("Order confirmation DEC-QUO-719923", 12, new UglyToad.PdfPig.Core.PdfPoint(50, 700), font);

        var text = AttachmentText.Extract(builder.Build(), "application/pdf", "AB.pdf");

        Assert.Contains("Order confirmation DEC-QUO-719923", text);
    }

    [Fact]
    public void Docx_text_is_extracted()
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open(), Encoding.UTF8);
            writer.Write("""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
                  <w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t xml:space="preserve"> world</w:t></w:r></w:p>
                  <w:p><w:r><w:t>Second line</w:t></w:r></w:p>
                </w:body></w:document>
                """);
        }

        var text = AttachmentText.Extract(stream.ToArray(), "application/octet-stream", "letter.docx");

        Assert.Equal("Hello world\nSecond line", text);
    }

    [Fact]
    public void Unknown_binary_formats_are_not_returned()
    {
        Assert.Null(AttachmentText.Extract([1, 2, 3], "application/zip", "archive.zip"));
    }
}
