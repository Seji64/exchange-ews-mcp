using System.ComponentModel;
using ExchangeEwsMcp.Exchange;
using Microsoft.Exchange.WebServices.Data;
using ModelContextProtocol.Server;

namespace ExchangeEwsMcp.Tools;

/// <summary>Only registered when Exchange:EnableDownloadLinks is true (needs a file tool such as Fileshed in the client).</summary>
[McpServerToolType]
public sealed class DownloadTools(EwsContext ews, DownloadStore downloads)
{
    [McpServerTool(Name = "prepare_attachment_download", ReadOnly = true, OpenWorld = false)]
    [Description("""
        Makes an attachment available as a file download without putting its content into the conversation.
        Returns a single-use URL (valid 10 minutes) that is only reachable by the chat server, e.g. for a file tool
        such as Fileshed to fetch (curl) and offer to the user. Use get_attachment instead to read the content.
        """)]
    public async Task<string> PrepareAttachmentDownload(
        [Description("The item id.")] string itemId,
        [Description("Attachment index (from read_mail), file name, or attachment id.")] string attachment,
        [Description(MailTools.ItemMailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var (_, match) = await ews.FindAttachmentAsync(itemId, attachment);

        DownloadStore.File file;
        if (match is ItemAttachment itemAttachment)
        {
            // Attached e-mails are offered as .eml files.
            await itemAttachment.Load(CancellationToken.None, ItemSchema.MimeContent);
            var name = Path.HasExtension(match.Name) ? match.Name : match.Name + ".eml";
            file = new DownloadStore.File(name, "message/rfc822", itemAttachment.Item?.MimeContent?.Content ?? []);
        }
        else
        {
            var fileAttachment = (FileAttachment)match;
            await fileAttachment.Load();
            file = new DownloadStore.File(fileAttachment.Name, fileAttachment.ContentType ?? "application/octet-stream", fileAttachment.Content ?? []);
        }

        var url = downloads.Add(file);
        return $"Download URL for '{file.Name}' ({file.ContentType}, {file.Content.Length} bytes), single use, valid 10 minutes: {url}";
    }
}
