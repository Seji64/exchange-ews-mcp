using System.ComponentModel;
using System.Text;
using ExchangeEwsMcp.Configuration;
using ExchangeEwsMcp.Exchange;
using Microsoft.Exchange.WebServices.Data;
using ModelContextProtocol;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using static ExchangeEwsMcp.Exchange.EwsHelpers;

namespace ExchangeEwsMcp.Tools;

[McpServerToolType]
public sealed class MailTools(EwsContext ews)
{
    internal const string MailboxDescription =
        "Optional shared or delegated mailbox (e-mail address, alias or name). Omit for the user's own mailbox.";

    private const string SignatureDescription =
        "Append the Outlook signature (from Outlook on the web) below the text. Ask the user whether to include it before creating the draft.";

    internal const string ItemMailboxDescription =
        "Shared or delegated mailbox the item is in (e-mail address, alias or name). Omit for the user's own mailbox.";

    private static readonly PropertySet SummaryProperties = new(BasePropertySet.IdOnly,
        ItemSchema.Subject, EmailMessageSchema.From, ItemSchema.DisplayTo, ItemSchema.DateTimeReceived,
        EmailMessageSchema.IsRead, ItemSchema.HasAttachments, ItemSchema.Importance, ItemSchema.IsDraft, ItemSchema.Preview);

    [McpServerTool(Name = "list_folders", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Lists the mail folders of a mailbox with item and unread counts.")]
    public async Task<IReadOnlyList<FolderInfo>> ListFolders(
        [Description(MailboxDescription)] string? mailbox = null,
        [Description("Include non-mail folders (calendar, contacts, tasks, ...). Default false.")] bool includeAllFolderTypes = false)
    {
        await ews.UseMailboxAsync(mailbox);
        var root = ews.WellKnown(WellKnownFolderName.MsgFolderRoot);
        var view = new FolderView(500)
        {
            Traversal = FolderTraversal.Deep,
            PropertySet = new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName, FolderSchema.ParentFolderId,
                FolderSchema.TotalCount, FolderSchema.ChildFolderCount, FolderSchema.FolderClass, FolderSchema.UnreadCount),
        };
        var folders = (await ews.Service.FindFolders(root, view)).Folders;
        var byId = folders.ToDictionary(f => f.Id.UniqueId!);

        string PathOf(Folder f)
        {
            var parts = new Stack<string>();
            for (var current = f; current is not null; current = byId.GetValueOrDefault(current.ParentFolderId.UniqueId ?? ""))
            {
                parts.Push(current.DisplayName);
                if (parts.Count > 20) break;
            }
            return string.Join('/', parts);
        }

        return folders
            .Where(f => includeAllFolderTypes || f.FolderClass is null or "" || f.FolderClass.StartsWith("IPF.Note", StringComparison.OrdinalIgnoreCase))
            .Select(f => new FolderInfo(f.Id.UniqueId!, f.DisplayName, PathOf(f), f.TotalCount, Try(() => f.UnreadCount), f.ChildFolderCount))
            .OrderBy(f => f.Path, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    [McpServerTool(Name = "search_mail", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("""
        Searches e-mails in a folder (default: Inbox), newest first. Returns ids, subjects, senders and a short preview.
        Use read_mail with an id to get the full message. 'query' uses Exchange search (AQS/KQL) such as
        'budget report' or 'from:alice subject:invoice'. Without 'query', the structured filters are applied exactly.
        """)]
    public async Task<SearchResult<MailSummary>> SearchMail(
        [Description("Free-text search (Exchange AQS/KQL). Optional.")] string? query = null,
        [Description("Folder: well-known name (inbox, sentitems, drafts, deleteditems, junkemail, archive), a path like 'Inbox/Projects' or a folder id. Default inbox.")] string? folder = null,
        [Description("Only mails from this sender (name or address).")] string? from = null,
        [Description("Only mails whose subject contains this text.")] string? subject = null,
        [Description("Only mails received on/after this date/time (ISO 8601).")] string? since = null,
        [Description("Only mails received before this date/time (ISO 8601).")] string? until = null,
        [Description("Only unread mails.")] bool unreadOnly = false,
        [Description("Only mails with attachments.")] bool hasAttachments = false,
        [Description("Maximum number of results (default 20).")] int maxResults = 20,
        [Description("Number of results to skip, for paging.")] int offset = 0,
        [Description(MailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var folderId = await ews.ResolveFolderAsync(folder, WellKnownFolderName.Inbox);
        var view = new ItemView(Math.Clamp(maxResults, 1, ews.Options.MaxResults), Math.Max(0, offset)) { PropertySet = SummaryProperties };
        view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);

        FindItemsResults<Item> results;
        if (!string.IsNullOrWhiteSpace(query) || !string.IsNullOrWhiteSpace(from))
        {
            results = await ews.Service.FindItems(folderId, BuildAqs(query, from, subject, since, until, unreadOnly, hasAttachments), view);
        }
        else
        {
            var filters = new List<SearchFilter>();
            if (!string.IsNullOrWhiteSpace(subject))
                filters.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, subject, ContainmentMode.Substring, ComparisonMode.IgnoreCase));
            if (!string.IsNullOrWhiteSpace(since))
                filters.Add(new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, ews.ParseDateTime(since, nameof(since))));
            if (!string.IsNullOrWhiteSpace(until))
                filters.Add(new SearchFilter.IsLessThan(ItemSchema.DateTimeReceived, ews.ParseDateTime(until, nameof(until))));
            if (unreadOnly)
                filters.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            if (hasAttachments)
                filters.Add(new SearchFilter.IsEqualTo(ItemSchema.HasAttachments, true));

            results = filters.Count == 0
                ? await ews.Service.FindItems(folderId, view)
                : await ews.Service.FindItems(folderId, new SearchFilter.SearchFilterCollection(LogicalOperator.And, filters), view);
        }

        var items = results.Items.Select(ToSummary).ToList();
        return new SearchResult<MailSummary>(results.TotalCount, results.MoreAvailable, items);
    }

    internal static string BuildAqs(string? query, string? from, string? subject, string? since, string? until, bool unreadOnly, bool hasAttachments)
    {
        static string Quote(string v) => "\"" + v.Replace("\"", "") + "\"";
        static string Day(string v) => DateTime.TryParse(v, System.Globalization.CultureInfo.InvariantCulture,
            System.Globalization.DateTimeStyles.None, out var d) ? d.ToString("yyyy-MM-dd") : v;

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(query)) parts.Add(query.Trim());
        if (!string.IsNullOrWhiteSpace(from)) parts.Add("from:" + Quote(from.Trim()));
        if (!string.IsNullOrWhiteSpace(subject)) parts.Add("subject:" + Quote(subject.Trim()));
        if (!string.IsNullOrWhiteSpace(since)) parts.Add("received>=" + Day(since));
        if (!string.IsNullOrWhiteSpace(until)) parts.Add("received<" + Day(until));
        if (unreadOnly) parts.Add("isread:false");
        if (hasAttachments) parts.Add("hasattachment:true");
        return string.Join(' ', parts);
    }

    private MailSummary ToSummary(Item item)
    {
        var message = item as EmailMessage;
        return new MailSummary(
            item.Id.UniqueId!,
            Try(() => item.Subject),
            message is null ? null : ToPerson(Try(() => message.From)),
            Try(() => item.DisplayTo),
            ews.TryFormat(() => item.DateTimeReceived),
            message is null || Try(() => message.IsRead),
            Try(() => item.HasAttachments),
            Try(() => item.Importance).ToString(),
            Try(() => item.IsDraft),
            NullIfEmpty(Try(() => item.Preview)));
    }

    [McpServerTool(Name = "read_mail", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Reads a single e-mail including body, recipients and the list of attachments.")]
    public async Task<MailDetails> ReadMail(
        [Description("The item id from search_mail.")] string id,
        [Description("Body format: 'text' (default, HTML converted to plain text by Exchange) or 'html'.")] string format = "text",
        [Description(MailTools.ItemMailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var props = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Body, EmailMessageSchema.BccRecipients)
        {
            RequestedBodyType = string.Equals(format, "html", StringComparison.OrdinalIgnoreCase) ? BodyType.HTML : BodyType.Text,
        };
        var item = await Item.Bind(ews.Service, new ItemId(id), props);
        var message = item as EmailMessage;

        var (body, truncated) = Truncate(Try(() => item.Body?.Text), ews.Options.MaxBodyLength);
        return new MailDetails(
            item.Id.UniqueId!,
            Try(() => item.Subject),
            message is null ? null : ToPerson(Try(() => message.From)),
            ToPeople(message is null ? null : Try(() => message.ToRecipients)),
            ToPeople(message is null ? null : Try(() => message.CcRecipients)),
            message is not null && Try(() => message.BccRecipients) is { Count: > 0 } bcc ? ToPeople(bcc) : null,
            ews.TryFormat(() => item.DateTimeSent),
            ews.TryFormat(() => item.DateTimeReceived),
            message is null || Try(() => message.IsRead),
            Try(() => item.IsDraft),
            Try(() => item.Importance).ToString(),
            Try(() => item.Categories) is { Count: > 0 } categories ? categories.ToList() : null,
            Try(() => item.ConversationId?.UniqueId),
            Try(() => item.Body?.BodyType) == BodyType.HTML ? "html" : "text",
            body,
            truncated,
            ListAttachments(item));
    }

    internal static IReadOnlyList<AttachmentInfo> ListAttachments(Item item) =>
        Try(() => item.Attachments)?.Select((a, i) => new AttachmentInfo(
            i, a.Id, a.Name, NullIfEmpty(a.ContentType), Try(() => a.Size), Try(() => a.IsInline), a is ItemAttachment)).ToList()
        ?? [];

    [McpServerTool(Name = "get_attachment", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("""
        Downloads an attachment of an e-mail or calendar item. Text-based files (txt, csv, json, xml, html, md, ics, eml)
        are returned as text, images as image content. Other binary files are only described unless includeBase64 is true.
        """)]
    public async Task<CallToolResult> GetAttachment(
        [Description("The item id.")] string itemId,
        [Description("Attachment index (from read_mail), file name, or attachment id.")] string attachment,
        [Description("Return other binary files as base64. Default false.")] bool includeBase64 = false,
        [Description(MailTools.ItemMailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var item = await Item.Bind(ews.Service, new ItemId(itemId), new PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments));
        var attachments = item.Attachments.ToList();
        var match = (int.TryParse(attachment, out var index) && index >= 0 && index < attachments.Count ? attachments[index] : null)
            ?? attachments.FirstOrDefault(a => a.Id == attachment)
            ?? attachments.FirstOrDefault(a => string.Equals(a.Name, attachment, StringComparison.OrdinalIgnoreCase))
            ?? throw new McpException($"Attachment '{attachment}' not found. Available: {string.Join(", ", attachments.Select((a, i) => $"[{i}] {a.Name}"))}");

        if (match.Size > ews.Options.MaxAttachmentBytes)
        {
            throw new McpException($"Attachment '{match.Name}' is {match.Size} bytes, which exceeds the limit of {ews.Options.MaxAttachmentBytes} bytes.");
        }

        if (match is ItemAttachment itemAttachment)
        {
            await itemAttachment.Load(CancellationToken.None, ItemSchema.MimeContent);
            var text = Encoding.UTF8.GetString(itemAttachment.Item?.MimeContent?.Content ?? []);
            return TextResult($"Attached item '{match.Name}' (MIME):\n\n{Truncate(text, ews.Options.MaxBodyLength).Text}");
        }

        var file = (FileAttachment)match;
        await file.Load();
        var content = file.Content ?? [];
        var contentType = file.ContentType ?? "application/octet-stream";

        if (IsText(contentType, file.Name))
        {
            var (text, truncated) = Truncate(Encoding.UTF8.GetString(content), ews.Options.MaxBodyLength);
            return TextResult($"Attachment '{file.Name}' ({contentType}, {content.Length} bytes){(truncated ? ", truncated" : "")}:\n\n{text}");
        }
        if (contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase))
        {
            return new CallToolResult
            {
                Content =
                [
                    new TextContentBlock { Text = $"Attachment '{file.Name}' ({contentType}, {content.Length} bytes)." },
                    ImageContentBlock.FromBytes(content, contentType),
                ],
            };
        }
        if (includeBase64)
        {
            return TextResult($"Attachment '{file.Name}' ({contentType}, {content.Length} bytes), base64:\n{Convert.ToBase64String(content)}");
        }
        return TextResult($"Attachment '{file.Name}' is a binary file ({contentType}, {content.Length} bytes) whose content cannot be displayed as text.");
    }

    private static bool IsText(string contentType, string? name)
    {
        var ext = Path.GetExtension(name ?? "").ToLowerInvariant();
        return contentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("json", StringComparison.OrdinalIgnoreCase)
            || contentType.Contains("xml", StringComparison.OrdinalIgnoreCase)
            || ext is ".txt" or ".csv" or ".json" or ".xml" or ".md" or ".html" or ".htm" or ".ics" or ".eml" or ".log" or ".yaml" or ".yml";
    }

    private static CallToolResult TextResult(string text) => new() { Content = [new TextContentBlock { Text = text }] };

    // ---- Drafts --------------------------------------------------------------------------------

    [McpServerTool(Name = "create_draft", Destructive = false, OpenWorld = false)]
    [Description("""
        Creates a new e-mail as draft in the Drafts folder. Nothing is sent; the user reviews and sends it in Outlook.
        Recipients may be e-mail addresses or names (names are resolved via the address book and must be unambiguous).
        """)]
    public async Task<DraftResult> CreateDraft(
        [Description("Subject.")] string subject,
        [Description("Body text. Plain text by default; set bodyFormat='html' for HTML.")] string body,
        [Description("To recipients.")] string[]? to = null,
        [Description("Cc recipients.")] string[]? cc = null,
        [Description("Bcc recipients.")] string[]? bcc = null,
        [Description("'text' (default) or 'html'.")] string bodyFormat = "text",
        [Description(SignatureDescription)] bool includeSignature = false,
        [Description("'low', 'normal' (default) or 'high'.")] string importance = "normal",
        [Description("Optional shared mailbox (e-mail address, alias or name) to create the draft in (it is sent from that mailbox).")] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var message = new EmailMessage(ews.Service)
        {
            Subject = subject,
            Body = await BodyAsync(body, bodyFormat, includeSignature),
            Importance = ParseImportance(importance),
        };
        message.ToRecipients.AddRange(await ResolveRecipientsAsync(ews.Service, to));
        message.CcRecipients.AddRange(await ResolveRecipientsAsync(ews.Service, cc));
        message.BccRecipients.AddRange(await ResolveRecipientsAsync(ews.Service, bcc));
        if (ews.Mailbox is not null)
        {
            message.From = new EmailAddress(ews.Mailbox);
        }
        await message.Save(ews.WellKnown(WellKnownFolderName.Drafts));
        return new DraftResult(message.Id.UniqueId!, subject, "Drafts", DraftSavedMessage());
    }

    [McpServerTool(Name = "create_reply_draft", Destructive = false, OpenWorld = false)]
    [Description("Creates a reply (or reply-all) to an e-mail as draft, keeping the quoted original. Nothing is sent.")]
    public async Task<DraftResult> CreateReplyDraft(
        [Description("Id of the e-mail to reply to.")] string id,
        [Description("Reply text placed above the quoted original.")] string body,
        [Description("Reply to all recipients. Default false.")] bool replyAll = false,
        [Description("'text' (default) or 'html'.")] string bodyFormat = "text",
        [Description(SignatureDescription)] bool includeSignature = false,
        [Description("Optional shared mailbox (e-mail address, alias or name) the original mail is in (draft is saved there).")] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var original = await EmailMessage.Bind(ews.Service, new ItemId(id));
        var reply = original.CreateReply(replyAll);
        reply.BodyPrefix = await BodyAsync(body, bodyFormat, includeSignature);
        var draft = await reply.Save(ews.WellKnown(WellKnownFolderName.Drafts));
        await draft.Load(new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject));
        return new DraftResult(draft.Id.UniqueId!, draft.Subject, "Drafts", DraftSavedMessage());
    }

    [McpServerTool(Name = "create_forward_draft", Destructive = false, OpenWorld = false)]
    [Description("Creates a forward of an e-mail (including its attachments) as draft. Nothing is sent.")]
    public async Task<DraftResult> CreateForwardDraft(
        [Description("Id of the e-mail to forward.")] string id,
        [Description("Recipients (addresses or names).")] string[] to,
        [Description("Text placed above the forwarded message.")] string? body = null,
        [Description("Cc recipients.")] string[]? cc = null,
        [Description("'text' (default) or 'html'.")] string bodyFormat = "text",
        [Description(SignatureDescription)] bool includeSignature = false,
        [Description("Optional shared mailbox (e-mail address, alias or name) the original mail is in (draft is saved there).")] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var original = await EmailMessage.Bind(ews.Service, new ItemId(id));
        var forward = original.CreateForward();
        forward.ToRecipients.AddRange(await ResolveRecipientsAsync(ews.Service, to));
        forward.CcRecipients.AddRange(await ResolveRecipientsAsync(ews.Service, cc));
        if (!string.IsNullOrEmpty(body))
        {
            forward.BodyPrefix = await BodyAsync(body, bodyFormat, includeSignature);
        }
        var draft = await forward.Save(ews.WellKnown(WellKnownFolderName.Drafts));
        await draft.Load(new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject));
        return new DraftResult(draft.Id.UniqueId!, draft.Subject, "Drafts", DraftSavedMessage());
    }

    [McpServerTool(Name = "update_draft", Destructive = false, OpenWorld = false)]
    [Description("""
        Updates an existing draft. Only the given fields change. Recipient lists replace the existing ones.
        Note: 'body' replaces the whole body, including quoted text of a reply; to rewrite a reply,
        delete the draft and use create_reply_draft again.
        """)]
    public async Task<DraftResult> UpdateDraft(
        [Description("Draft id.")] string id,
        [Description("New subject.")] string? subject = null,
        [Description("New body (replaces the existing body).")] string? body = null,
        [Description("New To recipients.")] string[]? to = null,
        [Description("New Cc recipients.")] string[]? cc = null,
        [Description("New Bcc recipients.")] string[]? bcc = null,
        [Description("'text' (default) or 'html'.")] string bodyFormat = "text",
        [Description(SignatureDescription)] bool includeSignature = false,
        [Description("'low', 'normal' or 'high'.")] string? importance = null,
        [Description(MailTools.ItemMailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var draft = await BindDraftAsync(id);
        if (subject is not null) draft.Subject = subject;
        if (body is not null) draft.Body = await BodyAsync(body, bodyFormat, includeSignature);
        if (importance is not null) draft.Importance = ParseImportance(importance);
        if (to is not null) { draft.ToRecipients.Clear(); draft.ToRecipients.AddRange(await ResolveRecipientsAsync(ews.Service, to)); }
        if (cc is not null) { draft.CcRecipients.Clear(); draft.CcRecipients.AddRange(await ResolveRecipientsAsync(ews.Service, cc)); }
        if (bcc is not null) { draft.BccRecipients.Clear(); draft.BccRecipients.AddRange(await ResolveRecipientsAsync(ews.Service, bcc)); }
        await draft.Update(ConflictResolutionMode.AlwaysOverwrite);
        return new DraftResult(draft.Id.UniqueId!, draft.Subject, "Drafts", "Draft updated. It has not been sent.");
    }

    [McpServerTool(Name = "add_attachment_to_draft", Destructive = false, OpenWorld = false)]
    [Description("Adds a file attachment to a draft. Provide the content either as text or base64.")]
    public async Task<DraftResult> AddAttachmentToDraft(
        [Description("Draft id.")] string id,
        [Description("File name including extension, e.g. 'notes.txt'.")] string fileName,
        [Description("File content as text (UTF-8).")] string? contentText = null,
        [Description("File content as base64.")] string? contentBase64 = null,
        [Description(MailTools.ItemMailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        byte[] content;
        if (contentBase64 is not null)
        {
            try
            {
                content = Convert.FromBase64String(contentBase64);
            }
            catch (FormatException)
            {
                throw new McpException("contentBase64 is not valid base64.");
            }
        }
        else if (contentText is not null)
        {
            content = Encoding.UTF8.GetBytes(contentText);
        }
        else
        {
            throw new McpException("Either contentText or contentBase64 is required.");
        }
        if (content.Length > ews.Options.MaxAttachmentBytes)
        {
            throw new McpException($"Attachment exceeds the limit of {ews.Options.MaxAttachmentBytes} bytes.");
        }

        var draft = await BindDraftAsync(id);
        draft.Attachments.AddFileAttachment(fileName, content);
        await draft.Update(ConflictResolutionMode.AlwaysOverwrite);
        return new DraftResult(draft.Id.UniqueId!, draft.Subject, "Drafts", $"Attachment '{fileName}' added. The draft has not been sent.");
    }

    [McpServerTool(Name = "delete_draft", Destructive = true, OpenWorld = false)]
    [Description("Deletes a draft (moves it to Deleted Items). Only works for unsent drafts.")]
    public async Task<string> DeleteDraft([Description("Draft id.")] string id,
        [Description(MailTools.ItemMailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var draft = await BindDraftAsync(id);
        await draft.Delete(DeleteMode.MoveToDeletedItems);
        return "Draft moved to Deleted Items.";
    }

    private async Task<MessageBody> BodyAsync(string? body, string bodyFormat, bool includeSignature)
    {
        var result = CreateBody(body, bodyFormat);
        if (!includeSignature) return result;

        // Outlook on the web keeps the signature in the mailbox (where tools like AutoSignature.NET write it);
        // Exchange does not add it to drafts created via EWS.
        var options = await UserConfiguration.Bind(ews.Service, "OWA.UserOptions", ews.WellKnown(WellKnownFolderName.Root),
            UserConfigurationProperties.Dictionary);
        var signature =
            options.Dictionary.TryGetValue("signaturehtml", out var html) && html is string h && h.Trim().Length > 0 ? h
            : options.Dictionary.TryGetValue("signaturetext", out var text) && text is string t && t.Trim().Length > 0 ? CreateBody(t, "text").Text
            : throw new McpException("No Outlook on the web signature is configured for this mailbox. The draft was not created.");
        return new MessageBody(BodyType.HTML, result.Text + "<br>" + signature);
    }

    private async Task<EmailMessage> BindDraftAsync(string id)
    {
        var draft = await EmailMessage.Bind(ews.Service, new ItemId(id));
        if (!draft.IsDraft)
        {
            throw new McpException("This item is not a draft; only drafts can be changed.");
        }
        return draft;
    }

    private string DraftSavedMessage() => ews.Options.AllowSend
        ? "Draft saved in the Drafts folder. It has NOT been sent. Use send_draft only after the user explicitly confirmed."
        : "Draft saved in the Drafts folder. It has NOT been sent; the user can review and send it from Outlook.";

    private static Importance ParseImportance(string? value) => value?.ToLowerInvariant() switch
    {
        "low" => Importance.Low,
        "high" => Importance.High,
        _ => Importance.Normal,
    };

    // ---- Organizing ----------------------------------------------------------------------------

    [McpServerTool(Name = "move_mail", Destructive = false, OpenWorld = false)]
    [Description("Moves one or more items to another folder (e.g. 'archive', 'deleteditems', 'Inbox/Projects').")]
    public async Task<string> MoveMail(
        [Description("Item ids.")] string[] ids,
        [Description("Target folder (well-known name, path or folder id).")] string folder,
        [Description(MailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        var target = await ews.ResolveFolderAsync(folder, WellKnownFolderName.Inbox);
        if (ews.Options.AccessLevel != AccessLevel.Full && await IsDeletedItemsAsync(target))
        {
            throw new McpException("Deleting is disabled on this server, so items cannot be moved to Deleted Items.");
        }
        var responses = await ews.Service.MoveItems(ids.Select(i => new ItemId(i)), target, false);
        var failed = responses.Where(r => r.Result != ServiceResult.Success).ToList();
        return failed.Count == 0
            ? $"Moved {ids.Length} item(s) to '{folder}'."
            : $"Moved {ids.Length - failed.Count} item(s); {failed.Count} failed: {string.Join("; ", failed.Select(f => f.ErrorMessage))}";
    }

    // shortcut: subfolders of Deleted Items are not detected; walk ParentFolderId if that matters.
    private async Task<bool> IsDeletedItemsAsync(FolderId target)
    {
        if (target.FolderName is not null)
        {
            return target.FolderName is WellKnownFolderName.DeletedItems or WellKnownFolderName.RecoverableItemsRoot
                or WellKnownFolderName.RecoverableItemsDeletions or WellKnownFolderName.RecoverableItemsPurges;
        }
        var idOnly = new PropertySet(BasePropertySet.IdOnly);
        var deletedItems = await Folder.Bind(ews.Service, ews.WellKnown(WellKnownFolderName.DeletedItems), idOnly);
        return deletedItems.Id.UniqueId == (await Folder.Bind(ews.Service, target, idOnly)).Id.UniqueId;
    }

    [McpServerTool(Name = "mark_mail", Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Marks e-mails as read/unread and/or sets the follow-up flag.")]
    public async Task<string> MarkMail(
        [Description("Item ids.")] string[] ids,
        [Description("true = read, false = unread. Omit to leave unchanged.")] bool? isRead = null,
        [Description("Follow-up flag: 'flagged', 'complete' or 'none'. Omit to leave unchanged.")] string? flag = null,
        [Description(MailTools.ItemMailboxDescription)] string? mailbox = null)
    {
        await ews.UseMailboxAsync(mailbox);
        if (isRead is null && flag is null) throw new McpException("Nothing to change: specify isRead and/or flag.");
        foreach (var id in ids)
        {
            var message = await EmailMessage.Bind(ews.Service, new ItemId(id), new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.IsRead, ItemSchema.Flag));
            if (isRead is not null) message.IsRead = isRead.Value;
            if (flag is not null)
            {
                message.Flag = new Flag
                {
                    FlagStatus = flag.ToLowerInvariant() switch
                    {
                        "flagged" => ItemFlagStatus.Flagged,
                        "complete" or "completed" => ItemFlagStatus.Complete,
                        _ => ItemFlagStatus.NotFlagged,
                    },
                };
            }
            await message.Update(ConflictResolutionMode.AutoResolve);
        }
        return $"Updated {ids.Length} item(s).";
    }
}
