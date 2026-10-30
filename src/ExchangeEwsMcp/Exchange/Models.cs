namespace ExchangeEwsMcp.Exchange;

// Result types returned by the tools. Null values are omitted from the JSON sent to the model.

public sealed record PersonInfo(string? Name, string? Email)
{
    public override string ToString() => Name is null ? Email ?? "" : $"{Name} <{Email}>";
}

// Folders are addressed by path in the other tools; ids would only cost context.
public sealed record FolderInfo(string Path, int TotalCount, int? UnreadCount, int ChildFolderCount);

public sealed record MailSummary(
    string Id,
    string? Subject,
    PersonInfo? From,
    string? To,
    string? Received,
    bool IsRead,
    bool? HasAttachments,
    string? Importance,
    bool? IsDraft,
    string? Preview);

public sealed record AttachmentInfo(int Index, string Id, string Name, string? ContentType, int Size, bool IsInline, bool IsItemAttachment);

public sealed record MailDetails(
    string Id,
    string? Subject,
    PersonInfo? From,
    IReadOnlyList<PersonInfo> To,
    IReadOnlyList<PersonInfo> Cc,
    IReadOnlyList<PersonInfo>? Bcc,
    string? Sent,
    string? Received,
    bool IsRead,
    bool IsDraft,
    string? Importance,
    IReadOnlyList<string>? Categories,
    string? ConversationId,
    string BodyType,
    string Body,
    bool BodyTruncated,
    bool? QuotedHistoryOmitted,
    IReadOnlyList<AttachmentInfo> Attachments,
    string? WebLink);

public sealed record SearchResult<T>(int TotalCount, bool MoreAvailable, IReadOnlyList<T> Items);

public sealed record DraftResult(string Id, string? Subject, string Folder, string Message);

public sealed record EventSummary(
    string Id,
    string? Subject,
    string? Start,
    string? End,
    bool IsAllDay,
    string? Location,
    PersonInfo? Organizer,
    bool IsMeeting,
    bool IsRecurring,
    bool IsCancelled,
    string? ShowAs,
    string? MyResponse);

public sealed record AttendeeStatus(string? Name, string? Email, string? Response);

public sealed record EventDetails(
    string Id,
    string? Subject,
    string? Start,
    string? End,
    bool IsAllDay,
    string? Location,
    PersonInfo? Organizer,
    IReadOnlyList<AttendeeStatus> RequiredAttendees,
    IReadOnlyList<AttendeeStatus> OptionalAttendees,
    bool IsMeeting,
    bool IsRecurring,
    bool IsCancelled,
    string? ShowAs,
    string? MyResponse,
    bool? IsOrganizer,
    IReadOnlyList<string>? Categories,
    string Body,
    bool BodyTruncated);

public sealed record BusySlot(string Start, string End, string Status, string? Subject, string? Location);

public sealed record PersonAvailability(string Email, string? Error, IReadOnlyList<BusySlot> Busy);

public sealed record SuggestedTime(string Start, string Quality, IReadOnlyList<string>? Conflicts);

public sealed record AvailabilityResult(
    string TimeZone,
    IReadOnlyList<PersonAvailability> Attendees,
    IReadOnlyList<SuggestedTime>? Suggestions);

public sealed record ResolvedName(string? Name, string? Email, string? Type, string? Department, string? JobTitle, string? Office, string? Phone, string Source);

public sealed record ListMember(string? Name, string? Email, string Type, string? ViaGroup);

public sealed record GroupExpansion(string Group, int Count, bool Complete, IReadOnlyList<ListMember> Members);

public sealed record OutOfOfficeInfo(string Mailbox, string State, string? Start, string? End, string ExternalAudience, string? InternalReply, string? ExternalReply);

public sealed record UserInfo(string? DisplayName, string? Email, string? UserPrincipalName, string TimeZone, string Now, bool SendingEnabled, string AccessLevel);
