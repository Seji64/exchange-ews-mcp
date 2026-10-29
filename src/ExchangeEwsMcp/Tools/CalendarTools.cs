using System.ComponentModel;
using ExchangeEwsMcp.Configuration;
using ExchangeEwsMcp.Exchange;
using Microsoft.Exchange.WebServices.Data;
using ModelContextProtocol;
using ModelContextProtocol.Server;
using static ExchangeEwsMcp.Exchange.EwsHelpers;

namespace ExchangeEwsMcp.Tools;

[McpServerToolType]
public sealed class CalendarTools(EwsContext ews)
{
    private static readonly PropertySet EventSummaryProperties = new(BasePropertySet.IdOnly,
        ItemSchema.Subject, AppointmentSchema.Start, AppointmentSchema.End, AppointmentSchema.IsAllDayEvent,
        AppointmentSchema.Location, AppointmentSchema.Organizer, AppointmentSchema.IsMeeting, AppointmentSchema.IsRecurring,
        AppointmentSchema.IsCancelled, AppointmentSchema.LegacyFreeBusyStatus, AppointmentSchema.MyResponseType);

    [McpServerTool(Name = "list_events", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Lists calendar events (recurring events expanded) in a time range. Default: today and the next 7 days.")]
    public async Task<IReadOnlyList<EventSummary>> ListEvents(
        [Description("Range start (ISO 8601). Default: today 00:00.")] string? start = null,
        [Description("Range end (ISO 8601). Default: start + 7 days.")] string? end = null,
        [Description("Maximum number of events (default 50).")] int maxResults = 50,
        [Description(MailTools.MailboxDescription)] string? mailbox = null)
    {
        var from = start is null ? ews.Now.Date : ews.ParseDateTime(start, nameof(start));
        var to = end is null ? from.AddDays(7) : ews.ParseDateTime(end, nameof(end));
        if (to <= from) throw new McpException("'end' must be after 'start'.");

        var view = new CalendarView(from, to, Math.Clamp(maxResults, 1, Math.Max(ews.Options.MaxResults, 100))) { PropertySet = EventSummaryProperties };
        var results = await ews.Service.FindAppointments(ews.WellKnown(WellKnownFolderName.Calendar, mailbox), view);
        return results.Items.Select(a => new EventSummary(
            a.Id.UniqueId!,
            Try(() => a.Subject),
            ews.TryFormat(() => a.Start),
            ews.TryFormat(() => a.End),
            Try(() => a.IsAllDayEvent),
            NullIfEmpty(Try(() => a.Location)),
            ToPerson(Try(() => a.Organizer)),
            Try(() => a.IsMeeting),
            Try(() => a.IsRecurring),
            Try(() => a.IsCancelled),
            Try(() => a.LegacyFreeBusyStatus).ToString(),
            Try(() => a.IsMeeting) ? Try(() => a.MyResponseType).ToString() : null)).ToList();
    }

    [McpServerTool(Name = "get_event", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("Gets details of a calendar event including attendees, their responses and the description.")]
    public async Task<EventDetails> GetEvent([Description("Event id from list_events.")] string id)
    {
        var props = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Body) { RequestedBodyType = BodyType.Text };
        var a = await Appointment.Bind(ews.Service, new ItemId(id), props);
        return ToDetails(a);
    }

    private EventDetails ToDetails(Appointment a)
    {
        static IReadOnlyList<AttendeeStatus> Attendees(AttendeeCollection? c) =>
            c?.Select(x => new AttendeeStatus(NullIfEmpty(x.Name), NullIfEmpty(x.Address), x.ResponseType?.ToString())).ToList() ?? [];

        var (body, truncated) = Truncate(Try(() => a.Body?.Text), ews.Options.MaxBodyLength);
        return new EventDetails(
            a.Id.UniqueId!,
            Try(() => a.Subject),
            ews.TryFormat(() => a.Start),
            ews.TryFormat(() => a.End),
            Try(() => a.IsAllDayEvent),
            NullIfEmpty(Try(() => a.Location)),
            ToPerson(Try(() => a.Organizer)),
            Attendees(Try(() => a.RequiredAttendees)),
            Attendees(Try(() => a.OptionalAttendees)),
            Try(() => a.IsMeeting),
            Try(() => a.IsRecurring),
            Try(() => a.IsCancelled),
            Try(() => a.LegacyFreeBusyStatus).ToString(),
            Try(() => a.IsMeeting) ? Try(() => a.MyResponseType).ToString() : null,
            Try(() => a.MyResponseType) == MeetingResponseType.Organizer ? true : Try(() => a.IsMeeting) ? false : null,
            Try(() => a.Categories) is { Count: > 0 } categories ? categories.ToList() : null,
            body,
            truncated);
    }

    [McpServerTool(Name = "get_availability", ReadOnly = true, Idempotent = true, OpenWorld = false)]
    [Description("""
        Gets free/busy information of people (and rooms) and optionally suggests meeting times when everyone is available.
        The signed-in user is included automatically.
        """)]
    public async Task<AvailabilityResult> GetAvailability(
        [Description("Attendees (e-mail addresses or names). Rooms can be added by address.")] string[] attendees,
        [Description("Range start (ISO 8601).")] string start,
        [Description("Range end (ISO 8601). Max. 42 days after start.")] string end,
        [Description("Meeting duration in minutes, used for suggestions. Default 30.")] int durationMinutes = 30,
        [Description("Return suggested meeting times. Default true.")] bool suggestTimes = true)
    {
        var from = ews.ParseDateTime(start, nameof(start));
        var to = ews.ParseDateTime(end, nameof(end));
        if (to <= from) throw new McpException("'end' must be after 'start'.");

        var addresses = (await ResolveRecipientsAsync(ews.Service, attendees)).Select(a => a.Address).ToList();
        var me = await ews.GetCurrentUserSmtpAsync();
        if (!addresses.Contains(me, StringComparer.OrdinalIgnoreCase)) addresses.Insert(0, me);

        // Availability windows must start/end at midnight.
        var windowStart = from.Date;
        var windowEnd = to.TimeOfDay == TimeSpan.Zero ? to.Date : to.Date.AddDays(1);
        var options = new AvailabilityOptions
        {
            MeetingDuration = Math.Clamp(durationMinutes, 5, 1440),
            RequestedFreeBusyView = FreeBusyViewType.Detailed,
            MaximumNonWorkHoursSuggestionsPerDay = 0,
            MaximumSuggestionsPerDay = 10,
            DetailedSuggestionsWindow = new TimeWindow(windowStart, windowEnd),
        };
        var results = await ews.Service.GetUserAvailability(
            addresses.Select(a => new Microsoft.Exchange.WebServices.Data.AttendeeInfo(a)),
            new TimeWindow(windowStart, windowEnd),
            suggestTimes ? AvailabilityData.FreeBusyAndSuggestions : AvailabilityData.FreeBusy,
            options);

        var people = results.AttendeesAvailability.Select((r, i) => new PersonAvailability(
            addresses[i],
            r.ErrorCode == ServiceError.NoError ? null : r.ErrorMessage,
            r.CalendarEvents
                .Where(e => e.EndTime > from && e.StartTime < to && e.FreeBusyStatus != LegacyFreeBusyStatus.Free)
                .Select(e => new BusySlot(ews.Format(e.StartTime), ews.Format(e.EndTime), e.FreeBusyStatus.ToString(),
                    NullIfEmpty(e.Details?.Subject), NullIfEmpty(e.Details?.Location)))
                .ToList())).ToList();

        List<SuggestedTime>? suggestions = null;
        if (suggestTimes && results.Suggestions is not null)
        {
            suggestions = results.Suggestions
                .SelectMany(day => day.TimeSuggestions)
                .Where(s => s.MeetingTime >= from && s.MeetingTime.AddMinutes(options.MeetingDuration) <= to)
                .OrderByDescending(s => s.Quality == SuggestionQuality.Excellent)
                .ThenBy(s => s.MeetingTime)
                .Take(15)
                .Select(s => new SuggestedTime(
                    ews.Format(s.MeetingTime),
                    s.Quality.ToString(),
                    s.Conflicts.Count > 0 ? s.Conflicts.Select(c => $"{c.ConflictType} ({c.FreeBusyStatus})").ToList() : null))
                .ToList();
        }

        return new AvailabilityResult(ews.TimeZone.Id, people, suggestions);
    }

    [McpServerTool(Name = "create_event", Destructive = false, OpenWorld = false)]
    [Description("""
        Creates a calendar event. With attendees it becomes a meeting that is SAVED WITHOUT SENDING invitations
        (unless sendInvitations=true and sending is enabled on the server). The user can send it from Outlook.
        """)]
    public async Task<EventDetails> CreateEvent(
        [Description("Subject.")] string subject,
        [Description("Start (ISO 8601). For all-day events the date.")] string start,
        [Description("End (ISO 8601). For all-day events the day after the last day.")] string end,
        [Description("Required attendees (addresses or names).")] string[]? requiredAttendees = null,
        [Description("Optional attendees (addresses or names).")] string[]? optionalAttendees = null,
        [Description("Location text.")] string? location = null,
        [Description("Description / agenda (plain text).")] string? body = null,
        [Description("All-day event. Default false.")] bool isAllDay = false,
        [Description("Show as: 'busy' (default), 'free', 'tentative', 'oof', 'workingelsewhere'.")] string showAs = "busy",
        [Description("Reminder minutes before start. Default 15; 0 disables the reminder.")] int reminderMinutes = 15,
        [Description("Send invitations to attendees immediately (requires sending to be enabled on the server).")] bool sendInvitations = false,
        [Description(MailTools.MailboxDescription)] string? mailbox = null)
    {
        if (sendInvitations) ews.EnsureSendAllowed("meeting invitations");

        var appointment = new Appointment(ews.Service)
        {
            Subject = subject,
            Start = ews.ParseDateTime(start, nameof(start)),
            End = ews.ParseDateTime(end, nameof(end)),
            IsAllDayEvent = isAllDay,
            LegacyFreeBusyStatus = ParseShowAs(showAs),
            IsReminderSet = reminderMinutes > 0,
        };
        if (appointment.End <= appointment.Start) throw new McpException("'end' must be after 'start'.");
        if (reminderMinutes > 0) appointment.ReminderMinutesBeforeStart = reminderMinutes;
        if (location is not null) appointment.Location = location;
        if (body is not null) appointment.Body = CreateBody(body, "text");
        foreach (var a in await ResolveRecipientsAsync(ews.Service, requiredAttendees)) appointment.RequiredAttendees.Add(new Attendee(a));
        foreach (var a in await ResolveRecipientsAsync(ews.Service, optionalAttendees)) appointment.OptionalAttendees.Add(new Attendee(a));

        await appointment.Save(ews.WellKnown(WellKnownFolderName.Calendar, mailbox),
            sendInvitations ? SendInvitationsMode.SendToAllAndSaveCopy : SendInvitationsMode.SendToNone);
        return await GetEvent(appointment.Id.UniqueId!);
    }

    [McpServerTool(Name = "update_event", Destructive = false, OpenWorld = false)]
    [Description("""
        Updates a calendar event. Only given fields change; attendee lists replace the existing ones.
        Updates are NOT sent to attendees unless sendUpdates=true and sending is enabled on the server.
        """)]
    public async Task<EventDetails> UpdateEvent(
        [Description("Event id.")] string id,
        [Description("New subject.")] string? subject = null,
        [Description("New start (ISO 8601).")] string? start = null,
        [Description("New end (ISO 8601).")] string? end = null,
        [Description("New required attendees.")] string[]? requiredAttendees = null,
        [Description("New optional attendees.")] string[]? optionalAttendees = null,
        [Description("New location.")] string? location = null,
        [Description("New description (plain text).")] string? body = null,
        [Description("Show as: 'busy', 'free', 'tentative', 'oof', 'workingelsewhere'.")] string? showAs = null,
        [Description("Send the update to attendees (requires sending to be enabled on the server).")] bool sendUpdates = false)
    {
        if (sendUpdates) ews.EnsureSendAllowed("meeting updates");

        var appointment = await Appointment.Bind(ews.Service, new ItemId(id));
        if (subject is not null) appointment.Subject = subject;
        if (start is not null) appointment.Start = ews.ParseDateTime(start, nameof(start));
        if (end is not null) appointment.End = ews.ParseDateTime(end, nameof(end));
        if (location is not null) appointment.Location = location;
        if (body is not null) appointment.Body = CreateBody(body, "text");
        if (showAs is not null) appointment.LegacyFreeBusyStatus = ParseShowAs(showAs);
        if (requiredAttendees is not null)
        {
            appointment.RequiredAttendees.Clear();
            foreach (var a in await ResolveRecipientsAsync(ews.Service, requiredAttendees)) appointment.RequiredAttendees.Add(new Attendee(a));
        }
        if (optionalAttendees is not null)
        {
            appointment.OptionalAttendees.Clear();
            foreach (var a in await ResolveRecipientsAsync(ews.Service, optionalAttendees)) appointment.OptionalAttendees.Add(new Attendee(a));
        }

        await appointment.Update(ConflictResolutionMode.AlwaysOverwrite,
            sendUpdates ? SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy : SendInvitationsOrCancellationsMode.SendToNone);
        return await GetEvent(appointment.Id.UniqueId!);
    }

    [McpServerTool(Name = "delete_event", Destructive = true, OpenWorld = false)]
    [Description("""
        Deletes a calendar event (moved to Deleted Items). Cancelling a meeting you organize with attendees sends
        cancellations and therefore requires sending to be enabled on the server.
        """)]
    public async Task<string> DeleteEvent(
        [Description("Event id.")] string id,
        [Description("Optional cancellation message for attendees.")] string? cancellationMessage = null)
    {
        var appointment = await Appointment.Bind(ews.Service, new ItemId(id), new PropertySet(BasePropertySet.IdOnly,
            AppointmentSchema.IsMeeting, AppointmentSchema.MyResponseType, AppointmentSchema.RequiredAttendees, AppointmentSchema.OptionalAttendees));
        var isOrganizedMeeting = appointment.IsMeeting
            && appointment.MyResponseType == MeetingResponseType.Organizer
            && appointment.RequiredAttendees.Count + appointment.OptionalAttendees.Count > 0;

        if (isOrganizedMeeting)
        {
            ews.EnsureSendAllowed("meeting cancellation");
            await appointment.CancelMeeting(cancellationMessage ?? "");
            return "Meeting cancelled and cancellations sent to attendees.";
        }
        if (appointment.IsMeeting)
        {
            throw new McpException("This is a meeting organized by someone else. Use respond_to_meeting with 'decline' instead.");
        }
        await appointment.Delete(DeleteMode.MoveToDeletedItems, SendCancellationsMode.SendToNone);
        return "Event moved to Deleted Items.";
    }

    [McpServerTool(Name = "respond_to_meeting", Destructive = false, OpenWorld = false)]
    [Description("""
        Accepts, tentatively accepts or declines a meeting invitation. By default the response is recorded without
        notifying the organizer; sendResponse=true requires sending to be enabled on the server.
        """)]
    public async Task<string> RespondToMeeting(
        [Description("Id of the calendar event or meeting request.")] string id,
        [Description("'accept', 'tentative' or 'decline'.")] string response,
        [Description("Notify the organizer. Default false.")] bool sendResponse = false,
        [Description("Optional message to the organizer (only with sendResponse=true).")] string? message = null)
    {
        if (sendResponse) ews.EnsureSendAllowed("meeting response");

        var kind = response.ToLowerInvariant();
        if (kind is not ("accept" or "tentative" or "decline"))
        {
            throw new McpException("response must be 'accept', 'tentative' or 'decline'.");
        }
        if (kind == "decline" && ews.Options.AccessLevel != AccessLevel.Full)
        {
            throw new McpException("Declining moves the meeting to Deleted Items, and deleting is disabled on this server.");
        }

        var item = await Item.Bind(ews.Service, new ItemId(id));
        var appointment = item switch
        {
            Appointment a => a,
            MeetingRequest request => await Appointment.Bind(ews.Service, request.AssociatedAppointmentId ?? throw new McpException("Meeting request has no calendar item.")),
            _ => throw new McpException("The item is neither a calendar event nor a meeting request."),
        };

        if (sendResponse && !string.IsNullOrWhiteSpace(message))
        {
            if (kind == "decline")
            {
                var decline = appointment.CreateDeclineMessage();
                decline.Body = CreateBody(message, "text");
                await decline.SendAndSaveCopy();
            }
            else
            {
                var accept = appointment.CreateAcceptMessage(tentative: kind == "tentative");
                accept.Body = CreateBody(message, "text");
                await accept.SendAndSaveCopy();
            }
        }
        else
        {
            _ = kind switch
            {
                "accept" => await appointment.Accept(sendResponse),
                "tentative" => await appointment.AcceptTentatively(sendResponse),
                _ => await appointment.Decline(sendResponse),
            };
        }
        return $"Meeting {kind} recorded{(sendResponse ? " and the organizer was notified" : " without notifying the organizer")}.";
    }

    private static LegacyFreeBusyStatus ParseShowAs(string? value) => value?.Replace(" ", "").ToLowerInvariant() switch
    {
        "free" => LegacyFreeBusyStatus.Free,
        "tentative" => LegacyFreeBusyStatus.Tentative,
        "oof" or "outofoffice" or "away" => LegacyFreeBusyStatus.OOF,
        "workingelsewhere" => LegacyFreeBusyStatus.WorkingElsewhere,
        _ => LegacyFreeBusyStatus.Busy,
    };
}
