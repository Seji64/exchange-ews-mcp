using System.ComponentModel;
using ExchangeEwsMcp.Exchange;
using Microsoft.Exchange.WebServices.Data;
using ModelContextProtocol;
using ModelContextProtocol.Server;

namespace ExchangeEwsMcp.Tools;

/// <summary>Tools that send mail. Only registered when Exchange:AllowSend is true.</summary>
[McpServerToolType]
public sealed class SendTools(EwsContext ews)
{
    [McpServerTool(Name = "send_draft", Destructive = true, OpenWorld = true)]
    [Description("""
        Sends an existing draft. ONLY call this after the user has explicitly confirmed sending this specific draft
        (show recipients and subject first). A copy is saved in Sent Items.
        """)]
    public async Task<string> SendDraft([Description("Draft id.")] string id)
    {
        ews.EnsureSendAllowed("send_draft");
        var draft = await EmailMessage.Bind(ews.Service, new ItemId(id), new PropertySet(BasePropertySet.IdOnly,
            ItemSchema.IsDraft, ItemSchema.Subject, EmailMessageSchema.ToRecipients, EmailMessageSchema.CcRecipients,
            EmailMessageSchema.BccRecipients, EmailMessageSchema.From));
        if (!draft.IsDraft) throw new McpException("This item is not a draft.");
        if (draft.ToRecipients.Count + draft.CcRecipients.Count + draft.BccRecipients.Count == 0)
        {
            throw new McpException("The draft has no recipients.");
        }
        await draft.SendAndSaveCopy();
        var recipients = string.Join(", ", draft.ToRecipients.Concat(draft.CcRecipients).Select(r => r.Address));
        return $"Sent '{draft.Subject}' to {recipients}.";
    }

    [McpServerTool(Name = "send_meeting_invitations", Destructive = true, OpenWorld = true)]
    [Description("""
        Sends the invitations (or updates) of a meeting that was saved without sending. ONLY call this after the user
        explicitly confirmed.
        """)]
    public async Task<string> SendMeetingInvitations([Description("Event id.")] string id)
    {
        ews.EnsureSendAllowed("send_meeting_invitations");
        var appointment = await Appointment.Bind(ews.Service, new ItemId(id), new PropertySet(BasePropertySet.IdOnly,
            ItemSchema.Subject, AppointmentSchema.RequiredAttendees, AppointmentSchema.OptionalAttendees, AppointmentSchema.MyResponseType));
        if (appointment.RequiredAttendees.Count + appointment.OptionalAttendees.Count == 0)
        {
            throw new McpException("The event has no attendees.");
        }
        if (appointment.MyResponseType is not (MeetingResponseType.Organizer or MeetingResponseType.Unknown))
        {
            throw new McpException("Only the organizer can send invitations for this meeting.");
        }
        // Touch the subject so EWS performs an update that sends the invitations.
        appointment.Subject = appointment.Subject;
        await appointment.Update(ConflictResolutionMode.AlwaysOverwrite, SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy);
        return $"Invitations for '{appointment.Subject}' sent.";
    }
}
