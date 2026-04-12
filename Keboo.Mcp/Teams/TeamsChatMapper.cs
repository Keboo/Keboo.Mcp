using Microsoft.Graph.Models;

namespace Keboo.Mcp.Teams;

internal static class TeamsChatMapper
{
    public static TeamsChatSummary ToSummary(Chat chat, string currentUserId)
    {
        TeamsChatMemberSummary[] members = chat.Members?
            .Select(ToMemberSummary)
            .ToArray() ?? [];

        return new TeamsChatSummary(
            chat.Id ?? throw new InvalidOperationException("Microsoft Graph returned a chat without an ID."),
            chat.ChatType?.ToString() ?? "Unknown",
            ResolveTitle(chat, members, currentUserId),
            chat.Topic,
            chat.WebUrl,
            chat.LastUpdatedDateTime,
            ToPreview(chat.LastMessagePreview),
            members);
    }

    public static TeamsSendMessageResult ToSendMessageResult(
        string chatId,
        ChatMessage message,
        string recipient,
        string deliveryMode)
    {
        string messageId = message.Id ?? throw new InvalidOperationException("Microsoft Graph returned a sent message without an ID.");

        return new TeamsSendMessageResult(
            chatId,
            messageId,
            message.CreatedDateTime,
            recipient,
            deliveryMode);
    }

    private static string ResolveTitle(
        Chat chat,
        IReadOnlyList<TeamsChatMemberSummary> members,
        string currentUserId)
    {
        if (!string.IsNullOrWhiteSpace(chat.Topic))
        {
            return chat.Topic;
        }

        if (chat.ChatType is ChatType.OneOnOne)
        {
            TeamsChatMemberSummary? otherParticipant = members.FirstOrDefault(member =>
                !string.IsNullOrWhiteSpace(member.UserId) &&
                !string.Equals(member.UserId, currentUserId, StringComparison.OrdinalIgnoreCase));

            otherParticipant ??= members.FirstOrDefault();

            return otherParticipant?.DisplayName
                ?? otherParticipant?.Email
                ?? "One-on-one chat";
        }

        return chat.ChatType is ChatType.Group ? "Group chat" : "Chat";
    }

    private static TeamsMessagePreview? ToPreview(ChatMessageInfo? preview)
    {
        if (preview?.Body is null)
        {
            return null;
        }

        return new TeamsMessagePreview(preview.Body.Content, preview.Body.ContentType?.ToString());
    }

    private static TeamsChatMemberSummary ToMemberSummary(ConversationMember member)
    {
        if (member is AadUserConversationMember aadMember)
        {
            return new TeamsChatMemberSummary(
                aadMember.UserId,
                aadMember.DisplayName,
                aadMember.Email,
                aadMember.Roles?.ToArray() ?? []);
        }

        return new TeamsChatMemberSummary(
            member.Id,
            member.DisplayName,
            null,
            member.Roles?.ToArray() ?? []);
    }
}
