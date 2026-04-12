namespace Keboo.Mcp.Teams;

internal sealed record CurrentUserProfile(string Id, string? DisplayName, string? UserPrincipalName);

internal sealed record TeamsChatListResult(IReadOnlyList<TeamsChatSummary> Chats);

internal sealed record TeamsChatSummary(
    string Id,
    string ChatType,
    string Title,
    string? Topic,
    string? WebUrl,
    DateTimeOffset? LastUpdatedDateTime,
    TeamsMessagePreview? LastMessagePreview,
    IReadOnlyList<TeamsChatMemberSummary> Members);

internal sealed record TeamsMessagePreview(string? Content, string? ContentType);

internal sealed record TeamsChatMemberSummary(
    string? UserId,
    string? DisplayName,
    string? Email,
    IReadOnlyList<string> Roles);

internal sealed record TeamsSendMessageResult(
    string ChatId,
    string MessageId,
    DateTimeOffset? CreatedDateTime,
    string? Recipient,
    string DeliveryMode);
