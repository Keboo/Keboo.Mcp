using Keboo.Mcp.Teams;

namespace Keboo.Mcp.Services;

internal interface ITeamsChatService
{
    Task<TeamsChatListResult> ListChatsAsync(int maxResults, CancellationToken cancellationToken);

    Task<TeamsSendMessageResult> SendMessageToChatAsync(string chatId, string message, CancellationToken cancellationToken);

    Task<TeamsSendMessageResult> SendDirectMessageAsync(string recipientUserPrincipalNameOrId, string message, CancellationToken cancellationToken);
}
