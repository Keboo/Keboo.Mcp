using Keboo.Mcp.Teams;
using Microsoft.Graph.Models;

namespace Keboo.Mcp.Services;

internal interface ITeamsGraphClient
{
    Task<CurrentUserProfile> GetCurrentUserAsync(CancellationToken cancellationToken);

    Task<IReadOnlyList<Chat>> ListChatsAsync(int maxResults, CancellationToken cancellationToken);

    Task<Chat> CreateOrGetOneOnOneChatAsync(string currentUserId, string recipientUserPrincipalNameOrId, CancellationToken cancellationToken);

    Task<ChatMessage> SendChatMessageAsync(string chatId, string message, CancellationToken cancellationToken);
}
