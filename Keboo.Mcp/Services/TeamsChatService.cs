using Keboo.Mcp.Configuration;
using Keboo.Mcp.Teams;
using Microsoft.Extensions.Options;
using Microsoft.Graph.Models;

namespace Keboo.Mcp.Services;

internal sealed class TeamsChatService : ITeamsChatService
{
    private readonly ITeamsGraphClient _graphClient;
    private readonly TeamsGraphOptions _options;

    public TeamsChatService(
        ITeamsGraphClient graphClient,
        IOptions<TeamsGraphOptions> options)
    {
        _graphClient = graphClient;
        _options = options.Value;
    }

    public async Task<TeamsChatListResult> ListChatsAsync(int maxResults, CancellationToken cancellationToken)
    {
        int normalizedMaxResults = NormalizeMaxResults(maxResults);
        CurrentUserProfile currentUser = await _graphClient.GetCurrentUserAsync(cancellationToken);
        IReadOnlyList<Chat> chats = await _graphClient.ListChatsAsync(normalizedMaxResults, cancellationToken);

        return new TeamsChatListResult(chats.Select(chat => TeamsChatMapper.ToSummary(chat, currentUser.Id)).ToArray());
    }

    public async Task<TeamsSendMessageResult> SendMessageToChatAsync(string chatId, string message, CancellationToken cancellationToken)
    {
        string normalizedChatId = NormalizeRequiredValue(chatId, nameof(chatId));
        string normalizedMessage = NormalizeRequiredValue(message, nameof(message));
        ChatMessage sentMessage = await _graphClient.SendChatMessageAsync(normalizedChatId, normalizedMessage, cancellationToken);

        return TeamsChatMapper.ToSendMessageResult(normalizedChatId, sentMessage, normalizedChatId, "existingChat");
    }

    public async Task<TeamsSendMessageResult> SendDirectMessageAsync(
        string recipientUserPrincipalNameOrId,
        string message,
        CancellationToken cancellationToken)
    {
        string normalizedRecipient = NormalizeRequiredValue(recipientUserPrincipalNameOrId, nameof(recipientUserPrincipalNameOrId));
        string normalizedMessage = NormalizeRequiredValue(message, nameof(message));

        CurrentUserProfile currentUser = await _graphClient.GetCurrentUserAsync(cancellationToken);
        Chat chat = await _graphClient.CreateOrGetOneOnOneChatAsync(currentUser.Id, normalizedRecipient, cancellationToken);
        string chatId = chat.Id ?? throw new InvalidOperationException("Microsoft Graph returned a one-on-one chat without an ID.");
        ChatMessage sentMessage = await _graphClient.SendChatMessageAsync(chatId, normalizedMessage, cancellationToken);

        return TeamsChatMapper.ToSendMessageResult(chatId, sentMessage, normalizedRecipient, "directMessage");
    }

    private int NormalizeMaxResults(int maxResults)
    {
        int effectiveValue = maxResults <= 0 ? _options.DefaultChatPageSize : maxResults;
        return Math.Clamp(effectiveValue, 1, Math.Max(_options.MaxChatPageSize, 1));
    }

    private static string NormalizeRequiredValue(string value, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("Value cannot be empty.", parameterName);
        }

        return value.Trim();
    }
}
