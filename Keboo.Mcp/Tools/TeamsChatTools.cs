using System.ComponentModel;
using Keboo.Mcp.Services;
using Keboo.Mcp.Teams;
using ModelContextProtocol.Server;

namespace Keboo.Mcp.Tools;

[McpServerToolType]
internal sealed class TeamsChatTools
{
    private readonly ITeamsChatService _teamsChatService;

    public TeamsChatTools(ITeamsChatService teamsChatService)
    {
        _teamsChatService = teamsChatService;
    }

    [McpServerTool]
    [Description("Lists the signed-in user's Microsoft Teams chats, including chat IDs, members, and preview metadata.")]
    public Task<TeamsChatListResult> ListChats(
        [Description("Maximum number of chats to return. Defaults to 50 and is capped by server configuration.")] int maxResults = 50)
    {
        return _teamsChatService.ListChatsAsync(maxResults, CancellationToken.None);
    }

    [McpServerTool]
    [Description("Sends a plain-text Microsoft Teams message to an existing chat by chat ID.")]
    public Task<TeamsSendMessageResult> SendChatMessage(
        [Description("The Microsoft Teams chat ID to send to. Use ListChats first to discover existing chat IDs.")] string chatId,
        [Description("The plain-text message to send.")] string message)
    {
        return _teamsChatService.SendMessageToChatAsync(chatId, message, CancellationToken.None);
    }

    [McpServerTool]
    [Description("Sends a plain-text direct Microsoft Teams message to a user identified by work-account UPN/email or Microsoft Entra user ID. The tool reuses an existing one-on-one chat when possible and otherwise creates it.")]
    public Task<TeamsSendMessageResult> SendDirectMessage(
        [Description("The recipient's work-account user principal name (email) or Microsoft Entra user ID.")] string recipientUserPrincipalNameOrId,
        [Description("The plain-text message to send.")] string message)
    {
        return _teamsChatService.SendDirectMessageAsync(recipientUserPrincipalNameOrId, message, CancellationToken.None);
    }
}
