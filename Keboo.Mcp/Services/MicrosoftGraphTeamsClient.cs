using Azure.Core;
using Azure.Identity;
using Keboo.Mcp.Configuration;
using Keboo.Mcp.Teams;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Graph.Models;

namespace Keboo.Mcp.Services;

internal sealed class MicrosoftGraphTeamsClient : ITeamsGraphClient
{
    private static readonly string[] GraphScopes = ["User.Read", "Chat.Read", "Chat.Create", "ChatMessage.Send"];

    private readonly ILogger<MicrosoftGraphTeamsClient> _logger;
    private readonly TeamsGraphOptions _options;
    private readonly Lazy<GraphServiceClient> _graphClient;
    private readonly SemaphoreSlim _currentUserLock = new(1, 1);
    private CurrentUserProfile? _currentUser;

    public MicrosoftGraphTeamsClient(
        ILogger<MicrosoftGraphTeamsClient> logger,
        IOptions<TeamsGraphOptions> options)
    {
        _logger = logger;
        _options = options.Value;
        _graphClient = new Lazy<GraphServiceClient>(CreateGraphServiceClient, isThreadSafe: true);
    }

    public async Task<CurrentUserProfile> GetCurrentUserAsync(CancellationToken cancellationToken)
    {
        if (_currentUser is not null)
        {
            return _currentUser;
        }

        await _currentUserLock.WaitAsync(cancellationToken);

        try
        {
            if (_currentUser is not null)
            {
                return _currentUser;
            }

            User me = await _graphClient.Value.Me.GetAsync(cancellationToken: cancellationToken)
                ?? throw new InvalidOperationException("Microsoft Graph did not return a user profile for the signed-in account.");

            string userId = me.Id ?? throw new InvalidOperationException("Microsoft Graph returned a user profile without an ID.");
            _currentUser = new CurrentUserProfile(userId, me.DisplayName, me.UserPrincipalName ?? me.Mail);

            return _currentUser;
        }
        finally
        {
            _currentUserLock.Release();
        }
    }

    public async Task<IReadOnlyList<Chat>> ListChatsAsync(int maxResults, CancellationToken cancellationToken)
    {
        int requestedResults = Math.Max(maxResults, 1);
        int pageSize = Math.Min(requestedResults, Math.Max(_options.MaxChatPageSize, 1));
        List<Chat> chats = [];

        ChatCollectionResponse? page = await _graphClient.Value.Me.Chats.GetAsync(
            requestConfiguration =>
            {
                requestConfiguration.QueryParameters.Top = pageSize;
                requestConfiguration.QueryParameters.Expand = ["members", "lastMessagePreview"];
            },
            cancellationToken);

        while (page is not null)
        {
            if (page.Value is not null)
            {
                foreach (Chat chat in page.Value)
                {
                    chats.Add(chat);

                    if (chats.Count >= requestedResults)
                    {
                        return chats;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(page.OdataNextLink))
            {
                break;
            }

            page = await _graphClient.Value.Me.Chats
                .WithUrl(page.OdataNextLink)
                .GetAsync(cancellationToken: cancellationToken);
        }

        return chats;
    }

    public async Task<Chat> CreateOrGetOneOnOneChatAsync(
        string currentUserId,
        string recipientUserPrincipalNameOrId,
        CancellationToken cancellationToken)
    {
        Chat request = new()
        {
            ChatType = ChatType.OneOnOne,
            Members =
            [
                CreateConversationMember(currentUserId),
                CreateConversationMember(recipientUserPrincipalNameOrId)
            ]
        };

        return await _graphClient.Value.Chats.PostAsync(request, cancellationToken: cancellationToken)
            ?? throw new InvalidOperationException("Microsoft Graph did not return a chat after the one-on-one chat request.");
    }

    public async Task<ChatMessage> SendChatMessageAsync(string chatId, string message, CancellationToken cancellationToken)
    {
        ChatMessage request = new()
        {
            Body = new ItemBody
            {
                Content = message,
                ContentType = BodyType.Text
            }
        };

        return await _graphClient.Value.Chats[chatId].Messages.PostAsync(request, cancellationToken: cancellationToken)
            ?? throw new InvalidOperationException("Microsoft Graph did not return a sent message payload.");
    }

    private GraphServiceClient CreateGraphServiceClient()
    {
        EnsureGraphConfiguration();

        TokenCredential credential = _options.AuthenticationMode switch
        {
            TeamsGraphAuthenticationMode.DeviceCode => CreateDeviceCodeCredential(),
            TeamsGraphAuthenticationMode.InteractiveBrowser => CreateInteractiveBrowserCredential(),
            _ => throw new InvalidOperationException($"Unsupported Teams Graph authentication mode '{_options.AuthenticationMode}'.")
        };

        return new GraphServiceClient(credential, GraphScopes);
    }

    private InteractiveBrowserCredential CreateInteractiveBrowserCredential()
    {
        return new InteractiveBrowserCredential(new InteractiveBrowserCredentialOptions
        {
            TenantId = _options.TenantId,
            ClientId = _options.ClientId,
            TokenCachePersistenceOptions = new TokenCachePersistenceOptions
            {
                Name = NormalizeTokenCacheName()
            }
        });
    }

    private DeviceCodeCredential CreateDeviceCodeCredential()
    {
        return new DeviceCodeCredential(new DeviceCodeCredentialOptions
        {
            TenantId = _options.TenantId,
            ClientId = _options.ClientId,
            TokenCachePersistenceOptions = new TokenCachePersistenceOptions
            {
                Name = NormalizeTokenCacheName()
            },
            DeviceCodeCallback = (deviceCodeInfo, _) =>
            {
                _logger.LogWarning("{DeviceCodeMessage}", deviceCodeInfo.Message);
                return Task.CompletedTask;
            }
        });
    }

    private void EnsureGraphConfiguration()
    {
        List<string> missingValues = [];

        if (string.IsNullOrWhiteSpace(_options.TenantId))
        {
            missingValues.Add("KebooMcp:Graph:TenantId");
        }

        if (string.IsNullOrWhiteSpace(_options.ClientId))
        {
            missingValues.Add("KebooMcp:Graph:ClientId");
        }

        if (missingValues.Count == 0)
        {
            return;
        }

        throw new InvalidOperationException(
            "Microsoft Graph for Teams is not configured. " +
            $"Set {string.Join(" and ", missingValues)} via user secrets or environment variables (for example KebooMcp__Graph__TenantId). " +
            "This server uses delegated Microsoft Graph permissions and requires a work or school account.");
    }

    private string NormalizeTokenCacheName()
    {
        return string.IsNullOrWhiteSpace(_options.TokenCacheName)
            ? "Keboo.Mcp"
            : _options.TokenCacheName.Trim();
    }

    private static AadUserConversationMember CreateConversationMember(string userIdOrUpn)
    {
        string normalizedUserIdOrUpn = userIdOrUpn.Trim();

        return new AadUserConversationMember
        {
            OdataType = "#microsoft.graph.aadUserConversationMember",
            Roles = ["owner"],
            AdditionalData = new Dictionary<string, object>
            {
                ["user@odata.bind"] = $"https://graph.microsoft.com/v1.0/users('{EscapeODataString(normalizedUserIdOrUpn)}')"
            }
        };
    }

    private static string EscapeODataString(string value)
    {
        return value.Replace("'", "''", StringComparison.Ordinal);
    }
}
