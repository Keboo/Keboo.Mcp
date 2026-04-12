using Keboo.Mcp.Configuration;
using Keboo.Mcp.Services;
using Keboo.Mcp.Teams;
using Microsoft.Extensions.Options;
using Microsoft.Graph.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Keboo.Mcp.Tests;

[TestClass]
public sealed class TeamsChatServiceTests
{
    [TestMethod]
    public async Task ListChatsAsync_UsesConfiguredDefaultWhenRequestedValueIsNotPositive()
    {
        Mock<ITeamsGraphClient> graphClient = new();
        graphClient
            .Setup(client => client.GetCurrentUserAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new CurrentUserProfile("me", "Me", "me@contoso.com"));
        graphClient
            .Setup(client => client.ListChatsAsync(7, It.IsAny<CancellationToken>()))
            .ReturnsAsync(
            [
                new Chat
                {
                    Id = "chat-1",
                    ChatType = ChatType.Group,
                    Topic = "Operations"
                }
            ]);

        TeamsChatService service = CreateService(graphClient.Object, defaultChatPageSize: 7, maxChatPageSize: 25);

        TeamsChatListResult result = await service.ListChatsAsync(0, CancellationToken.None);

        Assert.AreEqual(1, result.Chats.Count);
        Assert.AreEqual("Operations", result.Chats[0].Title);
        graphClient.Verify(client => client.ListChatsAsync(7, It.IsAny<CancellationToken>()), Times.Once);
    }

    [TestMethod]
    public async Task SendDirectMessageAsync_CreatesOrResolvesChatBeforeSending()
    {
        Mock<ITeamsGraphClient> graphClient = new();
        graphClient
            .Setup(client => client.GetCurrentUserAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new CurrentUserProfile("me", "Me", "me@contoso.com"));
        graphClient
            .Setup(client => client.CreateOrGetOneOnOneChatAsync("me", "alex.wilber@contoso.com", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new Chat { Id = "chat-42" });
        graphClient
            .Setup(client => client.SendChatMessageAsync("chat-42", "Hello Alex", It.IsAny<CancellationToken>()))
            .ReturnsAsync(new ChatMessage { Id = "message-99" });

        TeamsChatService service = CreateService(graphClient.Object);

        TeamsSendMessageResult result = await service.SendDirectMessageAsync(" alex.wilber@contoso.com ", " Hello Alex ", CancellationToken.None);

        Assert.AreEqual("chat-42", result.ChatId);
        Assert.AreEqual("message-99", result.MessageId);
        Assert.AreEqual("alex.wilber@contoso.com", result.Recipient);
        Assert.AreEqual("directMessage", result.DeliveryMode);
    }

    [TestMethod]
    public async Task SendMessageToChatAsync_BlankChatId_Throws()
    {
        TeamsChatService service = CreateService(Mock.Of<ITeamsGraphClient>());

        try
        {
            await service.SendMessageToChatAsync(" ", "Hello", CancellationToken.None);
            Assert.Fail("Expected SendMessageToChatAsync to throw an ArgumentException for a blank chat ID.");
        }
        catch (ArgumentException exception)
        {
            Assert.AreEqual("chatId", exception.ParamName);
        }
    }

    private static TeamsChatService CreateService(
        ITeamsGraphClient graphClient,
        int defaultChatPageSize = 50,
        int maxChatPageSize = 100)
    {
        return new TeamsChatService(
            graphClient,
            Options.Create(new TeamsGraphOptions
            {
                DefaultChatPageSize = defaultChatPageSize,
                MaxChatPageSize = maxChatPageSize
            }));
    }
}
