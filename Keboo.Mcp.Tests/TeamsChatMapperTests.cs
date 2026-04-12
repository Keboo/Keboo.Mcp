using Keboo.Mcp.Teams;
using Microsoft.Graph.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Keboo.Mcp.Tests;

[TestClass]
public sealed class TeamsChatMapperTests
{
    [TestMethod]
    public void ToSummary_OneOnOneChat_UsesOtherParticipantAsTitle()
    {
        Chat chat = new()
        {
            Id = "chat-1",
            ChatType = ChatType.OneOnOne,
            Members =
            [
                new AadUserConversationMember
                {
                    UserId = "me",
                    DisplayName = "Me",
                    Roles = ["owner"]
                },
                new AadUserConversationMember
                {
                    UserId = "user-2",
                    DisplayName = "Alex Wilber",
                    Email = "alex.wilber@contoso.com",
                    Roles = ["owner"]
                }
            ],
            LastMessagePreview = new ChatMessageInfo
            {
                Body = new ItemBody
                {
                    Content = "Hello from Alex",
                    ContentType = BodyType.Text
                }
            }
        };

        TeamsChatSummary result = TeamsChatMapper.ToSummary(chat, "me");

        Assert.AreEqual("Alex Wilber", result.Title);
        Assert.AreEqual("Hello from Alex", result.LastMessagePreview?.Content);
        Assert.AreEqual("Text", result.LastMessagePreview?.ContentType);
        Assert.AreEqual(2, result.Members.Count);
        Assert.AreEqual("alex.wilber@contoso.com", result.Members[1].Email);
    }

    [TestMethod]
    public void ToSendMessageResult_MapsExpectedFields()
    {
        ChatMessage message = new()
        {
            Id = "message-9",
            CreatedDateTime = DateTimeOffset.Parse("2026-04-09T16:30:00Z")
        };

        TeamsSendMessageResult result = TeamsChatMapper.ToSendMessageResult("chat-1", message, "alex.wilber@contoso.com", "directMessage");

        Assert.AreEqual("chat-1", result.ChatId);
        Assert.AreEqual("message-9", result.MessageId);
        Assert.AreEqual("alex.wilber@contoso.com", result.Recipient);
        Assert.AreEqual("directMessage", result.DeliveryMode);
        Assert.AreEqual(DateTimeOffset.Parse("2026-04-09T16:30:00Z"), result.CreatedDateTime);
    }
}
