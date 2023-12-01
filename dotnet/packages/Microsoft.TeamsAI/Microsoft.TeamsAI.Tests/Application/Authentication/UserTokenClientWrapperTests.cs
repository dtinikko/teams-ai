using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Microsoft.Teams.AI.Tests.TestUtils;
using Moq;
using Microsoft.Bot.Connector.Authentication;

namespace Microsoft.Teams.AI.Tests.Application.Authentication
{
    public class UserTokenClientWrapperTests
    {
        [Fact]
        public async void Test_GetUserToken()
        {
            // arrange
            TurnContext context = MockTurnContext();
            TokenResponse expectedResult = new(token: "test token");
            var turnStateMock = new Mock<TurnContextStateCollection>();
            turnStateMock.Setup(mock => mock.Get<It.IsAnyType>()).Returns(It.IsAny<UserTokenClient>);
            var userTokenClientMock = new Mock<UserTokenClient>();
            userTokenClientMock.Setup(mock => mock.GetUserTokenAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>())).ReturnsAsync(expectedResult);

            // act
            var result = await UserTokenClientWrapper.GetUserTokenAsync(context, "test connection", "123456");

            // assert
            Assert.NotNull(result);
            Assert.Equal(expectedResult.Token, result.Token);
        }

        private static TurnContext MockTurnContext()
        {
            return new TurnContext(new SimpleAdapter(), new Activity()
            {
                Type = ActivityTypes.Invoke,
                Recipient = new() { Id = "recipientId" },
                Conversation = new() { Id = "conversationId" },
                From = new() { Id = "fromId" },
                ChannelId = "channelId",
                Name = ""
            });
        }
    }
}
