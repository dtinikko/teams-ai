﻿using Microsoft.Teams.AI.AI;
using Microsoft.Teams.AI.AI.Moderator;
using Microsoft.Teams.AI.AI.Planners;
using Microsoft.Teams.AI.AI.Prompts;
using Microsoft.Bot.Schema;
using Moq;
using System.Reflection;
using Microsoft.Teams.AI.Tests.TestUtils;
using Microsoft.Bot.Builder;
using Azure.AI.ContentSafety;
using Azure;
using Microsoft.Teams.AI.State;
using Record = Microsoft.Teams.AI.State.Record;

#pragma warning disable CS8604 // Possible null reference argument.
namespace Microsoft.Teams.AI.Tests.AITests
{
    public class AzureContentSafetyModeratorTests
    {
        [Fact]
        public async void Test_ReviewPrompt_ThrowsException()
        {
            // Arrange
            var apiKey = "randomApiKey";
            var endpoint = "https://test.cognitiveservices.azure.com";

            var botAdapterMock = new Mock<BotAdapter>();
            var activity = new Activity()
            {
                Text = "input",
            };

            var turnContext = TurnStateConfig.CreateConfiguredTurnContext();
            var turnStateMock = TurnStateConfig.GetTurnStateWithConversationStateAsync(turnContext);
            var promptTemplate = new PromptTemplate(
                "prompt",
                new(new() { })
            )
            {
                Configuration = new PromptTemplateConfiguration
                {
                    Completion =
                        {
                            MaxTokens = 2000,
                            Temperature = 0.2,
                            TopP = 0.5,
                        }
                }
            };

            var clientMock = new Mock<ContentSafetyClient>(new Uri(endpoint), new AzureKeyCredential(apiKey));
            var exception = new RequestFailedException("Exception Message");
            clientMock.Setup(client => client.AnalyzeTextAsync(It.IsAny<AnalyzeTextOptions>(), It.IsAny<CancellationToken>())).ThrowsAsync(exception);

            var options = new AzureContentSafetyModeratorOptions(apiKey, endpoint, ModerationType.Both);
            var moderator = new AzureContentSafetyModerator<TurnState>(options);
            moderator.GetType().GetField("_client", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(moderator, clientMock.Object);

            // Act
            var result = await Assert.ThrowsAsync<RequestFailedException>(async () => await moderator.ReviewInput(turnContext, turnStateMock.Result));

            // Assert
            Assert.Equal("Exception Message", result.Message);
        }

        [Theory]
        [InlineData(ModerationType.Input)]
        [InlineData(ModerationType.Output)]
        [InlineData(ModerationType.Both)]
        public async void Test_ReviewPrompt_Flagged(ModerationType moderate)
        {
            // Arrange
            var apiKey = "randomApiKey";
            var endpoint = "https://test.cognitiveservices.azure.com";

            var botAdapterMock = new Mock<BotAdapter>();
            // TODO: when TurnState is implemented, get the user input
            var activity = new Activity()
            {
                Text = "input",
            };
            var turnContext = TurnStateConfig.CreateConfiguredTurnContext();
            var turnStateMock = TurnStateConfig.GetTurnStateWithConversationStateAsync(turnContext);
            var promptTemplate = new PromptTemplate(
                "prompt",
                new(new() { })
            )
            {
                Configuration = new PromptTemplateConfiguration
                {
                    Completion =
                        {
                            MaxTokens = 2000,
                            Temperature = 0.2,
                            TopP = 0.5,
                        }
                }
            };

            var clientMock = new Mock<ContentSafetyClient>(new Uri(endpoint), new AzureKeyCredential(apiKey));
            AnalyzeTextResult analyzeTextResult = ContentSafetyModelFactory.AnalyzeTextResult(hateResult: ContentSafetyModelFactory.TextAnalyzeSeverityResult(TextCategory.Hate, 2));
            Response? response = null;
            clientMock.Setup(client => client.AnalyzeTextAsync(It.IsAny<AnalyzeTextOptions>(), It.IsAny<CancellationToken>())).ReturnsAsync(Response.FromValue(analyzeTextResult, response));

            var options = new AzureContentSafetyModeratorOptions(apiKey, endpoint, moderate);
            var moderator = new AzureContentSafetyModerator<TurnState>(options);
            moderator.GetType().GetField("_client", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(moderator, clientMock.Object);

            // Act
            var result = await moderator.ReviewInput(turnContext, turnStateMock.Result);

            // Assert
            if (moderate == ModerationType.Input || moderate == ModerationType.Both)
            {
                Assert.NotNull(result);
                Assert.Equal(AIConstants.DoCommand, result.Commands[0].Type);
                Assert.Equal(AIConstants.FlaggedInputActionName, ((PredictedDoCommand)result.Commands[0]).Action);
                Assert.NotNull(((PredictedDoCommand)result.Commands[0]).Parameters);
                Assert.True(((PredictedDoCommand)result.Commands[0]).Parameters!.ContainsKey("Result"));
                Assert.StrictEqual(analyzeTextResult, ((PredictedDoCommand)result.Commands[0]).Parameters!.GetValueOrDefault("Result"));
            }
            else
            {
                Assert.Null(result);
            }
        }

        [Theory]
        [InlineData(ModerationType.Input)]
        [InlineData(ModerationType.Output)]
        [InlineData(ModerationType.Both)]
        public async void Test_ReviewPrompt_NotFlagged(ModerationType moderate)
        {
            // Arrange
            var apiKey = "randomApiKey";
            var endpoint = "https://test.cognitiveservices.azure.com";

            var botAdapterMock = new Mock<BotAdapter>();
            // TODO: when TurnState is implemented, get the user input
            var activity = new Activity()
            {
                Text = "input",
            };
            var turnContext = TurnStateConfig.CreateConfiguredTurnContext();
            var turnStateMock = TurnStateConfig.GetTurnStateWithConversationStateAsync(turnContext);
            var promptTemplate = new PromptTemplate(
                "prompt",
                new(new() { })
            )
            {
                Configuration = new PromptTemplateConfiguration
                {
                    Completion =
                        {
                            MaxTokens = 2000,
                            Temperature = 0.2,
                            TopP = 0.5,
                        }
                }
            };

            var clientMock = new Mock<ContentSafetyClient>(new Uri(endpoint), new AzureKeyCredential(apiKey));
            AnalyzeTextResult analyzeTextResult = ContentSafetyModelFactory.AnalyzeTextResult(hateResult: ContentSafetyModelFactory.TextAnalyzeSeverityResult(TextCategory.Hate, 0));
            Response? response = null;
            clientMock.Setup(client => client.AnalyzeTextAsync(It.IsAny<AnalyzeTextOptions>(), It.IsAny<CancellationToken>())).ReturnsAsync(Response.FromValue(analyzeTextResult, response));

            var options = new AzureContentSafetyModeratorOptions(apiKey, endpoint, moderate);
            var moderator = new AzureContentSafetyModerator<TurnState>(options);
            moderator.GetType().GetField("_client", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(moderator, clientMock.Object);

            // Act
            var result = await moderator.ReviewInput(turnContext, turnStateMock.Result);

            // Assert
            Assert.Null(result);
        }

        [Fact]
        public async void Test_ReviewPlan_ThrowsException()
        {
            // Arrange
            var apiKey = "randomApiKey";
            var endpoint = "https://test.cognitiveservices.azure.com";

            var turnContext = TurnStateConfig.CreateConfiguredTurnContext();
            var turnStateMock = TurnStateConfig.GetTurnStateWithConversationStateAsync(turnContext);
            var plan = new Plan(new List<IPredictedCommand>()
            {
                new PredictedDoCommand("action"),
                new PredictedSayCommand("response"),
            });

            var clientMock = new Mock<ContentSafetyClient>(new Uri(endpoint), new AzureKeyCredential(apiKey));
            var exception = new RequestFailedException("Exception Message");
            clientMock.Setup(client => client.AnalyzeTextAsync(It.IsAny<AnalyzeTextOptions>(), It.IsAny<CancellationToken>())).ThrowsAsync(exception);

            var options = new AzureContentSafetyModeratorOptions(apiKey, endpoint, ModerationType.Both);
            var moderator = new AzureContentSafetyModerator<TurnState>(options);
            moderator.GetType().GetField("_client", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(moderator, clientMock.Object);

            // Act
            var result = await Assert.ThrowsAsync<RequestFailedException>(async () => await moderator.ReviewOutput(turnContext, turnStateMock.Result, plan));

            // Assert
            Assert.Equal("Exception Message", result.Message);
        }

        [Theory]
        [InlineData(ModerationType.Input)]
        [InlineData(ModerationType.Output)]
        [InlineData(ModerationType.Both)]
        public async void Test_ReviewPlan_Flagged(ModerationType moderate)
        {
            // Arrange
            var apiKey = "randomApiKey";
            var endpoint = "https://test.cognitiveservices.azure.com";

            var turnContext = TurnStateConfig.CreateConfiguredTurnContext();
            var turnStateMock = TurnStateConfig.GetTurnStateWithConversationStateAsync(turnContext);
            var plan = new Plan(new List<IPredictedCommand>()
            {
                new PredictedDoCommand("action"),
                new PredictedSayCommand("response"),
            });

            var clientMock = new Mock<ContentSafetyClient>(new Uri(endpoint), new AzureKeyCredential(apiKey));
            AnalyzeTextResult analyzeTextResult = ContentSafetyModelFactory.AnalyzeTextResult(hateResult: ContentSafetyModelFactory.TextAnalyzeSeverityResult(TextCategory.Hate, 2));
            Response? response = null;
            clientMock.Setup(client => client.AnalyzeTextAsync(It.IsAny<AnalyzeTextOptions>(), It.IsAny<CancellationToken>())).ReturnsAsync(Response.FromValue(analyzeTextResult, response));

            var options = new AzureContentSafetyModeratorOptions(apiKey, endpoint, moderate);
            var moderator = new AzureContentSafetyModerator<TurnState>(options);
            moderator.GetType().GetField("_client", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(moderator, clientMock.Object);

            // Act
            var result = await moderator.ReviewOutput(turnContext, turnStateMock.Result, plan);

            // Assert
            if (moderate == ModerationType.Output || moderate == ModerationType.Both)
            {
                Assert.NotNull(result);
                Assert.Equal(AIConstants.DoCommand, result.Commands[0].Type);
                Assert.Equal(AIConstants.FlaggedOutputActionName, ((PredictedDoCommand)result.Commands[0]).Action);
                Assert.NotNull(((PredictedDoCommand)result.Commands[0]).Parameters);
                Assert.True(((PredictedDoCommand)result.Commands[0]).Parameters!.ContainsKey("Result"));
                Assert.StrictEqual(analyzeTextResult, ((PredictedDoCommand)result.Commands[0]).Parameters!.GetValueOrDefault("Result"));
            }
            else
            {
                Assert.StrictEqual(plan, result);
            }
        }

        [Theory]
        [InlineData(ModerationType.Input)]
        [InlineData(ModerationType.Output)]
        [InlineData(ModerationType.Both)]
        public async void Test_ReviewPlan_NotFlagged(ModerationType moderate)
        {
            // Arrange
            var apiKey = "randomApiKey";
            var endpoint = "https://test.cognitiveservices.azure.com";

            var turnContext = TurnStateConfig.CreateConfiguredTurnContext();
            var turnStateMock = TurnStateConfig.GetTurnStateWithConversationStateAsync(turnContext);
            var plan = new Plan(new List<IPredictedCommand>()
            {
                new PredictedDoCommand("action"),
                new PredictedSayCommand("response"),
            });

            var clientMock = new Mock<ContentSafetyClient>(new Uri(endpoint), new AzureKeyCredential(apiKey));
            AnalyzeTextResult analyzeTextResult = ContentSafetyModelFactory.AnalyzeTextResult(hateResult: ContentSafetyModelFactory.TextAnalyzeSeverityResult(TextCategory.Hate, 0));
            Response? response = null;
            clientMock.Setup(client => client.AnalyzeTextAsync(It.IsAny<AnalyzeTextOptions>(), It.IsAny<CancellationToken>())).ReturnsAsync(Response.FromValue(analyzeTextResult, response));

            var options = new AzureContentSafetyModeratorOptions(apiKey, endpoint, moderate);
            var moderator = new AzureContentSafetyModerator<TurnState>(options);
            moderator.GetType().GetField("_client", BindingFlags.Instance | BindingFlags.NonPublic)!.SetValue(moderator, clientMock.Object);

            // Act
            var result = await moderator.ReviewOutput(turnContext, turnStateMock.Result, plan);

            // Assert
            Assert.StrictEqual(plan, result);
        }
    }
}
#pragma warning restore CS8604 // Possible null reference argument.
