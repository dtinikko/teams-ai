﻿using Microsoft.Bot.Builder;
using Microsoft.Teams.AI.Exceptions;
using Microsoft.Teams.AI.State;

namespace Microsoft.Teams.AI.Tests.Application.Authentication
{
    public class MockedAuthentication<TState> : IAuthentication<TState>
        where TState : TurnState, new()
    {
        private string _mockedToken;
        private SignInStatus _mockedStatus;
        private bool _validActivity;

        public MockedAuthentication(SignInStatus mockedStatus = SignInStatus.Complete, string mockedToken = "mocked token", bool validActivity = true)
        {
            _mockedToken = mockedToken;
            _mockedStatus = mockedStatus;
            _validActivity = validActivity;
        }

        public void Initialize(Application<TState> app, string name, IStorage? storage = null)
        {
            return;
        }

        public Task<bool> IsValidActivity(ITurnContext context)
        {
            return Task.FromResult(_validActivity);
        }

        public IAuthentication<TState> OnUserSignInFailure(Func<ITurnContext, TState, TeamsAIAuthException, Task> handler)
        {
            return this;
        }

        public IAuthentication<TState> OnUserSignInSuccess(Func<ITurnContext, TState, Task> handler)
        {
            return this;
        }

        public Task<SignInResponse> SignInUser(ITurnContext context, TState state)
        {
            var result = new SignInResponse(_mockedStatus);
            if (_mockedStatus == SignInStatus.Complete)
            {
                result.Token = _mockedToken;
            }
            return Task.FromResult(result);
        }

        public Task SignOutUser(ITurnContext context, TState state)
        {
            return Task.CompletedTask;
        }
    }
}
