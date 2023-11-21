using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Schema;
using Microsoft.Teams.AI.Exceptions;
using Microsoft.Teams.AI.State;
using Microsoft.Teams.AI.Utilities;

namespace Microsoft.Teams.AI.Authentication
{
    internal delegate Task UserSignInSuccessHandler<TState>(ITurnContext turnContext, TState turnState) where TState : ITurnState<StateBase, StateBase, TempState>;
    internal delegate Task UserSignInFailureHandler<TState>(ITurnContext turnContext, TState turnState, Exception exception) where TState : ITurnState<StateBase, StateBase, TempState>;

    internal abstract class BotAuthenticationBase<TState, TTurnStateManager>
        where TState : ITurnState<StateBase, StateBase, TempState>
        where TTurnStateManager : ITurnStateManager<TState>, new()
    {
        protected IStorage _storage;
        protected string _settingName;
        private UserSignInSuccessHandler<TState>? _userSignInSuccessHandler;
        private UserSignInFailureHandler<TState>? _userSignInFailureHandler;

        public BotAuthenticationBase(Application<TState, TTurnStateManager> app, string settingName, IStorage? storage)
        {
            this._settingName = settingName;
            this._storage = storage ?? new MemoryStorage();

            app.AddRoute(async (turnContext, cancellationToken) =>
            {
                return await this._VerifyStateRouteSelector(turnContext);
            }, this.HandleSignInActivityAsync, true);

            app.AddRoute(async (turnContext, cancellationToken) =>
            {
                return await this._TokenExchangeRouteSelector(turnContext);
            }, this.HandleSignInActivityAsync, true);
        }

        public async Task<string?> AuthenticateAsync(TurnContext turnContext, TState turnState, CancellationToken cancellationToken)
        {
            // Get property names to use
            string userAuthStatePropertyName = this._GetUserAuthStatePropertyName(turnContext);
            string userDialogStatePropertyName = this._GetDialogAuthStatePropertyName(turnContext);

            // Save message if not signed in
            if (this._GetUserAuthState(turnContext, turnState) == null)
            {
                turnState.Conversation?.Set(userAuthStatePropertyName, new UserAuthState() { Message = turnContext.Activity.Text });
            }

            DialogTurnResult result = await this.RunDialogAsync(turnContext, turnState, userDialogStatePropertyName, cancellationToken);

            if (result.Status == DialogTurnStatus.Complete)
            {
                // Delete user auth state
                this.DeleteAuthFlowState(turnContext, turnState);

                if (result.Result is TokenResponse tokenResponse && tokenResponse.Token != null)
                {
                    return tokenResponse.Token;
                }
                else
                {
                    // Completed dialog without a token
                    // This could mean the user declined the consent prompt in the previous turn
                    // Retry authentication flow.
                    return await this.AuthenticateAsync(turnContext, turnState, cancellationToken);
                }
            }

            return null;
        }

        public bool IsValidActivity(TurnContext turnContext)
        {
            return turnContext.Activity.Type == ActivityTypes.Message && !string.IsNullOrEmpty(turnContext.Activity.Text);
        }

        public void OnUserSignInSuccess(UserSignInSuccessHandler<TState> handler)
        {
            this._userSignInSuccessHandler = handler;
        }

        public void OnUserSignInFailure(UserSignInFailureHandler<TState> handler)
        {
            this._userSignInFailureHandler = handler;
        }

        public async Task HandleSignInActivityAsync(ITurnContext turnContext, TState turnState, CancellationToken cancellationToken = default)
        {
            try
            {
                string userDialogStatePropertyName = this._GetDialogAuthStatePropertyName(turnContext);
                DialogTurnResult result = await this.ContinueDialogAsync(turnContext, turnState, userDialogStatePropertyName, cancellationToken);

                if (result.Status == DialogTurnStatus.Complete)
                {
                    // OAuthPrompt dialog should have sent an invoke response already.

                    if (result.Result is TokenResponse tokenResponse && tokenResponse.Token != null)
                    {
                        // Successful sign in
                        AuthenticationUtils.SetTokenInState(turnState, this._settingName, tokenResponse.Token);

                        // Get user auth state
                        UserAuthState? userAuthState = this._GetUserAuthState(turnContext, turnState);

                        turnContext.Activity.Text = userAuthState?.Message ?? "";

                        if (this._userSignInSuccessHandler != null)
                        {
                            await this._userSignInSuccessHandler(turnContext, turnState);
                        }
                    }
                    else
                    {
                        if (this._userSignInFailureHandler != null)
                        {
                            await this._userSignInFailureHandler(turnContext, turnState, new AuthException("Authentication flow completed without a token.", AuthExceptionReason.CompletionWithoutToken));
                        }
                    }
                }

            }
            catch (Exception e)
            {
                string message = $"Unexpected error encountered while signing in: {e.Message}. Incomming activites details: type: {turnContext.Activity.Type}, name: {turnContext.Activity.Name}";

                if (this._userSignInFailureHandler != null)
                {
                    await this._userSignInFailureHandler.Invoke(turnContext, turnState, new AuthException(message));
                }
            }
        }

        protected Task<bool> _VerifyStateRouteSelector(ITurnContext context)
        {
            // TODO: use global name instead of string
            return Task.FromResult(context.Activity.Type == ActivityTypes.Invoke && context.Activity.Name == "signin/verifyState");
        }

        protected Task<bool> _TokenExchangeRouteSelector(ITurnContext context)
        {
            // TODO: use global name instad of string
            return Task.FromResult(context.Activity.Type == ActivityTypes.Invoke && context.Activity.Name == "signin/tokenExchange");
        }

        public abstract Task<DialogTurnResult> RunDialogAsync(ITurnContext turnContext, TState turnState, string dialogStateProperty, CancellationToken cancellationToken);
        public abstract Task<DialogTurnResult> ContinueDialogAsync(ITurnContext turnContext, TState turnState, string dialogStateProperty, CancellationToken cancellationToken);

        #region private

        private void DeleteAuthFlowState(ITurnContext turnContext, TState turnState)
        {
            // Delete user auth state
            string userAuthStatePropertyName = this._GetUserAuthStatePropertyName(turnContext);
            if (this._GetUserAuthState(turnContext, turnState) != null)
            {
                turnState.Conversation?.Remove(userAuthStatePropertyName);
            }

            // Delete user dialog state
            string userDialogStatePropertyName = this._GetDialogAuthStatePropertyName(turnContext);
            if (this._GetDialogAuthState(turnContext, turnState) != null)
            {
                turnState.Conversation?.Remove(userDialogStatePropertyName);
            }
        }


        private string _GetUserAuthStatePropertyName(ITurnContext turnContext)
        {
            return $"__{turnContext.Activity.From.Id}:{this._settingName}:Bot:AuthState__";
        }

        private string _GetDialogAuthStatePropertyName(ITurnContext turnContext)
        {
            return $"__{turnContext.Activity.From.Id}:{this._settingName}:DialogState__";
        }

        private UserAuthState? _GetUserAuthState(ITurnContext context, TState turnState)
        {
            return turnState.Conversation?.Get<UserAuthState>(_GetUserAuthStatePropertyName(context));
        }

        private DialogState? _GetDialogAuthState(ITurnContext context, TState turnState)
        {
            return turnState.Conversation?.Get<DialogState>(_GetDialogAuthStatePropertyName(context));
        }

        #endregion
    }

    internal class UserAuthState
    {
        public string Message { get; set; } = string.Empty;
    }
}
