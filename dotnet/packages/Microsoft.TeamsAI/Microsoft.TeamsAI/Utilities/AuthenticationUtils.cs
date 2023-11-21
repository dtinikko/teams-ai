using Microsoft.Teams.AI.State;

namespace Microsoft.Teams.AI.Utilities
{
    internal class AuthenticationUtils
    {
        /// <summary>
        /// Sets the token in the the turn state
        /// </summary>
        /// <param name="turnState">The turn state.</param>
        /// <param name="settingName">The name of the setting.</param>
        /// <param name="token">The token to set.</param>
        public static void SetTokenInState<TState>(TState turnState, string settingName, string token)
            where TState : ITurnState<StateBase, StateBase, TempState>
        {
            if (turnState.Temp.AuthTokens.Count < 1)
            {
                turnState.Temp.AuthTokens.Add(settingName, token);
            }

            turnState.Temp.AuthTokens[settingName] = token;
        }

        /// <summary>
        /// Deletes the token from the turn state
        /// </summary>
        /// <param name="turnState">The turn state.</param>
        /// <param name="settingName">The name of the setting.</param>
        /// <param name="token">The token to set.</param>
        public static void DeleteTokenFromState<TState>(TState turnState, string settingName)
            where TState : ITurnState<StateBase, StateBase, TempState>
        {
            turnState.Temp.AuthTokens.Remove(settingName);
        }
    }
}
