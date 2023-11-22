﻿using Microsoft.Bot.Builder;
using Microsoft.Teams.AI.State;

namespace Microsoft.Teams.AI.AI.Action
{
    /// <summary>
    /// Handler to perform the action.
    /// </summary>
    /// <typeparam name="TState">Type of the turn state.</typeparam>
    public interface IActionHandler<TState> where TState : TurnState
    {
        /// <summary>
        /// Perform the action.
        /// </summary>
        /// <param name="turnContext">Current turn context.</param>
        /// <param name="turnState">Current turn state.</param>
        /// <param name="entities">Optional enti to be used to perform the action.</param>
        /// <param name="action">The actual action name.</param>
        /// <returns>The result of the action handler.</returns>
        Task<string> PerformAction(ITurnContext turnContext, TState turnState, object? entities = null, string? action = null);
    }
}
