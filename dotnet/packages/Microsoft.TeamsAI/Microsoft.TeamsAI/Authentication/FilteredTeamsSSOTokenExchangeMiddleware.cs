using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Text;

namespace Microsoft.Teams.AI.Authentication
{
    public class FilteredTeamsSSOTokenExchangeMiddleware : TeamsSSOTokenExchangeMiddleware
    {
        private string _oauthConnectionName;

        public FilteredTeamsSSOTokenExchangeMiddleware(IStorage storage, string oauthConnectionName) : base(storage, oauthConnectionName)
        {
            this._oauthConnectionName = oauthConnectionName;
        }

        public async Task OnTurnAsync(ITurnContext turnContext, NextDelegate next, CancellationToken cancellationToken = default)
        {
            // If connection name matches then continue to the Teams SSO Token Exchange Middleware.
            if ((turnContext.Activity.Value as JObject).Value<string>("ConnectionName") == this._oauthConnectionName)
            {
                await base.OnTurnAsync(turnContext, next, cancellationToken);
            }
            else
            {
                await next(cancellationToken);
            }
        }
    }
}
