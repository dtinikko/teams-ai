﻿using Microsoft.Bot.Builder;
using Microsoft.Bot.Schema;
using Newtonsoft.Json.Linq;

namespace Microsoft.Teams.AI
{
    /// <summary>
    /// Handles authentication for Message Extensions in Teams using OAuth Connection.
    /// </summary>
    internal class OAuthMessageExtensionsAuthentication : MessageExtensionsAuthenticationBase
    {
        private string _oauthConnectionName;

        public OAuthMessageExtensionsAuthentication(string oauthConnectionName)
        {
            _oauthConnectionName = oauthConnectionName;
        }

        /// <summary>
        /// Gets the sign in link for the user.
        /// </summary>
        /// <param name="context">The turn context</param>
        /// <returns>The sign in link</returns>
        public override async Task<string> GetSignInLink(ITurnContext context)
        {
            SignInResource signInResource = await UserTokenClientWrapper.GetSignInResourceAsync(context, _oauthConnectionName);
            return signInResource.SignInLink;
        }

        /// <summary>
        /// Handles the user sign in.
        /// </summary>
        /// <param name="context">The turn context</param>
        /// <param name="magicCode">The magic code from user sign-in.</param>
        /// <returns>The token response if successfully verified the magic code</returns>
        public override async Task<TokenResponse> HandleUserSignIn(ITurnContext context, string magicCode)
        {
            return await UserTokenClientWrapper.GetUserTokenAsync(context, _oauthConnectionName, magicCode);
        }

        /// <summary>
        /// Handles the SSO token exchange.
        /// </summary>
        /// <param name="context">The turn context</param>
        /// <returns>The token response if token exchange success</returns>
        public override async Task<TokenResponse> HandleSsoTokenExchange(ITurnContext context)
        {
            JObject value = JObject.FromObject(context.Activity.Value);
            TokenExchangeRequest? tokenExchangeRequest = value["authentication"]?.ToObject<TokenExchangeRequest>();
            if (tokenExchangeRequest != null && !string.IsNullOrEmpty(tokenExchangeRequest.Token))
            {
                return await UserTokenClientWrapper.ExchangeTokenAsync(context, _oauthConnectionName, tokenExchangeRequest);
            }

            return new TokenResponse();
        }
    }
}
