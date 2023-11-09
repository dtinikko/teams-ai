// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TokenExchangeRequest, TokenResponse, TurnContext } from "botbuilder";
import { OAuthPromptSettings } from "botbuilder-dialogs";
import { MessagingExtensionAuthenticationBase } from "./MessagingExtensionAuthenticationBase";
import * as UserTokenAccess from './UserTokenAccess';

export class OAuthPromptMessagingExtensionAuthentication extends MessagingExtensionAuthenticationBase {

    public constructor(private readonly settings: OAuthPromptSettings) {
        super();
    }

    public async handleSsoTokenExchange(
        context: TurnContext,
        tokenExchangeRequest: TokenExchangeRequest
    ): Promise<TokenResponse | undefined> {
        return await UserTokenAccess.exchangeToken(context, this.settings, tokenExchangeRequest);
    }

    public async handleUserSignIn(context: TurnContext, magicCode: string): Promise<TokenResponse | undefined> {
        return await UserTokenAccess.getUserToken(context, this.settings, magicCode);
    }

    public async getSignInLink(context: TurnContext): Promise<string|undefined> {
        const signInResource = await UserTokenAccess.getSignInResource(context, this.settings);
        return signInResource.signInLink;
    }
}