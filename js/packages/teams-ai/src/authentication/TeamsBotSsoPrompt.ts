import { Dialog, DialogContext, PromptRecognizerResult } from "botbuilder-dialogs";
import { ActionTypes, Activity, ActivityTypes, CardFactory, MessageFactory, OAuthCard, StatusCodes, TeamsChannelAccount, TeamsInfo, TokenExchangeInvokeRequest, TokenExchangeResource, TokenResponse, TurnContext, tokenExchangeOperationName, verifyStateOperationName } from "botbuilder";
import { v4 as uuidv4 } from "uuid";
import { AuthenticationResult, ConfidentialClientApplication, Configuration } from "@azure/msal-node";

const invokeResponseType = "invokeResponse";

export interface TeamsSsoPromptSettings {
    scopes: string[];
    msalConfig: Configuration;
    signInLink: string;
    timeout?: number;
    endOnInvalidMessage?: boolean;
}

class TokenExchangeInvokeResponse {
    /**
     * Response id
     */
    id: string;

    /**
     * Detailed error message
     */
    failureDetail: string;

    constructor(id: string, failureDetail: string) {
        this.id = id;
        this.failureDetail = failureDetail;
    }
}

export class TeamsSsoPrompt extends Dialog {
    private settings: TeamsSsoPromptSettings;
    private msal: ConfidentialClientApplication;

    constructor(
        dialogId: string,
        settings: TeamsSsoPromptSettings,
    ) {
        super(dialogId);
        this.settings = settings;
        this.msal = new ConfidentialClientApplication(settings.msalConfig);

        this.validateScopesType(settings.scopes);
    }

    public async beginDialog(dc: any, options: any): Promise<any> {
        const default_timeout = 900000;
        let timeout: number = default_timeout;
        if (this.settings.timeout) {
            if (typeof this.settings.timeout != "number") {
                const errorMsg = "type of timeout property in teamsBotSsoPromptSettings should be number.";
                throw new Error(errorMsg);
            }
            if (this.settings.timeout <= 0) {
                const errorMsg =
                    "value of timeout property in teamsBotSsoPromptSettings should be positive.";
                throw new Error(errorMsg);
            }
            timeout = this.settings.timeout;
        }

        if (this.settings.endOnInvalidMessage === undefined) {
            this.settings.endOnInvalidMessage = true;
        }
        const state = dc.activeDialog?.state;
        state.state = {};
        state.options = {};
        state.expires = new Date().getTime() + timeout;

        const loginHint = await this.getLoginHint(dc.context);
        if (loginHint) {
            const token = await this.acquireTokenFromCache(loginHint);
            if (token) {
                const tokenResponse: TokenResponse = {
                    connectionName: "", // No connection name is avaiable in this implementation
                    token: token.accessToken,
                    expiration: token.expiresOn?.toISOString() ?? "",
                };
                return await dc.endDialog(tokenResponse);
            }
        }

        // Cannot get token from cache, send OAuth card to get SSO token
        await this.sendOAuthCardAsync(dc.context, loginHint);
        return Dialog.EndOfTurn;
    }

    public async continueDialog(dc: any): Promise<any> {
        const state = dc.activeDialog?.state;
        const isMessage: boolean = dc.context.activity.type === ActivityTypes.Message;
        const isTimeoutActivityType: boolean =
            isMessage ||
            this.isTeamsVerificationInvoke(dc.context) ||
            this.isTokenExchangeRequestInvoke(dc.context);

        const hasTimedOut: boolean = isTimeoutActivityType && new Date().getTime() > state.expires;
        if (hasTimedOut) {
            return await dc.endDialog(undefined);
        } else {
            if (
                this.isTeamsVerificationInvoke(dc.context) ||
                this.isTokenExchangeRequestInvoke(dc.context)
            ) {
                const recognized: PromptRecognizerResult<any> =
                    await this.recognizeToken(dc);

                if (recognized.succeeded) {
                    return await dc.endDialog(recognized.value);
                }
            } else if (isMessage && this.settings.endOnInvalidMessage) {
                return await dc.endDialog(undefined);
            }

            return Dialog.EndOfTurn;
        }
    }

    private async recognizeToken(
        dc: DialogContext
    ): Promise<PromptRecognizerResult<any>> {
        const context = dc.context;
        let tokenResponse: TokenResponse | undefined;

        if (this.isTokenExchangeRequestInvoke(context)) {
            // Received activity is not a token exchange request
            if (!(context.activity.value && this.isTokenExchangeRequest(context.activity.value))) {
                const warningMsg =
                    "The bot received an InvokeActivity that is missing a TokenExchangeInvokeRequest value. This is required to be sent with the InvokeActivity.";

                await context.sendActivity(
                    this.getTokenExchangeInvokeResponse(StatusCodes.BAD_REQUEST, warningMsg)
                );
            } else {
                const ssoToken = context.activity.value.token;
                let exchangedToken: AuthenticationResult | null;

                try {
                    exchangedToken = await this.msal.acquireTokenOnBehalfOf({
                        oboAssertion: ssoToken,
                        scopes: this.settings.scopes,
                    });

                    if (exchangedToken) {
                        await context.sendActivity(
                            this.getTokenExchangeInvokeResponse(StatusCodes.OK, "", context.activity.value.id)
                        );
                        tokenResponse = {
                            connectionName: "",
                            token: exchangedToken.accessToken,
                            expiration: exchangedToken.expiresOn?.toISOString() ?? "",
                        };
                    }
                } catch (error) {
                    const warningMsg = "The bot is unable to exchange token. Ask for user consent.";
                    await context.sendActivity(
                        this.getTokenExchangeInvokeResponse(
                            StatusCodes.PRECONDITION_FAILED,
                            warningMsg,
                            context.activity.value.id
                        )
                    );
                }
            }
        } else if (this.isTeamsVerificationInvoke(context)) {
            const loginHint = await this.getLoginHint(context);
            await this.sendOAuthCardAsync(dc.context, loginHint);
            await context.sendActivity({ type: invokeResponseType, value: { status: StatusCodes.OK } });
        }

        return tokenResponse !== undefined
            ? { succeeded: true, value: tokenResponse }
            : { succeeded: false };
    }

    private async getLoginHint(context: TurnContext): Promise<string | undefined> {
        const account: TeamsChannelAccount = await TeamsInfo.getMember(
            context,
            context.activity.from.id
        );
        return account.userPrincipalName;
    }

    private async acquireTokenFromCache(
        loginHint: string
    ): Promise<AuthenticationResult | null> {
        try {
            const accounts = await this.msal.getTokenCache().getAllAccounts();
            const account = accounts.find((account) => account.username === loginHint);
            if (account) {
                const silentRequest = {
                    account: account,
                    scopes: this.settings.scopes,
                };
                return await this.msal.acquireTokenSilent(silentRequest);
            }
        } catch (error) {
            return null;
        }
        return null;
    }

    private getTokenExchangeInvokeResponse(
        status: number,
        failureDetail: string,
        id?: string
    ): Activity {
        const invokeResponse: Partial<Activity> = {
            type: invokeResponseType,
            value: { status, body: new TokenExchangeInvokeResponse(id as string, failureDetail) },
        };
        return invokeResponse as Activity;
    }

    private async sendOAuthCardAsync(context: TurnContext, loginHint?: string): Promise<void> {
        const signInResource = await this.getSignInResource(loginHint);
        const card = CardFactory.oauthCard(
            "",
            "Teams SSO Sign In",
            "Sign In",
            signInResource.signInLink,
            signInResource.tokenExchangeResource
        );
        (card.content as OAuthCard).buttons[0].type = ActionTypes.Signin;
        const msg: Partial<Activity> = MessageFactory.attachment(card);

        // Send prompt
        await context.sendActivity(msg);
    }

    private async getSignInResource(loginHint?: string) {
        const clientId = this.settings.msalConfig.auth.clientId;
        const scope = encodeURI(this.settings.scopes.join(" "));
        const authority = this.settings.msalConfig.auth.authority ?? "https://login.microsoftonline.com/common/";
        const tenantId = authority.match(/https:\/\/[^\/]+\/([^\/]+)\/?/)?.[1];

        const signInLink = `${this.settings.signInLink}?scope=${scope}&clientId=${clientId}&tenantId=${tenantId}&loginHint=${loginHint}`;

        const tokenExchangeResource: TokenExchangeResource = {
            id: uuidv4(),
        };

        return {
            signInLink: signInLink,
            tokenExchangeResource: tokenExchangeResource,
        };
    }

    private validateScopesType(value: any): void {
        // empty array
        if (Array.isArray(value) && value.length === 0) {
            return;
        }

        // string array
        if (Array.isArray(value) && value.length > 0 && value.every((item) => typeof item === "string")) {
            return;
        }

        const errorMsg = "The type of scopes is not valid, it must be string array";
        throw new Error(errorMsg);
    }

    private isTeamsVerificationInvoke(context: TurnContext): boolean {
        const activity: Activity = context.activity;

        return activity.type === ActivityTypes.Invoke && activity.name === verifyStateOperationName;
    }

    private isTokenExchangeRequestInvoke(context: TurnContext): boolean {
        const activity: Activity = context.activity;

        return activity.type === ActivityTypes.Invoke && activity.name === tokenExchangeOperationName;
    }

    private isTokenExchangeRequest(obj: any): obj is TokenExchangeInvokeRequest {
        return obj.hasOwnProperty("token");
    }
}