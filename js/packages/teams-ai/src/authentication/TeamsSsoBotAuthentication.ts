// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { Storage, TokenResponse, TurnContext } from "botbuilder";
import { Application } from "../Application";
import { DefaultTurnState } from "../DefaultTurnStateManager";
import { TurnState } from "../TurnState";
import { BotAuthenticationBase } from "./BotAuthenticationBase";
import { TeamsSsoPrompt, TeamsSsoPromptSettings } from "./TeamsBotSsoPrompt";
import { DialogSet, DialogState, DialogTurnResult, DialogTurnStatus } from "botbuilder-dialogs";
import { TurnStateProperty } from "../TurnStateProperty";

export class TeamsSsoBotAuthentication<TState extends TurnState = DefaultTurnState> extends BotAuthenticationBase<TState> {
    private _prompt: TeamsSsoPrompt;

    public constructor(
        app: Application<TState>,
        promptSettings: TeamsSsoPromptSettings, // Child classes will have different types for this
        settingName: string,
        storage?: Storage
    ) {
        super(app, settingName, storage);

        this._prompt = new TeamsSsoPrompt('TeamsSsoPrompt', promptSettings);
    }

    public async runDialog(context: TurnContext, state: TState, dialogStateProperty: string): Promise<DialogTurnResult<TokenResponse>> {
        const accessor = new TurnStateProperty<DialogState>(state, 'conversation', dialogStateProperty);
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this._prompt);
        const dialogContext = await dialogSet.createContext(context);
        let results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            results = await dialogContext.beginDialog(this._prompt.id);
        }
        return results;
    }

    // private async shouldDedup(context: TurnContext): Promise<boolean> {
    //     const storeItem = {
    //         eTag: context.activity.value.id,
    //     };

    //     const key = this.getStorageKey(context);
    //     const storeItems = { [key]: storeItem };

    //     try {
    //         await this._storage.write(storeItems);
    //         this._dedupStorageKeys.push(key);
    //     } catch (err) {
    //         if (err instanceof Error && err.message.indexOf("eTag conflict")) {
    //             return true;
    //         }
    //         throw err;
    //     }
    //     return false;
    // }

    // private getStorageKey(context: TurnContext): string {
    //     if (!context || !context.activity || !context.activity.conversation) {
    //         throw new Error("Invalid context, can not get storage key!");
    //     }
    //     const activity = context.activity;
    //     const channelId = activity.channelId;
    //     const conversationId = activity.conversation.id;
    //     if (
    //         activity.type !== ActivityTypes.Invoke ||
    //         activity.name !== tokenExchangeOperationName
    //     ) {
    //         throw new Error(
    //             "TokenExchangeState can only be used with Invokes of signin/tokenExchange."
    //         );
    //     }
    //     const value = activity.value;
    //     if (!value || !value.id) {
    //         throw new Error(
    //             "Invalid signin/tokenExchange. Missing activity.value.id."
    //         );
    //     }
    //     return `${channelId}/${conversationId}/${value.id}`;
    // }
}