// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { Attachment, CardFactory, } from 'botbuilder';

/**
 * @returns {Attachment} The adaptive card attachment for the sign-in request.
 */
export function createSignOutCard(): Attachment {
    return CardFactory.adaptiveCard({
        version: '1.0.0',
        type: 'AdaptiveCard',
        body: [
            {
                type: 'TextBlock',
                text: 'You have been signed out.'
            }
        ],
        actions: [
            {
                type: 'Action.Submit',
                title: 'Close',
                data: {
                    key: 'close'
                }
            }
        ]
    });
}
