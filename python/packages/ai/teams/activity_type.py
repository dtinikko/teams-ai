"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

from typing import Literal

ActivityType = Literal[
    "message",
    "contactRelationUpdate",
    "conversationUpdate",
    "typing",
    "endOfConversation",
    "event",
    "invoke",
    "invokeResponse",
    "deleteUserData",
    "messageUpdate",
    "messageDelete",
    "installationUpdate",
    "messageReaction",
    "suggestion",
    "trace",
    "handoff",
    "command",
    "commandResult",
]