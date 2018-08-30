// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as utils from "./utils";
import * as logger from "winston";
import * as constants from "./constants";
import * as config from "config";
import { RootDialog } from "./dialogs/RootDialog";
import { fetchTemplates, cardTemplates } from "./dialogs/CardTemplates";
import { renderACAttachment } from "./utils";

export class TeamsBot extends builder.UniversalBot {

    constructor(
        public _connector: builder.IConnector,
        public _botSettings: any,
    )
    {
        super(_connector, _botSettings);
        this.set("persistentConversationData", true);

        // Handle generic invokes
        let teamsConnector = this._connector as msteams.TeamsChatConnector;
        teamsConnector.onInvoke(async (event, cb) => {
            try {
                await this.onInvoke(event, cb);
            } catch (e) {
                logger.error("Invoke handler failed", e);
                cb(e, null, 500);
            }
        });

        // Register dialogs
        new RootDialog().register(this);
    }

    // Handle incoming invoke
    private async onInvoke(event: builder.IEvent, cb: (err: Error, body: any, status?: number) => void): Promise<void> {
        console.log(JSON.stringify(utils.getContext(event)));
        let session = await utils.loadSessionAsync(this, event);
        if (session) {
            // Invokes don't participate in middleware

            // If the message is not task/*, simulate a normal message and route it, but remember the original invoke message
            let invokeType = (event as any).name;
            let invokeValue = (event as any).value;
            if (invokeType === undefined) {
                invokeType = null;
            }
            switch (invokeType) {
                case "task/fetch": {
                    if (invokeValue !== undefined && fetchTemplates[invokeValue.taskModule.toLowerCase()] !== undefined) {
                        // Return the specified task module response to the bot
                        cb(null, fetchTemplates[invokeValue.taskModule.toLowerCase()], 200);
                    }
                    else {
                        console.log(`Error: task module template for ${(invokeValue.taskModule === undefined ? "<undefined>" : invokeValue.taskModule)} not found.`);
                    }
                    break;
                }
                case "task/submit": {
                    if (invokeValue.data !== undefined) {
                        // It's a valid task module response
                        switch (invokeValue.data.taskResponse) {
                            case "message":
                                // Return HTTP 200 (OK) as the invoke response and echo the results to the chat stream
                                // cb(null, fetchTemplates.submitNullResponse, 200);
                                session.send("**task/submit results from the Adaptive Card:**\n```" + JSON.stringify(invokeValue) + "```");
                                break;
                            case "continue":
                                let fetchResponse = fetchTemplates.submitResponse;
                                fetchResponse.task.value.card = renderACAttachment(cardTemplates.acSubmitResponse, { results: JSON.stringify(invokeValue.data) });
                                cb(null, fetchResponse, 200);
                                break;
                            case "final":
                                // cb(null, fetchTemplates.submitNullResponse, 200);
                                break;
                            default:
                                // It's a response from an HTML task module
                                cb(null, fetchTemplates.submitMessageResponse, 200);
                                session.send("**task/submit results from HTML:**\n\n```" + JSON.stringify(invokeValue.data) + "```");
                        }
                    }
                    break;
                }
                case null: {
                    let fakeMessage: any = {
                        ...event,
                        text: invokeValue.command + " " + JSON.stringify(invokeValue),
                        originalInvoke: event,
                    };

                    session.message = fakeMessage;
                    session.dispatch(session.sessionState, session.message, () => {
                        session.routeToActiveDialog();
                    });
                }
            }
        }
        cb(null, "");
    }
}
