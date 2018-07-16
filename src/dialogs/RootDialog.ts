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
import * as constants from "../constants";
import * as utils from "../utils";
import * as logger from "winston";
import * as config from "config";
import { renderACAttachment } from "../utils/AdaptiveCardUtils";
import { cardTemplates } from "./CardTemplates";

export class RootDialog extends builder.IntentDialog
{
    constructor() {
        super();
    }

    // Register the dialogs with the bot
    public register(bot: builder.UniversalBot): void {
        bot.dialog(constants.DialogId.Root, this);

        this.onBegin((session, args, next) => { logger.verbose("onDialogBegin called"); this.onDialogBegin(session, args, next); });
        this.onDefault((session) => { logger.verbose("onDefault called"); this.onMessageReceived(session); } );
        logger.verbose("register called for dialog: " + constants.DialogId.Root);
    }

    // Handle start of dialog
    private async onDialogBegin(session: builder.Session, args: any, next: () => void): Promise<void> {
        next();
    }

    // Handle message
    private async onMessageReceived(session: builder.Session): Promise<void> {
        if (session.message.text === "") {
            console.log("AC payload: " + JSON.stringify(session.message.value));
        }
        else {
            // Message might contain @mentions which we would like to strip off in the response
            let text = utils.getTextWithoutMentions(session.message);

            let appInfo = {
                appId: config.get("bot.appId"),
                appRoot: config.get("app.baseUri"),
            };
            let taskModuleInfo = {
                button1: "YouTube",
                url1: encodeURI(`https://teams.microsoft.com/l/task/${appInfo.appId}?url=${appInfo.appRoot}/youtube&height=large&width=large&title=${encodeURIComponent("Satya Nadella's Build 2018 Keynote")}`),
                button2: "PowerApp",
                url2: encodeURI(`https://teams.microsoft.com/l/task/${appInfo.appId}?url=${appInfo.appRoot}/powerapps&height=large&width=large&title=${encodeURIComponent("PowerApp: Asset Checkout")}`),
                button3: "Custom Form",
                url3: encodeURI(`https://teams.microsoft.com/l/task/${appInfo.appId}?url=${appInfo.appRoot}/customform&height=medium&width=medium&title=${encodeURIComponent("Custom Form")}`),
            };

            let cardData: any = {
                title: "Task Module",
                subTitle: "Task Module Test Card",
                instructions: "Click on the buttons below below to open task modules in various ways.",
                linkbutton1: taskModuleInfo.button1,
                url1: taskModuleInfo.url1,
                markdown1: `[${taskModuleInfo.button1}](${taskModuleInfo.url1})`,
                linkbutton2: taskModuleInfo.button2,
                url2: taskModuleInfo.url2,
                markdown2: `[${taskModuleInfo.button2}](${taskModuleInfo.url2})`,
                linkbutton3: taskModuleInfo.button3,
                url3: taskModuleInfo.url3,
                markdown3: `[${taskModuleInfo.button3}](${taskModuleInfo.url3})`,
            };

            session.send(new builder.Message(session).addAttachment(
                renderACAttachment(cardTemplates.taskModule, cardData),
            ));
            // session.send("You said: %s", text);
        }
    }
}
