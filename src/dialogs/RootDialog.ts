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
import { ACGeneratorDialog } from "./ACGenerator";
import { renderACAttachment } from "../utils/AdaptiveCardUtils";
import { cardTemplates, fetchTemplates } from "./CardTemplates";

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
        new ACGeneratorDialog(constants.DialogId.ACTester).register(bot, this);
        this.matches(/actester/i, constants.DialogId.ACTester);
    }

    // Handle start of dialog
    private async onDialogBegin(session: builder.Session, args: any, next: () => void): Promise<void> {
        next();
    }

    // Handle message
    private async onMessageReceived(session: builder.Session): Promise<void> {
        if (session.message.text === "") {
            console.log("Empty message received");
            // This is a response from a generated AC card
            if (session.message.value !== undefined) {
                session.send("**Action.Submit results:** " + JSON.stringify(session.message.value));
            }
        }
        else {
            // Message might contain @mentions which we would like to strip off in the response
            let text = utils.getTextWithoutMentions(session.message);

            let appInfo = {
                appId: config.get("bot.appId"),
                appRoot: config.get("app.baseUri"),
            };
            let taskModuleUrls = {
                url1: encodeURI(`https://teams.microsoft.com/l/task/${appInfo.appId}?url=${appInfo.appRoot}/youtube&height=large&width=large&title=${encodeURIComponent("Satya Nadella's Build 2018 Keynote")}`),
                url2: encodeURI(`https://teams.microsoft.com/l/task/${appInfo.appId}?url=${appInfo.appRoot}/powerapps&height=large&width=large&title=${encodeURIComponent("PowerApp: Asset Checkout")}`),
                url3: encodeURI(`https://teams.microsoft.com/l/task/${appInfo.appId}?url=${appInfo.appRoot}/customform&height=medium&width=medium&title=${encodeURIComponent("Custom Form")}`),
            };

            let cardData: any = {
                title: "Task Module",
                subTitle: "Task Module Test Card",
                instructions: "Click on the buttons below below to open task modules in various ways.",
                linkbutton1: constants.TaskModuleStrings.YouTubeName,
                url1: taskModuleUrls.url1,
                markdown1: `[${constants.TaskModuleStrings.YouTubeName}](${taskModuleUrls.url1})`,
                linkbutton2: constants.TaskModuleStrings.PowerAppName,
                url2: taskModuleUrls.url2,
                markdown2: `[${constants.TaskModuleStrings.PowerAppName}](${taskModuleUrls.url2})`,
                linkbutton3: constants.TaskModuleStrings.CustomFormName,
                url3: taskModuleUrls.url3,
                markdown3: `[${constants.TaskModuleStrings.CustomFormName}](${taskModuleUrls.url3})`,
                fetchButtonId1: `${constants.TaskModuleIds.YouTube}`,
                fetchButtonTitle1: `${constants.TaskModuleStrings.YouTubeName}`,
                fetchButtonId2: `${constants.TaskModuleIds.PowerApp}`,
                fetchButtonTitle2: `${constants.TaskModuleStrings.PowerAppName}`,
                fetchButtonId3: `${constants.TaskModuleIds.CustomForm}`,
                fetchButtonTitle3: `${constants.TaskModuleStrings.CustomFormName}`,
                taskFetchJSON: `**${constants.TaskModuleStrings.YouTubeName}:** ${JSON.stringify(fetchTemplates[constants.TaskModuleIds.YouTube])}\r**${constants.TaskModuleStrings.PowerAppName}:** ${JSON.stringify(fetchTemplates[constants.TaskModuleIds.PowerApp])}\r**${constants.TaskModuleStrings.CustomFormName}:** ${JSON.stringify(fetchTemplates[constants.TaskModuleIds.CustomForm])}`,
            };

            session.send(new builder.Message(session).addAttachment(
                renderACAttachment(cardTemplates.taskModule, cardData),
            ));
            // session.send("You said: %s", text);
        }
    }
}
