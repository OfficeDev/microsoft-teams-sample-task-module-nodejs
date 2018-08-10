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

// Activity types
export const messageType = "message";
export const invokeType = "invoke";

// Dialog ids
// tslint:disable-next-line:variable-name
export const DialogId = {
    Root: "/",
    ACTester: "actester",
};

// Telemetry events
// tslint:disable-next-line:variable-name
export const TelemetryEvent = {
    UserActivity: "UserActivity",
    BotActivity: "BotActivity",
};

// URL Placeholders
// tslint:disable-next-line:variable-name
export const UrlPlaceholders = "loginHint={loginHint}&upn={userPrincipalName}&aadId={userObjectId}&theme={theme}&groupId={groupId}&tenantId={tid}&locale={locale}";

// Task Module Strings
// tslint:disable-next-line:variable-name
export const TaskModuleStrings = {
    YouTubeTitle: "Satya Nadella's Build 2018 Keynote",
    PowerAppTitle: "PowerApp: Asset Checkout",
    CustomFormTitle: "Custom Form",
    AdaptiveCardTitle: "Adaptive Card: Inputs",
    ActionSubmitResponseTitle: "Action.Submit Response",
    YouTubeName: "YouTube",
    PowerAppName: "PowerApp",
    CustomFormName: "Custom Form",
    AdaptiveCardName: "Adaptive Card",
};

// Task Module Ids
// tslint:disable-next-line:variable-name
export const TaskModuleIds = {
    YouTube: "youtube",
    PowerApp: "powerapp",
    CustomForm: "customform",
    AdaptiveCard1: "adaptivecard1",
    AdaptiveCard2: "adaptivecard2",
};

// Task Module Sizes
// tslint:disable-next-line:variable-name
export const TaskModuleSizes = {
    youtube: {
        height: 80,
        width: 80,
    },
    powerapp: {
        height: 60,
        width: 70,
    },
    customform: {
        height: 40,
        width: 30,
    },
    // youtube: {
    //     height: 20,
    //     width: 20,
    // },
    // powerapp: {
    //     height: "medium",
    //     width: "medium",
    // },
    // customform: {
    //     height: "80",
    //     width: "medium",
    // },
    adaptivecard: {
        height: "large",
        width: "medium",
    },
};
