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

import * as express from "express";
import * as favicon from "serve-favicon";
import * as bodyParser from "body-parser";
import * as path from "path";
import * as logger from "winston";
import * as winston from "winston";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as config from "config";
import * as storage from "./storage";
import { TeamsBot } from "./TeamsBot";
import { MessagingExtension } from "./MessagingExtension";

// initLogger();
winston.verbose("hello world");

// initLogger();
winston.verbose("hello world");

let app = express();
app.set("port", process.env.PORT || 3333);
app.use(express.static(path.join(__dirname, "../../public")));
app.use(favicon(path.join(__dirname, "../../public/images", "favicon.ico")));
app.use(bodyParser.json());

// Configure bot storage
let botStorageProvider = config.get("storage");
let botStorage = null;
switch (botStorageProvider) {
    case "mongoDb":
        botStorage = new storage.MongoDbBotStorage(config.get("mongoDb.botStateCollection"), config.get("mongoDb.connectionString"));
        break;
    case "memory":
        botStorage = new builder.MemoryBotStorage();
        break;
    case "null":
        botStorage = new storage.NullBotStorage();
        break;
}

// Create bot
let connector = new msteams.TeamsChatConnector({
    appId: config.get("bot.appId"),
    appPassword: config.get("bot.appPassword"),
});
let botSettings = {
    storage: botStorage,
};
let bot = new TeamsBot(connector as builder.ChatConnector, botSettings);

// Adding a messaging extension to our bot
let messagingExtension = new MessagingExtension(bot);

// Set up route for the bot to listen.
// NOTE: This endpoint cannot be changed and must be api/messages
app.post("/api/messages", connector.listen());

// Log bot errors
bot.on("error", (error: Error) => {
    logger.error(error.message);
});

// Adding tabs to our app. This will setup routes to various views
let tabs = require("./tabs");
tabs.setup(app);

// Configure ping route
app.get("/ping", (req, res) => {
    res.status(200).send("OK");
});

// Start our nodejs app
app.listen(app.get("port"), function(): void {
    console.log("Express server listening on port " + app.get("port"));
    console.log("Bot messaging endpoint: " + config.get("app.baseUri") + "/api/messages");

    logger.verbose("Express server listening on port " + app.get("port"));
    logger.verbose("Bot messaging endpoint: " + config.get("app.baseUri") + "/api/messages");
});

function initLogger(): void {

    logger.addColors({
        error: "red",
        warn:  "yellow",
        info:  "green",
        verbose: "cyan",
        debug: "blue",
        silly: "magenta",
    });

    logger.remove(logger.transports.Console);
    logger.add(logger.transports.Console,
        {
            timestamp: () => { return new Date().toLocaleTimeString(); },
            colorize: (process.env.MONOCHROME_CONSOLE) ? false : true,
            prettyPrint: true,
            level: "debug",
        },
    );
}
