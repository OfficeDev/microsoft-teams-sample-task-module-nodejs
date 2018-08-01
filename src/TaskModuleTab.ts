import * as microsoftTeams from "@microsoft/teams-js";
import * as constants from "./constants";
import { cardTemplates } from "./dialogs/CardTemplates";

// Set the desired theme
function setTheme(theme: string): void {
    if (theme) {
        // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
        document.body.className = "theme-" + (theme === "default" ? "light" : theme);
    }
}

// Create the URL that Microsoft Teams will load in the tab. You can compose any URL even with query strings.
function createTabUrl(): string {
    let tabChoice = document.getElementById("tabChoice");
    let selectedTab = tabChoice[(tabChoice as HTMLSelectElement).selectedIndex].value;

    return window.location.protocol + "//" + window.location.host + "/" + selectedTab;
}

// Call the initialize API first
microsoftTeams.initialize();

// Check the initial theme user chose and respect it
microsoftTeams.getContext(function(context: microsoftTeams.Context): void {
    if (context && context.theme) {
        setTheme(context.theme);
    }
});

// Handle theme changes
microsoftTeams.registerOnThemeChangeHandler(function(theme: string): void {
    setTheme(theme);
});

// Save configuration changes
microsoftTeams.settings.registerOnSaveHandler(function(saveEvent: microsoftTeams.settings.SaveEvent): void {
    // Let the Microsoft Teams platform know what you want to load based on
    // what the user configured on this page
    microsoftTeams.settings.setSettings({
        contentUrl: createTabUrl(), // Mandatory parameter
        entityId: createTabUrl(), // Mandatory parameter
    });

    // Tells Microsoft Teams platform that we are done saving our settings. Microsoft Teams waits
    // for the app to call this API before it dismisses the dialog. If the wait times out, you will
    // see an error indicating that the configuration settings could not be saved.
    saveEvent.notifySuccess();
});

// Logic to let the user configure what they want to see in the tab being loaded
document.addEventListener("DOMContentLoaded", function(): void {
    let tabChoice = document.getElementById("tabChoice");
    if (tabChoice) {
        tabChoice.onchange = function(): void {
            let selectedTab = this[(this as HTMLSelectElement).selectedIndex].value;

            // This API tells Microsoft Teams to enable the 'Save' button. Since Microsoft Teams always assumes
            // an initial invalid state, without this call the 'Save' button will never be enabled.
            microsoftTeams.settings.setValidityState(selectedTab === "first" || selectedTab === "second" || selectedTab === "taskmodule");
        };
    }

    let taskModuleButtons = document.getElementsByClassName("taskModuleButton");

    // Initialize deep links
    let appRoot = `${window.location.protocol}//${window.location.host}/`;
    let taskInfo = {
        appId: "bdc707d5-48e0-48f8-bbe7-6131e0565a4c",
        title: null,
        height: null,
        width: null,
        url: null,
        card: null,
    };
    let deepLink = document.getElementById("dlYouTube") as HTMLAnchorElement;
    deepLink.href = encodeURI(`https://teams.microsoft.com/l/task/${taskInfo.appId}?url=${appRoot}youtube&height=${taskInfo.height}&width=${taskInfo.width}&title=${encodeURIComponent("Satya Nadella's Build 2018 Keynote")}`);
    deepLink = document.getElementById("dlPowerApps") as HTMLAnchorElement;
    deepLink.href = encodeURI(`https://teams.microsoft.com/l/task/${taskInfo.appId}?url=${appRoot}powerapps&height=${taskInfo.height}&width=${taskInfo.width}&title=${encodeURIComponent("PowerApp: Asset Checkout")}`);
    deepLink = document.getElementById("dlCustomForm") as HTMLAnchorElement;
    deepLink.href = encodeURI(`https://teams.microsoft.com/l/task/${taskInfo.appId}?url=${appRoot}customform&height=medium&width=medium&title=${encodeURIComponent("Custom Form")}`);
    deepLink = document.getElementById("dlAdaptiveCard") as HTMLAnchorElement;
    deepLink.href = encodeURI(`https://teams.microsoft.com/l/task/${taskInfo.appId}?height=large&width=medium&card=${encodeURIComponent(cardTemplates.adaptiveCard)}`);

    for (let btn of taskModuleButtons) {
        btn.addEventListener("click",
            function (): void {
                taskInfo.url = appRoot + this.id.toLowerCase();
                let completionHandler = (err: string, result: any): void => { console.log("Result: " + result); };
                switch (this.id.toLowerCase()) {
                    case constants.TaskModuleIds.YouTube:
                        taskInfo.title = constants.TaskModuleStrings.YouTubeTitle;
                        taskInfo.height = "large";
                        taskInfo.width = "large";
                        microsoftTeams.tasks.startTask(taskInfo, completionHandler);
                        break;
                    case constants.TaskModuleIds.PowerApp:
                        taskInfo.title = constants.TaskModuleStrings.PowerAppTitle;
                        taskInfo.height = "large";
                        taskInfo.width = "large";
                        microsoftTeams.tasks.startTask(taskInfo, completionHandler);
                        break;
                    case constants.TaskModuleIds.CustomForm:
                        taskInfo.title = constants.TaskModuleStrings.CustomFormTitle;
                        taskInfo.height = "medium";
                        taskInfo.width = "medium";
                        microsoftTeams.tasks.startTask(taskInfo, completionHandler);
                        break;
                    case constants.TaskModuleIds.AdaptiveCard:
                        taskInfo.title = constants.TaskModuleStrings.AdaptiveCardTitle;
                        taskInfo.url = null;
                        taskInfo.height = "large";
                        taskInfo.width = "medium";
                        taskInfo.card = cardTemplates.adaptiveCard;
                        microsoftTeams.tasks.startTask(taskInfo, completionHandler);
                        break;
                    default:
                        console.log("Unexpected button ID");
                        return;
                }
                console.log("URL: " + taskInfo.url);
            });
    }
});
