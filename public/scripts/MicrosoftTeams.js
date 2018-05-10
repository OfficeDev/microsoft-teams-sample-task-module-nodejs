/**
 * This is the root namespace for the JavaScript SDK.
 */
var microsoftTeams;
(function (microsoftTeams) {
    "use strict";
    const version = "1.2";
    const validOrigins = [
        "https://teams.microsoft.com",
        "https://teams.microsoft.us",
        "https://int.teams.microsoft.com",
        "https://devspaces.skype.com",
        "https://ssauth.skype.com",
        "http://dev.local" // local development
    ];
    const handlers = {};
    // Ensure these declarations stay in sync with the framework.
    const frameContexts = {
        settings: "settings",
        content: "content",
        authentication: "authentication",
        remove: "remove"
    };
    const hostClientTypes = {
        desktop: "desktop",
        web: "web"
    };
    let TeamType;
    (function (TeamType) {
        TeamType[TeamType["Standard"] = 0] = "Standard";
        TeamType[TeamType["Edu"] = 1] = "Edu";
        TeamType[TeamType["Class"] = 2] = "Class";
        TeamType[TeamType["Plc"] = 3] = "Plc";
        TeamType[TeamType["Staff"] = 4] = "Staff";
    })(TeamType = microsoftTeams.TeamType || (microsoftTeams.TeamType = {}));
    let UserTeamRole;
    (function (UserTeamRole) {
        UserTeamRole[UserTeamRole["Admin"] = 0] = "Admin";
        UserTeamRole[UserTeamRole["User"] = 1] = "User";
        UserTeamRole[UserTeamRole["Guest"] = 2] = "Guest";
    })(UserTeamRole = microsoftTeams.UserTeamRole || (microsoftTeams.UserTeamRole = {}));
    // This indicates whether initialize was called (started).
    // It does not indicate whether initialization is complete. That can be inferred by whether parentOrigin is set.
    let initializeCalled = false;
    let currentWindow;
    let parentWindow;
    let parentOrigin;
    let parentMessageQueue = [];
    let childWindow;
    let childOrigin;
    let childMessageQueue = [];
    let nextMessageId = 0;
    let callbacks = {};
    let frameContext;
    let hostClientType;
    let themeChangeHandler;
    handlers["themeChange"] = handleThemeChange;
    let fullScreenChangeHandler;
    handlers["fullScreenChange"] = handleFullScreenChange;
    let backButtonPressHandler;
    handlers["backButtonPress"] = handleBackButtonPress;
    /**
     * Initializes the library. This must be called before any other SDK calls
     * but after the frame is loaded successfully.
     */
    function initialize() {
        if (initializeCalled) {
            // Independent components might not know whether the SDK is initialized so might call it to be safe.
            // Just no-op if that happens to make it easier to use.
            return;
        }
        initializeCalled = true;
        // Undocumented field used to mock the window for unit tests
        currentWindow = this._window || window;
        // Listen for messages post to our window
        let messageListener = (evt) => processMessage(evt);
        currentWindow.addEventListener("message", messageListener, false);
        // If we are in an iframe, our parent window is the one hosting us (i.e., window.parent); otherwise,
        // it's the window that opened us (i.e., window.opener)
        parentWindow =
            currentWindow.parent !== currentWindow.self
                ? currentWindow.parent
                : currentWindow.opener;
        try {
            // Send the initialized message to any origin, because at this point we most likely don't know the origin
            // of the parent window, and this message contains no data that could pose a security risk.
            parentOrigin = "*";
            let messageId = sendMessageRequest(parentWindow, "initialize", [version]);
            callbacks[messageId] = (context, clientType) => {
                frameContext = context;
                hostClientType = clientType;
            };
        }
        finally {
            parentOrigin = null;
        }
        // Undocumented function used to clear state between unit tests
        this._uninitialize = () => {
            if (frameContext) {
                registerOnThemeChangeHandler(null);
                registerFullScreenHandler(null);
                registerBackButtonHandler(null);
            }
            if (frameContext === frameContexts.settings) {
                settings.registerOnSaveHandler(null);
            }
            if (frameContext === frameContexts.remove) {
                settings.registerOnRemoveHandler(null);
            }
            initializeCalled = false;
            parentWindow = null;
            parentOrigin = null;
            parentMessageQueue = [];
            childWindow = null;
            childOrigin = null;
            childMessageQueue = [];
            nextMessageId = 0;
            callbacks = {};
            frameContext = null;
            hostClientType = null;
            currentWindow.removeEventListener("message", messageListener, false);
        };
    }
    microsoftTeams.initialize = initialize;
    /**
     * Retrieves the current context the frame is running in.
     * @param callback The callback to invoke when the {@link Context} object is retrieved.
     */
    function getContext(callback) {
        ensureInitialized();
        let messageId = sendMessageRequest(parentWindow, "getContext");
        callbacks[messageId] = callback;
    }
    microsoftTeams.getContext = getContext;
    /**
     * Registers a handler for theme changes.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     * @param handler The handler to invoke when the user changes their theme.
     */
    function registerOnThemeChangeHandler(handler) {
        ensureInitialized();
        themeChangeHandler = handler;
    }
    microsoftTeams.registerOnThemeChangeHandler = registerOnThemeChangeHandler;
    function handleThemeChange(theme) {
        if (themeChangeHandler) {
            themeChangeHandler(theme);
        }
        if (childWindow) {
            sendMessageRequest(childWindow, "themeChange", [theme]);
        }
    }
    /**
     * Registers a handler for changes from or to full-screen view for a tab.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     * @param handler The handler to invoke when the user toggles full-screen view for a tab.
     */
    function registerFullScreenHandler(handler) {
        ensureInitialized();
        fullScreenChangeHandler = handler;
    }
    microsoftTeams.registerFullScreenHandler = registerFullScreenHandler;
    function handleFullScreenChange(isFullScreen) {
        if (fullScreenChangeHandler) {
            fullScreenChangeHandler(isFullScreen);
        }
    }
    /**
     * Registers a handler for user presses of the Team client's back button. Experiences that maintain an internal
     * navigation stack should use this handler to navigate the user back within their frame. If an app finds
     * that after running its back button handler it cannot handle the event it should call the navigateBack
     * method to ask the Teams client to handle it instead.
     * @param handler The handler to invoke when the user presses their Team client's back button.
     */
    function registerBackButtonHandler(handler) {
        ensureInitialized();
        backButtonPressHandler = handler;
    }
    microsoftTeams.registerBackButtonHandler = registerBackButtonHandler;
    function handleBackButtonPress() {
        if (!backButtonPressHandler || !backButtonPressHandler()) {
            navigateBack();
        }
    }
    /**
     * Navigates back in the Teams client. See registerBackButtonHandler for more information on when
     * it's appropriate to use this method.
     */
    function navigateBack() {
        ensureInitialized();
        let messageId = sendMessageRequest(parentWindow, "navigateBack", []);
        callbacks[messageId] = (success) => {
            if (!success) {
                throw new Error("Back navigation is not supported in the current client or context.");
            }
        };
    }
    microsoftTeams.navigateBack = navigateBack;
    /**
     * Navigates the frame to a new cross-domain URL. The domain of this URL must match at least one of the
     * valid domains specified in the validDomains block of the manifest; otherwise, an exception will be
     * thrown. This function needs to be used only when navigating the frame to a URL in a different domain
     * than the current one in a way that keeps the app informed of the change and allows the SDK to
     * continue working.
     * @param url The URL to navigate the frame to.
     */
    function navigateCrossDomain(url) {
        ensureInitialized(frameContexts.content, frameContexts.settings, frameContexts.remove);
        let messageId = sendMessageRequest(parentWindow, "navigateCrossDomain", [
            url
        ]);
        callbacks[messageId] = (success) => {
            if (!success) {
                throw new Error("Cross-origin navigation is only supported for URLs matching the pattern registered in the manifest.");
            }
        };
    }
    microsoftTeams.navigateCrossDomain = navigateCrossDomain;
    /**
     * Allows an app to retrieve for this user tabs that are owned by this app.
     * If no TabInstanceParameters are passed, the app defaults to favorite teams and favorite channels.
     * @param callback The callback to invoke when the {@link TabInstanceParameters} object is retrieved.
     * @param tabInstanceParameters OPTIONAL Flags that specify whether to scope call to favorite teams or channels.
     */
    function getTabInstances(callback, tabInstanceParameters) {
        ensureInitialized();
        let messageId = sendMessageRequest(parentWindow, "getTabInstances", [
            tabInstanceParameters
        ]);
        callbacks[messageId] = callback;
    }
    microsoftTeams.getTabInstances = getTabInstances;
    /**
     * Allows an app to retrieve the most recently used tabs for this user.
     * @param callback The callback to invoke when the {@link TabInformation} object is retrieved.
     * @param tabInstanceParameters OPTIONAL Ignored, kept for future use
     */
    function getMruTabInstances(callback, tabInstanceParameters) {
        ensureInitialized();
        let messageId = sendMessageRequest(parentWindow, "getMruTabInstances", [
            tabInstanceParameters
        ]);
        callbacks[messageId] = callback;
    }
    microsoftTeams.getMruTabInstances = getMruTabInstances;
    /**
     * Shares a deep link that a user can use to navigate back to a specific state in this page.
     * @param deepLinkParameters ID and label for the link and fallback URL.
     */
    function shareDeepLink(deepLinkParameters) {
        ensureInitialized(frameContexts.content);
        sendMessageRequest(parentWindow, "shareDeepLink", [
            deepLinkParameters.subEntityId,
            deepLinkParameters.subEntityLabel,
            deepLinkParameters.subEntityWebUrl
        ]);
    }
    microsoftTeams.shareDeepLink = shareDeepLink;
    /**
     * Opens a client-friendly preview of the specified file.
     * @param file The file to preview.
     */
    function openFilePreview(filePreviewParameters) {
        ensureInitialized(frameContexts.content);
        sendMessageRequest(parentWindow, "openFilePreview", [
            filePreviewParameters.entityId,
            filePreviewParameters.title,
            filePreviewParameters.description,
            filePreviewParameters.type,
            filePreviewParameters.objectUrl,
            filePreviewParameters.downloadUrl,
            filePreviewParameters.webPreviewUrl,
            filePreviewParameters.webEditUrl
        ]);
    }
    microsoftTeams.openFilePreview = openFilePreview;
    /**
     * Navigates the Microsoft Teams app to the specified tab instance.
     * @param tabInstance The tab instance to navigate to.
     */
    function navigateToTab(tabInstance) {
        ensureInitialized();
        let messageId = sendMessageRequest(parentWindow, "navigateToTab", [
            tabInstance
        ]);
        callbacks[messageId] = (success) => {
            if (!success) {
                throw new Error("Invalid internalTabInstanceId and/or channelId were/was provided");
            }
        };
    }
    microsoftTeams.navigateToTab = navigateToTab;
    /**
     * Namespace to interact with the settings-specific part of the SDK.
     * This object is usable only on the settings frame.
     */
    let settings;
    (function (settings) {
        let saveHandler;
        let removeHandler;
        handlers["settings.save"] = handleSave;
        handlers["settings.remove"] = handleRemove;
        /**
         * Sets the validity state for the settings.
         * The initial value is false, so the user cannot save the settings until this is called with true.
         * @param validityState Indicates whether the save or remove button is enabled for the user.
         */
        function setValidityState(validityState) {
            ensureInitialized(frameContexts.settings, frameContexts.remove);
            sendMessageRequest(parentWindow, "settings.setValidityState", [
                validityState
            ]);
        }
        settings.setValidityState = setValidityState;
        /**
         * Gets the settings for the current instance.
         * @param callback The callback to invoke when the {@link Settings} object is retrieved.
         */
        function getSettings(callback) {
            ensureInitialized(frameContexts.settings, frameContexts.remove);
            let messageId = sendMessageRequest(parentWindow, "settings.getSettings");
            callbacks[messageId] = callback;
        }
        settings.getSettings = getSettings;
        /**
         * Sets the settings for the current instance.
         * This is an asynchronous operation; calls to getSettings are not guaranteed to reflect the changed state.
         * @param settings The desired settings for this instance.
         */
        function setSettings(instanceSettings) {
            ensureInitialized(frameContexts.settings);
            sendMessageRequest(parentWindow, "settings.setSettings", [
                instanceSettings
            ]);
        }
        settings.setSettings = setSettings;
        /**
         * Registers a handler for when the user attempts to save the settings. This handler should be used
         * to create or update the underlying resource powering the content.
         * The object passed to the handler must be used to notify whether to proceed with the save.
         * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
         * @param handler The handler to invoke when the user selects the save button.
         */
        function registerOnSaveHandler(handler) {
            ensureInitialized(frameContexts.settings);
            saveHandler = handler;
        }
        settings.registerOnSaveHandler = registerOnSaveHandler;
        /**
         * Registers a handler for user attempts to remove content. This handler should be used
         * to remove the underlying resource powering the content.
         * The object passed to the handler must be used to indicate whether to proceed with the removal.
         * Only one handler may be registered at a time. Subsequent registrations will override the first.
         * @param handler The handler to invoke when the user selects the remove button.
         */
        function registerOnRemoveHandler(handler) {
            ensureInitialized(frameContexts.remove);
            removeHandler = handler;
        }
        settings.registerOnRemoveHandler = registerOnRemoveHandler;
        function handleSave(result) {
            let saveEvent = new SaveEventImpl(result);
            if (saveHandler) {
                saveHandler(saveEvent);
            }
            else {
                // If no handler is registered, we assume success.
                saveEvent.notifySuccess();
            }
        }
        /**
         * @private
         * Hide from docs, since this class is not directly used.
         */
        class SaveEventImpl {
            constructor(result) {
                this.notified = false;
                this.result = result ? result : {};
            }
            notifySuccess() {
                this.ensureNotNotified();
                sendMessageRequest(parentWindow, "settings.save.success");
                this.notified = true;
            }
            notifyFailure(reason) {
                this.ensureNotNotified();
                sendMessageRequest(parentWindow, "settings.save.failure", [reason]);
                this.notified = true;
            }
            ensureNotNotified() {
                if (this.notified) {
                    throw new Error("The SaveEvent may only notify success or failure once.");
                }
            }
        }
        function handleRemove() {
            let removeEvent = new RemoveEventImpl();
            if (removeHandler) {
                removeHandler(removeEvent);
            }
            else {
                // If no handler is registered, we assume success.
                removeEvent.notifySuccess();
            }
        }
        /**
         * @private
         * Hide from docs, since this class is not directly used.
         */
        class RemoveEventImpl {
            constructor() {
                this.notified = false;
            }
            notifySuccess() {
                this.ensureNotNotified();
                sendMessageRequest(parentWindow, "settings.remove.success");
                this.notified = true;
            }
            notifyFailure(reason) {
                this.ensureNotNotified();
                sendMessageRequest(parentWindow, "settings.remove.failure", [reason]);
                this.notified = true;
            }
            ensureNotNotified() {
                if (this.notified) {
                    throw new Error("The removeEvent may only notify success or failure once.");
                }
            }
        }
    })(settings = microsoftTeams.settings || (microsoftTeams.settings = {}));
    /**
     * Namespace to interact with the authentication-specific part of the SDK.
     * This object is used for starting or completing authentication flows.
     */
    let authentication;
    (function (authentication) {
        let authParams;
        let authWindowMonitor;
        handlers["authentication.authenticate.success"] = handleSuccess;
        handlers["authentication.authenticate.failure"] = handleFailure;
        /**
         * Registers the authentication handlers
         * @param authenticateParameters A set of values that configure the authentication pop-up.
         */
        function registerAuthenticationHandlers(authenticateParameters) {
            authParams = authenticateParameters;
        }
        authentication.registerAuthenticationHandlers = registerAuthenticationHandlers;
        /**
         * Initiates an authentication request, which opens a new window with the specified settings.
         */
        function authenticate(authenticateParameters) {
            let authenticateParams = authenticateParameters !== undefined
                ? authenticateParameters
                : authParams;
            ensureInitialized(frameContexts.content, frameContexts.settings, frameContexts.remove);
            if (hostClientType === hostClientTypes.desktop) {
                // Convert any relative URLs into absolute URLs before sending them over to the parent window.
                let link = document.createElement("a");
                link.href = authenticateParams.url;
                // Ask the parent window to open an authentication window with the parameters provided by the caller.
                let messageId = sendMessageRequest(parentWindow, "authentication.authenticate", [link.href, authenticateParams.width, authenticateParams.height]);
                callbacks[messageId] = (success, response) => {
                    if (success) {
                        authenticateParams.successCallback(response);
                    }
                    else {
                        authenticateParams.failureCallback(response);
                    }
                };
            }
            else {
                // Open an authentication window with the parameters provided by the caller.
                openAuthenticationWindow(authenticateParams);
            }
        }
        authentication.authenticate = authenticate;
        /**
         * @private
         * Hide from docs.
         * ------
         * Requests an Azure AD token to be issued on behalf of the app. The token is acquired from the cache
         * if it is not expired. Otherwise a request is sent to Azure AD to obtain a new token.
         * @param authTokenRequest A set of values that configure the token request.
         */
        function getAuthToken(authTokenRequest) {
            ensureInitialized();
            let messageId = sendMessageRequest(parentWindow, "authentication.getAuthToken", [authTokenRequest.resources]);
            callbacks[messageId] = (success, result) => {
                if (success) {
                    authTokenRequest.successCallback(result);
                }
                else {
                    authTokenRequest.failureCallback(result);
                }
            };
        }
        authentication.getAuthToken = getAuthToken;
        /**
         * @private
         * Hide from docs.
         * ------
         * Requests the decoded Azure AD user identity on behalf of the app.
         */
        function getUser(userRequest) {
            ensureInitialized();
            let messageId = sendMessageRequest(parentWindow, "authentication.getUser");
            callbacks[messageId] = (success, result) => {
                if (success) {
                    userRequest.successCallback(result);
                }
                else {
                    userRequest.failureCallback(result);
                }
            };
        }
        authentication.getUser = getUser;
        function closeAuthenticationWindow() {
            // Stop monitoring the authentication window
            stopAuthenticationWindowMonitor();
            // Try to close the authentication window and clear all properties associated with it
            try {
                if (childWindow) {
                    childWindow.close();
                }
            }
            finally {
                childWindow = null;
                childOrigin = null;
            }
        }
        function openAuthenticationWindow(authenticateParameters) {
            authParams = authenticateParameters;
            // Close the previously opened window if we have one
            closeAuthenticationWindow();
            // Start with a sensible default size
            let width = authParams.width || 600;
            let height = authParams.height || 400;
            // Ensure that the new window is always smaller than our app's window so that it never fully covers up our app
            width = Math.min(width, currentWindow.outerWidth - 400);
            height = Math.min(height, currentWindow.outerHeight - 200);
            // Convert any relative URLs into absolute URLs before sending them over to the parent window
            let link = document.createElement("a");
            link.href = authParams.url;
            // We are running in the browser, so we need to center the new window ourselves
            let left = typeof currentWindow.screenLeft !== "undefined"
                ? currentWindow.screenLeft
                : currentWindow.screenX;
            let top = typeof currentWindow.screenTop !== "undefined"
                ? currentWindow.screenTop
                : currentWindow.screenY;
            left += currentWindow.outerWidth / 2 - width / 2;
            top += currentWindow.outerHeight / 2 - height / 2;
            // Open a child window with a desired set of standard browser features
            childWindow = currentWindow.open(link.href, "_blank", "toolbar=no, location=yes, status=no, menubar=no, scrollbars=yes, top=" +
                top +
                ", left=" +
                left +
                ", width=" +
                width +
                ", height=" +
                height);
            if (childWindow) {
                // Start monitoring the authentication window so that we can detect if it gets closed before the flow completes
                startAuthenticationWindowMonitor();
            }
            else {
                // If we failed to open the window, fail the authentication flow
                handleFailure("FailedToOpenWindow");
            }
        }
        function stopAuthenticationWindowMonitor() {
            if (authWindowMonitor) {
                clearInterval(authWindowMonitor);
                authWindowMonitor = 0;
            }
            delete handlers["initialize"];
            delete handlers["navigateCrossDomain"];
        }
        function startAuthenticationWindowMonitor() {
            // Stop the previous window monitor if one is running
            stopAuthenticationWindowMonitor();
            // Create an interval loop that
            // - Notifies the caller of failure if it detects that the authentication window is closed
            // - Keeps pinging the authentication window while it is open to re-establish
            //   contact with any pages along the authentication flow that need to communicate
            //   with us
            authWindowMonitor = currentWindow.setInterval(() => {
                if (!childWindow || childWindow.closed) {
                    handleFailure("CancelledByUser");
                }
                else {
                    let savedChildOrigin = childOrigin;
                    try {
                        childOrigin = "*";
                        sendMessageRequest(childWindow, "ping");
                    }
                    finally {
                        childOrigin = savedChildOrigin;
                    }
                }
            }, 100);
            // Set up an initialize-message handler that gives the authentication window its frame context
            handlers["initialize"] = () => {
                return [frameContexts.authentication, hostClientType];
            };
            // Set up a navigateCrossDomain message handler that blocks cross-domain re-navigation attempts
            // in the authentication window. We could at some point choose to implement this method via a call to
            // authenticationWindow.location.href = url; however, we would first need to figure out how to
            // validate the URL against the tab's list of valid domains.
            handlers["navigateCrossDomain"] = (url) => {
                return false;
            };
        }
        /**
         * Notifies the frame that initiated this authentication request that the request was successful.
         * This function is usable only on the authentication window.
         * This call causes the authentication window to be closed.
         * @param result Specifies a result for the authentication. If specified, the frame that initiated the authentication pop-up receives this value in its callback.
         * @param callbackUrl Specifies the url to redirect back to if the client is Win32 Outlook.
         */
        function notifySuccess(result, callbackUrl) {
            redirectIfWin32Outlook(callbackUrl, "result", result);
            ensureInitialized(frameContexts.authentication);
            sendMessageRequest(parentWindow, "authentication.authenticate.success", [
                result
            ]);
            // Wait for the message to be sent before closing the window
            waitForMessageQueue(parentWindow, () => setTimeout(() => currentWindow.close(), 200));
        }
        authentication.notifySuccess = notifySuccess;
        /**
         * Notifies the frame that initiated this authentication request that the request failed.
         * This function is usable only on the authentication window.
         * This call causes the authentication window to be closed.
         * @param result Specifies a result for the authentication. If specified, the frame that initiated the authentication pop-up receives this value in its callback.
         * @param callbackUrl Specifies the url to redirect back to if the client is Win32 Outlook.
         */
        function notifyFailure(reason, callbackUrl) {
            redirectIfWin32Outlook(callbackUrl, "reason", reason);
            ensureInitialized(frameContexts.authentication);
            sendMessageRequest(parentWindow, "authentication.authenticate.failure", [
                reason
            ]);
            // Wait for the message to be sent before closing the window
            waitForMessageQueue(parentWindow, () => setTimeout(() => currentWindow.close(), 200));
        }
        authentication.notifyFailure = notifyFailure;
        function handleSuccess(result) {
            try {
                if (authParams && authParams.successCallback) {
                    authParams.successCallback(result);
                }
            }
            finally {
                authParams = null;
                closeAuthenticationWindow();
            }
        }
        function handleFailure(reason) {
            try {
                if (authParams && authParams.failureCallback) {
                    authParams.failureCallback(reason);
                }
            }
            finally {
                authParams = null;
                closeAuthenticationWindow();
            }
        }
        /**
         * Validates that the callbackUrl param is a valid connector url, appends the result/reason and authSuccess/authFailure as URL fragments and redirects the window
         * @param callbackUrl - the connectors url to redirect to
         * @param key - "result" in case of success and "reason" in case of failure
         * @param value - the value of the passed result/reason parameter
         */
        function redirectIfWin32Outlook(callbackUrl, key, value) {
            if (callbackUrl) {
                let link = document.createElement("a");
                link.href = decodeURIComponent(callbackUrl);
                if (link.host &&
                    link.host !== window.location.host &&
                    link.host === "outlook.office.com" &&
                    link.search.indexOf("client_type=Win32_Outlook") > -1) {
                    if (key && key === "result") {
                        if (value) {
                            link.href = updateUrlParameter(link.href, "result", value);
                        }
                        currentWindow.location.assign(updateUrlParameter(link.href, "authSuccess", ""));
                    }
                    if (key && key === "reason") {
                        if (value) {
                            link.href = updateUrlParameter(link.href, "reason", value);
                        }
                        currentWindow.location.assign(updateUrlParameter(link.href, "authFailure", ""));
                    }
                }
            }
        }
        /**
         * Appends either result or reason as a fragment to the 'callbackUrl'
         * @param uri - the url to modify
         * @param key - the fragment key
         * @param value - the fragment value
         */
        function updateUrlParameter(uri, key, value) {
            let i = uri.indexOf("#");
            let hash = i === -1 ? "#" : uri.substr(i);
            hash = hash + "&" + key + (value !== "" ? "=" + value : "");
            uri = i === -1 ? uri : uri.substr(0, i);
            return uri + hash;
        }
    })(authentication = microsoftTeams.authentication || (microsoftTeams.authentication = {}));
    function ensureInitialized(...expectedFrameContexts) {
        if (!initializeCalled) {
            throw new Error("The library has not yet been initialized");
        }
        if (frameContext &&
            expectedFrameContexts &&
            expectedFrameContexts.length > 0) {
            let found = false;
            for (let i = 0; i < expectedFrameContexts.length; i++) {
                if (expectedFrameContexts[i] === frameContext) {
                    found = true;
                    break;
                }
            }
            if (!found) {
                throw new Error("This call is not allowed in the '" + frameContext + "' context");
            }
        }
    }
    function processMessage(evt) {
        // Process only if we received a valid message
        if (!evt || !evt.data || typeof evt.data !== "object") {
            return;
        }
        // Process only if the message is coming from a different window and a valid origin
        let messageSource = evt.source || evt.originalEvent.source;
        let messageOrigin = evt.origin || evt.originalEvent.origin;
        if (messageSource === currentWindow ||
            (messageOrigin !== currentWindow.location.origin &&
                validOrigins.indexOf(messageOrigin.toLowerCase()) === -1)) {
            return;
        }
        // Update our parent and child relationships based on this message
        updateRelationships(messageSource, messageOrigin);
        // Handle the message
        if (messageSource === parentWindow) {
            handleParentMessage(evt);
        }
        else if (messageSource === childWindow) {
            handleChildMessage(evt);
        }
    }
    function updateRelationships(messageSource, messageOrigin) {
        // Determine whether the source of the message is our parent or child and update our
        // window and origin pointer accordingly
        if (!parentWindow || messageSource === parentWindow) {
            parentWindow = messageSource;
            parentOrigin = messageOrigin;
        }
        else if (!childWindow || messageSource === childWindow) {
            childWindow = messageSource;
            childOrigin = messageOrigin;
        }
        // Clean up pointers to closed parent and child windows
        if (parentWindow && parentWindow.closed) {
            parentWindow = null;
            parentOrigin = null;
        }
        if (childWindow && childWindow.closed) {
            childWindow = null;
            childOrigin = null;
        }
        // If we have any messages in our queue, send them now
        flushMessageQueue(parentWindow);
        flushMessageQueue(childWindow);
    }
    function handleParentMessage(evt) {
        if ("id" in evt.data) {
            // Call any associated callbacks
            const message = evt.data;
            const callback = callbacks[message.id];
            if (callback) {
                callback.apply(null, message.args);
                // Remove the callback to ensure that the callback is called only once and to free up memory.
                delete callbacks[message.id];
            }
        }
        else if ("func" in evt.data) {
            // Delegate the request to the proper handler
            const message = evt.data;
            const handler = handlers[message.func];
            if (handler) {
                // We don't expect any handler to respond at this point
                handler.apply(this, message.args);
            }
        }
    }
    function handleChildMessage(evt) {
        if ("id" in evt.data && "func" in evt.data) {
            // Try to delegate the request to the proper handler
            const message = evt.data;
            const handler = handlers[message.func];
            if (handler) {
                let result = handler.apply(this, message.args);
                if (result) {
                    sendMessageResponse(childWindow, message.id, Array.isArray(result) ? result : [result]);
                }
            }
            else {
                // Proxy to parent
                let messageId = sendMessageRequest(parentWindow, message.func, message.args);
                // tslint:disable-next-line:no-any
                callbacks[messageId] = (...args) => {
                    if (childWindow) {
                        sendMessageResponse(childWindow, message.id, args);
                    }
                };
            }
        }
    }
    function getTargetMessageQueue(targetWindow) {
        return targetWindow === parentWindow
            ? parentMessageQueue
            : targetWindow === childWindow
                ? childMessageQueue
                : [];
    }
    function getTargetOrigin(targetWindow) {
        return targetWindow === parentWindow
            ? parentOrigin
            : targetWindow === childWindow
                ? childOrigin
                : null;
    }
    function flushMessageQueue(targetWindow) {
        let targetOrigin = getTargetOrigin(targetWindow);
        let targetMessageQueue = getTargetMessageQueue(targetWindow);
        while (targetWindow && targetOrigin && targetMessageQueue.length > 0) {
            targetWindow.postMessage(targetMessageQueue.shift(), targetOrigin);
        }
    }
    function waitForMessageQueue(targetWindow, callback) {
        let messageQueueMonitor = currentWindow.setInterval(() => {
            if (getTargetMessageQueue(targetWindow).length === 0) {
                clearInterval(messageQueueMonitor);
                callback();
            }
        }, 100);
    }
    function sendMessageRequest(targetWindow, actionName, 
    // tslint:disable-next-line:no-any
    args) {
        let request = createMessageRequest(actionName, args);
        let targetOrigin = getTargetOrigin(targetWindow);
        // If the target window isn't closed and we already know its origin, send the message right away; otherwise,
        // queue the message and send it after the origin is established
        if (targetWindow && targetOrigin) {
            targetWindow.postMessage(request, targetOrigin);
        }
        else {
            getTargetMessageQueue(targetWindow).push(request);
        }
        return request.id;
    }
    function sendMessageResponse(targetWindow, id, 
    // tslint:disable-next-line:no-any
    args) {
        let response = createMessageResponse(id, args);
        let targetOrigin = getTargetOrigin(targetWindow);
        if (targetWindow && targetOrigin) {
            targetWindow.postMessage(response, targetOrigin);
        }
    }
    // tslint:disable-next-line:no-any
    function createMessageRequest(func, args) {
        return {
            id: nextMessageId++,
            func: func,
            args: args || []
        };
    }
    // tslint:disable-next-line:no-any
    function createMessageResponse(id, args) {
        return {
            id: id,
            args: args || []
        };
    }
    /**
     * Namespace to interact with the task module-specific part of the SDK.
     * This object is usable only on the content frame.
     */
    let task;
    (function (task) {
        /**
        * Allows an app to open the task module.
        * @param taskInfo An object containing the parameters of the task module
        * @param completionHandler Handler to call when the task module is completed
        */
        function start(taskInfo, completionHandler) {
            // Ensure that the tab content is initialized
            ensureInitialized(frameContexts.content);
            let messageId = sendMessageRequest(parentWindow, "start", [taskInfo]);
            callbacks[messageId] = completionHandler;
        }
        task.start = start;
        /**
        * Complete the task module.
        * @param result Contains the result to be sent to the bot or teh app. Typically a JSON object or a serialized version of it
        * @param appId Helps to validate that the call originates from the same appId as the one that invoked the task module
        */
        function complete(result, appId) {
            // Ensure that the tab content is initialized
            ensureInitialized(frameContexts.content);
            sendMessageRequest(parentWindow, "complete", [
                result,
                appId
            ]);
        }
        task.complete = complete;
    })(task = microsoftTeams.task || (microsoftTeams.task = {}));
})(microsoftTeams || (microsoftTeams = {}));

//# sourceMappingURL=MicrosoftTeams.js.map
