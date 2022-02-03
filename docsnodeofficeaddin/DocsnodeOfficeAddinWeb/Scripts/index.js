// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/*
    This file provides the functionality for the welcome task pane page.
*/

/// /*<reference path="./App.js" />*/
var GraphAPIToken = "";
(function () {
    "use strict";

    var Auth0AccountData = Auth0AccountData || {};

    // Replace the placeholders in the next two lines.
    Auth0AccountData.subdomain = 'login.microsoftonline.com';
    Auth0AccountData.clientID = '2f4ba6ed-0bce-4f8f-af35-7f9eb5a6f815'; // Development ID
    Auth0AccountData.clientUrl = 'https://bd97a965997b.ngrok.io/'; // Development App Service URL

    //Auth0AccountData.clientID = 'a24ca10c-4c37-422d-803d-5d00cbcd03da'; // Production ID
    //Auth0AccountData.clientUrl = 'https://docsnodetemplateswordaddin.azurewebsites.net/'; // Production App Service URL

    // The Auth0 subdomain and client ID need to be shared with the popup dialog
    localStorage.setItem('Auth0Subdomain', Auth0AccountData.subdomain);
    localStorage.setItem('Auth0ClientID', Auth0AccountData.clientID);
    localStorage.setItem('Auth0ClientUrl', Auth0AccountData.clientUrl);

    var sharePointTenantName;
    var authContext;
    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function () {
        $(document).ready(function () {
            openLoader();
            app.initialize();
            if (Office.context.platform == "OfficeOnline") {
                sharePointTenantName = Office.context.document.url;
                sharePointTenantName = sharePointTenantName.split("/")[2].split(".")[0];
                if (sharePointTenantName.indexOf("-my") > -1) {
                    sharePointTenantName = sharePointTenantName.split("-")[0];
                }
                $('#startDiv').attr('disabled', false);
                // Internet Explorer 6-11
                var isIE = /*@cc_on!@*/ false || !!document.documentMode;
                // Edge 20+
                var isEdge = !isIE && !!window.StyleMedia;

                if (isIE || isEdge) {
                    window.config = {
                        clientId: localStorage.getItem('Auth0ClientID'),
                        postlogoutredirecturi: window.location.origin,
                        redirectUri: localStorage.getItem('Auth0ClientUrl') + "TemplateChooserHome.html",
                        cachelocation: 'sessionstorage', // enable this for ie, as sessionstorage does not work for localhost.
                        callback: callbackFunction,
                        popup: true
                    };
                } else {
                    //Window.confid for online
                    window.config = {
                        clientId: localStorage.getItem('Auth0ClientID'),
                        postlogoutredirecturi: window.location.origin,
                        redirectUri: localStorage.getItem('Auth0ClientUrl') + "TemplateChooserHome.html",
                        cachelocation: 'sessionstorage', // enable this for ie, as sessionstorage does not work for localhost.
                        callback: callbackFunction,
                        displayCall: function (urlNavigate) {
                            var popupWindow = window.open(urlNavigate, "login", 'width=483, height=600');
                            if (popupWindow == null || typeof (popupWindow) == 'undefined') {
                                $("#dvPopBlock").show();
                                return;
                            }
                            if (popupWindow && popupWindow.focus)
                                popupWindow.focus();
                            var registeredRedirectUri = this.redirectUri;
                            var pollTimer = window.setInterval(function () {
                                if (!popupWindow || popupWindow.closed || popupWindow.closed === undefined) {
                                    window.clearInterval(pollTimer);
                                }
                                try {
                                    if (popupWindow.document.URL.indexOf(registeredRedirectUri) != -1) {
                                        window.clearInterval(pollTimer);
                                        window.location.hash = popupWindow.location.hash;
                                        authContext.handleWindowCallback();
                                        popupWindow.close();
                                    }
                                } catch (e) {}
                            }, 20)
                        },
                        popup: true
                    };
                }
            } else {
                $('.insertTenantName').css('display', 'block');
                window.config = {
                    clientId: localStorage.getItem('Auth0ClientID'),
                    postlogoutredirecturi: window.location.origin,
                    redirectUri: localStorage.getItem('Auth0ClientUrl') + "TemplateChooserHome.html",
                    cachelocation: 'sessionstorage', // enable this for ie, as sessionstorage does not work for localhost.
                    callback: callbackFunction
                };
            }

            authContext = new AuthenticationContext(config);
            var isCallback = authContext.isCallback(window.location.hash);
            authContext.handleWindowCallback();
            localStorage.setItem('platform', Office.context.platform);
            var user = authContext.getCachedUser();
            if (user) {
                $(".logo-title-box").addClass('hidden');
                window.location.replace(localStorage.getItem('Auth0ClientUrl') + "TemplateChooserHome.html");
            } else {
                authContext.login();
            }
            localStorage.setItem('userDisplayName', user.profile.name);
        });
    };

    function callbackFunction(errorDesc, token, error, tokenType) {
        var user = authContext.getCachedUser();
        if (user) {
            // Use the logged in user information to call your own api
            var user = authContext.getCachedUser();
            var username = user.userName;
            var upn = user.profile.upn;
            authContext.acquireToken("https://graph.microsoft.com", function (errorDesc, token, error) {
                if (error) { //acquire token failure
                    if (config.popUp) {
                        // If using popup flows
                        authContext.acquireTokenPopup("https://graph.microsoft.com", null, null, function (errorDesc, token, error) {});
                    } else {
                        // In this case the callback passed in the Authentication request constructor will be called.
                        authContext.acquireTokenRedirect("https://graph.microsoft.com", null, null);
                    }
                } else {
                    GraphAPIToken = token;
                    Window.location.replace(localStorage.getItem('Auth0ClientUrl') + "TemplateChooserHome.html");
                }
            });

        } else {
            // Initiate login
            authContext.login();
        }
    }

    // This handler responds to the success or failure message that the pop-up dialog receives from the identity provider.
    function openLoader() {
        $(".logo-title-box").removeClass('hidden');
        setTimeout(function () {
            $(".logo-title-box").addClass('hidden');
            $('.welcome_page').css('display', 'block');
        }, 300);
    }
}());