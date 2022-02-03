// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/* This file provides the functionality for the page that opens in the popup 
   after the user has logged in with a provider 
*/

(function () {
    "use strict";
    window.tokenDefer = $.Deferred();
    //var sharePointTenantName ;
    var response = { "status": "none"};
    // Office.initialize must be called on every page where Office JavaScript is 
    // called. Other initialization code should go inside it.
    Office.initialize = function () {
        $(document).ready(SignIn)        
    };
        function SignIn(){  
            try {
                console.log("doc ready from AD Login");
                window.config = { 
                    //tenant: sharePointTenantName + '.onmicrosoft.com',
                    clientId: localStorage.getItem('Auth0ClientID'),
                    postLogoutRedirectUri: window.location.origin,
                    redirectUri: "https://docsnodewordtemplafyprod.azurewebsites.net/popupRedirect.html",//knut prod
                   // redirectUri: "https://docsnodeofficewordaddin.azurewebsites.net/popupRedirect.html",//knut devlp
                    //redirectUri: "https://docsnodeexcel.azurewebsites.net/popupRedirect.html",//knut devlp2
                    //redirectUri: "https://localhost:44335/popupRedirect.html",//local development
                    cacheLocation: 'sessionStorage' // enable this for IE, as sessionStorage does not work for localhost.
                };
                var authContext = new AuthenticationContext(config);

                var isCallback = authContext.isCallback(window.location.hash);
                authContext.handleWindowCallback();
                
                var user = authContext.getCachedUser();
                //  if (isCallback && !authContext.getLoginError()) {
                //     window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
                //}
                // If not logged in force login
                var cachedToken = authContext.getCachedToken(window.config.clientId);
                if (cachedToken) {
                    console.log("user already logged in");
                    // Logged in already
                    authContext.acquireToken(authContext.config.loginResource, function (error, token) {
                        if (error || !token) {
                            console.log("ADAL error occurred: " + error);
                            window.tokenDefer.reject();
                        }
                        console.log("got the token.. resolving tokendefer");
                        response.status = "success";                        
                        //accessTokenForAuth0 = getHashStringParameter(token);
                        var messageObject = { outcome: "success"};
                        var jsonMessage = JSON.stringify(messageObject);

                        // Tell the task pane about the outcome.
                        Office.context.ui.messageParent(jsonMessage);
                        window.tokenDefer.resolve(token);                       
                    });
                }
                else {
                    // NOTE: you may want to render the page for anonymous users and render
                    // a login button which runs the login function upon click.
                    console.log("calling login");
                    response.status = "error";                    
                    authContext.login();                  
                }
                // Auth0 adds its access token as a hash (#) value on the URL
                //var accessTokenForAuth0 = getHashStringParameter(response.accessToken);            

                // //Create the outcome message and send it to the task pane.
                //var messageObject = {outcome: "success", auth0Token: accessTokenForAuth0};            
                //var jsonMessage = JSON.stringify(messageObject);

                //// Tell the task pane about the outcome.
                //Office.context.ui.messageParent(jsonMessage);
            }
            catch(err) {
                
                // Create the outcome message and send it to the task pane.
                var messageObject = {outcome: "failure", error: err.message};            
                var jsonMessage = JSON.stringify(messageObject);

                // Tell the task pane about the outcome.
                Office.context.ui.messageParent("message to parent " + jsonMessage); 
            }
        
    }

    // Function to retrieve a hash string value when the hash
    // value is structured like query parameters.
    function getHashStringParameter(paramToRetrieve) {
        var hash = location.hash.replace('#', '');
        var params = hash.split("&");
        for (var i = 0; i < params.length; i = i + 1) {
            var singleParam = params[i].split("=");
            if (singleParam[0] == paramToRetrieve)
                return singleParam[1];
        }
    }
}());    
 