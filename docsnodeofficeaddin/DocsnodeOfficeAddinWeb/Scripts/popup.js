// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/* 
   This file provides the functionality for the page that opens in the popup. 
*/

(function () {
    "use strict";

    try {

        // Redirect to Auth0 and tell it which provider to use.
        var auth0AuthorizeEndPoint = 'https://' + localStorage.getItem('Auth0Subdomain') + '/common/oauth2/authorize/';

        $(document).ready(function () {
            redirectToIdentityProvider('windowslive');
            //$("#msAccountButton").click(function () {
            //    redirectToIdentityProvider('windowslive');
            //});
        })

        function redirectToIdentityProvider(provider) {
            window.location.replace(auth0AuthorizeEndPoint
                + '?'
                + 'response_type=id_token'
                + '&grant_type= authorization_code'
                + '&client_id=' + localStorage.getItem('Auth0ClientID')
                + '&redirect_uri=https://docsnodewordtemplafyprod.azurewebsites.net/popupRedirect.html'
                + '&scope=openid profile email offline_access https://graph.microsoft.com/User.ReadWrite'
                + '&state=123456'
                + '&nonce=8cca77c5-3758-48df-a476-e407db5225f5');
        }

    }
    catch(err) {
        console.log(err.message);
    }
}());  

