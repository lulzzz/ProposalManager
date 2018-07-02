/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import { UserAgentApplication, Logger, User } from 'msal';

const clientId = 'aff7757b-d5e5-4c60-8bb5-a8773d387c0b';
const webApiScope = [clientId];
const authority = null; // Null for login as common (multi-tenant also) eg. https://login.microsoftonline.com/common/oauth2/v2.0/authorize
const graphScopes = ["user.readbasic.all", "mail.send", "files.read"];
const graphTokenStoreKey = 'AD.GraphToken';
const webApiTokenStoreKey = 'AD.WebApiToken';
const userStoreKey = 'AD.User';
const optionsUserAgentApp = {
    cacheLocation: 'localStorage',
    logger: new Logger((level, message, containsPII) => {
        const logger = level === 0 ? console.error : level === 1 ? console.warn : console.log;
        logger(`AD: ${message}`);
    })
};

// Initialize th library
var userAgentApplication = new UserAgentApplication(
    clientId,
    authority,
    tokenReceivedCallback,
    optionsUserAgentApp);

function handleError(error) {
    console.error(`AD: ${error}`);
}

function tokenReceivedCallback(errorMessage, token, error, tokenType) {
    //This function is called after loginRedirect and acquireTokenRedirect. Use tokenType to determine context. 
    //For loginRedirect, tokenType = "id_token". For acquireTokenRedirect, tokenType:"access_token".

    if (!errorMessage && token) {
        this.acquireTokenSilent(graphScopes).then(accessToken => {
            // Store token in localStore
            localStorage.setItem(graphTokenStoreKey, accessToken);
            localStorage.setItem(webApiTokenStoreKey, token);
        }, function (error) {
            handleError(error);
            this.acquireTokenPopup(graphScopes).then(accessToken => {
                // Store token in localStore
                localStorage.setItem(graphTokenStoreKey, accessToken);
                localStorage.setItem(webApiTokenStoreKey, token);
            }, function (error) {
                handleError(error);
            });
        });
    } else {
        handleError(error);
    }
}


export class AuthClient {
  
    constructor() {
    }

    public login(): Promise<User>
	{
		//var user = getUser();

		//if (user)
		//{
		//	return;
		//}
        return new Promise((resolve, reject) =>
        {
            userAgentApplication.loginPopup(graphScopes).then(function (idToken)
            {
                //Login Success
                userAgentApplication.acquireTokenSilent(webApiScope)
                .then(
                    accessToken =>
                    {
                        let user = userAgentApplication.getUser();
                        localStorage.setItem(userStoreKey, JSON.stringify(user));
                        localStorage.setItem(webApiTokenStoreKey, accessToken);
                        return resolve(user);
                    },
                    error =>
                    {
                        return reject(error);
                    }
                )
            });
        });
    }
    
    public getWebApiToken()
	{
		return localStorage.getItem(webApiTokenStoreKey);
	}

	public getUser()
	{
        let userData = localStorage.getItem(userStoreKey);
        
        if(userData)
        {
            return JSON.parse(userData);
        }

        return null;
	}
}
