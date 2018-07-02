/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import { UserAgentApplication, Logger } from 'msal';
import { appUri, clientId, redirectUri, graphScopes, webApiScopes, graphScopesAdmin, authority } from '../helpers/AppSettings';
import Promise from 'promise';

const graphTokenStoreKey = 'GraphToken';
const webApiTokenStoreKey = 'WebApiToken';
const graphAdminTokenStoreKey = 'AdminGraphToken';
//const logger = new Msal.Logger(loggerCallback, { level: Msal.LogLevel.Verbose });
//const logger = new Msal.Logger({ level: Msal.LogLevel.Verbose, piiLoggingEnabled: true });

const level = 3;
const containsPII = false;

const optionsUserAgentApp = {
	cacheLocation: 'localStorage',
	logger: new Logger((level, message, containsPII) => {
		const logger = level === 0 ? console.error : level === 1 ? console.warn : console.log;
		//logger(`AD: ${message}`);
		console.log(`AD: ${message}`);
	}),
	redirectUri: redirectUri
};


// Initialize th library
var userAgentApplication = new UserAgentApplication(
	clientId,
	authority,
	tokenReceivedCallback,
	optionsUserAgentApp);


function getUserAgentApplication() {
	return userAgentApplication;
}

function handleToken(accesstoken) {
	if (accesstoken) {
		localStorage.setItem(graphTokenStoreKey, accesstoken);
	}
}

function handleWebApiToken(idToken) {
	if (idToken) {
		console.log("handleWebApiToken-not empty");
		localStorage.setItem(webApiTokenStoreKey, idToken);
	}
}

function handleGraphAdminToken(idToken) {
	if (idToken) {
		console.log("handleGraphAdminToken-not empty");
		localStorage.setItem(graphAdminTokenStoreKey, idToken);
	}
}

function handleRemoveToken() {
	localStorage.removeItem(graphTokenStoreKey);
}

function handleRemoveWebApiToken() {
	localStorage.removeItem(webApiTokenStoreKey);
}

function handleRemoveGraphAdminToken() {
	localStorage.removeItem(graphAdminTokenStoreKey);
}

function handleError(error) {
	console.log(`AuthHelper: ${error}`);
}


function tokenReceivedCallback(errorMessage, token, error, tokenType) {
	//This function is called after loginRedirect and acquireTokenRedirect. Use tokenType to determine context. 
	//For loginRedirect, tokenType = "id_token". For acquireTokenRedirect, tokenType:"access_token".
	localStorage.setItem("loginRedirect", "tokenReceivedCallback");
	if (!errorMessage && token) {
		this.acquireTokenSilent(graphScopes)
			.then(accessToken => {
			// Store token in localStore
			handleToken(accessToken);
			handleWebApiToken(token);
			})
			.catch(error => {
				handleError("tokenReceivedCallback-acquireTokenSilent: " + error);
				// TODO: need to add aquiretokenpopup or similar
			});
	} else {
		handleError("tokenReceivedCallback: " + error);
	}
}


export default class AuthClient {
	constructor(props) {

		// Get the instance of UserAgentApplication.
		this.authClient = getUserAgentApplication();

		this.userProfile = [];
	}

	loginPopup() {
		return new Promise((resolve, reject) => {
			this.authClient.loginPopup(graphScopes)
				.then(function (idToken) {
					handleWebApiToken(idToken);
					resolve(idToken);
				})
				.catch((err) => {
					reject(err);
				});
		});
	}

	loginPopupGraphAdmin() {
		return new Promise((resolve, reject) => {
			this.authClient.loginPopup(graphScopesAdmin)
				.then(function (idToken) {
					handleGraphAdminToken(idToken);
					resolve(idToken);
				})
				.catch((err) => {
					reject(err);
				});
		});
	}

	loginRedirectPromise() {
		return new Promise((resolve, reject) => {
			localStorage.setItem("loginRedirect", "loginRedirect start");
			this.authClient.loginRedirect(graphScopes)
				.then(function (idToken) {
					handleWebApiToken(idToken);
					localStorage.setItem("loginRedirect", "loginRedirect got access_token");
					resolve(idToken);
				})
				.catch((err) => {
					localStorage.setItem("loginRedirect", "AuthHelper_loginRedirect error: " + err);
					console.log("AuthHelper_loginRedirect error: " + err);
					reject(err);
				});
		});
	}

	loginRedirect() {
		localStorage.setItem("loginRedirect", "loginRedirect2 start");
		localStorage.setItem("AuthRedirect", "start");
		return this.authClient.loginRedirect(graphScopes);
	}

	acquireTokenSilent() {
		return new Promise((resolve, reject) => {
			this.authClient.acquireTokenSilent(graphScopes, authority)
				.then(function (accessToken) {
					handleToken(accessToken);
					resolve(accessToken);
				})
				.catch((err) => {
					reject(err);
				});
		});
	}

	acquireTokenSilentParam(extraQueryParameters) {
		return new Promise((resolve, reject) => {
			localStorage.setItem("AuthError", "acquireTokenSilentParam started");
			this.authClient.acquireTokenSilent(graphScopes, authority, null, extraQueryParameters)
				.then(function (accessToken) {
					console.log("AuthHelper_acquireTokenSilentParam got access_token");
					handleToken(accessToken);
					resolve(accessToken);
				})
				.catch((err) => {
					localStorage.setItem("AuthError", "AuthHelper_acquireTokenSilentParam error: " + err);
					console.log("AuthHelper_acquireTokenSilentParam error: " + err);
					reject(err);
				});
		});
	}

	acquireWebApiTokenSilent() {
		return new Promise((resolve, reject) => {
			this.authClient.acquireTokenSilent(webApiScopes, authority)
				.then(function (accessToken) {
					handleWebApiToken(accessToken);
					resolve(accessToken);
				})
				.catch((err) => {
					reject(err);
				});
		});
	}

	acquireWebApiTokenSilentParam(extraQueryParameters) {
		return new Promise((resolve, reject) => {
			localStorage.setItem("AuthError", "acquireTokenSilentParam started");
			this.authClient.acquireTokenSilent(webApiScopes, authority, null, extraQueryParameters)
				.then(function (accessToken) {
					console.log("AuthHelper_acquireTokenSilentParam got access_token");
					handleToken(accessToken);
					resolve(accessToken);
				})
				.catch((err) => {
					localStorage.setItem("AuthError", "AuthHelper_acquireTokenSilentParam error: " + err);
					console.log("AuthHelper_acquireTokenSilentParam error: " + err);
					reject(err);
				});
		});
	}

	acquireGraphAdminTokenSilent() {
		return new Promise((resolve, reject) => {
			this.authClient.acquireTokenSilent(graphScopesAdmin, authority)
				.then(function (accessToken) {
					handleGraphAdminToken(accessToken);
					resolve(accessToken);
				})
				.catch((err) => {
					reject(err);
				});
		});
	}

	accuireTokenAndWebTokenSilent() {
		this.acquireTokenSilent()
			.then(res => {
				this.acquireWebApiTokenSilent()
					.then(res => {
						localStorage.setItem("AuthError", "AuthHelper_accuireTokenAndWebTokenSilent_acquireWebApiTokenSilent done");
						return res;
					})
					.catch((err) => {
						localStorage.setItem("AuthError", "AuthHelper_accuireTokenAndWebTokenSilent_acquireWebApiTokenSilent error: " + err);
						console.log("AuthHelper_accuireTokenAndWebTokenSilent_acquireWebApiTokenSilent error: " + err);
						return err;
					});
			})
			.catch((err) => {
				localStorage.setItem("AuthError", "AuthHelper_accuireTokenAndWebTokenSilent_acquireTokenSilent error: " + err);
				console.log("AuthHelper_accuireTokenAndWebTokenSilent_acquireTokenSilent error: " + err);
				return err;
			});
	}

	getUserAsync() {
		return new Promise((resolve, reject) => {
			let res = this.authClient.getUser();
			if (res) {
				resolve(res);
			} else {
				reject(res);
			}
		});
	}

	getUser() {
		let userResult = this.authClient.getUser();
		if (userResult === null) {
			return 'null if'; // TODO: Temporal return for debug
		}
		return userResult;
	}

	getUserProfile() {
		return new Promise((resolve, reject) => {
			if (this.userProfile) {
				let userResult = this.getUser();
				if (userResult.displayableId === this.userProfile.userPrincipalName) {
					resolve(this.userProfile);
				}
				reject('null if');
			} else {
				reject('null if'); // TODO: Temporal return for debug
			}
		});
	}

	callGetUserProfile() {
		return new Promise((resolve, reject) => {
			// Call the Web API with the AccessToken
			//const accessToken = this.getWebApiToken();
			const userPrincipalName = this.getUser();
			console.log("AuthHelper_callGetUserProfile getUser: " + userPrincipalName.displayableId);
			if (userPrincipalName.displayableId.length > 0) {
				const endpoint = appUri + "/api/UserProfile?upn=" + userPrincipalName.displayableId;

				this.callWebApiWithToken(endpoint, "GET")
					.then(data => {
						if (data) {
							this.userProfile = data;
							resolve({
								roles: data.userRoles,
								id: data.id,
								displayName: data.displayName,
								mail: data.mail,
								userPrincipalName: data.userPrincipalName
							});
						} else {
							console.log("Error callGetUserProfile: " + JSON.stringify(data));
							reject(data);
						}
					})
					.catch(function (err) {
						handleError('Error when calling endpoint in callGetUserProfile: ' + JSON.stringify(err));
						reject(err);
					});
			} else {
				reject("Error when calling endpoint in callGetUserProfile: no current user exists in context");
			}
		});
	}

	isAuthenticated() {
		let tokenResult = localStorage.getItem(graphTokenStoreKey);
		let userResult = this.getUser();
		let msalError = this.getAuthError();

		if (!userResult) {
			return false;
		}
		if (!tokenResult) {
			return false;
		}
		if (msalError) {
			return false;
		}
		return true;
	}

	isCallBack(hash) {
		let isCallback = this.authClient.isCallback(hash);
		return isCallback;
	}

	getGraphToken() {
		if (!this.isAuthenticated()) {
			console.log("getGraphToken isAuth: false");
		}
		return localStorage.getItem(graphTokenStoreKey);
	}

	getWebApiToken() {
		if (!this.isAuthenticated()) {
			console.log("getWebApiToken isAuth: false");
		}
		return localStorage.getItem(webApiTokenStoreKey);
	}

	getGraphAdminToken() {
		if (!this.isAuthenticated()) {
			console.log("getGraphAdminToken isAuth: false");
		}
		return localStorage.getItem(graphAdminTokenStoreKey);
	}

	getAuthError() {
		return localStorage.getItem("msal.error");
	}

	getAuthRedirectState() {
		return localStorage.getItem("AuthRedirect");
	}

	getIdToken() {
		console.log("getIdToken");
		return localStorage.getItem('msal.idtoken');
	}

	logout() {
		handleRemoveToken();
		handleRemoveWebApiToken();
		handleRemoveGraphAdminToken();
		return new Promise((resolve, reject) => {
			this.authClient.logout()
				.then(res => {
					resolve(res);
				})
				.catch(err => {
					reject(err);
				});
		});
	}

	callWebApiWithToken(endpoint, method) {
		return new Promise((resolve, reject) => {
			let headers = new Headers();
			let bearer = "Bearer " + this.getGraphToken();

			headers.append("Authorization", bearer);
			let options = {
				method: method,
				headers: headers
			};

			fetch(endpoint, options)
				.then(function (response) {
					var contentType = response.headers.get("content-type");
					if (response.status === 200 && contentType && contentType.indexOf("application/json") !== -1) {
						response.json()
							.then(function (data) {
								// return response
								resolve(data);
							})
							.catch(function (err) {
								handleError(err + ' when calling endpoint: ' + endpoint + ' method:' + options[method]);
								reject(err);
							});
					} else {
						response.json()
							.then(function (data) {
								// Display response as error in the page
								reject(data);
							})
							.catch(function (err) {
								handleError(err + ' when calling endpoint: ' + endpoint);
								reject(err);
							});
					}
				})
				.catch(function (err) {
					handleError(err + ' when calling endpoint: ' + endpoint);
					reject(err);
				});
		});
	}
}