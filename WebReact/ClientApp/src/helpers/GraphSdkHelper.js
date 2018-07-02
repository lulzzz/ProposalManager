/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import async from 'async';
import Promise from 'promise';
import AuthHelper from './AuthHelper';
import Utils from './Utils';
import { teamsAppInstanceId } from './AppSettings';

export default class GraphSdkHelper {
	constructor(props) {
		const MicrosoftGraph = require("@microsoft/microsoft-graph-client");

		if (window.authHelper) {
			this.authHelper = window.authHelper;
		} else {
			// Initilize the AuthService and save it in the window object.
			this.authHelper = new AuthHelper();
			window.authHelper = this.authHelper;
		}

		let graphToken = this.authHelper.getGraphToken();
		if (!graphToken) {
			console.log("GraphSdkHelper_Constructor graphToken is null");
		}

		// Initialize the Graph SDK.
		this.client = MicrosoftGraph.Client.init({
			debugLogging: true,
			authProvider: (done) => {
				done(null, graphToken);
			}
		});

        this.utils = new Utils();

		this._handleError = this._handleError.bind(this);
		this.props = props;
	}

    initClientAdmin() {
        const MicrosoftGraph = require("@microsoft/microsoft-graph-client");

        let graphAdminToken = this.authHelper.getGraphAdminToken();

        if (graphAdminToken) {
            this.clientAdmin = MicrosoftGraph.Client.init({
                debugLogging: true,
                authProvider: (done) => {
                    done(null, graphAdminToken);
                }
            });
        } else {
            console.log("GraphSdkHelper_initClientAdmin graphAdminToken = empty");
        }
    }

	// GET me (OLD TODO: Deprecate after replacement is tested)
	getMeOld(callback) {
		this.client
			.api('/me')
			.select('displayName')
			.get((err, res) => {
				if (!err) {
					callback(null, res);
				}
				else this._handleError(err);
			});
	}

	// GET me
	getMe() {
		// debugger;
        return new Promise((resolve, reject) => {
            this.client
                .api('/me')
                .select('displayName')
                .get()
                .then(res => {
                    resolve(res);
                })
                .catch(err => {
                    this._handleError(err);
                    reject(err);
                });
        });
	}

	// GET me UPN
	getMeUpn() {
        return new Promise((resolve, reject) => {
            this.client
                .api('/me')
                .select('userPrincipalName')
                .get()
                .then(res => {
                    resolve(res);
                })
                .catch(err => {
                    this._handleError(err);
                    reject(err);
                });
        });
	}

	// GET me Details
	getMeDetails() {
        return new Promise((resolve, reject) => {
            this.client
                .api('/me')
                .select('displayName,givenName,surname,emailAddresses,userPrincipalName')
                .get()
                .then(res => {
                    resolve(res);
                })
                .catch(err => {
                    this._handleError(err);
                    reject(err);
                });
        });
	}

	// GET me/people
	getPeople(callback) {
		this.client
			.api('/me/people')
			.version('beta')
			.filter(`personType eq 'Person'`)
			.select('displayName,givenName,surname,emailAddresses,userPrincipalName')
			.top(20)
			.get((err, res) => {
				if (err) {
					this._handleError(err);
				}
				callback(err, res ? res.value : []);
			});
	}

	// *** TEST *** //
	// GET me/memberOf
	getMemberships(callback) {
		this.client
			.api('/me/memberof')
			.version('beta')
			.select('id')
			.get((err, res) => {
				if (err) {
					this._handleError(err);
				}
				callback(err, res ? res.value : []);
			});
	}

	// GET user/id/photo/$value for each person
	getProfilePics(personas, callback) {
		const pic = (p, done) => {
			this.client
				.api(`users/${p.id}/photo/$value`)
				.header('Cache-Control', 'no-cache')
				.responseType('blob')
				.get((err, res, rawResponse) => {
					if (err) {
						done(err);
					}
					else {
						p.imageUrl = window.URL.createObjectURL(rawResponse.xhr.response);
						p.initialsColor = null;
						done();
					}
				});
		};
		async.each(personas, pic, (err) => {
			callback(err);
		});
	}

	// GET users?$filter=displayName startswith('{searchText}')
	searchForPeople(searchText, callback) {
		this.client
			.api('/users')
			.filter(`startswith(displayName,'${searchText}')`)
			.select('displayName,givenName,surname,mail,userPrincipalName,id')
			.top(20)
			.get((err, res) => {
				if (err) {
					this._handleError(err);
				}
				callback(err, res ? res.value : []);
			});
	}

	// POST me/sendMail
	sendMail(recipients, callback) {
		const email = {
			Subject: 'Email from the Microsoft Graph Sample with Office UI Fabric',
			Body: {
				ContentType: 'HTML',
				Content: `<p>Thanks for trying out Office UI Fabric!</p>
		  <p>See what else you can do with <a href="http://dev.office.com/fabric#/components">
		  Fabric React components</a>.</p>`
			},
			ToRecipients: recipients
		};
		this.client
			.api('/me/sendMail')
			.post({ 'message': email, 'saveToSentItems': true }, (err, res, rawResponse) => {
				if (err) {
					this._handleError(err);
				}
				callback(err, rawResponse.req._data.message.ToRecipients);
			});
	}

	// GET drive/root/children
	getFiles(nextLink, callback) {
		let request;
		if (nextLink) {
            request = this.client
                .api(nextLink);
		}
		else {
            request = this.client
                .api('/me/drive/root/children')
                .select('name,createdBy,createdDateTime,lastModifiedBy,lastModifiedDateTime,webUrl,file')
                .top(100); // default result set is 200
		}
		request.get((err, res) => {
			if (err) {
				this._handleError(err);
			}
			callback(err, res);
		});
	}

	// Create Team in 2 steps 1st the group then the team
	// POST /groups
	createTeamGroup(displayName, description) {
        return new Promise((resolve, reject) => {
            this.initClientAdmin();

            if (this.clientAdmin) {
                //let mailNickname = displayName.replace(/[\s`~!@#$%^&*()_|+\-=?;:'",.<>{}[]\\\/]/gi, '');
                const regExpr = /[^a-zA-Z0-9-.\/s]/g;
                let mailNickname = displayName.replace(regExpr, "");

                console.log("GrapSDKHelper_createTeamGroup nickname: " + mailNickname);
                const groupSettings = {
                    "description": description,
                    "displayName": displayName,
                    "groupTypes": ["Unified"],
                    "mailEnabled": true,
                    "mailNickname": mailNickname,
                    "securityEnabled": false
                };

                const teamSettings = `{
			        "memberSettings": {
				        "allowCreateUpdateChannels": "true"
			        },
			        "messagingSettings": {
				        "allowUserEditMessages": "true",
				        "allowUserDeleteMessages": "true"
			        },
			        "funSettings": {
				        "allowGiphy": "true",
				        "giphyContentRating": "strict"
			        }}`;
                this.clientAdmin
                    .api('/groups')
                    .post(groupSettings)
                    .then(res => {
                        this.clientAdmin
                            .api('/groups/' + res.id + '/team')
                            .version('beta')
                            .put(teamSettings)
                            .then(res => {
                                console.log("GraphSdkHelper_createTeamGroup group created: " + JSON.stringify(res));
                                resolve(res.id);
                            })
                            .catch(err => {
                                this._handleError("createTeamGroup_put: " + err);
                                reject(err);
                            });
                    })
                    .catch(err => {
                        this._handleError("createTeamGroup_post: " + err);
                        reject(err);
                    });
            } else {
                console.log("GraphSdkHelper_createTeamGroup clientAdmin invalid");
                reject(false);
            }
        });
    }

	// POST /groups/{id}/channels
	createChannel(displayName, description, teamId) {
        return new Promise((resolve, reject) => {
            this.initClientAdmin();

            if (this.clientAdmin) {
                const channelSettings = {
                    "displayName": displayName,
                    "description": description
                };
                this.clientAdmin
                    .api('/groups/' + teamId + '/channels')
                    .version('beta')
                    .post(channelSettings)
                    .then(res => {
                        console.log("createChannel created: ");
                        console.log(res);
                        resolve(res);
                    })
                    .catch(err => {
                        this._handleError("createChannel: " + err);
                        //reject(err);
                        resolve(err); // Forcing resolve so it continues creating other channels
                    });
            } else {
                console.log("GraphSdkHelper_createChannel clientAdmin invalid");
                reject(false);
            }
        });
	}

    // POST /beta/teams/{id}/apps
    addAppToTeam(teamId) {
        return new Promise((resolve, reject) => {
            this.initClientAdmin();

            if (this.clientAdmin) {
                const jsonBody = {
                    "id": teamsAppInstanceId
                };
                this.clientAdmin
                    .api('/teams/' + teamId + '/apps')
                    .version('beta')
                    .post(jsonBody)
                    .then(res => {
                        console.log("addAppToTeam completed");
                        console.log(res);
                        resolve(res);
                    })
                    .catch(err => {
                        this._handleError("addAppToTeam: " + err);
                        //reject(err);
                        resolve(err); // Forcing resolve so it continues creating other channels
                    });
            } else {
                console.log("addAppToTeam clientAdmin invalid");
                reject(false);
            }
        });
    }

	// GET /groups/{id}/channels/{id}
	getChannel(teamId, channelName) {
        return new Promise((resolve, reject) => {
            this.client
                .api('/groups/' + teamId + '/channels/' + channelName)
                .version('beta')
                .get()
                .then(res => {
                    console.log("getChannel: " + res);
                    resolve(res);
                })
                .catch(err => {
                    this._handleError(err);
                    reject(err);
                });
        });
	}

	// GET /groups/{id}/channels/{id}
	getTeamByName(teamName) {
        return new Promise((resolve, reject) => {
            this.client
                .api("/groups?filter=startswith(displayName,'" + teamName + "')")
                .version('beta')
                .get()
                .then(res => {
                    console.log("getTeamByName: " + res);
                    resolve(res);
                })
                .catch(err => {
                    this._handleError(err);
                    reject(err);
                });
        });
	}


	getTeamById(teamId) {
        return new Promise((resolve, reject) => {
            this.client
                .api("/groups/" + teamId + "/team")
                .version('beta')
                .get()
                .then(res => {
                    console.log("getTeamById: " + res);
                    resolve(res);
                })
                .catch(err => {
                    this._handleError(err);
                    reject(err);
                });
        });
	}


	// POST /groups/{id}/channels/{id}/chatthreads
	createChatThread(groupId, channelId, requestBody) {
        return new Promise((resolve, reject) => {
            this.client
                .api('/groups/' + groupId + '/channels/' + channelId + '/chatthreads')
                .version('beta')
                .post(requestBody)
                .then(res => {
                    console.log("createChatThread: " + res);
                    resolve(res);
                })
                .catch(err => {
                    this._handleError(err);
                    reject(err);
                });
        });
	}

	// PUT /to sharepoint
	uploadFile(file, siteId) {
		let url = 'https://graph.microsoft.com/v1.0/sites/' + siteId + 'drive/root:/' + file.name + ':/content';
		this.client
			.api(url)
			.put(file, (err, res) => {
				if (err) {
					console.log(err);
					return;
				}
				console.log("We've updated your picture!");
			});
	}

	uploadFileToChannel(file, siteId) {
		let url = 'https://graph.microsoft.com/v1.0/sites/' + siteId + 'drive/root:/' + file.name + ':/content';
		this.client
			.api(url)
			.put(file, (err, res) => {
				if (err) {
					console.log(err);
					return;
				}
				console.log("We've updated your picture!");
			});
	}

	_handleError(err) {
		console.log("GraphSdkHelper error: " + err.statusCode + " - " + err.message);

		// This method just redirects to the login function when the token is expired.
		// For production need to implement some token management with refresh token, etc.
		if (err.statusCode === 401 && err.message === 'Access token has expired.') {
			console.log('GraphSdkHelper_handleError tryAcquireTokenSilent 401');
			this.authHelper.acquireTokenSilent()
				.then(res => {
					console.log('GraphSdkHelper_handleError tryAcquireTokenSilent executed');
				})
				.catch(err => {
                    console.log('GraphSdkHelper_handleError tryAcquireTokenSilent failed error code: ');
                    console.log(err);
				});
		}
	}
}