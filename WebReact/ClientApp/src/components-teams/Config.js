/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import AuthHelper from '../helpers/AuthHelper';
import GraphSdkHelper from '../helpers/GraphSdkHelper';
import { appUri } from '../helpers/AppSettings';
import Utils from '../helpers/Utils';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export class Config extends Component {
	displayName = Config.name

	constructor(props) {
		super(props);


		if (window.authHelper) {
			this.authHelper = window.authHelper;
		} else {
			// Initilize the AuthService and save it in the window object.
			this.authHelper = new AuthHelper();
			window.authHelper = this.authHelper;
		}

		if (window.sdkHelper) {
			this.sdkHelper = window.sdkHelper;
		} else {
			// Initilize the AuthService and save it in the window object.
			this.sdkHelper = new GraphSdkHelper();
			window.sdkHelper = this.sdkHelper;
		}

		this.utils = new Utils();

		this.authInProgress = false;

		try {
			microsoftTeams.initialize();
		}
		catch (err) {
			console.log("ProposalManagement_ConfigTAB error initializing teams: " + JSON.stringify(err));
		}
		finally {
			this.state = {
				isAuthenticated: this.authHelper.isAuthenticated(),
				channelName: "",
				channelId: "",
				teamName: "",
				groupId: "",
                contextUpn: "",
                userRoleList: []
			};

			/** Pass the Context interface to the initialize function below */
			//microsoftTeams.getContext(context => this.initialize(context));
		}
	}


	componentWillMount() {
		// Get the teams context
		this.getTeamsContext();

		this.setState({
			isAuthenticated: this.authHelper.isAuthenticated(),
			channelName: this.getQueryVariable('channelName'),
			teamName: this.getQueryVariable('teamName'),
			groupId: this.getQueryVariable('groupId'),
			channelId: this.getQueryVariable('channelId')
			
		});

        this.setChannelConfig();

		if (this.state.isAuthenticated) {
			this.authHelper.callGetUserProfile()
				.then(userProfile => {
					microsoftTeams.settings.setValidityState(true);
					this.setState({
						userProfile: userProfile,
						isAuthenticated: true,
						displayName: `Hello ${userProfile.displayName}!`
					});
				});
		}
	}

	componentDidUpdate() {
		if (this.state.teamName.length > 0) {
			this.authUser();
		}

		microsoftTeams.settings.setValidityState(this.authHelper.isAuthenticated());
    }

    getUserRoles() {
        // call to API fetch data
        return new Promise((resolve, reject) => {
            console.log("Config_getUserRoles userRoleList fetch");
            let requestUrl = 'api/RoleMapping';
            fetch(requestUrl, {
                method: "GET",
                headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
            })
                .then(response => response.json())
                .then(data => {
                    try {
                        console.log("Config_getUserRoles userRoleList data lenght: " + data.length);

                        let userRoleList = [];
                        //console.log(data);
                        for (let i = 0; i < data.length; i++) {
                            let userRole = {};
                            userRole.id = data[i].id;
                            userRole.adGroupName = data[i].adGroupName;
                            userRole.roleName = data[i].roleName;
                            userRole.processStep = data[i].processStep;
                            userRole.channel = data[i].channel;
                            userRole.adGroupId = data[i].adGroupId;
                            userRole.processType = data[i].processType;
                            userRoleList.push(userRole);
                        }
                        this.setState({ userRoleList: userRoleList });
                        console.log("Config_getUserRoles userRoleList lenght: " + userRoleList.length);
                        resolve(true);
                    }
                    catch (err) {
                        reject(err);
                    }

                });
        });
    }

    setChannelConfig() {
        this.getUserRoles()
            .then(res => {
                let chName = this.getQueryVariable('channelName');
                let teamName = this.getQueryVariable('teamName');
                let channelId = this.getQueryVariable('channelId');
                let channelName = this.getQueryVariable('channelName');
                let channelMapping = this.state.userRoleList.filter(x => x.channel.toLowerCase() === channelName.toLowerCase());
                let tabName = "";

                if (channelName === "General") {
                    tabName = "rootTab";
                } else if (channelMapping.length > 0) {
                    if (channelMapping.processType !== "Base" && channelMapping.processType !== "Administration") {
                        tabName = channelMapping[0].processType;
                    }
                }

                console.log("Config_setChannelConfig tabName: " + tabName + " ChannelName: " + channelName);
                console.log(channelMapping);

                if (tabName !== "") {
                    microsoftTeams.settings.registerOnSaveHandler(function (saveEvent) {
                        microsoftTeams.settings.setSettings({
                            entityId: "PM" + channelName,
                            contentUrl: appUri + "/tab/" + tabName + "?channelName=" + chName + "&teamName=" + teamName + "&channelId=" + channelId,
                            suggestedDisplayName: "Proposal Manager",
                            websiteUrl: appUri + "/tabMob/" + tabName + "?channelName=" + chName + "&teamName=" + teamName + "&channelId=" + channelId

                        });
                        saveEvent.notifySuccess();
                    });

                    this.setState({
                        validityState: microsoftTeams.settings.validityState
                    });
                }
            })
            .catch(err => {
                console.log("Config_getUserRoles error: ");
                console.log(err);
            });
    }

	getTeamsContext() {
		microsoftTeams.getContext(context => {
			this.setState({
				channelName: context.channelName,
				channelId: context.channelId,
				teamName: context.teamName,
				groupId: context.groupId,
				contextUpn: context.upn,
				validityState: microsoftTeams.settings.validityState
			});
		});
	}

	authUser() {
		//if (localStorage.getItem("configLoginState") !== "1") {
		if (!this.authInProgress) {
			//localStorage.setItem("configLoginState", "1");
			this.authInProgress = true;
			
			microsoftTeams.authentication.authenticate({
				url: window.location.origin + "/tab/tabauth",
				width: 670,
				height: 570,
				successCallback: function (result) {
					//getUserProfile(result.accessToken);
					microsoftTeams.settings.setValidityState(true);
				},
				failureCallback: function (reason) {
					//handleAuthError(reason);
					if (reason === "FailedToOpenWindow") {
						// Try loginPopup since it may be browser client
						localStorage.setItem("configLoginState", "1");
					}
				}
			});
			if (localStorage.getItem("configLoginState") === "1") {
				//this.login();
			}
		}
	}

	refresh() {
		window.location.reload();
	}

	// Returns the value of a query variable.
	getQueryVariable = (variable) => {
		const query = window.location.search.substring(1);
		const vars = query.split('&');
		for (const varPairs of vars) {
			const pair = varPairs.split('=');
			if (decodeURIComponent(pair[0]) === variable) {
				return decodeURIComponent(pair[1]);
			}
		}
		return null;
	}

	render() {

		return (
			<div className="BgConfigImage">
				
				<br /><br /><br /><br /><br /><br /><br />	<br />
				<p className="WhiteFont">{this.state.displayName ? this.state.displayName : 'Welcome'}</p>
				
				<PrimaryButton className='pull-right refreshbutton' onClick={this.refresh.bind(this)}>
					Refresh 
				</PrimaryButton>
			<br />
			</div>
		);
	}
}
