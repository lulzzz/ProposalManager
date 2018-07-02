/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// Global imports
import React, { Component } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import { Route } from 'react-router';
import GraphSdkHelper from './helpers/GraphSdkHelper';
import AuthHelper from './helpers/AuthHelper';
import Utils from './helpers/Utils';
// Teams Add-in imports
import { ThemeStyle } from 'msteams-ui-components-react';

import { Home } from './components-teams/Home';

import { Config } from './components-teams/Config';
import { Privacy } from './components-teams/Privacy';
import { TermsOfUse } from './components-teams/TermsOfUse';

//import { Checklist } from './views-teams/Proposal/Checklist';
import { Checklist } from './components-teams/Checklist';
import { RootTab } from './components-teams/RootTab'; //'./views-teams/Proposal/RootTab';
import { TabAuth } from './components-teams/TabAuth';
import { ProposalStatus } from './components-teams/ProposalStatus';
import { CustomerDecision } from './components-teams/CustomerDecision';
// Components mobile
import { RootTab as RootTabMob} from './components-mobile/RootTab';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

export class AppTeams extends Component {
	displayName = AppTeams.name

	constructor(props) {
		super(props);

		initializeIcons();

		console.log("ProposalManagement_AppTeams");

		if (window.authHelper) {
			console.log("AppTeams: Auth already initialized");
			this.authHelper = window.authHelper;
		} else {
			// Initilize the AuthService and save it in the window object.
			console.log("AppTeams: Initialize auth");
			this.authHelper = new AuthHelper();
			window.authHelper = this.authHelper;
		}

		if (window.sdkHelper) {
			this.sdkHelper = window.sdkHelper;
		} else {
			// Initialize the GraphService and save it in the window object.
			this.sdkHelper = new GraphSdkHelper();
			window.sdkHelper = this.sdkHelper;
		}

		this.utils = new Utils();

		this.authInProgress = false;

		// Set the isAuthenticated prop and the (empty) Fabric example selection.
		let isAuth = this.authHelper.isAuthenticated();
		console.log("AppTeams_Constructor call is Authenticated: " + isAuth);


		try {
			/* Initialize the Teams library before any other SDK calls.
			 * Initialize throws if called more than once and hence is wrapped in a try-catch to perform a safe initialization.
			 */
			microsoftTeams.initialize();
		}
		catch (err) {
			console.log(err);
		}
		finally {
			this.state = {
				isTeams: this.utils.inTeams(),
				isAuthenticated: isAuth,
				theme: ThemeStyle.Light,
				fontSize: 16,
				channelName: "",
				channelId: "",
				teamName: "",
				groupId: "",
				contextUpn: ""
			};
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
			channelId: this.getQueryVariable('channelId'),
			fontSize: this.pageFontSize()
		});

		// If you are deploying your site as a MS Teams static or configurable tab, you should add ?theme={theme} to
		// your tabs URL in the manifest. That way you will get the current theme on start up (calling getContext on
		// the MS Teams SDK has a delay and may cause the default theme to flash before the real one is returned).
		this.updateTheme(this.getQueryVariable('theme'));

		microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);

		if (this.state.isAuthenticated) {
			this.authHelper.callGetUserProfile()
				.then(userProfile => {
					console.log("RESPONSE AppBrowser_componentWillMount: " + userProfile.userPrincipalName);
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
        console.log("AppTeams_componentDidUpdate teamname: " + this.state.teamName);
		if (this.state.teamName !== null) {
            if (!this.isTokenHandlerExcluded()) {
                console.log("AppTeams_componentDidUpdate isTokenHandlerExcluded: false");
				this.authUser();
			}
		}

		console.log("DEBUG AppTeams_componentDidUpdate isAuth: " + this.authHelper.isAuthenticated());
	}

	isTokenHandlerExcluded() {
		if (window.location.pathname.substring(0, 7) === "/tabmob" ||
			window.location.pathname.substring(0, 11) === "/tab/config" ||
			window.location.pathname.substring(0, 12) === "/tab/tabauth") {
			return true;
		} else {
			return false;
		}
	}

	// Login user
	authUser() {
		//if (localStorage.getItem("configLoginState") !== "1") {
		console.log("DEBUG AppTeams_authUser start 1");
		if (!this.authInProgress) {
			//localStorage.setItem("configLoginState", "1");
			this.authInProgress = true;
			console.log("DEBUG AppTeams_authUser start 2");
			microsoftTeams.authentication.authenticate({
				url: window.location.origin + "/tab/tabauth",
				width: 670,
				height: 570,
				successCallback: function (result) {
					//getUserProfile(result.accessToken);
                    console.log("DEBUG AppTeams_authUser: " + JSON.stringify(result));
                    // TODO: See if we can detect teams browser so we don't try to set validity state
					microsoftTeams.settings.setValidityState(true);
				},
				failureCallback: function (reason) {
					//handleAuthError(reason);
					console.log("DEBUG AppTeams_authUser failureCallback: " + JSON.stringify(reason));
					if (reason === "FailedToOpenWindow") {
						// Try loginPopup since it may be browser client
						localStorage.setItem("configLoginState", "1");
					}
				}
			});
			console.log("DEBUG AppTeams_authUser start 3");
			if (localStorage.getItem("configLoginState") === "1") {
				console.log("DEBUG AppTeams_authUser configLoginState = 1");
			}
		}
	}

	// Tries to get a token silently
	acquireTokenSilent() {
		console.log("DEBUG AppTeams_acquireTokenSilent inTeams: " + this.state.isTeams + " - isTokenHandlerExcluded: " + this.isTokenHandlerExcluded());
		if (this.state.isTeams) {
			if (this.isTokenHandlerExcluded()) {
				// TBD
			} else {
				console.log("AppTeams_acquireTokenSilent Started ");
				this.authHelper.acquireWebApiTokenSilent()
					.then(() => {
						console.log("AppTeams_acquireTokenSilent_acquireWebApiTokenSilent completed ");
						this.authHelper.callGetUserProfile()
							.then(userProfile => {
								console.log("AppTeams_acquireTokenSilent: " + userProfile.userPrincipalName);
								if (this.state.userProfile.displayName.length === 0 || !this.state.isAuthenticated) {
									this.setState({
										userProfile: userProfile,
										isAuthenticated: true,
										displayName: `Hello ${userProfile.displayName}!`
									});
								}
							});
					})
					.catch((err) => {
						console.log("AppTeams_acquireTokenSilent error: " + err);
						if (err === "user_login_error:User login is required") {
							//this.authUser();
						}
					});
			}
		}
	}

	// Sign the user out of the session.
	logout() {
		this.authHelper.logout().then(() => {
			this.setState({
				isAuthenticated: false,
				displayName: ''
			});
		});
	}

	getTeamsContext() {
		microsoftTeams.getContext(context => {
			if (context) {
				this.setState({
					channelName: context.channelName,
					channelId: context.channelId,
					teamName: context.teamName,
					groupId: context.groupId,
					contextUpn: context.upn,
					validityState: microsoftTeams.settings.validityState
				});
			}

		});
	}

	// Grabs the font size in pixels from the HTML element on your page.
	pageFontSize = () => {
		let sizeStr = window.getComputedStyle(document.getElementsByTagName('html')[0]).getPropertyValue('font-size');
		sizeStr = sizeStr.replace('px', '');
		let fontSize = parseInt(sizeStr, 10);
		if (!fontSize) {
			fontSize = 16;
		}
		return fontSize;
	}

	// Sets the correct theme type from the query string parameter.
	updateTheme = (themeStr) => {
		let theme;
		switch (themeStr) {
			case 'dark':
				theme = ThemeStyle.Dark;
				break;
			case 'contrast':
				theme = ThemeStyle.HighContrast;
				break;
			case 'default':
			default:
				theme = ThemeStyle.Light;
		}
		this.setState({ theme });
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
			<div className="ms-font-m show">
				<Route exact path='/tabmob/rootTab' component={RootTabMob} />
				<Route exact path='/tabmob/proposalStatusTab' component={ProposalStatus} />
                <Route exact path='/tabmob/checklistTab' component={Checklist} />
                <Route exact path='/tabmob/customerDecisionTab' component={CustomerDecision} />
				
				<Route exact path='/tab' component={Home} />
				<Route exact path='/tab/config' component={Config} />
				<Route exact path='/tab/tabauth' component={TabAuth} />
				<Route exact path='/tab/privacy' component={Privacy} />
				<Route exact path='/tab/termsofuse' component={TermsOfUse} />

				<Route exact path='/tab/proposalStatusTab' component={ProposalStatus} />
				<Route exact path='/tab/checklistTab' component={Checklist} />
				<Route exact path='/tab/rootTab' component={RootTab} />
				<Route exact path='/tab/customerDecisionTab' component={CustomerDecision} />
			</div>
		);
	}
}
