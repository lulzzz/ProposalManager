/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// Global imports
import React, { Component } from 'react';
import Promise from 'promise';
import AuthHelper from './helpers/AuthHelper';
import GraphSdkHelper from './helpers/GraphSdkHelper';
import Utils from './helpers/Utils';
import { Route } from 'react-router';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Image } from 'office-ui-fabric-react/lib/Image';

import { AppSettings } from './helpers/AppSettings';
import { Layout } from './components/Layout';
import { Opportunities } from './components/Opportunities';
//import { Notifications } from './components/Notifications';
import { Administration } from './components/Administration/Administration';
import { Settings } from './components/Administration/Settings';

import { OpportunitySummary } from './components/Opportunity/OpportunitySummary';
import { OpportunityNotes } from './components/Opportunity/OpportunityNotes';
import { OpportunityStatus } from './components/Opportunity/OpportunityStatus';
import { OpportunityChooseTeam } from './components/OpportunityChooseTeam';

// compoents-mobile
import { getQueryVariable } from './common';


export class AppBrowser extends Component {
	displayName = AppBrowser.name

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
			// Initialize the GraphService and save it in the window object.
			this.sdkHelper = new GraphSdkHelper();
			window.sdkHelper = this.sdkHelper;
		}

		this.utils = new Utils();

		const userProfile = { id: "", displayName: "", mail: "", phone: "", picture: "", userPrincipalName: "", roles: [] };

		this.state = {
			isAuthenticated: false,
			userProfile: userProfile,
			isLoading: false
		};
	}

	componentWillMount() {
		if (!this.utils.inTeams() && !this.utils.iframed() && this.authHelper.getAuthRedirectState() !== "start") {
			if (this.authHelper.isAuthenticated()) {
				if (!this.state.isAuthenticated) {
					this.authHelper.callGetUserProfile()
						.then(userProfile => {
							this.setState({
								userProfile: userProfile,
								isAuthenticated: true,
								displayName: `Hello ${userProfile.displayName}!`
							});
						});
				}
			}
		}
	}

	componentDidMount() {
		this.acquireTokenSilent();
	}

	// Login user
	loginPrev() {
        this.authHelper.loginPopup()
            .then(() => {
                this.setState({
                    isLoading: true
                });
                this.authHelper.acquireTokenSilent()
                    .then(() => {
                        this.authHelper.acquireWebApiTokenSilent()
                            .then(res => {
                                this.authHelper.callGetUserProfile()
                                    .then(userProfile => {
                                        this.setState({
                                            userProfile: userProfile,
                                            isAuthenticated: true,
                                            displayName: `Hello ${userProfile.displayName}!`
                                        });
                                    });
                            });
                    });
            });
	}


	login() {
		let extraQueryParameters = {
			login_hint: this.state.contextUpn
		};
		this.setState({
			isLoading: true
		});

        this.authHelper.acquireWebApiTokenSilentParam(extraQueryParameters)
            .then(res => {
                this.authHelper.acquireTokenSilentParam(extraQueryParameters)
                    .then(res => {
                        this.authHelper.callGetUserProfile()
                            .then(userProfile => {
                                this.setState({
                                    //isLoading: false,
                                    userProfile: userProfile,
                                    isAuthenticated: true,
                                    displayName: `Hello ${userProfile.displayName}!`
                                });
                            });
                    });
				
			})
			.catch(err => {
				this.authHelper.loginPopup()
					.then(res => {
						this.authHelper.acquireTokenSilent()
							.then(res => {
								this.authHelper.acquireWebApiTokenSilent()
									.then(res => {
										this.authHelper.callGetUserProfile()
											.then(userProfile => {
												this.setState({
													//isLoading: false,
													userProfile: userProfile,
													isAuthenticated: true,
													displayName: `Hello ${userProfile.displayName}!`
												});
											});
									});
							});
					});
			});
	}

	// Tries to get a token silently
	acquireTokenSilent() {
		if (!this.utils.inTeams() && !this.utils.iframed() && this.authHelper.getAuthRedirectState() !== "start") {
			this.authHelper.acquireWebApiTokenSilent()
				.then(() => {
					this.authHelper.callGetUserProfile()
						.then(userProfile => {
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
					if (err === "user_login_error:User login is required") {
						this.login();
					}
				});
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


	render() {
		const userProfileData = this.state.userProfile;
		const userDisplayName = this.state.displayName;
		const isAuthenticated = this.state.isAuthenticated;

		const isLoading = this.state.isLoading;

		//get opportunity details
		const oppId = getQueryVariable('opportunityId') ? getQueryVariable('opportunityId') : "";

		//Inject props to components
		const OpportunitiesView = ({ match }) => {
			return <Opportunities userProfile={userProfileData} />;
		};

		//const NotificationsView = ({ match }) => {
			//return <Notifications userProfile={userProfileData} />;
		//};

		const AdministrationView = ({ match }) => {
			return <Administration userProfile={userProfileData} />;
        };

        const SettingsView = ({ match }) => {
            return <Settings userProfile={userProfileData} />;
        };

		const Summary = ({ match }) => {
			return <OpportunitySummary userProfile={userProfileData} opportunityId={oppId} />;
		};

		const Notes = ({ match }) => {
			return <OpportunityNotes userProfile={userProfileData} opportunityId={oppId} />;
		};

		const Status = ({ match }) => {
			return <OpportunityStatus userProfile={userProfileData} opportunityId={oppId} />;
		};

		const ChooseTeam = ({ match }) => {
			return <OpportunityChooseTeam opportunityId={oppId} />;
		};

		// Route formatting:
		// <Route path="/greeting/:name" render={(props) => <Greeting text="Hello, " {...props} />} />

		return (
			<div>
                <CommandBar farItems={
                    [

                        {

                            key: 'display-name',
                            name: userDisplayName
                        },
						{
							key: 'log-in-out=button',
							name: this.state.isAuthenticated ? 'Sign out' : 'Sign in',
							onClick: this.state.isAuthenticated ? this.logout.bind(this) : this.login.bind(this)
						}
                    ]
                }
                />
				
				<div className="ms-font-m show">
					{
						isAuthenticated ?
                            <Layout userProfile={userProfileData}>
								<Route exact path='/' component={OpportunitiesView} />
								
                                <Route exact path='/Administration' component={AdministrationView} />
                                <Route exact path='/Settings' component={SettingsView} />

								<Route exact path='/OpportunitySummary' component={Summary} />
								<Route exact path='/OpportunityNotes' component={Notes} />
								<Route exact path='/OpportunityStatus' component={Status} />
								<Route exact path='/OpportunityChooseTeam' component={ChooseTeam} />
							</Layout>
							:
							<div className="BgImage">
								<div className="Caption">
									<h3> <span> EMPOWER </span> Banking    </h3>
									<h2> Proposal Manager</h2>
								</div>
								{
									isLoading &&
									<div className='Loading-spinner'>
                                        <Spinner className="Homelaoder Homespinnner" size={SpinnerSize.medium} label='Loading your experience...' ariaLive='assertive' />
									</div>
								}
							</div>
					}
				</div>
			</div>
		);
	}
}
