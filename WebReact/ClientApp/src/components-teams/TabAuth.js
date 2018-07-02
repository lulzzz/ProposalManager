/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import Utils from '../helpers/Utils';
import * as microsoftTeams from '@microsoft/teams-js';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export class TabAuth extends Component {
    displayName = TabAuth.name

    constructor(props) {
        super(props);

        console.log("Constructor TabAuth");

        this.authHelper = window.authHelper;
        this.sdkHelper = window.sdkHelper;
        this.utils = new Utils();

        try {
            microsoftTeams.initialize();
        }
        catch (err) {
            console.log("ProposalManagement_ConfigTAB error initializing teams: ");
            console.log(err);
        }
        finally {
            this.state = {
                isAuthenticated: this.authHelper.isAuthenticated(),
                channelName: "",
                channelId: "",
                teamName: "",
                groupId: "",
                contextUpn: ""
            };

            /** Pass the Context interface to the initialize function below */
            //microsoftTeams.getContext(context => this.initialize(context));
        }

    }

    componentWillMount() {
        // Get the teams context
        if (this.state.teamName.length === 0) {
            this.getTeamsContext();
        }

        if (window.location.pathname.substring(0, 7) === "/tabmob") {
            console.log("TabAuth componentWillMount in location: " + window.location.pathname);
            this.acquireTokenSilentParam();
        }
    }

    componentDidUpdate() {
        if (this.state.teamName.length > 0) {
            console.log("TabAuth componentDidUpdate in context of upn: " + this.state.contextUpn);
            this.acquireTokenSilentParam();
        }
        //TODO: Handle popup error
        //msal.error  popup_window_error  msal.error.description
    }

    getTeamsContext() {
        microsoftTeams.getContext(context => {
            this.setState({
                channelName: context.channelName,
                channelId: context.channelId,
                teamName: context.teamName,
                groupId: context.groupId,
                contextUpn: context.upn
            });
        });
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


    // Tries to get a token silently
    acquireTokenSilentParam() {
        console.log("TabAuth acquireTokenSilentParam login_hint: " + this.state.contextUpn);

        //let extraQueryParameters = "login_hint=" + this.state.contextUpn;
        let extraQueryParameters = {
            login_hint: this.state.contextUpn
        };

        this.authHelper.acquireWebApiTokenSilentParam(extraQueryParameters)
            .then(res => {
                console.log("TabAuth acquireTokenSilentParam_acquireTokenSilentParam_then: " + JSON.stringify(res));
                this.notifySuccess();
            })
            .catch(err => {
                console.log("TabAuth_acquireTokenSilentParam error: ");
                console.log(err);
                this.authHelper.loginRedirect();
                console.log("TabAuth_Login_loginRedirect");
                //this.notifySuccess();
            });
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

    notifySuccess() {
        microsoftTeams.authentication.notifySuccess();
    }


    render() {

        return (
            <div className="BgConfigImage ">
                <h2 className='font-white text-center darkoverlay'>Proposal Manager</h2>
                <div className='ms-Grid-row'>
                    <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 mt50 mb50 text-center'>
                <div className='TabAuthLoader'>
                    <Spinner size={SpinnerSize.large} label='Loading your experience...' ariaLive='assertive' />
                    </div>
                    </div>
                </div>

                <div className='ms-Grid-row mt50'>
                <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12  text-center'>
                        <PrimaryButton className='' onClick={this.logout.bind(this)}>
                            Reset Token
                        </PrimaryButton>
              
                        <PrimaryButton className='ml10 backbutton ' onClick={this.notifySuccess.bind(this)}>
                            Close
                        </PrimaryButton>
                    </div>
                </div>
            </div>
        );
    }
}
