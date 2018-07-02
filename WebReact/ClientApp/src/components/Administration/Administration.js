/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import { OpportunityListCompact } from '../Opportunity/OpportunityListCompact';
import { oppStatusClassName } from '../../common';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import Utils from '../../helpers/Utils';
import '../../Style.css';

export class Administration extends Component {
    displayName = Administration.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;
        this.utils = new Utils();

        this.state = {
            userProfile: this.props.userProfile,
            loading: true,
            refreshing: false,
            items: [],
            itemsOriginal: [],
            userRoleList: [],
            channelCounter: 0
        };
    }

    componentWillMount() {
        this.acquireGraphAdminTokenSilent(); // Call acquire token so it is ready when calling graph using admin token

        if (this.state.itemsOriginal.length === 0) {
            console.log("Administration_componentWillMount getOpportunityIndex");
            this.getOpportunityIndex()
                .then(data => {
                    console.log("Administration_componentWillMount getUserRoles");
                    this.getUserRoles()
                        .then(res => {
                            console.log("Administration_componentWillMount getUserRoles done" + res);
                            this.setState({
                                loading: false
                            });
                        })
                        .catch(err => {
                            // TODO: Add error message
                            this.errorHandler(err, "Administration_componentWillMount_getUserRoles");
                        });
                })
                .catch(err => {
                    // TODO: Add error message
                    this.errorHandler(err, "Administration_componentWillMount_getOpportunityIndex");
                });
        }
    }

    fetchResponseHandler(response, referenceCall) {
        if (response.status === 401) {
            //TODO: Handle refresh token in vNext;
        }
    }

    errorHandler(err, referenceCall) {
        console.log("Administration Ref: " + referenceCall + " error: " + JSON.stringify(err));
    }

    acquireGraphAdminTokenSilent() {
        if (this.utils.getQueryVariable("adminconsent")) {
            let isAdmin = this.state.userProfile.roles.filter(x => x.displayName === "Administrator");
            if (isAdmin) {
                this.authHelper.loginPopupGraphAdmin()
                    .then(access_token => {
                        // TODO: For future expansion sice the toke has been handled by authHelper
                    })
                    .catch(err => {
                        console.log(err);
                        this.errorHandler(err, "Administration_acquireGraphAdminTokenSilent");
                    });
            }
        } else {
            let isAdmin = this.state.userProfile.roles.filter(x => x.displayName === "Administrator");
            if (isAdmin) {
                this.authHelper.acquireGraphAdminTokenSilent()
                    .then(access_token => {
                        // TODO: For future expansion sice the toke has been handled by authHelper
                    })
                    .catch(err => {
                        console.log(err);
                        this.errorHandler(err, "Administration_acquireGraphAdminTokenSilent");
                        //this.showMessageBar("Error while requesting an admin token for Graph API, please try refreshing your browser and sign-in again.", MessageBarType.error);
                    });
            }
        }
    }

    getOpportunityIndex() {
        return new Promise((resolve, reject) => {
            // To get the List of Opportunities to Display on Dashboard page
            let requestUrl = 'api/Opportunity?page=1';

            fetch(requestUrl, {
                method: "GET",
                headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
            })
                .then(response => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        this.fetchResponseHandler(response, "Administration_getOpportunityIndex");
                        reject(response);
                    }
                })
                .then(data => {
                    let itemslist = [];
                    
                    if (data.ItemsList.length > 0) {
                        for (let i = 0; i < data.ItemsList.length; i++) {

                            let item = data.ItemsList[i];

                            let newItem = {};

                            newItem.id = item.id;
                            newItem.opportunity = item.displayName;
                            newItem.client = item.customer.displayName;
                            newItem.dealsize = item.dealSize;
                            newItem.openedDate = new Date(item.openedDate).toLocaleDateString();
                            newItem.stausValue = item.opportunityState;
                            newItem.status = oppStatusClassName[item.opportunityState];
                            itemslist.push(newItem);
                        }
                    }

                    let filteredItems = itemslist.filter(itm => itm.stausValue < 2);


                    this.setState({
                        items: filteredItems,
                        itemsOriginal: itemslist
                    });

                    resolve(true);
                })
                .catch(err => {
                    this.errorHandler(err, "Administration_getOpportunityIndex");
                    this.setState({
                        loading: false,
                        items: [],
                        itemsOriginal: []
                    });
                    reject(err);
                });
        });
    }

    getOpportunity(oppId) {
        return new Promise((resolve, reject) => {
            // To get the List of Opportunities to Display on Dashboard page
            this.requestUrl = 'api/Opportunity?id=' + oppId;

            fetch(this.requestUrl, {
                method: "GET",
                headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
            })
                .then(response => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        this.fetchResponseHandler(response, "Administration_getOpportunityIndex");
                        reject(response);
                    }
                })
                .then(data => {
                    resolve(data);
                })
                .catch(err => {
                    this.errorHandler(err, "Administration_getOpportunityIndex");
                    reject(err);
                });

        });
    }

    getUserRoles() {
        // call to API fetch data
        return new Promise((resolve, reject) => {
            let requestUrl = 'api/RoleMapping';
            fetch(requestUrl, {
                method: "GET",
                headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
            })
                .then(response => response.json())
                .then(data => {
                    try {
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
                        console.log("Administration_getUserRoles userRoleList lenght: " + userRoleList.length);
                        resolve(true);
                    }
                    catch (err) {
                        reject(err);
                    }

                });
        });
    }

    updateOpportunity(opportunity) {
        return new Promise((resolve, reject) => {
            let requestUrl = 'api/opportunity';

            var options = {
                method: "PATCH",
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'authorization': 'Bearer    ' + this.authHelper.getWebApiToken()
                },
                body: JSON.stringify(opportunity)
            };

            fetch(requestUrl, options)
                .then(response => this.fetchResponseHandler(response, "Administration_updateOpportunity_fetch"))
                .then(data => {
                    resolve(data);
                })
                .catch(err => {
                    this.errorHandler(err, "Administration_updateOpportunity");
                    reject(err);
                });
        });

    }

    chngeOpportunityState(id) {
        console.log("Administration_chngeOpportunityState timer for: " + id);
    }

    createTeam(teamName) {
        return new Promise((resolve, reject) => {
            let grpDisplayName = teamName;
            let grpDescription = "This is the team group for " + grpDisplayName;

            this.sdkHelper.createTeamGroup(grpDisplayName, grpDescription)
                .then((res) => {
                    resolve(res, null);
                })
                .catch(err => {
                    console.log(err);
                    resolve(null, err);
                });
        });
    }

    createChannel(teamId, name, description) {
        return new Promise((resolve, reject) => {
            this.sdkHelper.createChannel(name, description, teamId)
                .then((res) => {
                    resolve(res, null);
                })
                .catch(err => {
                    console.log("Administration_createChannel error: ");
                    console.log(err);
                    resolve(null, err);
                });
        });
    }

    createNextChannel(teamId, item) {
        let channelCounter = this.state.channelCounter;
        const roleMappings = this.state.userRoleList;
        console.log("Administration_createNextChannel start channelCounter: " + channelCounter);
        if (roleMappings.length > channelCounter) {
            let channelName = roleMappings[channelCounter].channel;

            if (channelName !== "General" && channelName !== "None") {
                this.createChannel(teamId, channelName, channelName + " channel")
                    .then((res, err) => {
                        console.log("Administration_createNextChannel channelCounter: " + channelCounter + " lenght: " + roleMappings.length);
                        this.setState({ channelCounter: channelCounter + 1 });
                        this.createNextChannel(teamId, item);
                    })
                    .catch(err => {
                        this.errorHandler(err, "createNextChannel_createChannel: " + channelName);
                    });
            } else {
                this.setState({ channelCounter: channelCounter + 1 });
                this.createNextChannel(teamId, item);
            }
        } else {
            //this.setState({ channelCounter: 0 });
            console.log("Administration_createNextChannel finished channelCounter: " + channelCounter);
            this.showMessageBar("Updating opportunity state and moving files to team for " + item.opportunity + ", please do not close or browse to other items until creation process is complete.", MessageBarType.warning);
            setTimeout(this.chngeOpportunityState, 4000, item.id);
            this.getOpportunity(item.id)
                .then(res => {
                    res.opportunityState = 2;
                    this.updateOpportunity(res)
                        .then(res => {
                            this.hideMessageBar();
                            this.setState({
                                loading: true
                            });
                            setTimeout(this.chngeOpportunityState, 2000, item.id);
                            this.getOpportunityIndex()
                                .then(data => {
                                    console.log("Administration_createNextChannel finished after getOpportunityIndex channelCounter: " + channelCounter);
                                    console.log("Administration_createNextChannel Adding team app to teamId: " + teamId);
                                    this.sdkHelper.addAppToTeam(teamId)
                                        .then(res => {
                                            this.setState({
                                                loading: false,
                                                channelCounter: 0
                                            });
                                        })
                                        .catch(err => {
                                            // TODO: Add error message
                                            this.errorHandler(err, "Administration_createNextChannel_addAppToTeam");
                                        });
                                })
                                .catch(err => {
                                    // TODO: Add error message
                                    this.errorHandler(err, "Administration_createNextChannel_getOpportunityIndex");
                                });
                        })
                        .catch(err => {
                            this.showMessageBar("There was a problem trying to update the opportunity, please try again in a few minutes.", MessageBarType.error);
                            this.errorHandler(err, "createNextChannel_updateOpportunity");
                        });
                })
                .catch(err => {
                    this.showMessageBar("There was a problem trying to get the opportunity, please try again.", MessageBarType.error);
                    this.errorHandler(err, "createNextChannel_getOpportunity");
                });
        }
    }

    showMessageBar(text, messageBarType) {
        this.setState({
            result: {
                type: messageBarType,
                text: text
            }
        });
        // MessageBar types:
        // MessageBarType.error
        // MessageBarType.info
        // MessageBarType.severeWarning
        // MessageBarType.success
        // MessageBarType.warning
    }

    hideMessageBar() {
        this.setState({
            result: null
        });
    }


    //Event handlers

    onActionItemClickOld(item) {
        if (this.state.items.length > 0) {
            this.showMessageBar("Creating team and channels for " + item.opportunity + ", please do not close or browse to other items until creation process is complete.", MessageBarType.warning);
            this.createTeam(item.opportunity)
                .then((res, err) => {
                    let teamId = res;
                    if (err) {
                        // Try to get teamId if error is due to existing team
                    }

                    this.createChannel(teamId, "Risk Assessment", "Risk Assessment channel")
                        .then((res, err) => {
                            this.createChannel(teamId, "Credit Check", "Credit Check channel")
                                .then((res, err) => {
                                    this.createChannel(teamId, "Compliance", "Compliance channel")
                                        .then((res, err) => {
                                            this.createChannel(teamId, "Formal Proposal", "Formal Proposal channel")
                                                .then((res, err) => {
                                                    this.createChannel(teamId, "Customer Decision", "Customer Decision channel")
                                                        .then((res, err) => {
                                                            this.showMessageBar("Updating opportunity state and moving files to team for " + item.opportunity + ", please do not close or browse to other items until creation process is complete.", MessageBarType.warning);
                                                            setTimeout(this.chngeOpportunityState, 4000, item.id);
                                                            this.getOpportunity(item.id)
                                                                .then(res => {
                                                                    res.opportunityState = 2;
                                                                    this.updateOpportunity(res)
                                                                        .then(res => {
                                                                            this.hideMessageBar();
                                                                            this.setState({
                                                                                loading: true
                                                                            });
                                                                            setTimeout(this.chngeOpportunityState, 2000, item.id);
                                                                            this.getOpportunityIndex()
                                                                                .then(data => {
                                                                                    this.setState({
                                                                                        loading: false
                                                                                    });
                                                                                })
                                                                                .catch(err => {
                                                                                    // TODO: Add error message
                                                                                    this.errorHandler(err, "Administration_chngeOpportunityState_getOpportunityIndex");
                                                                                });
                                                                        })
                                                                        .catch(err => {
                                                                            this.showMessageBar("There was a problem trying to update the opportunity, please try again in a few minutes.", MessageBarType.error);
                                                                            this.errorHandler(err, "onActionItemClick_updateOpportunity");
                                                                        });
                                                                })
                                                                .catch(err => {
                                                                    this.showMessageBar("There was a problem trying to get the opportunity, please try again.", MessageBarType.error);
                                                                    this.errorHandler(err, "onActionItemClick_getOpportunity");
                                                                });
                                                        })
                                                        .catch(err => {
                                                            this.errorHandler(err, "onActionItemClick_createChannel: Customer Decision");
                                                        });
                                                })
                                                .catch(err => {
                                                    this.errorHandler(err, "onActionItemClick_createChannel: Formal Proposal");
                                                });
                                        })
                                        .catch(err => {
                                            this.errorHandler(err, "onActionItemClick_createChannel: Compliance");
                                        });
                                })
                                .catch(err => {
                                    this.errorHandler(err, "onActionItemClick_createChannel: Credit Check");
                                });
                        })
                        .catch(err => {
                            this.errorHandler(err, "onActionItemClick_createChannel:Risk Assessment ");
                        });
                })
                .catch(err => {
                    this.errorHandler(err, "onActionItemClick_createTeam");
                });
        }
    }

    onActionItemClick(item) {
        if (this.state.items.length > 0) {
            this.showMessageBar("Creating team and channels for " + item.opportunity + ", please do not close or browse to other items until creation process is complete.", MessageBarType.warning);
            this.createTeam(item.opportunity)
                .then((res, err) => {
                    let teamId = res;
                    if (err) {
                        // Try to get teamId if error is due to existing team
                    }
                    console.log("onActionItemClick_createTeam start channel creation");
                    this.createNextChannel(teamId, item);
                })
                .catch(err => {
                    this.errorHandler(err, "onActionItemClick_createTeam");
                });
        }
    }

    render() {
        const items = this.state.items;

        let isAdmin = false;
        if (this.state.userProfile.roles.filter(x => x.displayName === "Administrator").length > 0) {
            isAdmin = true;
        }

        return (
            <div className='ms-Grid'>
                <div className='ms-Grid-row'>
                    <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6 pageheading'>
                        <h3>Manage Team Creation</h3>
                    </div>
                </div>
                {
                    this.state.result &&
                    <MessageBar
                        messageBarType={this.state.result.type}
                        onDismiss={this.hideMessageBar.bind(this)}
                        isMultiline={false}
                        dismissButtonAriaLabel='Close'
                    >
                        {this.state.result.text}
                    </MessageBar>
                }

                {
                    isAdmin ?
                        this.state.loading ?
                            <div>
                                <br /><br /><br />
                                <Spinner size={SpinnerSize.medium} label='Loading opportunities...' ariaLive='assertive' />
                            </div>
                            :
                            items.length > 0 ?
                                <OpportunityListCompact opportunityIndex={items} onActionItemClick={this.onActionItemClick.bind(this)} />
                                :
                                <div>There are no opportunities with status 'Creating'.</div>
                        :
                        <div>
                            <br /><br /><br />
                            <h3>You must be an administrator to access this functionality</h3>
                        </div>
                }

                <div className='ms-grid-row'>
                    <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 pl0'><br />

                    </div>
                    <div className='ms-Grid-col ms-sm12 ms-md6 ms-lg6'><br />
                        {
                            this.state.refreshing ?
                                <div className='ms-Grid-col ms-sm12 ms-md3 ms-lg6 pull-right'>
                                    <Spinner size={SpinnerSize.small} label='Loading opportunities...' ariaLive='assertive' />
                                </div>
                                :
                                <br />
                        }
                        <br /><br /><br />
                    </div>
                </div>
            </div>
        );
    }
}