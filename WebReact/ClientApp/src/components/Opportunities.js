/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { LinkContainer } from 'react-router-bootstrap';
import { Glyphicon, Nav, Navbar, NavItem } from 'react-bootstrap';
//import { Link } from 'office-ui-fabric-react/lib/Link';
import { Dashboard } from './Opportunity/Dashboard';

import { NewOpportunity } from './Opportunity/NewOpportunity';
import { NewOpportunityDocuments } from './Opportunity/NewOpportunityDocuments';
import { NewOpportunityOthers } from './Opportunity/NewOpportunityOthers';
import { applicationId, redirectUri, graphScopes, resourceUri, webApiScopes } from '../helpers/AppSettings';

import { debug } from 'util';
import { oppStatus, userRoles, oppStatusText, oppStatusClassName } from '../common';
import '../Style.css';


export class Opportunities extends Component {
    displayName = Opportunities.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;

        const userProfile = this.props.userProfile;
        this.filesToUpload = [];

        this.state = {
            messageBarEnabled: false,
            viewState: "dashboard",
            userProfile: userProfile,
            industryList: [],
            regionList: [],
            categoryList: [],
            dashboardList: [],
            teamMembers: [],
            reverseList: false,
            loading: true
        };

        this.onClickCreateOpp = this.onClickCreateOpp.bind(this);
        this.onClickOppCancel = this.onClickOppCancel.bind(this);
        this.onClickOppBack = this.onClickOppBack.bind(this);
    }

    componentWillMount() {
        // NOTE: We could just set this directly on state so during the session no calls are made to MT although volume does not justify for now
        this.getOpportunityIndex()
            .then(data => {
                //TODO: something may replace and use authHelper fetcher
            })
            .catch(err => {
                this.errorHandler(err, "Opportunities_componentWillMount_getOpportunityIndex error:");
            });

        if (this.state.regionList.length === 0) {
            this.getRegions();
        }
        if (this.state.industryList.length === 0) {
            this.getIndustries();
        }
        if (this.state.categoryList.length === 0) {
            this.getCategories();
        }
        if (this.state.teamMembers.length === 0) {
            this.getUserProfiles();
        }
    }

    // Class methods

    getOpportunityIndex() {
        return new Promise((resolve, reject) => {
            // To get the List of Opportunities to Display on Dashboard page
            let requestUrl = 'api/Opportunity?page=1';

            fetch(requestUrl, {
                method: "GET",
                headers: { 'authorization': 'Bearer ' + window.authHelper.getWebApiToken() }
            })
                .then(response => {
                    if (response.ok) {
                        return response.json();
                    } else {
                        this.fetchResponseHandler(response, "getOpportunityIndex");
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
                    if (itemslist.length > 0) {
                        this.setState({ reverseList: true });
                    }

                    let sortedList = this.state.reverseList ? itemslist.reverse() : itemslist;
                    this.setState({
                        loading: false,
                        dashboardList: sortedList
                    });

                    resolve(true);
                })
                .catch(err => {
                    this.errorHandler(err, "getOpportunityIndex");
                    this.setState({
                        loading: false,
                        dashboardList: []
                    });
                    reject(err);
                });
        });
    }

    getRegions() {
        // call to API fetch Regions
        let requestUrl = 'api/Region';
        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + window.authHelper.getWebApiToken() }
        })
            .then(response => response.json())
            .then(data => {
                try {
                    let regionList = [];
                    for (let i = 0; i < data.length; i++) {
                        let region = {};
                        region.key = data[i].id;
                        region.text = data[i].name;
                        regionList.push(region);
                    }
                    this.setState({ regionList: regionList });
                }
                catch (err) {
                    return false;
                }

            });
    }

    getIndustries() {
        // call to API fetch Industry
        let requestUrl = 'api/Industry';
        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
        })
            .then(response => response.json())
            .then(data => {
                try {
                    let industryList = [];
                    for (let i = 0; i < data.length; i++) {
                        let industry = {};
                        industry.key = data[i].id;
                        industry.text = data[i].name;
                        industryList.push(industry);
                    }
                    this.setState({ industryList: industryList });
                }
                catch (err) {
                    return false;
                }

            });
    }

    getCategories() {
        // call to API fetch Categories
        let requestUrl = 'api/Category';
        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
        })
            .then(response => response.json())
            .then(data => {
                try {
                    let categoryList = [];
                    for (let i = 0; i < data.length; i++) {
                        let category = {};
                        category.key = data[i].id;
                        category.text = data[i].name;
                        categoryList.push(category);
                    }
                    this.setState({ categoryList: categoryList });
                }
                catch (err) {
                    return false;
                }

            });
    }

    getUserProfiles() {
        let requestUrl = 'api/UserProfile/';
        fetch(requestUrl, {
            method: "GET",
            headers: {
                'authorization': 'Bearer ' + this.authHelper.getWebApiToken()
            }
        })
            .then(response => {
                if (response.ok) {
                    return response.json();
                } else {
                    this.fetchResponseHandler(response, "getUserProfiles");
                    return [];
                }
            })
            .then(data => {
                let itemslist = [];

                if (data.ItemsList.length > 0) {
                    for (let i = 0; i < data.ItemsList.length; i++) {

                        let item = data.ItemsList[i];

                        let newItem = {};

                        newItem.id = item.id;
                        newItem.displayName = item.displayName;
                        newItem.mail = item.mail;
                        newItem.userPrincipalName = item.userPrincipalName;
                        newItem.userRoles = item.userRoles;

                        itemslist.push(newItem);
                    }
                }

                this.setState({
                    teamMembers: itemslist
                });
            })
            .catch(err => {
                console.log("Opportunities_getUserProfiles error: " + JSON.stringify(err));
            });
    }

    fetchResponseHandler(response, referenceCall) {
        if (response.status === 401) {
            // Handle refresh token
        }

    }

    errorHandler(err, referenceCall) {
        console.log("Opportunities Ref: " + referenceCall + " error: " + JSON.stringify(err));
    }


    onClickCreateOpp() {
        this.newOpportunity = {
            id: "",
            displayName: "",
            customer: {
                id: "",
                displayName: "",
                referenceId: ""
            },
            dealSize: 0,
            annualRevenue: 0,
            openedDate: new Date(),
            industry: {
                id: "",
                name: ""
            },
            region: {
                id: "",
                name: ""
            },
            margin: 0,
            rate: 0,
            debtRatio: 0,
            purpose: "",
            disbursementSchedule: "",
            collateralAmount: 0,
            guarantees: "",
            riskRating: 0,
            teamMembers: [
                {
                    status: 0,
                    id: this.state.userProfile.id,
                    displayName: this.state.userProfile.displayName,
                    mail: this.state.userProfile.mail,
                    userPrincipalName: this.state.userProfile.userPrincipalName,
                    userRoles: this.state.userProfile.roles,
                    assignedRole: this.state.userProfile.roles.filter(x => x.displayName === "RelationshipManager")[0]
                }
            ],
            notes: [],
            documentAttachments: []
        };

        this.setState({
            viewState: "createStep1"
        });
    }

    onClickOppCancel() {
        this.setState({
            viewState: "dashboard"
        });
    }

    onClickOppBack() {
        if (this.state.viewState === "createStep1") {
            this.setState({
                viewState: "dashboard"
            });

        } else if (this.state.viewState === "createStep2") {
            this.setState({
                viewState: "createStep1"
            });

        } else if (this.state.viewState === "createStep3") {
            this.setState({
                viewState: "createStep2"
            });

        } else {
            this.setState({
                viewState: "dashboard"
            });
        }
    }

    onClickCreateOppNext() {

        if (this.state.viewState === "createStep1") {
            this.setState({
                viewState: "createStep2"
            });

        } else if (this.state.viewState === "createStep2") {
            this.setState({
                viewState: "createStep3"
            });

        } else if (this.state.viewState === "createStep3") {
            this.setState({
                viewState: "dashboard"
            });

            // Save data
            this.setMessageBar(true, "Saving opportunity data...", MessageBarType.info);
            this.createOpportunity()
                .then(res => {
                    this.setMessageBar(true, "Uploading files...", MessageBarType.info);
                    this.uploadFiles()
                        .then(res => {
                            this.setMessageBar(false, "", MessageBarType.info);
                            this.setState({
                                loading: true
                            });
                            this.getOpportunityIndex()
                                .then(data => {
                                    this.setMessageBar(false, "", MessageBarType.info);
                                })
                                .catch(err => {
                                    this.setMessageBar(false, "", MessageBarType.info);
                                    this.errorHandler(err, "Opportunities_onClickCreateOppNext_getOpportunityIndex");
                                });
                        })
                        .catch(err => {
                            this.setMessageBar(false, "", MessageBarType.info); // TODO: Set error message with timer
                            this.errorHandler(err, "Opportunities_onClickCreateOppNext_uploadFiles");
                            this.setState({
                                loading: true
                            });
                            this.getOpportunityIndex()
                                .then(data => {
                                    this.setMessageBar(false, "", MessageBarType.info);
                                })
                                .catch(err => {
                                    this.errorHandler(err, "Opportunities_onClickCreateOppNext_getOpportunityIndex");
                                });
                        });
                })
                .catch(err => {
                    this.errorHandler(err, "Opportunities_onClickCreateOppNext_createOpportunity");
                });

        } else {
            this.setState({
                viewState: "dashboard"
            });
        }
    }

    setMessageBar(enabled, text, type) {
        this.setState({
            messageBarEnabled: enabled,
            messageBarText: text,
            messageBarType: type

        });
    }

    // Create New Opportunity
    createOpportunity() {
        return new Promise((resolve, reject) => {
            // Clean attachments prior to submit then put them back so upload has the actual file to upload
            let currentAttchments = [];
            this.filesToUpload = currentAttchments.concat(this.newOpportunity.documentAttachments);

            let cleanAttachments = [];
            for (let i = 0; i < this.filesToUpload.length; i++) {
                cleanAttachments.push({
                    id: this.filesToUpload[i].id,
                    fileName: this.filesToUpload[i].file.name,
                    note: this.filesToUpload[i].note,
                    category: {
                        id: this.filesToUpload[i].category.id,
                        displayName: this.filesToUpload[i].category.name
                    },
                    tags: this.filesToUpload[i].tags,
                    documentUri: ""
                });
            }

            this.newOpportunity.documentAttachments = cleanAttachments;

            let requestUrl = 'api/opportunity/';

            let options = {
                method: "POST",
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'authorization': 'Bearer    ' + this.authHelper.getWebApiToken()
                },
                body: JSON.stringify(this.newOpportunity)
                //body: this.newOpportunity
            };

            fetch(requestUrl, options)
                .then(response => this.fetchResponseHandler(response, "createOpportunity"))
                .then(data => {
                    resolve(data);
                })
                .catch(err => {
                    this.setMessageBar(true, "Error saving opportunity data", MessageBarType.error);
                    reject(err);
                });
        });

    }

    // Upload files
    uploadFiles() {
        return new Promise((resolve, reject) => {

            let files = this.filesToUpload;
            for (let i = 0; i < files.length; i++) {
                this.setMessageBar(true, "Uploading files " + (i + 1) + "/" + this.filesToUpload.length, MessageBarType.info);
                let fd = new FormData();
                fd.append('opportunity', "NewOpportunity");
                fd.append('file', files[i].file);
                fd.append('opportunityName', this.newOpportunity.displayName);
                fd.append('fileName', files[i].file.name);

                let requestUrl = 'api/document/UploadFile/' + encodeURIComponent(this.newOpportunity.displayName) + "/Attachment";

                let options = {
                    method: "PUT",
                    headers: {
                        'authorization': 'Bearer    ' + this.authHelper.getWebApiToken()
                    },
                    body: fd
                };

                fetch(requestUrl, options)
                    .then(response => this.fetchResponseHandler(response, "uploadFile"))
                    .then(data => {
                        resolve(data);
                    })
                    .catch(err => {
                        reject(false);
                    });
            }
        });
    }

    render() {
        const viewState = this.state.viewState;
        const isLoading = this.state.loading;

        const userProfile = this.state.userProfile;

        return (
            <div>
                {
                    isLoading ?
                        <div>
                            <br /><br /><br />
                            <Spinner size={SpinnerSize.medium} label='Loading opportunities...' ariaLive='assertive' />
                        </div>
                        :
                        <div>
                            {
                                this.state.messageBarEnabled ?
                                    <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                        <MessageBar messageBarType={this.state.messageBarType} isMultiline={false}>
                                            {this.state.messageBarText}
                                        </MessageBar>
                                    </div>
                                    : ""
                            }

                            {
                                viewState === "dashboard" &&
                                <Dashboard
                                    userProfile={userProfile}
                                    dashboardList={this.state.dashboardList}
                                    onClickCreateOpp={this.onClickCreateOpp}
                                />
                            }

                            {
                                viewState === "createStep1" &&
                                <NewOpportunity
                                    userProfile={userProfile}
                                    opportunity={this.newOpportunity}
                                    regions={this.state.regionList}
                                    industries={this.state.industryList}
                                    dashboardList={this.state.dashboardList}
                                    onClickCancel={this.onClickOppCancel}
                                    onClickNext={this.onClickCreateOppNext.bind(this, this.newOpportunity)}
                                />
                            }

                            {
                                viewState === "createStep2" &&
                                <NewOpportunityDocuments
                                    userProfile={userProfile}
                                    opportunity={this.newOpportunity}
                                    categories={this.state.categoryList}
                                    onClickBack={this.onClickOppBack}
                                    onClickNext={this.onClickCreateOppNext.bind(this, this.newOpportunity)}
                                />
                            }

                            {
                                viewState === "createStep3" &&
                                <NewOpportunityOthers
                                    userProfile={userProfile}
                                    opportunity={this.newOpportunity}
                                    teamMembers={this.state.teamMembers}
                                    onClickBack={this.onClickOppBack}
                                    onClickNext={this.onClickCreateOppNext.bind(this, this.newOpportunity)}
                                />
                            }
                        </div>
                }
            </div>
        );
    }

}