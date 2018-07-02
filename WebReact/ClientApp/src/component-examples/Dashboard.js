/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import * as ReactDOM from 'react-dom';
import GraphSdkHelper from '../helpers/GraphSdkHelper';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
//List
import { getRTL } from 'office-ui-fabric-react/lib/Utilities';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { List } from 'office-ui-fabric-react/lib/List';

import { Glyphicon, Nav, Navbar, NavItem } from 'react-bootstrap';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { LinkContainer } from 'react-router-bootstrap';
//import { Link } from 'office-ui-fabric-react/lib/Link';
import { Link } from 'react-router-dom';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { NewOpportunity } from '../components/Opportunity/NewOpportunity';
import { NewOpportunityDocuments } from '../components/Opportunity/NewOpportunityDocuments';
import { NewOpportunityOthers } from '../components/Opportunity/NewOpportunityOthers';
import { applicationId, redirectUri, graphScopes, resourceUri, webApiScopes } from '../helpers/AppSettings';

import { debug } from 'util';
import { oppStatus, userRoles, oppStatusText, oppStatusClassName } from '../common';
import { MessageBarButton } from 'office-ui-fabric-react/lib/Button';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import '../../Style.css';


function DashboardView(props) {
    return (
        <div>
            {
                this.props.context.messageBarEnabled ?
                    <MessageBar
                        messageBarType={this.props.context.messageBarType}
                        isMultiline={false}
                    >
                        {this.props.context.messageBarText}
                    </MessageBar>
                    : ""
            }

            <div className='ms-Grid-row'>
                <div className='ms-Grid-col ms-sm6 ms-md4 ms-lg12'>&nbsp;</div>
            </div>
            <div className='ms-Grid-row'>
                <div className='ms-Grid-col ms-sm6 ms-md4 ms-lg6 pageheading'>
                    <h3>Dashboard</h3>
                </div>
                {
                    this.props.context.AuthUserRole === 2 ?
                        <div className='ms-Grid-col ms-sm6 ms-md4 ms-lg6 createButton '>
                            <PrimaryButton className='pull-right' onClick={this.fnCreateStep1}> <i className="ms-Icon ms-Icon--Add pr10" aria-hidden="true"></i> Create New</PrimaryButton>
                        </div>
                        :
                        <div className='ms-Grid-col ms-sm6 ms-md4 ms-lg6 createButton '>
                        </div>

                }
            </div><br />
            <div className='ms-Grid'>
                <div className='ms-Grid-row ms-SearchBoxSmallExample'>
                    <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg3'>
                        <span>Client Name</span>
                        <SearchBox
                            placeholder='Search'
                            onChange={this._onFilterByNameChanged}
                        />
                    </div>
                    <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg3'>
                        <span>Deal Size</span>
                        <SearchBox
                            placeholder='Search'
                            onChange={this._onFilterByDealChanged}
                        />
                    </div>
                </div><br />
                <div className='ms-Grid-row'>
                    <div>
                        {opportunitiesListHeading}
                    </div>
                    {
                        isLoading ?
                            <div>
                                <br /><br /><br />
                                <Spinner size={SpinnerSize.medium} label='Loading opportunities...' ariaLive='assertive' />
                            </div>
                            :
                            listHasItems ?
                                <div>
                                    <br /><br />
                                    {opportunitiesListComponent}
                                </div>
                                :
                                <p><em>No opportunities were found</em></p>
                    }
                </div><br /><br />
            </div>
        </div>
    );
}

export class DashboardOld extends Component {
	displayName = Dashboard.name

	constructor(props) {
		super(props);

		this.sdkHelper = window.sdkHelper;
		this.authHelper = window.authHelper;

        let userProfile = this.props.userProfile;

		this.state = {
			refreshToken: false,
			filterClient: '',
			filterDeal: '',
			items: new Array(),
			itemsOriginal: new Array(),
			loading: true,
			reverseList: false, //Seems there are issues with Reverse function on arrays
			createStep: 0, //'choose    ,
			currentCount: 0,
			OfficersList: [{}],
            AuthUserId: userProfile.id,
            AuthUserDisplayName: userProfile.displayName,
            AuthUserMail: userProfile.mail,
            AuthUserPhone: userProfile.phone,
            AuthUserPicture: userProfile.picture,
            AuthUserUPN: userProfile.userPrincipalName,
            AuthUserRole: userProfile.role,
			file: '',
			OppNameError: false,
			CustNameError: false,
			DealSizeError: false,
			AnnualRevenueError: false,
			messageBarEnabled: false,
			messageBarText: "",
			MessagebarTextOpp: "",
			MessagebarTexCust: "",
			MessagebarTexDealSize: "",
			loadSpinner:true,
			MessagebarTexAnnualRev: "",
			isLoadingOpportunity: false,
			submitPressed:false
		};

		this._onFilterByNameChanged = this._onFilterByNameChanged.bind(this);
        this._onFilterByDealChanged = this._onFilterByDealChanged.bind(this);

		this.fnCreateStep1 = this.fnCreateStep1.bind(this);
		this.fnCreateStep2 = this.fnCreateStep2.bind(this);
		this.fnCreateStep3 = this.fnCreateStep3.bind(this);
        this.fnBackStep = this.fnBackStep.bind(this);

		this.submit = this.submit.bind(this);
		//this.submitComplete = this.submitComplete.bind(this);
		this.uploadFile = this.uploadFile.bind(this);
		this.createOpportunity = this.createOpportunity.bind(this);

        this.getOpportunityIndex();
	}

	login() {
		window.authHelper.loginPopup()
			.then(() => {
				window.authHelper.acquireTokenSilent()
					.then(() => {
						console.log("Dashboard_login_acquireTokenSilent");
						this.forceUpdate();
					});
			});
	}

	fetchResponseHandler(response, referenceCall, reject) {
		console.log("fetchResponseHandler refcall: " + referenceCall + " response: " + response.status + " - " + response.statusText)
		if (response.status === 401) {
			this.setState({
				refreshToken: true
			});
		}
	}

	errorHandler(err, referenceCall) {
		console.log("Dashboard Ref: " + referenceCall + " error: " + JSON.stringify(err));
	}

	getOpportunityIndex() {
		return new Promise((resolve, reject) => {
			// To get the List of Opportunities to Display on Dashboard page
			this.requestUrl = 'api/Opportunity?page=1';

			fetch(this.requestUrl, {
				method: "GET",
				headers: { 'authorization': 'Bearer ' + window.authHelper.getWebApiToken() }
			})
				.then(response => {
					if (response.ok) {
						return response.json();
					} else {
						this.fetchResponseHandler(response, "getOpportunityIndex");
						reject(response.statusText);
					}
				})
				.then(data => {
					let itemslist = new Array();
					console.log("getOpportunityIndex number of items: " + data.ItemsList.length);
					if (data.ItemsList.length > 0) {
						for (var i = 0; i < data.ItemsList.length; i++) {

							var item = data.ItemsList[i];

							var newItem = new Object();

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

					let sortedList = this.state.reverseList ? itemslist.reverse() : itemslist;
					this.setState({
						loading: false,
						items: sortedList,
						itemsOriginal: sortedList
					});

					resolve(true);
				})
				.catch(err => {
					this.errorHandler(err, "getOpportunityIndex");
					this.setState({
						loading: false,
						items: new Array(),
						itemsOriginal: new Array()
					});
					reject(err);
				});

		})
	}

	getUserDetails() {
		return new Promise((resolve, reject) => {
			let data = this.authHelper.getUserProfile();
			if (data) {
				this.setState({
					AuthUserRole: data.userRole,
					AuthUserId: data.id,
					AuthUserDisplayName: data.displayName,
					AuthUserMail: data.mail,
					AuthUserPhone: data.phone,
					AuthUserPicture: data.picture,
					AuthUserUPN: data.userPrincipalName
				});

				resolve(data.userRole);
			} else {
				reject();
			}
		});
	}

    userProfileHandlerDEPRECATE() {
        console.log("TEST userProfileHandler");
        this.authHelper.getUserProfile()
            .then(data => {
                console.log("Dashboard_componentWillMount getUserProfile: " + JSON.stringify(data));
                this.setState({
                    AuthUserId: data.id,
                    AuthUserDisplayName: data.displayName,
                    AuthUserMail: data.mail,
                    AuthUserPhone: data.phone,
                    AuthUserPicture: data.picture,
                    AuthUserUPN: data.userPrincipalName,
                    AuthUserRole: data.userRole
                });
            })
            .catch((err) => {
                console.log("Dashboard_componentWillMount error: " + JSON.stringify(err));
                return (err);
            });
    }

    componentWillMount() {
        
	}

	componentDidMount() {
        
	}

	// Call NewOpportunity.js  - 1st Page
	fnCreateStep1() {
		this.state.newOpportunityState = null;
		this.state.newOpportunityDocument = null;
		this.state.newOpportunityOthers = null;
		this.setState({
			createStep: 1
		});
	}

	// Call NewOpportunityDocuments.js  - 2nd Page
	fnCreateStep2() {
		this.setState({
			createStep: 2
		});
	}

	// Call NewOpportunityOthers.js  - 3rd Page
	fnCreateStep3(event) {
		this.setState({
			createStep: 3
		});
	}

	// Back button
	fnBackStep() {
		this.setState({createStep: this.state.createStep - 1});
	}

	// Submit - Create New Opportunity
	submit() {
		this.submitPressed = true;
		this.setState({
			submitPressed: true,
			isLoadingOpportunity: true,
			createStep: 0,
			messageBarEnabled: true,
			messageBarText: "Creating Opportunity",
			messageBarType: MessageBarType.info

		});
	}

	// Submit complete
	submitComplete() {
		this.getOpportunityIndex()
			.then(res => {
				this.setState({
					createStep: 0,
					messageBarEnabled: false,
					messageBarText: "Creating Opportunity Complete",
					messageBarType: MessageBarType.success,
					submitPressed: false,
					isLoadingOpportunity: false

				});
			})
			.catch(err => {
				console.log("getOpportunityIndex error: " + err)
			});
	}

	setMessageBar(enabled, text, type) {
		this.setState({
			messageBarEnabled: enabled,
			messageBarText: text,
			messageBarType: type

		});
		// MessageBar types:
		// MessageBarType.error
		// MessageBarType.info
		// MessageBarType.severeWarning
		// MessageBarType.success
		// MessageBarType.warning
	}

	// Callback New Opportunity
	callbackStep1 = (dataFromChild) => {

		this.state.newOpportunityState = dataFromChild;
	  
		if (this.state.newOpportunityState.OpportunityName.length <= 0) {
			this.setState({
				MessagebarTextOpp: "Please enter a valid Opportunity Name",
				OppNameError: true
			});
			//this.setState({ isUpdate: false, MessagebarText: "" });
			this.setState({
				createStep: 1
			});
		}
		if (this.state.newOpportunityState.CustomerName.length <= 0) {
			//alert("Please enter a valid CustomerName!!!");
			this.setState({
				MessagebarTextCust: "Please enter a valid CustomerName",
				CustNameError: true
			});
			this.setState({
				createStep: 1
			});
		}
		if (this.state.newOpportunityState.AnnualRevenue.length > 0 && /^\d+$/.test(this.state.newOpportunityState.AnnualRevenue) === false) {
			this.setState({
				MessagebarTextAnnualRev: "AnnualRevenue should only contain Digits",
				AnnualRevenueError: true
			});
		   // alert("AnnualRevenue should only contain Digits!!!");
			this.setState({
				createStep: 1
			});
		}
		if (this.state.newOpportunityState.DealSize.length > 0 && /^\d+$/.test(this.state.newOpportunityState.DealSize) === false) {
		   // alert("DealSize should only contain Digits!!!");
			this.setState({
				MessagebarTextDealSize: "DealSize should only contain Digits",
				DealSizeError: true
			});
			this.setState({
				createStep: 1
			});
		}
	}

	// CallBack New Opportunity Document
    callbackStep2 = (dataFromChild) => {
		this.setState({
			newOpportunityDocument: dataFromChild
		});
	}

	// Callback New Opportunity Others
    callbackStep3 = (dataFromChild) => {
		this.state.newOpportunityOthers = dataFromChild;
		if (this.submitPressed) {

			if (this.state.newOpportunityOthers.DebtRatio.length > 0 && /^\d+$/.test(this.state.newOpportunityOthers.DebtRatio) === false) {
				this.setState({
					createStep: 3
				});
			}
			if (this.state.newOpportunityOthers.CollateralAmount.length > 0 && /^\d+$/.test(this.state.newOpportunityOthers.CollateralAmount) === false) {
				this.setState({
					createStep: 3
				});
			}
			if (this.state.newOpportunityOthers.RiskRating.length > 0 && /^\d+$/.test(this.state.newOpportunityOthers.RiskRating) === false) {
				this.setState({
					createStep: 3
				});
			}
			if (this.state.newOpportunityOthers.Margin.length > 0 && /^\d+$/.test(this.state.newOpportunityOthers.Margin) === false) {
				this.setState({
					createStep: 3
				});
			}
			else
				this.setOpportunityDetail();
		}
	}

    // Create New Opportunity
    createOpportunity(entityDto, grpDisplayName, opportunity) {
        return new Promise((resolve, reject) => {
            let requestUrl = 'api/opportunity/';

            var headers = new Headers();
            var bearer = "Bearer " + window.authHelper.getWebApiToken();
            var options = {
                method: "POST",
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'authorization': 'Bearer    ' + window.authHelper.getWebApiToken()
                },
                body: JSON.stringify(entityDto),
            };

            fetch(requestUrl, options)
                .then(response => this.fetchResponseHandler(response, "createOpportunity"))
                .then(data => {
                    this.setMessageBar(true, "Uploading files for the new opportunity", MessageBarType.info);
                    this.uploadFile(grpDisplayName, opportunity)
                        .then(respUpload => {
                            // TODO: Check response to inform user, remove spinners, eyc
                            console.log("Response upload: " + respUpload);
                            resolve(respUpload);
                        });
                })
                .catch(err => {
                    console.log("createOpportunity error: " + err);
                    reject(err);
                });
        })

    }

    uploadFile(grpDisplayName, opportunity) {
        return new Promise((resolve, reject) => {
            let files = this.state.newOpportunityDocument.selectorFiles;
            for (var i = 0; i < files.length; i++) {
                var fd = new FormData()
                fd.append('opportunity', "NewOpportunity")
                fd.append('file', files[i])
                fd.append('opportunityName', opportunity.OpportunityName);
                fd.append('fileName', files[i].name);
                let requestUrl = 'api/context/UploadFile/' + opportunity.OpportunityName + "/Attachment";
                var options = {
                    method: "PUT",
                    headers: {
                        'authorization': 'Bearer    ' + window.authHelper.getWebApiToken()
                    },
                    body: fd
                };

                fetch(requestUrl, options)
                    .then(response => this.fetchResponseHandler(response, "uploadFile"))
                    .then(data => {
                        this.setMessageBar(true, "Creating Opportunity Complete", MessageBarType.success);
                        resolve(data);
                    })
                    .catch(err => {
                        console.log("Error uploaing file: " + err);
                        reject(false);
                    });
            }
        })
    }

    // Creating New Opporunity Object and Call to Controller
    setOpportunityDetail() {
        var documentJSON = this.state.newOpportunityDocument;
        var others = this.state.newOpportunityOthers;
        var opportunity = this.state.newOpportunityState;
        var assignOpp = { "Name": "Not Started", "Value": 1 }
        var notes = opportunity.Notes + "" + documentJSON.notes
        var date = new Date()
        var teamMemberList = [];
        var entityDto = new Object();
        entityDto.id = "1";
        entityDto.displayName = opportunity.OpportunityName;
        entityDto.version = ""
        entityDto.reference = "";
        entityDto.opportunityState = 0;

        entityDto.dealSize = parseInt(opportunity.DealSize);
        entityDto.annualRevenue = parseInt(opportunity.AnnualRevenue);
        entityDto.openedDate = date;

        var industry = new Object();
        industry.displayName = opportunity.Industry;
        industry.id = "1";
        entityDto.industry = industry;

        var region = new Object();
        region.displayName = opportunity.Region;
        region.id = "1";

        entityDto.region = region;

        entityDto.margin = parseInt(others.Margin);
        entityDto.rate = parseInt(others.Rate);
        entityDto.purpose = others.Purpose;
        entityDto.disbursementSchedule = others.DisbursementSchedule;
        entityDto.collateralAmount = parseInt(others.CollateralAmount);
        entityDto.guarantees = others.Guarantees;
        entityDto.debtRatio = parseInt(others.DebtRatio);
        entityDto.riskRating = parseInt(others.RiskRating);
        var randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        var Customer = new Object();
        Customer.displayName = opportunity.CustomerName;
        Customer.referenceId = randomId.toString();
        Customer.id = "";
        entityDto.customer = Customer;

        var TeamMemberList = [];

        var loanOfficer =
            {
                "id": typeof others.currentSelectedItems[0] !== 'undefined' ? others.currentSelectedItems[0].key : "",
                "displayName": typeof others.currentSelectedItems[0] !== 'undefined' ? others.currentSelectedItems[0].primaryText : "",
                "mail": typeof others.currentSelectedItems[0] !== 'undefined' ? others.currentSelectedItems[0].userPrincipalName : "",
                "phoneNumber": "",
                "userPrincipalName": typeof others.currentSelectedItems[0] !== 'undefined' ? others.currentSelectedItems[0].userPrincipalName : "",
                "userPicture": "",
                "userRole": 1,
                "status": 0
            };


        var relationshipmanager =
            {
                "id": this.state.AuthUserId,
                "displayName": this.state.AuthUserDisplayName,
                "mail": this.state.AuthUserMail,
                "phoneNumber": this.state.AuthUserPhone,
                "userPrincipalName": this.state.AuthUserUPN,
                "userPicture": this.state.AuthUserPicture,
                "userRole": 2,
                "status": 0

            };

        var creditanalyst =
            {
                "id": "",
                "displayName": "",
                "mail": "",
                "phoneNumber": "",
                "userPrincipalName": "",
                "userPicture": "",
                "userRole": 3,
                "status": 0
            };


        var legalcounsel =
            {
                "id": "",
                "displayName": "",
                "mail": "",
                "phoneNumber": "",
                "userPrincipalName": "",
                "userPicture": "",
                "userRole": 4,
                "status": 0
            };


        var seniorriskofficer =
            {
                "id": "",
                "displayName": "",
                "mail": "",
                "phoneNumber": "",
                "userPrincipalName": "",
                "userPicture": "",
                "userRole": 5,
                "status": 0
            };

        var contentType = {
            "Name": "",
            "Value": 0
        }

        TeamMemberList.push(loanOfficer);
        TeamMemberList.push(relationshipmanager);
        TeamMemberList.push(creditanalyst);
        TeamMemberList.push(legalcounsel);
        TeamMemberList.push(seniorriskofficer);

        entityDto.teamMembers = TeamMemberList;
        var Notes = [];


        var userProfile =
            {
                "id": others.RelationshipManagerID,
                "displayName": others.RelationshipManagerName,
                "mail": others.RelationshipManagerUPN,
                "phoneNumber": "",
                "userPrincipalName": others.RelationshipManagerUPN,
                "userPicture": "",
                "userRole": 2,

            };
        var randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        var note1 =
            {
                "id": randomId.toString(),
                "createdBy": userProfile,
                "noteBody": notes,
                "createdDateTime": date
            };
        Notes.push(note1);

        entityDto.notes = Notes;

        var checkLists = [];

        var checkListRiskAssesment = new Object();
        var checkListCreditCheck = new Object();
        var checkListCompliance = new Object();


        var checklistTaskList = [];
        var checkListTask = {
            "checklistItem": "",
            "completed": false,
            "fileUri": ""
        };
        checklistTaskList.push(checkListTask);
        randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        checkListRiskAssesment.id = randomId.toString();
        checkListRiskAssesment.checklistChannel = "Risk Assessment";
        checkListRiskAssesment.checklistStatus = 0;
        checkListRiskAssesment.checklistTaskList = checklistTaskList;

        randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        checkListCreditCheck.id = randomId.toString();
        checkListCreditCheck.checklistChannel = "Credit Check";
        checkListCreditCheck.checklistStatus = 0;
        checkListCreditCheck.checklistTaskList = checklistTaskList;

        randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        checkListCompliance.id = randomId.toString();
        checkListCompliance.checklistChannel = "Compliance";
        checkListCompliance.checklistStatus = 0;
        checkListCompliance.checklistTaskList = checklistTaskList;

        checkLists.push(checkListRiskAssesment);
        checkLists.push(checkListCreditCheck);
        checkLists.push(checkListCompliance);

        entityDto.checklists = checkLists;

        randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        var CM =
            {
                "displayName": "",
                "id": randomId.toString()
            };

        randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        var DSM =
            {
                "id": randomId.toString(),
                "displayName": "",
                "owner": relationshipmanager,
                "sectionStatus": 0,
                "subSectionId": "",
                "lastModifiedDateTime": date
            };

        var DocSectionModelList = [];
        DocSectionModelList.push(DSM);


        //var ProposalSectionList1 = new List<ProposalDocumentContentModel>();

        var ProposalDocumentContentModel1 =
            {
                "proposalSectionList": DocSectionModelList
            };

        var NoteModel1 =
            {
                "id": "",
                "createdBy": userProfile,
                "reatedDateTime": date,
                "noteBody": ""
            };

        var Notes1 = [];
        Notes1.push(NoteModel1);
        randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        var ProposalDoc =
            {
                "id": randomId.toString(),
                "category": CM,
                "content": ProposalDocumentContentModel1,
                "contentType": contentType,
                "displayName": "",
                "documentUri": "",
                "reference": "",
                "tags": "",
                "version": "",
                "notes": Notes
            };
        entityDto.proposalDocument = ProposalDoc;

        randomId = Math.floor(Math.random() * (100000 - 6458) + 340)
        entityDto.artifactsBag = [];
        var custDecision = {
            "id": randomId.toString(),
            "approved": false,
            "approvedDate": date,
            "loanDisbursed": date
        };

        entityDto.customerDecision = custDecision;

        // Create team and channels
        let grpDisplayName = this.state.newOpportunityState.OpportunityName;
        let grpDescription = "This is the team group for " + grpDisplayName;
        this.setMessageBar(true, "Creating Team & Channels for the new opportunity", MessageBarType.info);
        this.sdkHelper.createTeamGroup(grpDisplayName, grpDescription)
            .then((resTeam) => {
                this.sdkHelper.createChannel("Risk Assessment", "Risk assessment channel", resTeam)
                    .then((res) => {
                        //console.log("2 ");
                        //console.log(resTeam);
                        this.sdkHelper.createChannel("Credit Check", "Credit check channel", resTeam)
                            .then((res) => {
                                //console.log("3 ");
                                //console.log(resTeam);
                                this.sdkHelper.createChannel("Compliance", "Compliance channel", resTeam)
                                    .then((res) => {
                                        //console.log("4 " + resTeam);
                                        this.sdkHelper.createChannel("Formal Proposal", "Formal proposal channel", resTeam)
                                            .then((res) => {
                                                //console.log("5 " + resTeam);
                                                this.sdkHelper.createChannel("Customer Decision", "Customer decision channel", resTeam)
                                                    .then((res) => {
                                                        this.setMessageBar(true, "Saving the new opportunity", MessageBarType.info);
                                                        grpDisplayName = grpDisplayName.replace(/\s/g, '');
                                                        this.createOpportunity(entityDto, grpDisplayName, opportunity)
                                                            .then(resp => {
                                                                console.log("finished last task in create opportunity");
                                                                this.submitComplete();
                                                            })
                                                            .catch(err => {
                                                                console.log("create opportunity chained tasks - last error: " + err)
                                                            });
                                                        // moving this call to inside createOpportunitythis.uploadFile(grpDisplayName, opportunity);
                                                    })
                                            })
                                            .catch(err => {
                                                //console.log(err.code + ' - ' + err.message);
                                            })
                                    })
                                    .catch(err => {
                                        //console.log(err.code + ' - ' + err.message);
                                    })
                            })
                            .catch(err => {
                                //console.log(err.code + ' - ' + err.message);
                            })
                    })
                    .catch(err => {
                        //console.log(err.code + ' - ' + err.message);
                    })
            })
            .catch(err => {
                //console.log(err.code + ' - ' + err.message);
            });

    }

	// List functions
	opportunitiesListHeading() {
		return (
			<div className='ms-List-th'>
				<div className='ms-List-th-itemName'>Opportunity</div>
				<div className='ms-List-th-itemClient'>Client</div>
				<div className='ms-List-th-itemDealsize'>Deal Size</div>
				<div className='ms-List-th-itemDate'>Opened Date</div>
				<div className='ms-List-th-itemState'>Status</div>
			</div>
		);
	}

	_onFilterByNameChanged(text) {
		const items = this.state.itemsOriginal;
		
		this.setState({
			filterClient: text,
			items: text ?
				items.filter(item => item.client.toString().toLowerCase().indexOf(text.toString().toLowerCase()) > -1) :
				items
		});
	}

	_onFilterByDealChanged(value) {
		const items = this.state.itemsOriginal;

		this.setState({
			filterDeal: value,
			items: value ?
				items.filter(item => item.dealsize >= value) :
				items
		});
	}

	_onRenderCell(item, index) {

		//<div className='ms-List-itemIndex'>{`Item ${index}`}</div>
		//<Icon
		//    className='ms-List-chevron'
		//    iconName={getRTL() ? 'ChevronLeft' : 'ChevronRight'}
		///>

		return (
			<div className='ms-List-itemCell' data-is-focusable={true}>
				<div className='ms-List-itemContent'>
					<div className='ms-List-itemName'>
                        <Link to={'/OpportunitySummary?opportunityId=' + item.id} >
                            {item.opportunity}
                        </Link>
					</div>
					<div className='ms-List-itemClient'>{item.client}</div>
					<div className='ms-List-itemDealsize'>{item.dealsize}</div>
					<div className='ms-List-itemDate'>{item.openedDate}</div>
					<div className='ms-List-itemState'>{oppStatusText[item.stausValue]}</div>
				</div>
			</div>
		);
	}

	opportunitiesList(itemsList, itemsListOriginal) {
		const lenght = typeof itemsList !== 'undefined' ? itemsList.length : 0;
		const lenghtOriginal = typeof itemsListOriginal !== 'undefined' ? itemsListOriginal.length : 0;
		const originalItems = itemsListOriginal;
		const items = itemsList;
		const resultCountText = lenght === lenghtOriginal ? '' : ` (${items.length} of ${originalItems.length} shown)`;

		console.log("list lenght = " + lenght + " resultcount: " + resultCountText);

		return (
			<FocusZone direction={FocusZoneDirection.vertical}>
				<List
					items={items}
					onRenderCell={this._onRenderCell}
					className='ms-List'
				/>
			</FocusZone>
		);
	}

	// Create button method to be deprecated
	createButtonDEPRECATE(authUserRole) {
		if (authUserRole === 2) {
			return (
				<div className='ms-Grid-col ms-sm6 ms-md4 ms-lg6 createButton '>
					<PrimaryButton className='pull-right' onClick={this.fnCreateStep1}> <i className="ms-Icon ms-Icon--Add pr10" aria-hidden="true"></i> Create New</PrimaryButton>
				</div>
			);
		} else if (authUserRole === 0) {
			return (
				<div className='ms-Grid-col ms-sm6 ms-md4 ms-lg6 createButton '>
					<Spinner size={SpinnerSize.small} label='Loading actions...' ariaLive='assertive' />
				</div>
			);
		}
		return (
			<div className='ms-Grid-col ms-sm6 ms-md4 ms-lg6 createButton '>
			</div>
		);
	}

    


    render() {
		const refreshToken = this.state.refreshToken;
		if (refreshToken) {
			this.login();
		}

		const isLoading = this.state.loading;

		const authUserRole = this.state.AuthUserRole;

		const itemsOriginal = this.state.itemsOriginal;
		const items = this.state.items;

		const lenghtOriginal = typeof itemsOriginal !== 'undefined' ? itemsOriginal.length : 0;
		const listHasItems = lenghtOriginal > 0 ? true : false;

		const opportunitiesListHeading = this.opportunitiesListHeading();
        const opportunitiesListComponent = this.opportunitiesList(items, itemsOriginal);

        let showDashboard = this.state.createStep === 0 ? true : false;
        let showStep1 = this.state.createStep === 1 ? true : false;
        let showStep2 = this.state.createStep === 2 ? true : false;
        let showStep3 = this.state.createStep === 3 ? true : false;
        let showChooseTeam = this.state.createStep === "chooseteam" ? true : false;

        const createStep = this.state.createStep;

        return (
            <div className='ms-Grid'>
                <DashboardView context={this.state} />
                <div>
                    {
                        createStep === 1 &&
                        <div>
                            <NewOpportunity parentState={this.state} callbackFromParent={this.callbackStep1} />
                            <div className='ms-Grid-row'>
                                <div className='docs-TextFieldExample ms-Grid-col ms-sm6 ms-md8 ms-lg6'><br />
                                    <PrimaryButton className='backbutton pull-left' onClick={this.setState({ createStep: this.state.createStep - 1 })}>Back</PrimaryButton>
                                </div>
                                <div className='docs-TextFieldExample ms-Grid-col ms-sm6 ms-md8 ms-lg6'><br />
                                    <PrimaryButton className='pull-right' onClick={this.fnCreateStep2}>Next</PrimaryButton>
                                </div>
                            </div>
                        </div>
                    }
                </div>
                <div>
                    {
                        createStep === 2 &&
                        <div>
                            <NewOpportunityDocuments parentState={this.state} callbackFromParent={this.callbackStep2} />
                            <div className='ms-Grid-row'>
                                <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg6'>
                                    <PrimaryButton className='backbutton pull-left' onClick={this.fnBackStep}>Back</PrimaryButton>
                                </div>
                                <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg6'>
                                    <PrimaryButton className='pull-right' onClick={this.fnCreateStep3}>Next</PrimaryButton>
                                </div>
                            </div>
                        </div>
                    }
                </div>
                <div>
                    {
                        createStep === 3 &&
                        <div>
                            <NewOpportunityOthers parentState={this.state} callbackFromParent={this.callbackStep3} />
                            <div className='ms-grid-row'>
                                <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg6'><br />
                                    <PrimaryButton className='backbutton pull-left' onClick={this.fnBackStep}>Back</PrimaryButton>
                                </div>
                                <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg6'><br />
                                    <LinkContainer to='/'>
                                        <PrimaryButton className='pull-right' onClick={this.submit}>Submit</PrimaryButton>
                                    </LinkContainer>
                                </div>
                            </div>
                        </div>
                    }
                </div>
                <div>
                    {
                        createStep === 'chooseteam' &&
                        <div className='ms-grid-row'>
                            <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg8'>
                                Load choose team component
                                </div>
                            <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg4'>
                            </div>
                        </div>
                    }
                </div>
            </div >
            );



		if (this.state.createStep === 1) {
			return (
				<div className='ms-Grid'>
                    
				</div>

			);
		}

		if (this.state.createStep === 2) {
			return (
				<div className='ms-Grid'>
                    
				</div>
			);
		}

		if (this.state.createStep === 3) {
			return (
				<div className='ms-Grid'>
                    
				</div>
			);
		}

		if (this.state.createStep === 'chooseteam') {
			return (
				<div className='ms-Grid'>
					
				</div>
			);
		}
	}

}
