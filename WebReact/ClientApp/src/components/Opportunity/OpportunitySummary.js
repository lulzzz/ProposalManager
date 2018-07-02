/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { PrimaryButton, Button } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Link as LinkRoute } from 'react-router-dom';
import { TeamMembers } from '../../components/Opportunity/TeamMembers';
import {
    Persona,
    PersonaSize
} from 'office-ui-fabric-react/lib/Persona';
import {
    Spinner,
    SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar } from 'office-ui-fabric-react/lib/MessageBar';
import { userRoles, oppStatusText } from '../../common';
import { PeoplePickerTeamMembers } from '../PeoplePickerTeamMembers';



export class OpportunitySummary extends Component {
    displayName = OpportunitySummary.name
    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
		this.authHelper = window.authHelper;

		const userProfile = this.props.userProfile;
		
		const oppId = this.props.opportunityId;

		const teamUrl = "https://teams.microsoft.com";

		let isAdmin = false;

		if (this.props.userProfile.roles.filter(x => x.displayName === "Administrator").length > 0) {
			isAdmin = true;
		}

        this.state = {
            loading: true,
            loadView: 'summary',
            menuLevel: 'Level2',
            LoanOfficer: [],
            OppDetails: [],
            teamMembers: [],
            oppId: oppId,
            showPicker: false,
            TeamMembersAll: [],
            peopleList: [],
            mostRecentlyUsed: [],
            currentSelectedItems: [],
            oppData: [],
			btnSaveDisable: false,
			//userRoles: userProfile.roles,
			userId: userProfile.id,
            usersPickerLoading: true,
            loanOfficerPic: '',
            loanOfficerName: '',
            loanOfficerRole:'',
			teamUrl: teamUrl ,
			isAdmin: isAdmin,
			userAssignedRole : ""
        };
    }

    componentWillMount() {
        if (this.state.peopleList.length === 0) {
            this.getUserProfiles();
        }

        this.getOppDetails();
    }

    getOppDetails() {
        //return new Promise((resolve, reject) => {
            //fetch starts
            let requestUrl = 'api/Opportunity/?id=' + this.state.oppId;

            fetch(requestUrl, {
                method: "GET",
                headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
            })
                .then(response => response.json())
                .then(data => {
                    try {
                        // filter loan officers

                        let loanOfficerObj = data.teamMembers.filter(function (k) {
                            return k.assignedRole.displayName === "LoanOfficer";
                        });
                        let officer = {};
                        if (loanOfficerObj.length > 0) {
                            officer.loanOfficerPic = "";
                            officer.loanOfficerName = loanOfficerObj[0].text;
                            officer.loanOfficerRole = "";
                        }
                        let teamMembers = [];

                        teamMembers = data.teamMembers;

						let currentUserId = this.state.userId;
						

						let teamMemberDetails = data.teamMembers.filter(function (k) {
							return k.id === currentUserId;
						});

						let userAssignedRole = teamMemberDetails[0].assignedRole.displayName;

                        this.setState({
                            loading: false,
                            teamMembers: teamMembers,
                            oppData: data,
                            LoanOfficer: loanOfficerObj.length === 0 ? loanOfficerObj : [],
							showPicker: loanOfficerObj.length === 0 ? true : false,
							userAssignedRole: userAssignedRole

                        });
                    }
                    catch (err) {
                        //console.log("Error")
                    }
                });
       // });
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
                    this.setState({ usersPickerLoading: false });
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

                // filter to just loan officers
                //let filteredList = itemslist.filter(itm => itm.userRole === 1);
                
                this.setState({
                    usersPickerLoading: true,
                    peopleList: itemslist
                });

                this.filterUserProfiles();
            })
            .catch(err => {
                console.log("Opportunities_getUserProfiles error: " + JSON.stringify(err));
            });
    }

    filterUserProfiles() {
        let data = this.state.peopleList;
        let teamlist = [];

        for (let i = 0; i < data.length; i++) {
            let item = data[i];

            if ((item.userRoles.filter(x => x.displayName === "LoanOfficer")).length > 0) {
                teamlist.push(item);
            }
        }

        this.setState({
            usersPickerLoading: false,
            peopleList: teamlist
        });
    }

    fetchResponseHandler(response, referenceCall) {
        if (response.status === 401) {
            // TODO_ Next version handling of refresh token
        }
    }

    errorHandler(err, referenceCall) {
        console.log("Opportunities Ref: " + referenceCall + " error: " + JSON.stringify(err));
    }

    toggleHiddenPicker() {
        this.setState({
            showPicker: !this.state.showPicker
        });
    }

    static funcThis(fields) {
        if (typeof fields !== "undefined")
            return fields.displayName;
    }
    renderSummaryDetails(oppDeatils) {
        let loanOfficerArr = [];
        loanOfficerArr = oppDeatils.teamMembers.filter(function (k) {
			return k.assignedRole.displayName === "LoanOfficer";


        });

		let enableConnectWithTeam;
		if (this.state.oppData.opportunityState !== 1 ) {
			enableConnectWithTeam = true;
		}
		else {
			enableConnectWithTeam = false;
		}
		return (
			
            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 p10A'>
                <div className='ms-Grid-row bg-white pt15'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Opportunity Name </Label>
                        <span>{oppDeatils.displayName}</span>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Client Name </Label>
                        <span>{oppDeatils.customer.displayName}</span>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Opened Date  </Label>
                        <span>{new Date(oppDeatils.openedDate).toLocaleDateString()} </span>
                    </div>
                </div>
                <div className='ms-Grid-row bg-white none'>
                    <div className='ms-Grid ms-sm12 ms-md12 ms-lg12 pb10'>
                        &nbsp;
                    </div>
                </div>
                <div className='ms-Grid-row bg-white'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Deal Size </Label>
                        <span>{oppDeatils.dealSize.toLocaleString()} </span>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Annual Revenue </Label>
                        <span>{oppDeatils.annualRevenue.toLocaleString()}</span>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Industry </Label>
                        <span>{oppDeatils.industry.name} </span>
                    </div>
                </div>
                <div className='ms-Grid-row bg-white none'>
                    <div className='ms-Grid ms-sm12 ms-md12 ms-lg12 pb10'>
                        &nbsp;
                    </div>
                </div>
                <div className='ms-Grid-row bg-white'>
                    <div className=' ms-Grid-col ms-sm12 ms-md6 ms-lg4 pb10'>
                        <Label>Region </Label>
                        <span>{oppDeatils.region.name} </span>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Status </Label>
                        <span> {oppStatusText[oppDeatils.opportunityState]} </span>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Margin ($M) </Label>
                        <span>{oppDeatils.margin}</span>
                    </div>
                </div>
                <div className='ms-Grid-row bg-white none'>
                    <div className='ms-Grid ms-sm12 ms-md12 ms-lg12  '>
                        &nbsp;
                    </div>
                </div>
                <div className='ms-Grid-row bg-white'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Rate </Label>
                        <span>{oppDeatils.rate} </span>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>
                        <Label>Debt Ratio </Label>
                        <span>{oppDeatils.debtRatio}</span>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg4 pb10'>

                        <Label>Loan Officer </Label>
                        
                        {
                            //loanOfficerName.length > 0 ?
                            loanOfficerArr.length > 0 ?
                                <div>
                                    {this.state.showPicker ? "" :
                                        <div>
                                            <Persona
                                                { ...{ imageUrl: loanOfficerArr[0].UserPicture } }
                                                size={PersonaSize.size40}
                                                primaryText={loanOfficerArr[0].displayName}
                                                secondaryText="Loan Officer"
                                            />
                                            {
                                                this.state.oppData.opportunityState === 10 ?
                                                <Link className="pull-right" disabled>Change</Link>
                                                :
                                                <Link onClick={this.toggleHiddenPicker.bind(this)} className="pull-right">Change</Link>
                                            }
                                        </div>

                                    }
                                </div>
                                :
                                ""

                        }
                        {this.state.showPicker ?
                            <div>
                                {this.state.usersPickerLoading
                                    ? <div className='ms-BasicSpinnersExample'>
                                        <Spinner size={SpinnerSize.large} label='loading...' ariaLive='assertive' />
                                    </div>
                                    :
                                    <div>
                                        <PeoplePickerTeamMembers teamMembers={this.state.peopleList} onChange={(e) => this.fnChangeLoanOfficer(e)} />
                                        <br />
                                        <Button
                                            buttonType={0}
                                            onClick={this._fnUpdateLoanOfficer.bind(this)}
                                            disabled={(!(this.state.currentSelectedItems.length === 1))}
                                        >
                                            Save
                                        </Button>
                                    </div>
                                }
                                {
                                   this.state.isUpdate ?
                                          <Spinner size={SpinnerSize.large} label='updating' ariaLive='assertive' />
                                        : ""
                                }

                                </div>
                            : ""
                        }
                        <br />

                        {
                            this.state.result &&
                            <MessageBar
                                messageBarType={this.state.result.type}
                            >
                                {this.state.result.text}
                            </MessageBar>
                        }

                    </div>
                </div>
                <div className='ms-Grid-row bg-white'>
					<div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3 '>
                        {
							enableConnectWithTeam
								?
								<a href={this.state.teamUrl} target="_blank" rel="noopener noreferrer"><PrimaryButton className='' >Connect With Team</PrimaryButton></a>
								:
								<PrimaryButton className='' disabled >Connect With Team</PrimaryButton>
                            
						}
                    </div>
                </div>
                <div className='ms-Grid-row bg-white'>
                    <div className='ms-Grid ms-sm12 ms-md12 ms-lg12'>
                        &nbsp;
                    </div>
                </div>
            </div>
        );
    }

    _renderSubComp() {
        let oppDetails = this.state.loading ? <div className='bg-white'><p><em>Loading...</em></p></div> : this.renderSummaryDetails(this.state.oppData);
        switch (this.state.loadView) {
            case 'summary': return (
                <div>
                    <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg9 p-5 bg-grey'>
                        <div className='ms-Grid-row'>
                            <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg6 pageheading'>
                                <h3>Opportunity Details</h3>
                            </div>
                            <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg6'><br />
                                <LinkRoute to={'/'} className='pull-right'>Back to Dashboard </LinkRoute>
                            </div>
                        </div>
                        <div className='ms-Grid-row  p-r-10'>
                            {oppDetails}
                        </div>
                    </div>
                </div>

            );
            case 'chooseteam': return (
                <div>
                    <h2>Choose Team</h2>
                </div>
            );
            default:
                break;

        }
    }

	fnChangeLoanOfficer(item) {
        this.setState({ currentSelectedItems: item });
        if (this.state.currentSelectedItems.length > 1) {
            this.setState({
                btnSaveDisable: true
            });
        } else {
            this.setState({
                btnSaveDisable: false
            });
        }        
    }

    _fnUpdateLoanOfficer() {
        let oppDetails = this.state.oppData; //oppData;
        let selLoanOfficer = this.state.currentSelectedItems;

        this.setState({
            loanOfficerName : selLoanOfficer[0].text,
            loanOfficerPic : '', //selLoanOfficer[0].imageUrl,
            loanOfficerRole : userRoles[0]
        });
        let updloanOfficer =
            {
                "id": selLoanOfficer[0].id,
                "displayName": selLoanOfficer[0].text,
                "mail": selLoanOfficer[0].mail,
                "phoneNumber": "",
                "userPrincipalName": selLoanOfficer[0].userPrincipalName,
                //"userPicture": selLoanOfficer[0].imageUrl,
                "userRole": selLoanOfficer[0].userRoles,
                "status": 0,
                "assignedRole": selLoanOfficer[0].userRoles.filter(x => x.displayName === "LoanOfficer")[0]
            };

        let isLoanOfficerExists = false;
        for (let t = 0; t < oppDetails.teamMembers.length; t++) {
            if (oppDetails.teamMembers[t].assignedRole.displayName === "LoanOfficer") {
                oppDetails.teamMembers[t] = updloanOfficer;
                isLoanOfficerExists = true;
            }
        }

        if (!isLoanOfficerExists) {
            oppDetails.teamMembers.push(updloanOfficer);
        }

        this.setState({ memberslist: oppDetails.teamMembers });
        this.fnUpdateCustDecision(oppDetails.teamMembers);
    }

    fnUpdateCustDecision(updTeamMembersObj) {
        this.setState({ isUpdate: true });

        let oppViewData = this.state.oppData;
        oppViewData.teamMembers = updTeamMembersObj;

        // API Update call        
        this.requestUpdUrl = 'api/opportunity?id=' + oppViewData.id; 
        let options = {
            method: "PATCH",
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'authorization': 'Bearer    ' + window.authHelper.getWebApiToken()
            },
            body: JSON.stringify(oppViewData)
        };

        fetch(this.requestUpdUrl, options)
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    return response.json;
                } else {
                    //console.log('Error...: ');
                }
            }).then(json => {
                //console.log(json);
                this.setState({ isUpdate: false, showPicker: false });
            });


    }

	render() {
		
        if (this.state.loading) {
            return (
                <div className='ms-BasicSpinnersExample ibox-content pt15 '>
                    <Spinner size={SpinnerSize.large} label='loading...' ariaLive='assertive' />
                </div>
            );
        } else {
            return (
                <div className='ms-Grid'>
                    <div className='ms-Grid-row'>
                        {this._renderSubComp()}
                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3 p-l-10 TeamMembersBG'>
                            <h3>Team Members</h3>
							<TeamMembers memberslist={this.state.teamMembers} createTeamId={this.state.oppData.id} opportunityName={this.state.oppData.displayName} opportunityState={this.state.oppData.opportunityState} userRole={this.state.userAssignedRole} />
                        </div>
                    </div>
                </div>
            );
        }
    }
    
}