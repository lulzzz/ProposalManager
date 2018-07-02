/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/
import React, { Component } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';

import { Pivot, PivotItem, PivotLinkFormat, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Workflow } from './Proposal/Workflow';
import { TeamUpdate } from './Proposal/TeamUpdate';
import { TeamsComponentContext } from 'msteams-ui-components-react';
import './teams.css';
import { GroupEmployeeStatusCard } from '../components/Opportunity/GroupEmployeeStatusCard';


export class RootTab extends Component {
	displayName = RootTab.name

	constructor(props) {
		super(props);

		this.authHelper = window.authHelper;
		this.sdkHelper = window.sdkHelper;
		

		try {
			microsoftTeams.initialize();
		}
		catch (err) {
			console.log(err);
		}
		finally {
			this.state = {
				teamName: "",
				channelId: "",
				groupId: "",
				teamMembers: [],
				errorLoading: false,
                OppName: "",
                oppDetails: "",
                UserRoleList: [],
                OtherRoleTeamMembers: []
			};
		}
	}

	componentWillMount() {
		// Get the teams context
        this.getTeamsContext();
        this.getUserRoles();
	}

	componentDidUpdate() {
		const teamName = this.state.teamName;
		const teamMembers = this.state.teamMembers;

		if (teamName.length > 0 && teamMembers.length === 0 && this.state.errorLoading === false) {
			this.getOpportunityData(teamName);
		}
	}

	getTeamsContext() {
		microsoftTeams.getContext(context => {
			if (context) {
				this.setState({
					teamName: context.teamName,
					channelId: context.channelId,
					groupId: context.groupId
				});
			}
		});
	}

	getOpportunityData(teamName) {
		// API - Fetch call
		let requestUrl = "api/Opportunity?name='" + teamName + "'";
		fetch(requestUrl, {
			method: "GET",
			headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }

		})
			.then(response => {
				if (response.status === 200) {
					return response.json();
				} else {
					this.setState({
						errorLoading: true
					});
                    return null;
				}
			})
			.then(data => {
				if (data !== null) {
                    let teamMembers = data.teamMembers;
                    // Get Other role officers list
                    let otherRolesMapping = this.state.UserRoleList.filter(function (k) {
                        return k.processType.toLowerCase() !== "base" && k.processType.toLowerCase() !== "administration" && k.processType.toLowerCase() !== "customerdecisiontab" && k.processType.toLowerCase() !== "proposalstatustab";
                    });

                    let otherRolesArr1 = [];
                    for (let j = 0; j < otherRolesMapping.length; j++) {
                        let teamMember = data.teamMembers.filter(function (k) {
                            if (k.assignedRole.displayName.toLowerCase() === otherRolesMapping[j].roleName.toLowerCase()) {
                                //ProcessStep
                                k.processStep = otherRolesMapping[j].processStep;
                                //ProcessStatus
                                let processStatus = [];
                                processStatus = data.checklists.filter(function (p) {
                                    if (p.id.toLowerCase() === otherRolesMapping[j].processStep.toLowerCase()) {
                                        return p;
                                    }
                                });
                                if (processStatus.length > 0)
                                    k.processStatus = processStatus[0].checklistStatus ? processStatus[0].checklistStatus : 0;
                                else k.processStatus = 0;
                                return k.assignedRole.displayName.toLowerCase() === otherRolesMapping[j].roleName.toLowerCase();
                            }
                        });
                        if (teamMember.length === 0) {
                            teamMember = [{
                                "displayName": "",
                                "assignedRole": {
                                    "displayName": otherRolesMapping[j].roleName,
                                    "adGroupName": otherRolesMapping[j].adGroupName
                                },
                                "processStep": otherRolesMapping[j].processStep,
                                "processStatus": 0,
                                "status": 0
                            }];
                        }
                        otherRolesArr1 = otherRolesArr1.concat(teamMember);
                    }


                    let UserRolesList = this.state.UserRoleList;
                    let otherRolesArr = otherRolesArr1.reduce(function (res, currentValue) {
                        if (res.indexOf(currentValue.assignedRole.displayName) === -1) {
                            res.push(currentValue.assignedRole.displayName);
                        }
                        return res;
                    }, []).map(function (group) {
                        return {
                            group: group,
                            users: otherRolesArr1.filter(function (_el) {
                                return _el.assignedRole.displayName === group;
                            }).map(function (_el) { return _el; })
                        };
                    });
                    let otherRolesObj = [];
                    if (otherRolesArr.length > 1) {
                        for (let r = 0; r < otherRolesArr.length; r++) {
                            otherRolesObj.push(otherRolesArr[r].users);
                        }
                    }
					this.setState({
						loading: false,
						teamMembers: teamMembers,
						oppDetails: data,
						oppStatus: data.opportunityState,
                        OppName: data.displayName,
                        OtherRoleTeamMembers: otherRolesObj
					});
				}
			})
			.catch(err => {
				console.log("RootTab_getOpportunityData error");
				console.log(err);
			});
	}

	resetToken() {
		this.authHelper.logout().then(() => {
            window.location.reload();
		});
	}

    getUserRoles() {
        // call to API fetch data
        let requestUrl = 'api/RoleMapping';
        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + window.authHelper.getWebApiToken() }
        })
            .then(response => response.json())
            .then(data => {
                try {
                    let userRoleList = [];
                    for (let i = 0; i < data.length; i++) {
                        let userRole = {};
                        userRole.id = data[i].id;
                        userRole.roleName = data[i].roleName;
                        userRole.adGroupName = data[i].adGroupName;
                        userRole.processStep = data[i].processStep;
                        userRole.processType = data[i].processType;
                        userRoleList.push(userRole);
                    }
                    this.setState({ UserRoleList: userRoleList });
                }
                catch (err) {
                    return false;
                }

            });
    }


	render() {
		const team = this.state.teamMembers;
		const channelId = this.state.channelId;

		let loanOfficerRealManagerArr = [];
		let otherTeamMembersArr = [];

        let loanOfficerRealManagerArr1 = team.filter(x => x.assignedRole.displayName === "LoanOfficer");
        if (loanOfficerRealManagerArr1.length === 0) {
            loanOfficerRealManagerArr1 = [{
                "displayName": "",
                "assignedRole": {
                    "displayName": "CreditAnalyst"
                }
            }];
        }
		let loanOfficerRealManagerArr2 = team.filter(x => x.assignedRole.displayName === "RelationshipManager");

		loanOfficerRealManagerArr = loanOfficerRealManagerArr1.concat(loanOfficerRealManagerArr2);

        let otherTeamMembersArr1 = team.filter(x => x.assignedRole.displayName === "CreditAnalyst");
        if (otherTeamMembersArr1.length === 0) {
            otherTeamMembersArr1 = [{
                "displayName": "",
                "assignedRole": {
                    "displayName": "CreditAnalyst"
                }
            }];
        }

        let otherTeamMembersArr2 = team.filter(x => x.assignedRole.displayName === "LegalCounsel");
        if (otherTeamMembersArr2.length === 0) {
            otherTeamMembersArr2 = [{
                "displayName": "",
                "assignedRole": {
                    "displayName": "LegalCounsel"
                }
            }];
        }

		let otherTeamMembersArr3 = team.filter(x => x.assignedRole.displayName === "SeniorRiskOfficer");
        if (otherTeamMembersArr3.length === 0) {
            otherTeamMembersArr3 = [{
                "displayName": "",
                "assignedRole": {
                    displayName: "SeniorRiskOfficer"
                }
            }];
        }

		otherTeamMembersArr = otherTeamMembersArr1.concat(otherTeamMembersArr2, otherTeamMembersArr3);

		return (
			
			<TeamsComponentContext>
				<div className='ms-Grid'>
					<div className='ms-Grid-row'>
						{this.state.teamName} teamName here
						<div className='ms-Grid-col ms-sm6 ms-md8 ms-lg12 bgwhite tabviewUpdates noscroll pL0' >
							{
								this.state.errorLoading ?
									<div>
										We found an error while reading the opportunity data, please refresh the tab by clicking on the refresh button on the top right section.
										<br /><br />
                                        <PrimaryButton className='pull-right refreshbutton' onClick={() => this.resetToken()}>
											Reset Tab
										</PrimaryButton>
									</div>
									:
									<Pivot className='tabcontrols' linkFormat={PivotLinkFormat.tabs} linkSize={PivotLinkSize.large}>
										<br />
										<br />
                                        <PivotItem linkText='Workflow' width='100%' >
                                            <Label><Workflow memberslist={this.state.teamMembers} oppStaus={this.state.oppStatus} oppDetails={this.state.oppDetails} /></Label>
										</PivotItem>
										<PivotItem linkText='Team Update'>
                                            <div className='ms-Grid-row mt20 mr20 pl15'>
                                                {
                                                    this.state.OtherRoleTeamMembers.map((obj, ind) =>
                                                        obj.length > 1 ?
                                                            <div className=' ms-Grid-col ms-sm12 ms-md8 ms-lg4 p-5' key={ind}>
                                                                <GroupEmployeeStatusCard members={obj} status={obj[0].status} isDispOppStatus='false' role={obj[0].assignedRole.adGroupName} isTeam='true' />
                                                            </div>
                                                            :
                                                            obj.map((member, j) =>
                                                                <div className=' ms-Grid-col ms-sm12 ms-md8 ms-lg4 p-5' key={j}>
                                                                    <TeamUpdate memberslist={member} channelId={channelId} groupId={this.state.groupId} OppName={this.state.OppName} />
                                                                </div>
                                                            )

                                                    )

                                                }
											</div>
											<div className='ms-Grid-row'>
												<div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12' > </div>
											</div>
                                            <div className='ms-Grid-row pl15'>
												{
                                                    loanOfficerRealManagerArr.map((member, ind) =>
                                                        <div className=' ms-Grid-col ms-sm12 ms-md8 ms-lg4 p-5' key={ind} >
															<TeamUpdate memberslist={member} channelId={channelId} groupId={this.state.groupId} OppName={this.state.OppName} />
														</div>
													)
												}
											</div>
										</PivotItem>
									</Pivot>
							}
						</div>
					</div>
					<div className='ms-Grid-row'>
						<div className='ms-Grid-col ms-sm6 ms-md8 ms-lg10'>
						</div>
					</div>
				</div>
			</TeamsComponentContext>
		);
	}
}