/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Link as LinkRoute } from 'react-router-dom';
import { TeamMembers } from '../../components/Opportunity/TeamMembers';
import { EmployeeStatusCard } from '../../components/Opportunity/EmployeeStatusCard';
import { GroupEmployeeStatusCard } from '../../components/Opportunity/GroupEmployeeStatusCard';
import {
    Spinner,
    SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';

import {
    Persona,
    PersonaSize
} from 'office-ui-fabric-react/lib/Persona';


export class OpportunityStatus extends Component {
    displayName = OpportunityStatus.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;

        const userProfile = this.props.userProfile;

        const oppId = this.props.opportunityId;

        this.state = {
            oppId: oppId,
            loading: true,
            teamMembers: [],
            LoanOfficer: [],
            userId: userProfile.id,
            UserRoleList: [],
            OtherRolesMapping: []
        };

    }

    componentWillMount() {
        this.getUserRoles();
        this.getOppDetails();
    }

    getOppDetails() {
        let requestUrl = 'api/Opportunity/?id=' + this.state.oppId;

        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + window.authHelper.getWebApiToken() }
        })
            .then(response => response.json())
            .then(data => {
                
                let loanOfficerObj = data.teamMembers.filter(function (k) {
                    return k.assignedRole.displayName === "LoanOfficer"; // "loan officer";
                });

                let relManagerObj = data.teamMembers.filter(function (k) {
                    return k.assignedRole.displayName === "RelationshipManager"; // "relationshipmanager";
                });

                // Get Other role officers list

                let otherRolesMapping = this.state.UserRoleList.filter(function (k) {
                    //return (k.roleName.toLowerCase() !== "relationshipmanager" && k.roleName.toLowerCase() !== "loanofficer" && k.roleName.toLowerCase() !== "administrator");
                    return k.processType.toLowerCase() !== "base" && k.processType.toLowerCase() !== "administration" && k.processType.toLowerCase() !== "customerdecisiontab" && k.processType.toLowerCase() !== "proposalstatustab";
                });

                this.setState({ OtherRolesMapping: otherRolesMapping });
                let otherRolesArr1 = [];
                for (let j = 0; j < otherRolesMapping.length; j++) {
                    let teamMember = data.teamMembers.filter(function (k) {
                        if (k.assignedRole.displayName.toLowerCase() === otherRolesMapping[j].roleName.toLowerCase()) {
                            //ProcessStep
                            k.processStep = otherRolesMapping[j].processStep;
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


                let userId = this.state.userId;
                let currentUser = data.teamMembers.filter(function (k) {
                    return k.id === userId;
                });

                let assignedUserRole;
                if (currentUser.length > 0) {
                    assignedUserRole = currentUser[0].assignedRole.displayName;
                }
                else {
                    assignedUserRole = "";
                }

                this.setState({
                    LoanOfficer: loanOfficerObj,
                    RelationShipOfficer: relManagerObj,
                    OtherRoleOfficers: otherRolesObj,
                    teamMembers: data.teamMembers,
                    oppData: data,
                    userRole: assignedUserRole,
                    loading: false
                });

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
                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg9 p-r-10 bg-white'>
                            <div className='ms-Grid-row'>
                                <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg6 pageheading'>
                                    <h3>Status</h3>
                                </div>
                                <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg6'><br />
                                    <LinkRoute to={'/'} className='pull-right'>Back to Dashboard </LinkRoute>
                                </div>
                            </div>
                            <div className='ms-Grid-row p-5'>
                                <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12 bg-white pb20'>
                                    <div className='ms-Grid-row pt35'>
                                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3  bg-gray newOpBg p20A'>
                                                <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                                &nbsp;&nbsp;<span>New Opportunity</span>
                                                {this.state.RelationShipOfficer.length > 0 ?
                                                    this.state.RelationShipOfficer.map(officer =>
                                                        <EmployeeStatusCard key={officer.id}
                                                            { ...{
                                                                id: officer.id,
                                                                name: officer.displayName,
                                                                image: "",
                                                                role: officer.assignedRole.adGroupName,
                                                                status: officer.status,
                                                                isDispOppStatus: false
                                                            }
                                                            }
                                                        />
                                                    )
                                                    : ""
                                                }
                                        </div>
                                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3 newOpBg p20A'>
                                            <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                            &nbsp;&nbsp;<span>Start Process</span>
                                            {
                                                this.state.LoanOfficer.length === 0 ?
                                                    <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12 bg-grey p-5'>
                                                        <div className='ms-PersonaExample'>
                                                            <div className='ms-Grid-row'>
                                                                <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg4'>
                                                                    <Label>Status</Label>

                                                                </div>
                                                                <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg8'>
                                                                    <Label><span className='notstarted'> Not Started </span></Label>

                                                                </div>
                                                            </div>
                                                            <div className='ms-Grid-row'>
                                                                <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12'>
                                                                    <Persona
                                                                        { ...{
                                                                            imageUrl: "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==",
                                                                            imageInitials: ""
                                                                        } }
                                                                        size={PersonaSize.size40}
                                                                        text="User Not Selected"
                                                                        secondaryText="Loan Officer"
                                                                    />

                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                    :
                                                    this.state.LoanOfficer.length === 1 ?
                                                        this.state.LoanOfficer.map(officer =>
                                                            <EmployeeStatusCard key={officer.id}
                                                                { ...{
                                                                    id: officer.id,
                                                                    name: officer.displayName,
                                                                    image: "",
                                                                    role: officer.assignedRole.adGroupName,
                                                                    status: officer.status,
                                                                    isDispOppStatus: false
                                                                }
                                                                }
                                                            />
                                                        )
                                                        :
                                                        <GroupEmployeeStatusCard members={this.state.LoanOfficer} status={this.state.LoanOfficer[0].status} isDispOppStatus='false' role='Loan Officer' />
                                            }
                                        </div>
                                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3 divUserRolegroup-arrow'>
                                            <div className='  divUserRolegroup'>
                                                <div>
                                                    {
                                                        this.state.OtherRoleOfficers.length > 1 ?
                                                            <div>
                                                                {
                                                                    this.state.OtherRoleOfficers.map((obj, ind) =>
                                                                        obj.length > 1 ?
                                                                            <div key={ind}>
                                                                                <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                                                                &nbsp;&nbsp;<span>{obj[0].processStep}</span>
                                                                                <GroupEmployeeStatusCard members={obj} status={obj[0].status} role={obj[0].assignedRole.adGroupName} />
                                                                            </div>
                                                                            :
                                                                            obj.length === 1 ?
                                                                                obj.map(officer =>
                                                                                    <div key={ind}>
                                                                                        <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                                                                        &nbsp;&nbsp;<span>{officer.processStep}</span>
                                                                                        <EmployeeStatusCard key={officer.id}
                                                                                            { ...{
                                                                                                id: officer.id,
                                                                                                name: officer.displayName,
                                                                                                image: "",
                                                                                                role: officer.assignedRole.adGroupName,
                                                                                                status: officer.status,
                                                                                                isDispOppStatus: false
                                                                                            }
                                                                                            }
                                                                                        />
                                                                                    </div>

                                                                                )
                                                                                :
                                                                                <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12' key={ind}>
                                                                                    <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                                                                    &nbsp;&nbsp;<span>{this.state.OtherRolesMapping[ind].processStep}</span>
                                                                                    <div className='ms-PersonaExample bg-grey p-5'>
                                                                                        <div className='ms-Grid-row'>
                                                                                            <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg4'>
                                                                                                <Label>Status</Label>

                                                                                            </div>
                                                                                            <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg8'>
                                                                                                <Label><span className='notstarted'> Not Started </span></Label>

                                                                                            </div>
                                                                                        </div>
                                                                                        <div className='ms-Grid-row'>
                                                                                            <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12'>
                                                                                                <Persona
                                                                                                    { ...{
                                                                                                        imageUrl: "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==",
                                                                                                        imageInitials: ""
                                                                                                    } }
                                                                                                    size={PersonaSize.size40}
                                                                                                    text="User Not Selected"
                                                                                                    secondaryText=""
                                                                                                />

                                                                                            </div>
                                                                                        </div>
                                                                                    </div>
                                                                                </div>
                                                                    )
                                                                }
                                                            </div>
                                                            : 
                                                            ""
                                                    }
                                                </div>
                                           
                                        </div>
                                        </div>
                                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3'>
                                            <div>
                                                <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                                &nbsp;&nbsp;<span>Draft Proposal</span>
                                                {
                                                    this.state.LoanOfficer.length === 0 ?
                                                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12 bg-grey p-5'>
                                                            <div className='ms-PersonaExample'>
                                                                <div className='ms-Grid-row'>
                                                                    <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg4'>
                                                                        <Label>Status</Label>

                                                                    </div>
                                                                    <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg8'>
                                                                        <Label><span className='notstarted'> Not Started </span></Label>

                                                                    </div>
                                                                </div>
                                                                <div className='ms-Grid-row'>
                                                                    <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12'>
                                                                        <Persona
                                                                            { ...{
                                                                                imageUrl: "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==",
                                                                                imageInitials: ""
                                                                            } }
                                                                            size={PersonaSize.size40}
                                                                            text="User Not Selected"
                                                                            secondaryText="Loan Officer"
                                                                        />

                                                                    </div>
                                                                </div>
                                                            </div>
                                                        </div>
                                                        :

                                                        this.state.LoanOfficer.length === 1 ?
                                                            this.state.LoanOfficer.map(officer =>
                                                                <EmployeeStatusCard key={officer.id}
                                                                    { ...{
                                                                        id: officer.id,
                                                                        name: officer.displayName,
                                                                        image: "",
                                                                        role: officer.assignedRole.adGroupName,
                                                                        status: this.state.oppData.opportunityState,
                                                                        isDispOppStatus: true
                                                                    }
                                                                    }
                                                                />
                                                            )
                                                            :
                                                            <GroupEmployeeStatusCard members={this.state.LoanOfficer} status={this.state.oppData.opportunityState} isDispOppStatus='true' role='Loan Officer' />
                                                }
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>

                        </div>
                        <div className=' ms-Grid-col ms-sm12 ms-md8 ms-lg3 p-l-10 TeamMembersBG'>
                            <h3>Team Members</h3>
                            <TeamMembers memberslist={this.state.teamMembers} createTeamId={this.state.oppId} opportunityState={this.state.oppData.opportunityState} userRole={this.state.userRole} />
                        </div>
                    </div>


                </div>

            );
        }
    }
}