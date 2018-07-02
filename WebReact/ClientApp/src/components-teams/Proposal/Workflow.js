/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import * as ReactDOM from 'react-dom';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { LinkContainer } from 'react-router-bootstrap';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Glyphicon } from 'react-bootstrap';
import { TeamMembers } from '../../components/Opportunity/TeamMembers';
import { EmployeeStatusCard } from '../../components/Opportunity/EmployeeStatusCard';
import { GroupEmployeeStatusCard } from '../../components/Opportunity/GroupEmployeeStatusCard';
import {
    Spinner,
    SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';
import {
    Persona,
    PersonaSize,
    PersonaPresence
} from 'office-ui-fabric-react/lib/Persona';
import '../teams.css';
import { oppStatus } from '../../common';

export class Workflow extends Component {
    displayName = Workflow.name

    constructor(props) {
        super(props);
        this.authHelper = window.authHelper;
        this.sdkHelper = window.sdkHelper;

        this.state = {
            TeamMembers: [],
            UserRoleList: []
        };
    }

    componentWillMount() {
        this.getUserRoles();
    }
    getUserRoles() {
        // call to API fetch data
        let requestUrl = 'api/RoleMapping';
        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
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
        let loading = true;
        let oppDetails = this.props.oppDetails;
        let oppStatus = this.props.oppStaus;
        let teamMembersAll = [];
        teamMembersAll = this.props.memberslist;

        //if (this.props.memberslist.length > 0) {
        //    loading = false;
        //}

        let loanOfficerObj = teamMembersAll.filter(function (k) {
            return k.assignedRole.displayName === "LoanOfficer";
        });

        let relShipManagerObj = teamMembersAll.filter(function (k) {
            return k.assignedRole.displayName === "RelationshipManager";
        });


        // Get Other role officers list
        let otherRolesMapping = this.state.UserRoleList.filter(function (k) {
            return k.processType.toLowerCase() !== "base" && k.processType.toLowerCase() !== "administration" && k.processType.toLowerCase() !== "customerdecisiontab" && k.processType.toLowerCase() !== "proposalstatustab";
        });

        let otherRolesArr1 = [];
        for (let j = 0; j < otherRolesMapping.length; j++) {
            let teamMember = teamMembersAll.filter(function (k) {
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

        if (otherRolesObj.length > 0) {
            loading = false;
        }

        return (
            <div>
                {
                    loading ?
                        <div className='ms-BasicSpinnersExample pull-center'>
                            <br /><br />
                            <Spinner size={SpinnerSize.medium} label='loading...' ariaLive='assertive' />
                        </div>
                        :
                        <div className='ms-Grid'>
                            <div className='ms-Grid-row'>
                                <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12 p-r-10 '>
                                    <div className='ms-Grid-row'>
                                    </div>
                                    <div className='ms-Grid-row p-5 mt20'>
                                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3'>
                                            <div className='bg-gray newOpBg p20A'>
                                                <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                                &nbsp;&nbsp;<span>New Opportunity</span>
                                                {
                                                    relShipManagerObj.length > 0 ?
                                                        relShipManagerObj.map((officer, ind) =>
                                                            <EmployeeStatusCard key={ind}
                                                                {...{
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
                                        </div>
                                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3'>
                                            <div className='newOpBg p20A'>
                                                <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                                &nbsp;&nbsp;<span>Start Process</span>
                                                {
                                                    loanOfficerObj.length > 1 ?
                                                        <GroupEmployeeStatusCard members={loanOfficerObj} status={loanOfficerObj[0].status} isDispOppStatus='false' role='Loan Officer' />
                                                        :
                                                        loanOfficerObj.length === 1 ?
                                                            loanOfficerObj.map(officer =>
                                                                <EmployeeStatusCard key={officer.id}
                                                                    {...{
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
                                                                                {...{
                                                                                    imageUrl: "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==",
                                                                                    imageInitials: ""
                                                                                }}
                                                                                size={PersonaSize.size40}
                                                                                primaryText="User Not Selected"
                                                                                secondaryText="Loan Officer"
                                                                            />
                                                                        </div>
                                                                    </div>
                                                                </div>
                                                            </div>

                                                }

                                            </div>
                                        </div>
                                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg3 divUserRolegroup-arrow'>
                                            <div className='  divUserRolegroup'>
                                                
                                                <div>
                                                    {
                                                        otherRolesObj.length > 1 ?
                                                            <div>
                                                                {
                                                                    otherRolesObj.map((obj, ind) =>
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
                                                                                    &nbsp;&nbsp;<span>{otherRolesMapping[ind].processStep}</span>
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
                                            <i className="ms-Icon ms-Icon--ArrangeBringForward" aria-hidden="true"></i>
                                            &nbsp;&nbsp;<span>Draft Proposal</span>
                                            {
                                                loanOfficerObj.length > 1 ?
                                                    <GroupEmployeeStatusCard members={loanOfficerObj} status={oppStatus} isDispOppStatus='true' role='Loan Officer' />
                                                    :
                                                    loanOfficerObj.length === 1 ?
                                                        loanOfficerObj.map(officer =>
                                                            <EmployeeStatusCard key={officer.id}
                                                                {...{
                                                                    id: officer.id,
                                                                    name: officer.displayName,
                                                                    image: "",
                                                                    role: officer.assignedRole.adGroupName,
                                                                    status: oppStatus, //officer.status,
                                                                    isDispOppStatus: true
                                                                }
                                                                }
                                                            />
                                                        )
                                                        :
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
                                                                            {...{
                                                                                imageUrl: "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==",
                                                                                imageInitials: ""
                                                                            }}
                                                                            size={PersonaSize.size40}
                                                                            primaryText="User Not Selected"
                                                                            secondaryText="Loan Officer"
                                                                        />
                                                                    </div>
                                                                </div>
                                                            </div>
                                                        </div>
                                            }
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div className='ms-Grid-row'>
                                <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12  '>
                                    <hr />
                                </div>
                            </div>
                        </div>
                }
            </div>
        );
    }
}