/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { PeoplePickerTeamMembers } from '../PeoplePickerTeamMembers';



export class NewOpportunityOthers extends Component {
    displayName = NewOpportunityOthers.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;

        this.opportunity = this.props.opportunity;

        this.state = {
            showModal: false,
            currentPicker: 1,
            delayResults: false,
            teamMembers: [],
            teamMembersAll: []
        };
    }

    componentWillMount() {
        this.filterUserProfiles();
    }

    // Class methods
    filterUserProfiles() {
        let data = this.props.teamMembers;
        let teamlist = [];

        for (let i = 0; i < data.length; i++) {
            let item = data[i];

            if ((item.userRoles.filter(x => x.displayName === "LoanOfficer")).length > 0) {
                teamlist.push(item);
            }
        }

        this.setState({
            teamMembers: teamlist,
            teamMembersAll: data
        });
    }

    getSelectedUsers() {
        let selectedUsers = this.opportunity.teamMembers.filter(x => x.assignedRole.displayName === "LoanOfficer");

        return selectedUsers;
    }

    onBlurMargin(e) {
        this.opportunity.margin = e.target.value;
    }

    onChangeLoanOfficer(value) {
        if (value.length > 0) {
            let updatedTeamMembers = this.opportunity.teamMembers.filter(x => x.assignedRole.displayName !== "LoanOfficer");

            let newMember = {};
            newMember.id = value[0].id;
            newMember.displayName = value[0].text;
            newMember.mail = value[0].mail;
            newMember.userPrincipalName = value[0].userPrincipalName;
            newMember.userRoles = value[0].userRoles;
            newMember.status = 0;
            newMember.assignedRole = value[0].userRoles.filter(x => x.displayName === "LoanOfficer")[0];

            updatedTeamMembers.push(newMember);
            this.opportunity.teamMembers = updatedTeamMembers;
        }
    }

    onBlurRate(e) {
        this.opportunity.rate = e.target.value;
    }

    onBlurDebtRatio(e) {
        this.opportunity.debtRatio = e.target.value;
    }

    onBlurPurpose(e) {
        this.opportunity.purpose = e.target.value;
    }

    onBlurDisbursementSchedule(e) {
        this.opportunity.disbursementSchedule = e.target.value;
    }

    onBlurCollateralAmount(e) {
        this.opportunity.collateralAmount = e.target.value;
    }

    onBlurGuarantees(e) {
        this.opportunity.guarantees = e.target.value;
    }

    onBlurRiskRating(e) {
        this.opportunity.riskRating = e.target.value;
    }


    render() {
        let selectedUsers = this.getSelectedUsers();

        return (
            <div>
                <div className='ms-Grid'>
                    <div className='ms-grid-row'>
                        <h3 className='pageheading'>Customer Input</h3>
                        <div className='ms-lg12 ibox-content'>
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg4'>
                                <TextField
                                    label='Margin ($M)'
                                    value={this.opportunity.margin}
                                    onBlur={(e) => this.onBlurMargin(e)}
                                />
                            </div>
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg4'>
                                <TextField
                                    label='Rate'
                                    value={this.opportunity.rate}
                                    onBlur={(e) => this.onBlurRate(e)}
                                />
                            </div>
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg4'>
                                <TextField
                                    label='Debt Ratio'
                                    value={this.opportunity.debtRatio}
                                    onBlur={(e) => this.onBlurDebtRatio(e)}
                                />

                            </div>
                        </div>
                    </div>
                </div>

                <div className='ms-Grid'>
                    <div className='ms-grid-row'>
                        <h3 className="pageheading">Credit Facility</h3>
                        <div className='ms-lg12 ibox-content pb20'>
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <TextField
                                    label='Purpose'
                                    value={this.opportunity.purpose}
                                    onBlur={(e) => this.onBlurPurpose(e)}
                                />
                            </div>
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <TextField
                                    label='Disbursement Schedule'
                                    value={this.opportunity.disbursementSchedule}
                                    onBlur={(e) => this.onBlurDisbursementSchedule(e)}
                                />
                            </div>

                            <div className='ms-grid-row'>
                                <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                    <TextField
                                        label='Collateral Amount'
                                        value={this.opportunity.collateralAmount}
                                        onBlur={(e) => this.onBlurCollateralAmount(e)}
                                    />
                                </div>
                                <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                    <TextField
                                        label='Guarantees'
                                        value={this.opportunity.guarantees}
                                        onBlur={(e) => this.onBlurGuarantees(e)}
                                    />
                                </div>
                            </div>
                            <div className='ms-grid-row'>
                                <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                    <TextField
                                        label='Risk Rating'
                                        value={this.opportunity.riskRating}
                                        onBlur={(e) => this.onBlurRiskRating(e)}
                                    />
                                </div>
                                <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                    &nbsp;
								</div>
                            </div>
                        </div>
                    </div>
                </div>

                <div className='ms-Grid'>
                    <div className='ms-grid-row'>
                        <h3 className="pageheading">Loan Officer</h3>
                        <div className='ms-lg12 ibox-content pb20'>
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <PeoplePickerTeamMembers teamMembers={this.state.teamMembers} defaultSelectedUsers={selectedUsers} onChange={(e) => this.onChangeLoanOfficer(e)} />
                            </div>
                        </div>
                    </div>
                    <div className='ms-grid-row '>
                        <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 pb20'><br />
                            <PrimaryButton className='backbutton pull-left' onClick={this.props.onClickBack}>Back</PrimaryButton>
                        </div>
                        <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 pb20'><br />
                            <PrimaryButton className='pull-right' onClick={this.props.onClickNext}>Submit</PrimaryButton>
                        </div>
                    </div><br /><br />
                </div>
            </div>
        );
    }
}