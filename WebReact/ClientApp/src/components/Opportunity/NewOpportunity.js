/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import Utils from '../../helpers/Utils';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';


export class NewOpportunity extends Component {
    displayName = NewOpportunity.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;
        this.utils = new Utils();

        this.state = {
            industryList: this.props.industries,
            regionList: this.props.regions,
            custNameError: false,
            messagebarTextCust: "",
            oppNameError: false,
            messagebarTextOpp: "",
            dealSizeError: false,
            messagebarTextDealSize: "",
            annualRevenueError: false,
            messagebarTextAnnualRev: "",
            nextDisabled: true
        };
    }

    componentWillMount() {
        this.opportunity = this.props.opportunity;
        this.dashboardList = this.props.dashboardList;
    }


    // Class methods
    onBlurCustomer(e) {
        if (e.target.value.length === 0) {
            this.setState({
                messagebarTextCust: "Customer name can not be empty.",
                custNameError: false
            });
            this.opportunity.customer.displayName = "";
        } else {
            this.setState({
                messagebarTextCust: "",
                custNameError: false
            });
            this.opportunity.customer.displayName = e.target.value;
        }
    }

    onBlurOpportunityName(e) {
        if (e.target.value.length > 0) {
            this.opportunity.displayName = e.target.value;
            this.setState({
                messagebarTextOpp: "",
                oppNameError: false
            });
            // let uniqueResponse = this.oppNameIsUnique(e.target.value);
        } else {
            this.opportunity.displayName = "";
            this.setState({
                messagebarTextOpp: "Opportunity name can not be empty.",
                oppNameError: false
            });
        }
    }

    onBlurDealSize(e) {
        this.opportunity.dealSize = e.target.value;
    }

    onBlurAnnualRevenue(e) {
        this.opportunity.annualRevenue = e.target.value;
    }

    onChangeIndustry(e) {
        this.opportunity.industry.id = e.key;
        this.opportunity.industry.name = e.text;
    }

    onChangeRegion(e) {
        this.opportunity.region.id = e.key;
        this.opportunity.region.name = e.text;
    }

    onBlurNotes(e) {
        // TODO: Add createdby propeties
        let note = {
            id: this.utils.guid(),
            noteBody: e.target.value,
            createdDateTime: "",
            createdBy: {
                id: "",
                displayName: "",
                userPrincipalName: "",
                userRoles: []
            }
        };

        this.opportunity.notes.push(note);
    }

    oppNameIsUnique(name) {
        if (this.opportunity.displayName.length > 0) {
            if (this.dashboardList.find(itm => itm.opportunity === name)) {
                this.setState({
                    messagebarTextOpp: "Opportunity name must be unique.",
                    oppNameError: false
                });
                return true;
            } else {
                return false;
            }
        } else {
            // If empty also return false
            return false;
        }
    }

    render() {
        let nextDisabled = true;
        if (this.opportunity.customer.displayName.length > 0 && this.opportunity.displayName.length > 0) {
            nextDisabled = false;
        }

        //TODO: set focus on initial load of component: this.customerName.focusInput()

        return (
            <div className='ms-Grid'>
                <div className='ms-Grid-row'>
                    <h3 className='pageheading'>Create New Opportunity</h3>
                    <div className='ms-lg12 ibox-content'>
                        <div className="ms-Grid-row">
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <TextField
                                    id='customerName'
                                    label='Customer Name' value={this.opportunity.customer.displayName}
                                    errorMessage={this.state.messagebarTextCust}
                                    onBlur={(e) => this.onBlurCustomer(e)}
                                />
                                {this.state.custNameError ?
                                    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                                        {this.state.messagebarTextCust}
                                    </MessageBar>
                                    : ""
                                }
                            </div>


                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <TextField
                                    label='Opportunity Name' value={this.opportunity.displayName}
                                    errorMessage={this.state.messagebarTextOpp}
                                    onBlur={(e) => this.onBlurOpportunityName(e)}

                                />
                                {this.state.oppNameError ?
                                    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                                        {this.state.messagebarTextOpp}
                                    </MessageBar>
                                    : ""
                                }
                            </div>
                        </div>
                        <div className="ms-Grid-row">
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <TextField
                                    label='Deal Size' value={this.opportunity.dealSize}
                                    onBlur={(e) => this.onBlurDealSize(e)}
                                />
                                {this.state.dealSizeError ?
                                    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                                        {this.state.messagebarTextDealSize}
                                    </MessageBar>
                                    : ""
                                }
                            </div>
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <TextField
                                    label='Annual Revenue' value={this.opportunity.annualRevenue}
                                    onBlur={(e) => this.onBlurAnnualRevenue(e)}

                                />
                                {this.state.annualRevenueError ?
                                    <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
                                        {this.state.messagebarTextAnnualRev}
                                    </MessageBar>
                                    : ""
                                }
                            </div>
                        </div>

                        <div className="ms-Grid-row">
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <Dropdown
                                    placeHolder='Select Industry'
                                    label='Industry'
                                    id='Basicdrop1'
                                    ariaLabel='Industry'
                                    value={this.opportunity.industry.id}
                                    options={this.state.industryList}
                                    defaultSelectedKey={this.opportunity.industry.id}
                                    componentRef={this.ddlIndustry}
                                    onChanged={(e) => this.onChangeIndustry(e)}
                                />
                            </div>
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg6'>
                                <Dropdown
                                    placeHolder='Select Region'
                                    label='Region'
                                    id='ddlRegion'
                                    ariaLabel='Region'
                                    value={this.opportunity.region.id}
                                    options={this.state.regionList}
                                    defaultSelectedKey={this.opportunity.region.id}
                                    componentRef=''
                                    onChanged={(e) => this.onChangeRegion(e)}
                                />
                            </div>
                        </div>
                        <div className="ms-Grid-row">
                            <div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                                <TextField
                                    label='Notes'
                                    multiline
                                    rows={6}
                                    value={this.opportunity.notes.noteBody}
                                    onBlur={(e) => this.onBlurNotes(e)}
                                />
                            </div>
                        </div>
                    </div>
                </div>
                <div className='ms-Grid-row pb20'>
                    <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 pl0'><br />
                        <PrimaryButton className='backbutton pull-left' onClick={this.props.onClickCancel}>Cancel</PrimaryButton>
                    </div>
                    <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 pr0'><br />
                        <PrimaryButton className='pull-right' onClick={this.props.onClickNext} disabled={nextDisabled}>Next</PrimaryButton>
                    </div>
                </div>
            </div>
        );
    }
}