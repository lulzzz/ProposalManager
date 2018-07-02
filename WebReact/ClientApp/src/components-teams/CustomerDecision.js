/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/
import React, { Component } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import { TeamsComponentContext, Panel, PanelBody, PanelFooter, PanelHeader} from 'msteams-ui-components-react';
import { DatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import {
    Spinner,
    SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { getQueryVariable } from '../common';

let teamContext = {};
const DayPickerStrings = {
    months: [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December'
    ],

    shortMonths: [
        'Jan',
        'Feb',
        'Mar',
        'Apr',
        'May',
        'Jun',
        'Jul',
        'Aug',
        'Sep',
        'Oct',
        'Nov',
        'Dec'
    ],

    days: [
        'Sunday',
        'Monday',
        'Tuesday',
        'Wednesday',
        'Thursday',
        'Friday',
        'Saturday'
    ],

    shortDays: [
        'S',
        'M',
        'T',
        'W',
        'T',
        'F',
        'S'
    ],

    goToToday: 'Go to today',
    prevMonthAriaLabel: 'Go to previous month',
    nextMonthAriaLabel: 'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year'
};

export class CustomerDecision extends Component {
    displayName = CustomerDecision.name
    constructor(props) {
        super(props);
        
        this.onChangedTxtApprovedDate = this.onChangedTxtApprovedDate.bind(this);
        this.onChangedTxtLoadDisbursed = this.onChangedTxtLoadDisbursed.bind(this);
        this.fnDdlCustomerApproved = this.fnDdlCustomerApproved.bind(this);

        this.state = {
            CustomerDecision: {},
            LoadDisbursed: "",
            ApprovedDate: "",
            ApprovedStatus: false,
            loading: true,
            oppData: [],
            isUpdate: false,
            MessagebarText: ""
        };

        
       
	}

    componentWillMount() {
        // condition to check loading from mobile view
        if (window.location.href.indexOf("tabMob") > -1) {
            let teamName = getQueryVariable('teamName');
            this.fnGetOpportunityData(teamName);
        } else {
            // API call to fetch details
            microsoftTeams.getContext(context => this.initialize(context));
        }
    }

	initialize({ groupId, channelName, teamName }) {
		
		let tc = {
			group: groupId,
			channel: channelName,
			team: teamName
		};
		teamContext = tc;

        this.fnGetOpportunityData(teamName);
	}


    fnGetOpportunityData(teamName) {
        // API - Fetch call
        this.requestUrl = "api/Opportunity?name='" + teamName + "'";
        fetch(this.requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + window.authHelper.getWebApiToken() }

        })
            .then(response => response.json())
            .then(data => {

                let customerDesionObj = data.customerDecision;
                this.setState({
                    loading: false,
                    CustomerDecision: customerDesionObj,
                    CustomerDecisionId: customerDesionObj.id,
                    LoadDisbursed: new Date(customerDesionObj.loanDisbursed),
                    ApprovedDate: new Date(customerDesionObj.approvedDate),
                    ApprovedStatus: customerDesionObj.approved,
                    oppData: data,
                    isUpdate: false
                });
            });
    }

    fnUpdateCustDecision(custDecisionObj) {
        this.setState({ isUpdate: true, MessagebarText: "Updating..." });

        let oppViewData = this.state.oppData;
        oppViewData.customerDecision = custDecisionObj;

        // API Update call        
        this.requestUpdUrl = 'api/opportunity?id=' + oppViewData.id;
        let options = {
            method: "PATCH",
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'authorization': 'Bearer    ' + window.authHelper.getWebApiToken()
            },
            body: JSON.stringify(oppViewData),
            id: this.props.match.params.id
        };

        fetch(this.requestUpdUrl, options)
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    return response.json;
                } else {
                    console.log('Error...: ');
                }
            }).then(json => {
                this.setState({ MessagebarText: "Updated successfully." });
                console.log(json);
                // this.setState({ isUpdate: false, MessagebarText: "" });
                setTimeout(function () { this.setState({ isUpdate: false, MessagebarText: "" }); }.bind(this), 3000);
            });
            

    }

    onChangedTxtApprovedDate(event) {
        this.setState(Object.assign({}, this.state, { txtApprovedDate: event.target.value }));
    }

    onChangedTxtLoadDisbursed(event) {
        this.setState(Object.assign({}, this.state, { txtLoadDisbursed: event.target.value }));
    }

    fnDdlCustomerApproved = (event) => {
        this.setState({ ApprovedStatus: event.key });
        let custDecisionObj = {
            "id": this.state.CustomerDecisionId,
            "approved": event.key,
            "approvedDate": this.state.ApprovedDate,
            "loanDisbursed": this.state.LoadDisbursed
        };
        this.fnUpdateCustDecision(custDecisionObj);
    }

    _onSelectApproved = (date) => {
        this.setState({ ApprovedDate: date });
        let custDecisionObj = {
            "id": this.state.CustomerDecisionId,
            "approved": this.state.ApprovedStatus,
            "approvedDate": date,
            "loanDisbursed": this.state.LoadDisbursed
        };
        this.fnUpdateCustDecision(custDecisionObj);
    }


    _onSelectLoanDisbursed = (date) => {
        this.setState({ LoadDisbursed: date });
        let custDecisionObj = {
            "id": this.state.CustomerDecisionId,
            "approved": this.state.ApprovedStatus,
            "approvedDate": this.state.ApprovedDate,
            "loanDisbursed": date
        };
        this.fnUpdateCustDecision(custDecisionObj);
    }

    _onFormatDate = (date) => {
        return (
            date.getMonth() + 1 +
            '/' +
            date.getDate() +
            '/' +
            date.getFullYear()
        );
    }

    _onParseDateFromString = (value) => {
        const date = this.state.value || new Date();
        const values = (value || '').trim().split('/');
        const day =
            values.length > 0
                ? Math.max(1, Math.min(31, parseInt(values[0], 10)))
                : date.getDate();
        const month =
            values.length > 1
                ? Math.max(1, Math.min(12, parseInt(values[1], 10))) - 1
                : date.getMonth();
        let year = values.length > 2 ? parseInt(values[2], 10) : date.getFullYear();
        if (year < 100) {
            year += date.getFullYear() - date.getFullYear() % 100;
        }
        return new Date(year, month, day);
    }

    _setItemDate(dt) {
        let lmDate = new Date(dt);
        if (lmDate.getFullYear() === 1 || lmDate.getFullYear() === 0) {
            return new Date();
        } else return new Date(dt);
    }


    renderContent(customerObj, isUpdate) {
        return (
            <div className='ms-Grid-row'>
            <div className='docs-TextFieldExample ms-Grid-col ms-sm6 ms-md8 ms-lg3'>
                <Dropdown
                    label='Customer Approved'
                    selectedKey={this.state.ApprovedStatus}
                    onChanged={this.fnDdlCustomerApproved}
                    options={
                        [
                            { key: true, text: "Yes" },
                            { key: false, text: "No" }
                        ]
                    }
                />
            </div>
            <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg3 docs-TextFieldExample'>
                <DatePicker strings={DayPickerStrings}
                    showWeekNumbers={false}
                    firstWeekOfYear={1}
                    showMonthPickerAsOverlay='true'
                    label="Approved Date"
                    placeholder='Approved Date'
                    iconProps={{ iconName: 'Calendar' }}
                    value={this._setItemDate(this.state.ApprovedDate)}
                    onSelectDate={this._onSelectApproved}
                    formatDate={this._onFormatDate}
                    parseDateFromString={this._onParseDateFromString}
                />
            </div>
            <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg3 docs-TextFieldExample'>
                <DatePicker strings={DayPickerStrings}
                    showWeekNumbers={false}
                    firstWeekOfYear={1}
                    showMonthPickerAsOverlay='true'
                    label="Loan Disbursed"
                    placeholder='Loan Disbursed'
                    iconProps={{ iconName: 'Calendar' }}
                    value={this._setItemDate(this.state.LoadDisbursed)}
                    onSelectDate={this._onSelectLoanDisbursed}
                    formatDate={this._onFormatDate}
                    parseDateFromString={this._onParseDateFromString}
                />
            </div>
            <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg2 hide'>
                {
                    this.state.isUpdate ?
                        <Spinner size={SpinnerSize.large} label='' ariaLive='assertive' className="pt15 pull-center" />
                        : ""
                }
            </div>
        </div>
        );
    }

    render() {
        let isUpdate = this.state.isUpdate;
        let content = this.state.loading
            ? <p><em>Loading...</em></p>
            : this.renderContent(this.state.CustomerDecision, isUpdate);

        if (this.state.loading) {
            return (
                <div className='ms-BasicSpinnersExample pull-center'>
                    <Spinner size={SpinnerSize.medium} label='loading...' ariaLive='assertive' />
                </div>
            );
        } else {
            return (
                <div>
                    <TeamsComponentContext>
                        <Panel>
                            <PanelHeader>
                                <h3 className="pl10">Customer Decision</h3>
                            </PanelHeader>
                            <PanelBody>
                                <div className='ms-Grid'>
                                    <div className='ms-Grid-row'>
                                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12 hide'>
                                            
                                            {
                                                isUpdate ?
                                                    <Spinner size={SpinnerSize.large} label='' ariaLive='assertive' className="pt15 pull-center" />
                                                    : ""
                                            }
                                        </div>
                                    </div>
                                </div>

                                <div className='ms-Grid'>
                                    {content}
                                </div>
                                <br /><br /><br /><br />

                            </PanelBody>
                            <PanelFooter>
                                <div className='ms-Grid'>
                                    <div className='ms-Grid-row'>
                                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg8'>
                                        </div>
                                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg4'>
                                            {this.state.isUpdate ?
                                                <MessageBar
                                                    messageBarType={MessageBarType.success}
                                                    isMultiline={false}
                                                >
                                                    {this.state.MessagebarText}
                                            </MessageBar>
                                            : ""
                                            }

                                        </div>
                                    </div>
                                </div>
                                

                            </PanelFooter>
                        </Panel>


                    </TeamsComponentContext>


                </div>
            );
        }
    }
}