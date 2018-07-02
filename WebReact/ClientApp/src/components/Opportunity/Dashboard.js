/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import { Link } from 'react-router-dom';
//import { Link } from 'office-ui-fabric-react/lib/Link';
import { PrimaryButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { oppStatusText, oppStatusClassName } from '../../common';
import '../../Style.css';
import { DetailsList, DetailsListLayoutMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';


export class Dashboard extends Component {
    displayName = Dashboard.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;

        //const userProfile = this.props.userProfile;
        const dashboardList = this.props.dashboardList;

        let columns = [
            {
                key: 'column1',
                name: 'Opportunity',
                headerClassName: 'ms-List-th',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg3',
                fieldName: 'name',
                minWidth: 150,
                maxWidth: 350,
                isRowHeader: true,
                isResizable: true,
                onColumnClick: this.onColumnClick,
                onRender: (item) => {
                    return (
                        <div className='ms-List-itemName'>
                            <Link to={'/OpportunitySummary?opportunityId=' + item.id} >
                                {item.opportunity}
                            </Link>
                        </div>
                    );
                }
            },
            {
                key: 'column2',
                name: 'Client',
                headerClassName: 'ms-List-th',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg3 clientcolum',
                fieldName: 'client',
                minWidth: 150,
                maxWidth: 350,
                isRowHeader: true,
                isResizable: true,
                onColumnClick: this.onColumnClick,
                onRender: (item) => {
                    return (
                        <div className='ms-List-itemClient'>{item.client}</div>
                    );
                },
                isPadded: true
            },
            {
                key: 'column3',
                name: 'Deal Size',
                headerClassName: 'ms-List-th',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg3 clientcolum',
                fieldName: 'client',
                minWidth: 150,
                maxWidth: 350,
                isRowHeader: true,
                isResizable: true,
                onColumnClick: this.onColumnClick,
                onRender: (item) => {
                    return (
                        <div className='ms-List-itemClient'>{item.dealsize}</div>
                    );
                },
                isPadded: true
            },
            {
                key: 'column4',
                name: 'Opened Date',
                headerClassName: 'ms-List-th',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg3',
                fieldName: 'openedDate',
                minWidth: 150,
                maxWidth: 350,
                isRowHeader: true,
                isResizable: true,
                onColumnClick: this.onColumnClick,
                onRender: (item) => {
                    return (
                        <div className='ms-List-itemDate AdminDate'>{item.openedDate}</div>
                    );
                },
                isPadded: true
            },
            {
                key: 'column5',
                name: 'Status',
                headerClassName: 'ms-List-th',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg2',
                fieldName: 'staus',
                minWidth: 150,
                maxWidth: 350,
                isRowHeader: true,
                isResizable: true,
                onColumnClick: this.onColumnClick,
                onRender: (item) => {
                    return (
                        <div className={oppStatusClassName[item.stausValue].toLowerCase()}>{oppStatusText[item.stausValue]}</div>
                    );
                },
                isPadded: true
            }
        ];

        const actionColumn = {
            key: 'column6',
            name: 'Action',
            headerClassName: 'ms-List-th delectBTNwidth',
            className: 'DetailsListExample-cell--FileIcon actioniconAlign ',
            minWidth: 30,
            maxWidth: 30,
            onColumnClick: this.onColumnClick,
            onRender: (item) => {
                return (
                    <div className='OpportunityDelete'>
                        <TooltipHost content='Delete' calloutProps={{ gapSpace: 0 }} closeDelay={200}>
                            <IconButton iconProps={{ iconName: 'Delete' }} onClick={e => this.deleteRow(item)} />
                        </TooltipHost>
                    </div>
                );
            }
        };

        if (this.props.userProfile.roles.filter(x => x.displayName === "RelationshipManager").length > 0) {
            columns.push(actionColumn);
        }


        this.state = {
            filterClient: '',
            filterDeal: '',
            items: dashboardList,
            itemsOriginal: dashboardList,
            loading: false,
            reverseList: false, //Seems there are issues with Reverse function on arrays
            authUserId: this.props.userProfile.id,
            authUserDisplayName: this.props.userProfile.displayName,
            authUserMail: this.props.userProfile.mail,
            authUserPhone: this.props.userProfile.phone,
            authUserPicture: this.props.userProfile.picture,
            authUserUPN: this.props.userProfile.userPrincipalName,
            authUserRoles: this.props.userProfile.roles,
            messageBarEnabled: false,
            messageBarText: "",
            MessagebarTextOpp: "",
            MessagebarTexCust: "",
            MessagebarTexDealSize: "",
            loadSpinner: true,
            columns: columns,
            isCompactMode: false,
            isDelteOpp: false,
            MessageDeleteOpp: "",
            MessageBarTypeDeleteOpp: ""
        };

        this._onFilterByNameChanged = this._onFilterByNameChanged.bind(this);
        this._onFilterByDealChanged = this._onFilterByDealChanged.bind(this);
    }

    fetchResponseHandler(response, referenceCall) {
        if (response.status === 401) {
            // TODO: This has been deprecated with the new token refresh functionality leaving the code for future expansion
        }
    }

    errorHandler(err, referenceCall) {
        console.log("Dashboard Ref: " + referenceCall + " error: " + JSON.stringify(err));
    }

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
                        items: [],
                        itemsOriginal: []
                    });
                    reject(err);
                });

        });
    }

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

    deleteRow(item) {
        let fetchData = {
            method: 'delete',
            //body: JSON.stringify(item.id),
            headers: {
                'authorization': 'Bearer ' + window.authHelper.getWebApiToken()
            }
        };
        this.requestUrl = 'api/opportunity/' + item.id;
        this.setState({ isDelteOpp: true, MessageDeleteOpp: " Deleting Opportunity - " + item.opportunity, MessageBarTypeDeleteOpp: MessageBarType.success });

        fetch(this.requestUrl, fetchData)
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    return response.json;
                } else {
                    //console.log('Error...: ');
                }
            }).then(json => {
                let currentItems = this.state.items.filter(x => x.id !== item.id);

                this.setState({
                    items: currentItems
                });
                this.setState({ MessageDeleteOpp: " Deleted Opportunity - " + item.opportunity });
                setTimeout(function () {
                    this.setState({ isDelteOpp: false, MessageDeleteOpp: "", MessageBarTypeDeleteOpp: MessageBarType.success });
                }.bind(this), 3000);
            });
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


        return (
            <div className='ms-List-itemCell' data-is-focusable='true'>
                <div className='ms-List-itemContent'>
                    <div className='ms-List-itemName'>
                        <Link to={'/OpportunitySummary?opportunityId=' + item.id} >
                            {item.opportunity}
                        </Link>
                    </div>
                    <div className='ms-List-itemClient'>{item.client}</div>
                    <div className='ms-List-itemDealsize'>{item.dealsize}</div>
                    <div className='ms-List-itemDate'>{item.openedDate}</div>
                    <div className={"ms-List-itemState " + oppStatusClassName[item.stausValue].toLowerCase()}>{oppStatusText[item.stausValue]}</div>
                    <div className="OpportunityDelete ">
                        <IconButton iconProps={{ iconName: 'Delete' }} onClick={e => this.deleteRow(item)} />
                    </div>
                </div>
            </div>
        );
    }

    opportunitiesList(itemsList, itemsListOriginal) {
        //const lenght = typeof itemsList !== 'undefined' ? itemsList.length : 0;
        //const lenghtOriginal = typeof itemsListOriginal !== 'undefined' ? itemsListOriginal.length : 0;
        //const originalItems = itemsListOriginal;
        const items = itemsList;
        //const resultCountText = lenght === lenghtOriginal ? '' : ` (${items.length} of ${originalItems.length} shown)`;

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

    render() {
        const { columns, isCompactMode, items } = this.state;

        const isLoading = this.state.loading;

        let isRelationshipManager = false;
        if ((this.state.authUserRoles.filter(x => x.displayName === "RelationshipManager")).length > 0) {
            isRelationshipManager = true;
        }

        const itemsOriginal = this.state.itemsOriginal;
        //const items = this.state.items;

        const lenghtOriginal = typeof itemsOriginal !== 'undefined' ? itemsOriginal.length : 0;
        const listHasItems = lenghtOriginal > 0 ? true : false;

        const opportunitiesListHeading = this.opportunitiesListHeading();
        const opportunitiesListComponent = this.opportunitiesList(items, itemsOriginal);


        return (
            <div className='ms-Grid'>
                {
                    this.state.messageBarEnabled ?
                        <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                            <MessageBar messageBarType={this.props.context.messageBarType} isMultiline={false}>
                                {this.props.context.messageBarText}
                            </MessageBar>
                        </div>
                        : ""
                }

                
                <div className='ms-Grid-row'>
                    <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 pageheading'>
                        <h3>Dashboard</h3>
                    </div>
                    {
                        isRelationshipManager ?
                            isLoading ?
                                <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg4 createButton pull-right '>
                                  
                                    <Spinner size={SpinnerSize.small} label='Loading actions...' ariaLive='assertive' />
                                </div>
                                :
                                <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 createButton pt15 '>
                                    <PrimaryButton className='pull-right' onClick={this.props.onClickCreateOpp}> <i className="ms-Icon ms-Icon--Add pr10" aria-hidden="true"></i> Create New</PrimaryButton>
                                </div>
                            :
                            <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 createButton ' />
                    }
                </div>
                <div className='ms-Grid'>
                    <div className='ms-Grid-row ms-SearchBoxSmallExample'>
                        <div className='ms-Grid-col ms-sm4 ms-md4 ms-lg3 pl0'>
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
                        <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                            {
                                this.state.isDelteOpp ?
                                    <MessageBar messageBarType={this.state.MessageBarTypeDeleteOpp} isMultiline={false}>
                                        {this.state.MessageDeleteOpp}
                                    </MessageBar>
                                    : ""
                            }
                        </div>
                    </div>
                    <div className='ms-Grid-row'>
                        {
                            items.length > 0
                                ?
                                <DetailsList
                                    items={items}
                                    compact={isCompactMode}
                                    columns={columns}
                                    selectionMode={SelectionMode.none}
                                    setKey='key'
                                    layoutMode={DetailsListLayoutMode.justified}
                                    enterModalSelectionOnTouch='false'
                                />
                                :
                                <div>There are no opportunities.</div>
                        }
                        
                    </div>
                    <br /><br />
                </div>
            </div>
        );
    }

}