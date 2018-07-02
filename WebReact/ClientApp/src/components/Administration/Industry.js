/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { DetailsList, DetailsListLayoutMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import Utils from '../../helpers/Utils';
import {
    Spinner,
    SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export class Industry extends Component {
    displayName = Industry.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;

        this.utils = new Utils();
        const columns = [
            {
                key: 'column1',
                name: 'Industry',
                headerClassName: 'ms-List-th browsebutton',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg8',
                fieldName: 'Region',
                minWidth: 150,
                maxWidth: 250,
                isRowHeader: true,
                onRender: (item) => {
                    return (
                        <TextField
                            id={'txtIndustry' + item.id}
                            value={item.name}
                            onBlur={(e) => this.onBlurIndustryName(e, item, item.operation)}
                        />
                    );
                }
            },
            {
                key: 'column2',
                name: 'Action',
                headerClassName: 'ms-List-th',
                className: 'ms-Grid-col ms-sm12 ms-md12 ms-lg4',
                minWidth: 16,
                maxWidth: 16,
                onRender: (item) => {
                    return (
                        <div>
                            <IconButton iconProps={{ iconName: 'Delete' }} onClick={e => this.deleteRow(item)} />
                        </div>
                    );
                }
            }
        ];

        let rowCounter = 0;

        this.state = {
            items: [],
            rowItemCounter: rowCounter,
            columns: columns,
            isCompactMode: false,
            loading: true,
            isUpdate: false,
            updatedItems: [],
            MessagebarText: "",
            MessageBarType: MessageBarType.success,
            isUpdateMsg: false
        };
    }

    componentWillMount() {
        this.getIndustries();
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
                        industry.id = data[i].id;
                        industry.name = data[i].name;
                        industry.operation = "update";
                        industryList.push(industry);
                    }
                    console.log(industryList);
                    this.setState({ items: industryList, loading: false, rowItemCounter: industryList.length });
                }
                catch (err) {
                    return false;
                }

            });
    }

    createItem(key) {
        return {
            id: key,
            name: "",
            operation: "add"
        };
    }

    onAddRow() {
        let rowCounter = this.state.rowItemCounter + 1;
        let newItems = [];
        newItems.push(this.createItem(rowCounter));

        let currentItems = this.state.items.concat(newItems);

        this.setState({
            items: currentItems,
            rowItemCounter: rowCounter
        });
    }

    deleteRow(item) {
        this.setState({ isUpdate: true });
        //deleteIndustry
        this.deleteIndustry(item);
    }


    industryList(columns, isCompactMode, items, selectionDetails) {
        return (
            <div className='ms-Grid-row LsitBoxAlign p20ALL'>
                <DetailsList
                    items={items}
                    compact={isCompactMode}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    selectionPreservedOnEmptyClick='true'
                    setKey='set'
                    layoutMode={DetailsListLayoutMode.justified}
                    enterModalSelectionOnTouch='false'
                />
            </div>
        );
    }

    onBlurIndustryName(e, item, operation) {
        this.setState({ isUpdate: true });
        delete item['operation'];

        let updatedItems = this.state.items;
        let itemIdx = updatedItems.indexOf(item);
        updatedItems[itemIdx].name = e.target.value;
        //this.industry = updatedItems;
        //this.setState({
        //    items: updatedItems
        //});
        this.setState({
            updatedItems: updatedItems
        });

        if (operation === "add") {
            this.addIndustry(updatedItems[itemIdx]);
        } else if (operation === "update") {
            this.updateIndustry(updatedItems[itemIdx]);
        }

    }

    addIndustry(industryItem) {
        console.log(industryItem);
        let industryObj = industryItem;
        // API Update call        
        this.requestUpdUrl = 'api/Industry';
        let options = {
            method: "POST",
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'authorization': 'Bearer    ' + window.authHelper.getWebApiToken()
            },
            body: JSON.stringify(industryObj)
        };

        fetch(this.requestUpdUrl, options)
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    this.setState({
                        items: this.state.updatedItems,
                        MessagebarText: "Industry added successfully.",
                        isUpdate: false,
                        isUpdateMsg: true
                    });
                    setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.success, MessagebarText: "" }); }.bind(this), 3000);
                    return response.json;
                } else {
                    this.setState({
                        MessagebarText: "Error occured. Please try again!",
                        isUpdate: false,
                        isUpdateMsg: true
                    });
                    setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.error, MessagebarText: "" }); }.bind(this), 3000);
                }
            }).then(json => {
                //console.log(json);
                this.setState({ isUpdate: false });
            });


    }

    updateIndustry(industryItem) {
        console.log(industryItem);
        let industryObj = industryItem;
        // API Update call        
        this.requestUpdUrl = 'api/Industry';
        let options = {
            method: "PATCH",
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'authorization': 'Bearer    ' + window.authHelper.getWebApiToken()
            },
            body: JSON.stringify(industryObj)
        };

        fetch(this.requestUpdUrl, options)
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    this.setState({
                        items: this.state.updatedItems,
                        MessagebarText: "Industry updated successfully.",
                        isUpdate: false,
                        isUpdateMsg: true
                    });
                    setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.success, MessagebarText: "" }); }.bind(this), 3000);

                    return response.json;
                } else {
                    this.setState({
                        MessagebarText: "Error occured. Please try again!",
                        isUpdate: false,
                        isUpdateMsg: true
                    });
                    setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.error, MessagebarText: "" }); }.bind(this), 3000);
                }
            }).then(json => {
                //console.log(json);
                this.setState({ isUpdate: false });
            });


    }

    deleteIndustry(industryItem) {
        console.log(industryItem);
        let industryObj = industryItem;
        delete industryObj['operation'];
        // API Update call        
        this.requestUpdUrl = 'api/Industry/' + industryObj.id;
        console.log(industryObj);

        fetch(this.requestUpdUrl, {
            method: "DELETE",
            headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
        })
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    let currentItems = this.state.items.filter(x => x.id !== industryObj.id);
                    this.industry = currentItems;
                    this.setState({
                        items: currentItems,
                        isUpdate: false
                    });
                    this.setState({
                        items: currentItems,
                        MessagebarText: "Industry deleted successfully.",
                        isUpdate: false,
                        isUpdateMsg: true
                    });

                    setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.success, MessagebarText: "" }); }.bind(this), 3000);
                    return response.json;
                } else {
                    this.setState({
                        MessagebarText: "Error occured. Please try again!",
                        isUpdate: false,
                        isUpdateMsg: true
                    });
                    setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.error, MessagebarText: "" }); }.bind(this), 3000);
                }
            }).then(json => {
                //console.log(json);
                this.setState({ isUpdate: false });
            });


    }


    render() {
        const { columns, isCompactMode, items, selectionDetails } = this.state;
        const industryList = this.industryList(columns, isCompactMode, items, selectionDetails);
        if (this.state.loading) {
            return (
                <div className='ms-BasicSpinnersExample ibox-content pt15 '>
                    <Spinner size={SpinnerSize.large} label='loading...' ariaLive='assertive' />
                </div>
            );
        } else {
            return (

                <div className='ms-Grid bg-white ibox-content'>
                   
                    <div className='ms-Grid-row'>
                        <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 pt10'>
                                <Link href='' className='pull-left' onClick={() => this.onAddRow()} >+ Add New</Link>
                            </div>
                            {industryList}
                        </div>
                    </div>
                    <div className='ms-Grid-row'>
                        <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                            <div className='ms-BasicSpinnersExample p-10'>
                                {
                                    this.state.isUpdate ?
                                        <Spinner size={SpinnerSize.large} ariaLive='assertive' />
                                        : ""
                                }
                                {
                                    this.state.isUpdateMsg ?
                                        <MessageBar
                                            messageBarType={this.state.MessageBarType}
                                            isMultiline={false}
                                        >
                                            {this.state.MessagebarText}
                                        </MessageBar>
                                        : ""
                                }
                            </div>
                        </div>
                    </div>
                </div>
            );

        }
    }

}