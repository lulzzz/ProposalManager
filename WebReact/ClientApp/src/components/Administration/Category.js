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


export class Category extends Component {
    displayName = Category.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;

        this.utils = new Utils();
        const columns = [
            {
                key: 'column1',
                name: 'Category',
                headerClassName: 'ms-List-th browsebutton',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg8',
                fieldName: 'Category',
                minWidth: 150,
                maxWidth: 250,
                isRowHeader: true,
                onRender: (item) => {
                    return (
                        <TextField
                            id={'txtCategory' + item.id}
                            value={item.name}
                            onBlur={(e) => this.onBlurCategoryName(e, item, item.operation)}
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

        this.category = [];

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
        this.getCategories();
    }

    getCategories() {
        // call to API fetch Categories
        let requestUrl = 'api/Category';
        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
        })
            .then(response => response.json())
            .then(data => {
                try {
                    let categoryList = [];
                    for (let i = 0; i < data.length; i++) {
                        let category = {};
                        category.id = data[i].id;
                        category.name = data[i].name;
                        category.operation = "update";
                        categoryList.push(category);
                    }
                    this.setState({ items: categoryList, loading: false, rowItemCounter: categoryList.length });
                }
                catch (err) {
                    return false;
                }

            });
    }

    createCategoryItem(key) {
        return {
            id: key,
            name: "",
            operation:"add"
        };
    }

    onAddRow() {
        let rowCounter = this.state.rowItemCounter + 1;
        let newItems = [];
        newItems.push(this.createCategoryItem(rowCounter));

        let currentItems = this.state.items.concat(newItems);

        this.setState({
            items: currentItems,
            rowItemCounter: rowCounter
        });
    }

    deleteRow(item) {
        this.setState({ isUpdate: true });

        //deleteCategory
        this.deleteCategory(item);
    }

    //Category List - Details
    categoryList(columns, isCompactMode, items, selectionDetails) {
        return (
            <div className='ms-Grid-row LsitBoxAlign p20ALL '>
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

    onBlurCategoryName(e, item, operation) {
        this.setState({ isUpdate: true });
        delete item['operation'];
        let updatedItems = this.state.items;
        let itemIdx = updatedItems.indexOf(item);
        updatedItems[itemIdx].name = e.target.value;
        //this.category = updatedItems;
        //this.setState({
        //    items: updatedItems
        //});
        this.setState({
            updatedItems: updatedItems
        });

        if (operation === "add") {
            this.addCategory(updatedItems[itemIdx]);
        } else if (operation === "update") {
            this.updateCategory(updatedItems[itemIdx]);
        }
        
    }

    addCategory(categoryItem) {
        console.log(categoryItem);
        let categoriesObj = categoryItem;
        // API Update call        
        this.requestUpdUrl = 'api/Category';
        let options = {
            method: "POST",
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'authorization': 'Bearer    ' + window.authHelper.getWebApiToken()
            },
            body: JSON.stringify(categoriesObj)
        };

        fetch(this.requestUpdUrl, options)
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    this.setState({
                        items: this.state.updatedItems,
                        MessagebarText: "Category added successfully.",
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

    updateCategory(categoryItem) {
        console.log(categoryItem);
        let categoriesObj = categoryItem;
        // API Update call        
        this.requestUpdUrl = 'api/Category';
        let options = {
            method: "PATCH",
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'authorization': 'Bearer    ' + window.authHelper.getWebApiToken()
            },
            body: JSON.stringify(categoriesObj)
        };

        fetch(this.requestUpdUrl, options)
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    this.setState({
                        items: this.state.updatedItems,
                        MessagebarText: "Category updated successfully.",
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

    deleteCategory(categoryItem) {
        // API Update call        
        this.requestUpdUrl = 'api/Category/'+categoryItem.id;

        fetch(this.requestUpdUrl, {
            method: "DELETE",
            headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
        })
            .catch(error => console.error('Error:', error))
            .then(response => {
                if (response.ok) {
                    let currentItems = this.state.items.filter(x => x.id !== categoryItem.id);
                    this.category = currentItems;
                    this.setState({
                        items: currentItems,
                        MessagebarText: "Category deleted successfully.",
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
        const categoryList = this.categoryList(columns, isCompactMode, items, selectionDetails);
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
                            {categoryList}
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