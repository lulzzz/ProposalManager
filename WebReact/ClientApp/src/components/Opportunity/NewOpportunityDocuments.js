/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { PrimaryButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { DetailsList, DetailsListLayoutMode, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { FilePicker } from '../FilePicker';
import Utils from '../../helpers/Utils';
import '../../Style.css';


export class NewOpportunityDocuments extends Component {
    displayName = NewOpportunityDocuments.name

    constructor(props) {
        super(props);

        this.utils = new Utils();
        this.opportunity = this.props.opportunity;

        const columns = [
            {
                key: 'column1',
                name: 'File',
                headerClassName: 'ms-List-th browsebutton',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg4 browsebutton',
                fieldName: 'file',
                minWidth: 150,
                maxWidth: 250,
                isRowHeader: true,
                onRender: (item) => {
                    //if (item.file.name)
                    let itemFileUri = "";
                    return (
                        <FilePicker
                            id={'fp' + item.id}
                            fileUri={itemFileUri}
                            file={item.file}
                            showBrowse='true'
                            showLabel='true'
                            onChange={(e) => this.onChangeFile(e, item)}
                        />
                    );
                }
            },
            {
                key: 'column2',
                name: 'Notes',
                headerClassName: 'ms-List-th',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg3',
                fieldName: 'notes',
                minWidth: 150,
                maxWidth: 550,
                isRowHeader: false,
                isResizable: true,
                isCollapsable: true,
                onRender: (item) => {
                    return (
                        <TextField
                            id={'txtNotes' + item.id}
                            value={item.note}
                            onBlur={(e) => this.onBlurNotes(e, item)}
                        />
                    );
                },
                isPadded: true
            },
            {
                key: 'column3',
                name: 'Category',
                headerClassName: 'ms-List-th',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg2 categoryResponssive',
                fieldName: 'category',
                minWidth: 150,
                maxWidth: 550,
                isRowHeader: false,
                isResizable: true,
                onRender: (item) => {
                    return (
                        <div>
                            <Dropdown
                                id={'ddCat' + item.id}
                                ariaLabel='Category'
                                options={this.props.categories}
                                defaultSelectedKey={item.category.id}
                                onChanged={(e) => this.onChangeCategory(e, item)}
                            />
                        </div>
                    );
                },
                isPadded: true
            },
            {
                key: 'column4',
                name: 'Tags',
                headerClassName: 'ms-List-th',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg2 tagsfield',
                fieldName: 'tags',
                minWidth: 150,
                maxWidth: 550,
                isRowHeader: false,
                isResizable: true,
                isCollapsable: true,
                onRender: (item) => {
                    return (
                        <TextField
                            id={'txtTags' + item.id}
                            value={item.tags}
                            onBlur={(e) => this.onBlurTags(e, item)}
                        />
                    );
                },
                isPadded: true
            },
            {
                key: 'column5',
                name: 'Action',
                headerClassName: 'ms-List-th',
                className: 'DetailsListExample-cell--FileIcon',
                iconClassName: 'DetailsListExample-Header-FileTypeIcon',
                iconName: 'Page',
                //isIconOnly: true,
                //fieldName: 'name',
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

        let rowCounter = 1;
        if (this.opportunity.documentAttachments.length > 0) {
            rowCounter = this.opportunity.documentAttachments.length + 1;
        } else {
            let currentItems = this.opportunity.documentAttachments;
            currentItems.push(this.createListItem(rowCounter));
            this.opportunity.documentAttachments = currentItems;
        }

        this.state = {
            items: this.opportunity.documentAttachments,
            rowItemCounter: rowCounter,
            columns: columns,
            isCompactMode: false
        };
    }


    // Class methods
    onAddRow() {
        let rowCounter = this.state.rowItemCounter + 1;
        let newItems = [];
        newItems.push(this.createListItem(rowCounter));

        let currentItems = newItems.concat(this.state.items);

        this.opportunity.documentAttachments = currentItems;

        this.setState({
            items: currentItems,
            rowItemCounter: rowCounter
        });
    }

    deleteRow(item) {
        let currentItems = this.state.items.filter(x => x.id !== item.id);

        this.opportunity.documentAttachments = currentItems;
        this.setState({
            items: currentItems
        });
    }

    createListItem(key) {
        return {
            key: key,
            id: this.utils.guid(),
            file: {},
            fileName: "",
            note: "",
            category: {
                id: "",
                displayName: ""
            },
            tags: "",
            documentUri: ""
        };
    }

    onChangeFile(e, item) {
        let updatedItems = this.state.items;
        let itemIdx = updatedItems.indexOf(item);
        updatedItems[itemIdx].file = e;
        updatedItems[itemIdx].fileName = e.name;
        this.opportunity.documentAttachments = updatedItems;
        this.setState({
            items: updatedItems
        });
    }

    onChangeCategory(e, item) {
        let updatedItems = this.state.items;
        let itemIdx = updatedItems.indexOf(item);
        updatedItems[itemIdx].category.id = e.key;
        updatedItems[itemIdx].category.name = e.text;
        this.opportunity.documentAttachments = updatedItems;
        this.setState({
            items: updatedItems
        });
    }

    onBlurNotes(e, item) {
        let updatedItems = this.state.items;
        let itemIdx = updatedItems.indexOf(item);
        updatedItems[itemIdx].note = e.target.value;
        this.opportunity.documentAttachments = updatedItems;
        this.setState({
            items: updatedItems
        });
    }

    onBlurTags(e, item) {
        let updatedItems = this.state.items;
        let itemIdx = updatedItems.indexOf(item);
        updatedItems[itemIdx].tags = e.target.value;
        this.opportunity.documentAttachments = updatedItems;
        this.setState({
            items: updatedItems
        });
    }


    // For DeatlsList
    documentsList(columns, isCompactMode, items, selectionDetails) {
        //selection={this.selection}
        return (
            <div className='ms-Grid-row ibox-content'>
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

    onColumnClick = (ev, column) => {
        const { columns, items } = this.state;
        let newItems = items.slice();
        const newColumns = columns.slice();
        const currColumn = newColumns.filter((currCol, idx) => {
            return column.key === currCol.key;
        })[0];

        newColumns.forEach((newCol) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                currColumn.isSorted = true;
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });

        newItems = this.sortItems(newItems, currColumn.fieldName, currColumn.isSortedDescending);

        this.setState({
            columns: newColumns,
            items: newItems
        });
    }

    sortItems = (items, sortBy, descending = false) => {
        if (descending) {
            return items.sort((a, b) => {
                if (a[sortBy] < b[sortBy]) {
                    return 1;
                }
                if (a[sortBy] > b[sortBy]) {
                    return -1;
                }
                return 0;
            });
        } else {
            return items.sort((a, b) => {
                if (a[sortBy] < b[sortBy]) {
                    return -1;
                }
                if (a[sortBy] > b[sortBy]) {
                    return 1;
                }
                return 0;
            });
        }
    }

    getSelectionDetails() {
        const selectionCount = this.selection.getSelectedCount();
        switch (selectionCount) {
            case 0:
                return 'No items selected';
            case 1:
                return '1 item selected: ' + this.selection.getSelection()[0].name;
            default:
                return `${selectionCount} items selected`;
        }
    }


    render() {
        const { columns, isCompactMode, items, selectionDetails } = this.state;
        const documentsList = this.documentsList(columns, isCompactMode, items, selectionDetails);

        return (
            <div className='ms-Grid'>
                <div className='ms-Grid-row'>
                    <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6 pageheading'>
                        <h3>Add Documents</h3>
                    </div>
                    <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6 pt15 pr18 '>
                        <h5><Link href='' className='pull-right' onClick={() => this.onAddRow()} >+ Add New</Link></h5>
                    </div>
                </div>
                {documentsList}
                <div className='ms-grid-row '>
                    <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 pl0 pb20'><br />
                        <PrimaryButton className='backbutton pull-left' onClick={this.props.onClickBack}>Back</PrimaryButton>
                    </div>
                    <div className='ms-Grid-col ms-sm6 ms-md6 ms-lg6 pb20'><br />
                        <PrimaryButton className='pull-right' onClick={this.props.onClickNext}>Next</PrimaryButton>
                    </div>
                </div><br /><br />
            </div>
        );
    }
}