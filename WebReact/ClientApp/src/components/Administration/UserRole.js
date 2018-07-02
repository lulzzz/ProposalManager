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
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';


export class UserRole extends Component {

	displayName = UserRole.name
	
	constructor(props) {
		super(props);

		this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;

        this.processTypes = [
            {
                "key": 1,
                "text": "Base"
            },
            {
                "key": 2,
                "text": "Administration"
            },
            {
                "key": 3,
                "text": "ChecklistTab"
            },
            {
                "key": 4,
                "text":"CustomerDecisionTab"
            },
            {
                "key": 5,
                "text": "ProposalStatusTab"
            }
        ];


		this.utils = new Utils();
		const columns = [
			{
				key: 'column1',
				name: 'ADGroupName',
				headerClassName: 'ms-List-th browsebutton',
				className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg8',
				fieldName: 'ADGroup',
				minWidth: 150,
				maxWidth: 200,
				isRowHeader: true,
				onRender: (item) => {
					return (
						
                        <TextField
                            id={'txtADGroup' + item.id}
                            value={item.adGroupName}
                            onBlur={(e) => this.onBlurUserRole(e, item)}
                        />
					);
				}
			},
			{
				key: 'column2',
				name: 'RoleName',
				headerClassName: 'ms-List-th browsebutton',
				className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg8',
				fieldName: 'Role',
				minWidth: 150,
				maxWidth: 150,
				isRowHeader: true,
				onRender: (item) => {
					return (
                        
                            <TextField
                                id={'txtRole' + item.id}
                                value={item.roleName}
                                onBlur={(e) => this.onBlurUserRole(e, item)}
                            />
					);
				}
			},
			{
				key: 'column3',
				name: 'ProcessStep',
				headerClassName: 'ms-List-th browsebutton',
				className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg8',
				fieldName: 'Process',
				minWidth: 150,
				maxWidth: 150,
				isRowHeader: true,
				onRender: (item) => {
					return (
                        
                            <TextField
                                id={'txtProcessStep' + item.id}
                                value={item.processStep}
                                onBlur={(e) => this.onBlurUserRole(e, item)}
                            />
					);
				}
			},
			{
				key: 'column4',
				name: 'Process Type',
				headerClassName: 'ms-List-th browsebutton',
				className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg8',
                fieldName: 'processType',
				minWidth: 150,
				maxWidth: 200,
				isRowHeader: true,
                onRender: (item) => {
                    return (
                        <Dropdown
                            id={'txtProcessType' + item.id}
                            ariaLabel='ProcessType'
                            options={this.processTypes}
                            defaultSelectedKey={item.ddlprocessType.id}
                            onChanged={(e) => this.onBlurUserRole(e, item)}
                            
                        />
					);
				}
			},
			{
				key: 'column5',
				name: 'Channel',
				headerClassName: 'ms-List-th browsebutton chanelheading',
                className: 'docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg8 chanelheading-Column',
				fieldName: 'Channel',
				minWidth: 150,
				maxWidth: 300,
				isRowHeader: true,
				onRender: (item) => {
					return (
                        
                            <TextField
                                id={'txtChannel' + item.id}
                                value={item.channel}
                                onBlur={(e) => this.onBlurUserRole(e, item)}
                            />
					);
				}
			},
			{
				key: 'column6',
				name: 'Action',
				headerClassName: 'ms-List-th ActionHeading',
				className: 'ms-Grid-col ms-sm12 ms-md12 ms-lg4 ActionDelete',
				minWidth: 10,
				maxWidth: 10,
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
			loading: true
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
						userRole.adGroupName = data[i].adGroupName;
						userRole.roleName = data[i].roleName;
						userRole.processStep = data[i].processStep;
						userRole.channel = data[i].channel;
						userRole.adGroupId = data[i].adGroupId;
                        userRole.processType = data[i].processType;
                        userRole.ddlprocessType = {};
                        switch (data[i].processType.toLowerCase()) {
                            case "base":
                                userRole.ddlprocessType.id = 1;
                                userRole.ddlprocessType.processType = data[i].processType;
                                userRole.ddlprocessType.disabled = true;
                                break;
                            case "administration":
                                userRole.ddlprocessType.id = 2;
                                userRole.ddlprocessType.processType = data[i].processType;
                                userRole.ddlprocessType.disabled = true;
                                break;
                            case "checklisttab":
                                userRole.ddlprocessType.id = 3;
                                userRole.ddlprocessType.processType = data[i].processType;
                                userRole.ddlprocessType.disabled = false;
                                break;
                            case "customerdecisiontab":
                                userRole.ddlprocessType.id = 4;
                                userRole.ddlprocessType.processType = data[i].processType;
                                userRole.ddlprocessType.disabled = true;
                                break;
                            case "proposalstatustab":
                                userRole.ddlprocessType.id = 5;
                                userRole.ddlprocessType.processType = data[i].processType;
                                userRole.ddlprocessType.disabled = true;
                                break;
                            
                            default:
                                userRole.ddlprocessType.id = 3;
                                userRole.ddlprocessType.processType = "checklistTab";
                                userRole.ddlprocessType.disabled = false;
                                break;
                        }
                        
						userRole.type = "old";
						userRoleList.push(userRole);
					}
					this.setState({ items: userRoleList, loading: false, rowItemCounter: userRoleList.length });
				}
				catch (err) {
					return false;
				}

			});
	}

	createItem(key) {
		return {
			id: key,
			adGroupName: "",
			roleName: "",
			processStep:"",
			channel:"",
            processType: "",
            ddlprocessType: {
                id: "",
                processType: ""
            },
			type: "new"
				
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
        // check atleast one cheklist type should be exist
        let chklstTypeArr = this.state.items.filter(x => x.processType.toLowerCase() === "checklisttab");
        if (chklstTypeArr.length === 1) {
            alert("There should be at least on CheklistTab of Process Type.");
            return false;
        }
		
		let currentItems = this.state.items.filter(x => x.id !== item.id);

		this.userRole = currentItems;
		
		this.deleteItem(item.id);
		this.setState({
			items: currentItems,
			MessagebarText: "User Role Mapping deleted successfully.",
			isUpdate: false,
			isUpdateMsg: true
		});

		setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.success, MessagebarText: "" }); }.bind(this), 3000);
		this.getUserRoles();
	}

	deleteItem(id) {
		return new Promise((resolve, reject) => {

			let requestUrl = 'api/RoleMapping/' + id;
			
			let options = {
				method: "DELETE",
				headers: {
					'Accept': 'application/json',
					'Content-Type': 'application/json',
					'authorization': 'Bearer ' + this.authHelper.getWebApiToken()
				}
			};

			fetch(requestUrl, options)
				.then(response => {
					console.log("Delete User Role Mapping response: " + response.status + " - " + response.statusText);
					if (response.status === 401) {
						// TODO: For v2 see how we pass to authHelper to force token refresh
					}
					return response;
				})
				.then(data => {

					resolve(data);
				})
				.catch(err => {
					//this.errorHandler(err, "");
					this.setState({
						MessagebarText: "Error occured. Please try again!",
						isUpdate: false,
						isUpdateMsg: true
					});
					setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.error, MessagebarText: "" }); }.bind(this), 3000);
					//this.hideMessagebar();
					reject(err);
				});
		});
	}


	onBlurUserRole(e, item) {
		let updatedItems = this.state.items;
		let itemIdx = updatedItems.indexOf(item);
		updatedItems[itemIdx] = item;
        if (e.key) {
            updatedItems[itemIdx].processType = e.text;
        } else {
            if (e.target.id.match("txtADGroup"))
                updatedItems[itemIdx].adGroupName = e.target.value;
            else if (e.target.id.match("txtRole"))
                updatedItems[itemIdx].roleName = e.target.value;
            else if (e.target.id.match("txtProcessStep"))
                updatedItems[itemIdx].processStep = e.target.value;
            else if (e.target.id.match("txtChannel"))
                updatedItems[itemIdx].channel = e.target.value;
        }
		this.userRole = updatedItems;

		
		if (!updatedItems[itemIdx].id || !updatedItems[itemIdx].adGroupName || !updatedItems[itemIdx].roleName || !updatedItems[itemIdx].processStep || !updatedItems[itemIdx].channel || !updatedItems[itemIdx].processType) {
			return;
		}
				
		let currentuserRole = {};
		currentuserRole.id = item.id;
		currentuserRole.adGroupName = item.adGroupName;
		currentuserRole.processStep = item.processStep;
		currentuserRole.roleName = item.roleName;
		currentuserRole.channel = item.channel;
		currentuserRole.adGroupId = item.adGroupId;
		currentuserRole.processType = item.processType;

		let messagebarText;
		
		let FilteredItems = updatedItems.filter(function (k) {
			return k.processStep.toLowerCase() !== "new opportunity" && k.processStep.toLowerCase() !== "start process" && k.processStep.toLowerCase() !== "draft proposal" && k.processStep.toLowerCase() !== "none";
		});
		
        var processStepList = FilteredItems.map(item => item.processStep)
            .filter((value, index, self) => self.indexOf(value) === index);
		
		if (processStepList.length > 5) {
			alert("A maximum of 8 unique process Steps are currently supported.");
			this.getUserRoles();
			return;
		}
		this.setState({
			items: updatedItems
			//MessagebarText: messagebarText,
			//isUpdate: false,
			//isUpdateMsg: true
		});
		if (item.type === "new") {
			this.createNewItem(currentuserRole);
			messagebarText = "Role Mapping added successfully.";
		}
		else {
			this.updateItem(currentuserRole);
			messagebarText = "Role Mapping updated successfully.";
		}
		updatedItems[itemIdx].type = "old";
		this.setState({
			items: updatedItems,
			MessagebarText: messagebarText,
			isUpdate: false,
			isUpdateMsg: true
		});
		setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.success, MessagebarText: "" }); }.bind(this), 3000);
		this.getUserRoles();
		//console.log(updatedItems);


	}

	createNewItem(userRoleMapping) {
		return new Promise((resolve, reject) => {

			let requestUrl = 'api/RoleMapping';
			
			let options = {
				method: "POST",
				headers: {
					'Accept': 'application/json',
					'Content-Type': 'application/json',
					'authorization': 'Bearer ' + this.authHelper.getWebApiToken()
				},
				body: JSON.stringify(userRoleMapping)
			};

			fetch(requestUrl, options)
				.then(response => {
					console.log("Update Role Mapping response: " + response.status + " - " + response.statusText);
					if (response.status === 401) {
						// TODO: For v2 see how we pass to authHelper to force token refresh
					}
					return response;
				})
				.then(data => {

					resolve(data);
				})
				.catch(err => {
					//this.errorHandler(err, "");
					this.setState({
						MessagebarText: "Error occured. Please try again!",
						isUpdate: false,
						isUpdateMsg: true
					});
					setTimeout(function () { this.setState({ isUpdateMsg: false, MessageBarType: MessageBarType.error, MessagebarText: "" }); }.bind(this), 3000);
					//this.hideMessagebar();
					reject(err);
				});
		});
	}



	updateItem(userRoleMapping) {
		return new Promise((resolve, reject) => {

			let requestUrl = 'api/RoleMapping';
			
			let options = {
				method: "PATCH",
				headers: {
					'Accept': 'application/json',
					'Content-Type': 'application/json',
					'authorization': 'Bearer ' + this.authHelper.getWebApiToken()
				},
				body: JSON.stringify(userRoleMapping)
			};

			fetch(requestUrl, options)
				.then(response => {
					console.log("Update User Role Mapping response: " + response.status + " - " + response.statusText);
					if (response.status === 401) {
						// TODO: For v2 see how we pass to authHelper to force token refresh
					}
					return response;
				})
				.then(data => {
					resolve(data);
				})
				.catch(err => {
					//this.errorHandler(err, "");
					this.setState({
						updateStatus: true,
						MessagebarText: "Error while updating user Role Mapping, please try again."
					});
					this.hideMessagebar();
					reject(err);
				});
		});
	}


	userRoleList(columns, isCompactMode, items, selectionDetails) {
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



	render() {
		const { columns, isCompactMode, items, selectionDetails } = this.state;
		const userRoleList = this.userRoleList(columns, isCompactMode, items, selectionDetails);
		const itemCount = items.length;

		if (this.state.loading) {
			return (
				<div className='ms-BasicSpinnersExample ibox-content pt15 '>
					<Spinner size={SpinnerSize.large} label='loading...' ariaLive='assertive' />
				</div>
			);
		} else {
			return (

                <div className='ms-Grid bg-white ibox-content  '>
				

					<div className='ms-Grid-row'>
						<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                            <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 pt10'>
                                <Link href='' className='pull-left' onClick={() => this.onAddRow()} >+ Add New</Link>
                            </div>
							{itemCount === 0 ?
								<div>
									<br />
									<h5>No records to display. Please 'Add New' records.</h5>
								</div>

								:
								<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>

									{userRoleList}
								</div>
							}
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