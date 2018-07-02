/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { setIconOptions } from 'office-ui-fabric-react/lib/Styling';
import { Link as LinkRoute } from 'react-router-dom';
import { FilePicker } from '../FilePicker';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { PeoplePickerTeamMembers } from '../PeoplePickerTeamMembers';


export class ChooseTeam extends Component {
	displayName = ChooseTeam.name
	constructor(props) {
		super(props);

		this.sdkHelper = window.sdkHelper;
		this.authHelper = window.authHelper;

        const oppID = this.props.oppID;

		// Suppress icon warnings.
		setIconOptions({
			disableWarnings: false
        });

		this.state = {
			isChecked: false,
			checked: false,
			OfficersList: [],
			teamcount: 0,
			Team: [],
			selectedRole: {},
			selectorFiles: [],
			selectedTeamMember: '',
			filterOfficersList: [],
			currentSelectedItems: [],
			peopleList: [],
			OppDetails: {},
			mostRecentlyUsed: [],
			allOfficersList: [],
			oppName: "",
			MessagebarText: "",
			MessagebarTextFinalizeTeam: "",
			MessageBarTypeFinalizeTeam: "",
			otherPeopleList: [],
			oppTeamMembers: [],
			loading: true,
			usersPickerLoading: true,
			oppID: oppID,
			proposalDocumentFileName: "",
			creditAnalystList: [],
			legalCounselList: [],
			seniorRiskOfficerList: [],
			selectedCreditAnalyst: [],
			selectedLegalCounsel: [],
            selectedSeniorRiskOfficer: [],
            UserRoleMapList: [],
            isDisableCreditAnalyst: false,
            isDisableLegalCounsel: false,
            isDisableSeniorRiskOfficer: false,
            isEnableFinalizeTeamButton: false
		};

		//this._ddlRolechangeState = this._ddlRolechangeState.bind(this);
		this.onFinalizeTeam = this.onFinalizeTeam.bind(this);
		this.handleFileUpload = this.handleFileUpload.bind(this);
		this.saveFile = this.saveFile.bind(this);
		this.selectSeniorRiskOfficer = this.selectSeniorRiskOfficer.bind(this);
		this.selectLegalCounsel = this.selectLegalCounsel.bind(this);
		this.selectCreditAnalyst = this.selectCreditAnalyst.bind(this);
	}

    componentWillMount() {
        this.getUserRoles();
        this.getOpportunity();
	}

    getOpportunity() {
        let oppDetails = {};
        let requestUrl = 'api/Opportunity/?id=' + this.state.oppID;

        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
        })
            .then(response => response.json())
            .then(data => {

                //get users for people picker 
                this.getUserProfiles();

                let oppSelTeam = [];
                if (data.teamMembers.length > 0) {
                    for (let m = 0; m < data.teamMembers.length; m++) {
                        //let member = {};
                        let item = data.teamMembers[m];
                        if (item.displayName.length > 0) {
                            let newItem = {};

                            newItem.id = item.id;
                            newItem.displayName = item.displayName;
                            newItem.mail = item.mail;
                            newItem.userPrincipalName = item.userPrincipalName;
                            newItem.userRoles = item.userRoles;
                            newItem.status = 0;
                            newItem.assignedRole = item.assignedRole;

                            oppSelTeam.push(newItem);
                        }
                    }
                }

                let creditAnalyst = [];
                let creditAnalystCheck = oppSelTeam.filter(x => x.assignedRole.displayName === "CreditAnalyst");
                if (creditAnalystCheck.length > 0) {
                    creditAnalyst.push(creditAnalystCheck[0]);
                }

                let legalCounsel = [];
                let legalCounselCheck = oppSelTeam.filter(x => x.assignedRole.displayName === "LegalCounsel");
                if (legalCounselCheck.length > 0) {
                    legalCounsel.push(legalCounselCheck[0]);
                }

                let seniorRiskOfficer = [];
                let seniorRiskOfficerCheck = oppSelTeam.filter(x => x.assignedRole.displayName === "SeniorRiskOfficer");
                if (seniorRiskOfficerCheck.length > 0) {
                    seniorRiskOfficer.push(seniorRiskOfficerCheck[0]);
                }

                let fileName = this.getDocumentName(data.proposalDocument.documentUri);

                this.setState({
                    oppData: data,
                    oppName: data.displayName,
                    oppTeamMembers: oppSelTeam,
                    oppID: data.id,
                    currentSelectedItems: oppSelTeam,
                    loading: false,
                    proposalDocumentFileName: fileName,
                    selectedCreditAnalyst: creditAnalyst,
                    selectedLegalCounsel: legalCounsel,
                    selectedSeniorRiskOfficer: seniorRiskOfficer

                });
            });
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
                    this.setState({ UserRoleMapList: userRoleList });
                }
                catch (err) {
                    return false;
                }

            });
    }

	getDocumentName(fileUri) {
		const vars = fileUri.split('&');
		for (const varPairs of vars) {
			const pair = varPairs.split('=');
			if (decodeURIComponent(pair[0]) === "file") {
				return decodeURIComponent(pair[1]);
			}
		}
	}

	getUserProfiles() {
		let requestUrl = 'api/UserProfile/';
		fetch(requestUrl, {
			method: "GET",
			headers: {
				'authorization': 'Bearer ' + this.authHelper.getWebApiToken()
			}
		})
			.then(response => {
				if (response.ok) {
					return response.json();
				} else {
					this.setState({ usersPickerLoading: false });
					this.fetchResponseHandler(response, "getUserProfiles");
					return [];
				}
			})
			.then(data => {
				let itemslist = [];
				let CreditAnalystList = [];
				let LegalCounselList = [];
				let SeniorRiskOfficerList = [];

                //Check "CreditAnalyst,LegaCounsel,SeniorRiskOfficer" roles exist in UserRoleMapList"
                let creditAnalystRoleArr = this.state.UserRoleMapList.filter(x => x.roleName.toLowerCase() === "creditanalyst");
                let isDisableCreditAnalyst = creditAnalystRoleArr.length > 0 ? false :  true;

                let legalCounselRoleArr = this.state.UserRoleMapList.filter(x => x.roleName.toLowerCase() === "legalcounsel");
                let isDisableLegalCounsel = legalCounselRoleArr.length > 0 ? false : true;

                let seniorRiskOfficerRoleArr = this.state.UserRoleMapList.filter(x => x.roleName.toLowerCase() === "seniorriskofficer");
                let isDisableSeniorRiskOfficer = seniorRiskOfficerRoleArr.length > 0 ? false : true;


				if (data.ItemsList.length > 0) {
					for (let i = 0; i < data.ItemsList.length; i++) {
						
						let item = data.ItemsList[i];

						let newItem = {};
							
                        newItem.id = item.id;
                        newItem.displayName = item.displayName;
						newItem.mail = item.mail;
						newItem.userPrincipalName = item.userPrincipalName;
						newItem.userRoles = item.userRoles;
						newItem.status = 0;
						// newItem.assignedRole = item.assigned; //RoleuserRoles.filter(x => (x.displayName === "CreditAnalyst") || (x.displayName === "LegalCounsel") || (x.displayName === "SeniorRiskOfficer") )[0];                   						
						let creditAnalyst = newItem.userRoles.filter(x => x.displayName === "CreditAnalyst");
						if (creditAnalyst.length > 0) {
							CreditAnalystList.push(newItem);
						}

						let legalCounsel = newItem.userRoles.filter(x => x.displayName === "LegalCounsel");
						if (legalCounsel.length > 0) {
							LegalCounselList.push(newItem);
						}

						let seniorRiskOfficer = newItem.userRoles.filter(x => x.displayName === "SeniorRiskOfficer");
						if (seniorRiskOfficer.length > 0) {
							SeniorRiskOfficerList.push(newItem);
						}

						itemslist.push(newItem);
					}
				}

				this.setState({
					allOfficersList: itemslist,
					usersPickerLoading: false,
					otherPeopleList: [],
					creditAnalystList: CreditAnalystList,
					legalCounselList: LegalCounselList,
                    seniorRiskOfficerList: SeniorRiskOfficerList,
                    isDisableCreditAnalyst: isDisableCreditAnalyst,
                    isDisableLegalCounsel: isDisableLegalCounsel,
                    isDisableSeniorRiskOfficer: isDisableSeniorRiskOfficer,
                    isDisableFinalizeTeamButton: isDisableCreditAnalyst && isDisableLegalCounsel && isDisableSeniorRiskOfficer ? true : false
				});
			})
			.catch(err => {
				console.log("Opportunities_getUserProfiles error: " + JSON.stringify(err));
			});
	}

	saveFile() {
		let files = this.state.selectorFiles;
		for (let i = 0; i < files.length; i++) {
			let fd = new FormData();
			fd.append('opportunity', "ProposalDocument");
			fd.append('file', files[0]);
			fd.append('opportunityName', this.state.oppName);
            fd.append('fileName', files[0].name);

			this.setState({
				IsfileUpload: true
			});

            let requestUrl = "api/document/UploadFile/" + encodeURIComponent(this.state.oppName) + "/ProposalTemplate";
            
			let options = {
				method: "PUT",
				headers: {
					'authorization': 'Bearer ' + this.authHelper.getWebApiToken()
				},
				body: fd
			};
			try {
				fetch(requestUrl, options)
					.then(response => {
						if (response.ok) {
							return response.json;
						} else {
							console.log('Error...: ');
						}
					}).then(data => {
						this.setState({ IsfileUpload: false, fileUploadMsg: true, MessagebarText: "Template uploaded successfully" });
						setTimeout(function () { this.setState({ fileUploadMsg: false, MessagebarText: "" }); }.bind(this), 3000);
					});
			}
			catch (err) {
				this.setState({
					IsfileUpload: false,
					fileUploadMsg: true,
					MessagebarText: "Error while uploading template. Please try again."
				});
				//alert("Error Uploading File");
				return false;
			}
		}
	}

	handleFileUpload(file) {
		this.setState({ selectorFiles: this.state.selectorFiles.concat([file]) });
	}

	onFinalizeTeam() {
		let oppID = this.state.oppID;
        let teamsSelected = this.state.currentSelectedItems;
        let oppDetails = {};

		this.setState({
			isFinalizeTeam: true
        });

		let data = this.state.oppData;
		data.teamMembers = teamsSelected;
		
		let fetchData = {
			method: 'PATCH',
			body: JSON.stringify(data),
			headers: {
				'Content-Type': 'application/json',
				'authorization': 'Bearer ' + this.authHelper.getWebApiToken()
			}
        };

		let requestUrl = 'api/opportunity';

        fetch(requestUrl, fetchData)
			.catch(error => console.error('Error:', error))
			.then(response => {
				this.setState({ isFinalizeTeam: false, finazlizeTeamMsg: true, MessagebarTextFinalizeTeam: "Finalize Team Complete", MessageBarTypeFinalizeTeam: MessageBarType.success });
				setTimeout(function () {
					this.setState({ finazlizeTeamMsg: false, MessagebarTextFinalizeTeam: "" });
				}.bind(this), 3000);
			});
	}

	selectCreditAnalyst(item) {
		let tempSelectedTeamMembers = this.state.currentSelectedItems;
		let finalTeam = [];
			
		for (let i = 0; i < tempSelectedTeamMembers.length; i++) {

			if (tempSelectedTeamMembers[i].assignedRole.displayName !== "CreditAnalyst") {

				finalTeam.push(tempSelectedTeamMembers[i]);
			}
		}
			if (item.length === 0) {
			this.setState({
				currentSelectedItems: finalTeam
			});
			return;
		}
		else {

			let newMember = {};
			newMember.id = item[0].id;
			newMember.displayName = item[0].text;
			newMember.mail = item[0].mail;
			newMember.userPrincipalName = item[0].userPrincipalName;
			newMember.userRoles = item[0].userRoles;
			newMember.status = 0;
			newMember.assignedRole = item[0].userRoles.filter(x => x.displayName === "CreditAnalyst")[0];

			finalTeam.push(newMember);

			this.setState({
				currentSelectedItems: finalTeam
			});
		}
	}

	selectLegalCounsel(item) {

		let tempSelectedTeamMembers = this.state.currentSelectedItems;
		let finalTeam = [];
		
		for (let i = 0; i < tempSelectedTeamMembers.length; i++) {
			
			if (tempSelectedTeamMembers[i].assignedRole.displayName !== "LegalCounsel") {
				finalTeam.push(tempSelectedTeamMembers[i]);
			}
		}
		if (item.length === 0) {
			this.setState({
				currentSelectedItems: finalTeam
			});
			return;
		}
		else {
			let newMember = {};
			newMember.id = item[0].id;
			newMember.displayName = item[0].text;
			newMember.mail = item[0].mail;
			newMember.userPrincipalName = item[0].userPrincipalName;
			newMember.userRoles = item[0].userRoles;
			newMember.status = 0;
			newMember.assignedRole = item[0].userRoles.filter(x => x.displayName === "LegalCounsel")[0];

			finalTeam.push(newMember);

			this.setState({
				currentSelectedItems: finalTeam
			});
		}
	}

	selectSeniorRiskOfficer(item) {
		let tempSelectedTeamMembers = this.state.currentSelectedItems;
		let finalTeam = [];
		
		for (let i = 0; i < tempSelectedTeamMembers.length; i++) {

			if (tempSelectedTeamMembers[i].assignedRole.displayName !== "SeniorRiskOfficer") {
				finalTeam.push(tempSelectedTeamMembers[i]);
			}
		}
		if (item.length === 0) {
			this.setState({
				currentSelectedItems: finalTeam
			});
			return;
		}
		else {

			let newMember = {};
			newMember.id = item[0].id;
			newMember.displayName = item[0].text;
			newMember.mail = item[0].mail;
			newMember.userPrincipalName = item[0].userPrincipalName;
			newMember.userRoles = item[0].userRoles;
			newMember.status = 0;
			newMember.assignedRole = item[0].userRoles.filter(x => x.displayName === "SeniorRiskOfficer")[0];

			finalTeam.push(newMember);

			this.setState({
				currentSelectedItems: finalTeam
			});
		}
	}

	render() {
		const { isChecked, selectedRole } = this.state;
		let oppTeamMembers = this.state.oppTeamMembers;
		let oppID = this.state.oppID;
		let itemFileUri = "";
        let uploadedFile = { name: this.state.proposalDocumentFileName };

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
						<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg8 '>
							<div className='ms-Grid-row'>
								<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6 pageheading'>
									<h3>Update Team</h3>
								</div>
								<div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg6'><br />
									<LinkRoute to={"/OpportunitySummary?opportunityId=" + oppID} className='pull-right'> Back to Opportunity </LinkRoute>
								</div>
							</div>
							<div className='ms-Grid-row'>
								
								<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg3 hide'>
									<span>Search</span>
                                    <SearchBox
                                        placeholder='Search'
                                        className='bg-white'
                                    />
								</div>
								<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6 '>
									<span></span>
								</div>
							</div>
							
							{
								this.state.usersPickerLoading
									?
									<div className='ms-Grid-row bg-white '>
										<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 TeamsBGnew pull-right pb15'>
											<div className='ms-BasicSpinnersExample ibox-content pt15 '>
												<Spinner size={SpinnerSize.large} label='loading...' ariaLive='assertive' />
											</div>
										</div>
									</div>
									:
									<div>

										<div className='ms-Grid-row bg-white'>
											<div className='ms-Grid-col ms-sm11 ms-md11 ms-lg11 light-grey '> 
												<h5>Credit Analyst</h5>

												<span className="p-b-10"> </span>
                                                <PeoplePickerTeamMembers teamMembers={this.state.creditAnalystList} itemLimit='1' defaultSelectedUsers={this.state.selectedCreditAnalyst} onChange={(e) => this.selectCreditAnalyst(e)} isDisableTextBox={this.state.isDisableCreditAnalyst} />

											</div>

											<div className='ms-Grid-col ms-sm11 ms-md11 ms-lg11 light-grey '> 
												<h5>Legal Counsel</h5>

												<span className="p-b-10">  </span>
                                                <PeoplePickerTeamMembers teamMembers={this.state.legalCounselList} itemLimit='1' defaultSelectedUsers={this.state.selectedLegalCounsel} onChange={(e) => this.selectLegalCounsel(e)} isDisableTextBox={this.state.isDisableLegalCounsel} />

											</div>

											<div className='ms-Grid-col ms-sm11 ms-md11 ms-lg11 light-grey '> 
												<h5>Senior Risk Officer</h5>

												<span className="p-b-10"> </span>
                                                <PeoplePickerTeamMembers teamMembers={this.state.seniorRiskOfficerList} itemLimit='1' defaultSelectedUsers={this.state.selectedSeniorRiskOfficer} onChange={(e) => this.selectSeniorRiskOfficer(e)} isDisableTextBox={this.state.isDisableSeniorRiskOfficer} />

											</div>
										</div>

										<div className='ms-Grid-row bg-white'>
											<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg10 TeamsBGnew pb15'>
												{
													this.state.isFinalizeTeam ?
														<Spinner size={SpinnerSize.small} label='Finalizing Team..' ariaLive='assertive' className="pull-right p-5" />
														: ""
												}
												{
													this.state.finazlizeTeamMsg ?
                                                        <MessageBar
                                                            messageBarType={this.state.MessageBarTypeFinalizeTeam}
                                                            isMultiline={false}
                                                        >
                                                            {this.state.MessagebarTextFinalizeTeam}
														</MessageBar>
														: ""
												}
											</div>
											<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg4 pull-right TeamsBGnew pb15'>

                                                <PrimaryButton onClick={this.onFinalizeTeam} className='pull-right' disabled={this.state.isFinalizeTeam || this.state.isDisableFinalizeTeamButton} >Finalize Team</PrimaryButton >

											</div>

										</div>
									</div>
							}
						</div>
						<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg3 bg-white p10 pr0 pull-right'>
							<div className='ms-Grid-row'>
								<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 pl0'>
									<h4 className='p15'> Selected Team</h4>
									{
										this.state.currentSelectedItems.map((member, index) =>
											member.displayName !== "" ?
                                                <div className='ms-Grid-col ms-sm6 ms-md4 ms-lg12 p15' key={index}>
                                                    <Persona
                                                        { ...{ imageUrl: member.UserPicture, imageInitials: '' } }
                                                        size={PersonaSize.size40}
                                                        primaryText={member.displayName}
                                                        secondaryText={member.assignedRole.adGroupName}
                                                    />

												</div>
												: ""

										)

									}
								</div>
							</div>
						</div>
					</div>
					<div className='ms-Grid-row'>
						<div className='ms-Grid-col ms-sm12 ms-md12 ms-lg8 mt20 '>
							<div className='ms-Grid-row'>
                                <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12 pageheading bg-white pb20'>
                                    <h4 className=" mb0 pt15">Update Template</h4>
									<div className='docs-TextFieldExample ms-Grid-col ms-sm12 ms-md12 ms-lg12 pt10 '>
										<div className='ms-Grid-col ms-sm12 ms-md6 ms-lg9 pl0 pull-left' >
                                            <FilePicker
                                                id='filePicker'
                                                fileUri={this.state.oppData.proposalDocument.documentUri}
                                                file={uploadedFile}
                                                showBrowse='true'
                                                showLabel='true'
                                                onChange={(e) => this.handleFileUpload(e)}
                                                btnCaption={this.state.oppData.proposalDocument.documentUri ? "Change File" : ""}
                                            />
										</div>
										<div className='ms-Grid-col ms-sm12 ms-md6 ms-lg3 '>
											{
												this.state.IsfileUpload ?
													<Spinner size={SpinnerSize.small} ariaLive='assertive' className="pull-right p-5" />
													: ""
											}


											<PrimaryButton className='pull-right' onClick={this.saveFile} disabled={this.state.IsfileUpload}>Save</PrimaryButton >
											{
												this.state.fileUploadMsg ?
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

							</div>
						</div>
					</div>
				</div>

			);
		}
	}

}
