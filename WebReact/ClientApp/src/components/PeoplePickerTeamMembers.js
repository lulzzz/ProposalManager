/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { NormalPeoplePicker } from 'office-ui-fabric-react/lib/Pickers';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

export class PeoplePickerTeamMembers extends Component {
    displayName = PeoplePickerTeamMembers.name

    constructor(props) {
        super(props);

        // Set the initial state for the picker data source.
        // The people list is populated in the _onFilterChanged function.
        this._peopleList = [];
        this._searchResults = [];

        // Helper that uses the JavaScript SDK to communicate with Microsoft Graph.
        this.sdkHelper = window.sdkHelper;

        this.authHelper = window.authHelper;

        this._showError = this._showError.bind(this);

        let filteredList = this.props.teamMembers;
        let isDisableTextBox = this.props.isDisableTextBox;

        this.state = {
            teamMembers: filteredList,
            defaultSelectedItems: [],
            isLoadingPeople: false,
            isLoadingPics: false,
            isDisableTextBox: isDisableTextBox
        };
    }

    componentWillMount() {

        if (this.props.defaultSelectedUsers && this.props.defaultSelectedUsers.length > 0 && this.props.defaultSelectedUsers[0].displayName.length > 0) {
            this.mapDefaultSelectedItems();
        }
    }

    fetchResponseHandler(response, referenceCall) {
        // console.log("PeoplePickerLoanOfficer fetchResponseHandler refcall: " + referenceCall + " response: " + response.status + " - " + response.statusText);
        if (response.status === 401) {
            // TODO: placeholder for future logic
        }
    }

    errorHandler(err, referenceCall) {
        console.log("PeoplePickerTeamMembers Ref: " + referenceCall + " error: " + JSON.stringify(err));
    }

    getUserProfilesSearch(searchText, callback) {
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
                    this.fetchResponseHandler(response, "getUserProfiles");
                    let err = "Error in fetch get users";
                    callback(err, []);
                }
            })
            .then(data => {
                let itemslist = [];

                if (data.ItemsList.length > 0) {
                    for (let i = 0; i < data.ItemsList.length; i++) {

                        let item = data.ItemsList[i];

                        let newItem = {};

                        newItem.id = item.id;
                        newItem.displayName = item.displayName;
                        newItem.mail = item.mail;
                        newItem.userPrincipalName = item.userPrincipalName;
                        newItem.userRoles = item.userRoles;

                        itemslist.push(newItem);
                    }
                }

                let filteredList = itemslist.filter(itm => itm.userRole === 1);

                this.setState({
                    teamMembers: filteredList,
                    isLoadingPeople: false
                });

                callback(null, filteredList ? filteredList : []);
            })
            .catch(err => {
                this._showError(err);
                callback(err, []);
            });

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
                    this.fetchResponseHandler(response, "getUserProfiles");
                    return [];
                }
            })
            .then(data => {
                let itemslist = [];

                if (data.ItemsList.length > 0) {
                    for (let i = 0; i < data.ItemsList.length; i++) {

                        let item = data.ItemsList[i];

                        let newItem = {};

                        newItem.id = item.id;
                        newItem.displayName = item.displayName;
                        newItem.mail = item.mail;
                        newItem.phoneNumber = item.phoneNumber;
                        newItem.UserPicture = item.UserPicture;
                        newItem.userPrincipalName = item.userPrincipalName;
                        newItem.userRole = item.userRole;

                        itemslist.push(newItem);
                    }
                }

                // filter to just loan officers
                let filteredList = itemslist.filter(itm => itm.userRole === 1);
                
                this.setState({
                    teamMembers: filteredList,
                    isLoadingPeople: false
                });
            })
            .catch(err => {
                this._showError(err);
            });
    }

    // Create the persona for the defaultSelectedItems using user ID
    mapDefaultSelectedItems() {
        this.setState({
            defaultSelectedItems: this._mapUsersToPersonas(this.props.defaultSelectedUsers, false)
        });
    }

    // Map user properties to persona properties.
    _mapUsersToPersonas(users, useMailProp) {
        return users.map((p) => {

            // The email property is returned differently from the /users and /people endpoints.
            let email = p.mail ? p.mail : p.userPrincipalName;
            
            let persona = {
                id: p.id,
                text: p.displayName,
                secondaryText: p.userPrincipalName,
                presence: PersonaPresence.none,
                imageInitials: p.displayName.substring(0, 2),
                initialsColor: Math.floor(Math.random() * 15) + 0,
                mail: email,
                userPrincipalName: p.userPrincipalName,
                userRoles: p.userRoles,
                status: 0
            };

            return persona;
        });
    }

    // Gets the profile photo for each user.
    _getPics(personas) {

        // Make suggestions available before retrieving profile pics.
        this.setState({
            isLoadingPics: false
        });

        // TODO: Retrieve pictures
        //this.sdkHelper.getProfilePics(personas, (err) => {
        //    this.setState({
        //        isLoadingPics: false
        //    });
        //});
    }

    // Remove currently selected people from the suggestions list.
    _listContainsPersona(persona, items) {
        if (!items || !items.length || items.length === 0) {
            return false;
        }
        return items.filter(item => item.primaryText === persona.primaryText).length > 0;
    }

    // Handler for when text is entered into the picker control.
    // Populate the people list.
    _onFilterChanged(filterText, items) {
        if (this._peopleList) {
            return filterText ? this._peopleList.concat(this._searchResults)
                .filter(item => item.primaryText.toLowerCase().indexOf(filterText.toLowerCase()) === 0)
                .filter(item => !this._listContainsPersona(item, items)) : [];
        }
        else {
            //TODO: Fetch more people from MT right now MT returns all so no need for this. For future expansion.
            //return new Promise((resolve, reject) => this.getUserProfiles((err, people) => {
            //    console.log("ONFILTER: " + JSON.stringify(people));
            //    if (!err) {
            //        this._peopleList = this._mapUsersToPersonas(people, false);
            //        this._getPics(this._peopleList);
            //        resolve(this._peopleList);
            //    }
            //    else { this._showError(err); }
            //})).then(value => value.concat(this._searchResults)
            //    .filter(item => item.primaryText.toLowerCase().indexOf(filterText.toLowerCase()) === 0)
            //    .filter(item => !this._listContainsPersona(item, items)));
        }
    }

    // Handler for when the Search button is clicked.
    // This method returns the first 20 matches as suggestions.
    _onGetMoreResults(searchText) {
        this.setState({
            isLoadingPeople: true,
            isLoadingPics: false
        });
        return new Promise((resolve) => {
            this.getUserProfilesSearch(searchText.toLowerCase(), (err, people) => {
                // console.log("ONMORERES: " + JSON.stringify(people));
                if (!err) {
                    this._searchResults = this._mapUsersToPersonas(people, true);
                    this.setState({
                        isLoadingPeople: false
                    });
                    this._getPics(this._searchResults);
                    resolve(this._searchResults);
                }
            });
        });
    }

    // Handler for when the picker gets focus
    onEmptyInputFocusHandler() {
        return new Promise((resolve) => {
            this._peopleList = this._mapUsersToPersonas(this.state.teamMembers, true);

            this._getPics(this._peopleList);
            resolve(this._peopleList);
        });
    }

    // Show the results of the `/me/people` query.
    // For sample purposes only.
    _showPeopleResults() {
        let message = 'Query loading. Please try again.';
        if (!this.state.isLoadingPeople) {
            const people = this._peopleList.map((p) => {
                return `\n${p.primaryText}`;
            });
            message = people.toString();
        }
        alert(message);
    }

    // Configure the error message.
    _showError(err) {
        this.setState({
            result: {
                type: MessageBarType.error,
                text: `Error ${err.statusCode}: ${err.code} - ${err.message}`
            }
        });
    }

    // Renders the people picker using the NormalPeoplePicker template.
    render() {
        //onChange={this._onSelectionChanged.bind(this)}
        return (
            <div>
                {
                    this.state.isLoadingPeople ?
                        <div>
                            <Spinner size={SpinnerSize.xSmall} label='Loading loan officers list ...' ariaLive='assertive' /><br />
                        </div>
                        :
                        <div />
                }
                <NormalPeoplePicker
                    onResolveSuggestions={this._onFilterChanged.bind(this)}
                    pickerSuggestionsProps={{
                        suggestionsHeaderText: 'Team Members',
                        noResultsFoundText: 'No results found',
                        searchForMoreText: 'Search',
                        loadingText: 'Loading pictures...',
                        isLoading: this.state.isLoadingPics
                    }}
                    getTextFromItem={(persona) => persona.primaryText}
                    onEmptyInputFocus={this.onEmptyInputFocusHandler.bind(this)}
                    onChange={this.props.onChange}
                    onGetMoreResults={this._onGetMoreResults.bind(this)}
                    className='ms-PeoplePicker normalPicker'
                    key='normal-people-picker'
                    itemLimit={this.props.itemLimit ? this.props.itemLimit : '1'}
                    defaultSelectedItems={this.state.defaultSelectedItems ? this.state.defaultSelectedItems : []}
                    disabled={this.state.isLoadingPeople || this.state.isDisableTextBox}
                />
                <br />
                {
                    this.state.result &&
                    <MessageBar messageBarType={this.state.result.type}>
                        {this.state.result.text}
                    </MessageBar>
                }
            </div>
        );
    }
}
