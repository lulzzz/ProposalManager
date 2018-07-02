/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { FocusZone, FocusZoneDirection } from 'office-ui-fabric-react/lib/FocusZone';
import { List } from 'office-ui-fabric-react/lib/List';
import { Persona, PersonaSize, PersonaPresence} from 'office-ui-fabric-react/lib/Persona';
import { NotificationDetails } from './Notification/NotificationDetails';
import '../Style.css'


export class Notifications extends Component {
    displayName = Notifications.name

    constructor(props) {
        super(props);

        this.authHelper = window.authHelper;

        const userProfile = this.props.userProfile;

        this.state = {
            loading: true,
            view: 'list',
            userProfile: userProfile,
            itemsList: [],
            itemsListOriginal: [],
            notificationDetails: {}
        };

        this.onClickBack = this.onClickBack.bind(this);
        this.onClickHandler = this.onClickHandler.bind(this);

        this.fetchNotifications();
    }

    errorHandler(err, referenceCall) {
        console.log("Notifications Ref: " + referenceCall + " error: " + JSON.stringify(err));
    }

    fetchResponseHandler(response, referenceCall) {
        console.log("Notifications fetchResponseHandler refcall: " + referenceCall + " response: " + response.status + " - " + response.statusText);
        if (response.status === 401) {
            this.setState({
                refreshToken: true
            });
        }
    }

    fetchNotifications() {
        let requestUrl = 'api/Notification/?page=1';
        fetch(requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + this.authHelper.getWebApiToken() }
        })
            .then(response => {
                if (response.ok) {
                    return response.json();
                } else {
                    this.fetchResponseHandler(response, "Notifications_fetchNotifications");
                    this.errorHandler(response, "Notifications_fetchNotifications");
                    return false;
                }
            })
            .then(data => {
                let notificationslist = [];
                if (data) {
                    for (let i = 0; i < data.length; i++) {

                        let item = data[i];
                        let newItem = {};

                        newItem.id = item.id;
                        newItem.subject = item.title;
                        newItem.sentTo = item.sentTo;
                        newItem.sentFrom = item.sentFrom;
                        newItem.date = new Date(item.sentDate).toLocaleDateString();
                        newItem.isRead = item.isRead;
                        newItem.message = item.message;

                        newItem.onClick = () => this.onClickHandler(item);

                        notificationslist.push(newItem);
                    }
                }

                this.setState({
                    loading: false,
                    items: notificationslist,
                    itemsOriginal: notificationslist
                });
            });
    }

    updateNotification(notification) {

        if (notification.id !== null) {
            // API Update notification IsRead flag
            let newNotification = {
                id: notification.id,
                title: notification.subject,
                sentTo: notification.sentTo,
                sentFrom: notification.sentFrom,
                sentDate: notification.date,
                isRead: "true",
                message: notification.message
            };

            let requestUrl = 'api/Notification';
            let options = {
                method: "PATCH",
                headers: {
                    'Accept': 'application/json',
                    'Content-Type': 'application/json',
                    'authorization': 'Bearer    ' + this.authHelper.getWebApiToken()
                },
                body: JSON.stringify(newNotification)
            };

            fetch(requestUrl, options)
                .then(response => {
                    //window.location.reload();
                    console.log('Success: ', response);
                })
                .catch(err => {
                    this.errorHandler(err, "Notifications_updateNotification_fetch error");
                });
        }
    }

    _onRenderCell(item, index) {
        const renderPersonaDetails = true;
        let isReadClass = "";
        if (!item.isRead) {
            isReadClass = "f-bold";
        }

        const onClickHandler = item.onClick;

        return (
            <div className='ms-List-itemCell' data-is-focusable='true'>
                <div className='ms-Grid ibox-content'>
                    <div className='ms-Grid-row' key={item.id}>
                        <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg1'>
                            <div className='ms-PersonaExample'>
                                <Persona
                                    {...{ imageUrl: '', imageInitials: item.sentFrom.match(/\b(\w)/g).join('') }}
                                    size={PersonaSize.size40}
                                    presence={PersonaPresence.none}
                                    hidePersonaDetails={!renderPersonaDetails}
                                />

                            </div>
                        </div>
                        <div className={' ms-Grid-col ms-sm12 ms-md12 ms-lg9 pull-left pl0 ' + isReadClass}>
                            <Link href='' onClick={onClickHandler} msgId={item.id}>
                                <Label><h5>{item.sentFrom} </h5>
                                    <span>{item.subject}</span>
                                </Label>
                            </Link>

                        </div>
                        <div className={' ms-Grid-col ms-sm12 ms-md12 ms-lg2 ' + isReadClass}>
                            <span className='pull-right'>{new Date(item.date).toLocaleDateString()} </span>
                        </div>
                    </div>
                </div>
            </div>
        );
    }

    notificationsList(itemsList, itemsListOriginal) {
        const lenght = typeof itemsList !== 'undefined' ? itemsList.length : 0;
        const lenghtOriginal = typeof itemsListOriginal !== 'undefined' ? itemsListOriginal.length : 0;
        const originalItems = itemsListOriginal;
        const items = itemsList;
        const resultCountText = lenght === lenghtOriginal ? '' : ` (${items.length} of ${originalItems.length} shown)`;

        console.log("notificationsList lenght = " + lenght + " resultcount: " + resultCountText);

        //<NotificationList notification={notificationItem} viewMessageBind={() => this.fnViewNotification(notificationItem)} value={notificationItem.id} />

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

    onClickHandler(item) {
        this.updateNotification(item);

        this.setState({
            notificationDetails: item,
            view: 'viewmessage'
        });
    }

    onClickBack() {
        this.setState({
            view: 'list'
        });
    }


    render() {
        const isLoading = this.state.loading;

        const itemsOriginal = this.state.itemsOriginal;
        const items = this.state.items;
        const lenghtOriginal = typeof itemsOriginal !== 'undefined' ? itemsOriginal.length : 0;
        const listHasItems = lenghtOriginal > 0 ? true : false;

        const notificationsListComponent = this.notificationsList(items, itemsOriginal);

        const notification = this.state.notificationDetails;

        if (this.state.view === "list") {
            return (
                <div className='ms-Grid'>
                    <div className='ms-Grid-row br-gray p-5'>
                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg6 pageheading'>
                            <h3>Notifications</h3>
                        </div>
                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg3'>
                        </div>
                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg3 hide'><br />
                            <SearchBox placeholder='Search' />
                        </div>
                    </div>
                    {
                        isLoading ?
                            <div className='ms-BasicSpinnersExample ibox-content pt15 '>
                                <br /><br />
                                <Spinner size={SpinnerSize.large} label='loading...' ariaLive='assertive' />
                            </div>
                            :
                            listHasItems ?
                                <div>
                                    {notificationsListComponent}
                                </div>
                                :
                                <div className='ms-BasicSpinnersExample ibox-content pt15 '>
                                    <p><em>No notifications were found</em></p>
                                </div>
                    }
                </div>
            );
        } else if (this.state.view === "viewmessage") {
            return (
                <div>
                    <NotificationDetails notification={notification} onClickBack={() => this.onClickBack()} />
                </div>
                );
        }
        return null;
    }
}