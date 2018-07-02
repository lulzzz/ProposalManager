/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import * as ReactDOM from 'react-dom';
import { Glyphicon } from 'react-bootstrap';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Persona, PersonaSize, PersonaPresence } from 'office-ui-fabric-react/lib/Persona';
import '../../Style.css';

export class NotificationDetails extends Component {
    displayName = NotificationDetails.name

    constructor(props) {
        super(props);

    }

    render() {
        const notification = this.props.notification;

        return (
            <div className='ms-Grid'>
                <div className='ms-Grid-row br-gray p-5'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <h2 className='pageheading'>Notification</h2>
                    </div>
                    
                </div>
                <div className='ms-Grid-row br-gray p-5'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <Link onClick={this.props.onClickBack} className='pull-right'>
                            <span>Back to Notifications </span>
                        </Link>
                    </div>
                </div>
                <div className='ms-Grid-row ibox-content'>
                    <div className=' ms-Grid-col ms-sm8 ms-md8 ms-lg8'>
                        <div className='ms-PersonaExample'>
                            <Persona
                                {...{ imageUrl: '', imageInitials: notification.sentFrom.match(/\b(\w)/g).join('') } }
                                size={PersonaSize.size40}
                                presence={PersonaPresence.none}
                                primaryText={notification.sentFrom}
                                secondaryText={notification.subject}
                            />
                        </div>
                    </div>
                    <div className=' ms-Grid-col ms-sm4 ms-md4 ms-lg4'>
                        <Label className='pull-right'>{new Date(notification.sentDate).toLocaleDateString()} </Label>
                    </div>
                </div>
                <div className='ms-Grid-row ibox-content'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <div dangerouslySetInnerHTML={{ __html: notification.message }} /> <br /><br /><br /><br />        
                    </div>
                </div>
                <div className='ms-Grid-row hide'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <TextField label='' multiline rows={6} />
                    </div>
                </div>
                <div className='ms-Grid-row'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12'>&nbsp;</div>
                </div>

                <div className='ms-Grid-row hide'>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <Label>Reply</Label>
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <Glyphicon glyph='file' className='pull-left' />
                    </div>
                    <div className=' ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <PrimaryButton className='backbutton pull-right'><Glyphicon glyph='play' /></PrimaryButton>
                    </div>
                </div>
            </div>
        );
    }
}