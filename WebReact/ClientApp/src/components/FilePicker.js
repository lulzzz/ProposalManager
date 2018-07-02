/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { DefaultButton, PrimaryButton, IconButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { TextField } from 'office-ui-fabric-react/lib/TextField';

export class FilePicker extends Component {
    displayName = FilePicker.name

    constructor(props) {
        super(props);

        let fileUri = "";
        if (this.props.fileUri) {
            fileUri = this.props.fileUri;
        }

        let showLabel = true;
        if (this.props.showLabel === false) {
            showLabel = false;
        }

        let showBrowse = true;
        if (this.props.showBrowse === false) {
            showBrowse = false;
        }

        this.state = {
            file: this.props.file,
            fileUri: fileUri,
            showLabel: showLabel,
            showBrowse: showBrowse
        };
    }

    onChangeFile(e) {
        let labelFileElem = document.getElementById('lblFile' + this.props.id);
        labelFileElem.textContent = e.target.files[0].name;

        //this.props.file = e.target.files[0];

        this.setState({
            file: e.target.files[0]
        });

        this.props.onChange(e.target.files[0]);
    }

    onClickFileAdd(e) {
        let fileElem = document.getElementById('selFile' + this.props.id);
        fileElem.click();
    }

    render() {
        let btnCaption = "Browse...";
        if (this.props.btnCaption) {
            btnCaption = this.props.btnCaption;
        }

        let showLink = false;
        if (this.state.fileUri && this.state.fileUri !== "") {
            showLink = true;
        }

        return (
            <div>
                <input type="file" id={'selFile' + this.props.id} name={'selFile' + this.props.id} style={{ display: 'none' }} onChange={(e) => this.onChangeFile(e)} />
                {
                    this.state.showLabel &&
                    <TextField
                        id={'lblFile' + this.props.id}
                        value={this.state.file.name}
                        disabled='true'
                        className='filepickerAlign'
                    />
                }

                {
                    showLink &&
                    <Link id={'lnkFile' + this.props.id} href={this.state.fileUri} target='_blank'>
                        &nbsp; <Icon iconName='View' /> &nbsp; View Document
                    </Link>
                }

                {
                    this.state.showBrowse &&
                    <DefaultButton className='pull-right greybutton is-disabled' onClick={e => this.onClickFileAdd(e)} disabled={this.props.disabled}>{btnCaption}</DefaultButton>
                }
            </div>
        );
    }
}