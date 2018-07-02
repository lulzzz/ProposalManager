/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

export class ComponntTemplate extends Component {
    displayName = ComponntTemplate.name

    constructor(props) {
        super(props);

        this.state = {
            showDialog: this.props.showDialog
        };
    }

    _showDialog = () => {
        this.setState({ showDialog: true });
    }

    _closeDialog = () => {
        this.setState({ showDialog: false });

    }

    render() {
        return (
            <div>
                <DefaultButton
                    description='Opens the Sample Dialog'
                    onClick={this._showDialog}
                    text='Open Dialog'
                />
                <label id='myLabelId' className='screenReaderOnly'>Session expired</label>
            </div>
        );
    }
}