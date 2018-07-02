/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import * as ReactDOM from 'react-dom';
import { ChooseTeam } from './Opportunity/ChooseTeam';
import {
    Spinner,
    SpinnerSize
} from 'office-ui-fabric-react/lib/Spinner';

export class OpportunityChooseTeam extends Component {
    displayName = OpportunityChooseTeam.name

    constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;
        this.authHelper = window.authHelper;

        //const userProfile = this.props.userProfile;
        const oppId = this.props.opportunityId;

        this.state = {
            oppID: oppId
        };
    }

    
    componentWillMount() {

    }

    render() {
            return (
                <div>
                    <ChooseTeam oppID={this.state.oppID} />
                </div>
            );
    }
}