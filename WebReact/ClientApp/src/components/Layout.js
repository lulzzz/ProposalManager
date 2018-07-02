/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { Col, Grid, Row } from 'react-bootstrap';
import { NavMenu } from './NavMenu';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import'../Style.css'

export class Layout extends Component {
    displayName = Layout.name

    constructor(props) {
        super(props);

        this.state = {
            userProfile: this.props.userProfile
        };

        initializeIcons();
    }

    render() {
        const userProfileData = this.state.userProfile;

        return (
            <Grid fluid>
                <Row>
                    <Col sm={2}>
                        <NavMenu userProfile={userProfileData} />
                    </Col>

                    <Col sm={10} className='mainpanel'>
                        {this.props.children}
                    </Col>
                </Row>
            </Grid>
        );
    }
}
