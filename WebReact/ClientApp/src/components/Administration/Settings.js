/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import { Pivot, PivotItem, PivotLinkFormat, PivotLinkSize } from 'office-ui-fabric-react/lib/Pivot';
import { Label } from 'office-ui-fabric-react/lib/Label';

import { Category } from '../Administration/Category';
import { Industry } from '../Administration/Industry';
import { Region } from '../Administration/Region';
import { UserRole } from '../Administration/UserRole';

import '../../Style.css';

export class Settings extends Component { 
    displayName = Settings.name

    constructor(props) {
        super(props);
    }

 

    render() {

        return (
            <div className='ms-Grid'>
                <div className='ms-Grid-row'>
                    <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg6 pageheading'>
                        <h3>Settings</h3>
                    </div>
                </div>
                <div className='ms-Grid-row ibox-content'>
                    <div className='ms-Grid-col ms-sm12 ms-md12 ms-lg12'>
                        <Pivot linkFormat={PivotLinkFormat.tabs} linkSize={PivotLinkSize.large}>
                            <PivotItem linkText="Category" className="TabBorder">
                                <Category/> 
                                </PivotItem>    
                            <PivotItem linkText="Industry" className="TabBorder">
                                <Industry/>
                                </PivotItem>
                            <PivotItem linkText="Region" className="TabBorder">
                                <Region/>
                            </PivotItem>
                            <PivotItem linkText="Role Mapping" className="TabBorder">
                                <UserRole/>
                            </PivotItem>
                        </Pivot>
                    </div>
                </div>
            </div>

        );


    }
}