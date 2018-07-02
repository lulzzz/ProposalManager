/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import * as ReactDOM from 'react-dom';
import { Label } from 'office-ui-fabric-react/lib/Label';
import {
    Persona,
    PersonaSize,
    PersonaPresence
} from 'office-ui-fabric-react/lib/Persona';
import '../../Style.css';
import { oppStatus, oppStatusText } from '../../common';
import { TooltipHost } from 'office-ui-fabric-react/lib/Tooltip';


export class GroupEmployeeStatusCard extends Component {
    displayName = GroupEmployeeStatusCard.name

    constructor(props) {
        super(props);

    }


    getUsersDetails(users, role) {
        return (
            users.map((officer,ind) =>
                <div className="p-5" key={ind}>
                    <div className='ms-Grid-row'>
                        <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg4'>
                            &nbsp;&nbsp;
                        </div>
                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg8'>
                            &nbsp;&nbsp;
                        </div>
                    </div>
                    <Persona
                        { ...{ imageUrl: '', imageInitials: "" } }
                        size={PersonaSize.size40}
                        text={officer.displayName}
                        secondaryText={role}
                    />
                </div>
            )
        );
    }
    render() {
        let _menuButtonElement;

        let role = this.props.role;

        let members = this.props.members;

        let userContent = this.getUsersDetails(members, role);

        let status = "";
        let statusClassName = "";
        if (this.props.isDispOppStatus) {
            status = oppStatusText[this.props.status];
            statusClassName = this.props.status === 0 ? "status" + this.props.status : "status" + (this.props.status - 1);
        } else {
            status = oppStatus[this.props.status];
            statusClassName = "status" + this.props.status;
        }
        // Team view not display the status
        let isDispStatusClass = this.props.isTeam === true ? "hide" : "";

        return (
            <div className='ms-Grid'>
                <div className=' ms-Grid-row bg-grey p-5 mr5A'  >
                <div>
                        <TooltipHost
                            tooltipProps={{
                                onRenderContent: () => {
                                    return (
                                        userContent
                                        //content={userContent}
                                    );
                                }
                            }}
                            id="groupCardID" calloutProps={{ gapSpace: 0 }}
                        >
                            <div aria-describedby="groupCardID">
                                <div className={'ms-Grid-row ' + isDispStatusClass}>
                                    <div className='ms-Grid-col ms-sm6 ms-md8 ms-lg4'>
                                        <Label>Status</Label>

                                    </div>
                                    <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg8'>
                                        <Label><span className={statusClassName}> {status} </span></Label>
                                    </div>
                                </div>
                                <div className='ms-PersonaExample'>
                                    <div className='ms-Grid-row'>
                                        <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12'>
                                            <Persona
                                                { ...{
                                                    imageUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADoAAAAuCAIAAAD2np3yAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAK3SURBVGhD7Zctd7MwFIDfH1YThaqaQaFQNShUVQ0qCo+aQmGiplAxUSgUBoVC8QOWj0tXOkKSLuycnZfHkaTJ03Bzb/g3/SkO3T05dPfk0N2T/0R37Bpa1zVl7QAtzox9y+QcTTdCk5EXdMe2TIPTFyjCdQ99tvQ1jhD8XhCkZWvj7Kzbk0Stg85RFIXgHebMeoemkeWh+lkQ8jnOMF9CzH/aVbcp3sTccdGAXl9nsiUltlExkFT8AF1nvbF9T0TLW9GoBj2Oun0lJ87qh71sy0i0YQbPJhgWw6OyhWfOWGeiLalM++uo21UXMe+qLoVnE1Sre6k6aNDxYjCEeI7V14MhmeX2Cwbbo9Y3pMyzlI9Is7wkzeIl/+JR62lxecxiChTdqjkRPecoxUO2G9vqtjIiuBTU6Ouk25E54QYh3zUBvsbQhKKcDV0Fe38K4it+HsFf/8ByUEXneUSWwAbz9Eu2o9dBdw6D8PY058CKWPYgpJYNs6e6wQNcvX4YgeKCLUO9Izc1Yjs72OsyLFe6H5AFI8XywHE0MXiPeX6kMF0EOsDfjOhF2WqvwlqX5WI1/VxdGYvFNo43JJVTXGpe+EgzuUauz+C2ulAfNl6Vyv5b6Vcl3I16Yl7E6ahtAmsZdc2lawN/ulCZjMGwqIiu+NOdho+r0EHXj7X6xnvlUVvvtcWjrtg/lYz0iSw019lNvOrKZATpSlMmYNyreNYVCU2euBUSXQJzwKvuvTQJ5PVFMF+DON8Koiv+dO9xgGJMmuV5GhqCVZ3+YUR40r3fCUNMdSd/oBjGuHzYLfGjCxVYk8O+mLOZtg6b8KILBcDmgwI+JSw+HFbxoetUrozFbxMPuvxqKARs7wLzPcbiU+c7HnQHVopyUD5duHU4Dl/i56j9Gofunhy6e3Lo7smhuyd/SneaPgFYTrJ09Dnc1AAAAABJRU5ErkJggg==",
                                                    imageInitials: ""
                                                } }
                                                size={PersonaSize.size40}
                                                text='Group Members'
                                                secondaryText={role}
                                            />
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </TooltipHost>
                </div>

                
            </div>
            </div>
            );
    }
}