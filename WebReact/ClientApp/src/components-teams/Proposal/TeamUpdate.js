/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';
import { Glyphicon } from 'react-bootstrap';
import {
	Persona,
	PersonaSize
} from 'office-ui-fabric-react/lib/Persona';
import { TeamsComponentContext, Anchor } from 'msteams-ui-components-react';
import {  redirectUri } from '../../helpers/AppSettings';
import { appUri, clientId} from '../../helpers/AppSettings';


export class TeamUpdate extends Component {
    displayName = TeamUpdate.name
	constructor(props) {
		super(props);

		this.sdkHelper = window.sdkHelper;
		let teamName = this.props.OppName;
        let channelId = this.props.channelId;
        try {
			/* Initialize the Teams library before any other SDK calls.
			 * Initialize throws if called more than once and hence is wrapped in a try-catch to perform a safe initialization.
			 */
            microsoftTeams.initialize();
        }
        catch (err) {
            console.log(err);
        }
        finally {
            this.state = {
				teamContext: {},
				channelId: channelId,
				teamName: teamName
            };
        }
	}

    componentWillMount() {
        // Get the teams context
		this.getTeamsContext();
	
    }


    getTeamsContext() {
        microsoftTeams.getContext(context => {
            let tc = {
                group: context.groupId,
                channel: context.channelName,
                team: context.teamName,
                entityId: context.entityId
            };

            this.setState({
                teamContext: tc
            });
        });
    }


	render() {
	
        let tablinkContext = {
            canvasUrl: appUri + "/rootTab",
            channelId: this.state.channelId
        };
		let entityId = "Root";
		let tablinkContextString = JSON.stringify(tablinkContext);
		
		let convDeepLink = "/l/entity/" + clientId + "/" + entityId + "?webUrl=" + redirectUri + "&label=Conversations" + "&context=" + tablinkContextString;

        let role = "";

        role = this.props.memberslist.assignedRole.adGroupName;
        return (
            <div className='ms-Grid'>
                <TeamsComponentContext>
				<div className='ms-Grid-row bg-grey p-5 mr5A' key={this.props.memberslist.id}>
                    <div className=' ms-Grid-col ms-sm12 ms-md8 ms-lg9'>
                            {this.props.memberslist.displayName ? 
                                <div>
                                    <Persona
                                        { ...{ imageUrl: "", imageInitials: "" } }
                                        size={PersonaSize.size40}
                                        text={this.props.memberslist.displayName}
                                        secondaryText={role}
                                    />
                                </div>
                                :
                                <div>
                                    <Persona
                                        {...{
                                            imageUrl: "data:image/svg+xml;base64,PD94bWwgdmVyc2lvbj0iMS4wIiBlbmNvZGluZz0iVVRGLTgiPz4NCjwhRE9DVFlQRSBzdmcgIFBVQkxJQyAnLS8vVzNDLy9EVEQgU1ZHIDEuMS8vRU4nICAnaHR0cDovL3d3dy53My5vcmcvR3JhcGhpY3MvU1ZHLzEuMS9EVEQvc3ZnMTEuZHRkJz4NCjxzdmcgd2lkdGg9IjQwMXB4IiBoZWlnaHQ9IjQwMXB4IiBlbmFibGUtYmFja2dyb3VuZD0ibmV3IDMxMi44MDkgMCA0MDEgNDAxIiB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjMxMi44MDkgMCA0MDEgNDAxIiB4bWw6c3BhY2U9InByZXNlcnZlIiB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciPg0KPGcgdHJhbnNmb3JtPSJtYXRyaXgoMS4yMjMgMCAwIDEuMjIzIC00NjcuNSAtODQzLjQ0KSI+DQoJPHJlY3QgeD0iNjAxLjQ1IiB5PSI2NTMuMDciIHdpZHRoPSI0MDEiIGhlaWdodD0iNDAxIiBmaWxsPSIjRTRFNkU3Ii8+DQoJPHBhdGggZD0ibTgwMi4zOCA5MDguMDhjLTg0LjUxNSAwLTE1My41MiA0OC4xODUtMTU3LjM4IDEwOC42MmgzMTQuNzljLTMuODctNjAuNDQtNzIuOS0xMDguNjItMTU3LjQxLTEwOC42MnoiIGZpbGw9IiNBRUI0QjciLz4NCgk8cGF0aCBkPSJtODgxLjM3IDgxOC44NmMwIDQ2Ljc0Ni0zNS4xMDYgODQuNjQxLTc4LjQxIDg0LjY0MXMtNzguNDEtMzcuODk1LTc4LjQxLTg0LjY0MSAzNS4xMDYtODQuNjQxIDc4LjQxLTg0LjY0MWM0My4zMSAwIDc4LjQxIDM3LjkgNzguNDEgODQuNjR6IiBmaWxsPSIjQUVCNEI3Ii8+DQo8L2c+DQo8L3N2Zz4NCg==",
                                            imageInitials: ""
                                        } }
                                        size={PersonaSize.size40}
                                        text="User Not Selected"
                                        secondaryText={role}
                                    />
                                </div>
                            }
					</div>
					<div className=' ms-Grid-col ms-sm12 ms-md8 ms-lg3'>
                        <span>
                            {
                                this.props.memberslist.displayName ? 
                                    <Anchor href={"mailto:" + this.props.memberslist.mail}> <Glyphicon glyph='envelope' /></Anchor>
                                    :
                                    null
                            }
                            {
                                this.props.memberslist.displayName ?
                                    this.props.isMobile ?
                                        ""
                                        :
										<Anchor href={"https://teams.microsoft.com" + convDeepLink} target="main" className={this.props.opportunityState === 1 ? "noPointer-event hide" : ""}> <Glyphicon glyph='comment' /></Anchor>
                                    :
                                    null
                            }
						</span>
					</div>
					<div className='ms-Grid-row'>
						<div className='ms-Grid ms-sm12 ms-md8 ms-lg12'>
							&nbsp;
                        </div>
					</div>
                    </div>
                </TeamsComponentContext>
			</div>
		);
	}

	
}