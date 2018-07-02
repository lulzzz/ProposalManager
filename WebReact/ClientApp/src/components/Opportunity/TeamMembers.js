/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Glyphicon } from 'react-bootstrap';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { LinkContainer } from 'react-router-bootstrap';
import {
    Persona,
    PersonaSize
} from 'office-ui-fabric-react/lib/Persona';
import { oppStatus } from '../../common';

export class TeamMembers extends Component {
    displayName = TeamMembers.name
	constructor(props) {
        super(props);

        this.sdkHelper = window.sdkHelper;

        this.state = {
            redirect: false,
            teamName: this.props.opportunityName,
            channelId: "",
			teamWebUrl: "",
			isAdmin: this.props.isAdmin,
			useRole: this.props.useRole
        };
	}

	componentWillMount() {
	/*
		this.sdkHelper.getTeamByName(this.state.teamName)
			.then(data => {

				this.sdkHelper.getTeamById(data.value[0].id)
					.then(teamdata => {
				
						this.setState({ teamWebUrl: teamdata.webUrl });
					});
			});
			*/
	}


    render() {
	
		let enableEditTeam;
		
		if (this.props.opportunityState !== 1 && this.props.userRole.toLowerCase() === "loanofficer" ) {
			enableEditTeam = true;
		}
		else {
			enableEditTeam = false;
		}
			
        return (
            <div className='ms-Grid'>
                {typeof this.props.memberslist === 'undefined' ? "" :
                    this.props.memberslist.map((member, index) =>
                        member.displayName !== "" ?
                            <div className='ms-Grid-row bg-grey p-5 mr5A' key={index}>
                                <div className=' ms-Grid-col ms-sm6 ms-md8 ms-lg12'>
                                    <Persona
                                        { ...{ imageUrl: member.UserPicture, imageInitials: member.displayName ? member.displayName.match(/\b(\w)/g) ? member.displayName.match(/\b(\w)/g).join('') : "" : "" } }
                                        size={PersonaSize.size40}
                                        text={member.displayName}
                                        secondaryText={member.assignedRole.displayName}
                                    />
                                    <span>Status: {oppStatus[member.status]}
                                        <p className="pull-right">
                                            <Link href={"mailto:" + member.userPrincipalName}> <Glyphicon glyph='envelope' /></Link>&nbsp;&nbsp;&nbsp;
                                           
                                        </p>
                                    </span>
                                </div>
                            </div>
                            : ""
                    )
                }

                {
                    <div className='ms-Grid-row p-10'> 
						<div className='ms-Grid ms-sm12 ms-md12 ms-lg12'>
							{
							enableEditTeam
								?
								<LinkContainer to={'/OpportunityChooseTeam?opportunityId=' + this.props.createTeamId} >
									<PrimaryButton className='ModifyButton'>Edit Team Collaboration </PrimaryButton>
								</LinkContainer>
								:
								<PrimaryButton className='ModifyButton' disabled>Edit Team Collaboration</PrimaryButton>

							}
							<br />
                            
						</div>
						<div className='ms-Grid ms-sm12 ms-md12 ms-lg12'>
							{this.props.opportunityState === 1
								?
								<Label>Team is being setup, please contact the admin. </Label>

								:
								""
							}
							</div>
                        </div>
                       // : ""
                }
            </div>
        );
    }
}
export default TeamMembers;