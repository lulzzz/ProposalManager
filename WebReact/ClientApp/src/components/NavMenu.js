/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

import React, { Component } from 'react';
import { Link } from 'react-router-dom';
import { Glyphicon, Nav, Navbar, NavItem } from 'react-bootstrap';
import { LinkContainer } from 'react-router-bootstrap';
import '../Style.css';
import { getQueryVariable } from '../common';

export class NavMenu extends Component {
	displayName = NavMenu.name

	constructor(props) {
		super(props);

		this.state = {
			userProfile: this.props.userProfile
		};
	}

	componentDidMount() {

	}


	render() {
		let menuRoot = true;
		if (window.location.pathname === "/" || window.location.pathname === "/Notifications") {
			menuRoot = true;
		} else if (window.location.pathname === "/OpportunitySummary" || window.location.pathname === "/OpportunityNotes" || window.location.pathname === "/OpportunityStatus" || window.location.pathname === "/OpportunityChooseTeam") {
			menuRoot = false;
		}

		const oppId = getQueryVariable('opportunityId');

		let isAdmin = false;
		if (this.state.userProfile.roles.filter(x => x.displayName === "Administrator").length > 0) {
			console.log("NavMenu Render isAdmin = true");
			isAdmin = true;
		}

		return (
			<Navbar inverse fixedTop fluid collapseOnSelect>
				<Navbar.Header>
					<Navbar.Brand>
						<Link to={'/'}>Proposal Manager</Link>
					</Navbar.Brand>
					<Navbar.Toggle />
				</Navbar.Header>
				<Navbar.Collapse>
					{
						menuRoot ?
							<Nav >
								<LinkContainer to={'/'} exact>
									<NavItem eventKey={1} >
										<i className="ms-Icon ms-Icon--HomeSolid pr10" aria-hidden="true"></i> Dashboard
                                    </NavItem>
								</LinkContainer>
								
								{
									isAdmin &&
									<Nav>
										<LinkContainer to={'/Administration'} exact >
											<NavItem eventKey={3}>
												<i className="ms-Icon ms-Icon--Admin pr10" aria-hidden="true"></i> Administration
                                        </NavItem>
										</LinkContainer>
										<LinkContainer to={'/Settings'} exact >
											<NavItem eventKey={4}>
												<i className="ms-Icon ms-Icon--Settings pr10" aria-hidden="true"></i>Settings
                                        </NavItem>
										</LinkContainer>
									</Nav>
								}
							</Nav>
							:
							<Nav >
								<LinkContainer to={'/OpportunitySummary?opportunityId=' + oppId} >
									<NavItem eventKey={1}>
										<Glyphicon glyph='list' /> Summary
                                    </NavItem>
								</LinkContainer>
								<LinkContainer to={'/OpportunityNotes?opportunityId=' + oppId} >
									<NavItem eventKey={2}>
										<Glyphicon glyph='edit' /> Notes
                                    </NavItem>
								</LinkContainer>
								<LinkContainer to={'/OpportunityStatus?opportunityId=' + oppId} >
									<NavItem eventKey={3}>
										<Glyphicon glyph='check' /> Status
                                    </NavItem>
								</LinkContainer>
							</Nav>
					}
				</Navbar.Collapse>
			</Navbar>
		);
	}
}
