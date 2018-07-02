/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import React, { Component } from 'react';
import * as ReactDOM from 'react-dom';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { applicationId, redirectUri, graphScopes, resourceUri, webApiScopes } from '../helpers/config';

export class Opportunity extends Component {
    displayName = Opportunity.name

    constructor(props) {
        super(props);

        //TODO: Use this component as container for the 3 components of edit opportunity

        this.requestUrl = 'api/Data/CreateOpportunity?accessToken=' + window.authHelper.getIdToken();
        this.state = { results: [], loading: true };
        this.createOpportunity = this.createOpportunity.bind(this);
    }

    createOpportunity() {
        fetch(this.requestUrl, {
            method: "GET",
            headers: { 'authorization': 'Bearer ' + window.authHelper.getWebApiToken() }
        })
            .then(response => response.json())
            .then(data => {
                this.setState({ results: data, loading: false });
            });
    }

    static renderResults(results) {
        return (
            <div>
                {results}
            </div>
        );
    }

    render() {
        let contents = this.state.loading
            ? <p><em>Loading...</em></p>
            : Opportunity.renderResults(this.state.results);

        return (
            <div>
                <h1>Opportunity</h1>

                <p>This is a POC for creating the teams/channel for an opportunity.</p>

                <PrimaryButton onClick={this.createOpportunity}>Create Opportunity</PrimaryButton>

                <p>Results: {contents}</p>
            </div>
        );
    }
}