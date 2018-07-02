/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

// Global imports
import React, { Component } from 'react';
import GraphSdkHelper from './helpers/GraphSdkHelper';
import AuthHelper from './helpers/AuthHelper';
import { AppBrowser } from './AppBrowser';
import { AppTeams } from './AppTeams';
//import { loadTheme } from 'office-ui-fabric-react/lib/Styling';


export default class App extends Component {
    displayName = App.name

    constructor(props) {
        super(props);

        if (window.authHelper) {
            this.authHelper = window.authHelper;
        } else {
            // Initilize the AuthService and save it in the window object.
            this.authHelper = new AuthHelper();
            window.authHelper = this.authHelper;
        }

        if (window.sdkHelper) {
            this.sdkHelper = window.sdkHelper;
        } else {
            // Initilize the AuthService and save it in the window object.
            this.sdkHelper = new GraphSdkHelper();
            window.sdkHelper = this.sdkHelper;
        }

        // Officec-fabrix-ui theme overrides
        //loadTheme({
        //    palette: {
        //        'themePrimary': 'red'
        //    }
        //});

        this.state = {
            inTeams: this.inTeams()
        };
    }

    componentWillMount() {
        let inTeams = this.inTeams();
        if (inTeams) {
            this.setState({
                inTeams: true
            });
        }
    }

    // This is a simple method to check if your webpage is running inside of MS Teams.
    // This just checks to make sure that you are or are not iframed.
    iframed = () => {
        try {
            return window.self !== window.top;
        } catch (err) {
            return true;
        }
    }

    // All routes for teams are under /tab
    inTeams = () => {
        console.log("APP_inTeams href: " + window.location.pathname);

        if (window.location.pathname.substring(0, 4) === "/tab") {
            return true;
        } else {
            return false;
        }
    }

    render() {
        console.log("App_render inTeams: " + this.state.inTeams + " iframed: " + this.iframed());
        
        if (this.state.inTeams) {
            return (
                <AppTeams />
            );
        } else {
            return (
                <AppBrowser />
            );
        }
    }
}
