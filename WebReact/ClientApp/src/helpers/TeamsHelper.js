/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

import Msal, { UserAgentApplication, Logger } from 'msal';
import { clientId, redirectUri, graphScopes, resourceUri, webApiScopes, clientSecret, authority } from './AppSettings';
import Promise from 'promise';
import * as microsoftTeams from '@microsoft/teams-js';

export class TeamsHelper extends Component {
    displayName = TeamsHelper.name

    constructor(props) {
        super(props);

        if (!window.authHelper) {
            this.authHelper = new AuthHelper();
            window.authHelper = this.authHelper;
        } else {
            this.authHelper = window.authHelper;
        }

        try {
            /* Initialize the Teams library before any other SDK calls.
             * Initialize throws if called more than once and hence is wrapped in a try-catch to perform a safe initialization.
             */
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
        }
        catch (err) {
            console.log(err);
        }
        finally {
            this.state = {
                inTeams: this.inTeams(),
                theme: ThemeStyle.Light,
                fontSize: 16
            };
        }
    }

    componentWillMount() {
        // If you are deploying your site as a MS Teams static or configurable tab, you should add ?theme={theme} to
        // your tabs URL in the manifest. That way you will get the current theme on start up (calling getContext on
        // the MS Teams SDK has a delay and may cause the default theme to flash before the real one is returned).
        this.updateTheme(this.getQueryVariable('theme'));

        this.setState({
            fontSize: this.pageFontSize()
        });

        // If you are not using the MS Teams web SDK, you can remove this entire if block, otherwise if you want theme
        // changes in the MS Teams client to propogate to the page, you should leave this here.
        if (this.inTeams()) {
            this.setState({
                inTeams: true
            });
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
        }
    }

    // Grabs the font size in pixels from the HTML element on your page.
    pageFontSize = () => {
        let sizeStr = window.getComputedStyle(document.getElementsByTagName('html')[0]).getPropertyValue('font-size');
        sizeStr = sizeStr.replace('px', '');
        let fontSize = parseInt(sizeStr, 10);
        if (!fontSize) {
            fontSize = 16;
        }
        return fontSize;
    }

    // This is a simple method to check if your webpage is running inside of MS Teams.
    // This just checks to make sure that you are or are not iframed.
    inTeams = () => {
        try {
            return window.self !== window.top;
        } catch (err) {
            console.log(err);
            return true;
        }
    }

    // Sets the correct theme type from the query string parameter.
    updateTheme = (themeStr) => {
        let theme;
        switch (themeStr) {
            case 'dark':
                theme = ThemeStyle.Dark;
                break;
            case 'contrast':
                theme = ThemeStyle.HighContrast;
                break;
            case 'default':
            default:
                theme = ThemeStyle.Light;
        }
        this.setState({ theme });
    }

    // Returns the value of a query variable.
    getQueryVariable = (variable) => {
        const query = window.location.search.substring(1);
        const vars = query.split('&');
        for (const varPairs of vars) {
            const pair = varPairs.split('=');
            if (decodeURIComponent(pair[0]) === variable) {
                return decodeURIComponent(pair[1]);
            }
        }
        return null;
    }
}