import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Fabric } from 'office-ui-fabric-react/lib';
import App from './components/App';
import './assets/styles/global.less';
import 'office-ui-fabric-react/dist/css/fabric.min.css';
import * as Msal from 'msal';
import { AppConfig } from './config/appconfig';

initializeIcons();

let isOfficeInitialized = false;

const title = AppConfig.title;

var clientApplication = new Msal.UserAgentApplication(AppConfig.applicationId, null, function(errorDesc, token, error, tokenType) {
    if(errorDesc)
    {
        console.log(errorDesc);
    }
    else
    {
        this.acquireTokenSilent([AppConfig.applicationId])
        .then(token => {
            (window as any).sessionStorage[AppConfig.accessTokenKey] = token;
            render(App);
        })
        .catch(err => console.log(err));
    }
});
(window as any).authorization = clientApplication;

const render = (Component) => {
    ReactDOM.render(
        <Fabric>
            <Component title={title} isOfficeInitialized={isOfficeInitialized} />
        </Fabric>
        ,document.getElementById('container')
    );
};

Office.initialize = () =>
{
    isOfficeInitialized = true;
    render(App);
}

/* Initial render showing a progress bar */
render(App);
