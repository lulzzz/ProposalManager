import * as React from 'react';
import { SignIn } from './SignIn';
import { Requirements } from './Requirements';

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface AppState {
    user: string
}

export default class App extends React.Component<AppProps, AppState> {
    constructor(props, context) {
        super(props, context);
    }

    checkRequirements = () =>
    {
        let isDocumentValid = (Office.context.document.url && Office.context.document.url.toUpperCase().indexOf("HTTPS") > -1);
        let isVersionValid = Office.context.requirements.isSetSupported("WordApi", 1.2) === true;

        return {
            documentValid: isDocumentValid,
            versionValid: isVersionValid
        };
    }

    render() {
        const { isOfficeInitialized } = this.props;
 
        if (isOfficeInitialized) {
            let reqs = this.checkRequirements();

            if(reqs.documentValid)
            {
                return (
                    <SignIn
                        title={this.props.title}
                        logo='dist/assets/logo-filled.png'
                        message='You are not logged in. Please sign in.'
                    />
                );
            }
            else
            {
                return (
                    <Requirements invalidVersion={!reqs.versionValid} invalidDocument={!reqs.documentValid} />
                );
            }
        }

        return (
            <div>
                <section className='ms-welcome__progress ms-u-fadeIn500'>
                    <h1 className='ms-font-xl'>Initializing Add in</h1>
                </section>
            </div>
        );
    }
}
