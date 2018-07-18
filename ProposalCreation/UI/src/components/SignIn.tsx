import * as React from 'react';
import { DefaultButton } from 'office-ui-fabric-react';
import {Main} from './Main';
import * as microsoftTeams from '@microsoft/teams-js';
import { AppConfig } from '../config/appconfig';
import { LocalizationService } from '../services/LocalizationService';

export interface IProgressProps {
    logo: string;
    message: string;
    title: string;
    localizationService: LocalizationService;
}

export interface IProgressState
{
    token: string;
    loginDisabled: boolean;
    error: string;
}

export class SignIn extends React.Component<IProgressProps, IProgressState> 
{
    constructor(props)
    {
        super(props);
        this.login = this.login.bind(this);
        this.state = {token : '', loginDisabled: false, error: ''};
    }

    private isInIframe()
    {
        try 
        {
            return window.self !== window.top;
        }
        catch
        {
            return true;
        }
    }

    componentDidMount()
    {
        let token = (window as any).sessionStorage[AppConfig.accessTokenKey];
        if(token)
        {
            this.setState({token: token, error: '', loginDisabled: false});
        }
    }

    private getTokenFromOfficeContext()
    {
        (window as any).authorization.loginRedirect([AppConfig.applicationId]);
    }

    private getTokenFromTeams()
    {
        microsoftTeams.initialize();

        microsoftTeams.authentication.authenticate({
            url: '/auth',
            width: 600,
            height: 535,
            successCallback: (result: any) => {
                // set state with token
                this.setState({token: result.idToken, error: '', loginDisabled: false});
            },
            failureCallback: (err) => {
                console.log(err);
                // set state with error
                this.setState({token: '', error: err, loginDisabled: false});
            }
        });
    }

    private login()
    {
        this.setState({loginDisabled: true});

        if(!this.isInIframe())
        {
            this.getTokenFromOfficeContext();
        }
        else
        {
            this.getTokenFromTeams();
        }
    }

    render() 
    {
        const {
            logo,
            title,
            localizationService
        } = this.props;

        const { token, loginDisabled, error } = this.state;

        if(!token)
        {
            const errorMessage = () =>
            {
                if(error)
                {
                    return <div className='ms-fontSize-m' style={{paddingBottom:'10px'}}>
                            <span>{localizationService.getString("LoginError")}</span> <br/>
                            <span>{error}</span>
                        </div>;
                }

                return null;
            };
            
            return (
                <div>
                    <section className='ms-welcome__progress ms-u-fadeIn500'>
                        <img width='90' height='90' src={logo} alt={title} title={title} />
                        <h1 className='ms-fontSize-xxl ms-fontWeight-light ms-fontColor-neutralPrimary'>{title}</h1>
                        {errorMessage()}
                        <DefaultButton onClick={this.login} disabled={loginDisabled}>
                            {localizationService.getString("Login")}
                        </DefaultButton>
                    </section>
                </div>
            );
        }
        else
        {
            return (
                <Main token={token} localizationService={localizationService}/>
            )
        }
    }
}
