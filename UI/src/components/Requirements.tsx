import * as React from 'react';

export interface IRequirementsProps
{
    invalidVersion: boolean;
    invalidDocument: boolean;
}

export class Requirements extends React.Component<IRequirementsProps>
{
    constructor(props)
    {
        super(props);
    }

    render() 
    {
        const { invalidDocument, invalidVersion} = this.props;
        
        const renderInvalidDoc = () => { 
            return (
                <div className="ms-welcome__requirement">
                    <div className="ms-font-m">The add-ins need to open in the word from SharePoint site. Please close current open word file and follow the instructions below:</div>
                    <div className="ms-font-m">1. Login O365 site where the word documents are stored.</div>
                    <div className="ms-font-m">2. Select the word file and click "Open in Word" link.</div>
                    <img src="dist/assets/mode-word.png"/>
            </div>);
        };

        const renderInvalidVersion = () =>
        {
            return <div className="ms-welcome__requirement">
                <div className="ms-font-m">The version of the client word should be 2016 or higher.</div>
            </div>;
        };

        if(!invalidDocument)
        {
            return renderInvalidDoc();
        }
        
        if(invalidVersion)
        {
            return renderInvalidVersion();
        }

        const both = [];
        both.push(renderInvalidDoc());
        both.push(renderInvalidVersion());

        return <div>{both}</div>;
    }
}