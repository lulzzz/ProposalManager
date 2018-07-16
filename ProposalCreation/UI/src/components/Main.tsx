import * as React from 'react';
import * as $ from 'jquery'; 
import {
    Pivot,
    PivotItem,
    PivotLinkFormat
} from 'office-ui-fabric-react/lib/Pivot';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Home } from './Home';
import { Documents } from './Documents';
import { Notes } from './Notes';
import { ErrorBoundary } from './ErrorBoundary';

export interface IMainProps
{
    token: string
}

export class Main extends React.Component<IMainProps> {
    constructor(props, context) {
        super(props, context);
    }

    componentDidMount()
    {
        $(".icon_085d752b").hide();
        $(".dismissSingleLine_085d752b").hide();
    }

    public render() {
        const paddingLeft = {
            paddingLeft: "5px"
        };
        
        const log = (text: string): () => void =>
        (): void => console.log(text);

        return (
            <div className='ms-welcome'>
                <MessageBar
                    messageBarType={ MessageBarType.info }
                    isMultiline={ false }
                    onDismiss={ log('test') }
                    dismissButtonAriaLabel='Close'
                    truncated={ true }
                    overflowButtonAriaLabel='Overflow'
                    >
                    Opportunity description
                </MessageBar>
                <div className='ms-welcome__pivot' style={paddingLeft}>
                    <ErrorBoundary>
                        <Pivot linkFormat={PivotLinkFormat.links}>
                            <PivotItem linkText='Home' >
                                <Home token={this.props.token}/>
                            </PivotItem>

                            <PivotItem linkText='Documents'>
                            <Documents token={this.props.token}/>
                            </PivotItem>

                            <PivotItem linkText='Call Reports'>
                                <Notes token={this.props.token}/>
                            </PivotItem>
                        </Pivot>
                    </ErrorBoundary>
                </div>
            </div>
        );
    };
};