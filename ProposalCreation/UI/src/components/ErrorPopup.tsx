import * as $ from 'jquery';
import * as React from 'react';
import { IconButton } from '../../node_modules/office-ui-fabric-react/lib';

export interface IErrorPopupProps
{
    error: any;
}

export class ErrorPopup extends React.Component<IErrorPopupProps>
{
    render()
    {
        const { error } = this.props;

        return (
            <div id="errorPopup" className="overlay on">
                <div className="errorContainer">
                    <div className="errorHeader">
                        <span className="errorTitle ms-font-m">An error occurred</span>
                        <IconButton
                            iconProps={ { iconName: 'ChromeClose' } }
                            title='Close'
                            ariaLabel='ChromeClose'
                            onClick={
                                (e) => {
                                    e.preventDefault();
                                    $("#errorPopup").removeClass("on").addClass("off");
                                }
                            }
                        />
                    </div>
                    <details className="errorDetails ms-font-s" style={{ whiteSpace: 'pre-wrap' }}>
                        {error.responseText}
                    </details>
                </div>
            </div>
        );
    }
}