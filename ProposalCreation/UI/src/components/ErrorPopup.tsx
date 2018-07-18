import * as $ from 'jquery';
import * as React from 'react';
import { IconButton } from '../../node_modules/office-ui-fabric-react/lib';
import { LocalizationService } from '../services/LocalizationService';

export interface IErrorPopupProps
{
    error: any;
    localizationService: LocalizationService;
}

export class ErrorPopup extends React.Component<IErrorPopupProps>
{
    render()
    {
        const { error, localizationService } = this.props;

        return (
            <div id="errorPopup" className="overlay on">
                <div className="errorContainer">
                    <div className="errorHeader">
                        <span className="errorTitle ms-font-m">{localizationService.getString("ErrorHeader")}</span>
                        <IconButton
                            iconProps={ { iconName: 'ChromeClose' } }
                            title={localizationService.getString("Close")}
                            ariaLabel='ChromeClose'
                            onClick={
                                (e) => {
                                    e.preventDefault();
                                    $("#errorPopup").removeClass("on").addClass("off");
                                }
                            }
                        />
                    </div>
                    <details className="errorDetails ms-font-s" style={{ whiteSpace: 'pre-wrap' }} open>
                        {error.message ? error.message : error.responseText}
                    </details>
                </div>
            </div>
        );
    }
}