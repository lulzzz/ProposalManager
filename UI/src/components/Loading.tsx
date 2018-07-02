import * as React from 'react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib';

export interface ILoadingProps
{
    message: string;
    overlay?: boolean;
}

export class Loading extends React.Component<ILoadingProps>
{
    render()
    {
        const { message, overlay } = this.props;

        const spinner = <div className="text ms-font-xxl" style={{paddingTop: '50px'}}>
                            <Spinner size={SpinnerSize.large} label={message} />
                        </div>;

        if(overlay)
        {
            return (
                <div id="popup" className="overlay on">
                    {spinner}
                </div>
            );
        }

        return (
            <div>
                {spinner}
            </div>
        );
    }
}