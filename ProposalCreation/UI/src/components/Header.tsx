import * as React from 'react';

export interface HeaderProps {
    message: string;
}

export default class Header extends React.Component<HeaderProps> {
    render() {
        const {
            message
        } = this.props;

        return (
            <section className='ms-welcome__header ms-bgColor-neutralLighter'>
                <h1 className='ms-fontSize-l ms-fontWeight-light ms-fontColor-neutralPrimary'>{message}</h1>
            </section>
        );
    }
}
