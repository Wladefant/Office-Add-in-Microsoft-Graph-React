import * as React from 'react';
import { Button } from 'office-ui-fabric-react';

export interface HeaderProps {
    title: string;
    logo: string;
    message: string;
    onBackClick: () => void;
    onForwardClick: () => void;
}

export default class Header extends React.Component<HeaderProps> {
    render() {
        const { title, logo, message, onBackClick, onForwardClick } = this.props;

        return (
            <section className='ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500'>
                <Button iconProps={{ iconName: 'Back' }} onClick={onBackClick} />
                <Button iconProps={{ iconName: 'Forward' }} onClick={onForwardClick} />
                <img width='80' height='80' src={logo} alt={title} title={title} />
                <h1 className='ms-fontSize-xxl ms-fontWeight-light ms-fontColor-neutralPrimary'>{message}</h1>
            </section>
        );
    }
}
