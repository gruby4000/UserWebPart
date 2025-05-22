import * as React from 'react';
import styles from './UserInfoWebpart.module.scss';
import type { IUserInfoWebpartProps } from './IUserInfoWebpartProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class UserInfoWebpart extends React.Component<IUserInfoWebpartProps> {
    public render(): React.ReactElement<IUserInfoWebpartProps> {
        const {
            hasTeamsContext,
            userDisplayName,
            siteTitleFontName 
        } = this.props;

        const nameParts = userDisplayName ? userDisplayName.split(' ') : [];
        const firstName = nameParts[0] || ''; 

        return (
            <section className={`${styles.userInfoWebpart} ${hasTeamsContext ? styles.teams : ''}`}>
                <div className={styles.welcome} style={{ fontFamily: siteTitleFontName, fontSize: this.props.fontSize || '16px' }}>
                    <div>
                        <p>User name: {escape(firstName)}</p>
                    </div>
                </div>
            </section>
        );
    }
}
