import * as React from 'react';
import styles from '../EntraContacts.module.scss';
import { UserCardProps } from './userCardProps';


const UserCard: React.FC<UserCardProps> = (props) => {
    const { user } = props;

    return (
        <div className={styles.userCard}>
            <img src={user.photoUrl} alt={user.displayName} />
            <div className={styles.userInfo}>
                <p className={styles.name} >{user.displayName}</p>
                <p className={styles.email} >{user.userPrincipalName}</p>
            </div>
        </div>
    );
}

export default UserCard;