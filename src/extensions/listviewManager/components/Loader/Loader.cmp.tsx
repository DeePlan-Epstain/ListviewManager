import React from 'react';
import styles from './Loader.module.scss';

interface LoaderProps {
    msg: string;
}

export default function Loader({ msg }: LoaderProps) {

    return (
        <div className={styles.loaderContainer}>
            <div className={styles.spinner}></div>
            <div className={styles.loaderMsg}>
                <span>{msg}</span>
            </div>
        </div>
    );
}