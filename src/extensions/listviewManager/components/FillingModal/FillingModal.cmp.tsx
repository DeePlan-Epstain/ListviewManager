import React from 'react';
import modalStyles from "../../styles/modalStyles.module.scss";
import { SPFI } from '@pnp/sp';
import { Button, IconButton } from '@mui/material';
import { Close } from '@mui/icons-material';

export interface FillingModalProps {
    sp: SPFI;
    foldersMap: Map<string, File[]>;
    closeModal: () => void;
};

// const 

export function FillingModal({ sp, foldersMap, closeModal }: FillingModalProps) {

    const submit = async () => { };

    return (
        <div className={modalStyles.modalScreen} dir="ltr">
            <div className={modalStyles.modal} style={{ width: '85%', height: '75%' }} onClick={(ev: any) => ev.stopPropagation()}>
                <div className={modalStyles.modalHeader}>
                    <span>תיוק מסמכים</span>

                    <IconButton onClick={closeModal}>
                        <Close />
                    </IconButton>
                </div>
                <div>

                </div>
                <div className={modalStyles.modalFooter}>
                    <Button onClick={submit}>תיוק</Button>
                    <Button onClick={closeModal} color='error'>ביטול</Button>
                </div>
            </div>
        </div>
    )
};