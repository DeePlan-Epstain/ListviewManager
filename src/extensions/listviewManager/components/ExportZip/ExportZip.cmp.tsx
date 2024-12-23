import * as React from 'react';
import Box from '@mui/material/Box';
import Button from '@mui/material/Button';
import Typography from '@mui/material/Typography';
import Modal from '@mui/material/Modal';
import styles from './ExportZip.module.scss'
import { ExportZipModalProps } from './ExportZipProps';
import { ExportZipModalState } from './ExportZipState';
import { downloadToPC, exportToZip, saveZipToSharePoint } from '../../service/zip.service';
import { FolderPicker, IFolder } from "@pnp/spfx-controls-react/lib/FolderPicker";
import Swal from 'sweetalert2';
import modalStyles from "../../styles/modalStyles.module.scss";


const style = {
    position: 'absolute',
    top: '50%',
    left: '50%',
    transform: 'translate(-50%, -50%)',
    width: 400,
    bgcolor: 'background.paper',
    boxShadow: 24,
    p: 4,
    borderRadius: '8px',
};

const buttonContainerStyle = {
    display: 'flex',
    gap: '3em',
    justifyContent: 'center',
    paddingTop: '1em'
};


export default class ExportZipModal extends React.Component<ExportZipModalProps, ExportZipModalState> {
    constructor(props: any) {
        super(props);
        this.state = {

        }
    }

    private download = async () => {
        this.props.unMountDialog();
        const archive = await exportToZip(this.props.selectedItems, this.props.context);
        Swal.fire({
            title: "יצירת הקובץ בוצעה בהצלחה",
            text: "ההורדה תחל בשניות הקרובות",
            icon: "success",
        });
        downloadToPC(archive);
    }

    private saveToSharepoint = async () => {
        this.props.unMountDialog();
        const archive = await exportToZip(this.props.selectedItems, this.props.context);
        Swal.fire({
            title: "יצירת הקובץ בוצעה בהצלחה",
            text: "ההורדה תחל בשניות הקרובות",
            icon: "success",
        });
        saveZipToSharePoint(archive, this.props.selectedItems, this.props.sp);
    }

    public render(): React.ReactElement<ExportZipModalProps> {
        const font = 'Tahoma';
        return (
            <div>
                <Modal
                    open={this.props.status}
                    onClose={this.props.unMountDialog}
                    aria-labelledby="modal-modal-title"
                    aria-describedby="modal-modal-description"
                >
                    <Box sx={style}>
                        <Typography
                            id="modal-modal-title"
                            align='center'
                            className={styles.modal_title}>
                            איפה תרצה לשמור את הקובץ?
                        </Typography>
                        <Typography
                            id="modal-modal-title"
                            align='center'
                            className={styles.modal_text}>
                            פעולה זו עשויה לקחת זמן, בסיום היצירה תופיע התראה.
                        </Typography>
                        <div className={modalStyles.modalFooter}>
                        <Button
                                color="error"
                                onClick={this.props.unMountDialog}
                                className={`${styles.button}`}>
                                ביטול
                            </Button>
                            <Button
                                onClick={async () => this.download()}
                                className={`${styles.button}`}>
                                הורדה למחשב
                            </Button>
                            <Button
                                color='success'
                                onClick={() => this.saveToSharepoint()}
                                className={`${styles.button}`}>
                                שמירה באתר
                            </Button>
                        </div>
                    </Box>
                </Modal>
            </div>
        );
    }
}