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

const style = {
    position: 'absolute',
    top: '50%',
    left: '50%',
    transform: 'translate(-50%, -50%)',
    width: 400,
    bgcolor: 'background.paper',
    boxShadow: 24,
    p: 4,
};

const buttonContainerStyle = {
    display: 'flex',
    gap: '3em',
    justifyContent: 'center',
    paddingTop: '2em'
};


export default class ExportZipModal extends React.Component<ExportZipModalProps, ExportZipModalState> {
    constructor(props: any) {
        super(props);
        this.state = {

        }
    }

    private download = async () => {
        const archive = await exportToZip(this.props.selectedItems, this.props.context);
        downloadToPC(archive);
        this.props.unMountDialog();
    }

    private saveToSharepoint = async () => {
        const archive = await exportToZip(this.props.selectedItems, this.props.context);
        saveZipToSharePoint(archive, this.props.selectedItems, this.props.sp);
        this.props.unMountDialog();
    }

    public render(): React.ReactElement<ExportZipModalProps> {
        const font = 'Rubik';
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
                            variant="h6"
                            component="h2"
                            align='center'
                            fontFamily={font}>
                            איפה תרצה לשמור את הקובץ?
                        </Typography>
                        <Box sx={buttonContainerStyle}>
                            <Button
                                variant="contained"
                                color="primary"
                                onClick={async () => this.download()}
                                sx={{ fontFamily: font }}
                                className={`${styles.button} ${styles.downloadButton}`}>
                                הורדה למחשב
                            </Button>
                            <Button
                                variant="contained"
                                color="error"
                                onClick={this.props.unMountDialog}
                                sx={{ fontFamily: font }}
                                className={`${styles.button} ${styles.cancelButton}`}>
                                ביטול
                            </Button>
                            <Button
                                variant="contained"
                                style={{ backgroundColor: "#84C792" }}
                                onClick={() => this.saveToSharepoint()}
                                sx={{ fontFamily: font }}
                                className={`${styles.button} ${styles.saveButton}`}>
                                שמירה ב-Sharepoint
                            </Button>
                        </Box>
                    </Box>
                </Modal>
            </div>
        );
    }
}