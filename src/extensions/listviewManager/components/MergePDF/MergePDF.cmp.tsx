import React, { useEffect, useState } from 'react';
import Box from '@mui/material/Box';
import Button from '@mui/material/Button';
import Typography from '@mui/material/Typography';
import Modal from '@mui/material/Modal';
import styles from './MergePDF.module.scss';
import modalStyles from "../../styles/modalStyles.module.scss";
import DownloadIcon from '@mui/icons-material/Download';
import SaveIcon from '@mui/icons-material/Save';
import CloseIcon from '@mui/icons-material/Close';
import { SPFI } from '@pnp/sp';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';
import toast from 'react-hot-toast';
import { ConvertToPdf, downloadPdf } from '../../service/mergePdf.service';
import TreeFolders from '../TreeFolders/TreeFolders';

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

export interface MergePDFProps {
    selectedItems: any;
    context: ListViewCommandSetContext;
    unMountDialog: () => void;
}

const MergePDF: React.FC<MergePDFProps> = ({ selectedItems, context, unMountDialog }) => {
    const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
    const [fileName, setFileName] = useState<string>("CombinedFiles.pdf");

    const download = async () => {
        unMountDialog();
        let fileBlob: Blob;
        await toast.promise(
            ConvertToPdf(context, selectedItems).then((pdfBlob) => {
                fileBlob = pdfBlob;
            }),
            {
                loading: 'מאחד קבצים ל-PDF...',
                success: '',
                error: 'אירעה שגיאה בעת יצירת ה-PDF. אנא נסה שוב',
            }
        );
        downloadPdf(fileBlob);
    };

    const saveToSharepoint = () => {
        setFileName(fileName);
        setIsModalOpen(true);
    };

    return (
        <div>
            <Modal
                open={!isModalOpen}
                onClose={unMountDialog}
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
                    <div className={`${modalStyles.modalFooter} ${modalStyles.modalFooterSpaceAround}`}>
                        <Button
                            color="error"
                            onClick={unMountDialog}
                            className={`${styles.button}`}
                            startIcon={<CloseIcon style={{ color: "#f58383", paddingLeft: "8px", margin: "0px !important" }} />}>
                            ביטול
                        </Button>
                        <Button
                            onClick={download}
                            className={`${styles.button}`}
                            endIcon={<DownloadIcon style={{ color: '#1976d2', paddingRight: "8px" }} />}>
                            הורדה למחשב
                        </Button>
                        <Button
                            color='success'
                            endIcon={<SaveIcon style={{ color: '#2e7d32', paddingRight: "8px" }} />}
                            onClick={saveToSharepoint}
                            className={`${styles.button}`}>
                            שמירה באתר
                        </Button>
                    </div>
                </Box>
            </Modal>

            {/* Render the TreeFolders component if the modal for saving is open */}
            {isModalOpen && (
                <TreeFolders
                    context={context}
                    isClose={unMountDialog}
                    fileToSave={selectedItems}
                    fileName={fileName}
                />
            )}
        </div>
    );
};

export default MergePDF;
