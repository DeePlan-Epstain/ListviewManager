
import * as React from 'react';
import { ModalExtProps } from "./ModalExtProps"
import { ModalExtStates } from "./ModalExtstates"
import "./style.css"
import "@pnp/sp/site-users/web";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Box, TextField, Button, Dialog, DialogTitle, DialogContent, DialogContentText, DialogActions } from "@material-ui/core"
import { spfi, SPFI, SPFx } from '@pnp/sp';
import { Web } from '@pnp/sp/webs';
import "@pnp/sp/webs";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/folders";
import Autocomplete from "@material-ui/lab/Autocomplete";
import { ThemeProvider, StylesProvider } from "@material-ui/core/styles";
import { jss } from "../../models/jss";
import { theme } from "../../models/theme";

import CircularProgress from '@mui/material/CircularProgress';



export default class ModalCreateProject extends React.Component<ModalExtProps, ModalExtStates> {
    private _sp: SPFI;
    private _spWithCustomUrl: SPFI;

    constructor(props: ModalExtProps | Readonly<ModalExtProps>) {
        super(props)
        this.state = {
            open: false,
            error: false,
            isSave: false,
            FoldersHierarchy: [],
            FolderHierarchy: {},
            FolderHierarchyValidate: false,
            NewFolderName: "",
            NewFolderNameValidate: false,


        }
    }

    componentDidUpdate(prevProps: Readonly<ModalExtProps>, prevState: Readonly<ModalExtStates>, snapshot?: any): void {
        document.getElementById("modal-back2").className += " show-modal-back2"
        document.getElementById("modal-content2").className += " show-modal-content2"
    }
    componentDidMount(): void {
        this._sp = spfi().using(SPFx(this.props.context));
        this._spWithCustomUrl = spfi("https://epstin100.sharepoint.com/sites/EpsteinPortal/").using(SPFx(this.props.context));

        this.setState({
            open: true
        })
        this.ResetForm();
    }

    ResetForm = async () => {

        document.getElementById("modal-back2").className += " show-modal-back2" // to open modal.
        document.getElementById("modal-content2").className += " show-modal-content2" // to open modal.

        try {

            const listFolders = await this._spWithCustomUrl.web.lists.getByTitle("FolderHierarchy").rootFolder.folders();

            let listFoldersNoForms = listFolders.filter((folder: any) => folder.Name !== "Forms");

            console.log(listFoldersNoForms);


            // listFolders.forEach(async (item: any) => {
            //     console.log(item.Name);
            //     const destinationUrl = `${this.props.selectedRow.FileRef}/${item.Name}`;
            //     await this._sp.web.rootFolder.folders.getByUrl(`FolderHierarchy`).folders.getByUrl(`${item.Name}`).copyByPath(destinationUrl, true);
            // })
            this.setState({
                FoldersHierarchy: listFoldersNoForms,

            })
        } catch (err) {
            this.setState({

                error: true
            })
        }

        // this.props.unMountDialog()
    }

    createFolder = async () => {
        this.setState({ isSave: true })
        if (this.validate()) {
        
         const destinationUrl = `${this.props.finalPath}/${this.state.NewFolderName}`;
         console.log("destinationUrl",destinationUrl);
        

         
         await this._spWithCustomUrl.web.rootFolder.folders.getByUrl(`FolderHierarchy`).folders.getByUrl(`${this.state.FolderHierarchy.Name}`).copyByPath(destinationUrl, true);
            
         this.closeModal();
         location.reload();

        
        } else {
            this.setState({ isSave: true })
        }
    }

    validate = () => {
        const isEmpty = (value: any) => value === "" || value === undefined || value === null;
        let flag = true;

        // Validate FolderHierarchy
        if (isEmpty(this.state.FolderHierarchy?.Name)) {
            this.setState({ FolderHierarchyValidate: true });
            flag = false;
        }

        // Validate NewFolderName
        if (isEmpty(this.state.NewFolderName)) {
            this.setState({ NewFolderNameValidate: true });
            flag = false;
        }

        return flag;
    };


    clearModal = () => {
        this.setState({

        })
        return 0;
    }

    closeModal = async () => { // close the modal.

        let r: number = await this.clearModal();

        document.getElementById("modal-back2").className = "modal-back2"
        document.getElementById("modal-content2").className = "modal-content2"

    }

    onchange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const { name, value } = e.target;
    
        // עדכון ה-state עבור השדה שהשתנה
        const newState = { [e.target.name]: e.target.value } as unknown as Pick<ModalExtStates, keyof ModalExtStates>;
    
        // בדיקת ולידציה עבור NewFolderName
        if (name === "NewFolderName") {
            const hasInvalidChars = this.validateFolderName(value);
            newState.NewFolderNameValidate = hasInvalidChars; // true אם יש תווים לא חוקיים
        }
    
        this.setState(newState);
    };
    validateFolderName = (value:any) => {
        const invalidChars = /[\\/:*?"<>|]/; // תווים שאינם חוקיים
        return invalidChars.test(value);
      };

    public render(): React.ReactElement<ModalExtProps> {
        return (
            <StylesProvider jss={jss}>
                <ThemeProvider theme={theme}>


                    <div id='modal-back2' className='modal-back2' onClick={this.closeModal}>
                        <div style={{ display: "flex", justifyContent: "center", flexDirection: "column" }} id='modal-content2' className='modal-content2' onClick={(e) => { e.stopPropagation() }}>
                            <div className="modal-header" >
                                <h2  style={{margin:1}}>יצירת היררכית תיקיות</h2>
                            </div>
                            <div  className="modal-body">
                                <Autocomplete
                                    id="country-select-demo"
                                    onChange={(event, newValue) => {
                                        console.log(newValue);

                                        let s = this.state.FoldersHierarchy.find((folder) => folder.Name === newValue)


                                        this.setState({ FolderHierarchy: s, FolderHierarchyValidate: false });
                                    }}
                                    value={this.state.FolderHierarchy?.Name || ""}
                                    options={this.state.FoldersHierarchy.map((folder) => folder.Name)}
                                    renderInput={(params) => (
                                        <TextField
                                            {...params}
                                            variant="outlined"
                                            size="small"
                                            label="בחר היררכית תיקיות"
                                            fullWidth
                                            error={this.state.FolderHierarchyValidate}

                                        />
                                    )}
                                />
                                <TextField
                                    name="NewFolderName"
                                    error={this.state.NewFolderNameValidate}
                                    helperText={
                                      this.state.NewFolderNameValidate
                                        ? "שם התיקייה אינו יכול לכלול תווים כמו \\ / : * ? \" < > |"
                                        : ""
                                    }
                                    
                                    onChange={this.onchange}
                                    value={this.state.NewFolderName}
                                    className="text-field"
                                    id="outlined-basic"
                                    variant="outlined"
                                    label="שם התיקיה"
                                    size="small"
                                    fullWidth
                                    style={{ marginTop: '16px' }}
                                />
                            <div style={{fontSize: 12, textAlign: "right", direction: "rtl",marginTop:7}}>
    *יש לבחור את סוג ההיררכיה ואת שם התיקיה החדשה
</div>
                            </div>
                            <div className="modal-footer">
                                <Button
                                    variant="contained"
                                    style={{ backgroundColor: "green", color: "white" }}
                                    disabled={this.state.isSave}
                                    onClick={this.createFolder}
                                >
                                    אישור
                                </Button>
                                <Button
                                    variant="outlined"
                                    style={{ borderColor: "red", color: "red" }}

                                    onClick={this.closeModal}
                                >
                                    ביטול
                                </Button>
                            </div>
                            <div>
                            {this.state.isSave &&
                                <Box sx={{ display: 'flex', justifyContent:"center",margin:10 }}>
                            <CircularProgress />
                            </Box>
                            }

                            </div>
                            

                        </div >
                    </div >
                </ThemeProvider>
            </StylesProvider>
        );
    }
}
