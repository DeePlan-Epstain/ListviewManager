
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
import modalStyles from "../../styles/modalStyles.module.scss";
import CloseIcon from '@mui/icons-material/Close';
import CheckIcon from '@mui/icons-material/Check';
import CircularProgress from '@mui/material/CircularProgress';
import { IconButton } from '@mui/material';
import { light } from '@mui/material/styles/createPalette';



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
            success: false,
            FoldersHierarchyAfterChoosingDivision: [],
            DivisionValidate: false,
            Division: ""
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

            const listFolders = await this._spWithCustomUrl.web.lists.getById("430ade62-95d9-4540-99bd-e834dc4f55b5")
                .rootFolder.folders
                .select("Name", "ListItemAllFields/unit") // ציון השדות שברצונך לקבל
                .expand("ListItemAllFields") // הרחבה כדי לקבל את המטה-דאטה
                ();




            let listFoldersNoForms = listFolders.filter((folder: any) => folder.Name !== "Forms");




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
            try {
                
                const destinationUrl = `${this.props.finalPath}/${this.state.NewFolderName}`;
                const rootSiteUrl = `${window.location.origin}/sites/${window.location.pathname.split('/')[2]}`;
                const libraryName = this.props.finalPath.replace(/.*\/sites\/[^\/]+\//, '');
                console.log(rootSiteUrl);
                console.log(libraryName);

                
                console.log("destinationUrl", destinationUrl);
                const flowUrl = "https://prod-213.westeurope.logic.azure.com:443/workflows/5150c75426064ee28ba6dbad9948394d/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=T_aykz6E_7pktPhB30vv1u95t-568u8WoHprU6A6q7w";
                const requestBody = {
                    FolderHierarchyName: this.state.FolderHierarchy?.Name,
                    ToSiteUrl: rootSiteUrl,
                    LibraryTo: libraryName,
                    NewFolderName: this.state.NewFolderName,

                };
                const response = fetch(flowUrl, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify(requestBody)
                });
                //                const data = await response.json();

                this.setState({ isSave: false, success: true });




                setTimeout(() => {
                    this.closeModal();
                    location.reload();
                }, 3000);

            } catch (e) {
                console.log(e);

                this.setState({ isSave: false });

            }

        } else {
            this.setState({ isSave: false });
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
        if (isEmpty(this.state.Division)) {
            this.setState({ DivisionValidate: true });
            flag = false;
        }

        return flag;
    };


    clearModal = () => {
        this.setState({
            isSave: false,

            FolderHierarchy: {},
            FolderHierarchyValidate: false,
            NewFolderName: "",
            NewFolderNameValidate: false,
        })
        return 0;
    }

    closeModal = async () => { // close the modal.

        this.setState({
            isSave: false,
            success: false,
            DivisionValidate: false,
            Division: "",
            FolderHierarchy: {},
            FolderHierarchyValidate: false,
            NewFolderName: "",
            NewFolderNameValidate: false,
        }, () => {
            document.getElementById("modal-back2").className = "modal-back2"
            document.getElementById("modal-content2").className = "modal-content2"
        })
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
    validateFolderName = (value: any) => {
        const invalidChars = /[\\/:*?"<>|]/; // תווים שאינם חוקיים
        return invalidChars.test(value);
    };

    AfterChoosingDivision = (value: any) => {

        this.setState({ FoldersHierarchyAfterChoosingDivision: this.state.FoldersHierarchy.filter((folders: any) => { return folders.ListItemAllFields.unit === value }), Division: value, DivisionValidate: false })

    }

    public render(): React.ReactElement<ModalExtProps> {
        return (
            <StylesProvider jss={jss}>
                <ThemeProvider theme={theme}>


                    <div id='modal-back2' className='modal-back2' >
                        <div style={{ display: "flex", justifyContent: "center", flexDirection: "column" }} id='modal-content2' className='modal-content2' onClick={(e) => { e.stopPropagation() }}>
                            {!this.state.success &&
                                <>
                                    <div className="modal-header" >
                                        <h2 style={{ margin: 1 }}>יצירת היררכית תיקיות</h2>
                                    </div>

                                    <div className="modal-body">
                                        <Autocomplete
                                            id="country-select-demo"
                                            onChange={(event, newValue) => {


                                                this.AfterChoosingDivision(newValue)

                                            }}
                                            value={this.state.Division || ""}
                                            options={Array.from(
                                                new Set(
                                                    this.state.FoldersHierarchy
                                                        .filter(folder => folder.ListItemAllFields.unit) // מסנן ערכים ריקים
                                                        .map(folder => folder.ListItemAllFields.unit)    // ממפה את היחידה
                                                )
                                            )}
                                            renderInput={(params) => (
                                                <TextField
                                                    {...params}
                                                    variant="outlined"
                                                    size="small"
                                                    label="בחר חטיבה"
                                                    fullWidth
                                                    error={this.state.DivisionValidate}

                                                />
                                            )}
                                        />
                                        <Autocomplete
                                            id="country-select-demo"
                                            onChange={(event, newValue) => {
                                                console.log(newValue);

                                                let s = this.state.FoldersHierarchy.find((folder) => folder.Name === newValue)


                                                this.setState({ FolderHierarchy: s, FolderHierarchyValidate: false });
                                            }}
                                            value={this.state.FolderHierarchy?.Name || ""}
                                            options={this.state.FoldersHierarchyAfterChoosingDivision.map((folder) => folder.Name)}
                                            renderInput={(params) => (
                                                <TextField
                                                    {...params}
                                                    variant="outlined"
                                                    size="small"
                                                    label="בחר היררכית תיקיות"
                                                    fullWidth
                                                    error={this.state.FolderHierarchyValidate}
                                                    style={{ marginTop: '16px' }}

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
                                        <div style={{ fontSize: 12, textAlign: "right", direction: "rtl", marginTop: 7 }}>
                                            *יש לבחור את סוג ההיררכיה ואת שם התיקיה החדשה
                                        </div>
                                    </div>


                                    <div className={modalStyles.modalFooter} style={{ padding: "0 24%" }}>
                                        <Button
                                            disabled={this.state.isSave}
                                            onClick={this.closeModal}
                                            style={{ color: 'red' }}
                                            startIcon={<CloseIcon style={{ color: "#f58383", paddingLeft: 0, margin: "0px !important" }} />}
                                        >
                                            ביטול
                                        </Button>
                                        <Button
                                            style={{ color: '#1976d2' }}
                                            disabled={this.state.isSave}
                                            onClick={this.createFolder}
                                            endIcon={<CheckIcon style={{ color: this.state.isSave ? 'inherit' : '#1976d2', margin: "0px" }} />}
                                        >
                                            אישור
                                        </Button>
                                    </div>
                                    <div>
                                        {this.state.isSave &&
                                            <Box sx={{ display: 'flex', justifyContent: "center", margin: 10 }}>
                                                <CircularProgress />
                                            </Box>
                                        }

                                    </div>
                                </>
                            }
                            {this.state.success &&
                                <div style={{
                                    fontSize: '19px',
                                    textAlign: 'center',
                                    direction: 'rtl',
                                    marginTop: '10px',
                                    padding: '10px',
                                    color: 'green',
                                    borderRadius: '5px',
                                }}>
                                    ✅ הבקשה התקבלה בהצלחה! התיקייה תיווצר בעוד מספר דקות שם הפרויקט יתעדכן בסיום התהליך.
                                </div>

                            }
                        </div >
                    </div >
                </ThemeProvider>
            </StylesProvider>
        );
    }
}
