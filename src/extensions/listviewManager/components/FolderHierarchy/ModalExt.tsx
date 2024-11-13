
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





export default class ModalCreateProject extends React.Component<ModalExtProps, ModalExtStates> {
    private _sp: SPFI;
    constructor(props: ModalExtProps | Readonly<ModalExtProps>) {
        super(props)
        this.state = {
            open: false,
            error: false
        }
    }

    componentDidUpdate(prevProps: Readonly<ModalExtProps>, prevState: Readonly<ModalExtStates>, snapshot?: any): void {
        document.getElementById("modal-back2").className += " show-modal-back2"
        document.getElementById("modal-content2").className += " show-modal-content2"
    }
    componentDidMount(): void {
        this._sp = spfi().using(SPFx(this.props.context));
        this.setState({
            open: true
        })
        this.ResetForm();
    }

    ResetForm = async () => {

        document.getElementById("modal-back2").className += " show-modal-back2" // to open modal.
        document.getElementById("modal-content2").className += " show-modal-content2" // to open modal.

        try {

            const listFolders = await this._sp.web.lists.getByTitle("FolderHierarchy").rootFolder.folders();
            console.log(listFolders);
            
            listFolders.forEach(async (item: any) => {
                console.log(item.Name);
                
                const destinationUrl = `${this.props.selectedRow.FileRef}/${item.Name}`;
                await this._sp.web.rootFolder.folders.getByUrl(`FolderHierarchy`).folders.getByUrl(`${item.Name}`).copyByPath(destinationUrl, true);
            })
        }catch(err){
            this.setState({

                    error:true

            })
        }


        // this.props.unMountDialog()


    }
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

    public render(): React.ReactElement<ModalExtProps> {
        return (
            <div>


                <div id='modal-back2' className='modal-back2' onClick={this.closeModal}>
                    <div style={{ display: "flex", justifyContent: "center" }} id='modal-content2' className='modal-content2' onClick={(e) => { e.stopPropagation() }}>
                        {this.state.error ?



                            <div className="error alert">
                                <div className="alert-body">
                                    Hierarchical folder exit failed Contact the system administrator.
                                </div>
                            </div>

                            :
                            <div className="success alert">
                                <div className="alert-body">
                                    The folder hierarchy has been created successfully.

                                </div>
                            </div>

                        }
                    </div>
                </div>

            </div >
        );
    }
}
