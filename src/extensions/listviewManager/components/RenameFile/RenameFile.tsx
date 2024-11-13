// import * as React from "react";
// import {
//   Dialog,
//   DialogType,
//   DialogFooter,
// } from "office-ui-fabric-react/lib/Dialog";
// import {
//   PrimaryButton,
//   DefaultButton,
// } from "office-ui-fabric-react/lib/Button";
// import { TextField } from "office-ui-fabric-react/lib/TextField";
// import { SPFI } from "@pnp/sp";
// import { SelectedFile } from "../../models/global.model";

// export interface IRenameFileProps {
//   sp: SPFI;
//   context: any;
//   libraryName: any;
//   libraryID: any;
//   selectedRow: any;
//   unMountDialog: () => void;
// }

// export interface IRenameFileState {
//   newName: string;
//   setNewName: string;
//   showDialog: boolean;
//   showNotification: boolean;
//   isExceedingLimit: boolean;
//   isInputEmpty: boolean;
// }

// export default class RenameFile extends React.Component<
//   IRenameFileProps,
//   IRenameFileState
// > {
//   constructor(props: IRenameFileProps) {
//     super(props);
//     this.state = {
//       newName: "",
//       setNewName: "",
//       showDialog: true,
//       showNotification: false,
//       isExceedingLimit: false,
//       isInputEmpty: false,
//     };
//     this.handleInputChange = this.handleInputChange.bind(this);
//     this.handleRenameClick = this.handleRenameClick.bind(this);
//     this.closeDialog = this.closeDialog.bind(this);
//   }

//   componentDidMount(): void {
//     this.fileOrFolderNameHandler();
//   }

//   fileOrFolderNameHandler = () => {
//     const { selectedRow } = this.props;
//     console.log("selectedRow:", selectedRow);
//     const fileName = selectedRow.FileLeafRef;
//     const newName = fileName.substring(0, fileName.lastIndexOf("."));
//     const isFolder = selectedRow.FSObjType === "1";
//     const name = isFolder ? selectedRow.FileLeafRef : newName;

//     this.setState({
//       newName: name,
//       setNewName: selectedRow.FileLeafRef + "." + selectedRow.File_x0020_Type,
//     });
//   };

//   private handleInputChange(event: React.ChangeEvent<HTMLInputElement>) {
//     const inputValue = event.target.value;
//     if (inputValue.length <= 200) {
//       this.setState({
//         setNewName:
//           this.props.selectedRow.File_x0020_Type === ""
//             ? inputValue
//             : inputValue + "." + this.props.selectedRow.File_x0020_Type,
//         newName: inputValue,
//         isExceedingLimit: false,
//         isInputEmpty: false,
//       });
//     } else {
//       this.setState({ isExceedingLimit: true });
//     }
//   }

//   private async handleRenameClick() {
//     const { newName, setNewName } = this.state;

//     if (newName.trim() === "") {
//       this.setState({ isInputEmpty: true });
//       return;
//     } else {
//       try {
//         await this.props.sp.web.lists
//           .getById("8f3e6fdc-eeb8-4171-801a-cda59f667a92")
//           .items.add({
//             Title: newName.replace(/^[ \t]+|[ \t]+$/gm, ""), // Remove leading and trailing spaces
//             NewName: setNewName.replace(/^[ \t]+|[ \t]+$/gm, ""),
//             libraryName: this.props.libraryName,
//             libraryID: this.props.libraryID,
//             FileID: this.props.selectedRow.ID,
//             FileRef0: this.props.selectedRow.FileRef.toString(),
//             UniqueId0: this.props.selectedRow.UniqueId.toString(),
//           });
//         this.setState({
//           showDialog: false,
//           isInputEmpty: false,
//           newName: "",
//           showNotification: true,
//         });
//       } catch (error) {
//         console.log("Error renaming file:", error);
//       }
//     }
//   }

//   private closeDialog() {
//     this.props.unMountDialog();
//   }

//   public render(): React.ReactElement<IRenameFileProps> {
//     const {
//       newName,
//       showDialog,
//       showNotification,
//       isExceedingLimit,
//       isInputEmpty,
//     } = this.state;

//     return (
//       <>
//         <Dialog
//           hidden={!showDialog}
//           onDismiss={this.closeDialog}
//           dialogContentProps={{
//             type: DialogType.normal,
//             title: "Rename File",
//           }}
//           modalProps={{
//             isBlocking: true,
//             styles: { main: { maxWidth: 450 } },
//           }}
//         >
//           <TextField
//             label="New Name"
//             value={newName}
//             onChange={this.handleInputChange}
//           />
//           {isExceedingLimit ? (
//             <div style={{ color: "red", marginTop: "8px" }}>
//               limit exceeded (200 characters maximum)
//             </div>
//           ) : (
//             isInputEmpty && (
//               <div style={{ color: "red", marginTop: "8px" }}>
//                 Please enter a new name before saving
//               </div>
//             )
//           )}
//           <DialogFooter>
//             <PrimaryButton onClick={this.handleRenameClick} text="Rename" />
//             <DefaultButton onClick={this.closeDialog} text="Cancel" />
//           </DialogFooter>
//         </Dialog>

//         <Dialog
//           hidden={!showNotification}
//           onDismiss={this.closeDialog}
//           dialogContentProps={{
//             type: DialogType.normal,
//             title: "New name has been saved",
//           }}
//           modalProps={{
//             isBlocking: true,
//             styles: { main: { width: 600 } },
//           }}
//         >
//           <div>
//             <i aria-hidden="true" style={{ fontSize: "20px" }}>
//               It may take a few minutes to change the file name.
//             </i>
//           </div>
//           <DialogFooter styles={{ actionsRight: { justifyContent: "center" } }}>
//             <PrimaryButton onClick={this.closeDialog} text="ok" />
//           </DialogFooter>
//         </Dialog>
//       </>
//     );
//   }
// }
