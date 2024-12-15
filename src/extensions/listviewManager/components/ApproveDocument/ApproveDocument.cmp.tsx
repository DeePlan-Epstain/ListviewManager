import * as React from "react";
import { useState, useEffect } from "react";
import modalStyles from "../../styles/modalStyles.module.scss";
import styles from "./ApproveDocument.module.scss";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { TextField, Button, FormControlLabel, Checkbox } from "@mui/material";
import {
  PeoplePicker,
  PrincipalType,
  IPeoplePickerUserItem,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import axios from "axios";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import {
  DragDropContext,
  DraggableProvided,
  DraggableStateSnapshot,
  Droppable,
  DroppableProvided,
} from "react-beautiful-dnd";
import { Draggable } from "react-beautiful-dnd";
import Loader from "../Loader/Loader.cmp";
import { makeId } from "../../service/util.service";
import { DocumentProps, SelectedFile } from "../../models/global.model";

export interface ApproveDocumentProps {
  sp: SPFI;
  context: any;
  selectedRow: SelectedFile;
  currUser: ISiteUserInfo;
  modalInterface: "Review" | "Approval";
  unMountDialog: () => void;
}

export default function ReviewDocument({
  sp,
  selectedRow,
  context,
  currUser,
  modalInterface,
  unMountDialog,
}: ApproveDocumentProps) {
  const [loadingProps, setLoadingProps] = useState<{
    loading: boolean;
    msg: string;


  }>({
    loading: false,
    msg: "",
  });
  const [errorMsg, setErrorMsg] = useState<string>("");
  const [isAskCancelRunningFlow, setisAskCancelRunningFlow] =
    useState<boolean>(false);
  const [isHighlightFormErrors, setIsHighlightFormErrors] =
    useState<boolean>(false);
  const [selectedFile, setSelectedFile] = useState<DocumentProps>({
    Id: 0,
    Title: "",
    libraryId: "",
    latestApprovers:
      (modalInterface === "Approval"
        ? selectedRow.latestApprovers
        : selectedRow.latestReviewers) || [],
    approveStatus: selectedRow.approveStatus,
    comment: "",
    ApprovalStatus: "",
    Initiator: 0,
    ServerRelativeUrl: "",
  });

  useEffect((): void => {
    _getFile();
  }, [modalInterface]);

  const _getFile = async (isCancelFlow?: boolean): Promise<void> => {
    try {
      const file: any = await sp.web
        .getFileByServerRelativePath(selectedRow.FileRef)
        .expand("Properties")();
      const libraryId = file.Properties["vti_x005f_listid"].slice(1, -1);
      const fileListId = file.Properties["vti_x005f_doclibrowid"];
      const fileCheck: any = await sp.web.lists
        .getById(libraryId)
        .items.getById(fileListId)();

      setSelectedFile(
        (prev) =>
        (prev = {
          ...prev,
          Id: fileListId,
          Title: file.Name,
          libraryId,
          ApprovalStatus: file.Properties?.ApprovalStatus,
          ServerRelativeUrl: file.ServerRelativeUrl,
          approveStatus: isCancelFlow
            ? "Not Started"
            : fileCheck.ApprovalStatus,
        })
      );

      if (fileCheck.ApprovalStatus === "In progress") {
        setisAskCancelRunningFlow(true);
      }
    } catch (err) {
      console.log("_getFile", err);
    }
  };

  const onConfirm = async (): Promise<void> => {
    try {
      const spPortal= spfi("https://epstin100.sharepoint.com/sites/EpsteinPortal/").using(SPFx(context))
      

      // Resolve user IDs for all approvers
      const approvesIdPromises = selectedFile.latestApprovers.map(
        async (approver: any) => {
          try {
            const user = await spPortal.web.ensureUser(approver.loginName);
            return user.data.Id;
          } catch (error) {
            console.log(
              `Error resolving user ID for ${approver.loginName}: ${error}`
            );
            throw error;
          }
        }
      );

      // Wait for all promises to resolve
      const approvesId = await Promise.all(approvesIdPromises);

      if (approvesId.length === 0) {
        setErrorMsg("אנא בחר לפחות מאשר אחד");
        return;
      } else {
    const spPortal= spfi("https://epstin100.sharepoint.com/sites/EpsteinPortal/").using(SPFx(context))
    const currUser = await spPortal.web.currentUser();
    const fullUrl = context.pageContext.web.absoluteUrl; // כתובת האתר המלאה
    const baseUrl = fullUrl.split("/").slice(0, 5).join("/"); // חיתוך החלק הרצוי

        // Add item to the SharePoint list with resolved user IDs
        await spPortal.web.lists
          .getById("968d78ad-3e30-4657-baf7-8ef5f5c6c40f")
          .items.add({
            Title: currUser.Title,
            DocID: selectedFile.Id,
            ApproversId: approvesId,
            NumOfApprovals: approvesId.length,
            libraryId: selectedFile.libraryId,
            docName: selectedFile.Title,
            userId: currUser.Id,
            userEmail: currUser.Email,
            TaskComment: selectedFile.comment,
            ApproveStatus: selectedFile.approveStatus,
            InitiatorId: currUser.Id,
            DocServerRelativeUrl: selectedFile.ServerRelativeUrl,
            baseUrl:baseUrl
            
          });

        // Optionally, you can add a message to indicate successful addition
        console.log("Item added successfully.");
        unMountDialog();
      }
    } catch (error) {
      console.log(error);
    }
  };

  const handleApproversSelection = (
    approvers: IPeoplePickerUserItem[]
  ): void => {
    approvers.forEach(
      (A: any) => (A.sip = A.loginName.replace("i:0#.f|membership|", ""))
    );
    setSelectedFile({ ...selectedFile, latestApprovers: approvers });
  };

  const onDragEnd = (Res: any): void => {
    const { destination, source } = Res;
    if (
      !destination ||
      (destination.draggableId === source.droppableId &&
        destination.index === source.index)
    )
      return;
    const newApprovers = JSON.parse(
      JSON.stringify(selectedFile.latestApprovers)
    );
    const MovedCase = newApprovers[source.index];
    newApprovers.splice(source.index, 1);
    newApprovers.splice(destination.index, 0, MovedCase);
    setSelectedFile({ ...selectedFile, latestApprovers: newApprovers });
  };

  const cancelRunningFlow = async (): Promise<void> => {
    try {
      setLoadingProps({ loading: true, msg: "Cancelling approval flow..." });

      // Update the ApprovalStatus of the selected file
      await sp.web.lists
        .getById(selectedFile.libraryId)
        .items.getById(selectedFile.Id)
        .update({ ApprovalStatus: "Not started" });

      // Get the DocId of the selected file
      const docId = selectedFile.Id;

      // Query the list with ID 'fcca197d-c6ec-4cd3-9f0b-4316985fad65' to find items matching the DocId
      const itemsToUpdate = await sp.web.lists
        .getById("fcca197d-c6ec-4cd3-9f0b-4316985fad65")
        .items.select("Id")
        .filter(`DocID eq '${docId}'`)();

      // Update each item found in the query
      const updatePromises = itemsToUpdate.map(async (item: any) => {
        await sp.web.lists
          .getById("fcca197d-c6ec-4cd3-9f0b-4316985fad65")
          .items.getById(item.Id)
          .update({ ApproveStatus: "Cancelled" });
      });

      // Wait for all updates to complete
      await Promise.all(updatePromises);

      await _getFile(true);
      setLoadingProps({ loading: false, msg: "" });
      setisAskCancelRunningFlow(false);
    } catch (error) {
      setLoadingProps({ loading: false, msg: "" });
      setisAskCancelRunningFlow(false);
      setErrorMsg(
        "The file is currently open elsewhere, and so is locked for editing."
      );
      console.log(error);
    }
  };

  const onClose = (): void => {
    if (selectedFile.approveStatus === "In progress") unMountDialog();
    else setErrorMsg("");
  };

  return (
    <DragDropContext onDragEnd={onDragEnd}>
      <div
        className={modalStyles.modalScreen}
        onClick={() => (loadingProps.loading ? null : unMountDialog())}
      >
        <div
          className={modalStyles.modal}
          style={{
            width: 680,
            minHeight: errorMsg ? "unset" : 280,
            direction: "rtl"
          }}
          onClick={(ev: any) => ev.stopPropagation()}
        >
          {/* Loader */}
          {loadingProps.loading && <Loader msg={loadingProps.msg} />}
          {!loadingProps.loading && (
            <>
              <div className={modalStyles.modalHeader}>
                <span>
                  {modalInterface === "Approval" ? "אישור" : modalInterface}{" "}
                  מסמך
                </span>
              </div>

              {!errorMsg && (
                <>
                  <div className={styles.reviewDocumentModalContentContainer}>
                    <div
                      className={styles.reviewDocumentModalContent}
                      style={{ width: "calc(100% - 200px)" }}
                    >
                      <div dir="rtl">

                        <PeoplePicker
                          context={context}
                          titleText={
                            modalInterface === "Approval"
                              ? "בחר מאשרים *"
                              : "בחר מבקרים *"
                          }
                          personSelectionLimit={5}
                          showtooltip
                          // required
                          defaultSelectedUsers={
                            selectedFile.latestApprovers?.map(
                              (A: any) => A.sip
                            ) || []
                          }
                          onChange={(Users: IPeoplePickerUserItem[]) =>
                            handleApproversSelection(Users)
                          }
                          principalTypes={[PrincipalType.User]}
                          disabled={Boolean(errorMsg || isAskCancelRunningFlow)}
                          resolveDelay={1000}
                        />

                        {isHighlightFormErrors &&
                          !selectedFile.latestApprovers?.length && (
                            <span style={{ color: "red" }}>
                              דרושים{" "}
                              {modalInterface === "Approval"
                                ? "מאשרים"
                                : "מבקרים"}{" "}
                              להמשיך...
                            </span>
                          )}
                      </div>

                      <div>
                        <span style={{ fontWeight: 600, letterSpacing: ".02rem", padding: "5px 0px" }}>תגובה:</span>
                        <TextField
                          id="document-comment"
                          // label="תגובה"
                          value={selectedFile.comment}
                          sx={{ width: "100%" }}
                          disabled={Boolean(errorMsg || isAskCancelRunningFlow)}
                          multiline
                          rows={4}
                          InputProps={{
                            style: { direction: "rtl" },
                          }}
                          onChange={(
                            ev: React.ChangeEvent<HTMLInputElement | HTMLTextAreaElement>
                          ) =>
                            setSelectedFile({
                              ...selectedFile,
                              comment: ev.target.value,
                            })
                          }
                        />
                      </div>


                    </div>

                    <Droppable droppableId="approversList" direction="vertical">
                      {(provided: DroppableProvided) => (
                        <div
                          className={styles.approversOrderContainer}
                          ref={provided.innerRef}
                          {...provided.droppableProps}
                        >
                          <span>
                            סדר
                            {" "}
                            {modalInterface === "Approval"
                              ? "מאשרים"
                              : "מבקרים"}

                          </span>
                          <ul
                            id="approversOrder"
                            className={styles.approversOrderList}
                          >
                            {selectedFile.latestApprovers.map(
                              (A: any, idx: number) => (
                                <Draggable
                                  draggableId={`${A.id}`}
                                  key={`${A.id}`}
                                  index={idx}
                                >
                                  {(
                                    provided: DraggableProvided,
                                    snapshot: DraggableStateSnapshot
                                  ) => (
                                    <li
                                      {...provided.draggableProps}
                                      {...provided.dragHandleProps}
                                      ref={provided.innerRef}
                                      style={
                                        snapshot.isDragging
                                          ? {
                                            ...provided.draggableProps.style,
                                            color: "white",
                                            backgroundColor:
                                              "rgb(0, 30, 255, 0.6)",
                                          }
                                          : {
                                            ...provided.draggableProps.style,
                                          }
                                      }
                                      className={styles.approversOrderItem}
                                    >
                                      <span>{A.title || A.text}</span>
                                      <div
                                        className={
                                          styles.approversOrderItemBackground
                                        }
                                      ></div>
                                    </li>
                                  )}
                                </Draggable>
                              )
                            )}
                            {provided.placeholder}
                          </ul>
                        </div>
                      )}
                    </Droppable>
                  </div>

                  <div className={modalStyles.modalFooter}>
                    <Button
                      disabled={Boolean(errorMsg || isAskCancelRunningFlow)}
                      onClick={unMountDialog}
                      color="error"
                    >
                      ביטול
                    </Button>
                    <Button
                      disabled={Boolean(errorMsg || isAskCancelRunningFlow)}
                      onClick={onConfirm}
                    >
                      אישור
                    </Button>
                  </div>
                </>
              )}

              {isAskCancelRunningFlow && (
                <div className={modalStyles.modalDialogContainer}>
                  <span>
                   יש תהליך אישור פעיל על מסמך זה. האם ברצונך לבטל אותו?
                  </span>
                  <div className={modalStyles.modalDialogActions}>
                    <Button color="warning" onClick={cancelRunningFlow}>
                      כן
                    </Button>
                    <Button color="success" onClick={unMountDialog}>
                      לא
                    </Button>
                  </div>
                </div>
              )}

              {errorMsg && (
                <div className={modalStyles.modalDialogContainer}>
                  <span className={modalStyles.errMsg}>{errorMsg}</span>
                  <Button onClick={onClose}>סגור</Button>
                </div>
              )}
            </>
          )}
        </div>
      </div>
    </DragDropContext>
  );
}
