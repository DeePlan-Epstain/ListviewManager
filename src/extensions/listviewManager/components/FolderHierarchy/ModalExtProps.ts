export interface ModalExtProps {
  sp: any;
  context: any;
  finalPath: any;
  currUser: any;
  modalInterface: "Review" | "Approval";
  unMountDialog: () => void;
}