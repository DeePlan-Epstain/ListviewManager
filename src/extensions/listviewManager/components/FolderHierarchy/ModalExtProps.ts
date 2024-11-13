export interface ModalExtProps {
  sp: any;
  context: any;
  selectedRow: any;
  currUser: any;
  modalInterface: "Review" | "Approval";
  unMountDialog: () => void;
}