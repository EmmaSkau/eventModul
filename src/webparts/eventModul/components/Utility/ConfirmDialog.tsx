import * as React from "react";
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
} from "@fluentui/react";

export interface IConfirmDialogProps {
  hidden: boolean;
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  onConfirm: () => void;
  onCancel: () => void;
  dialogType?: DialogType;
}

const ConfirmDialog: React.FC<IConfirmDialogProps> = (props) => {
  const {
    hidden,
    title,
    message,
    confirmText = "Bekræft",
    cancelText = "Annuller",
    onConfirm,
    onCancel,
    dialogType = DialogType.normal,
  } = props;

  return (
    <Dialog
      hidden={hidden}
      onDismiss={onCancel}
      dialogContentProps={{
        type: dialogType,
        title: title,
        subText: message,
      }}
      modalProps={{
        isBlocking: true,
      }}
    >
      <DialogFooter>
        <PrimaryButton onClick={onConfirm} text={confirmText} />
        <DefaultButton onClick={onCancel} text={cancelText} />
      </DialogFooter>
    </Dialog>
  );
};

export default ConfirmDialog;
