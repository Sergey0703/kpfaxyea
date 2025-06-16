// src/webparts/xyea/components/EditDialog/IEditPropsDialogProps.ts

import { IConvertFileProps } from '../../models';

export interface IEditPropsDialogProps {
  isOpen: boolean;
  isEditMode: boolean;
  convertFileId: number;
  item?: IConvertFileProps | null;
  title: string;
  loading?: boolean;
  onSave: (convertFileId: number, title: string, prop: string, prop2: string) => Promise<void>;
  onCancel: () => void;
}