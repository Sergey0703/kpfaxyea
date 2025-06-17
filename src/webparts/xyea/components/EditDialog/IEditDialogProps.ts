// src/webparts/xyea/components/EditDialog/IEditDialogProps.ts

import { IConvertFile } from '../../models';

export interface IEditDialogProps {
  isOpen: boolean;
  isEditMode: boolean;
  item?: IConvertFile | undefined; // Use proper type instead of any
  title: string;
  loading?: boolean;
  onSave: (title: string) => Promise<void>;
  onCancel: () => void;
}