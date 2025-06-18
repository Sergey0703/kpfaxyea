// src/webparts/xyea/components/ConvertFilesPropsTable/IEditPropsDialogProps.ts - Updated with ConvertType support

import { IConvertFileProps } from '../../models';
import { IConvertType } from '../../models/IConvertType';

export interface IEditPropsDialogProps {
  isOpen: boolean;
  isEditMode: boolean;
  convertFileId: number;
  item?: IConvertFileProps | undefined;
  title: string;
  loading?: boolean;
  convertTypes: IConvertType[];
  onSave: (
    convertFileId: number, 
    title: string, 
    prop: string, 
    prop2: string,
    convertTypeId: number,
    convertType2Id: number
  ) => Promise<void>;
  onCancel: () => void;
}