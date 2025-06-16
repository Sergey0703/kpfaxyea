// src/webparts/xyea/components/ConvertFilesTable/IConvertFilesTableProps.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IConvertFile } from '../../models';

export interface IConvertFilesTableProps {
  context: WebPartContext;
  convertFiles: IConvertFile[];
  loading: boolean;
  onAdd: () => void;
  onEdit: (item: IConvertFile) => void;
  onDelete: (id: number) => void;
  onRowClick: (convertFileId: number) => void;
  expandedRows: number[];
}