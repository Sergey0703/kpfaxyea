// src/webparts/xyea/components/ConvertFilesTable/IConvertFilesTableProps.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IConvertFile } from '../../models';

// Define the type for selected files
export interface ISelectedFiles {
  [convertFileId: number]: {
    export?: File;
    import?: File;
  };
}

export interface IConvertFilesTableProps {
  context: WebPartContext;
  convertFiles: IConvertFile[];
  loading: boolean;
  onAdd: () => void;
  onEdit: (item: IConvertFile) => void;
  onDelete: (id: number) => void;
  onRowClick: (convertFileId: number) => void;
  expandedRows: number[];
  selectedFiles?: ISelectedFiles; // Selected files from parent
  onSelectedFilesChange?: (selectedFiles: ISelectedFiles) => void; // NEW: Callback to update parent
}