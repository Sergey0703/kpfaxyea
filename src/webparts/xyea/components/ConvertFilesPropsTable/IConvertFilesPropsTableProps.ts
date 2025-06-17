// src/webparts/xyea/components/ConvertFilesPropsTable/IConvertFilesPropsTableProps.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IConvertFileProps } from '../../models';
import { IExcelImportData } from '../ExcelImportButton/ExcelImportButton';
import { ISelectedFiles } from '../ConvertFilesTable/IConvertFilesTableProps';

export interface IConvertFilesPropsTableProps {
  context: WebPartContext;
  convertFileId: number;
  convertFileTitle: string;
  items: IConvertFileProps[];
  allItems: IConvertFileProps[]; // Для вычисления Priority
  loading: boolean;
  onAdd: (convertFileId: number) => void;
  onEdit: (item: IConvertFileProps) => void;
  onDelete: (id: number) => void;
  onMoveUp: (id: number) => void;
  onMoveDown: (id: number) => void;
  onToggleDeleted: (id: number, isDeleted: boolean) => void;
  onImportFromExcel?: (convertFileId: number, data: IExcelImportData[]) => Promise<void>; // New prop for Excel import
  selectedFiles?: ISelectedFiles; // NEW: Add selected files prop for validation
}