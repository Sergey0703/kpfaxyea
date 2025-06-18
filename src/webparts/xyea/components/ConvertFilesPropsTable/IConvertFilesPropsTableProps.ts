// src/webparts/xyea/components/ConvertFilesPropsTable/IConvertFilesPropsTableProps.ts - Updated with ConvertType support

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IConvertFileProps } from '../../models';
import { IConvertType } from '../../models/IConvertType';
import { IExcelImportData } from '../ExcelImportButton/ExcelImportButton';
import { ISelectedFiles } from '../ConvertFilesTable/IConvertFilesTableProps';

export interface IConvertFilesPropsTableProps {
  context: WebPartContext;
  convertFileId: number;
  convertFileTitle: string;
  items: IConvertFileProps[];
  allItems: IConvertFileProps[]; // Для вычисления Priority
  loading: boolean;
  convertTypes: IConvertType[]; // NEW: List of available convert types
  onAdd: (convertFileId: number) => void;
  onEdit: (item: IConvertFileProps) => void;
  onDelete: (id: number) => void;
  onMoveUp: (id: number) => void;
  onMoveDown: (id: number) => void;
  onToggleDeleted: (id: number, isDeleted: boolean) => void;
  onImportFromExcel?: (convertFileId: number, data: IExcelImportData[]) => Promise<void>;
  selectedFiles?: ISelectedFiles;
}