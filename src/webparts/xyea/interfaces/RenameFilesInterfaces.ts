// src/webparts/xyea/interfaces/RenameFilesInterfaces.ts

import { IExcelFile, IExcelSheet } from './ExcelInterfaces';

/**
 * Interface for WebPart context - ADDED: Specific interface instead of any
 */
export interface IWebPartContext {
  pageContext: {
    user: {
      displayName: string;
    };
    web: {
      absoluteUrl: string;
      serverRelativeUrl: string;
    };
  };
  aadTokenProviderFactory?: unknown;
}

/**
 * Interface for a custom column that can be added to the table
 */
export interface ICustomColumn {
  id: string; // unique identifier
  name: string; // display name
  isEditable: boolean;
  defaultValue?: string;
  width?: number;
}

/**
 * Interface for column order and configuration
 */
export interface IColumnConfiguration {
  id: string; // matches excel column or custom column id
  name: string; // display name
  originalIndex?: number; // original index from Excel file (if from Excel)
  currentIndex: number; // current display position
  isVisible: boolean;
  isCustom: boolean; // true if added by user, false if from Excel
  isEditable: boolean;
  width?: number;
  dataType?: 'text' | 'number' | 'date' | 'boolean';
}

/**
 * Interface for table cell data
 */
export interface ITableCell {
  value: string | number | boolean | Date | undefined;
  isEdited: boolean; // track if user modified this cell
  originalValue?: string | number | boolean | Date | undefined;
  columnId: string;
  rowIndex: number;
}

/**
 * Interface for table row data with custom columns support
 */
export interface IRenameTableRow {
  rowIndex: number;
  cells: { [columnId: string]: ITableCell };
  isVisible: boolean; // for filtering
  isEdited: boolean; // true if any cell in row was edited
}

/**
 * Interface for the main rename files data structure
 */
export interface IRenameFilesData {
  originalFile: IExcelFile | undefined;
  currentSheet: IExcelSheet | undefined;
  columns: IColumnConfiguration[];
  rows: IRenameTableRow[];
  customColumns: ICustomColumn[];
  totalRows: number;
  editedCellsCount: number;
}

/**
 * Interface for column reorder operation
 */
export interface IColumnReorderOperation {
  columnId: string;
  fromIndex: number;
  toIndex: number;
}

/**
 * Interface for cell edit operation
 */
export interface ICellEditOperation {
  columnId: string;
  rowIndex: number;
  oldValue: string | number | boolean | Date | undefined;
  newValue: string | number | boolean | Date | undefined;
  timestamp: Date;
}

/**
 * Interface for export configuration for renamed files
 */
export interface IRenameFilesExportConfig {
  fileName: string;
  includeOnlyEditedRows: boolean;
  includeCustomColumns: boolean;
  includeOriginalColumns: boolean;
  columnOrder: string[]; // array of column IDs in export order
  fileFormat: 'xlsx' | 'csv';
}

/**
 * Interface for component state
 */
export interface IRenameFilesState {
  // File and data
  data: IRenameFilesData;
  
  // UI state
  loading: boolean;
  error: string | undefined;
  uploadProgress: {
    stage: 'idle' | 'uploading' | 'parsing' | 'complete' | 'error';
    progress: number;
    message: string;
  };
  
  // Table state
  selectedCells: { [key: string]: boolean }; // key format: "columnId_rowIndex"
  editingCell: { columnId: string; rowIndex: number } | undefined;
  
  // Column management
  showColumnManager: boolean;
  draggedColumn: string | undefined;
  
  // Export state
  showExportDialog: boolean;
  exportConfig: IRenameFilesExportConfig;
  isExporting: boolean;
  
  // SharePoint folder selection
  selectedFolder: ISharePointFolder | undefined;
  showFolderDialog: boolean;
  availableFolders: ISharePointFolder[];
  loadingFolders: boolean;
}

/**
 * SharePoint folder interface
 */
export interface ISharePointFolder {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  TimeCreated: string;
  TimeLastModified: string;
}

/**
 * Interface for component props
 */
export interface IRenameFilesManagementProps {
  context: IWebPartContext; // FIXED: Changed from any to specific interface
  userDisplayName: string;
}

/**
 * Helper type for column operations
 */
export type ColumnOperation = 
  | { type: 'ADD_CUSTOM'; column: ICustomColumn }
  | { type: 'REMOVE_CUSTOM'; columnId: string }
  | { type: 'REORDER'; operation: IColumnReorderOperation }
  | { type: 'TOGGLE_VISIBILITY'; columnId: string }
  | { type: 'RENAME'; columnId: string; newName: string };

/**
 * Helper type for cell operations
 */
export type CellOperation = 
  | { type: 'EDIT'; operation: ICellEditOperation }
  | { type: 'CLEAR'; columnId: string; rowIndex: number }
  | { type: 'BULK_EDIT'; operations: ICellEditOperation[] };

/**
 * Utility interface for column statistics
 */
export interface IColumnStats {
  columnId: string;
  totalCells: number;
  editedCells: number;
  emptyCells: number;
  uniqueValues: number;
  mostCommonValue?: string | number;
}

/**
 * Interface for validation results
 */
export interface IValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  affectedRows?: number[];
}