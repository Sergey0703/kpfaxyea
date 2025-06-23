// src/webparts/xyea/components/RenameFilesManagement/types/RenameFilesTypes.ts

import { IExcelFile, IExcelSheet } from '../../../interfaces/ExcelInterfaces';

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
 * NEW: Search stages enum
 */
export enum SearchStage {
  IDLE = 'idle',
  ANALYZING_DIRECTORIES = 'analyzing_directories',
  CHECKING_EXISTENCE = 'checking_existence', 
  SEARCHING_FILES = 'searching_files',
  COMPLETED = 'completed',
  CANCELLED = 'cancelled',
  ERROR = 'error'
}

/**
 * NEW: Interface for search stage information
 */
export interface ISearchStageInfo {
  stage: SearchStage;
  title: string;
  description: string;
  progressMin: number; // minimum progress percentage for this stage
  progressMax: number; // maximum progress percentage for this stage
}

/**
 * NEW: Enhanced search progress interface with stages
 */
export interface ISearchProgress {
  // Stage information
  currentStage: SearchStage;
  stageProgress: number; // Progress within current stage (0-100)
  overallProgress: number; // Overall progress across all stages (0-100)
  
  // Current operation details
  currentRow: number;
  totalRows: number;
  currentFileName: string;
  currentDirectory?: string; // NEW: Current directory being processed
  
  // Stage-specific stats
  directoriesAnalyzed?: number;
  totalDirectories?: number;
  directoriesChecked?: number;
  existingDirectories?: number;
  filesSearched?: number;
  filesFound?: number;
  
  // Timing information
  stageStartTime?: Date;
  estimatedTimeRemaining?: number; // in seconds
  
  // Error information
  errors?: string[];
  warnings?: string[];
  
  // NEW: Search plan reference
  searchPlan?: ISearchPlan;
}

/**
 * NEW: Directory analysis result
 */
export interface IDirectoryAnalysis {
  directoryPath: string;
  normalizedPath: string;
  exists: boolean;
  fileCount: number;
  rowIndexes: number[];
  fullSharePointPath: string;
  hasValidPath: boolean;
}

/**
 * NEW: Search plan interface
 */
export interface ISearchPlan {
  totalRows: number;
  validRows: number;
  invalidRows: number;
  totalDirectories: number;
  existingDirectories: number;
  missingDirectories: number;
  directoryGroups: IDirectoryAnalysis[];
  estimatedDuration: number; // in seconds
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
  
  // File searching and renaming - UPDATED with new progress interface
  searchingFiles: boolean;
  fileSearchResults: { [rowIndex: number]: 'found' | 'not-found' | 'searching' };
  searchProgress: ISearchProgress; // UPDATED: Enhanced progress tracking
  searchPlan?: ISearchPlan; // NEW: Search plan for optimization
}

/**
 * Interface for component props
 */
export interface IRenameFilesManagementProps {
  context: any; // WebPartContext
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

/**
 * Interface for upload progress tracking
 */
export interface IUploadProgress {
  stage: 'idle' | 'uploading' | 'parsing' | 'complete' | 'error';
  progress: number; // 0-100
  message: string;
}

/**
 * Interface for file search result
 */
export interface IFileSearchResult {
  rowIndex: number;
  fileName: string;
  searchStatus: 'found' | 'not-found' | 'searching';
  foundPath?: string;
  searchTime?: number;
}

/**
 * NEW: Constants for search stages
 */
export const SEARCH_STAGES: { [key in SearchStage]: ISearchStageInfo } = {
  [SearchStage.IDLE]: {
    stage: SearchStage.IDLE,
    title: 'Ready',
    description: 'Ready to start search',
    progressMin: 0,
    progressMax: 0
  },
  [SearchStage.ANALYZING_DIRECTORIES]: {
    stage: SearchStage.ANALYZING_DIRECTORIES,
    title: 'Analyzing Directories',
    description: 'Extracting and analyzing directory structure from your data...',
    progressMin: 0,
    progressMax: 50 // NEW: Changed to 50% for two-button approach
  },
  [SearchStage.CHECKING_EXISTENCE]: {
    stage: SearchStage.CHECKING_EXISTENCE,
    title: 'Checking Directory Existence', 
    description: 'Verifying which directories exist in SharePoint...',
    progressMin: 50, // NEW: Changed to start at 50%
    progressMax: 100 // NEW: Changed to end at 100% for analysis phase
  },
  [SearchStage.SEARCHING_FILES]: {
    stage: SearchStage.SEARCHING_FILES,
    title: 'Searching Files',
    description: 'Looking for files in existing directories...',
    progressMin: 0, // NEW: Reset to 0% for separate search phase
    progressMax: 100 // NEW: Full 100% for file search
  },
  [SearchStage.COMPLETED]: {
    stage: SearchStage.COMPLETED,
    title: 'Completed',
    description: 'Operation completed successfully',
    progressMin: 100,
    progressMax: 100
  },
  [SearchStage.CANCELLED]: {
    stage: SearchStage.CANCELLED,
    title: 'Cancelled',
    description: 'Operation was cancelled',
    progressMin: 0,
    progressMax: 0
  },
  [SearchStage.ERROR]: {
    stage: SearchStage.ERROR,
    title: 'Error',
    description: 'An error occurred during operation',
    progressMin: 0,
    progressMax: 0
  }
};

/**
 * NEW: Helper functions for search progress
 */
export class SearchProgressHelper {
  
  /**
   * Calculate overall progress based on stage and stage progress
   */
  public static calculateOverallProgress(stage: SearchStage, stageProgress: number): number {
    const stageInfo = SEARCH_STAGES[stage];
    const stageRange = stageInfo.progressMax - stageInfo.progressMin;
    const stageContribution = (stageProgress / 100) * stageRange;
    return Math.min(100, Math.max(0, stageInfo.progressMin + stageContribution));
  }
  
  /**
   * Create initial search progress
   */
  public static createInitialProgress(): ISearchProgress {
    return {
      currentStage: SearchStage.IDLE,
      stageProgress: 0,
      overallProgress: 0,
      currentRow: 0,
      totalRows: 0,
      currentFileName: '',
      stageStartTime: new Date(),
      errors: [],
      warnings: []
    };
  }
  
  /**
   * Update progress for a specific stage
   */
  public static updateStageProgress(
    currentProgress: ISearchProgress,
    newStageProgress: number,
    updates?: Partial<ISearchProgress>
  ): ISearchProgress {
    const overallProgress = SearchProgressHelper.calculateOverallProgress(
      currentProgress.currentStage, 
      newStageProgress
    );
    
    return {
      ...currentProgress,
      stageProgress: newStageProgress,
      overallProgress,
      ...updates
    };
  }
  
  /**
   * Transition to next stage
   */
  public static transitionToStage(
    currentProgress: ISearchProgress,
    newStage: SearchStage,
    updates?: Partial<ISearchProgress>
  ): ISearchProgress {
    const stageInfo = SEARCH_STAGES[newStage];
    
    return {
      ...currentProgress,
      currentStage: newStage,
      stageProgress: 0,
      overallProgress: stageInfo.progressMin,
      stageStartTime: new Date(),
      ...updates
    };
  }
}