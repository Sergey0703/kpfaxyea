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
 * Interface for component state - UPDATED with skipped support
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
  fileSearchResults: { [rowIndex: number]: 'found' | 'not-found' | 'searching' | 'skipped' }; // UPDATED: Added 'skipped'
  searchProgress: ISearchProgress; // UPDATED: Enhanced progress tracking
  
  // NEW: Rename state with skipped support
  isRenaming: boolean;
  renameProgress?: {
    current: number;
    total: number;
    fileName: string;
    success: number;
    errors: number;
    skipped: number; // NEW: Track skipped files
  };
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
 * Interface for file search result - UPDATED with skipped support
 */
export interface IFileSearchResult {
  rowIndex: number;
  fileName: string;
  searchStatus: 'found' | 'not-found' | 'searching' | 'skipped'; // UPDATED: Added 'skipped'
  foundPath?: string;
  searchTime?: number;
  skipReason?: string; // NEW: Reason why file was skipped
}

/**
 * NEW: Interface for rename operation result
 */
export interface IRenameOperationResult {
  success: number;
  errors: number;
  skipped: number; // NEW: Number of skipped files
  errorDetails: string[];
  skippedDetails: string[]; // NEW: Details of skipped files
}

/**
 * NEW: Interface for file rename status
 */
export type FileRenameStatus = 'renaming' | 'renamed' | 'error' | 'skipped';

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

/**
 * NEW: Helper functions for rename progress
 */
export class RenameProgressHelper {
  
  /**
   * Create initial rename progress
   */
  public static createInitialProgress(totalFiles: number): {
    current: number;
    total: number;
    fileName: string;
    success: number;
    errors: number;
    skipped: number;
  } {
    return {
      current: 0,
      total: totalFiles,
      fileName: '',
      success: 0,
      errors: 0,
      skipped: 0
    };
  }
  
  /**
   * Update progress with new file
   */
  public static updateProgress(
    currentProgress: {
      current: number;
      total: number;
      fileName: string;
      success: number;
      errors: number;
      skipped: number;
    },
    fileName: string,
    status: 'success' | 'error' | 'skipped'
  ): {
    current: number;
    total: number;
    fileName: string;
    success: number;
    errors: number;
    skipped: number;
  } {
    const newProgress = {
      ...currentProgress,
      current: currentProgress.current + 1,
      fileName
    };
    
    switch (status) {
      case 'success':
        newProgress.success++;
        break;
      case 'error':
        newProgress.errors++;
        break;
      case 'skipped':
        newProgress.skipped++;
        break;
    }
    
    return newProgress;
  }
  
  /**
   * Calculate completion percentage
   */
  public static getCompletionPercentage(progress: {
    current: number;
    total: number;
    success: number;
    errors: number;
    skipped: number;
  }): number {
    if (progress.total === 0) return 0;
    return Math.round((progress.current / progress.total) * 100);
  }
  
  /**
   * Get summary message
   */
  public static getSummaryMessage(progress: {
    current: number;
    total: number;
    success: number;
    errors: number;
    skipped: number;
  }): string {
    const messages: string[] = [];
    
    if (progress.success > 0) {
      messages.push(`✅ ${progress.success} renamed`);
    }
    
    if (progress.skipped > 0) {
      messages.push(`⏭️ ${progress.skipped} skipped`);
    }
    
    if (progress.errors > 0) {
      messages.push(`❌ ${progress.errors} failed`);
    }
    
    return messages.join(', ') || 'No operations completed';
  }
}

/**
 * NEW: Enum for different types of file status icons
 */
export enum FileStatusIcon {
  SEARCHING = '🔍',
  FOUND = '✅',
  NOT_FOUND = '❌',
  SKIPPED = '⏭️',
  RENAMING = '🔄',
  RENAMED = '✅',
  ERROR = '❌'
}

/**
 * NEW: Helper functions for file status
 */
export class FileStatusHelper {
  
  /**
   * Get icon for file search status
   */
  public static getSearchIcon(status: 'found' | 'not-found' | 'searching' | 'skipped'): string {
    switch (status) {
      case 'searching':
        return FileStatusIcon.SEARCHING;
      case 'found':
        return FileStatusIcon.FOUND;
      case 'not-found':
        return FileStatusIcon.NOT_FOUND;
      case 'skipped':
        return FileStatusIcon.SKIPPED;
      default:
        return '';
    }
  }
  
  /**
   * Get icon for file rename status
   */
  public static getRenameIcon(status: FileRenameStatus): string {
    switch (status) {
      case 'renaming':
        return FileStatusIcon.RENAMING;
      case 'renamed':
        return FileStatusIcon.RENAMED;
      case 'skipped':
        return FileStatusIcon.SKIPPED;
      case 'error':
        return FileStatusIcon.ERROR;
      default:
        return '';
    }
  }
  
  /**
   * Get tooltip text for status
   */
  public static getTooltipText(status: 'found' | 'not-found' | 'searching' | 'skipped' | FileRenameStatus): string {
    switch (status) {
      case 'searching':
        return 'Searching for file...';
      case 'found':
        return 'File found in SharePoint';
      case 'not-found':
        return 'File not found in SharePoint';
      case 'skipped':
        return 'Skipped - target file already exists';
      case 'renaming':
        return 'Renaming file...';
      case 'renamed':
        return 'File renamed successfully';
      case 'error':
        return 'Error occurred during rename';
      default:
        return '';
    }
  }
}

/**
 * NEW: Interface for batch rename operations
 */
export interface IBatchRenameOperation {
  id: string;
  files: Array<{
    rowIndex: number;
    originalName: string;
    newName: string;
    staffID: string;
    status: FileRenameStatus;
  }>;
  startTime: Date;
  endTime?: Date;
  summary: {
    total: number;
    success: number;
    errors: number;
    skipped: number;
  };
}

/**
 * NEW: Interface for rename operation statistics
 */
export interface IRenameStatistics {
  totalOperations: number;
  successfulRenames: number;
  failedRenames: number;
  skippedRenames: number;
  averageTimePerFile: number;
  totalTimeElapsed: number;
  successRate: number;
}

/**
 * NEW: Type definitions for callback functions
 */
export type SearchProgressCallback = (progress: ISearchProgress) => void;
export type RenameProgressCallback = (progress: {
  current: number;
  total: number;
  fileName: string;
  success: number;
  errors: number;
  skipped: number;
}) => void;
export type FileStatusCallback = (rowIndex: number, status: FileRenameStatus) => void;
export type SearchResultCallback = (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void;