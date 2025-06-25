// src/webparts/xyea/components/RenameFilesManagement/RenameFilesManagement.tsx

import * as React from 'react';
import styles from './RenameFilesManagement.module.scss';
import { 
  IRenameFilesManagementProps, 
  IRenameFilesState,
  SearchStage,
  ISearchProgress,
  SearchProgressHelper,
  IRenameExportSettings,
  DirectoryStatus,
  FileSearchStatus,
  DirectoryStatusCallback,
  FileSearchResultCallback
} from './types/RenameFilesTypes';
import { ExcelFileProcessor } from './services/ExcelFileProcessor';
import { SharePointFolderService } from './services/SharePointFolderService';
import { FileSearchService } from './services/FileSearchService';
import { ColumnResizeHandler } from './handlers/ColumnResizeHandler';
import { CellEditingHandler } from './handlers/CellEditingHandler';
import { FolderSelectionDialog } from './components/FolderSelectionDialog';
import { RenameControlsPanel } from './components/RenameControlsPanel';
import { DataTableView } from './components/DataTableView';
import { ProgressIndicator } from './components/ProgressIndicator';
import { ExportControlsPanel } from './components/ExportControlsPanel';
import { ExcelExportService } from '../../services/ExcelExportService';

export default class RenameFilesManagement extends React.Component<IRenameFilesManagementProps, IRenameFilesState> {
  private fileInputRef: React.RefObject<HTMLInputElement>;
  private excelProcessor: ExcelFileProcessor;
  private folderService: SharePointFolderService;
  private fileSearchService: FileSearchService;
  private columnResizeHandler: ColumnResizeHandler;
  private cellEditingHandler: CellEditingHandler;

  constructor(props: IRenameFilesManagementProps) {
    super(props);
    
    this.state = {
      data: {
        originalFile: undefined,
        currentSheet: undefined,
        columns: [],
        rows: [],
        customColumns: [],
        totalRows: 0,
        editedCellsCount: 0
      },
      loading: false,
      error: undefined,
      uploadProgress: {
        stage: 'idle',
        progress: 0,
        message: ''
      },
      selectedCells: {},
      editingCell: undefined,
      showColumnManager: false,
      draggedColumn: undefined,
      showExportDialog: false,
      exportConfig: {
        fileName: 'renamed_files',
        includeOnlyEditedRows: false,
        includeCustomColumns: true,
        includeOriginalColumns: true,
        columnOrder: [],
        fileFormat: 'xlsx'
      },
      isExporting: false,
      // Export settings for status-based export
      exportSettings: {
        fileName: 'renamed_files_export',
        includeHeaders: true,
        includeStatusColumn: true,
        includeTimestamps: true,
        onlyCompletedRows: false,
        fileFormat: 'xlsx'
      },
      selectedFolder: undefined,
      showFolderDialog: false,
      availableFolders: [],
      loadingFolders: false,
      searchingFiles: false,
      // UPDATED: Separate directory and file status tracking
      directoryResults: {}, // NEW: Directory check results (after Stage 2)
      fileSearchResults: {}, // UPDATED: Only file search results (after Stage 3)
      fileRenameResults: {}, // Individual file rename status
      searchProgress: SearchProgressHelper.createInitialProgress(),
      // Rename state with skipped support
      isRenaming: false,
      renameProgress: undefined
    };

    this.fileInputRef = React.createRef<HTMLInputElement>();
    
    // Initialize services and handlers
    this.excelProcessor = new ExcelFileProcessor();
    this.folderService = new SharePointFolderService(props.context);
    this.fileSearchService = new FileSearchService(props.context);
    this.columnResizeHandler = new ColumnResizeHandler(this.updateColumnWidth);
    this.cellEditingHandler = new CellEditingHandler();
  }

  public componentDidMount(): void {
    this.columnResizeHandler.addEventListeners();
  }

  public componentWillUnmount(): void {
    this.columnResizeHandler.removeEventListeners();
    
    // Cancel any active search or rename
    if (this.state.searchingFiles) {
      this.fileSearchService.cancelSearch();
    }
  }

  // File handling methods
  private handleFileUpload = (): void => {
    if (this.fileInputRef.current) {
      this.fileInputRef.current.click();
    }
  }

  private handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const files = event.target.files;
    if (!files || files.length === 0) {
      return;
    }

    const file = files[0];
    await this.processExcelFile(file);

    // Reset file input
    if (this.fileInputRef.current) {
      this.fileInputRef.current.value = '';
    }
  }

  private processExcelFile = async (file: File): Promise<void> => {
    try {
      this.setState({ 
        loading: true, 
        error: undefined,
        uploadProgress: {
          stage: 'uploading',
          progress: 0,
          message: 'Reading file...'
        },
        // Reset ALL search and rename state when loading new file
        searchingFiles: false,
        isRenaming: false,
        directoryResults: {}, // NEW: Reset directory status
        fileSearchResults: {}, // Reset file search status
        fileRenameResults: {}, // Reset rename status
        searchProgress: SearchProgressHelper.createInitialProgress(),
        renameProgress: undefined,
        // Reset export settings with file name
        exportSettings: {
          ...this.state.exportSettings,
          fileName: this.generateExportFileName(file.name)
        }
      });

      const processedData = await this.excelProcessor.processFile(
        file, 
        this.handleUploadProgress
      );

      console.log('[RenameFilesManagement] File processed successfully:', {
        totalRows: processedData.totalRows,
        columns: processedData.columns.length,
        customColumns: processedData.customColumns.length
      });

      this.setState({
        data: processedData,
        loading: false,
        uploadProgress: {
          stage: 'complete',
          progress: 100,
          message: 'File loaded successfully!'
        }
      });

      // Clear progress message after delay
      setTimeout(() => {
        this.setState({
          uploadProgress: {
            stage: 'idle',
            progress: 0,
            message: ''
          }
        });
      }, 2000);

    } catch (error) {
      console.error('Error processing Excel file:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Failed to process Excel file',
        uploadProgress: {
          stage: 'error',
          progress: 0,
          message: 'Error processing file'
        }
      });
    }
  }

  private handleUploadProgress = (stage: string, progress: number, message: string): void => {
    this.setState({
      uploadProgress: {
        stage: stage as any,
        progress,
        message
      }
    });
  }

  // Generate export filename based on original file
  private generateExportFileName = (originalFileName: string): string => {
    const baseName = originalFileName.replace(/\.(xlsx|xls|csv)$/i, '');
    return `${baseName}_renamed`;
  }

  // Column management methods
  private handleAddCustomColumn = (): void => {
    const { data } = this.state;
    const updatedData = this.excelProcessor.addCustomColumn(data);
    this.setState({ data: updatedData });
  }

  private updateColumnWidth = (columnId: string, newWidth: number): void => {
    const { data } = this.state;
    const updatedData = this.excelProcessor.updateColumnWidth(data, columnId, newWidth);
    this.setState({ data: updatedData });
  }

  private updateCellData = (columnId: string, rowIndex: number, newValue: string): void => {
    const { data } = this.state;
    const updatedData = this.cellEditingHandler.updateCell(data, columnId, rowIndex, newValue);
    this.setState({ 
      data: updatedData,
      editingCell: undefined 
    });
  }

  // SharePoint folder methods
  private handleSelectFolder = async (): Promise<void> => {
    this.setState({ showFolderDialog: true, loadingFolders: true });
    
    try {
      const folders = await this.folderService.getDocumentLibraryFolders();
      this.setState({ 
        availableFolders: folders,
        loadingFolders: false 
      });
    } catch (error) {
      console.error('Error loading folders:', error);
      this.setState({ 
        error: 'Failed to load folders from Documents library',
        loadingFolders: false,
        showFolderDialog: false
      });
    }
  }

  private handleFolderSelect = (folder: any): void => {
    this.setState({ 
      selectedFolder: folder,
      showFolderDialog: false,
      // Clear ALL previous results when selecting new folder
      directoryResults: {}, // NEW: Clear directory status
      fileSearchResults: {}, // Clear file search results
      fileRenameResults: {}, // Clear rename results
      searchProgress: SearchProgressHelper.createInitialProgress(), // Reset progress
      isRenaming: false,
      renameProgress: undefined
    });
  }

  private handleCancelFolderSelection = (): void => {
    this.setState({ showFolderDialog: false });
  }

  private clearSelectedFolder = (): void => {
    this.setState({ 
      selectedFolder: undefined,
      directoryResults: {}, // NEW: Clear directory status
      fileSearchResults: {}, // Clear file search results
      fileRenameResults: {}, // Clear rename results
      searchProgress: SearchProgressHelper.createInitialProgress(),
      isRenaming: false,
      renameProgress: undefined
    });
  }

  // UPDATED: Handle directory analysis (Stages 1-2) with directory status callback
  private handleAnalyzeDirectories = async (): Promise<void> => {
    const { selectedFolder, data } = this.state;
    
    if (!selectedFolder) {
      this.setState({ error: 'Please select a SharePoint folder first' });
      return;
    }
    
    if (data.rows.length === 0) {
      this.setState({ error: 'No data rows to analyze' });
      return;
    }
    
    console.log('[RenameFilesManagement] Starting directory analysis...');
    
    this.setState({ 
      searchingFiles: true, 
      error: undefined,
      directoryResults: {}, // NEW: Reset directory results
      fileSearchResults: {}, // Reset file search results
      fileRenameResults: {}, // Reset rename results
      searchProgress: SearchProgressHelper.transitionToStage(
        this.state.searchProgress,
        SearchStage.ANALYZING_DIRECTORIES,
        {
          totalRows: data.rows.length,
          currentFileName: 'Starting directory analysis...'
        }
      ),
      isRenaming: false,
      renameProgress: undefined
    });
    
    try {
      const analysisProgress = await this.fileSearchService.analyzeDirectories(
        selectedFolder.ServerRelativeUrl,
        data.rows,
        this.updateSearchProgress,
        this.updateDirectoryStatus // NEW: Pass directory status callback
      );
      
      console.log('[RenameFilesManagement] Directory analysis completed:', {
        totalDirectories: analysisProgress.searchPlan?.totalDirectories,
        existingDirectories: analysisProgress.searchPlan?.existingDirectories
      });
      
      this.setState({ 
        searchProgress: analysisProgress
      });
      
    } catch (error) {
      console.error('Error analyzing directories:', error);
      this.setState({ 
        error: error instanceof Error ? error.message : 'Failed to analyze directories',
        searchProgress: SearchProgressHelper.transitionToStage(
          this.state.searchProgress,
          SearchStage.ERROR,
          {
            currentFileName: 'Analysis failed',
            errors: [error instanceof Error ? error.message : 'Unknown error']
          }
        )
      });
    } finally {
      this.setState({ searchingFiles: false });
    }
  }

  // UPDATED: Handle file search (Stage 3 only) with file search callback
  private handleSearchFiles = async (): Promise<void> => {
    const { data, searchProgress } = this.state;
    
    if (!searchProgress.searchPlan) {
      this.setState({ error: 'Please analyze directories first' });
      return;
    }
    
    if (data.rows.length === 0) {
      this.setState({ error: 'No data rows to search' });
      return;
    }
    
    console.log('[RenameFilesManagement] Starting file search...');
    
    this.setState({ 
      searchingFiles: true, 
      error: undefined,
      fileSearchResults: {}, // NEW: Reset only file search results
      fileRenameResults: {}, // Reset rename results
      isRenaming: false,
      renameProgress: undefined
    });
    
    try {
      const results = await this.fileSearchService.searchFilesInDirectories(
        searchProgress,
        data.rows,
        this.updateFileSearchResult, // NEW: Use separate file search callback
        this.updateSearchProgress
      );
      
      console.log('[RenameFilesManagement] File search completed:', {
        totalResults: Object.keys(results).length,
        foundFiles: Object.values(results).filter(r => r === 'found').length,
        notFoundFiles: Object.values(results).filter(r => r === 'not-found').length,
        directoryMissingFiles: Object.values(results).filter(r => r === 'directory-missing').length
      });
      
      this.setState({ 
        fileSearchResults: results
      });
      
    } catch (error) {
      console.error('Error searching files:', error);
      this.setState({ 
        error: error instanceof Error ? error.message : 'Failed to search files'
      });
    } finally {
      this.setState({ searchingFiles: false });
    }
  }

  // Handle file renaming with skipped support
  private handleRenameFiles = async (): Promise<void> => {
    const { data, fileSearchResults, selectedFolder } = this.state;
    
    if (!selectedFolder) {
      this.setState({ error: 'Please select a SharePoint folder first' });
      return;
    }
    
    const foundFilesCount = Object.values(fileSearchResults).filter(r => r === 'found').length;
    if (foundFilesCount === 0) {
      this.setState({ error: 'No files found to rename' });
      return;
    }
    
    console.log(`[RenameFilesManagement] Starting rename of ${foundFilesCount} files...`);
    
    this.setState({ 
      isRenaming: true, 
      error: undefined,
      fileRenameResults: {}, // Reset rename results
      renameProgress: {
        current: 0,
        total: foundFilesCount,
        fileName: '',
        success: 0,
        errors: 0,
        skipped: 0
      }
    });
    
    try {
      const results = await this.fileSearchService.renameFoundFiles(
        data.rows,
        fileSearchResults,
        selectedFolder.ServerRelativeUrl,
        this.updateRenameFileResult,
        this.updateRenameProgress
      );
      
      console.log('[RenameFilesManagement] Rename completed:', results);
      
      // Updated error handling for skipped files
      if (results.errors > 0 || results.skipped > 0) {
        let errorMessage = `Rename completed: ${results.success} files renamed successfully`;
        
        if (results.skipped > 0) {
          errorMessage += `, ${results.skipped} files skipped (target already exists)`;
        }
        
        if (results.errors > 0) {
          errorMessage += `, ${results.errors} files failed`;
        }
        
        errorMessage += '.';
        
        this.setState({ 
          error: errorMessage
        });
      } else {
        this.setState({ 
          error: undefined
        });
        // Success message could be shown here if needed
      }
      
    } catch (error) {
      console.error('Error renaming files:', error);
      this.setState({ 
        error: error instanceof Error ? error.message : 'Failed to rename files'
      });
    } finally {
      this.setState({ isRenaming: false });
    }
  }

  // Cancel rename operation
  private handleCancelRename = (): void => {
    console.log('[RenameFilesManagement] Cancelling rename...');
    
    this.fileSearchService.cancelSearch(); // Reuse the same cancel mechanism
    
    this.setState({ 
      isRenaming: false,
      renameProgress: undefined
    });
  }

  private handleCancelSearch = (): void => {
    console.log('[RenameFilesManagement] Cancelling search...');
    
    this.fileSearchService.cancelSearch();
    
    this.setState({ 
      searchingFiles: false,
      searchProgress: SearchProgressHelper.transitionToStage(
        this.state.searchProgress,
        SearchStage.CANCELLED,
        {
          currentFileName: 'Operation cancelled by user'
        }
      )
    });
  }

  private updateSearchProgress = (progress: ISearchProgress): void => {
    console.log('[RenameFilesManagement] Search progress update:', {
      stage: progress.currentStage,
      stageProgress: progress.stageProgress,
      overallProgress: progress.overallProgress,
      currentFile: progress.currentFileName
    });
    
    this.setState({
      searchProgress: progress
    });
  }

// NEW: Bulk directory status update (39 calls instead of 6000+)
private updateDirectoryStatus: DirectoryStatusCallback = (rowIndexes: number[], status: DirectoryStatus): void => {
  console.log(`[RenameFilesManagement] Bulk directory update: ${rowIndexes.length} rows -> ${status}`);
  
  this.setState(prevState => {
    const newDirectoryResults = { ...prevState.directoryResults };
    rowIndexes.forEach(rowIndex => {
      newDirectoryResults[rowIndex] = status;
    });
    return {
      directoryResults: newDirectoryResults
    };
  });
}

  // NEW: Separate callback for file search results (called after Stage 3)
  private updateFileSearchResult: FileSearchResultCallback = (rowIndex: number, result: FileSearchStatus): void => {
    console.log(`[RenameFilesManagement] File search result: Row ${rowIndex + 1} -> ${result}`);
    
    this.setState(prevState => ({
      fileSearchResults: {
        ...prevState.fileSearchResults,
        [rowIndex]: result
      }
    }));
  }

  // Update rename progress with skipped support
  private updateRenameProgress = (progress: { 
    current: number; 
    total: number; 
    fileName: string; 
    success: number; 
    errors: number; 
    skipped: number;
  }): void => {
    console.log('[RenameFilesManagement] Rename progress update:', progress);
    
    this.setState({
      renameProgress: progress
    });
  }

  // Update individual file rename result with skipped support
  private updateRenameFileResult = (rowIndex: number, status: 'renaming' | 'renamed' | 'error' | 'skipped'): void => {
    // Save the rename status for each individual file
    this.setState(prevState => ({
      fileRenameResults: {
        ...prevState.fileRenameResults,
        [rowIndex]: status
      }
    }));
    
    console.log(`[RenameFilesManagement] File ${rowIndex + 1} status: ${status}`);
    
    // Update progress callback to show correct icon
    if (status === 'skipped') {
      console.log(`[RenameFilesManagement] File ${rowIndex + 1} was skipped (target already exists)`);
    }
  }

  // Export functionality methods
  private handleExportSettingsChange = (newSettings: Partial<IRenameExportSettings>): void => {
    this.setState({
      exportSettings: {
        ...this.state.exportSettings,
        ...newSettings
      }
    });
  }

  private handleExport = async (): Promise<void> => {
    const { data, fileSearchResults, fileRenameResults, renameProgress, exportSettings } = this.state;
    
    if (!data.originalFile) {
      this.setState({ error: 'No data to export' });
      return;
    }

    // Validate export settings
    const validation = ExcelExportService.validateExportSettings(exportSettings);
    if (!validation.isValid) {
      this.setState({ error: `Export validation failed: ${validation.errors.join(', ')}` });
      return;
    }

    this.setState({ isExporting: true, error: undefined });

    try {
      console.log('[RenameFilesManagement] Starting export...');
      
      const result = await ExcelExportService.exportRenameFilesData(
        data,
        fileSearchResults,
        fileRenameResults, // Pass individual file rename results
        renameProgress,
        exportSettings
      );

      if (!result.success) {
        this.setState({ error: result.error || 'Export failed' });
      } else {
        console.log('[RenameFilesManagement] Export completed successfully:', result.fileName);
        // Could show success notification here
      }

    } catch (error) {
      console.error('[RenameFilesManagement] Export failed:', error);
      this.setState({ error: error instanceof Error ? error.message : 'Export failed' });
    } finally {
      this.setState({ isExporting: false });
    }
  }

  private clearError = (): void => {
    this.setState({ error: undefined });
  }

  // Render export controls section
  private renderExportControls = (): React.ReactNode => {
    const { data, fileSearchResults, fileRenameResults, renameProgress, exportSettings, isExporting } = this.state;

    if (!data.originalFile) {
      return null;
    }

    // Get export statistics
    const statistics = ExcelExportService.getRenameFilesExportStatistics(
      data,
      fileSearchResults,
      fileRenameResults, // Pass individual file rename results
      renameProgress,
      exportSettings
    );

    return (
      <ExportControlsPanel
        statistics={statistics}
        exportSettings={exportSettings}
        isExporting={isExporting}
        onExportSettingsChange={this.handleExportSettingsChange}
        onExport={this.handleExport}
      />
    );
  }

  public render(): React.ReactElement<IRenameFilesManagementProps> {
    const { 
      data, 
      loading, 
      error, 
      uploadProgress, 
      searchProgress,
      directoryResults, // NEW: Directory status
      fileSearchResults, // UPDATED: Only file search status
      selectedFolder,
      isRenaming,
      renameProgress
    } = this.state;
    
    const hasData = data.originalFile !== undefined;

    // UPDATED: Calculate statistics with separate directory and file status + FIXED logic
    const searchStats = {
      // Directory statistics (available after Stage 2)
      totalDirectories: Object.keys(directoryResults).length,
      existingDirectories: Object.values(directoryResults).filter(status => status === 'exists').length,
      missingDirectories: Object.values(directoryResults).filter(status => status === 'not-exists').length,
      
      // FIXED: File statistics with correct separation
      totalFiles: Object.keys(fileSearchResults).length,
      foundFiles: Object.values(fileSearchResults).filter(r => r === 'found').length,
      
      // CORRECTED: Only files NOT found in EXISTING directories
      notFoundFiles: Object.values(fileSearchResults).filter(r => r === 'not-found').length,
      
      // NEW: Files in missing directories (separate category)
      directoryMissingFiles: Object.values(fileSearchResults).filter(r => r === 'directory-missing').length,
      
      // Other statuses
      searchingFiles: Object.values(fileSearchResults).filter(r => r === 'searching').length,
      skippedFiles: Object.values(fileSearchResults).filter(r => r === 'skipped').length
    };

    return (
      <div className={styles.renameFilesManagement}>
        <div className={styles.header}>
          <div className={styles.title}>
            <h2>Rename Files Management</h2>
            <p>Upload Excel files with filename and directory columns to search and rename files in SharePoint</p>
          </div>
          
          <div className={styles.actions}>
            <input
              ref={this.fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={(e) => { this.handleFileChange(e).catch(console.error); }}
              style={{ display: 'none' }}
              disabled={loading || isRenaming}
            />
            
            <button
              className={styles.uploadButton}
              onClick={this.handleFileUpload}
              disabled={loading || isRenaming}
            >
              {loading ? (
                <>
                  <span className={styles.spinner} />
                  Loading...
                </>
              ) : (
                <>
                  📁 Open Excel File
                </>
              )}
            </button>

            {hasData && (
              <button
                className={styles.addColumnButton}
                onClick={this.handleAddCustomColumn}
                disabled={loading || isRenaming}
              >
                ➕ Add Column
              </button>
            )}
          </div>
        </div>

        <ProgressIndicator 
          uploadProgress={uploadProgress}
          error={error}
          onClearError={this.clearError}
        />

        {hasData ? (
          <div className={styles.content}>
            <div className={styles.tableInfo}>
              <div className={styles.fileInfo}>
                <strong>File:</strong> {data.originalFile?.name} | 
                <strong> Rows:</strong> {data.totalRows} | 
                <strong> Columns:</strong> {data.columns.length} |
                <strong> Edited Cells:</strong> {data.editedCellsCount}
                {/* UPDATED: Show directory statistics after Stage 2 */}
                {searchStats.totalDirectories > 0 && (
                  <>
                    {' | '}
                    <strong> Should be files in directories:</strong> {searchStats.existingDirectories}, {searchStats.missingDirectories} missing
                  </>
                )}
                {/* UPDATED: Show CORRECTED file statistics after Stage 3 */}
                {searchStats.totalFiles > 0 && (
                  <>
                    {' | '}
                    <strong> Files:</strong> {searchStats.foundFiles} found, {searchStats.notFoundFiles} not found
                    {searchStats.directoryMissingFiles > 0 && (
                      <>, {searchStats.directoryMissingFiles} dirs missing</>
                    )}
                    {searchStats.searchingFiles > 0 && (
                      <>, {searchStats.searchingFiles} searching</>
                    )}
                    {searchStats.skippedFiles > 0 && (
                      <>, {searchStats.skippedFiles} skipped</>
                    )}
                  </>
                )}
                {isRenaming && renameProgress && (
                  <>
                    {' | '}
                    <strong> Rename Progress:</strong> {renameProgress.success} renamed, {renameProgress.errors} errors
                    {renameProgress.skipped > 0 && (
                      <>, {renameProgress.skipped} skipped</>
                    )}
                  </>
                )}
              </div>
            </div>

            <RenameControlsPanel
              selectedFolder={selectedFolder}
              searchingFiles={this.state.searchingFiles}
              searchProgress={searchProgress}
              loading={loading}
              foundFilesCount={searchStats.foundFiles}
              isRenaming={isRenaming}
              renameProgress={renameProgress}
              onSelectFolder={this.handleSelectFolder}
              onClearFolder={this.clearSelectedFolder}
              onAnalyzeDirectories={this.handleAnalyzeDirectories}
              onSearchFiles={this.handleSearchFiles}
              onCancelSearch={this.handleCancelSearch}
              onRenameFiles={this.handleRenameFiles}
              onCancelRename={this.handleCancelRename}
            />

            {/* UPDATED: Pass both directory and file status to DataTableView */}
            <DataTableView
              data={data}
              directoryResults={directoryResults} // NEW: Pass directory status
              fileSearchResults={fileSearchResults} // UPDATED: Pass file search status
              columnResizeHandler={this.columnResizeHandler}
              onCellEdit={this.updateCellData}
            />

            {/* Export Controls Section */}
            {this.renderExportControls()}

            {/* Additional info for debugging/development */}
            {process.env.NODE_ENV === 'development' && searchProgress.searchPlan && (
              <div style={{
                marginTop: '20px',
                padding: '12px 16px',
                backgroundColor: '#f3f2f1',
                border: '1px solid #c8c6c4',
                borderRadius: '4px',
                fontFamily: "'Courier New', monospace",
                fontSize: '11px',
                color: '#605e5c'
              }}>
                <h4 style={{
                  margin: '0 0 8px 0',
                  fontSize: '12px',
                  fontWeight: 600,
                  color: '#323130',
                  fontFamily: '"Segoe UI", "Segoe UI Web (West European)", "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
                }}>Debug Information</h4>
                <p style={{ margin: '4px 0', lineHeight: 1.4 }}>
                  <strong style={{ color: '#323130' }}>Overall:</strong> {searchProgress.overallProgress.toFixed(1)}%
                </p>
                <p style={{ margin: '4px 0', lineHeight: 1.4 }}>
                  <strong style={{ color: '#323130' }}>Directory Status:</strong> {searchStats.existingDirectories} exist, {searchStats.missingDirectories} missing
                </p>
                <p style={{ margin: '4px 0', lineHeight: 1.4 }}>
                  <strong style={{ color: '#323130' }}>Found Files:</strong> {searchStats.foundFiles} ready for rename
                </p>
                <p style={{ margin: '4px 0', lineHeight: 1.4 }}>
                  <strong style={{ color: '#323130' }}>Not Found in Existing Dirs:</strong> {searchStats.notFoundFiles}
                </p>
                <p style={{ margin: '4px 0', lineHeight: 1.4 }}>
                  <strong style={{ color: '#323130' }}>Directory Missing:</strong> {searchStats.directoryMissingFiles}
                </p>
                {searchStats.skippedFiles > 0 && (
                  <p style={{ margin: '4px 0', lineHeight: 1.4 }}>
                    <strong style={{ color: '#323130' }}>Skipped Files:</strong> {searchStats.skippedFiles} (target already exists)
                  </p>
                )}
              </div>
            )}
          </div>
        ) : (
          <div className={styles.emptyState}>
            <div className={styles.emptyIcon}>📊</div>
            <h3>No Excel File Loaded</h3>
            <p>Click "Open Excel File" to start working with your data. The file should contain columns with file paths that will be automatically split into filename and directory columns.</p>
          </div>
        )}

        <FolderSelectionDialog
          isOpen={this.state.showFolderDialog}
          folders={this.state.availableFolders}
          loading={this.state.loadingFolders}
          onSelect={this.handleFolderSelect}
          onCancel={this.handleCancelFolderSelection}
        />
      </div>
    );
  }
} 