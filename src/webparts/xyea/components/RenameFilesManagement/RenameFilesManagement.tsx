// src/webparts/xyea/components/RenameFilesManagement/RenameFilesManagement.tsx

import * as React from 'react';
import styles from './RenameFilesManagement.module.scss';
import { 
  IRenameFilesManagementProps, 
  IRenameFilesState,
  SearchStage,
  ISearchProgress,
  SearchProgressHelper
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
      selectedFolder: undefined,
      showFolderDialog: false,
      availableFolders: [],
      loadingFolders: false,
      searchingFiles: false,
      fileSearchResults: {},
      searchProgress: SearchProgressHelper.createInitialProgress() // NEW: Use helper to create initial progress
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
    
    // Cancel any active search
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
        // Reset search state when loading new file
        searchingFiles: false,
        fileSearchResults: {},
        searchProgress: SearchProgressHelper.createInitialProgress()
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
      fileSearchResults: {}, // Clear previous search results
      searchProgress: SearchProgressHelper.createInitialProgress() // Reset progress
    });
  }

  private handleCancelFolderSelection = (): void => {
    this.setState({ showFolderDialog: false });
  }

  private clearSelectedFolder = (): void => {
    this.setState({ 
      selectedFolder: undefined,
      fileSearchResults: {},
      searchProgress: SearchProgressHelper.createInitialProgress()
    });
  }

  // NEW: Enhanced file search methods with three-stage progress
  private handleSearchFiles = async (): Promise<void> => {
    const { selectedFolder, data } = this.state;
    
    if (!selectedFolder) {
      this.setState({ error: 'Please select a SharePoint folder first' });
      return;
    }
    
    if (data.rows.length === 0) {
      this.setState({ error: 'No data rows to search' });
      return;
    }
    
    console.log('[RenameFilesManagement] Starting three-stage file search...');
    
    this.setState({ 
      searchingFiles: true, 
      error: undefined,
      fileSearchResults: {},
      searchProgress: SearchProgressHelper.transitionToStage(
        this.state.searchProgress,
        SearchStage.ANALYZING_DIRECTORIES,
        {
          totalRows: data.rows.length,
          currentFileName: 'Preparing search...'
        }
      )
    });
    
    try {
      const results = await this.fileSearchService.searchFiles(
        selectedFolder.ServerRelativeUrl,
        data.rows,
        this.updateSearchResult,
        this.updateSearchProgress
      );
      
      console.log('[RenameFilesManagement] Search completed with results:', {
        totalResults: Object.keys(results).length,
        foundFiles: Object.values(results).filter(r => r === 'found').length,
        notFoundFiles: Object.values(results).filter(r => r === 'not-found').length
      });
      
      // Final update with completion
      this.setState({ 
        fileSearchResults: results,
        searchProgress: SearchProgressHelper.transitionToStage(
          this.state.searchProgress,
          SearchStage.COMPLETED,
          {
            currentFileName: 'Search completed successfully',
            overallProgress: 100,
            filesFound: Object.values(results).filter(r => r === 'found').length,
            filesSearched: Object.keys(results).length
          }
        )
      });
      
    } catch (error) {
      console.error('Error searching files:', error);
      this.setState({ 
        error: error instanceof Error ? error.message : 'Failed to search files',
        searchProgress: SearchProgressHelper.transitionToStage(
          this.state.searchProgress,
          SearchStage.ERROR,
          {
            currentFileName: 'Search failed',
            errors: [error instanceof Error ? error.message : 'Unknown error']
          }
        )
      });
    } finally {
      this.setState({ searchingFiles: false });
    }
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
          currentFileName: 'Search cancelled by user'
        }
      )
    });
  }

  // NEW: Enhanced progress update method
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

  private updateSearchResult = (rowIndex: number, result: 'found' | 'not-found' | 'searching'): void => {
    this.setState(prevState => ({
      fileSearchResults: {
        ...prevState.fileSearchResults,
        [rowIndex]: result
      }
    }));
  }

  private clearError = (): void => {
    this.setState({ error: undefined });
  }

  public render(): React.ReactElement<IRenameFilesManagementProps> {
    const { data, loading, error, uploadProgress, searchProgress } = this.state;
    const hasData = data.originalFile !== undefined;

    // Calculate search statistics for display
    const searchStats = {
      totalFiles: Object.keys(this.state.fileSearchResults).length,
      foundFiles: Object.values(this.state.fileSearchResults).filter(r => r === 'found').length,
      notFoundFiles: Object.values(this.state.fileSearchResults).filter(r => r === 'not-found').length,
      searchingFiles: Object.values(this.state.fileSearchResults).filter(r => r === 'searching').length
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
              disabled={loading}
            />
            
            <button
              className={styles.uploadButton}
              onClick={this.handleFileUpload}
              disabled={loading}
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
                disabled={loading}
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
                {searchStats.totalFiles > 0 && (
                  <>
                    {' | '}
                    <strong> Search Results:</strong> {searchStats.foundFiles} found, {searchStats.notFoundFiles} not found
                    {searchStats.searchingFiles > 0 && (
                      <>, {searchStats.searchingFiles} searching</>
                    )}
                  </>
                )}
              </div>
            </div>

            <RenameControlsPanel
              selectedFolder={this.state.selectedFolder}
              searchingFiles={this.state.searchingFiles}
              searchProgress={searchProgress} // NEW: Pass enhanced progress
              loading={loading}
              onSelectFolder={this.handleSelectFolder}
              onClearFolder={this.clearSelectedFolder}
              onSearchFiles={this.handleSearchFiles}
              onCancelSearch={this.handleCancelSearch}
            />

            <DataTableView
              data={data}
              fileSearchResults={this.state.fileSearchResults}
              columnResizeHandler={this.columnResizeHandler}
              onCellEdit={this.updateCellData}
            />

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
                  <strong style={{ color: '#323130' }}>Search Plan:</strong> {searchProgress.searchPlan.totalDirectories} directories, {searchProgress.searchPlan.existingDirectories} exist
                </p>
                <p style={{ margin: '4px 0', lineHeight: 1.4 }}>
                  <strong style={{ color: '#323130' }}>Stage:</strong> {searchProgress.currentStage} ({searchProgress.stageProgress.toFixed(1)}%)
                </p>
                <p style={{ margin: '4px 0', lineHeight: 1.4 }}>
                  <strong style={{ color: '#323130' }}>Overall:</strong> {searchProgress.overallProgress.toFixed(1)}%
                </p>
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