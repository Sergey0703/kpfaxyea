// src/webparts/xyea/components/FolderStructureManagement/FolderStructureManagement.tsx

import * as React from 'react';
import styles from './FolderStructureManagement.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointFoldersService, ISharePointFolder } from '../../services/SharePointFoldersService';
import { FolderStructureExportService } from '../../services/FolderStructureExportService';

export interface IFolderStructureManagementProps {
  context: WebPartContext;
  userDisplayName: string;
}

export interface IFolderStructureState {
  folderPath: string;
  folders: ISharePointFolder[];
  loading: boolean;
  error: string | undefined;
  hasSearched: boolean;
  // Export state
  isExporting: boolean;
  exportError: string | undefined;
}

export default class FolderStructureManagement extends React.Component<IFolderStructureManagementProps, IFolderStructureState> {
  private foldersService: SharePointFoldersService;

  constructor(props: IFolderStructureManagementProps) {
    super(props);
    
    this.state = {
      folderPath: '',
      folders: [],
      loading: false,
      error: undefined,
      hasSearched: false,
      isExporting: false,
      exportError: undefined
    };

    this.foldersService = new SharePointFoldersService(this.props.context);
  }

  private handlePathChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ 
      folderPath: event.target.value,
      error: undefined,
      exportError: undefined
    });
  }

  private handleSearchFolders = async (): Promise<void> => {
    const { folderPath } = this.state;
    
    if (!folderPath.trim()) {
      this.setState({ 
        error: 'Please enter a folder path' 
      });
      return;
    }

    try {
      this.setState({ 
        loading: true, 
        error: undefined,
        exportError: undefined,
        hasSearched: false
      });

      console.log('[FolderStructureManagement] Starting recursive scan for path:', folderPath);

      // This now returns recursive structure like Python script
      const folders = await this.foldersService.getFoldersInPath(folderPath.trim());

      console.log('[FolderStructureManagement] Recursive scan completed:', {
        totalItems: folders.length,
        files: folders.filter(item => !(item as any).IsFile === false).length,
        folders: folders.filter(item => (item as any).IsFile === false).length
      });

      this.setState({
        folders,
        loading: false,
        hasSearched: true
      });

    } catch (error) {
      console.error('[FolderStructureManagement] Error during recursive scan:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Failed to load folder structure',
        hasSearched: true,
        folders: []
      });
    }
  }

  private handleExportToExcel = async (): Promise<void> => {
    const { folders, folderPath } = this.state;

    if (folders.length === 0) {
      this.setState({ exportError: 'No data to export. Please scan a folder first.' });
      return;
    }

    try {
      this.setState({ 
        isExporting: true, 
        exportError: undefined 
      });

      console.log('[FolderStructureManagement] Starting export to Excel:', {
        totalItems: folders.length,
        folderPath
      });

      // Create export settings
      const exportSettings = FolderStructureExportService.createDefaultExportSettings(
        `folder_structure_${folderPath.replace(/[^a-zA-Z0-9]/g, '_')}`
      );

      // Export to Excel
      const result = await FolderStructureExportService.exportFolderStructure(
        folderPath,
        folders,
        exportSettings
      );

      if (result.success) {
        console.log('[FolderStructureManagement] Export completed successfully:', result.fileName);
        // Show success message (optional)
        this.setState({ 
          isExporting: false,
          exportError: undefined
        });
      } else {
        throw new Error(result.error || 'Export failed');
      }

    } catch (error) {
      console.error('[FolderStructureManagement] Export failed:', error);
      this.setState({
        isExporting: false,
        exportError: error instanceof Error ? error.message : 'Export failed'
      });
    }
  }

  private handleKeyPress = (event: React.KeyboardEvent<HTMLInputElement>): void => {
    if (event.key === 'Enter') {
      this.handleSearchFolders().catch(console.error);
    }
  }

  private formatDate = (dateString: string): string => {
    try {
      return new Date(dateString).toLocaleDateString();
    } catch {
      return dateString;
    }
  }

  private formatFileSize = (itemCount: number): string => {
    if (itemCount === 0) return '0 bytes';
    
    if (itemCount < 1024) {
      return `${itemCount} bytes`;
    } else if (itemCount < 1024 * 1024) {
      return `${(itemCount / 1024).toFixed(1)} KB`;
    } else if (itemCount < 1024 * 1024 * 1024) {
      return `${(itemCount / (1024 * 1024)).toFixed(1)} MB`;
    } else {
      return `${(itemCount / (1024 * 1024 * 1024)).toFixed(1)} GB`;
    }
  }

  private renderHierarchicalStructure = (): React.ReactElement => {
    const { folders } = this.state;

    return (
      <div className={styles.hierarchicalView}>
        {folders.map((item, index) => {
          const level = (item as any).Level || 0;
          const isFile = (item as any).IsFile === true;
          const isError = item.Name.includes('[Permission Denied]') || item.Name.includes('[Error:');
          
          return (
            <div key={`${item.ServerRelativeUrl}-${index}`} className={styles.hierarchicalItem}>
              <div 
                className={styles.itemIndent}
                style={{ 
                  paddingLeft: `${level * 20 + 12}px`,
                  fontFamily: "'Consolas', 'Monaco', 'Courier New', monospace"
                }}
              >
                <span className={styles.itemIcon}>
                  {isFile ? '📄' : '📁'}
                </span>
                <span 
                  className={styles.itemName}
                  style={{
                    fontWeight: isFile ? 'normal' : '600',
                    color: isError ? '#d13438' : (isFile ? '#323130' : '#0078d4')
                  }}
                >
                  {item.Name}
                </span>
                {isFile && !isError && item.ItemCount > 0 && (
                  <span className={styles.fileSize}>
                    ({this.formatFileSize(item.ItemCount)})
                  </span>
                )}
                {isError && (
                  <span className={styles.errorIndicator}> ⚠️</span>
                )}
                {!isFile && !isError && (
                  <span className={styles.folderMeta}>
                    Modified: {this.formatDate(item.TimeLastModified)}
                  </span>
                )}
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  public render(): React.ReactElement<IFolderStructureManagementProps> {
    const { folderPath, folders, loading, error, hasSearched, isExporting, exportError } = this.state;

    // Determine if we have hierarchical data
    const hasHierarchicalData = folders.length > 0 && folders.some(item => (item as any).Level !== undefined);
    
    // Calculate statistics for hierarchical data
    const folderStats = hasHierarchicalData ? {
      totalItems: folders.length,
      files: folders.filter(item => (item as any).IsFile === true).length,
      folders: folders.filter(item => (item as any).IsFile === false).length,
      maxLevel: Math.max(...folders.map(item => (item as any).Level || 0))
    } : {
      totalItems: folders.length,
      files: 0,
      folders: folders.length,
      maxLevel: 0
    };

    // Get export statistics
    const exportStats = folders.length > 0 
      ? FolderStructureExportService.getFolderStructureExportStatistics(folders)
      : null;

    return (
      <div className={styles.folderStructure}>
        <div className={styles.container}>
          
          {/* Header */}
          <div className={styles.header}>
            <h2>Folder Structure Explorer</h2>
            <p>Enter a SharePoint folder path to explore its complete structure recursively (like Python script)</p>
          </div>

          {/* Search Section */}
          <div className={styles.searchSection}>
            <div className={styles.pathInputGroup}>
              <label htmlFor="folderPath" className={styles.pathLabel}>
                SharePoint Folder Path:
              </label>
              <div className={styles.inputWrapper}>
                <input
                  id="folderPath"
                  type="text"
                  value={folderPath}
                  onChange={this.handlePathChange}
                  onKeyPress={this.handleKeyPress}
                  placeholder="/Shared Documents"
                  className={styles.pathInput}
                  disabled={loading || isExporting}
                />
                <button
                  onClick={this.handleSearchFolders}
                  disabled={loading || !folderPath.trim() || isExporting}
                  className={styles.searchButton}
                >
                  {loading ? (
                    <>
                      <div className={styles.spinner}></div>
                      Scanning...
                    </>
                  ) : (
                    '🔍 Explore Folders'
                  )}
                </button>
              </div>
            </div>

            {/* Path Examples */}
            <div className={styles.examples}>
              <span className={styles.examplesLabel}>Examples (click to use):</span>
              <ul className={styles.examplesList}>
                <li onClick={() => this.setState({ folderPath: '/Shared Documents' })}>/Shared Documents</li>
                <li onClick={() => this.setState({ folderPath: '/sites/YourSite/Shared Documents' })}>/sites/YourSite/Shared Documents</li>
                <li onClick={() => this.setState({ folderPath: '/sites/KPFADataBackUp/Shared Documents' })}>/sites/KPFADataBackUp/Shared Documents</li>
              </ul>
            </div>

            {/* Algorithm Info */}
            <div className={styles.algorithmInfo}>
              <div className={styles.infoIcon}>🌲</div>
              <div className={styles.infoText}>
                <strong>Complete Recursive Scan:</strong> Explores ALL levels deep with no depth limit, 
                shows folders first then files with proper indentation and comprehensive error handling.
              </div>
            </div>
          </div>

          {/* Error Display */}
          {error && (
            <div className={styles.error}>
              <strong>Error:</strong> {error}
              <button 
                className={styles.retryButton}
                onClick={this.handleSearchFolders}
                disabled={loading || isExporting}
              >
                Try Again
              </button>
            </div>
          )}

          {/* Export Error Display */}
          {exportError && (
            <div className={styles.error}>
              <strong>Export Error:</strong> {exportError}
              <button 
                className={styles.retryButton}
                onClick={this.handleExportToExcel}
                disabled={loading || isExporting}
              >
                Try Export Again
              </button>
            </div>
          )}

          {/* Results Section */}
          {hasSearched && !loading && !error && (
            <div className={styles.resultsSection}>
              <div className={styles.resultsHeader}>
                <h3>
                  {hasHierarchicalData ? 'Complete Structure:' : 'Subfolders in:'} {folderPath}
                </h3>
                <div className={styles.resultsStats}>
                  {hasHierarchicalData ? (
                    <>
                      <span className={styles.statItem}>📁 {folderStats.folders} folders</span>
                      <span className={styles.statItem}>📄 {folderStats.files} files</span>
                      <span className={styles.statItem}>📊 {folderStats.maxLevel + 1} levels</span>
                      <span className={styles.statItem}>📋 {folderStats.totalItems} total</span>
                    </>
                  ) : (
                    <span className={styles.statItem}>
                      {folders.length} folder{folders.length !== 1 ? 's' : ''} found
                    </span>
                  )}
                </div>
              </div>

              {/* Export Button Section */}
              {folders.length > 0 && (
                <div style={{ 
                  marginBottom: '20px', 
                  padding: '16px', 
                  backgroundColor: '#f8f7f6', 
                  borderRadius: '8px',
                  border: '1px solid #edebe9'
                }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <div>
                      <h4 style={{ 
                        margin: '0 0 8px 0', 
                        color: '#323130', 
                        fontSize: '16px', 
                        fontWeight: '600' 
                      }}>
                        📊 Export Folder Structure
                      </h4>
                      {exportStats && (
                        <div style={{ fontSize: '12px', color: '#605e5c' }}>
                          Ready to export: {exportStats.totalFiles} files, {exportStats.totalFolders} folders 
                          ({exportStats.maxDepth} levels deep, ~{exportStats.estimatedFileSize})
                        </div>
                      )}
                    </div>
                    <button
                      onClick={this.handleExportToExcel}
                      disabled={loading || isExporting || folders.length === 0}
                      style={{
                        padding: '12px 24px',
                        backgroundColor: isExporting ? '#a19f9d' : '#107c10',
                        color: 'white',
                        border: 'none',
                        borderRadius: '4px',
                        fontSize: '14px',
                        fontWeight: '600',
                        cursor: isExporting ? 'not-allowed' : 'pointer',
                        transition: 'background-color 0.2s ease',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '8px',
                        minWidth: '160px',
                        justifyContent: 'center'
                      }}
                      onMouseOver={(e) => {
                        if (!isExporting && folders.length > 0) {
                          e.currentTarget.style.backgroundColor = '#0e6b0e';
                        }
                      }}
                      onMouseOut={(e) => {
                        if (!isExporting && folders.length > 0) {
                          e.currentTarget.style.backgroundColor = '#107c10';
                        }
                      }}
                    >
                      {isExporting ? (
                        <>
                          <div style={{
                            width: '16px',
                            height: '16px',
                            border: '2px solid transparent',
                            borderTop: '2px solid white',
                            borderRadius: '50%',
                            animation: 'spin 1s linear infinite'
                          }}></div>
                          Exporting...
                        </>
                      ) : (
                        <>
                          📊 Export to Excel
                        </>
                      )}
                    </button>
                  </div>
                </div>
              )}

              {folders.length === 0 ? (
                <div className={styles.noResults}>
                  <div className={styles.noResultsIcon}>📁</div>
                  <h4>No items found</h4>
                  <p>The specified path doesn't contain any items, or you may not have access to view them.</p>
                </div>
              ) : hasHierarchicalData ? (
                // Hierarchical tree view (like Python script output)
                <>
                  <div className={styles.treeViewHeader}>
                    <span className={styles.treeIcon}>🌳</span>
                    <span>Hierarchical Structure (Folders first, then files - like Python script)</span>
                  </div>
                  {this.renderHierarchicalStructure()}
                </>
              ) : (
                // Fallback: Regular grid view for simple folder lists
                <div className={styles.foldersGrid}>
                  {folders.map((folder, index) => (
                    <div key={`${folder.ServerRelativeUrl}-${index}`} className={styles.folderCard}>
                      <div className={styles.folderIcon}>📁</div>
                      <div className={styles.folderInfo}>
                        <h4 className={styles.folderName} title={folder.Name}>
                          {folder.Name}
                        </h4>
                        <div className={styles.folderDetails}>
                          <span className={styles.itemCount}>
                            {this.formatFileSize(folder.ItemCount)}
                          </span>
                          <span className={styles.lastModified}>
                            Modified: {this.formatDate(folder.TimeLastModified)}
                          </span>
                        </div>
                        <div className={styles.folderPath} title={folder.ServerRelativeUrl}>
                          {folder.ServerRelativeUrl}
                        </div>
                      </div>
                      <div className={styles.folderActions}>
                        <button 
                          className={styles.actionButton}
                          onClick={() => {
                            this.setState({ folderPath: folder.ServerRelativeUrl });
                          }}
                          title="Navigate to this folder"
                          disabled={loading || isExporting}
                        >
                          Navigate
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* Loading State */}
          {loading && !error && (
            <div className={styles.loadingState}>
              <div className={styles.loadingSpinner}></div>
              <h3>Scanning folder structure...</h3>
              <p>Please wait while we recursively scan the folder structure (like Python script). This may take a moment for large directories.</p>
              <div className={styles.loadingSteps}>
                <div className={styles.step}>📁 Reading folders and files...</div>
                <div className={styles.step}>🔍 Scanning subdirectories...</div>
                <div className={styles.step}>📋 Building hierarchical structure...</div>
              </div>
            </div>
          )}

        </div>
      </div>
    );
  }
}