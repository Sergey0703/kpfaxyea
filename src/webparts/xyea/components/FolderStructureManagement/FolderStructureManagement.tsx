// src/webparts/xyea/components/FolderStructureManagement/FolderStructureManagement.tsx

import * as React from 'react';
import styles from './FolderStructureManagement.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointFoldersService, ISharePointFolder } from '../../services/SharePointFoldersService';

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
      hasSearched: false
    };

    this.foldersService = new SharePointFoldersService(this.props.context);
  }

  private handlePathChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ 
      folderPath: event.target.value,
      error: undefined 
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
    const { folderPath, folders, loading, error, hasSearched } = this.state;

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
                  disabled={loading}
                />
                <button
                  onClick={this.handleSearchFolders}
                  disabled={loading || !folderPath.trim()}
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
                shows files first then folders with proper indentation and comprehensive error handling.
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
                disabled={loading}
              >
                Try Again
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
                    <span>Hierarchical Structure (Files first, then folders - like Python script)</span>
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