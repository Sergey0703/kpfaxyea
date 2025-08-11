// src/webparts/xyea/components/FolderStructureManagement/FolderStructureManagement.tsx

import * as React from 'react';
import styles from './FolderStructureManagement.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointFoldersService } from '../../services/SharePointFoldersService';

export interface IFolderStructureManagementProps {
  context: WebPartContext;
  userDisplayName: string;
}

export interface ISharePointFolder {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  TimeCreated: string;
  TimeLastModified: string;
  Exists: boolean;
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

      console.log('[FolderStructureManagement] Searching folders for path:', folderPath);

      const folders = await this.foldersService.getFoldersInPath(folderPath.trim());

      console.log('[FolderStructureManagement] Found folders:', {
        count: folders.length,
        folders: folders.slice(0, 3)
      });

      this.setState({
        folders,
        loading: false,
        hasSearched: true
      });

    } catch (error) {
      console.error('[FolderStructureManagement] Error searching folders:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Failed to load folders',
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
    return `${itemCount} items`;
  }

  public render(): React.ReactElement<IFolderStructureManagementProps> {
    const { folderPath, folders, loading, error, hasSearched } = this.state;

    return (
      <div className={styles.folderStructure}>
        <div className={styles.container}>
          
          {/* Header */}
          <div className={styles.header}>
            <h2>Folder Structure Explorer</h2>
            <p>Enter a SharePoint folder path to explore its subfolders</p>
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
                  placeholder="/sites/YourSite/Shared Documents/YourFolder"
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
                      Loading...
                    </>
                  ) : (
                    'Explore Folders'
                  )}
                </button>
              </div>
            </div>

            {/* Path Examples */}
            <div className={styles.examples}>
              <span className={styles.examplesLabel}>Examples:</span>
              <ul className={styles.examplesList}>
                <li>/sites/YourSite/Shared Documents</li>
                <li>/sites/YourSite/Shared Documents/Projects</li>
                <li>/Shared Documents/Department</li>
              </ul>
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
                <h3>Subfolders in: {folderPath}</h3>
                <span className={styles.resultsCount}>
                  {folders.length} folder{folders.length !== 1 ? 's' : ''} found
                </span>
              </div>

              {folders.length === 0 ? (
                <div className={styles.noResults}>
                  <div className={styles.noResultsIcon}>📁</div>
                  <h4>No subfolders found</h4>
                  <p>The specified path doesn't contain any subfolders, or you may not have access to view them.</p>
                </div>
              ) : (
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
              <h3>Exploring folders...</h3>
              <p>Please wait while we retrieve the folder structure from SharePoint.</p>
            </div>
          )}

        </div>
      </div>
    );
  }
}