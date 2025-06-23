// src/webparts/xyea/components/RenameFilesManagement/components/RenameControlsPanel.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { ISharePointFolder } from '../types/RenameFilesTypes';

export interface IRenameControlsPanelProps {
  selectedFolder: ISharePointFolder | undefined;
  searchingFiles: boolean;
  searchProgress: {
    currentRow: number;
    totalRows: number;
    currentFileName: string;
  };
  loading: boolean;
  onSelectFolder: () => void;
  onClearFolder: () => void;
  onSearchFiles: () => void;
  onCancelSearch: () => void;
}

export const RenameControlsPanel: React.FC<IRenameControlsPanelProps> = ({
  selectedFolder,
  searchingFiles,
  searchProgress,
  loading,
  onSelectFolder,
  onClearFolder,
  onSearchFiles,
  onCancelSearch
}) => {
  return (
    <>
      {/* SharePoint folder selection */}
      <div className={styles.folderSelection}>
        <div className={styles.folderControls}>
          <button
            className={styles.selectFolderButton}
            onClick={onSelectFolder}
            disabled={loading || searchingFiles}
          >
            📁 Select SharePoint Folder
          </button>
          
          {selectedFolder && (
            <div className={styles.selectedFolder}>
              <span className={styles.folderIcon}>📂</span>
              <span className={styles.folderName}>
                {selectedFolder.Name}
              </span>
              <button 
                className={styles.clearFolderButton}
                onClick={onClearFolder}
                title="Clear selection"
                disabled={searchingFiles}
              >
                ✕
              </button>
            </div>
          )}
        </div>
        
        {selectedFolder && (
          <div className={styles.folderInfo}>
            <small>Selected: {selectedFolder.ServerRelativeUrl}</small>
          </div>
        )}
      </div>

      {/* Rename Files Controls */}
      <div className={styles.renameControls}>
        <div className={styles.renameButtons}>
          <button
            className={styles.renameButton}
            onClick={onSearchFiles}
            disabled={loading || searchingFiles || !selectedFolder}
          >
            {searchingFiles ? (
              <>
                <span className={styles.spinner} />
                Searching...
              </>
            ) : (
              <>
                🔍 Rename
              </>
            )}
          </button>
          
          {searchingFiles && (
            <button
              className={styles.cancelButton}
              onClick={onCancelSearch}
            >
              ❌ Cancel
            </button>
          )}
        </div>
        
        {searchingFiles && searchProgress.totalRows > 0 && (
          <div className={styles.searchProgressInfo}>
            <div className={styles.progressText}>
              <strong>Row {searchProgress.currentRow} of {searchProgress.totalRows}</strong>
              {searchProgress.currentFileName && (
                <span className={styles.currentFile}>
                  Searching: {searchProgress.currentFileName}
                </span>
              )}
            </div>
            <div className={styles.progressBar}>
              <div 
                className={styles.progressFill}
                style={{ 
                  width: `${(searchProgress.currentRow / searchProgress.totalRows) * 100}%` 
                }}
              />
            </div>
            <div className={styles.progressStats}>
              {Math.round((searchProgress.currentRow / searchProgress.totalRows) * 100)}% complete
            </div>
          </div>
        )}
        
        {!selectedFolder && !searchingFiles && (
          <div className={styles.searchNote}>
            <small>Please select a SharePoint folder first</small>
          </div>
        )}
      </div>
    </>
  );
};