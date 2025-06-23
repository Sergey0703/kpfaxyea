// src/webparts/xyea/components/RenameFilesManagement/components/RenameControlsPanel.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { ISharePointFolder } from '../types/RenameFilesTypes';

export interface IRenameControlsPanelProps {
  selectedFolder: ISharePointFolder | undefined;
  searchingFiles: boolean;
  loading: boolean;
  onSelectFolder: () => void;
  onClearFolder: () => void;
  onSearchFiles: () => void;
}

export const RenameControlsPanel: React.FC<IRenameControlsPanelProps> = ({
  selectedFolder,
  searchingFiles,
  loading,
  onSelectFolder,
  onClearFolder,
  onSearchFiles
}) => {
  return (
    <>
      {/* SharePoint folder selection */}
      <div className={styles.folderSelection}>
        <div className={styles.folderControls}>
          <button
            className={styles.selectFolderButton}
            onClick={onSelectFolder}
            disabled={loading}
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
        <button
          className={styles.renameButton}
          onClick={onSearchFiles}
          disabled={loading || searchingFiles || !selectedFolder}
        >
          {searchingFiles ? (
            <>
              <span className={styles.spinner} />
              Searching Files...
            </>
          ) : (
            <>
              🔍 Rename
            </>
          )}
        </button>
        
        {searchingFiles && (
          <div className={styles.searchProgress}>
            Searching for files in selected folder...
          </div>
        )}
        
        {!selectedFolder && (
          <div className={styles.searchNote}>
            <small>Please select a SharePoint folder first</small>
          </div>
        )}
      </div>
    </>
  );
};