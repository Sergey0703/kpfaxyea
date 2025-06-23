// src/webparts/xyea/components/RenameFilesManagement/components/FolderSelectionDialog.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { ISharePointFolder } from '../types/RenameFilesTypes';

export interface IFolderSelectionDialogProps {
  isOpen: boolean;
  folders: ISharePointFolder[];
  loading: boolean;
  onSelect: (folder: ISharePointFolder) => void;
  onCancel: () => void;
}

export const FolderSelectionDialog: React.FC<IFolderSelectionDialogProps> = ({
  isOpen,
  folders,
  loading,
  onSelect,
  onCancel
}) => {
  if (!isOpen) {
    return null;
  }

  return (
    <div className={styles.folderDialog}>
      <div className={styles.dialogOverlay} onClick={onCancel} />
      <div className={styles.dialogContent}>
        <div className={styles.dialogHeader}>
          <h3>Select SharePoint Folder</h3>
          <button 
            className={styles.closeButton}
            onClick={onCancel}
          >
            ✕
          </button>
        </div>
        
        <div className={styles.dialogBody}>
          {loading ? (
            <div className={styles.loadingFolders}>
              <div className={styles.spinner} />
              <span>Loading folders...</span>
            </div>
          ) : (
            <div className={styles.folderList}>
              {folders.map((folder, index) => (
                <div
                  key={index}
                  className={styles.folderItem}
                  onClick={() => onSelect(folder)}
                >
                  <span className={styles.folderIcon}>📁</span>
                  <div className={styles.folderDetails}>
                    <div className={styles.folderName}>{folder.Name}</div>
                    <div className={styles.folderPath}>{folder.ServerRelativeUrl}</div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
        
        <div className={styles.dialogFooter}>
          <button 
            className={styles.cancelButton}
            onClick={onCancel}
          >
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
};