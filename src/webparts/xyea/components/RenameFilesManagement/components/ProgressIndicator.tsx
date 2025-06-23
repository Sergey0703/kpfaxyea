// src/webparts/xyea/components/RenameFilesManagement/components/ProgressIndicator.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { IUploadProgress } from '../types/RenameFilesTypes';

export interface IProgressIndicatorProps {
  uploadProgress: IUploadProgress;
  error: string | undefined;
  onClearError: () => void;
}

export const ProgressIndicator: React.FC<IProgressIndicatorProps> = ({
  uploadProgress,
  error,
  onClearError
}) => {
  const showProgress = uploadProgress.stage !== 'idle';
  const showError = error !== undefined;

  return (
    <>
      {/* Progress indicator */}
      {showProgress && (
        <div className={styles.progressContainer}>
          <div className={styles.progressMessage}>{uploadProgress.message}</div>
          <div className={styles.progressBar}>
            <div 
              className={styles.progressFill}
              style={{ width: `${uploadProgress.progress}%` }}
            />
          </div>
        </div>
      )}

      {/* Error display */}
      {showError && (
        <div className={styles.error}>
          <span className={styles.errorIcon}>⚠️</span>
          <span className={styles.errorMessage}>{error}</span>
          <button 
            className={styles.clearErrorButton}
            onClick={onClearError}
          >
            ✕
          </button>
        </div>
      )}
    </>
  );
};