// src/webparts/xyea/components/RenameFilesManagement/components/RenameControlsPanel.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { 
  ISharePointFolder, 
  ISearchProgress, 
  SearchStage, 
  SEARCH_STAGES 
} from '../types/RenameFilesTypes';

export interface IRenameControlsPanelProps {
  selectedFolder: ISharePointFolder | undefined;
  searchingFiles: boolean;
  searchProgress: ISearchProgress;
  loading: boolean;
  onSelectFolder: () => void;
  onClearFolder: () => void;
  onAnalyzeDirectories: () => void; // NEW: Stages 1-2
  onSearchFiles: () => void; // NEW: Stage 3 only
  onCancelSearch: () => void;
}

export const RenameControlsPanel: React.FC<IRenameControlsPanelProps> = ({
  selectedFolder,
  searchingFiles,
  searchProgress,
  loading,
  onSelectFolder,
  onClearFolder,
  onAnalyzeDirectories,
  onSearchFiles,
  onCancelSearch
}) => {

  /**
   * Get current stage information
   */
  const getCurrentStageInfo = () => {
    return SEARCH_STAGES[searchProgress.currentStage] || SEARCH_STAGES[SearchStage.IDLE];
  };

  /**
   * Format time remaining
   */
  const formatTimeRemaining = (seconds: number): string => {
    if (seconds < 60) {
      return `${Math.round(seconds)}s`;
    } else if (seconds < 3600) {
      return `${Math.round(seconds / 60)}m ${Math.round(seconds % 60)}s`;
    } else {
      return `${Math.round(seconds / 3600)}h ${Math.round((seconds % 3600) / 60)}m`;
    }
  };

  /**
   * Get stage-specific status text
   */
  const getStageStatusText = (): string => {
    const stage = searchProgress.currentStage;
    
    switch (stage) {
      case SearchStage.ANALYZING_DIRECTORIES:
        if (searchProgress.directoriesAnalyzed && searchProgress.totalDirectories) {
          return `Analyzed ${searchProgress.directoriesAnalyzed}/${searchProgress.totalDirectories} directories`;
        }
        return 'Extracting directory structure from data...';
        
      case SearchStage.CHECKING_EXISTENCE:
        if (searchProgress.directoriesChecked && searchProgress.totalDirectories) {
          const existing = searchProgress.existingDirectories || 0;
          const checked = searchProgress.directoriesChecked;
          return `Checked ${checked}/${searchProgress.totalDirectories} directories (${existing} exist)`;
        }
        return 'Verifying directories in SharePoint...';
        
      case SearchStage.SEARCHING_FILES:
        if (searchProgress.filesSearched && searchProgress.totalRows) {
          const found = searchProgress.filesFound || 0;
          const searched = searchProgress.filesSearched;
          return `Searched ${searched}/${searchProgress.totalRows} files (${found} found)`;
        }
        return 'Looking for files in directories...';
        
      case SearchStage.COMPLETED:
        if (searchProgress.filesFound && searchProgress.totalRows) {
          const found = searchProgress.filesFound;
          const total = searchProgress.totalRows;
          const percentage = ((found / total) * 100).toFixed(1);
          return `Search completed: ${found}/${total} files found (${percentage}%)`;
        }
        return 'Analysis completed successfully';
        
      case SearchStage.CANCELLED:
        return 'Operation was cancelled';
        
      case SearchStage.ERROR:
        return searchProgress.errors?.[0] || 'An error occurred';
        
      default:
        return 'Ready to start analysis';
    }
  };

  /**
   * Render stage indicators
   */
  const renderStageIndicators = () => {
    const stages = [
      SearchStage.ANALYZING_DIRECTORIES,
      SearchStage.CHECKING_EXISTENCE,
      SearchStage.SEARCHING_FILES
    ];

    return (
      <div className={styles.stageIndicators}>
        {stages.map((stage, index) => {
          const stageInfo = SEARCH_STAGES[stage];
          const isActive = searchProgress.currentStage === stage;
          const isCompleted = searchProgress.overallProgress > stageInfo.progressMax;
          const isCurrent = isActive;
          
          let stageClass = styles.stageIndicator;
          if (isCompleted) {
            stageClass += ` ${styles.completed}`;
          } else if (isCurrent) {
            stageClass += ` ${styles.active}`;
          } else {
            stageClass += ` ${styles.pending}`;
          }

          return (
            <div key={stage} className={stageClass}>
              <div className={styles.stageNumber}>{index + 1}</div>
              <div className={styles.stageName}>{stageInfo.title}</div>
            </div>
          );
        })}
      </div>
    );
  };

  /**
   * Determine which buttons to show based on current state
   */
  const getButtonState = () => {
    const hasSearchPlan = searchProgress.searchPlan && searchProgress.searchPlan.totalDirectories > 0;
    const isAnalysisComplete = searchProgress.currentStage === SearchStage.CHECKING_EXISTENCE || 
                              (searchProgress.searchPlan && searchProgress.existingDirectories !== undefined);
    
    return {
      showAnalyzeButton: !hasSearchPlan && !searchingFiles,
      showSearchButton: hasSearchPlan && isAnalysisComplete && !searchingFiles,
      showCancelButton: searchingFiles,
      analyzeButtonText: searchingFiles ? 'Analyzing...' : '🔍 Analyze Directories',
      searchButtonText: searchingFiles ? 'Searching...' : '🔍 Search Files'
    };
  };

  /**
   * Render detailed progress information
   */
  const renderDetailedProgress = () => {
    if (!searchingFiles || searchProgress.currentStage === SearchStage.IDLE) {
      return null;
    }

    const stageInfo = getCurrentStageInfo();

    return (
      <div className={styles.searchProgressInfo}>
        {/* Stage indicators */}
        {renderStageIndicators()}

        {/* Current stage title and description */}
        <div className={styles.stageHeader}>
          <h4 className={styles.stageTitle}>
            {stageInfo.title}
            {searchProgress.currentStage !== SearchStage.COMPLETED && 
             searchProgress.currentStage !== SearchStage.ERROR && 
             searchProgress.currentStage !== SearchStage.CANCELLED && (
              <span className={styles.stageProgress}>
                ({searchProgress.stageProgress.toFixed(0)}%)
              </span>
            )}
          </h4>
          <p className={styles.stageDescription}>{stageInfo.description}</p>
        </div>

        {/* Progress bar */}
        <div className={styles.progressContainer}>
          <div className={styles.progressText}>
            <strong className={styles.overallProgress}>
              Overall Progress: {searchProgress.overallProgress.toFixed(0)}%
            </strong>
            {searchProgress.estimatedTimeRemaining && searchProgress.estimatedTimeRemaining > 0 && (
              <span className={styles.timeRemaining}>
                ~{formatTimeRemaining(searchProgress.estimatedTimeRemaining)} remaining
              </span>
            )}
          </div>
          
          <div className={styles.progressBar}>
            <div 
              className={styles.progressFill}
              style={{ 
                width: `${searchProgress.overallProgress}%`,
                backgroundColor: searchProgress.currentStage === SearchStage.ERROR ? '#d13438' : 
                                searchProgress.currentStage === SearchStage.COMPLETED ? '#107c10' : '#0078d4'
              }}
            />
          </div>
          
          {/* Stage-specific progress bar */}
          <div className={styles.stageProgressContainer}>
            <div className={styles.stageProgressLabel}>
              Stage Progress: {searchProgress.stageProgress.toFixed(0)}%
            </div>
            <div className={styles.stageProgressBar}>
              <div 
                className={styles.stageProgressFill}
                style={{ width: `${searchProgress.stageProgress}%` }}
              />
            </div>
          </div>
        </div>

        {/* Current operation details */}
        <div className={styles.operationDetails}>
          {searchProgress.currentFileName && (
            <div className={styles.currentOperation}>
              <strong>Current:</strong> {searchProgress.currentFileName}
            </div>
          )}
          
          {searchProgress.currentDirectory && (
            <div className={styles.currentDirectory}>
              <strong>Directory:</strong> {searchProgress.currentDirectory}
            </div>
          )}
          
          {searchProgress.currentRow > 0 && searchProgress.totalRows > 0 && (
            <div className={styles.rowProgress}>
              <strong>Progress:</strong> {searchProgress.currentRow}/{searchProgress.totalRows} items
            </div>
          )}
        </div>

        {/* Statistics */}
        {(searchProgress.currentStage === SearchStage.SEARCHING_FILES || 
          searchProgress.currentStage === SearchStage.COMPLETED) && (
          <div className={styles.searchStats}>
            {searchProgress.filesSearched !== undefined && (
              <div className={styles.stat}>
                <span className={styles.statLabel}>Files Searched:</span>
                <span className={styles.statValue}>{searchProgress.filesSearched}</span>
              </div>
            )}
            {searchProgress.filesFound !== undefined && (
              <div className={styles.stat}>
                <span className={styles.statLabel}>Files Found:</span>
                <span className={styles.statValue}>{searchProgress.filesFound}</span>
              </div>
            )}
            {searchProgress.existingDirectories !== undefined && (
              <div className={styles.stat}>
                <span className={styles.statLabel}>Existing Directories:</span>
                <span className={styles.statValue}>{searchProgress.existingDirectories}</span>
              </div>
            )}
            {searchProgress.totalDirectories !== undefined && (
              <div className={styles.stat}>
                <span className={styles.statLabel}>Total Directories:</span>
                <span className={styles.statValue}>{searchProgress.totalDirectories}</span>
              </div>
            )}
          </div>
        )}

        {/* Errors and warnings */}
        {searchProgress.errors && searchProgress.errors.length > 0 && (
          <div className={styles.errorList}>
            <h5>Errors:</h5>
            {searchProgress.errors.map((error, index) => (
              <div key={index} className={styles.errorItem}>⚠️ {error}</div>
            ))}
          </div>
        )}

        {searchProgress.warnings && searchProgress.warnings.length > 0 && (
          <div className={styles.warningList}>
            <h5>Warnings:</h5>
            {searchProgress.warnings.map((warning, index) => (
              <div key={index} className={styles.warningItem}>⚠️ {warning}</div>
            ))}
          </div>
        )}
      </div>
    );
  };

  const buttonState = getButtonState();

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

      {/* Rename Files Controls with TWO BUTTONS */}
      <div className={styles.renameControls}>
        <div className={styles.renameButtons}>
          {/* BUTTON 1: Analyze Directories (Stages 1-2) */}
          {buttonState.showAnalyzeButton && (
            <button
              className={styles.renameButton}
              onClick={onAnalyzeDirectories}
              disabled={loading || searchingFiles || !selectedFolder}
            >
              {searchingFiles ? (
                <>
                  <span className={styles.spinner} />
                  Analyzing...
                </>
              ) : (
                <>
                  🔍 Analyze Directories
                </>
              )}
            </button>
          )}

          {/* BUTTON 2: Search Files (Stage 3) */}
          {buttonState.showSearchButton && (
            <button
              className={styles.renameButton}
              onClick={onSearchFiles}
              disabled={loading || searchingFiles}
            >
              {searchingFiles ? (
                <>
                  <span className={styles.spinner} />
                  Searching...
                </>
              ) : (
                <>
                  🔍 Search Files
                </>
              )}
            </button>
          )}
          
          {/* Cancel button */}
          {buttonState.showCancelButton && (
            <button
              className={styles.cancelButton}
              onClick={onCancelSearch}
            >
              ❌ Cancel
            </button>
          )}
        </div>
        
        {/* Main status text */}
        <div className={styles.searchStatus}>
          <div className={styles.searchStatusText}>
            {getStageStatusText()}
          </div>
        </div>
        
        {/* Detailed progress */}
        {renderDetailedProgress()}
        
        {/* Helper text */}
        {!selectedFolder && !searchingFiles && (
          <div className={styles.searchNote}>
            <small>Please select a SharePoint folder first</small>
          </div>
        )}

        {/* Show analysis results summary */}
        {searchProgress.searchPlan && !searchingFiles && (
          <div className={styles.searchNote}>
            <small>
              Analysis complete: {searchProgress.searchPlan.totalDirectories} directories found, 
              {' '}{searchProgress.existingDirectories || 0} exist in SharePoint.
              {buttonState.showSearchButton && ' Ready to search files.'}
            </small>
          </div>
        )}
      </div>
    </>
  );
};