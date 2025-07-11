// src/webparts/xyea/components/RenameFilesManagement/components/RenameControlsPanel.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { 
  ISharePointFolder, 
  ISearchProgress, 
  SearchStage, 
  SEARCH_STAGES,
  ITimeTrackingState,
  TimeTrackingHelper
} from '../types/RenameFilesTypes';

export interface IRenameControlsPanelProps {
  selectedFolder: ISharePointFolder | undefined;
  searchingFiles: boolean;
  searchProgress: ISearchProgress;
  loading: boolean;
  foundFilesCount: number;
  isRenaming: boolean;
  renameProgress?: {
    current: number;
    total: number;
    fileName: string;
    success: number;
    errors: number;
    skipped: number;
  };
  timeTracking: ITimeTrackingState; // NEW: Time tracking state
  onSelectFolder: () => void;
  onClearFolder: () => void;
  onAnalyzeDirectories: () => void;
  onSearchFiles: () => void;
  onCancelSearch: () => void;
  onRenameFiles: () => void;
  onCancelRename: () => void;
}

export const RenameControlsPanel: React.FC<IRenameControlsPanelProps> = ({
  selectedFolder,
  searchingFiles,
  searchProgress,
  loading,
  foundFilesCount,
  isRenaming,
  renameProgress,
  timeTracking, // NEW: Time tracking prop
  onSelectFolder,
  onClearFolder,
  onAnalyzeDirectories,
  onSearchFiles,
  onCancelSearch,
  onRenameFiles,
  onCancelRename
}) => {

  /**
   * Get current stage information
   */
  const getCurrentStageInfo = (): typeof SEARCH_STAGES[SearchStage] => {
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
   * Get stage-specific status text - UPDATED with new status texts
   */
  const getStageStatusText = (): string => {
    if (isRenaming && renameProgress) {
      return `Renaming files: ${renameProgress.current}/${renameProgress.total} (${renameProgress.success} renamed, ${renameProgress.errors} errors, ${renameProgress.skipped} skipped)`;
    }

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
   * NEW: Render compact current operation timer
   */
  const renderCurrentTimer = (): React.ReactNode => {
    if (!timeTracking.currentTimer || !timeTracking.currentTimer.isActive) {
      return null;
    }

    const timer = timeTracking.currentTimer;
    const formattedTime = TimeTrackingHelper.formatCompactTime(timer.elapsedSeconds);
    
    return (
      <div className={styles.currentTimer}>
        <span className={styles.timerIcon}>⏱️</span>
        <span className={styles.timerText}>{formattedTime}</span>
        <span className={styles.timerDescription}>{timer.operationDescription}</span>
      </div>
    );
  };

  /**
   * NEW: Render completed operations time badges (compact)
   */
  const renderCompletedOperations = (): React.ReactNode => {
    if (!timeTracking.showTimingInfo || timeTracking.completedOperations.length === 0) {
      return null;
    }

    // Show only the last 3 completed operations to save space
    const recentOperations = timeTracking.completedOperations
      .slice(-3)
      .reverse(); // Show newest first

    return (
      <div className={styles.completedOperations}>
        <span className={styles.completedLabel}>Recent:</span>
        {recentOperations.map((operation, index) => {
          const formattedTime = TimeTrackingHelper.formatCompactTime(operation.durationSeconds);
          const operationIcon = operation.operationType === 'analyze' ? '🔍' : 
                              operation.operationType === 'search' ? '🔎' : '🏷️';
          const statusIcon = operation.success ? '✅' : '❌';
          
          return (
            <div 
              key={`${operation.operationType}-${operation.startTime.getTime()}`}
              className={`${styles.completedOperation} ${operation.success ? styles.success : styles.error}`}
              title={`${operation.operationDescription} - ${operation.success ? 'Success' : 'Failed'} - ${formattedTime}${operation.itemsProcessed ? ` (${operation.itemsProcessed} items)` : ''}`}
            >
              <span className={styles.operationIcon}>{operationIcon}</span>
              <span className={styles.operationTime}>{formattedTime}</span>
              <span className={styles.operationStatus}>{statusIcon}</span>
            </div>
          );
        })}
      </div>
    );
  };

  /**
   * NEW: Render timing info section (compact)
   */
  const renderTimingInfo = (): React.ReactNode => {
    if (!timeTracking.showTimingInfo) {
      return null;
    }

    const hasCurrentTimer = timeTracking.currentTimer && timeTracking.currentTimer.isActive;
    const hasCompletedOperations = timeTracking.completedOperations.length > 0;

    if (!hasCurrentTimer && !hasCompletedOperations) {
      return null;
    }

    return (
      <div className={styles.timingInfo}>
        {renderCurrentTimer()}
        {renderCompletedOperations()}
      </div>
    );
  };

  /**
   * Render stage indicators
   */
  const renderStageIndicators = (): React.ReactNode => {
    if (isRenaming) {
      return null; // Hide stage indicators during rename
    }

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
   * Render rename progress with skipped files support - UPDATED with new status texts
   */
  const renderRenameProgress = (): React.ReactNode => {
    if (!renameProgress) return null;

    const progressPercentage = renameProgress.total > 0 ? (renameProgress.current / renameProgress.total) * 100 : 0;

    return (
      <div className={styles.searchProgressInfo}>
        <div className={styles.stageHeader}>
          <h4 className={styles.stageTitle}>
            🏷️ Renaming Files
            <span className={styles.stageProgress}>
              ({progressPercentage.toFixed(0)}%)
            </span>
          </h4>
          <p className={styles.stageDescription}>Adding staffID prefix to filenames in SharePoint...</p>
        </div>

        <div className={styles.progressContainer}>
          <div className={styles.progressText}>
            <strong className={styles.overallProgress}>
              Progress: {renameProgress.current}/{renameProgress.total} files
            </strong>
          </div>
          
          <div className={styles.progressBar}>
            <div 
              className={styles.progressFill}
              style={{ 
                width: `${progressPercentage}%`,
                backgroundColor: '#28a745'
              }}
            />
          </div>
        </div>

        <div className={styles.operationDetails}>
          <div className={styles.currentOperation}>
            <strong>Current File:</strong> {renameProgress.fileName}
          </div>
        </div>

        <div className={styles.searchStats}>
          <div className={styles.stat}>
            <span className={styles.statLabel}>Files Renamed:</span>
            <span className={styles.statValue}>{renameProgress.success}</span>
          </div>
          <div className={styles.stat}>
            <span className={styles.statLabel}>Rename Errors:</span>
            <span className={styles.statValue}>{renameProgress.errors}</span>
          </div>
          <div className={styles.stat}>
            <span className={styles.statLabel}>Files Skipped:</span>
            <span className={styles.statValue}>{renameProgress.skipped}</span>
          </div>
          <div className={styles.stat}>
            <span className={styles.statLabel}>Remaining:</span>
            <span className={styles.statValue}>{renameProgress.total - renameProgress.current}</span>
          </div>
        </div>
      </div>
    );
  };

  /**
   * Determine which buttons to show based on current state
   */
  const getButtonState = (): {
    showAnalyzeButton: boolean;
    showSearchButton: boolean;
    showRenameButton: boolean;
    showCancelButton: boolean;
    analyzeButtonText: string;
    searchButtonText: string;
    renameButtonText: string;
    cancelButtonText: string;
  } => {
    const hasSearchPlan = Boolean(searchProgress.searchPlan && searchProgress.searchPlan.totalDirectories > 0);
    const isAnalysisComplete = searchProgress.currentStage === SearchStage.CHECKING_EXISTENCE || 
                              Boolean(searchProgress.searchPlan && typeof searchProgress.existingDirectories === 'number');
    const isSearchComplete = searchProgress.currentStage === SearchStage.COMPLETED && foundFilesCount > 0;
    
    return {
      showAnalyzeButton: !hasSearchPlan && !searchingFiles && !isRenaming,
      showSearchButton: hasSearchPlan && isAnalysisComplete && !searchingFiles && !isRenaming,
      showRenameButton: isSearchComplete && !searchingFiles && !isRenaming,
      showCancelButton: searchingFiles || isRenaming,
      analyzeButtonText: searchingFiles ? 'Analyzing...' : '🔍 Analyze Directories',
      searchButtonText: searchingFiles ? 'Searching...' : '🔍 Search Files',
      renameButtonText: isRenaming ? 'Renaming...' : `🏷️ Rename ${foundFilesCount} Files`,
      cancelButtonText: isRenaming ? '❌ Cancel Rename' : '❌ Cancel Search'
    };
  };

  /**
   * Render detailed progress information
   */
  const renderDetailedProgress = (): React.ReactNode => {
    if ((!searchingFiles && !isRenaming) || searchProgress.currentStage === SearchStage.IDLE) {
      return null;
    }

    // Rename progress has priority over search progress
    if (isRenaming && renameProgress) {
      return renderRenameProgress();
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
            disabled={loading || searchingFiles || isRenaming}
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
                disabled={searchingFiles || isRenaming}
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

      {/* Rename Files Controls with THREE BUTTONS */}
      <div className={styles.renameControls}>
        {/* NEW: Timing info section (compact) */}
        {renderTimingInfo()}

        <div className={styles.renameButtons}>
          {/* BUTTON 1: Analyze Directories (Stages 1-2) */}
          {buttonState.showAnalyzeButton && (
            <button
              className={styles.renameButton}
              onClick={onAnalyzeDirectories}
              disabled={loading || searchingFiles || !selectedFolder || isRenaming}
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
              disabled={loading || searchingFiles || isRenaming}
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

          {/* BUTTON 3: Rename Files */}
          {buttonState.showRenameButton && (
            <button
              className={styles.renameFilesButton}
              onClick={onRenameFiles}
              disabled={loading || searchingFiles || isRenaming}
            >
              {isRenaming ? (
                <>
                  <span className={styles.spinner} />
                  Renaming...
                </>
              ) : (
                <>
                  🏷️ Rename {foundFilesCount} Files
                </>
              )}
            </button>
          )}
          
          {/* Cancel button */}
          {buttonState.showCancelButton && (
            <button
              className={styles.cancelButton}
              onClick={isRenaming ? onCancelRename : onCancelSearch}
            >
              {buttonState.cancelButtonText}
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
        {!selectedFolder && !searchingFiles && !isRenaming && (
          <div className={styles.searchNote}>
            <small>Please select a SharePoint folder first</small>
          </div>
        )}

        {/* Show analysis results summary */}
        {searchProgress.searchPlan && !searchingFiles && !isRenaming && (
          <div className={styles.searchNote}>
            <small>
              Analysis complete: {searchProgress.searchPlan.totalDirectories} directories found, 
              {' '}{searchProgress.existingDirectories || 0} exist in SharePoint.
              {buttonState.showSearchButton && ' Ready to search files.'}
              {buttonState.showRenameButton && ` Found ${foundFilesCount} files ready for renaming.`}
            </small>
          </div>
        )}
      </div>
    </>
  );
};