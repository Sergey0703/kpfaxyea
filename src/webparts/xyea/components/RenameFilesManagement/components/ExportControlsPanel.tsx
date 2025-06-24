// src/webparts/xyea/components/RenameFilesManagement/components/ExportControlsPanel.tsx

import * as React from 'react';
import styles from './ExportControlsPanel.module.scss';
import { 
  IRenameExportSettings,
  IRenameExportStatistics
} from '../types/RenameFilesTypes';

export interface IExportControlsPanelProps {
  statistics: IRenameExportStatistics;
  exportSettings: IRenameExportSettings;
  isExporting: boolean;
  onExportSettingsChange: (settings: Partial<IRenameExportSettings>) => void;
  onExport: () => void;
}

export interface IExportControlsPanelState {
  showAdvancedOptions: boolean;
}

export class ExportControlsPanel extends React.Component<IExportControlsPanelProps, IExportControlsPanelState> {

  constructor(props: IExportControlsPanelProps) {
    super(props);
    
    this.state = {
      showAdvancedOptions: false
    };
  }

  private handleFileNameChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.props.onExportSettingsChange({
      fileName: event.target.value
    });
  }

  private handleFileFormatChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    this.props.onExportSettingsChange({
      fileFormat: event.target.value as 'xlsx' | 'csv'
    });
  }

  private handleIncludeHeadersChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.props.onExportSettingsChange({
      includeHeaders: event.target.checked
    });
  }

  private handleIncludeStatusColumnChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.props.onExportSettingsChange({
      includeStatusColumn: event.target.checked
    });
  }

  private handleIncludeTimestampsChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.props.onExportSettingsChange({
      includeTimestamps: event.target.checked
    });
  }

  private handleOnlyCompletedRowsChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    this.props.onExportSettingsChange({
      onlyCompletedRows: event.target.checked
    });
  }

  private toggleAdvancedOptions = (): void => {
    this.setState(prevState => ({
      showAdvancedOptions: !prevState.showAdvancedOptions
    }));
  }

  private renderExportStatistics = (): React.ReactNode => {
    const { statistics } = this.props;

    return (
      <div className={styles.exportStats}>
        <div className={styles.stat}>
          <span className={styles.statLabel}>Total Rows:</span>
          <span className={styles.statValue}>{statistics.totalRows}</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statLabel}>Exportable:</span>
          <span className={styles.statValue}>{statistics.exportableRows}</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statLabel}>Files Found:</span>
          <span className={styles.statValue}>{statistics.foundFiles}</span>
        </div>
        <div className={styles.stat}>
          <span className={styles.statLabel}>Files Not Found:</span>
          <span className={styles.statValue}>{statistics.notFoundFiles}</span>
        </div>
        {statistics.renamedFiles > 0 && (
          <div className={styles.stat}>
            <span className={styles.statLabel}>Files Renamed:</span>
            <span className={styles.statValue}>{statistics.renamedFiles}</span>
          </div>
        )}
        {statistics.errorFiles > 0 && (
          <div className={styles.stat}>
            <span className={styles.statLabel}>Rename Errors:</span>
            <span className={styles.statValue}>{statistics.errorFiles}</span>
          </div>
        )}
        {statistics.skippedFiles > 0 && (
          <div className={styles.stat}>
            <span className={styles.statLabel}>Files Skipped:</span>
            <span className={styles.statValue}>{statistics.skippedFiles}</span>
          </div>
        )}
        {statistics.searchingFiles > 0 && (
          <div className={styles.stat}>
            <span className={styles.statLabel}>Folders Not Found:</span>
            <span className={styles.statValue}>{statistics.searchingFiles}</span>
          </div>
        )}
        <div className={styles.stat}>
          <span className={styles.statLabel}>File Size:</span>
          <span className={styles.statValue}>{statistics.estimatedFileSize}</span>
        </div>
      </div>
    );
  }

  private renderBasicSettings = (): React.ReactNode => {
    const { exportSettings, isExporting } = this.props;

    return (
      <div className={styles.basicSettings}>
        <div className={styles.settingGroup}>
          <label className={styles.settingLabel}>
            File Name:
            <input
              type="text"
              className={styles.settingInput}
              value={exportSettings.fileName}
              onChange={this.handleFileNameChange}
              disabled={isExporting}
              placeholder="Enter export filename"
            />
          </label>
        </div>

        <div className={styles.settingGroup}>
          <label className={styles.settingLabel}>
            Format:
            <select
              className={styles.settingSelect}
              value={exportSettings.fileFormat}
              onChange={this.handleFileFormatChange}
              disabled={isExporting}
            >
              <option value="xlsx">Excel (.xlsx)</option>
              <option value="csv">CSV (.csv)</option>
            </select>
          </label>
        </div>

        <div className={styles.settingGroup}>
          <label className={styles.checkboxLabel}>
            <input
              type="checkbox"
              checked={exportSettings.includeHeaders}
              onChange={this.handleIncludeHeadersChange}
              disabled={isExporting}
            />
            Include column headers
          </label>
        </div>

        <div className={styles.settingGroup}>
          <label className={styles.checkboxLabel}>
            <input
              type="checkbox"
              checked={exportSettings.includeStatusColumn}
              onChange={this.handleIncludeStatusColumnChange}
              disabled={isExporting}
            />
            Include status column
          </label>
        </div>
      </div>
    );
  }

  private renderAdvancedSettings = (): React.ReactNode => {
    const { exportSettings, isExporting } = this.props;
    const { showAdvancedOptions } = this.state;

    if (!showAdvancedOptions) {
      return null;
    }

    return (
      <div className={styles.advancedSettings}>
        <div className={styles.settingGroup}>
          <label className={styles.checkboxLabel}>
            <input
              type="checkbox"
              checked={exportSettings.includeTimestamps}
              onChange={this.handleIncludeTimestampsChange}
              disabled={isExporting}
            />
            Include export timestamp
          </label>
        </div>

        <div className={styles.settingGroup}>
          <label className={styles.checkboxLabel}>
            <input
              type="checkbox"
              checked={exportSettings.onlyCompletedRows}
              onChange={this.handleOnlyCompletedRowsChange}
              disabled={isExporting}
            />
            Export only completed rows (exclude "folders not found")
          </label>
        </div>
      </div>
    );
  }

  private renderExportButton = (): React.ReactNode => {
    const { statistics, isExporting, onExport } = this.props;

    const buttonText = isExporting 
      ? 'Exporting...' 
      : `📥 Export ${statistics.exportableRows} Rows`;

    return (
      <div className={styles.exportButtonContainer}>
        <button
          className={styles.exportButton}
          onClick={onExport}
          disabled={!statistics.canExport || isExporting}
        >
          {isExporting && <span className={styles.spinner} />}
          {buttonText}
        </button>
        
        {!statistics.canExport && (
          <div className={styles.exportDisabledMessage}>
            No data available for export. Please load a file and search for files first.
          </div>
        )}
      </div>
    );
  }

  public render(): React.ReactElement<IExportControlsPanelProps> {
    const { statistics } = this.props;
    const { showAdvancedOptions } = this.state;

    return (
      <div className={styles.exportControlsPanel}>
        <div className={styles.header}>
          <h4 className={styles.title}>📤 Export Data</h4>
          <p className={styles.description}>
            Export your rename files data with status information to Excel or CSV format.
          </p>
        </div>

        <div className={styles.content}>
          {/* Export Statistics */}
          <div className={styles.statisticsSection}>
            <h5 className={styles.sectionTitle}>Export Statistics</h5>
            {this.renderExportStatistics()}
          </div>

          {/* Export Settings */}
          <div className={styles.settingsSection}>
            <h5 className={styles.sectionTitle}>Export Settings</h5>
            
            {/* Basic Settings */}
            {this.renderBasicSettings()}

            {/* Advanced Options Toggle */}
            <div className={styles.advancedToggle}>
              <button
                className={styles.toggleButton}
                onClick={this.toggleAdvancedOptions}
                type="button"
              >
                {showAdvancedOptions ? '▼' : '▶'} Advanced Options
              </button>
            </div>

            {/* Advanced Settings */}
            {this.renderAdvancedSettings()}
          </div>

          {/* Export Button */}
          <div className={styles.exportSection}>
            {this.renderExportButton()}
          </div>
        </div>

        {/* Export Information */}
        <div className={styles.exportInfo}>
          <div className={styles.infoItem}>
            <span className={styles.infoIcon}>ℹ️</span>
            <span className={styles.infoText}>
              The exported file will include all your data columns plus a status column showing 
              whether each file was found, renamed, or had errors.
            </span>
          </div>
          
          {statistics.searchingFiles > 0 && (
            <div className={styles.infoItem}>
              <span className={styles.infoIcon}>⚠️</span>
              <span className={styles.infoText}>
                {statistics.searchingFiles} files have "Folder not found" status. 
                You can export now or wait for a complete search.
              </span>
            </div>
          )}

          {statistics.skippedFiles > 0 && (
            <div className={styles.infoItem}>
              <span className={styles.infoIcon}>⏭️</span>
              <span className={styles.infoText}>
                {statistics.skippedFiles} files were skipped because target files already exist.
              </span>
            </div>
          )}
        </div>
      </div>
    );
  }
}