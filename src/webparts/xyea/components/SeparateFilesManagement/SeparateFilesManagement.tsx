// src/webparts/xyea/components/SeparateFilesManagement/SeparateFilesManagement.tsx

import * as React from 'react';
import styles from './SeparateFilesManagement.module.scss';
import { IXyeaProps } from '../IXyeaProps';
import { 
  IExcelFile, 
  IExcelSheet, 
  IExcelColumn, 
  IFilterState, 
  IUploadProgress, 
  IExportSettings,
  ISeparateFilesState,
  UploadStage,
  ExportFormat
} from '../../interfaces/ExcelInterfaces';
import { ExcelParserService } from '../../services/ExcelParserService';
import { ExcelFilterService } from '../../services/ExcelFilterService';
import { ExcelExportService } from '../../services/ExcelExportService';
import FileUploader from './FileUploader';
import ExcelDataTable from './ExcelDataTable';

export interface ISeparateFilesManagementProps {
  context: IXyeaProps['context'];
  userDisplayName: string;
}

export default class SeparateFilesManagement extends React.Component<ISeparateFilesManagementProps, ISeparateFilesState> {
  
  constructor(props: ISeparateFilesManagementProps) {
    super(props);
    
    this.state = {
      uploadedFile: null,
      currentSheet: null,
      columns: [],
      filterState: {
        filters: {},
        totalRows: 0,
        filteredRows: 0,
        isAnyFilterActive: false
      },
      loading: false,
      error: null,
      uploadProgress: {
        stage: UploadStage.IDLE,
        progress: 0,
        message: '',
        isComplete: false,
        hasError: false
      },
      currentPage: 1,
      pageSize: 50,
      userPreferences: {
        autoDetectDataTypes: true,
        showRowNumbers: true,
        pageSize: 50,
        saveFilterState: true,
        defaultExportFormat: ExportFormat.XLSX,
        maxFileSize: 10
      },
      isExporting: false,
      exportSettings: {
        fileName: '',
        includeHeaders: true,
        onlyVisibleColumns: true,
        fileFormat: ExportFormat.XLSX
      }
    };
  }

  private handleFileSelect = async (file: File): Promise<void> => {
    this.setState({ 
      loading: true, 
      error: null,
      uploadedFile: null,
      currentSheet: null,
      columns: []
    });

    try {
      // Парсинг файла с прогрессом
      const result = await ExcelParserService.parseFile(file, this.handleUploadProgress);

      if (!result.success) {
        this.setState({
          loading: false,
          error: result.error || 'Failed to parse file'
        });
        return;
      }

      const excelFile = result.file!;
      const firstSheet = excelFile.sheets[0];

      if (!firstSheet || !firstSheet.isValid) {
        this.setState({
          loading: false,
          error: 'No valid sheets found in the file'
        });
        return;
      }

      // Анализируем колонки
      const columns = ExcelFilterService.analyzeColumns(firstSheet);
      
      // Создаем начальное состояние фильтров
      const filterState = ExcelFilterService.createInitialFilterState(columns, firstSheet.totalRows);

      // Применяем фильтры (изначально все видимо)
      const { filteredSheet } = ExcelFilterService.applyFilters(firstSheet, filterState);

      // Создаем настройки экспорта по умолчанию
      const exportSettings = ExcelExportService.createDefaultExportSettings(file.name);

      this.setState({
        uploadedFile: excelFile,
        currentSheet: filteredSheet,
        columns,
        filterState: {
          ...filterState,
          filteredRows: filteredSheet.totalRows
        },
        exportSettings,
        loading: false
      });

      console.log('[SeparateFilesManagement] File processed successfully:', {
        fileName: file.name,
        sheets: excelFile.sheets.length,
        rows: firstSheet.totalRows,
        columns: columns.length
      });

    } catch (error) {
      console.error('[SeparateFilesManagement] File processing failed:', error);
      this.setState({
        loading: false,
        error: error.message || 'Failed to process file'
      });
    }
  }

  private handleUploadProgress = (stage: UploadStage, progress: number, message: string): void => {
    this.setState({
      uploadProgress: {
        stage,
        progress,
        message,
        isComplete: stage === UploadStage.COMPLETE,
        hasError: stage === UploadStage.ERROR
      }
    });
  }

  private handleFilterChange = (columnName: string, selectedValues: any[]): void => {
    const { filterState, currentSheet } = this.state;

    if (!currentSheet) return;

    // Обновляем состояние фильтра
    const updatedFilterState = ExcelFilterService.updateColumnFilter(
      filterState,
      columnName,
      selectedValues
    );

    // Применяем фильтры к данным
    const { filteredSheet, statistics } = ExcelFilterService.applyFilters(
      currentSheet,
      updatedFilterState
    );

    this.setState({
      filterState: {
        ...updatedFilterState,
        filteredRows: statistics.visible
      },
      currentSheet: filteredSheet,
      currentPage: 1 // Сброс на первую страницу
    });

    console.log('[SeparateFilesManagement] Filter applied:', {
      column: columnName,
      selectedCount: selectedValues.length,
      visibleRows: statistics.visible,
      hiddenRows: statistics.hidden
    });
  }

  private handleClearFilters = (): void => {
    const { columns, uploadedFile } = this.state;

    if (!uploadedFile || !columns.length) return;

    const originalSheet = uploadedFile.sheets[0];
    const clearedFilterState = ExcelFilterService.clearAllFilters(this.state.filterState, columns);

    // Применяем очищенные фильтры
    const { filteredSheet } = ExcelFilterService.applyFilters(originalSheet, clearedFilterState);

    this.setState({
      filterState: {
        ...clearedFilterState,
        filteredRows: originalSheet.totalRows
      },
      currentSheet: filteredSheet,
      currentPage: 1
    });

    console.log('[SeparateFilesManagement] All filters cleared');
  }

  private handleExport = async (): Promise<void> => {
    const { uploadedFile, currentSheet, filterState, exportSettings } = this.state;

    if (!uploadedFile || !currentSheet) {
      this.setState({ error: 'No data to export' });
      return;
    }

    this.setState({ isExporting: true, error: null });

    try {
      const result = await ExcelExportService.exportFilteredData(
        uploadedFile.name,
        currentSheet,
        filterState,
        exportSettings
      );

      if (!result.success) {
        this.setState({ error: result.error || 'Export failed' });
      } else {
        console.log('[SeparateFilesManagement] Export completed:', result.fileName);
        // Можно добавить уведомление об успешном экспорте
      }

    } catch (error) {
      console.error('[SeparateFilesManagement] Export failed:', error);
      this.setState({ error: error.message || 'Export failed' });
    } finally {
      this.setState({ isExporting: false });
    }
  }

  private handleExportSettingsChange = (settings: Partial<IExportSettings>): void => {
    this.setState({
      exportSettings: {
        ...this.state.exportSettings,
        ...settings
      }
    });
  }

  private renderFileUploader = (): React.ReactNode => {
    const { loading, uploadProgress, userPreferences } = this.state;

    return (
      <FileUploader
        onFileSelect={this.handleFileSelect}
        loading={loading}
        progress={uploadProgress}
        disabled={loading}
        acceptedFormats={['.xlsx', '.xls', '.csv']}
        maxFileSize={userPreferences.maxFileSize}
      />
    );
  }

  private renderDataTable = (): React.ReactNode => {
    const { currentSheet, columns, filterState, loading } = this.state;

    if (!currentSheet || !columns.length) {
      return null;
    }

    return (
      <ExcelDataTable
        sheet={currentSheet}
        columns={columns}
        filterState={filterState}
        onFilterChange={this.handleFilterChange}
        onClearFilters={this.handleClearFilters}
        loading={loading}
        pageSize={this.state.userPreferences.pageSize}
      />
    );
  }

  private renderExportControls = (): React.ReactNode => {
    const { currentSheet, filterState, exportSettings, isExporting } = this.state;

    if (!currentSheet) {
      return null;
    }

    const statistics = ExcelExportService.getExportStatistics(currentSheet, filterState);

    return (
      <div className={styles.exportControls}>
        <div className={styles.exportInfo}>
          <h4>📤 Export Data</h4>
          <div className={styles.exportStats}>
            <span>Rows to export: <strong>{statistics.visibleRows}</strong></span>
            <span>Estimated size: <strong>{statistics.estimatedFileSize}</strong></span>
            {statistics.activeFilters > 0 && (
              <span>Active filters: <strong>{statistics.activeFilters}</strong></span>
            )}
          </div>
        </div>

        <div className={styles.exportSettings}>
          <div className={styles.settingGroup}>
            <label className={styles.settingLabel}>
              File name:
              <input
                type="text"
                className={styles.settingInput}
                value={exportSettings.fileName}
                onChange={(e) => this.handleExportSettingsChange({ fileName: e.target.value })}
                disabled={isExporting}
                placeholder="Enter file name"
              />
            </label>
          </div>

          <div className={styles.settingGroup}>
            <label className={styles.settingLabel}>
              Format:
              <select
                className={styles.settingSelect}
                value={exportSettings.fileFormat}
                onChange={(e) => this.handleExportSettingsChange({ 
                  fileFormat: e.target.value as ExportFormat 
                })}
                disabled={isExporting}
              >
                <option value={ExportFormat.XLSX}>Excel (.xlsx)</option>
                <option value={ExportFormat.CSV}>CSV (.csv)</option>
              </select>
            </label>
          </div>

          <div className={styles.settingGroup}>
            <label className={styles.checkboxLabel}>
              <input
                type="checkbox"
                checked={exportSettings.includeHeaders}
                onChange={(e) => this.handleExportSettingsChange({ 
                  includeHeaders: e.target.checked 
                })}
                disabled={isExporting}
              />
              Include headers
            </label>
          </div>
        </div>

        <button
          className={styles.exportButton}
          onClick={this.handleExport}
          disabled={!statistics.canExport || isExporting}
        >
          {isExporting ? 'Exporting...' : '📥 Export Filtered Data'}
        </button>
      </div>
    );
  }

  public render(): React.ReactElement<ISeparateFilesManagementProps> {
    const { loading, error, uploadedFile } = this.state;

    return (
      <section className={styles.separateFilesManagement}>
        <div className={styles.container}>
          {error && (
            <div className={styles.error}>
              <strong>Error:</strong> {error}
              <button 
                className={styles.clearErrorButton}
                onClick={() => this.setState({ error: null })}
              >
                ✕
              </button>
            </div>
          )}

          {!uploadedFile ? (
            <div className={styles.uploadSection}>
              <h2 className={styles.sectionTitle}>📂 Upload Excel File</h2>
              <p className={styles.sectionDescription}>
                Upload an Excel file to view, filter, and export data. The first row should contain column headers.
              </p>
              {this.renderFileUploader()}
            </div>
          ) : (
            <div className={styles.dataSection}>
              <div className={styles.sectionHeader}>
                <h2 className={styles.sectionTitle}>
                  📊 {uploadedFile.name}
                </h2>
                <button
                  className={styles.newFileButton}
                  onClick={() => this.setState({
                    uploadedFile: null,
                    currentSheet: null,
                    columns: [],
                    error: null
                  })}
                  disabled={loading}
                >
                  📂 Load New File
                </button>
              </div>

              {this.renderDataTable()}
              {this.renderExportControls()}
            </div>
          )}
        </div>
      </section>
    );
  }
}