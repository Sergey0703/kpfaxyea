// src/webparts/xyea/components/ConvertFilesPropsTable/ConvertFilesPropsTable.tsx - Updated with ConvertType support

import * as React from 'react';
import styles from './ConvertFilesPropsTable.module.scss';
import { IConvertFilesPropsTableProps } from './IConvertFilesPropsTableProps';
import { IConvertFileProps } from '../../models';
import { PriorityHelper } from '../../utils';
import ExcelImportButton, { IExcelImportData } from '../ExcelImportButton/ExcelImportButton';
import { ConfirmationDialog, IConfirmationDialogConfig } from '../ConfirmationDialog';
import * as XLSX from 'xlsx';

export interface IConvertFilesPropsTableState {
  error: string | undefined;
  // Conversion dialog state
  showConversionDialog: boolean;
  conversionDialogConfig: IConfirmationDialogConfig;
  conversionDialogLoading: boolean;
  // Column validation state
  columnValidation: {
    [itemId: number]: {
      propValid: boolean;
      prop2Valid: boolean;
      propExists?: boolean;
      prop2Exists?: boolean;
    }
  };
  // Files content cache
  filesCache: {
    exportFile?: { headers: string[]; data: string[][] };
    importFile?: { headers: string[]; data: string[][] };
  };
}

export default class ConvertFilesPropsTable extends React.Component<IConvertFilesPropsTableProps, IConvertFilesPropsTableState> {
  
  constructor(props: IConvertFilesPropsTableProps) {
    super(props);
    this.state = {
      error: undefined,
      showConversionDialog: false,
      conversionDialogConfig: {
        title: '',
        message: '',
        confirmText: 'Continue',
        type: 'info'
      },
      conversionDialogLoading: false,
      columnValidation: {},
      filesCache: {}
    };
  }

  public componentDidUpdate(prevProps: IConvertFilesPropsTableProps): void {
    // Re-validate columns when files change or items change
    if (prevProps.selectedFiles !== this.props.selectedFiles || 
        prevProps.items !== this.props.items) {
      this.validateAllColumns().catch(error => {
        console.error('[ConvertFilesPropsTable] Validation error in componentDidUpdate:', error);
      });
    }
  }

  private handleEdit = (item: IConvertFileProps): void => {
    this.props.onEdit(item);
  }

  private handleDelete = (id: number): void => {
    if (confirm('Are you sure you want to mark this item as deleted?')) {
      this.props.onToggleDeleted(id, true);
    }
  }

  private handleRestore = (id: number): void => {
    this.props.onToggleDeleted(id, false);
  }

  private handleMoveUp = (id: number): void => {
    this.props.onMoveUp(id);
  }

  private handleMoveDown = (id: number): void => {
    this.props.onMoveDown(id);
  }

  private handleAdd = (): void => {
    this.props.onAdd(this.props.convertFileId);
  }

  private handleExcelImport = async (data: IExcelImportData[]): Promise<void> => {
    try {
      console.log('[ConvertFilesPropsTable] Starting Excel import:', {
        convertFileId: this.props.convertFileId,
        dataCount: data.length
      });

      // Call the import handler from props
      if (this.props.onImportFromExcel) {
        await this.props.onImportFromExcel(this.props.convertFileId, data);
      } else {
        throw new Error('Excel import is not supported');
      }

      console.log('[ConvertFilesPropsTable] Excel import completed successfully');
    } catch (error) {
      console.error('[ConvertFilesPropsTable] Excel import failed:', error);
      this.setState({ 
        error: error instanceof Error ? error.message : 'Excel import failed' 
      });
      throw error;
    }
  }

  // All conversion-related methods remain the same...
  private handleConversion = (): void => {
    const selectedFiles = this.props.selectedFiles?.[this.props.convertFileId];
    
    if (!selectedFiles?.export && !selectedFiles?.import) {
      this.setState({
        showConversionDialog: true,
        conversionDialogConfig: {
          title: 'Files Required',
          message: 'Both Export file and Import file must be selected before starting conversion.\n\nPlease select the required files and try again.',
          confirmText: 'OK',
          cancelText: '',
          type: 'warning',
          showIcon: true
        }
      });
      return;
    }
    
    if (!selectedFiles?.export) {
      this.setState({
        showConversionDialog: true,
        conversionDialogConfig: {
          title: 'Export File Required',
          message: 'Export file must be selected before starting conversion.\n\nPlease select an Export file and try again.',
          confirmText: 'OK',
          cancelText: '',
          type: 'warning',
          showIcon: true
        }
      });
      return;
    }
    
    if (!selectedFiles?.import) {
      this.setState({
        showConversionDialog: true,
        conversionDialogConfig: {
          title: 'Import File Required',
          message: 'Import file must be selected before starting conversion.\n\nPlease select an Import file and try again.',
          confirmText: 'OK',
          cancelText: '',
          type: 'warning',
          showIcon: true
        }
      });
      return;
    }
    
    this.setState({
      showConversionDialog: true,
      conversionDialogConfig: {
        title: 'Start Conversion Process',
        message: 'Do you want to start the conversion process?\n\nThis will copy data from the Export file to the Import file based on the column mappings defined in this table.',
        confirmText: 'Yes, Start Conversion',
        cancelText: 'Cancel',
        type: 'info',
        showIcon: true
      }
    });
  }

  private handleConversionConfirm = async (): Promise<void> => {
    const { conversionDialogConfig } = this.state;
    
    if (!conversionDialogConfig.cancelText || conversionDialogConfig.cancelText === '') {
      this.setState({
        showConversionDialog: false,
        conversionDialogLoading: false
      });
      return;
    }
    
    this.setState({ conversionDialogLoading: true });

    try {
      const selectedFiles = this.props.selectedFiles?.[this.props.convertFileId];
      
      if (!selectedFiles?.export || !selectedFiles?.import) {
        throw new Error('Files are no longer available');
      }

      await this.performConversion(selectedFiles.export, selectedFiles.import);

      this.setState({
        showConversionDialog: false,
        conversionDialogLoading: false
      });

    } catch (error) {
      console.error('[ConvertFilesPropsTable] Conversion failed:', error);
      this.setState({
        error: error instanceof Error ? error.message : 'Conversion failed',
        conversionDialogLoading: false
      });
    }
  }

  private handleConversionCancel = (): void => {
    this.setState({
      showConversionDialog: false,
      conversionDialogLoading: false
    });
  }

  // All conversion logic methods remain the same...
  private performConversion = async (exportFile: File, importFile: File): Promise<void> => {
    console.log('[ConvertFilesPropsTable] Starting conversion process');

    const exportData = await this.readExcelFile(exportFile);
    const importData = await this.readExcelFile(importFile);

    const processedRows: number[] = [];
    const skippedRows: Array<{ row: number; reason: string; column?: string; file?: string }> = [];

    const sortedItems = this.getSortedItems().filter(item => !item.IsDeleted);
    
    for (let i = 0; i < sortedItems.length; i++) {
      const item = sortedItems[i];
      const rowNumber = i + 1;

      try {
        const sourceColumnIndex = exportData.headers.findIndex(header => 
          header.toLowerCase().trim() === item.Prop.toLowerCase().trim()
        );

        if (sourceColumnIndex === -1) {
          skippedRows.push({
            row: rowNumber,
            reason: `Column "${item.Prop}" not found`,
            column: item.Prop,
            file: 'Export file'
          });
          continue;
        }

        const targetColumnIndex = importData.headers.findIndex(header => 
          header.toLowerCase().trim() === item.Prop2.toLowerCase().trim()
        );

        if (targetColumnIndex === -1) {
          skippedRows.push({
            row: rowNumber,
            reason: `Column "${item.Prop2}" not found`,
            column: item.Prop2,
            file: 'Import file'
          });
          continue;
        }

        const exportDataRows = exportData.data.slice(1);
        
        for (let dataRowIndex = 0; dataRowIndex < exportDataRows.length; dataRowIndex++) {
          const sourceValue = exportDataRows[dataRowIndex][sourceColumnIndex];
          const importRowIndex = dataRowIndex + 1;
          
          if (importRowIndex >= importData.data.length) {
            const newRow = new Array(importData.headers.length).fill('');
            importData.data.push(newRow);
          }

          if (!importData.data[importRowIndex]) {
            importData.data[importRowIndex] = new Array(importData.headers.length).fill('');
          }
          
          while (importData.data[importRowIndex].length <= targetColumnIndex) {
            importData.data[importRowIndex].push('');
          }

          importData.data[importRowIndex][targetColumnIndex] = sourceValue || '';
          
          console.log(`[ConvertFilesPropsTable] Copied "${sourceValue}" from Export[${dataRowIndex + 1}][${sourceColumnIndex}] to Import[${importRowIndex}][${targetColumnIndex}]`);
        }

        processedRows.push(rowNumber);
        console.log(`[ConvertFilesPropsTable] Processed row ${rowNumber}: ${item.Prop} → ${item.Prop2}`);

      } catch (error) {
        console.error(`[ConvertFilesPropsTable] Error processing row ${rowNumber}:`, error);
        skippedRows.push({
          row: rowNumber,
          reason: `Processing error: ${error instanceof Error ? error.message : 'Unknown error'}`
        });
      }
    }

    await this.downloadUpdatedFile(importData, importFile.name);
    this.showConversionResults(processedRows, skippedRows);
  }

  // All file reading and validation methods remain the same...
  private readExcelFile = async (file: File): Promise<{ headers: string[]; data: string[][] }> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (event) => {
        try {
          const arrayBuffer = event.target?.result as ArrayBuffer;
          const workbook = XLSX.read(arrayBuffer, { 
            type: 'array',
            cellDates: true,
            cellStyles: true 
          });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          
          console.log('[ConvertFilesPropsTable] Reading Excel file:', file.name);
          console.log('[ConvertFilesPropsTable] Worksheet range:', worksheet['!ref']);
          
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
          const numCols = range.e.c + 1;
          
          const excelHeaders: string[] = [];
          for (let col = 0; col < numCols; col++) {
            excelHeaders.push(this.getExcelColumnName(col));
          }
          
          console.log('[ConvertFilesPropsTable] Excel column headers:', excelHeaders);
          
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: '',
            raw: false,
            dateNF: 'yyyy-mm-dd'
          }) as string[][];

          console.log('[ConvertFilesPropsTable] Raw JSON data rows:', jsonData.length);

          if (jsonData.length === 0) {
            reject(new Error(`File ${file.name} is empty`));
            return;
          }

          const headers = excelHeaders;
          const data = jsonData;
          
          console.log('[ConvertFilesPropsTable] Final headers (Excel columns):', headers);
          console.log('[ConvertFilesPropsTable] Data rows count:', data.length);

          resolve({ headers, data });
        } catch (error) {
          console.error('[ConvertFilesPropsTable] Error reading Excel file:', error);
          reject(new Error(`Failed to read ${file.name}: ${error instanceof Error ? error.message : 'Unknown error'}`));
        }
      };

      reader.onerror = () => reject(new Error(`Failed to read file ${file.name}`));
      reader.readAsArrayBuffer(file);
    });
  }

  private getExcelColumnName = (columnIndex: number): string => {
    let columnName = '';
    while (columnIndex >= 0) {
      columnName = String.fromCharCode((columnIndex % 26) + 65) + columnName;
      columnIndex = Math.floor(columnIndex / 26) - 1;
    }
    return columnName;
  }

  private downloadUpdatedFile = async (data: { headers: string[]; data: string[][] }, originalFileName: string): Promise<void> => {
    const workbook = XLSX.utils.book_new();
    const worksheetData = data.data;
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const fileName = originalFileName.replace(/\.[^/.]+$/, '_converted.xlsx');
    XLSX.writeFile(workbook, fileName);
  }

  private showConversionResults = (processedRows: number[], skippedRows: Array<{ row: number; reason: string; column?: string; file?: string }>): void => {
    let message = `Conversion completed!\n\n✅ Processed: ${processedRows.length} rows`;
    
    if (skippedRows.length > 0) {
      message += `\n❌ Skipped: ${skippedRows.length} rows\n\n`;
      
      const exportFileErrors = skippedRows.filter(item => item.file === 'Export file');
      const importFileErrors = skippedRows.filter(item => item.file === 'Import file');
      const otherErrors = skippedRows.filter(item => !item.file);
      
      if (exportFileErrors.length > 0) {
        message += 'Missing columns in Export file:\n';
        exportFileErrors.forEach(item => {
          message += `- Row ${item.row}: Column "${item.column}" not found\n`;
        });
        message += '\n';
      }
      
      if (importFileErrors.length > 0) {
        message += 'Missing columns in Import file:\n';
        importFileErrors.forEach(item => {
          message += `- Row ${item.row}: Column "${item.column}" not found\n`;
        });
        message += '\n';
      }
      
      if (otherErrors.length > 0) {
        message += 'Other errors:\n';
        otherErrors.forEach(item => {
          message += `- Row ${item.row}: ${item.reason}\n`;
        });
      }
    }
    
    message += '\nThe updated Import file has been downloaded.';
    alert(message);
  }

  private validateAllColumns = async (): Promise<void> => {
    const selectedFiles = this.props.selectedFiles?.[this.props.convertFileId];
    
    console.log('[ConvertFilesPropsTable] Starting column validation');
    
    try {
      const cachedExport = this.state.filesCache.exportFile;
      const cachedImport = this.state.filesCache.importFile;
      
      const fileDataToProcess: {
        exportFile?: { headers: string[]; data: string[][] };
        importFile?: { headers: string[]; data: string[][] };
      } = {
        exportFile: cachedExport,
        importFile: cachedImport
      };

      if (selectedFiles?.export && !cachedExport) {
        console.log('[ConvertFilesPropsTable] Reading export file:', selectedFiles.export.name);
        fileDataToProcess.exportFile = await this.readExcelFile(selectedFiles.export);
      }
      
      if (selectedFiles?.import && !cachedImport) {
        console.log('[ConvertFilesPropsTable] Reading import file:', selectedFiles.import.name);
        fileDataToProcess.importFile = await this.readExcelFile(selectedFiles.import);
      }

      this.setState({
        filesCache: {
          exportFile: fileDataToProcess.exportFile,
          importFile: fileDataToProcess.importFile
        }
      });

      console.log('[ConvertFilesPropsTable] Available files:', {
        hasExport: !!fileDataToProcess.exportFile,
        hasImport: !!fileDataToProcess.importFile,
        exportHeaders: fileDataToProcess.exportFile?.headers.length || 0,
        importHeaders: fileDataToProcess.importFile?.headers.length || 0
      });

      const validation: {
        [itemId: number]: {
          propValid: boolean;
          prop2Valid: boolean;
          propExists?: boolean;
          prop2Exists?: boolean;
        }
      } = {};
      
      this.props.items.forEach(item => {
        console.log(`[ConvertFilesPropsTable] Validating item ${item.Id}: Prop="${item.Prop}", Prop2="${item.Prop2}"`);
        
        let propExists: boolean | undefined = undefined;
        if (fileDataToProcess.exportFile && item.Prop.trim()) {
          propExists = fileDataToProcess.exportFile.headers.some(header => {
            const headerLower = header.toLowerCase().trim();
            const propLower = item.Prop.toLowerCase().trim();
            const matches = headerLower === propLower;
            if (matches) {
              console.log(`[ConvertFilesPropsTable] Prop match found: "${header}" === "${item.Prop}"`);
            }
            return matches;
          });
        }
        
        let prop2Exists: boolean | undefined = undefined;
        if (fileDataToProcess.importFile && item.Prop2.trim()) {
          prop2Exists = fileDataToProcess.importFile.headers.some(header => {
            const headerLower = header.toLowerCase().trim();
            const prop2Lower = item.Prop2.toLowerCase().trim();
            const matches = headerLower === prop2Lower;
            if (matches) {
              console.log(`[ConvertFilesPropsTable] Prop2 match found: "${header}" === "${item.Prop2}"`);
            }
            return matches;
          });
        }

        console.log(`[ConvertFilesPropsTable] Item ${item.Id} validation result: Prop exists=${propExists}, Prop2 exists=${prop2Exists}`);

        validation[item.Id] = {
          propValid: propExists === true,
          prop2Valid: prop2Exists === true,
          propExists,
          prop2Exists
        };
      });

      console.log('[ConvertFilesPropsTable] Final validation result:', validation);
      this.setState({ columnValidation: validation });

    } catch (error) {
      console.error('[ConvertFilesPropsTable] Error validating columns:', error);
      this.setState({ columnValidation: {}, filesCache: {} });
    }
  }

  private getColumnValidationStyle = (itemId: number, column: 'prop' | 'prop2'): React.CSSProperties => {
    const validation = this.state.columnValidation[itemId];
    const selectedFiles = this.props.selectedFiles?.[this.props.convertFileId];
    
    const item = this.props.items.find(i => i.Id === itemId);
    const fieldValue = column === 'prop' ? item?.Prop : item?.Prop2;
    const isFieldEmpty = !fieldValue || fieldValue.trim() === '';
    
    const hasFile = column === 'prop' ? !!selectedFiles?.export : !!selectedFiles?.import;
    
    console.log(`[getColumnValidationStyle] Item ${itemId}, Column ${column}:`, {
      fieldValue,
      isFieldEmpty,
      hasFile,
      validation,
      exists: validation ? (column === 'prop' ? validation.propExists : validation.prop2Exists) : 'no validation',
      isValid: validation ? (column === 'prop' ? validation.propValid : validation.prop2Valid) : 'no validation'
    });
    
    if (!hasFile || isFieldEmpty) {
      console.log(`[getColumnValidationStyle] → GRAY (no file or empty field)`);
      return { backgroundColor: '#e6e6e6' };
    }
    
    if (!validation) {
      console.log(`[getColumnValidationStyle] → GRAY (no validation data)`);
      return { backgroundColor: '#e6e6e6' };
    }

    const isValid = column === 'prop' ? validation.propValid : validation.prop2Valid;
    const exists = column === 'prop' ? validation.propExists : validation.prop2Exists;
    
    if (exists === undefined) {
      console.log(`[getColumnValidationStyle] → GRAY (exists is undefined)`);
      return { backgroundColor: '#e6e6e6' };
    }
    
    if (isValid) {
      console.log(`[getColumnValidationStyle] → GREEN (valid)`);
      return { backgroundColor: '#d4edda' };
    } else {
      console.log(`[getColumnValidationStyle] → RED (invalid)`);
      return { backgroundColor: '#f8d7da' };
    }
  }

  private canMoveUp = (item: IConvertFileProps): boolean => {
    return PriorityHelper.canMoveUp(this.props.allItems, item.Id, this.props.convertFileId);
  }

  private canMoveDown = (item: IConvertFileProps): boolean => {
    return PriorityHelper.canMoveDown(this.props.allItems, item.Id, this.props.convertFileId);
  }

  private getSortedItems = (): IConvertFileProps[] => {
    return PriorityHelper.sortByPriority(this.props.items);
  }

  private clearError = (): void => {
    this.setState({ error: undefined });
  }

  // Helper method to get convert type name
  private getConvertTypeName = (convertTypeId: number, convertTypes: { Id: number; Title: string }[] = []): string => {
    const convertType = convertTypes.find(ct => ct.Id === convertTypeId);
    return convertType ? convertType.Title : `Type ${convertTypeId}`;
  }

  public render(): React.ReactElement<IConvertFilesPropsTableProps> {
    const { convertFileTitle, loading, convertTypes = [] } = this.props;
    const { error, showConversionDialog, conversionDialogConfig, conversionDialogLoading } = this.state;
    const sortedItems = this.getSortedItems();

    return (
      <div className={styles.convertFilesPropsTable}>
        <div className={styles.header}>
          <h3 className={styles.title}>Properties for: {convertFileTitle}</h3>
          <div className={styles.headerActions}>
            <ExcelImportButton
              onImport={this.handleExcelImport}
              disabled={loading}
              existingItemsCount={sortedItems.length}
            />
            <button 
              className={styles.addButton}
              onClick={this.handleAdd}
              disabled={loading}
            >
              + Add Property
            </button>
            <button 
              onClick={this.handleConversion}
              disabled={loading}
              title="Start conversion process"
              style={{
                backgroundColor: '#107c10',
                color: 'white',
                border: 'none',
                borderRadius: '2px',
                padding: '6px 12px',
                cursor: loading ? 'not-allowed' : 'pointer',
                fontSize: '12px',
                fontWeight: '600',
                marginLeft: '8px',
                transition: 'all 0.2s ease',
                opacity: loading ? 0.6 : 1
              }}
              onMouseEnter={(e) => {
                if (!loading) {
                  e.currentTarget.style.backgroundColor = '#0e6e0e';
                  e.currentTarget.style.transform = 'translateY(-1px)';
                  e.currentTarget.style.boxShadow = '0 2px 4px rgba(16, 124, 16, 0.3)';
                }
              }}
              onMouseLeave={(e) => {
                if (!loading) {
                  e.currentTarget.style.backgroundColor = '#107c10';
                  e.currentTarget.style.transform = 'none';
                  e.currentTarget.style.boxShadow = 'none';
                }
              }}
            >
              🔄 Conversion
            </button>
          </div>
        </div>

        {error && (
          <div className={styles.error}>
            <span className={styles.errorIcon}>⚠️</span>
            <span className={styles.errorMessage}>Error: {error}</span>
            <button 
              className={styles.clearErrorButton}
              onClick={this.clearError}
              title="Clear error"
            >
              ✕
            </button>
          </div>
        )}

        {loading ? (
          <div className={styles.loading}>
            Loading properties...
          </div>
        ) : sortedItems.length === 0 ? (
          <div className={styles.empty}>
            <div className={styles.emptyMessage}>No properties found for this convert file.</div>
            <div className={styles.emptyActions}>
              <ExcelImportButton
                onImport={this.handleExcelImport}
                disabled={loading}
                existingItemsCount={0}
              />
              <button 
                className={styles.addButton}
                onClick={this.handleAdd}
              >
                Add First Property
              </button>
            </div>
          </div>
        ) : (
          <table className={styles.table}>
            <thead className={styles.tableHead}>
              <tr>
                <th className={styles.headerCell}>Priority</th>
                <th className={styles.headerCell}>Title</th>
                <th className={styles.headerCell}>Prop</th>
                <th className={styles.headerCell}>Prop2</th>
                <th className={styles.headerCell}>Convert Type</th>
                <th className={styles.headerCell}>Convert Type2</th>
                <th className={styles.headerCell}>Status</th>
                <th className={styles.headerCell}>Created</th>
                <th className={styles.headerCell}>Actions</th>
              </tr>
            </thead>
            <tbody className={styles.tableBody}>
              {sortedItems.map((item: IConvertFileProps) => (
                <tr 
                  key={item.Id}
                  className={`${styles.tableRow} ${item.IsDeleted ? styles.deleted : ''}`}
                >
                  <td className={`${styles.tableCell} ${styles.priorityCell}`}>
                    {item.Priority}
                  </td>
                  <td className={`${styles.tableCell} ${styles.titleCell}`}>
                    {item.Title}
                  </td>
                  <td 
                    className={`${styles.tableCell} ${styles.propCell}`}
                    style={{
                      ...this.getColumnValidationStyle(item.Id, 'prop'),
                      color: '#605e5c'
                    }}
                  >
                    {item.Prop}
                  </td>
                  <td 
                    className={`${styles.tableCell} ${styles.propCell}`}
                    style={{
                      ...this.getColumnValidationStyle(item.Id, 'prop2'),
                      color: '#605e5c'
                    }}
                  >
                    {item.Prop2}
                  </td>
                  <td className={`${styles.tableCell} ${styles.convertTypeCell}`}>
                    <span 
                      className={styles.convertTypeBadge}
                      title={`Convert Type ID: ${item.ConvertType}`}
                    >
                      {this.getConvertTypeName(item.ConvertType, convertTypes)}
                    </span>
                  </td>
                  <td className={`${styles.tableCell} ${styles.convertTypeCell}`}>
                    <span 
                      className={styles.convertTypeBadge}
                      title={`Convert Type2 ID: ${item.ConvertType2}`}
                    >
                      {this.getConvertTypeName(item.ConvertType2, convertTypes)}
                    </span>
                  </td>
                  <td className={`${styles.tableCell} ${styles.statusCell}`}>
                    <span className={`${styles.statusBadge} ${item.IsDeleted ? styles.deleted : styles.active}`}>
                      {item.IsDeleted ? 'Deleted' : 'Active'}
                    </span>
                  </td>
                  <td className={styles.tableCell}>
                    {item.Created ? new Date(item.Created).toLocaleDateString() : '-'}
                  </td>
                  <td className={`${styles.tableCell} ${styles.actionsCell}`}>
                    <button 
                      className={`${styles.actionButton} ${styles.moveButton}`}
                      onClick={() => this.handleMoveUp(item.Id)}
                      disabled={loading || !this.canMoveUp(item)}
                      title={item.IsDeleted ? "Move Deleted Item Up" : "Move Up"}
                    >
                      ↑
                    </button>
                    <button 
                      className={`${styles.actionButton} ${styles.moveButton}`}
                      onClick={() => this.handleMoveDown(item.Id)}
                      disabled={loading || !this.canMoveDown(item)}
                      title={item.IsDeleted ? "Move Deleted Item Down" : "Move Down"}
                    >
                      ↓
                    </button>
                    
                    <button 
                      className={`${styles.actionButton} ${styles.editButton}`}
                      onClick={() => this.handleEdit(item)}
                      disabled={loading || item.IsDeleted}
                      title={item.IsDeleted ? "Cannot edit deleted item" : "Edit"}
                    >
                      Edit
                    </button>
                    
                    {item.IsDeleted ? (
                      <button 
                        className={`${styles.actionButton} ${styles.restoreButton}`}
                        onClick={() => this.handleRestore(item.Id)}
                        disabled={loading}
                        title="Restore"
                      >
                        Restore
                      </button>
                    ) : (
                      <button 
                        className={`${styles.actionButton} ${styles.deleteButton}`}
                        onClick={() => this.handleDelete(item.Id)}
                        disabled={loading}
                        title="Mark as Deleted"
                      >
                        Delete
                      </button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}

        <ConfirmationDialog
          isOpen={showConversionDialog}
          title={conversionDialogConfig.title}
          message={conversionDialogConfig.message}
          confirmText={conversionDialogConfig.confirmText}
          cancelText={conversionDialogConfig.cancelText || 'Cancel'}
          type={conversionDialogConfig.type}
          showIcon={conversionDialogConfig.showIcon}
          loading={conversionDialogLoading}
          onConfirm={() => { this.handleConversionConfirm().catch(console.error); }}
          onCancel={this.handleConversionCancel}
        />
      </div>
    );
  }
}