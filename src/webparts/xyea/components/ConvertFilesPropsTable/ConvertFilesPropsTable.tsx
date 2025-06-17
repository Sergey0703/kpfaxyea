// src/webparts/xyea/components/ConvertFilesPropsTable/ConvertFilesPropsTable.tsx

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
    exportFile?: { headers: string[]; data: any[][] };
    importFile?: { headers: string[]; data: any[][] };
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
      this.validateAllColumns();
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

  private handleConversion = (): void => {
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
    this.setState({ conversionDialogLoading: true });

    try {
      // Check if both files are selected
      const selectedFiles = this.props.selectedFiles?.[this.props.convertFileId];
      if (!selectedFiles?.export || !selectedFiles?.import) {
        throw new Error('Both Export file and Import file must be selected before starting conversion.');
      }

      // Start conversion process
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

  private performConversion = async (exportFile: File, importFile: File): Promise<void> => {
    console.log('[ConvertFilesPropsTable] Starting conversion process');

    // Read both files
    const exportData = await this.readExcelFile(exportFile);
    const importData = await this.readExcelFile(importFile);

    const processedRows: number[] = [];
    const skippedRows: Array<{ row: number; reason: string; column?: string; file?: string }> = [];

    // Process each row in the properties table
    const sortedItems = this.getSortedItems().filter(item => !item.IsDeleted);
    
    for (let i = 0; i < sortedItems.length; i++) {
      const item = sortedItems[i];
      const rowNumber = i + 1;

      try {
        // Find source column in export file
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

        // Find target column in import file
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

        // Copy data from source to target column
        for (let dataRowIndex = 0; dataRowIndex < exportData.data.length; dataRowIndex++) {
          const sourceValue = exportData.data[dataRowIndex][sourceColumnIndex];
          
          // Ensure import data has enough rows
          if (dataRowIndex >= importData.data.length) {
            // Add new rows as needed
            const newRow = new Array(importData.headers.length).fill('');
            importData.data.push(newRow);
          }

          // Copy the value
          importData.data[dataRowIndex][targetColumnIndex] = sourceValue || '';
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

    // Generate and download the updated import file
    await this.downloadUpdatedFile(importData, importFile.name);

    // Show completion message
    this.showConversionResults(processedRows, skippedRows);
  }

  private readExcelFile = async (file: File): Promise<{ headers: string[]; data: any[][] }> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (event) => {
        try {
          const arrayBuffer = event.target?.result as ArrayBuffer;
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          
          // Convert to array format
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
            header: 1,
            defval: '',
            raw: false
          }) as any[][];

          if (jsonData.length === 0) {
            reject(new Error(`File ${file.name} is empty`));
            return;
          }

          const headers = jsonData[0]?.map((header: any) => String(header || '').trim()) || [];
          const data = jsonData.slice(1);

          resolve({ headers, data });
        } catch (error) {
          reject(new Error(`Failed to read ${file.name}: ${error instanceof Error ? error.message : 'Unknown error'}`));
        }
      };

      reader.onerror = () => reject(new Error(`Failed to read file ${file.name}`));
      reader.readAsArrayBuffer(file);
    });
  }

  private downloadUpdatedFile = async (data: { headers: string[]; data: any[][] }, originalFileName: string): Promise<void> => {
    // Create new workbook
    const workbook = XLSX.utils.book_new();
    
    // Combine headers and data
    const worksheetData = [data.headers, ...data.data];
    
    // Create worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    
    // Generate filename
    const fileName = originalFileName.replace(/\.[^/.]+$/, '_converted.xlsx');
    
    // Download file
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
    if (!selectedFiles?.export || !selectedFiles?.import) {
      // Clear validation if files not selected
      this.setState({ columnValidation: {}, filesCache: {} });
      return;
    }

    try {
      // Read files if not cached
      let exportData = this.state.filesCache.exportFile;
      let importData = this.state.filesCache.importFile;

      if (!exportData) {
        exportData = await this.readExcelFile(selectedFiles.export);
      }
      if (!importData) {
        importData = await this.readExcelFile(selectedFiles.import);
      }

      // Cache the data
      this.setState({
        filesCache: {
          exportFile: exportData,
          importFile: importData
        }
      });

      // Validate all items
      const validation: { [itemId: number]: any } = {};
      
      this.props.items.forEach(item => {
        const propExists = exportData!.headers.some(header => 
          header.toLowerCase().trim() === item.Prop.toLowerCase().trim()
        );
        const prop2Exists = importData!.headers.some(header => 
          header.toLowerCase().trim() === item.Prop2.toLowerCase().trim()
        );

        validation[item.Id] = {
          propValid: propExists,
          prop2Valid: prop2Exists,
          propExists,
          prop2Exists
        };
      });

      this.setState({ columnValidation: validation });

    } catch (error) {
      console.error('[ConvertFilesPropsTable] Error validating columns:', error);
      this.setState({ columnValidation: {}, filesCache: {} });
    }
  }

  private getColumnValidationClass = (itemId: number, column: 'prop' | 'prop2'): string => {
    const validation = this.state.columnValidation[itemId];
    if (!validation) return styles.columnUnknown;

    const isValid = column === 'prop' ? validation.propValid : validation.prop2Valid;
    const exists = column === 'prop' ? validation.propExists : validation.prop2Exists;

    if (exists === undefined) return styles.columnUnknown;
    return isValid ? styles.columnValid : styles.columnInvalid;
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

  public render(): React.ReactElement<IConvertFilesPropsTableProps> {
    const { convertFileTitle, loading } = this.props;
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
              className={styles.conversionButton}
              onClick={this.handleConversion}
              disabled={loading}
              title="Start conversion process"
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
                  <td className={`${styles.tableCell} ${styles.propCell} ${this.getColumnValidationClass(item.Id, 'prop')}`}>
                    {item.Prop}
                  </td>
                  <td className={`${styles.tableCell} ${styles.propCell} ${this.getColumnValidationClass(item.Id, 'prop2')}`}>
                    {item.Prop2}
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
}<button 
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
      </div>
    );
  }
}