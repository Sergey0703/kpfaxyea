// src/webparts/xyea/components/ExcelImportButton/ExcelImportButton.tsx

import * as React from 'react';
import * as XLSX from 'xlsx';
import styles from './ExcelImportButton.module.scss';
import { ConfirmationDialog, IConfirmationDialogConfig } from '../ConfirmationDialog';

export interface IExcelImportData {
  prop: string;
  prop2: string;
}

export interface IExcelImportButtonProps {
  onImport: (data: IExcelImportData[]) => Promise<void>;
  disabled?: boolean;
  existingItemsCount: number;
}

export interface IExcelImportButtonState {
  loading: boolean;
  error: string | undefined;
  showConfirmDialog: boolean;
  confirmDialogConfig: IConfirmationDialogConfig;
  confirmDialogLoading: boolean;
  pendingData: IExcelImportData[] | undefined;
  dialogType: 'initial' | 'final'; // Track which type of dialog we're showing
}

export default class ExcelImportButton extends React.Component<IExcelImportButtonProps, IExcelImportButtonState> {
  private fileInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: IExcelImportButtonProps) {
    super(props);
    
    this.state = {
      loading: false,
      error: undefined,
      showConfirmDialog: false,
      confirmDialogConfig: {
        title: '',
        message: '',
        confirmText: 'Continue',
        type: 'warning'
      },
      confirmDialogLoading: false,
      pendingData: undefined,
      dialogType: 'initial'
    };

    this.fileInputRef = React.createRef<HTMLInputElement>();
  }

  private handleButtonClick = (): void => {
    if (this.props.disabled) {
      return;
    }

    // Если есть существующие элементы, показываем предупреждение
    if (this.props.existingItemsCount > 0) {
      this.setState({
        showConfirmDialog: true,
        dialogType: 'initial',
        confirmDialogConfig: {
          title: 'Replace Existing Data',
          message: `You currently have ${this.props.existingItemsCount} item${this.props.existingItemsCount > 1 ? 's' : ''} in this table.\n\nImporting from Excel will delete all existing items and replace them with data from the Excel file.\n\nAre you sure you want to continue?`,
          confirmText: 'Yes, Continue',
          cancelText: 'Cancel',
          type: 'danger',
          showIcon: true
        }
      });
    } else {
      // Если нет существующих элементов, сразу открываем диалог выбора файла
      this.openFileDialog();
    }
  }

  private openFileDialog = (): void => {
    console.log('[ExcelImportButton] Opening file dialog');
    if (this.fileInputRef.current) {
      this.fileInputRef.current.click();
    } else {
      console.error('[ExcelImportButton] File input ref is null');
    }
  }

  private handleFileInputChange = async (event: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    console.log('[ExcelImportButton] File input changed');
    const files = event.target.files;
    if (!files || files.length === 0) {
      console.log('[ExcelImportButton] No files selected');
      return;
    }

    const file = files[0];
    console.log('[ExcelImportButton] Selected file:', file.name);
    await this.processExcelFile(file);

    // Reset file input
    if (this.fileInputRef.current) {
      this.fileInputRef.current.value = '';
    }
  }

  private processExcelFile = async (file: File): Promise<void> => {
    this.setState({ loading: true, error: undefined });

    try {
      // Validate file format
      const fileName = file.name.toLowerCase();
      if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
        throw new Error('Please select a valid Excel file (.xlsx or .xls)');
      }

      // Validate file size (max 5MB)
      if (file.size > 5 * 1024 * 1024) {
        throw new Error('File size is too large. Please select a file smaller than 5MB.');
      }

      console.log('[ExcelImportButton] Processing file:', file.name);

      // Read file
      const arrayBuffer = await this.readFileAsArrayBuffer(file);
      
      // Parse Excel
      const workbook = XLSX.read(arrayBuffer, { 
        type: 'array',
        cellDates: true,
        dateNF: 'yyyy-mm-dd'
      });

      // Get first sheet
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) {
        throw new Error('No sheets found in the Excel file');
      }

      const worksheet = workbook.Sheets[firstSheetName];
      
      // Convert to array format
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        defval: '',
        raw: false
      }) as (string | number | boolean | undefined)[][];

      if (jsonData.length <= 1) {
        throw new Error('Excel file must contain at least one data row (excluding header)');
      }

      // Process data starting from row 2 (index 1)
      const importData: IExcelImportData[] = [];
      for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        
        // Get values from columns A (index 0) and B (index 1)
        const prop = this.sanitizeValue(row[0]);
        const prop2 = this.sanitizeValue(row[1]);

        // Skip completely empty rows
        if (!prop && !prop2) {
          continue;
        }

        importData.push({
          prop: prop || '', // Allow empty strings
          prop2: prop2 || '' // Allow empty strings
        });
      }

      if (importData.length === 0) {
        throw new Error('No valid data found in the Excel file. Please check that columns A and B contain data starting from row 2.');
      }

      console.log('[ExcelImportButton] Parsed data:', {
        totalRows: importData.length,
        sample: importData.slice(0, 3)
      });

      // Set pending data and show final confirmation
      this.setState({
        pendingData: importData,
        showConfirmDialog: true,
        dialogType: 'final',
        confirmDialogConfig: {
          title: 'Confirm Data Import',
          message: `Excel file processed successfully!\n\nFound ${importData.length} row${importData.length > 1 ? 's' : ''} of data.\n\nThis will replace all ${this.props.existingItemsCount} existing item${this.props.existingItemsCount > 1 ? 's' : ''} in the table.\n\nProceed with import?`,
          confirmText: 'Yes, Import Data',
          cancelText: 'Cancel',
          type: 'warning',
          showIcon: true
        }
      });

    } catch (error) {
      console.error('[ExcelImportButton] Error processing file:', error);
      this.setState({ 
        error: error instanceof Error ? error.message : 'Failed to process Excel file',
        pendingData: undefined
      });
    } finally {
      this.setState({ loading: false });
    }
  }

  private sanitizeValue = (value: string | number | boolean | undefined): string => {
    if (value === undefined || value === null) {
      return '';
    }
    
    if (typeof value === 'string') {
      return value.trim();
    }
    
    return String(value).trim();
  }

  private readFileAsArrayBuffer = (file: File): Promise<ArrayBuffer> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (event) => {
        if (event.target?.result instanceof ArrayBuffer) {
          resolve(event.target.result);
        } else {
          reject(new Error('Failed to read file as ArrayBuffer'));
        }
      };
      
      reader.onerror = () => reject(new Error('FileReader error'));
      reader.readAsArrayBuffer(file);
    });
  }

  private handleConfirmAction = async (): Promise<void> => {
    const { pendingData, dialogType } = this.state;
    
    if (dialogType === 'initial') {
      // This is the initial "replace data" confirmation - open file dialog
      this.setState({ 
        showConfirmDialog: false,
        confirmDialogLoading: false 
      });
      
      // Small delay to ensure dialog is closed before opening file dialog
      setTimeout(() => {
        this.openFileDialog();
      }, 100);
      
    } else if (dialogType === 'final' && pendingData) {
      // This is the final "import data" confirmation - proceed with import
      this.setState({ confirmDialogLoading: true });

      try {
        await this.performImport(pendingData);
        this.setState({ 
          showConfirmDialog: false,
          pendingData: undefined,
          confirmDialogLoading: false,
          dialogType: 'initial'
        });
      } catch (error) {
        console.error('[ExcelImportButton] Import failed:', error);
        this.setState({ 
          error: error instanceof Error ? error.message : 'Import failed',
          confirmDialogLoading: false
        });
      }
    }
  }

  private performImport = async (data: IExcelImportData[]): Promise<void> => {
    try {
      await this.props.onImport(data);
      console.log('[ExcelImportButton] Import completed successfully');
    } catch (error) {
      console.error('[ExcelImportButton] Import failed:', error);
      throw error;
    }
  }

  private handleCancelConfirmation = (): void => {
    this.setState({ 
      showConfirmDialog: false,
      pendingData: undefined,
      confirmDialogLoading: false,
      dialogType: 'initial'
    });
  }

  private clearError = (): void => {
    this.setState({ error: undefined });
  }

  public render(): React.ReactElement<IExcelImportButtonProps> {
    const { disabled } = this.props;
    const { loading, error, showConfirmDialog, confirmDialogConfig, confirmDialogLoading } = this.state;

    return (
      <div className={styles.excelImportButton}>
        <input
          ref={this.fileInputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => { this.handleFileInputChange(e).catch(console.error); }}
          style={{ display: 'none' }}
          disabled={disabled || loading}
        />

        <button
          className={styles.importButton}
          onClick={this.handleButtonClick}
          disabled={disabled || loading}
          title="Import properties from Excel file"
        >
          {loading ? (
            <>
              <span className={styles.spinner} />
              Processing...
            </>
          ) : (
            <>
              📥 Import from Excel
            </>
          )}
        </button>

        {error && (
          <div className={styles.error}>
            <span className={styles.errorIcon}>⚠️</span>
            <span className={styles.errorMessage}>{error}</span>
            <button 
              className={styles.clearErrorButton}
              onClick={this.clearError}
              title="Clear error"
            >
              ✕
            </button>
          </div>
        )}

        <ConfirmationDialog
          isOpen={showConfirmDialog}
          title={confirmDialogConfig.title}
          message={confirmDialogConfig.message}
          confirmText={confirmDialogConfig.confirmText}
          cancelText={confirmDialogConfig.cancelText || 'Cancel'}
          type={confirmDialogConfig.type}
          showIcon={confirmDialogConfig.showIcon}
          loading={confirmDialogLoading}
          onConfirm={() => { this.handleConfirmAction().catch(console.error); }}
          onCancel={this.handleCancelConfirmation}
        />
      </div>
    );
  }
}