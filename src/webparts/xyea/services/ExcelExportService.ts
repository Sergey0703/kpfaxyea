// src/webparts/xyea/services/ExcelExportService.ts

import * as XLSX from 'xlsx';
import { 
  IExcelSheet, 
  IFilterState, 
  IExportSettings,
  ExportFormat,
  IExcelRow
} from '../interfaces/ExcelInterfaces';
import { 
  IRenameFilesData,
  IRenameExportSettings,
  IRenameExportStatistics
} from '../components/RenameFilesManagement/types/RenameFilesTypes';

// Type definitions for better type safety
type CellValue = string | number | boolean | Date | undefined; // Changed from null to undefined
type ExcelRowData = CellValue[];
type ExcelData = ExcelRowData[];

interface IExportStatistics {
  totalRows: number;
  visibleRows: number;
  hiddenRows: number;
  activeFilters: number;
  estimatedFileSize: string;
  canExport: boolean;
}

interface IExportPreview {
  headers: string[];
  sampleRows: ExcelRowData[];
  totalRows: number;
  hasMoreData: boolean;
}

interface IExportValidation {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

interface IColumnWidth {
  wch: number;
}

export class ExcelExportService {

  /**
   * Экспорт отфильтрованных данных для Separate Files Management
   */
  public static async exportFilteredData(
    originalFileName: string,
    sheet: IExcelSheet,
    filterState: IFilterState,
    settings: IExportSettings
  ): Promise<{ success: boolean; fileName?: string; error?: string }> {
    try {
      console.log('[ExcelExportService] Starting export:', {
        originalFileName,
        sheetName: sheet.name,
        totalRows: sheet.totalRows,
        settings
      });

      // Получаем видимые строки
      const visibleRows = sheet.data.filter(row => row.isVisible);
      
      if (visibleRows.length === 0) {
        return {
          success: false,
          error: 'No data to export. Please adjust your filters.'
        };
      }

      // Применяем ограничение по количеству строк
      const exportRows = settings.maxRows && settings.maxRows > 0 
        ? visibleRows.slice(0, settings.maxRows)
        : visibleRows;

      // Генерируем имя файла
      const fileName = this.generateExportFileName(originalFileName, filterState, settings.fileFormat);

      // Подготавливаем данные для экспорта
      const exportData = this.prepareExportData(sheet.headers, exportRows, settings);

      // Экспортируем в зависимости от формата
      if (settings.fileFormat === ExportFormat.CSV) {
        await this.exportAsCSV(exportData, fileName);
      } else {
        await this.exportAsExcel(exportData, fileName, sheet.name);
      }

      console.log('[ExcelExportService] Export completed:', {
        fileName,
        rowsExported: exportRows.length,
        format: settings.fileFormat
      });

      return {
        success: true,
        fileName
      };

    } catch (error) {
      console.error('[ExcelExportService] Export failed:', error);
      return {
        success: false,
        error: `Export failed: ${error instanceof Error ? error.message : 'Unknown error'}`
      };
    }
  }

  /**
   * NEW: Export rename files data with status information
   */
  public static async exportRenameFilesData(
    data: IRenameFilesData,
    fileSearchResults: { [rowIndex: number]: 'found' | 'not-found' | 'searching' | 'skipped' },
    renameProgress?: {
      current: number;
      total: number;
      fileName: string;
      success: number;
      errors: number;
      skipped: number;
    },
    exportSettings?: IRenameExportSettings
  ): Promise<{ success: boolean; fileName?: string; error?: string }> {
    try {
      console.log('[ExcelExportService] Starting rename files export:', {
        totalRows: data.totalRows,
        columns: data.columns.length,
        exportSettings
      });

      // Use default settings if not provided
      const settings: IRenameExportSettings = exportSettings || {
        fileName: 'renamed_files_export',
        includeHeaders: true,
        includeStatusColumn: true,
        includeTimestamps: true,
        onlyCompletedRows: false,
        fileFormat: 'xlsx'
      };

      // Prepare export data with status information
      const exportData = this.prepareRenameFilesExportData(
        data,
        fileSearchResults,
        renameProgress,
        settings
      );
      
      if (exportData.length === 0) {
        return {
          success: false,
          error: 'No data to export. Please check your export settings.'
        };
      }

      // Generate filename
      const fileName = this.generateRenameExportFileName(
        settings.fileName,
        settings.fileFormat
      );

      // Create and download file
      const blob = await this.createRenameFilesExportFile(exportData, settings);
      this.downloadFile(blob, fileName);

      console.log('[ExcelExportService] Rename files export completed:', {
        fileName,
        rowsExported: exportData.length - (settings.includeHeaders ? 1 : 0)
      });

      return {
        success: true,
        fileName
      };

    } catch (error) {
      console.error('[ExcelExportService] Rename files export failed:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Export failed'
      };
    }
  }

  /**
   * NEW: Get export statistics for rename files
   */
  public static getRenameFilesExportStatistics(
    data: IRenameFilesData,
    fileSearchResults: { [rowIndex: number]: 'found' | 'not-found' | 'searching' | 'skipped' },
    renameProgress?: {
      current: number;
      total: number;
      fileName: string;
      success: number;
      errors: number;
      skipped: number;
    },
    exportSettings?: IRenameExportSettings
  ): IRenameExportStatistics {
    
    const settings = exportSettings || {
      fileName: 'renamed_files_export',
      includeHeaders: true,
      includeStatusColumn: true,
      includeTimestamps: true,
      onlyCompletedRows: false,
      fileFormat: 'xlsx'
    };

    // Count different statuses
    let foundFiles = 0;
    let notFoundFiles = 0;
    let searchingFiles = 0;
    let skippedFiles = 0;
    let renamedFiles = 0;
    let errorFiles = 0;

    data.rows.forEach(row => {
      const searchStatus = fileSearchResults[row.rowIndex];
      
      switch (searchStatus) {
        case 'found':
          foundFiles++;
          break;
        case 'not-found':
          notFoundFiles++;
          break;
        case 'searching':
          searchingFiles++;
          break;
        case 'skipped':
          skippedFiles++;
          break;
      }
    });

    // Add rename statistics if available
    if (renameProgress) {
      renamedFiles = renameProgress.success;
      errorFiles = renameProgress.errors;
      skippedFiles += renameProgress.skipped;
    }

    // Calculate exportable rows
    let exportableRows = data.totalRows;
    if (settings.onlyCompletedRows) {
      exportableRows = foundFiles + notFoundFiles + renamedFiles + errorFiles + skippedFiles;
    }

    // Estimate file size
    const estimatedFileSize = this.formatFileSize(exportableRows * (data.columns.length + 1) * 20);

    return {
      totalRows: data.totalRows,
      exportableRows,
      foundFiles,
      notFoundFiles,
      renamedFiles,
      errorFiles,
      skippedFiles,
      searchingFiles,
      estimatedFileSize,
      canExport: exportableRows > 0
    };
  }

  /**
   * Генерация имени файла на основе фильтров для Separate Files
   */
  private static generateExportFileName(
    originalFileName: string,
    filterState: IFilterState,
    format: ExportFormat
  ): string {
    // Удаляем расширение из оригинального имени
    const baseName = originalFileName.replace(/\.[^/.]+$/, '');
    
    // Собираем активные фильтры
    const activeFilters = Object.values(filterState.filters)
      .filter(filter => filter.isActive)
      .map(filter => {
        const selectedCount = filter.selectedValues.length;
        const totalCount = filter.totalUniqueValues;
        
        if (selectedCount === 1) {
          // Если выбрано одно значение, добавляем его в имя
          const value = String(filter.selectedValues[0]).replace(/[^a-zA-Z0-9]/g, '');
          return `${filter.columnName}-${value}`;
        } else {
          // Если выбрано несколько, показываем количество
          return `${filter.columnName}-${selectedCount}of${totalCount}`;
        }
      });

    // Формируем финальное имя
    let fileName = baseName;
    
    if (activeFilters.length > 0) {
      const filterSuffix = activeFilters.slice(0, 3).join('_'); // Максимум 3 фильтра в имени
      fileName += `_${filterSuffix}`;
      
      if (activeFilters.length > 3) {
        fileName += '_etc';
      }
    }

    // Добавляем временную метку для уникальности
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    fileName += `_${timestamp}`;

    // Добавляем расширение
    const extension = format === ExportFormat.CSV ? 'csv' : 'xlsx';
    return `${fileName}.${extension}`;
  }

  /**
   * NEW: Generate export filename for rename files
   */
  private static generateRenameExportFileName(baseName: string, format: 'xlsx' | 'csv'): string {
    // Remove existing extension
    const nameWithoutExt = baseName.replace(/\.(xlsx|xls|csv)$/i, '');
    
    // Add timestamp
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, '_');
    const cleanName = nameWithoutExt.replace(/[^a-zA-Z0-9_-]/g, '_');
    
    // Determine extension
    const extension = format === 'csv' ? 'csv' : 'xlsx';
    
    return `${cleanName}_${timestamp}.${extension}`;
  }

  /**
   * Подготовка данных для экспорта (Separate Files)
   */
  private static prepareExportData(
    headers: string[],
    rows: IExcelRow[],
    settings: IExportSettings
  ): ExcelData {
    const exportData: ExcelData = [];

    // Добавляем заголовки если нужно
    if (settings.includeHeaders) {
      exportData.push([...headers]);
    }

    // Добавляем данные
    rows.forEach(row => {
      const rowData: ExcelRowData = [];
      headers.forEach(header => {
        const cellValue = row.data[header];
        // Преобразуем значения для экспорта
        rowData.push(this.formatCellForExport(cellValue));
      });
      exportData.push(rowData);
    });

    return exportData;
  }

  /**
   * NEW: Prepare rename files data for export
   */
  private static prepareRenameFilesExportData(
    data: IRenameFilesData,
    fileSearchResults: { [rowIndex: number]: 'found' | 'not-found' | 'searching' | 'skipped' },
    renameProgress?: {
      current: number;
      total: number;
      fileName: string;
      success: number;
      errors: number;
      skipped: number;
    },
    settings?: IRenameExportSettings
  ): any[][] {
    
    const exportData: any[][] = [];
    
    // Prepare headers
    if (settings?.includeHeaders) {
      const headers: string[] = [];
      
      // Add data columns in order
      data.columns
        .sort((a, b) => a.currentIndex - b.currentIndex)
        .filter(col => col.isVisible)
        .forEach(column => {
          headers.push(column.name);
        });
      
      // Add status column
      if (settings.includeStatusColumn) {
        headers.push('Status');
      }
      
      // Add timestamp column
      if (settings.includeTimestamps) {
        headers.push('Export Timestamp');
      }
      
      exportData.push(headers);
    }
    
    // Process each row
    data.rows.forEach(row => {
      const searchStatus = fileSearchResults[row.rowIndex];
      
      // Filter rows based on settings
      if (settings.onlyCompletedRows) {
        if (searchStatus === 'searching') {
          return; // Skip rows that are still searching
        }
      }
      
      const rowData: any[] = [];
      
      // Add cell values in column order
      data.columns
        .sort((a, b) => a.currentIndex - b.currentIndex)
        .filter(col => col.isVisible)
        .forEach(column => {
          const cell = row.cells[column.id];
          const value = cell ? cell.value : '';
          
          // Format the value appropriately
          if (value instanceof Date) {
            rowData.push(value.toLocaleDateString());
          } else if (typeof value === 'number') {
            rowData.push(value);
          } else {
            rowData.push(String(value || ''));
          }
        });
      
      // Add status information
      if (settings.includeStatusColumn) {
        const statusText = this.getRenameStatusText(searchStatus, renameProgress, row.rowIndex);
        rowData.push(statusText);
      }
      
      // Add timestamp
      if (settings.includeTimestamps) {
        rowData.push(new Date().toLocaleString());
      }
      
      exportData.push(rowData);
    });
    
    return exportData;
  }

  /**
   * NEW: Get human-readable status text
   */
  private static getRenameStatusText(
    searchStatus: 'found' | 'not-found' | 'searching' | 'skipped',
    renameProgress?: {
      current: number;
      total: number;
      fileName: string;
      success: number;
      errors: number;
      skipped: number;
    },
    rowIndex?: number
  ): string {
    
    // If rename is in progress or completed, show rename status
    if (renameProgress && renameProgress.current > 0) {
      if (renameProgress.success > 0) {
        return 'Renamed Successfully';
      } else if (renameProgress.errors > 0) {
        return 'Rename Failed';
      } else if (renameProgress.skipped > 0) {
        return 'Skipped (Target Exists)';
      }
    }
    
    // Otherwise show search status
    switch (searchStatus) {
      case 'found':
        return 'Found in SharePoint';
      case 'not-found':
        return 'Not Found';
      case 'searching':
        return 'Searching...';
      case 'skipped':
        return 'Skipped';
      default:
        return 'Unknown';
    }
  }

  /**
   * Форматирование значения ячейки для экспорта
   */
  private static formatCellForExport(value: CellValue): CellValue {
    if (value === undefined) {
      return '';
    }

    // Преобразуем даты в читаемый формат
    if (value instanceof Date) {
      return value.toISOString().split('T')[0]; // YYYY-MM-DD
    }

    // Преобразуем boolean в читаемый вид
    if (typeof value === 'boolean') {
      return value ? 'Yes' : 'No';
    }

    return value;
  }

  /**
   * Экспорт в Excel формат (Separate Files)
   */
  private static async exportAsExcel(
    data: ExcelData,
    fileName: string,
    sheetName: string
  ): Promise<void> {
    // Создаем новую книгу
    const workbook = XLSX.utils.book_new();
    
    // Создаем лист из данных
    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // Устанавливаем ширину колонок
    const colWidths = this.calculateColumnWidths(data);
    worksheet['!cols'] = colWidths;

    // Добавляем лист в книгу
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName || 'Filtered Data');

    // Записываем файл
    XLSX.writeFile(workbook, fileName);
  }

  /**
   * Экспорт в CSV формат (Separate Files)
   */
  private static async exportAsCSV(
    data: ExcelData,
    fileName: string
  ): Promise<void> {
    // Преобразуем данные в CSV строку
    const csvContent = data
      .map(row => 
        row.map(cell => {
          // Экранируем кавычки и добавляем кавычки если нужно
          const cellStr = String(cell || '');
          if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
            return `"${cellStr.replace(/"/g, '""')}"`;
          }
          return cellStr;
        }).join(',')
      )
      .join('\n');

    // Создаем Blob и скачиваем
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', fileName);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    URL.revokeObjectURL(url);
  }

  /**
   * NEW: Create export file for rename files data
   */
  private static async createRenameFilesExportFile(
    data: any[][],
    settings: IRenameExportSettings
  ): Promise<Blob> {
    
    if (settings.fileFormat === 'csv') {
      return this.createCSVBlobFromData(data);
    } else {
      return this.createExcelBlobFromData(data, 'Rename Files Export');
    }
  }

  /**
   * NEW: Create CSV blob from data array
   */
  private static createCSVBlobFromData(data: any[][]): Blob {
    const csvContent = data
      .map(row => 
        row.map(cell => {
          const cellValue = String(cell || '');
          // Escape quotes and wrap in quotes if contains comma, quote, or newline
          if (cellValue.includes(',') || cellValue.includes('"') || cellValue.includes('\n')) {
            return `"${cellValue.replace(/"/g, '""')}"`;
          }
          return cellValue;
        }).join(',')
      )
      .join('\n');
    
    return new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  }

  /**
   * NEW: Create Excel blob from data array
   */
  private static createExcelBlobFromData(data: any[][], sheetName: string = 'Sheet1'): Blob {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    
    // Auto-adjust column widths
    const columnWidths = this.calculateRenameColumnWidths(data);
    worksheet['!cols'] = columnWidths;
    
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    
    const excelBuffer = XLSX.write(workbook, { 
      bookType: 'xlsx', 
      type: 'array',
      compression: true
    });
    
    return new Blob([excelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  /**
   * Вычисление оптимальной ширины колонок (Separate Files)
   */
  private static calculateColumnWidths(data: ExcelData): IColumnWidth[] {
    if (data.length === 0) return [];

    const columnCount = data[0].length;
    const widths: number[] = new Array(columnCount).fill(10);

    // Анализируем каждую колонку
    for (let colIndex = 0; colIndex < columnCount; colIndex++) {
      let maxWidth = 10;

      for (let rowIndex = 0; rowIndex < Math.min(data.length, 100); rowIndex++) { // Анализируем первые 100 строк
        const cellValue = data[rowIndex][colIndex];
        const cellLength = String(cellValue || '').length;
        maxWidth = Math.max(maxWidth, cellLength);
      }

      // Ограничиваем максимальную ширину
      widths[colIndex] = Math.min(maxWidth + 2, 50);
    }

    return widths.map(width => ({ wch: width }));
  }

  /**
   * NEW: Calculate column widths for rename files export
   */
  private static calculateRenameColumnWidths(data: any[][]): Array<{ wch: number }> {
    if (data.length === 0) return [];
    
    const columnCount = data[0].length;
    const widths: Array<{ wch: number }> = [];
    
    for (let col = 0; col < columnCount; col++) {
      let maxWidth = 10; // Minimum width
      
      data.forEach(row => {
        if (row[col] !== undefined && row[col] !== null) {
          const cellLength = String(row[col]).length;
          maxWidth = Math.max(maxWidth, Math.min(cellLength, 50)); // Max width of 50
        }
      });
      
      widths.push({ wch: maxWidth });
    }
    
    return widths;
  }

  /**
   * NEW: Download file helper method
   */
  private static downloadFile(blob: Blob, fileName: string): void {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.style.display = 'none';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // Clean up the URL object
    setTimeout(() => {
      window.URL.revokeObjectURL(url);
    }, 100);
  }

  /**
   * Получение статистики экспорта (Separate Files)
   */
  public static getExportStatistics(
    sheet: IExcelSheet,
    filterState: IFilterState
  ): IExportStatistics {
    const visibleRows = sheet.data.filter(row => row.isVisible).length;
    const hiddenRows = sheet.totalRows - visibleRows;
    const activeFilters = Object.values(filterState.filters).filter(f => f.isActive).length;

    // Приблизительная оценка размера файла
    const avgCellSize = 20; // байт на ячейку
    const estimatedBytes = visibleRows * sheet.headers.length * avgCellSize;
    const estimatedFileSize = this.formatFileSize(estimatedBytes);

    return {
      totalRows: sheet.totalRows,
      visibleRows,
      hiddenRows,
      activeFilters,
      estimatedFileSize,
      canExport: visibleRows > 0
    };
  }

  /**
   * Форматирование размера файла
   */
  private static formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  /**
   * Предварительный просмотр экспорта (Separate Files)
   */
  public static getExportPreview(
    sheet: IExcelSheet,
    filterState: IFilterState,
    previewRows: number = 5
  ): IExportPreview {
    const visibleRows = sheet.data.filter(row => row.isVisible);
    const sampleRows = visibleRows
      .slice(0, previewRows)
      .map(row => sheet.headers.map(header => row.data[header] as CellValue));

    return {
      headers: [...sheet.headers],
      sampleRows,
      totalRows: visibleRows.length,
      hasMoreData: visibleRows.length > previewRows
    };
  }

  /**
   * Валидация настроек экспорта
   */
  public static validateExportSettings(
    settings: IExportSettings | IRenameExportSettings,
    visibleRowsCount?: number
  ): IExportValidation {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Проверка имени файла
    if (!settings.fileName || settings.fileName.trim().length === 0) {
      errors.push('File name is required');
    } else if (settings.fileName.length > 200) {
      errors.push('File name is too long (maximum 200 characters)');
    } else if (!/^[a-zA-Z0-9._\-\s]+$/.test(settings.fileName)) {
      warnings.push('File name contains special characters that may cause issues');
    }

    // Проверка количества строк
    if (visibleRowsCount !== undefined) {
      if (visibleRowsCount === 0) {
        errors.push('No data to export');
      } else if (visibleRowsCount > 100000) {
        warnings.push('Large dataset detected. Export may take some time.');
      }

      // Проверка ограничения строк (only for IExportSettings)
      if ('maxRows' in settings && settings.maxRows && settings.maxRows > 0 && settings.maxRows > visibleRowsCount) {
        warnings.push(`Max rows setting (${settings.maxRows}) is higher than available data (${visibleRowsCount})`);
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Создание настроек экспорта по умолчанию (Separate Files)
   */
  public static createDefaultExportSettings(originalFileName: string): IExportSettings {
    const baseName = originalFileName.replace(/\.[^/.]+$/, '');
    
    return {
      fileName: `${baseName}_filtered`,
      includeHeaders: true,
      onlyVisibleColumns: true,
      fileFormat: ExportFormat.XLSX
    };
  }

  /**
   * NEW: Create default export settings for Rename Files Management
   */
  public static createDefaultRenameExportSettings(baseName?: string): IRenameExportSettings {
    return {
      fileName: baseName ? `${baseName}_renamed` : 'renamed_files_export',
      includeHeaders: true,
      includeStatusColumn: true,
      includeTimestamps: true,
      onlyCompletedRows: false,
      fileFormat: 'xlsx'
    };
  }
}