// src/webparts/xyea/services/ExcelExportService.ts

import * as XLSX from 'xlsx';
import { 
  IExcelSheet, 
  IFilterState, 
  IExportSettings,
  ExportFormat,
  IExcelRow
} from '../interfaces/ExcelInterfaces';

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
   * Экспорт отфильтрованных данных
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
   * Генерация имени файла на основе фильтров
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
   * Подготовка данных для экспорта
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
   * Экспорт в Excel формат
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
   * Экспорт в CSV формат
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
   * Вычисление оптимальной ширины колонок
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
   * Получение статистики экспорта
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
   * Предварительный просмотр экспорта
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
    settings: IExportSettings,
    visibleRowsCount: number
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
    if (visibleRowsCount === 0) {
      errors.push('No data to export');
    } else if (visibleRowsCount > 100000) {
      warnings.push('Large dataset detected. Export may take some time.');
    }

    // Проверка ограничения строк
    if (settings.maxRows && settings.maxRows > 0 && settings.maxRows > visibleRowsCount) {
      warnings.push(`Max rows setting (${settings.maxRows}) is higher than available data (${visibleRowsCount})`);
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Создание настроек экспорта по умолчанию
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
}