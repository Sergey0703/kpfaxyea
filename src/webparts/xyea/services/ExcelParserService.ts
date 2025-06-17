// src/webparts/xyea/services/ExcelParserService.ts

import * as XLSX from 'xlsx';
import { 
  IExcelFile, 
  IExcelSheet, 
  IExcelRow, 
  IParseResult, 
  IValidationResult,
  ExcelDataType,
  UploadStage 
} from '../interfaces/ExcelInterfaces';

// Define proper types instead of any
type CellValue = string | number | boolean | Date | undefined;
type ExcelRowArray = (string | number | boolean | Date | undefined)[];

export class ExcelParserService {
  private static readonly MAX_FILE_SIZE = 10 * 1024 * 1024; // 10MB
  private static readonly SUPPORTED_FORMATS = [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
    'application/vnd.ms-excel', // .xls
    'text/csv' // .csv
  ];

  /**
   * Парсинг загруженного файла
   */
  public static async parseFile(
    file: File,
    onProgress?: (stage: UploadStage, progress: number, message: string) => void
  ): Promise<IParseResult> {
    const startTime = Date.now();
    
    try {
      console.log('[ExcelParserService] Starting file parsing:', {
        name: file.name,
        size: file.size,
        type: file.type
      });

      // Валидация файла
      onProgress?.(UploadStage.VALIDATING, 10, 'Validating file...');
      const validation = this.validateFile(file);
      if (!validation.isValid) {
        return {
          success: false,
          error: validation.errors[0],
          statistics: this.createEmptyStatistics(Date.now() - startTime)
        };
      }

      // Чтение файла
      onProgress?.(UploadStage.UPLOADING, 30, 'Reading file...');
      const arrayBuffer = await this.readFileAsArrayBuffer(file);

      // Парсинг Excel
      onProgress?.(UploadStage.PARSING, 50, 'Parsing Excel data...');
      const workbook = XLSX.read(arrayBuffer, { 
        type: 'array',
        cellStyles: true,
        cellDates: true,
        dateNF: 'yyyy-mm-dd'
      });

      // Анализ листов
      onProgress?.(UploadStage.ANALYZING, 70, 'Analyzing sheets...');
      const sheets = this.parseWorkbook(workbook);

      if (sheets.length === 0) {
        return {
          success: false,
          error: 'No valid sheets found in the file',
          statistics: this.createEmptyStatistics(Date.now() - startTime)
        };
      }

      // Создание результата
      onProgress?.(UploadStage.COMPLETE, 100, 'Parsing complete!');
      
      const excelFile: IExcelFile = {
        name: file.name,
        size: file.size,
        lastModified: new Date(file.lastModified),
        data: arrayBuffer,
        sheets
      };

      const processingTime = Date.now() - startTime;
      console.log('[ExcelParserService] Parsing completed successfully:', {
        sheets: sheets.length,
        totalRows: sheets.reduce((sum, sheet) => sum + sheet.totalRows, 0),
        processingTime: `${processingTime}ms`
      });

      return {
        success: true,
        file: excelFile,
        statistics: {
          totalSheets: sheets.length,
          totalRows: sheets.reduce((sum, sheet) => sum + sheet.totalRows, 0),
          totalColumns: sheets[0]?.headers.length || 0,
          fileSize: this.formatFileSize(file.size),
          processingTime
        }
      };

    } catch (error) {
      console.error('[ExcelParserService] Parsing failed:', error);
      return {
        success: false,
        error: `Failed to parse file: ${error instanceof Error ? error.message : 'Unknown error'}`,
        statistics: this.createEmptyStatistics(Date.now() - startTime)
      };
    }
  }

  /**
   * Валидация файла
   */
  private static validateFile(file: File): IValidationResult {
    const errors: string[] = [];
    const warnings: string[] = [];

    // Проверка размера
    if (file.size > this.MAX_FILE_SIZE) {
      errors.push(`File size (${this.formatFileSize(file.size)}) exceeds maximum allowed size (${this.formatFileSize(this.MAX_FILE_SIZE)})`);
    }

    // Проверка формата
    if (!this.SUPPORTED_FORMATS.includes(file.type) && !file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
      errors.push(`Unsupported file format. Supported formats: .xlsx, .xls, .csv`);
    }

    // Проверка имени файла
    if (file.name.length === 0) {
      errors.push('File name is empty');
    }

    // Предупреждения
    if (file.size > 5 * 1024 * 1024) { // 5MB
      warnings.push('Large file detected. Processing may take longer.');
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Чтение файла как ArrayBuffer
   */
  private static readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
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

  /**
   * Парсинг workbook
   */
  private static parseWorkbook(workbook: XLSX.WorkBook): IExcelSheet[] {
    const sheets: IExcelSheet[] = [];

    // Берем только первый лист согласно требованиям
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      return sheets;
    }

    console.log('[ExcelParserService] Parsing first sheet:', firstSheetName);
    
    const worksheet = workbook.Sheets[firstSheetName];
    const sheetData = this.parseWorksheet(worksheet, firstSheetName);
    
    if (sheetData) {
      sheets.push(sheetData);
    }

    return sheets;
  }

  /**
   * Парсинг отдельного листа
   */
  private static parseWorksheet(worksheet: XLSX.WorkSheet, sheetName: string): IExcelSheet | null {
    try {
      // Получаем диапазон данных
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
      
      if (range.e.r < 1) {
        console.warn('[ExcelParserService] Sheet has no data rows:', sheetName);
        return null;
      }

      // Конвертируем в JSON с заголовками
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        defval: '',
        raw: false,
        dateNF: 'yyyy-mm-dd'
      }) as ExcelRowArray[]; // Use proper type instead of any[][]

      if (jsonData.length === 0) {
        return null;
      }

      // Первая строка - заголовки
      const headers = jsonData[0]?.map((header, index) => 
        header ? String(header).trim() : `Column_${index + 1}`
      ) || [];

      if (headers.length === 0) {
        return null;
      }

      // Остальные строки - данные
      const dataRows = jsonData.slice(1);
      const rows: IExcelRow[] = [];

      dataRows.forEach((row, index) => {
        const rowData: { [key: string]: CellValue } = {}; // Use proper type
        
        headers.forEach((header, colIndex) => {
          rowData[header] = row[colIndex] || undefined; // Use undefined instead of empty string
        });

        rows.push({
          rowIndex: index + 2, // +2 потому что +1 для заголовка и +1 для Excel нумерации с 1
          data: rowData,
          isVisible: true
        });
      });

      console.log('[ExcelParserService] Sheet parsed successfully:', {
        name: sheetName,
        headers: headers.length,
        rows: rows.length
      });

      return {
        name: sheetName,
        headers,
        data: rows,
        totalRows: rows.length,
        isValid: true
      };

    } catch (error) {
      console.error('[ExcelParserService] Failed to parse worksheet:', sheetName, error);
      return {
        name: sheetName,
        headers: [],
        data: [],
        totalRows: 0,
        isValid: false,
        validationErrors: [`Failed to parse sheet: ${error instanceof Error ? error.message : 'Unknown error'}`]
      };
    }
  }

  /**
   * Определение типа данных в колонке
   */
  public static detectColumnDataType(values: CellValue[]): ExcelDataType { // Use proper type
    if (values.length === 0) {
      return ExcelDataType.TEXT;
    }

    const nonEmptyValues = values.filter(v => v !== undefined && v !== '');
    if (nonEmptyValues.length === 0) {
      return ExcelDataType.TEXT;
    }

    let numberCount = 0;
    let dateCount = 0;
    let booleanCount = 0;

    nonEmptyValues.forEach(value => {
      if (typeof value === 'number' && !isNaN(value)) {
        numberCount++;
      } else if (typeof value === 'boolean') {
        booleanCount++;
      } else if (this.isDateValue(value)) {
        dateCount++;
      }
    });

    const total = nonEmptyValues.length;
    const threshold = 0.8; // 80% значений должны быть одного типа

    if (numberCount / total >= threshold) {
      return ExcelDataType.NUMBER;
    } else if (dateCount / total >= threshold) {
      return ExcelDataType.DATE;
    } else if (booleanCount / total >= threshold) {
      return ExcelDataType.BOOLEAN;
    } else if ((numberCount + dateCount + booleanCount) / total >= 0.5) {
      return ExcelDataType.MIXED;
    } else {
      return ExcelDataType.TEXT;
    }
  }

  /**
   * Проверка, является ли значение датой
   */
  private static isDateValue(value: CellValue): boolean { // Use proper type and explicit return type
    if (value instanceof Date) {
      return true;
    }
    
    if (typeof value === 'string') {
      // Проверяем различные форматы дат
      const datePatterns = [
        /^\d{4}-\d{2}-\d{2}$/, // YYYY-MM-DD
        /^\d{2}\/\d{2}\/\d{4}$/, // MM/DD/YYYY
        /^\d{2}\.\d{2}\.\d{4}$/, // DD.MM.YYYY
      ];
      
      return datePatterns.some(pattern => pattern.test(value)) && !isNaN(Date.parse(value));
    }
    
    return false;
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
   * Создание пустой статистики
   */
  private static createEmptyStatistics(processingTime: number): {
    totalSheets: number;
    totalRows: number;
    totalColumns: number;
    fileSize: string;
    processingTime: number;
  } {
    return {
      totalSheets: 0,
      totalRows: 0,
      totalColumns: 0,
      fileSize: '0 Bytes',
      processingTime
    };
  }
}