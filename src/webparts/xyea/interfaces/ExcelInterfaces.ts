// src/webparts/xyea/interfaces/ExcelInterfaces.ts

/**
 * Интерфейс для загруженного Excel файла
 */
export interface IExcelFile {
  name: string;
  size: number;
  lastModified: Date;
  data: ArrayBuffer;
  sheets: IExcelSheet[];
}

/**
 * Интерфейс для листа Excel
 */
export interface IExcelSheet {
  name: string;
  headers: string[];
  data: IExcelRow[];
  totalRows: number;
  isValid: boolean;
  validationErrors?: string[];
}

/**
 * Интерфейс для строки данных Excel
 */
export interface IExcelRow {
  rowIndex: number;
  data: { [columnName: string]: any };
  isVisible: boolean; // для фильтрации
}

/**
 * Интерфейс для колонки Excel
 */
export interface IExcelColumn {
  name: string;
  index: number;
  dataType: ExcelDataType;
  uniqueValues: any[];
  totalValues: number;
  hasFilter: boolean;
  selectedValues: any[];
}

/**
 * Типы данных в Excel
 */
export enum ExcelDataType {
  TEXT = 'text',
  NUMBER = 'number',
  DATE = 'date',
  BOOLEAN = 'boolean',
  MIXED = 'mixed'
}

/**
 * Интерфейс для фильтра колонки
 */
export interface IColumnFilter {
  columnName: string;
  selectedValues: any[];
  isActive: boolean;
  totalUniqueValues: number;
  dataType: ExcelDataType;
}

/**
 * Интерфейс для состояния фильтров
 */
export interface IFilterState {
  filters: { [columnName: string]: IColumnFilter };
  totalRows: number;
  filteredRows: number;
  isAnyFilterActive: boolean;
}

/**
 * Интерфейс для настроек экспорта
 */
export interface IExportSettings {
  fileName: string;
  includeHeaders: boolean;
  onlyVisibleColumns: boolean;
  maxRows?: number;
  fileFormat: ExportFormat;
}

/**
 * Форматы экспорта
 */
export enum ExportFormat {
  XLSX = 'xlsx',
  CSV = 'csv'
}

/**
 * Интерфейс для результата парсинга
 */
export interface IParseResult {
  success: boolean;
  file?: IExcelFile;
  error?: string;
  warnings?: string[];
  statistics: {
    totalSheets: number;
    totalRows: number;
    totalColumns: number;
    fileSize: string;
    processingTime: number;
  };
}

/**
 * Интерфейс для результата валидации
 */
export interface IValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  suggestions?: string[];
}

/**
 * Интерфейс для прогресса загрузки
 */
export interface IUploadProgress {
  stage: UploadStage;
  progress: number; // 0-100
  message: string;
  isComplete: boolean;
  hasError: boolean;
  error?: string;
}

/**
 * Стадии загрузки
 */
export enum UploadStage {
  IDLE = 'idle',
  UPLOADING = 'uploading',
  PARSING = 'parsing',
  VALIDATING = 'validating',
  ANALYZING = 'analyzing',
  COMPLETE = 'complete',
  ERROR = 'error'
}

/**
 * Интерфейс для аналитики данных
 */
export interface IDataAnalytics {
  columnStats: { [columnName: string]: IColumnStats };
  dataQuality: {
    completeness: number; // % заполненности
    consistency: number; // % консистентности типов
    duplicates: number; // количество дубликатов
  };
  recommendations: string[];
}

/**
 * Статистика по колонке
 */
export interface IColumnStats {
  totalValues: number;
  uniqueValues: number;
  emptyValues: number;
  dataType: ExcelDataType;
  sampleValues: any[];
  minValue?: any;
  maxValue?: any;
  averageValue?: number;
}

/**
 * Интерфейс для пользовательских настроек
 */
export interface IUserPreferences {
  autoDetectDataTypes: boolean;
  showRowNumbers: boolean;
  pageSize: number;
  saveFilterState: boolean;
  defaultExportFormat: ExportFormat;
  maxFileSize: number; // в MB
}

/**
 * Интерфейс для состояния компонента SeparateFilesManagement
 */
export interface ISeparateFilesState {
  // Файл и данные
  uploadedFile: IExcelFile | null;
  currentSheet: IExcelSheet | null;
  columns: IExcelColumn[];
  
  // Фильтрация
  filterState: IFilterState;
  
  // UI состояние
  loading: boolean;
  error: string | null;
  uploadProgress: IUploadProgress;
  
  // Пагинация
  currentPage: number;
  pageSize: number;
  
  // Настройки
  userPreferences: IUserPreferences;
  
  // Экспорт
  isExporting: boolean;
  exportSettings: IExportSettings;
}