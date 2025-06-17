// src/webparts/xyea/services/ExcelFilterService.ts

import { 
  IExcelSheet, 
  IExcelRow, 
  IExcelColumn, 
  IColumnFilter, 
  IFilterState,
  ExcelDataType 
} from '../interfaces/ExcelInterfaces';
import { ExcelParserService } from './ExcelParserService';

// Type definitions for better type safety
type CellValue = string | number | boolean | Date | undefined; // Changed from null to undefined

interface IFilterStatistics {
  totalFilters: number;
  activeFilters: number;
  totalRows: number;
  filteredRows: number;
  hiddenRows: number;
  filterEfficiency: number;
}

interface IApplyFiltersResult {
  filteredSheet: IExcelSheet;
  statistics: {
    visible: number;
    hidden: number;
  };
}

export class ExcelFilterService {

  /**
   * Анализ колонок и создание метаданных
   */
  public static analyzeColumns(sheet: IExcelSheet): IExcelColumn[] {
    console.log('[ExcelFilterService] Analyzing columns for sheet:', sheet.name);

    const columns: IExcelColumn[] = [];

    sheet.headers.forEach((header, index) => {
      // Получаем все значения для этой колонки
      const columnValues: CellValue[] = sheet.data.map(row => row.data[header]);
      
      // Уникальные значения (исключая пустые)
      const uniqueValuesSet = new Set(columnValues.filter((value: CellValue) => 
        value !== undefined && value !== ''
      ));
      const uniqueValues = Array.from(uniqueValuesSet);
      
      // Определяем тип данных
      const dataType = ExcelParserService.detectColumnDataType(columnValues);

      // Сортируем уникальные значения
      const sortedUniqueValues = this.sortValuesByType(uniqueValues, dataType);

      const column: IExcelColumn = {
        name: header,
        index,
        dataType,
        uniqueValues: sortedUniqueValues,
        totalValues: columnValues.length,
        hasFilter: false,
        selectedValues: [...sortedUniqueValues] // Изначально все значения выбраны
      };

      columns.push(column);

      console.log('[ExcelFilterService] Column analyzed:', {
        name: header,
        dataType,
        uniqueCount: uniqueValues.length,
        totalCount: columnValues.length
      });
    });

    return columns;
  }

  /**
   * Создание начального состояния фильтров
   */
  public static createInitialFilterState(columns: IExcelColumn[], totalRows: number): IFilterState {
    const filters: { [columnName: string]: IColumnFilter } = {};

    columns.forEach(column => {
      filters[column.name] = {
        columnName: column.name,
        selectedValues: [...column.uniqueValues], // Все значения выбраны
        isActive: false,
        totalUniqueValues: column.uniqueValues.length,
        dataType: column.dataType
      };
    });

    return {
      filters,
      totalRows,
      filteredRows: totalRows,
      isAnyFilterActive: false
    };
  }

  /**
   * Применение фильтров к данным
   */
  public static applyFilters(
    sheet: IExcelSheet, 
    filterState: IFilterState
  ): IApplyFiltersResult {
    console.log('[ExcelFilterService] Applying filters to sheet:', sheet.name);

    let visibleCount = 0;
    let hiddenCount = 0;

    // Создаем копию листа
    const filteredSheet: IExcelSheet = {
      ...sheet,
      data: sheet.data.map(row => {
        const isVisible = this.isRowVisible(row, filterState);
        
        if (isVisible) {
          visibleCount++;
        } else {
          hiddenCount++;
        }

        return {
          ...row,
          isVisible
        };
      })
    };

    console.log('[ExcelFilterService] Filter results:', {
      visible: visibleCount,
      hidden: hiddenCount,
      total: visibleCount + hiddenCount
    });

    return {
      filteredSheet,
      statistics: {
        visible: visibleCount,
        hidden: hiddenCount
      }
    };
  }

  /**
   * Проверка видимости строки на основе фильтров
   */
  private static isRowVisible(row: IExcelRow, filterState: IFilterState): boolean {
    // Если нет активных фильтров, показываем все строки
    if (!filterState.isAnyFilterActive) {
      return true;
    }

    // Проверяем каждый активный фильтр
    const filterNames = Object.keys(filterState.filters);
    for (const filterName of filterNames) {
      // Use Object.prototype.hasOwnProperty to ensure we only check own properties
      if (Object.prototype.hasOwnProperty.call(filterState.filters, filterName)) {
        const filter = filterState.filters[filterName];
        
        if (filter.isActive) {
          const cellValue = row.data[filter.columnName];
          
          // Проверяем, входит ли значение в выбранные
          if (!this.isValueInSelection(cellValue, filter.selectedValues)) {
            return false; // Если хотя бы один фильтр не пройден, строка скрыта
          }
        }
      }
    }

    return true;
  }

  /**
   * Проверка, входит ли значение в выборку
   */
  private static isValueInSelection(value: CellValue, selectedValues: CellValue[]): boolean {
    // Обработка пустых значений
    if (value === undefined || value === '') {
      return selectedValues.some(selected => 
        selected === undefined || selected === '' || selected === '(Empty)'
      );
    }

    // Прямое сравнение
    return selectedValues.includes(value);
  }

  /**
   * Обновление фильтра для колонки
   */
  public static updateColumnFilter(
    filterState: IFilterState,
    columnName: string,
    selectedValues: CellValue[]
  ): IFilterState {
    const filter = filterState.filters[columnName];
    if (!filter) {
      console.warn('[ExcelFilterService] Filter not found for column:', columnName);
      return filterState;
    }

    // Создаем новое состояние фильтра
    const updatedFilter: IColumnFilter = {
      ...filter,
      selectedValues: [...selectedValues],
      isActive: selectedValues.length < filter.totalUniqueValues // Активен, если выбрано не все
    };

    // Создаем новое состояние фильтров
    const updatedFilters = {
      ...filterState.filters,
      [columnName]: updatedFilter
    };

    // Проверяем, есть ли активные фильтры
    const isAnyFilterActive = Object.values(updatedFilters).some(f => f.isActive);

    console.log('[ExcelFilterService] Filter updated:', {
      column: columnName,
      selectedCount: selectedValues.length,
      totalCount: filter.totalUniqueValues,
      isActive: updatedFilter.isActive,
      anyActive: isAnyFilterActive
    });

    return {
      ...filterState,
      filters: updatedFilters,
      isAnyFilterActive
    };
  }

  /**
   * Сброс всех фильтров
   */
  public static clearAllFilters(filterState: IFilterState, columns: IExcelColumn[]): IFilterState {
    console.log('[ExcelFilterService] Clearing all filters');

    const clearedFilters: { [columnName: string]: IColumnFilter } = {};

    columns.forEach(column => {
      clearedFilters[column.name] = {
        columnName: column.name,
        selectedValues: [...column.uniqueValues],
        isActive: false,
        totalUniqueValues: column.uniqueValues.length,
        dataType: column.dataType
      };
    });

    return {
      ...filterState,
      filters: clearedFilters,
      isAnyFilterActive: false
    };
  }

  /**
   * Получение статистики фильтрации
   */
  public static getFilterStatistics(filterState: IFilterState): IFilterStatistics {
    const totalFilters = Object.keys(filterState.filters).length;
    const activeFilters = Object.values(filterState.filters).filter(f => f.isActive).length;
    const hiddenRows = filterState.totalRows - filterState.filteredRows;
    const filterEfficiency = filterState.totalRows > 0 ? 
      (filterState.filteredRows / filterState.totalRows) * 100 : 0;

    return {
      totalFilters,
      activeFilters,
      totalRows: filterState.totalRows,
      filteredRows: filterState.filteredRows,
      hiddenRows,
      filterEfficiency: Math.round(filterEfficiency * 100) / 100
    };
  }

  /**
   * Сортировка значений по типу данных
   */
  private static sortValuesByType(values: CellValue[], dataType: ExcelDataType): CellValue[] {
    const sortedValues = [...values];

    switch (dataType) {
      case ExcelDataType.NUMBER:
        return sortedValues.sort((a, b) => {
          const numA = parseFloat(String(a));
          const numB = parseFloat(String(b));
          if (isNaN(numA) && isNaN(numB)) return 0;
          if (isNaN(numA)) return 1;
          if (isNaN(numB)) return -1;
          return numA - numB;
        });

      case ExcelDataType.DATE:
        return sortedValues.sort((a, b) => {
          const dateA = new Date(String(a));
          const dateB = new Date(String(b));
          if (isNaN(dateA.getTime()) && isNaN(dateB.getTime())) return 0;
          if (isNaN(dateA.getTime())) return 1;
          if (isNaN(dateB.getTime())) return -1;
          return dateA.getTime() - dateB.getTime();
        });

      case ExcelDataType.BOOLEAN:
        return sortedValues.sort((a, b) => {
          // false сначала, потом true
          if (a === b) return 0;
          if (a === false || a === 'false') return -1;
          if (b === false || b === 'false') return 1;
          return 0;
        });

      default: // TEXT и MIXED
        return sortedValues.sort((a, b) => {
          const strA = String(a).toLowerCase();
          const strB = String(b).toLowerCase();
          return strA.localeCompare(strB);
        });
    }
  }

  /**
   * Поиск значений в колонке
   */
  public static searchColumnValues(
    column: IExcelColumn, 
    searchTerm: string
  ): CellValue[] {
    if (!searchTerm.trim()) {
      return column.uniqueValues;
    }

    const lowerSearchTerm = searchTerm.toLowerCase();
    
    return column.uniqueValues.filter(value => {
      const strValue = String(value).toLowerCase();
      return strValue.includes(lowerSearchTerm);
    });
  }
}