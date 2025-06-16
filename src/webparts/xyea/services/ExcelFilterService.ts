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

export class ExcelFilterService {

  /**
   * Анализ колонок и создание метаданных
   */
  public static analyzeColumns(sheet: IExcelSheet): IExcelColumn[] {
    console.log('[ExcelFilterService] Analyzing columns for sheet:', sheet.name);

    const columns: IExcelColumn[] = [];

    sheet.headers.forEach((header, index) => {
      // Получаем все значения для этой колонки
      const columnValues = sheet.data.map(row => row.data[header]);
      
      // Уникальные значения (исключая пустые)
      const uniqueValuesSet = new Set(columnValues.filter(value => 
  value !== null && value !== undefined && value !== ''
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
  ): { filteredSheet: IExcelSheet; statistics: { visible: number; hidden: number } } {
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
    for (const filterName in filterState.filters) {
      const filter = filterState.filters[filterName];
      
      if (filter.isActive) {
        const cellValue = row.data[filter.columnName];
        
        // Проверяем, входит ли значение в выбранные
        if (!this.isValueInSelection(cellValue, filter.selectedValues)) {
          return false; // Если хотя бы один фильтр не пройден, строка скрыта
        }
      }
    }

    return true;
  }

  /**
   * Проверка, входит ли значение в выборку
   */
  private static isValueInSelection(value: any, selectedValues: any[]): boolean {
    // Обработка пустых значений
    if (value === null || value === undefined || value === '') {
      return selectedValues.some(selected => 
        selected === null || selected === undefined || selected === '' || selected === '(Empty)'
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
    selectedValues: any[]
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
  public static getFilterStatistics(filterState: IFilterState): {
    totalFilters: number;
    activeFilters: number;
    totalRows: number;
    filteredRows: number;
    hiddenRows: number;
    filterEfficiency: number;
  } {
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
  private static sortValuesByType(values: any[], dataType: ExcelDataType): any[] {
    const sortedValues = [...values];

    switch (dataType) {
      case ExcelDataType.NUMBER:
        return sortedValues.sort((a, b) => {
          const numA = parseFloat(a);
          const numB = parseFloat(b);
          if (isNaN(numA) && isNaN(numB)) return 0;
          if (isNaN(numA)) return 1;
          if (isNaN(numB)) return -1;
          return numA - numB;
        });

      case ExcelDataType.DATE:
        return sortedValues.sort((a, b) => {
          const dateA = new Date(a);
          const dateB = new Date(b);
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
  ): any[] {
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