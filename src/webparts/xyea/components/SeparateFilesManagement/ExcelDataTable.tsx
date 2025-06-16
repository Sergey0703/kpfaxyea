// src/webparts/xyea/components/SeparateFilesManagement/ExcelDataTable.tsx

import * as React from 'react';
import styles from './ExcelDataTable.module.scss';
import { 
  IExcelSheet, 
  IExcelColumn, 
  IFilterState, 
  ExcelDataType 
} from '../../interfaces/ExcelInterfaces';
import ColumnFilter from './ColumnFilter';

export interface IExcelDataTableProps {
  sheet: IExcelSheet;
  columns: IExcelColumn[];
  filterState: IFilterState;
  onFilterChange: (columnName: string, selectedValues: any[]) => void;
  onClearFilters: () => void;
  loading?: boolean;
  pageSize?: number;
}

export interface IExcelDataTableState {
  currentPage: number;
  openFilterColumn: string | null;
  sortColumn: string | null;
  sortDirection: 'asc' | 'desc';
}

export default class ExcelDataTable extends React.Component<IExcelDataTableProps, IExcelDataTableState> {
  
  constructor(props: IExcelDataTableProps) {
    super(props);
    
    this.state = {
      currentPage: 1,
      openFilterColumn: null,
      sortColumn: null,
      sortDirection: 'asc'
    };
  }

  private handleFilterClick = (columnName: string): void => {
    this.setState(prevState => ({
      openFilterColumn: prevState.openFilterColumn === columnName ? null : columnName
    }));
  }

  private handleFilterClose = (): void => {
    this.setState({ openFilterColumn: null });
  }

  private handleFilterChange = (columnName: string, selectedValues: any[]): void => {
    this.props.onFilterChange(columnName, selectedValues);
    this.setState({ currentPage: 1 }); // Reset to first page after filtering
  }

  private handleSort = (columnName: string): void => {
    this.setState(prevState => {
      const newDirection = prevState.sortColumn === columnName && prevState.sortDirection === 'asc' 
        ? 'desc' 
        : 'asc';
      
      return {
        sortColumn: columnName,
        sortDirection: newDirection
      };
    });
  }

  private handlePageChange = (newPage: number): void => {
    this.setState({ currentPage: newPage });
  }

  private getDataTypeIcon = (dataType: ExcelDataType): string => {
    switch (dataType) {
      case ExcelDataType.NUMBER:
        return '🔢';
      case ExcelDataType.DATE:
        return '📅';
      case ExcelDataType.BOOLEAN:
        return '☑️';
      case ExcelDataType.TEXT:
        return '📝';
      default:
        return '📋';
    }
  }

  private getSortIcon = (columnName: string): string | null => {
    const { sortColumn, sortDirection } = this.state;
    if (sortColumn !== columnName) return null;
    return sortDirection === 'asc' ? '↑' : '↓';
  }

  private getSortedData = (): any[] => {
    const { sheet } = this.props;
    const { sortColumn, sortDirection } = this.state;

    if (!sortColumn) {
      return sheet.data.filter(row => row.isVisible);
    }

    const visibleData = sheet.data.filter(row => row.isVisible);
    
    return [...visibleData].sort((a, b) => {
      const aValue = a.data[sortColumn];
      const bValue = b.data[sortColumn];

      // Handle null/undefined values
      if (aValue === null || aValue === undefined) return 1;
      if (bValue === null || bValue === undefined) return -1;

      // Compare based on data type
      let comparison = 0;
      if (typeof aValue === 'number' && typeof bValue === 'number') {
        comparison = aValue - bValue;
      } else if (aValue instanceof Date && bValue instanceof Date) {
        comparison = aValue.getTime() - bValue.getTime();
      } else {
        comparison = String(aValue).localeCompare(String(bValue));
      }

      return sortDirection === 'asc' ? comparison : -comparison;
    });
  }

  private getPaginatedData = (): any[] => {
    const { pageSize = 50 } = this.props;
    const { currentPage } = this.state;
    
    const sortedData = this.getSortedData();
    const startIndex = (currentPage - 1) * pageSize;
    const endIndex = startIndex + pageSize;
    
    return sortedData.slice(startIndex, endIndex);
  }

  private getTotalPages = (): number => {
    const { pageSize = 50 } = this.props;
    const totalRows = this.getSortedData().length;
    return Math.ceil(totalRows / pageSize);
  }

  private formatCellValue = (value: any, dataType: ExcelDataType): string => {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    switch (dataType) {
      case ExcelDataType.DATE:
        if (value instanceof Date) {
          return value.toLocaleDateString();
        }
        break;
      case ExcelDataType.NUMBER:
        if (typeof value === 'number') {
          return value.toLocaleString();
        }
        break;
      case ExcelDataType.BOOLEAN:
        return value ? 'Yes' : 'No';
      default:
        return String(value);
    }
    
    return String(value);
  }

  private isFilterActive = (columnName: string): boolean => {
    const filter = this.props.filterState.filters[columnName];
    return filter ? filter.isActive : false;
  }

  private getActiveFiltersCount = (): number => {
    return Object.values(this.props.filterState.filters).filter(f => f.isActive).length;
  }

  private renderPagination = (): React.ReactNode => {
    const { currentPage } = this.state;
    const totalPages = this.getTotalPages();
    const totalRows = this.getSortedData().length;
    const { pageSize = 50 } = this.props;
    
    if (totalPages <= 1) return null;

    const startRow = (currentPage - 1) * pageSize + 1;
    const endRow = Math.min(currentPage * pageSize, totalRows);

    return (
      <div className={styles.pagination}>
        <div className={styles.paginationInfo}>
          Showing {startRow}-{endRow} of {totalRows} rows
        </div>
        <div className={styles.paginationControls}>
          <button
            className={styles.paginationButton}
            onClick={() => this.handlePageChange(1)}
            disabled={currentPage === 1}
          >
            ⟨⟨
          </button>
          <button
            className={styles.paginationButton}
            onClick={() => this.handlePageChange(currentPage - 1)}
            disabled={currentPage === 1}
          >
            ⟨
          </button>
          <span className={styles.pageInfo}>
            Page {currentPage} of {totalPages}
          </span>
          <button
            className={styles.paginationButton}
            onClick={() => this.handlePageChange(currentPage + 1)}
            disabled={currentPage === totalPages}
          >
            ⟩
          </button>
          <button
            className={styles.paginationButton}
            onClick={() => this.handlePageChange(totalPages)}
            disabled={currentPage === totalPages}
          >
            ⟩⟩
          </button>
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<IExcelDataTableProps> {
    const { sheet, columns, filterState, loading, onClearFilters } = this.props;
    const { openFilterColumn } = this.state;

    if (loading) {
      return (
        <div className={styles.excelDataTable}>
          <div className={styles.loading}>
            <div className={styles.loadingSpinner}></div>
            <p>Loading data...</p>
          </div>
        </div>
      );
    }

    if (!sheet || !columns.length) {
      return (
        <div className={styles.excelDataTable}>
          <div className={styles.noData}>
            <div className={styles.noDataIcon}>📊</div>
            <p>No data available</p>
          </div>
        </div>
      );
    }

    const paginatedData = this.getPaginatedData();
    const totalVisibleRows = this.getSortedData().length;
    const activeFiltersCount = this.getActiveFiltersCount();

    return (
      <div className={styles.excelDataTable}>
        <div className={styles.tableHeader}>
          <div className={styles.tableInfo}>
            <h3>📊 {sheet.name}</h3>
            <div className={styles.dataStats}>
              <span>Total: {sheet.totalRows} rows</span>
              <span>Visible: {totalVisibleRows} rows</span>
              <span>Columns: {columns.length}</span>
              {activeFiltersCount > 0 && (
                <div className={styles.filterBadge}>
                  {activeFiltersCount} filter{activeFiltersCount > 1 ? 's' : ''} active
                  <button
                    className={styles.clearFiltersButton}
                    onClick={onClearFilters}
                    title="Clear all filters"
                  >
                    ✕
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>

        <div className={styles.tableContainer}>
          <table className={styles.table}>
            <thead className={styles.tableHead}>
              <tr>
                <th className={styles.rowNumberHeader}>#</th>
                {columns.map((column) => {
                  const sortIcon = this.getSortIcon(column.name);
                  const isFilterActive = this.isFilterActive(column.name);
                  
                  return (
                    <th key={column.name} className={styles.headerCell}>
                      <div className={styles.headerContent}>
                        <div className={styles.headerText}>
                          <span className={styles.dataTypeIcon}>
                            {this.getDataTypeIcon(column.dataType)}
                          </span>
                          <span 
                            className={styles.columnName}
                            onClick={() => this.handleSort(column.name)}
                            title={`Sort by ${column.name}`}
                          >
                            {column.name}
                            {sortIcon && (
                              <span className={styles.sortIcon}>{sortIcon}</span>
                            )}
                          </span>
                        </div>
                        <button
                          className={`${styles.filterButton} ${isFilterActive ? styles.active : ''}`}
                          onClick={() => this.handleFilterClick(column.name)}
                          title="Filter column"
                        >
                          ⚲
                        </button>
                      </div>
                      
                      {openFilterColumn === column.name && (
                        <ColumnFilter
                          column={column}
                          filter={filterState.filters[column.name]}
                          isOpen={true}
                          onFilterChange={this.handleFilterChange}
                          onClose={this.handleFilterClose}
                        />
                      )}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody className={styles.tableBody}>
              {paginatedData.map((row, index) => {
                const { currentPage } = this.state;
                const { pageSize = 50 } = this.props;
                const rowNumber = (currentPage - 1) * pageSize + index + 1;
                
                return (
                  <tr key={row.rowIndex} className={styles.tableRow}>
                    <td className={styles.rowNumberCell}>{rowNumber}</td>
                    {columns.map((column) => {
                      const cellValue = row.data[column.name];
                      const formattedValue = this.formatCellValue(cellValue, column.dataType);
                      
                      return (
                        <td 
                          key={column.name} 
                          className={styles.tableCell}
                          title={formattedValue}
                        >
                          {formattedValue}
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {this.renderPagination()}
      </div>
    );
  }
}