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
        sortColumn: