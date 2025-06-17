// src/webparts/xyea/components/SeparateFilesManagement/ColumnFilter.tsx

import * as React from 'react';
import styles from './ColumnFilter.module.scss';
import { IColumnFilter, ExcelDataType, IExcelColumn } from '../../interfaces/ExcelInterfaces';
import { ExcelFilterService } from '../../services/ExcelFilterService';

// Define proper types instead of any
type CellValue = string | number | boolean | Date | undefined;

export interface IColumnFilterProps {
  column: IExcelColumn;
  filter: IColumnFilter;
  isOpen: boolean;
  onFilterChange: (columnName: string, selectedValues: CellValue[]) => void; // Use specific type instead of any
  onClose: () => void;
}

export interface IColumnFilterState {
  searchTerm: string;
  filteredValues: CellValue[]; // Use specific type instead of any
  selectedValues: Set<CellValue>; // Use specific type instead of any
  selectAll: boolean;
}

export default class ColumnFilter extends React.Component<IColumnFilterProps, IColumnFilterState> {
  private dropdownRef: React.RefObject<HTMLDivElement>;

  constructor(props: IColumnFilterProps) {
    super(props);
    
    this.state = {
      searchTerm: '',
      filteredValues: [...props.column.uniqueValues],
      selectedValues: new Set(props.filter.selectedValues),
      selectAll: props.filter.selectedValues.length === props.column.uniqueValues.length
    };

    this.dropdownRef = React.createRef<HTMLDivElement>();
  }

  public componentDidMount(): void {
    document.addEventListener('mousedown', this.handleClickOutside);
  }

  public componentWillUnmount(): void {
    document.removeEventListener('mousedown', this.handleClickOutside);
  }

  public componentDidUpdate(prevProps: IColumnFilterProps): void {
    if (prevProps.filter !== this.props.filter) {
      this.setState({
        selectedValues: new Set(this.props.filter.selectedValues),
        selectAll: this.props.filter.selectedValues.length === this.props.column.uniqueValues.length
      });
    }
  }

  private handleClickOutside = (event: MouseEvent): void => {
    if (this.dropdownRef.current && !this.dropdownRef.current.contains(event.target as Node)) {
      this.props.onClose();
    }
  }

  private handleSearchChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const searchTerm = event.target.value;
    const filteredValues = ExcelFilterService.searchColumnValues(this.props.column, searchTerm);
    
    this.setState({
      searchTerm,
      filteredValues
    });
  }

  private handleSelectAll = (checked: boolean): void => {
    const { filteredValues } = this.state;
    const newSelectedValues = new Set(this.state.selectedValues); // Use const instead of let

    if (checked) {
      // Добавляем все отфильтрованные значения
      filteredValues.forEach(value => newSelectedValues.add(value));
    } else {
      // Убираем все отфильтрованные значения
      filteredValues.forEach(value => newSelectedValues.delete(value));
    }

    this.setState({
      selectedValues: newSelectedValues,
      selectAll: checked && filteredValues.length === this.props.column.uniqueValues.length
    });
  }

  private handleValueToggle = (value: CellValue, checked: boolean): void => { // Use specific type instead of any
    const newSelectedValues = new Set(this.state.selectedValues);

    if (checked) {
      newSelectedValues.add(value);
    } else {
      newSelectedValues.delete(value);
    }

    const selectAll = newSelectedValues.size === this.props.column.uniqueValues.length;

    this.setState({
      selectedValues: newSelectedValues,
      selectAll
    });
  }

  private handleApply = (): void => {
    const selectedArray = Array.from(this.state.selectedValues);
    this.props.onFilterChange(this.props.column.name, selectedArray);
    this.props.onClose();
  }

  private handleClear = (): void => {
    this.setState({
      selectedValues: new Set(),
      selectAll: false,
      searchTerm: '',
      filteredValues: [...this.props.column.uniqueValues]
    });
  }

  private handleSelectAllValues = (): void => {
    this.setState({
      selectedValues: new Set(this.props.column.uniqueValues),
      selectAll: true
    });
  }

  private formatValue = (value: CellValue): string => { // Use specific type instead of any
    if (value === undefined || value === '') {
      return '(Empty)';
    }
    
    if (this.props.column.dataType === ExcelDataType.DATE && value instanceof Date) {
      return value.toLocaleDateString();
    }
    
    if (this.props.column.dataType === ExcelDataType.NUMBER && typeof value === 'number') {
      return value.toLocaleString();
    }
    
    return String(value);
  }

  private getDataTypeIcon = (): string => {
    switch (this.props.column.dataType) {
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

  public render(): React.ReactElement<IColumnFilterProps> | undefined { // Changed from null to undefined
    if (!this.props.isOpen) {
      return undefined;
    }

    const { column } = this.props;
    const { searchTerm, filteredValues, selectedValues } = this.state;

    const isAllFilteredSelected = filteredValues.every(value => selectedValues.has(value));
    const isIndeterminate = filteredValues.some(value => selectedValues.has(value)) && !isAllFilteredSelected;

    return (
      <div ref={this.dropdownRef} className={styles.columnFilter}>
        <div className={styles.header}>
          <div className={styles.columnInfo}>
            <span className={styles.dataTypeIcon}>{this.getDataTypeIcon()}</span>
            <span className={styles.columnName}>{column.name}</span>
          </div>
          <button className={styles.closeButton} onClick={this.props.onClose}>
            ✕
          </button>
        </div>

        <div className={styles.searchContainer}>
          <input
            type="text"
            className={styles.searchInput}
            placeholder="Search values..."
            value={searchTerm}
            onChange={this.handleSearchChange}
          />
        </div>

        <div className={styles.selectAllContainer}>
          <label className={styles.checkboxLabel}>
            <input
              type="checkbox"
              checked={isAllFilteredSelected}
              ref={input => {
                if (input) input.indeterminate = isIndeterminate;
              }}
              onChange={(e) => this.handleSelectAll(e.target.checked)}
              className={styles.checkbox}
            />
            <span className={styles.checkboxText}>
              Select All ({filteredValues.length} items)
            </span>
          </label>
        </div>

        <div className={styles.valuesList}>
          {filteredValues.map((value, index) => {
            const isSelected = selectedValues.has(value);
            const displayValue = this.formatValue(value);
            
            return (
              <label key={index} className={styles.valueItem}>
                <input
                  type="checkbox"
                  checked={isSelected}
                  onChange={(e) => this.handleValueToggle(value, e.target.checked)}
                  className={styles.checkbox}
                />
                <span className={styles.valueText} title={displayValue}>
                  {displayValue}
                </span>
              </label>
            );
          })}
        </div>

        <div className={styles.footer}>
          <div className={styles.statistics}>
            Selected: {selectedValues.size} of {column.uniqueValues.length}
          </div>
          <div className={styles.actions}>
            <button className={styles.actionButton} onClick={this.handleClear}>
              Clear
            </button>
            <button className={styles.actionButton} onClick={this.handleSelectAllValues}>
              All
            </button>
            <button className={`${styles.actionButton} ${styles.primary}`} onClick={this.handleApply}>
              Apply
            </button>
          </div>
        </div>
      </div>
    );
  }
}