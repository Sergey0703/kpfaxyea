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
  onFilterChange: (columnName: string, selectedValues: CellValue[]) => void;
  onClose: () => void;
}

export interface IColumnFilterState {
  searchTerm: string;
  filteredValues: CellValue[];
  selectedValues: Set<CellValue>;
  selectAll: boolean;
  position: { top: number; left: number };
  positionClass: 'positionedBelow' | 'positionedAbove';
}

export default class ColumnFilter extends React.Component<IColumnFilterProps, IColumnFilterState> {
  private dropdownRef: React.RefObject<HTMLDivElement>;

  constructor(props: IColumnFilterProps) {
    super(props);
    
    this.state = {
      searchTerm: '',
      filteredValues: [...props.column.uniqueValues],
      selectedValues: new Set(props.filter.selectedValues),
      selectAll: props.filter.selectedValues.length === props.column.uniqueValues.length,
      position: { top: 0, left: 0 },
      positionClass: 'positionedBelow'
    };

    this.dropdownRef = React.createRef<HTMLDivElement>();
  }

  public componentDidMount(): void {
    document.addEventListener('mousedown', this.handleClickOutside);
    this.calculatePosition();
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

    if (this.props.isOpen && !prevProps.isOpen) {
      // Recalculate position when filter opens
      setTimeout(() => this.calculatePosition(), 10);
    }
  }

  private calculatePosition = (): void => {
    // Find the filter button that triggered this dropdown
    const filterButtons = document.querySelectorAll('[data-column-filter]');
    let triggerButton: HTMLElement | null = null;

    // Find the button for our column - use for loop instead of forEach for better type handling
    for (let i = 0; i < filterButtons.length; i++) {
      const button = filterButtons[i];
      if (button.getAttribute('data-column-filter') === this.props.column.name) {
        triggerButton = button as HTMLElement;
        break;
      }
    }

    if (!triggerButton) {
      // Fallback positioning if we can't find the trigger
      this.setState({
        position: { top: 100, left: 100 },
        positionClass: 'positionedBelow'
      });
      return;
    }

    const buttonRect = triggerButton.getBoundingClientRect();
    const viewportHeight = window.innerHeight;
    const viewportWidth = window.innerWidth;
    
    // Dropdown dimensions (approximate)
    const dropdownHeight = 400;
    const dropdownWidth = 300;

    // Calculate vertical position
    const spaceBelow = viewportHeight - buttonRect.bottom;
    const spaceAbove = buttonRect.top;
    
    let top: number;
    let positionClass: 'positionedBelow' | 'positionedAbove';

    if (spaceBelow >= dropdownHeight || spaceBelow > spaceAbove) {
      // Position below the button
      top = buttonRect.bottom + 5;
      positionClass = 'positionedBelow';
    } else {
      // Position above the button
      top = buttonRect.top - dropdownHeight - 5;
      positionClass = 'positionedAbove';
    }

    // Calculate horizontal position
    let left = buttonRect.left;

    // Ensure dropdown doesn't go off-screen horizontally
    if (left + dropdownWidth > viewportWidth) {
      left = viewportWidth - dropdownWidth - 10;
    }
    if (left < 10) {
      left = 10;
    }

    // Ensure dropdown doesn't go off-screen vertically
    if (top < 10) {
      top = 10;
    }
    if (top + dropdownHeight > viewportHeight - 10) {
      top = viewportHeight - dropdownHeight - 10;
    }

    this.setState({
      position: { top, left },
      positionClass
    });
  }

  private handleClickOutside = (event: MouseEvent): void => {
    if (this.dropdownRef.current && !this.dropdownRef.current.contains(event.target as Node)) {
      // Also check if the click was on the filter button to prevent immediate close/open
      const target = event.target as Element;
      const isFilterButton = target?.closest('[data-column-filter]');
      
      if (!isFilterButton || isFilterButton.getAttribute('data-column-filter') !== this.props.column.name) {
        this.props.onClose();
      }
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
    const newSelectedValues = new Set(this.state.selectedValues);

    if (checked) {
      // Add all filtered values
      filteredValues.forEach(value => newSelectedValues.add(value));
    } else {
      // Remove all filtered values
      filteredValues.forEach(value => newSelectedValues.delete(value));
    }

    this.setState({
      selectedValues: newSelectedValues,
      selectAll: checked && filteredValues.length === this.props.column.uniqueValues.length
    });
  }

  private handleValueToggle = (value: CellValue, checked: boolean): void => {
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

  private formatValue = (value: CellValue): string => {
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

  public render(): React.ReactElement<IColumnFilterProps> | undefined {
    if (!this.props.isOpen) {
      return undefined;
    }

    const { column } = this.props;
    const { searchTerm, filteredValues, selectedValues, position, positionClass } = this.state;

    const isAllFilteredSelected = filteredValues.every(value => selectedValues.has(value));
    const isIndeterminate = filteredValues.some(value => selectedValues.has(value)) && !isAllFilteredSelected;

    return (
      <div 
        ref={this.dropdownRef} 
        className={`${styles.columnFilter} ${styles[positionClass]}`}
        style={{
          top: `${position.top}px`,
          left: `${position.left}px`
        }}
      >
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