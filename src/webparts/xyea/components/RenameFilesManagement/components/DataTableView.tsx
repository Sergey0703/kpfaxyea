// src/webparts/xyea/components/RenameFilesManagement/components/DataTableView.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { 
  IRenameFilesData,
  DirectoryStatus,
  FileSearchStatus,
  FileStatusHelper
} from '../types/RenameFilesTypes';
import { ColumnResizeHandler } from '../handlers/ColumnResizeHandler';

export interface IDataTableViewProps {
  data: IRenameFilesData;
  directoryResults: { [rowIndex: number]: DirectoryStatus }; // NEW: Directory status
  fileSearchResults: { [rowIndex: number]: FileSearchStatus }; // UPDATED: Only file search status
  columnResizeHandler: ColumnResizeHandler;
  onCellEdit: (columnId: string, rowIndex: number, newValue: string) => void;
}

export const DataTableView: React.FC<IDataTableViewProps> = ({
  data,
  directoryResults, // NEW: Directory status
  fileSearchResults, // UPDATED: File search status
  columnResizeHandler,
  onCellEdit
}) => {
  const handleCellChange = (columnId: string, rowIndex: number, value: string) => {
    onCellEdit(columnId, rowIndex, value);
  };

  /**
   * NEW: Render directory status indicator (shows after Stage 2)
   */
  const renderDirectoryIndicator = (rowIndex: number) => {
    const directoryStatus = directoryResults[rowIndex];
    
    if (!directoryStatus) return null;

    const icon = FileStatusHelper.getDirectoryIcon(directoryStatus);
    const tooltip = FileStatusHelper.getDirectoryTooltipText(directoryStatus);

    if (!icon) return null;

    return (
      <span 
        className={getDirectoryIndicatorClass(directoryStatus)} 
        title={tooltip}
      >
        {icon}
      </span>
    );
  };

  /**
   * UPDATED: Render file search indicator (shows after Stage 3)
   */
  const renderFileSearchIndicator = (rowIndex: number) => {
    const fileSearchStatus = fileSearchResults[rowIndex];
    
    if (!fileSearchStatus) return null;

    const icon = FileStatusHelper.getFileSearchIcon(fileSearchStatus);
    const tooltip = FileStatusHelper.getFileTooltipText(fileSearchStatus);

    if (!icon) return null;

    return (
      <span 
        className={getFileSearchIndicatorClass(fileSearchStatus)} 
        title={tooltip}
      >
        {icon}
      </span>
    );
  };

  /**
   * NEW: Get CSS class for directory indicator
   */
  const getDirectoryIndicatorClass = (status: DirectoryStatus): string => {
    switch (status) {
      case 'checking':
        return styles.checkingDirectoryIndicator;
      case 'exists':
        return styles.directoryExistsIndicator;
      case 'not-exists':
        return styles.directoryNotExistsIndicator;
      case 'error':
        return styles.directoryErrorIndicator;
      default:
        return '';
    }
  };

  /**
   * UPDATED: Get CSS class for file search indicator
   */
  const getFileSearchIndicatorClass = (status: FileSearchStatus): string => {
    switch (status) {
      case 'searching':
        return styles.searchingIndicator;
      case 'found':
        return styles.foundIndicator;
      case 'not-found':
        return styles.notFoundIndicator;
      case 'skipped':
        return styles.skippedIndicator;
      default:
        return '';
    }
  };

  /**
   * NEW: Render combined status indicators (directory + file)
   */
  const renderStatusIndicators = (rowIndex: number) => {
    const directoryStatus = directoryResults[rowIndex];
    const fileSearchStatus = fileSearchResults[rowIndex];
    
    // Show priority: File status (if available) > Directory status
    if (fileSearchStatus) {
      // Stage 3 completed: Show file search result
      return renderFileSearchIndicator(rowIndex);
    } else if (directoryStatus) {
      // Stage 2 completed but Stage 3 not started: Show directory status
      return renderDirectoryIndicator(rowIndex);
    }
    
    // No status available
    return null;
  };



  /**
   * NEW: Determine if row should have special styling based on status
   */
  const getRowStatusClass = (rowIndex: number): string => {
    const directoryStatus = directoryResults[rowIndex];
    const fileSearchStatus = fileSearchResults[rowIndex];
    
    // Apply styling based on current status
    if (fileSearchStatus === 'found') {
      return styles.fileFoundRow;
    } else if (fileSearchStatus === 'not-found') {
      return styles.fileNotFoundRow;
    } else if (directoryStatus === 'not-exists') {
      return styles.directoryNotExistsRow;
    } else if (directoryStatus === 'error') {
      return styles.directoryErrorRow;
    }
    
    return '';
  };

  return (
    <div className={styles.tableContainer}>
      <table className={styles.dataTable}>
        <thead>
          <tr>
            <th className={styles.rowHeader}>#</th>
            {data.columns
              .sort((a, b) => a.currentIndex - b.currentIndex)
              .filter(col => col.isVisible)
              .map(column => (
                <th 
                  key={column.id} 
                  className={`${styles.columnHeader} ${column.isCustom ? styles.customColumn : styles.excelColumn}`}
                  style={{ width: column.width }}
                  data-column-id={column.id}
                >
                  <div className={styles.headerContent}>
                    <span className={styles.columnName}>{column.name}</span>
                    {column.isCustom && (
                      <span className={styles.customBadge}>Custom</span>
                    )}
                  </div>
                  <div 
                    className={styles.resizeHandle}
                    onMouseDown={(e) => columnResizeHandler.handleResizeStart(column.id, e)}
                    title="Drag to resize column"
                  />
                </th>
              ))}
          </tr>
        </thead>
        <tbody>
          {data.rows.map(row => {
            // NEW: Apply row styling based on status
            const rowStatusClass = getRowStatusClass(row.rowIndex);
            const baseRowClass = row.isEdited ? styles.editedRow : '';
            const finalRowClass = [baseRowClass, rowStatusClass].filter(Boolean).join(' ');
            
            return (
              <tr key={row.rowIndex} className={finalRowClass}>
                <td className={styles.rowNumber}>
                  <div className={styles.rowNumberContent}>
                    <span className={styles.rowNumberText}>{row.rowIndex + 1}</span>
                    {/* NEW: Show combined status indicators */}
                    <div className={styles.searchIndicator}>
                      {renderStatusIndicators(row.rowIndex)}
                    </div>
                  </div>
                </td>
                {data.columns
                  .sort((a, b) => a.currentIndex - b.currentIndex)
                  .filter(col => col.isVisible)
                  .map(column => {
                    const cell = row.cells[column.id];
                    return (
                      <td 
                        key={`${column.id}_${row.rowIndex}`}
                        className={`${styles.tableCell} ${cell?.isEdited ? styles.editedCell : ''}`}
                      >
                        <input
                          type="text"
                          value={String(cell?.value || '')}
                          onChange={(e) => handleCellChange(column.id, row.rowIndex, e.target.value)}
                          className={styles.cellInput}
                          placeholder={column.isCustom ? 'Enter value...' : ''}
                        />
                      </td>
                    );
                  })}
              </tr>
            );
          })}
        </tbody>
      </table>
      
      {/* NEW: Status legend for better user understanding */}
      <div className={styles.statusLegend}>
        <div className={styles.legendTitle}>Status Legend:</div>
        <div className={styles.legendItems}>
          <div className={styles.legendItem}>
            <span className={styles.legendIcon}>🔍</span>
            <span className={styles.legendText}>Checking directory...</span>
          </div>
          <div className={styles.legendItem}>
            <span className={styles.legendIcon}>📂</span>
            <span className={styles.legendText}>Directory exists</span>
          </div>
          <div className={styles.legendItem}>
            <span className={styles.legendIcon}>📂❌</span>
            <span className={styles.legendText}>Directory not found</span>
          </div>
          <div className={styles.legendItem}>
            <span className={styles.legendIcon}>✅</span>
            <span className={styles.legendText}>File found</span>
          </div>
          <div className={styles.legendItem}>
            <span className={styles.legendIcon}>❌</span>
            <span className={styles.legendText}>File not found</span>
          </div>
          <div className={styles.legendItem}>
            <span className={styles.legendIcon}>⏭️</span>
            <span className={styles.legendText}>File skipped</span>
          </div>
        </div>
      </div>
    </div>
  );
};