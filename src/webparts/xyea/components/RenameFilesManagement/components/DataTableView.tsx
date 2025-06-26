// src/webparts/xyea/components/RenameFilesManagement/components/DataTableView.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { 
  IRenameFilesData,
  DirectoryStatus,
  FileSearchStatus,
  StatusCode
} from '../types/RenameFilesTypes';
import { ColumnResizeHandler } from '../handlers/ColumnResizeHandler';

export interface IDataTableViewProps {
  data: IRenameFilesData;
  directoryResults: { [rowIndex: number]: DirectoryStatus };
  fileSearchResults: { [rowIndex: number]: FileSearchStatus };
  columnResizeHandler: ColumnResizeHandler;
  onCellEdit: (columnId: string, rowIndex: number, newValue: string) => void;
}

export const DataTableView: React.FC<IDataTableViewProps> = ({
  data,
  directoryResults,
  fileSearchResults,
  columnResizeHandler,
  onCellEdit
}) => {
  const handleCellChange = (columnId: string, rowIndex: number, value: string): void => {
    onCellEdit(columnId, rowIndex, value);
  };

  /**
   * NEW: Get current status code for a row with priority logic
   */
  const getCurrentStatusCode = (rowIndex: number): StatusCode => {
    const directoryStatus = directoryResults[rowIndex];
    const fileSearchStatus = fileSearchResults[rowIndex];
    
    // Priority logic:
    // 1. If directory doesn't exist or has error -> show directory status
    // 2. If directory exists and we have file search results -> show file status  
    // 3. If directory exists but no file search yet -> show directory status
    
    if (directoryStatus === 'not-exists' || directoryStatus === 'error') {
      // Map directory status to status code
      switch (directoryStatus) {
        case 'not-exists':
          return StatusCode.DIRECTORY_NOT_EXISTS;
        case 'error':
          return StatusCode.DIRECTORY_ERROR;
        default:
          return StatusCode.CHECKING_DIRECTORY;
      }
    } else if (fileSearchStatus && directoryStatus === 'exists') {
      // Map file search status to status code
      switch (fileSearchStatus) {
        case 'found':
          return StatusCode.FOUND;
        case 'not-found':
          return StatusCode.NOT_FOUND;
        case 'searching':
          return StatusCode.SEARCHING;
        case 'skipped':
          return StatusCode.SKIPPED;
        default:
          return StatusCode.SEARCHING;
      }
    } else if (directoryStatus) {
      // Map directory status to status code
      switch (directoryStatus) {
        case 'checking':
          return StatusCode.CHECKING_DIRECTORY;
        case 'exists':
          return StatusCode.DIRECTORY_EXISTS;
        default:
          return StatusCode.CHECKING_DIRECTORY;
      }
    }
    
    return StatusCode.CHECKING_DIRECTORY; // Default
  };

  /**
   * NEW: Get CSS class for status styling
   */
  const getStatusCssClass = (statusCode: StatusCode): string => {
    switch (statusCode) {
      case StatusCode.FOUND:
      case StatusCode.RENAMED:
      case StatusCode.DIRECTORY_EXISTS:
        return 'statusSuccess';
      case StatusCode.NOT_FOUND:
      case StatusCode.RENAME_ERROR:
      case StatusCode.DIRECTORY_NOT_EXISTS:
      case StatusCode.DIRECTORY_ERROR:
        return 'statusError';
      case StatusCode.SKIPPED:
      case StatusCode.RENAME_SKIPPED:
        return 'statusWarning';
      case StatusCode.CHECKING_DIRECTORY:
      case StatusCode.SEARCHING:
      case StatusCode.RENAMING:
        return 'statusProgress';
      default:
        return 'statusDefault';
    }
  };

  /**
   * NEW: Get tooltip text for status code
   */
  const getStatusTooltip = (statusCode: StatusCode): string => {
    switch (statusCode) {
      case StatusCode.CHECKING_DIRECTORY:
        return 'Checking directory...';
      case StatusCode.DIRECTORY_EXISTS:
        return 'Directory exists';
      case StatusCode.DIRECTORY_NOT_EXISTS:
        return 'Directory not found';
      case StatusCode.DIRECTORY_ERROR:
        return 'Directory check error';
      case StatusCode.SEARCHING:
        return 'Searching for file...';
      case StatusCode.FOUND:
        return 'File found';
      case StatusCode.NOT_FOUND:
        return 'File not found';
      case StatusCode.SKIPPED:
        return 'File skipped';
      case StatusCode.RENAMING:
        return 'Renaming file...';
      case StatusCode.RENAMED:
        return 'File renamed successfully';
      case StatusCode.RENAME_ERROR:
        return 'File rename error';
      case StatusCode.RENAME_SKIPPED:
        return 'File rename skipped';
      default:
        return 'Unknown status';
    }
  };

  /**
   * Get CSS class for row styling based on status
   */
  const getRowStatusClass = (rowIndex: number): string => {
    const directoryStatus = directoryResults[rowIndex];
    const fileSearchStatus = fileSearchResults[rowIndex];
    
    // Apply row styling based on most relevant status
    if (directoryStatus === 'not-exists') {
      return 'directoryNotExistsRow';
    } else if (directoryStatus === 'error') {
      return 'directoryErrorRow';
    } else if (fileSearchStatus && directoryStatus === 'exists') {
      if (fileSearchStatus === 'found') {
        return 'fileFoundRow';
      } else if (fileSearchStatus === 'not-found') {
        return 'fileNotFoundRow';
      }
    }
    
    return '';
  };

  /**
   * Fixed: Safe CSS class access using computed property names
   */
  const getStatusCellClass = (statusCode: StatusCode): string => {
    const statusClass = getStatusCssClass(statusCode);
    const baseClass = styles.statusCell;
    
    // Use array access instead of computed property to avoid TypeScript error
    const statusStyleClass = styles[statusClass as keyof typeof styles] as string;
    
    return `${baseClass} ${statusStyleClass || ''}`;
  };

  return (
    <div className={styles.tableContainer}>
      <table className={styles.dataTable}>
        <thead>
          <tr>
            <th className={styles.rowHeader}>#</th>
            {/* NEW: Status Column Header */}
            <th className={`${styles.columnHeader} ${styles.statusColumn}`}>
              <div className={styles.headerContent}>
                <span className={styles.columnName}>Status</span>
                <span className={styles.statusBadge}>Live</span>
              </div>
            </th>
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
            const rowStatusClass = getRowStatusClass(row.rowIndex);
            const baseRowClass = row.isEdited ? styles.editedRow : '';
            const finalRowClass = [baseRowClass, rowStatusClass].filter(Boolean).join(' ');
            
            // Get current status for this row
            const currentStatusCode = getCurrentStatusCode(row.rowIndex);
            const statusTooltip = getStatusTooltip(currentStatusCode);
            
            return (
              <tr key={row.rowIndex} className={finalRowClass}>
                {/* Row Number Column */}
                <td className={styles.rowNumber}>
                  <div className={styles.rowNumberContent}>
                    <span className={styles.rowNumberText}>{row.rowIndex + 1}</span>
                  </div>
                </td>
                
                {/* NEW: Status Column */}
                <td className={getStatusCellClass(currentStatusCode)}>
                  <div 
                    className={styles.statusCode}
                    title={statusTooltip}
                  >
                    {currentStatusCode}
                  </div>
                </td>
                
                {/* Data Columns */}
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
      
      {/* NEW: Status Legend */}
      <div className={styles.statusLegend}>
        <div className={styles.legendTitle}>Status Codes:</div>
        <div className={styles.legendItems}>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusProgress}`}>CHK</span>
            <span className={styles.legendText}>Checking directory</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusSuccess}`}>DIR</span>
            <span className={styles.legendText}>Directory exists</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusError}`}>NDF</span>
            <span className={styles.legendText}>Directory not found</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusProgress}`}>SCH</span>
            <span className={styles.legendText}>Searching files</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusSuccess}`}>FND</span>
            <span className={styles.legendText}>File found</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusError}`}>NFD</span>
            <span className={styles.legendText}>File not found</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusProgress}`}>RNG</span>
            <span className={styles.legendText}>Renaming</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusSuccess}`}>REN</span>
            <span className={styles.legendText}>Renamed</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusWarning}`}>SKP</span>
            <span className={styles.legendText}>Skipped</span>
          </div>
          <div className={styles.legendItem}>
            <span className={`${styles.legendCode} ${styles.statusError}`}>ERR</span>
            <span className={styles.legendText}>Error</span>
          </div>
        </div>
      </div>
    </div>
  );
};
