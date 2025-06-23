// src/webparts/xyea/components/RenameFilesManagement/components/DataTableView.tsx

import * as React from 'react';
import styles from '../RenameFilesManagement.module.scss';
import { IRenameFilesData } from '../types/RenameFilesTypes';
import { ColumnResizeHandler } from '../handlers/ColumnResizeHandler';

export interface IDataTableViewProps {
  data: IRenameFilesData;
  fileSearchResults: { [rowIndex: number]: 'found' | 'not-found' | 'searching' };
  columnResizeHandler: ColumnResizeHandler;
  onCellEdit: (columnId: string, rowIndex: number, newValue: string) => void;
}

export const DataTableView: React.FC<IDataTableViewProps> = ({
  data,
  fileSearchResults,
  columnResizeHandler,
  onCellEdit
}) => {
  const handleCellChange = (columnId: string, rowIndex: number, value: string) => {
    onCellEdit(columnId, rowIndex, value);
  };

  const renderSearchIndicator = (rowIndex: number) => {
    const searchResult = fileSearchResults[rowIndex];
    
    if (!searchResult) return null;

    switch (searchResult) {
      case 'searching':
        return <span className={styles.searchingIndicator} title="Searching...">🔍</span>;
      case 'found':
        return <span className={styles.foundIndicator} title="File found in SharePoint">✅</span>;
      case 'not-found':
        return <span className={styles.notFoundIndicator} title="File not found in SharePoint">❌</span>;
      default:
        return null;
    }
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
            const searchResult = fileSearchResults[row.rowIndex];
            
            return (
              <tr key={row.rowIndex} className={row.isEdited ? styles.editedRow : ''}>
                <td className={styles.rowNumber}>
                  <div className={styles.rowNumberContent}>
                    <span className={styles.rowNumberText}>{row.rowIndex + 1}</span>
                    {searchResult && (
                      <div className={styles.searchIndicator}>
                        {renderSearchIndicator(row.rowIndex)}
                      </div>
                    )}
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
    </div>
  );
};