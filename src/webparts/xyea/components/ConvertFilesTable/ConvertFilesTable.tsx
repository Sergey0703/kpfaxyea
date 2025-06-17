// src/webparts/xyea/components/ConvertFilesTable/ConvertFilesTable.tsx

import * as React from 'react';
import styles from './ConvertFilesTable.module.scss';
import { IConvertFilesTableProps } from './IConvertFilesTableProps';
import { IConvertFile } from '../../models';

export interface IConvertFilesTableState {
  error: string | undefined;
}

export default class ConvertFilesTable extends React.Component<IConvertFilesTableProps, IConvertFilesTableState> {
  
  constructor(props: IConvertFilesTableProps) {
    super(props);
    this.state = {
      error: undefined
    };
  }

  private handleRowClick = (event: React.MouseEvent, convertFileId: number): void => {
    // Проверяем, что клик не был по кнопке действия
    const target = event.target as HTMLElement;
    if (target.closest('button') || target.closest('input')) {
      return;
    }
    
    this.props.onRowClick(convertFileId);
  }

  private handleEdit = (event: React.MouseEvent, item: IConvertFile): void => {
    event.stopPropagation();
    this.props.onEdit(item);
  }

  private handleDelete = (event: React.MouseEvent, id: number): void => {
    event.stopPropagation();
    if (confirm('Are you sure you want to delete this item?')) {
      this.props.onDelete(id);
    }
  }

  private handleExportFile = (event: React.MouseEvent, item: IConvertFile): void => {
    event.stopPropagation();
    console.log('[ConvertFilesTable] Export file for:', item.Title);
    
    // Create a hidden file input for export
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.style.display = 'none';
    
    input.onchange = (e: Event) => {
      const target = e.target as HTMLInputElement;
      if (target.files && target.files.length > 0) {
        const file = target.files[0];
        console.log('[ConvertFilesTable] Export file selected:', file.name);
        
        // Update state with selected export file
        const currentFiles = this.props.selectedFiles || {};
        const updatedFiles = {
          ...currentFiles,
          [item.Id]: {
            ...currentFiles[item.Id],
            export: file
          }
        };
        
        if (this.props.onSelectedFilesChange) {
          this.props.onSelectedFilesChange(updatedFiles);
        }
      }
      document.body.removeChild(input);
    };
    
    document.body.appendChild(input);
    input.click();
  }

  private handleImportFile = (event: React.MouseEvent, item: IConvertFile): void => {
    event.stopPropagation();
    console.log('[ConvertFilesTable] Import file for:', item.Title);
    
    // Create a hidden file input for import
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.style.display = 'none';
    
    input.onchange = (e: Event) => {
      const target = e.target as HTMLInputElement;
      if (target.files && target.files.length > 0) {
        const file = target.files[0];
        console.log('[ConvertFilesTable] Import file selected:', file.name);
        
        // Update state with selected import file
        const currentFiles = this.props.selectedFiles || {};
        const updatedFiles = {
          ...currentFiles,
          [item.Id]: {
            ...currentFiles[item.Id],
            import: file
          }
        };
        
        if (this.props.onSelectedFilesChange) {
          this.props.onSelectedFilesChange(updatedFiles);
        }
      }
      document.body.removeChild(input);
    };
    
    document.body.appendChild(input);
    input.click();
  }

  private clearExportFile = (event: React.MouseEvent, item: IConvertFile): void => {
    event.stopPropagation();
    const currentFiles = this.props.selectedFiles || {};
    const updatedFiles = {
      ...currentFiles,
      [item.Id]: {
        ...currentFiles[item.Id],
        export: undefined
      }
    };
    
    if (this.props.onSelectedFilesChange) {
      this.props.onSelectedFilesChange(updatedFiles);
    }
  }

  private clearImportFile = (event: React.MouseEvent, item: IConvertFile): void => {
    event.stopPropagation();
    const currentFiles = this.props.selectedFiles || {};
    const updatedFiles = {
      ...currentFiles,
      [item.Id]: {
        ...currentFiles[item.Id],
        import: undefined
      }
    };
    
    if (this.props.onSelectedFilesChange) {
      this.props.onSelectedFilesChange(updatedFiles);
    }
  }

  private truncateFileName = (fileName: string, maxLength: number = 12): string => {
    if (fileName.length <= maxLength) {
      return fileName;
    }
    const extension = fileName.split('.').pop();
    const nameWithoutExt = fileName.substring(0, fileName.lastIndexOf('.'));
    const truncated = nameWithoutExt.substring(0, maxLength - 4 - (extension?.length || 0));
    return `${truncated}...${extension ? '.' + extension : ''}`;
  }

  private isRowExpanded = (convertFileId: number): boolean => {
    return this.props.expandedRows.includes(convertFileId);
  }

  public render(): React.ReactElement<IConvertFilesTableProps> {
    const { convertFiles, loading, onAdd, selectedFiles } = this.props;
    const { error } = this.state;

    return (
      <div className={styles.convertFilesTable}>
        <div className={styles.tableHeader}>
          <h2 className={styles.title}>Convert Files</h2>
          <button 
            className={styles.addButton}
            onClick={onAdd}
            disabled={loading}
          >
            + Add New
          </button>
        </div>

        {error && (
          <div className={styles.error}>
            Error: {error}
          </div>
        )}

        {loading ? (
          <div className={styles.loading}>
            Loading convert files...
          </div>
        ) : convertFiles.length === 0 ? (
          <div className={styles.empty}>
            <div className={styles.emptyMessage}>No convert files found.</div>
            <button 
              className={styles.addButton}
              onClick={onAdd}
            >
              Create First Convert File
            </button>
          </div>
        ) : (
          <table className={styles.table}>
            <thead className={styles.tableHead}>
              <tr>
                <th className={styles.headerCell} /> {/* Self-closing empty th */}
                <th className={styles.headerCell}>ID</th>
                <th className={styles.headerCell}>Title</th>
                <th className={styles.headerCell}>Actions</th>
                <th className={styles.headerCell}>Export file</th>
                <th className={styles.headerCell}>Import file</th>
              </tr>
            </thead>
            <tbody className={styles.tableBody}>
              {convertFiles.map((item: IConvertFile) => {
                const itemFiles = selectedFiles?.[item.Id] || {};
                const exportFile = itemFiles.export;
                const importFile = itemFiles.import;
                
                return (
                  <tr 
                    key={item.Id}
                    className={`${styles.tableRow} ${this.isRowExpanded(item.Id) ? styles.expanded : ''}`}
                    onClick={(e) => this.handleRowClick(e, item.Id)}
                  >
                    <td className={`${styles.tableCell} ${styles.expandCell}`}>
                      <span className={`${styles.expandIcon} ${this.isRowExpanded(item.Id) ? styles.expanded : ''}`}>
                        ▶
                      </span>
                    </td>
                    <td className={styles.tableCell}>{item.Id}</td>
                    <td className={`${styles.tableCell} ${styles.titleCell}`}>{item.Title}</td>
                    <td className={`${styles.tableCell} ${styles.actionsCell}`}>
                      <button 
                        className={`${styles.actionButton} ${styles.editButton}`}
                        onClick={(e) => this.handleEdit(e, item)}
                        title="Edit"
                      >
                        Edit
                      </button>
                      <button 
                        className={`${styles.actionButton} ${styles.deleteButton}`}
                        onClick={(e) => this.handleDelete(e, item.Id)}
                        title="Delete"
                      >
                        Delete
                      </button>
                    </td>
                    <td className={`${styles.tableCell} ${styles.fileActionsCell}`}>
                      <button 
                        className={`${styles.fileButton} ${styles.exportButton} ${exportFile ? styles.hasFile : ''}`}
                        onClick={(e) => this.handleExportFile(e, item)}
                        title={exportFile ? `Selected: ${exportFile.name}` : "Select export file"}
                        disabled={loading}
                      >
                        <span className={styles.buttonContent}>
                          📤 {exportFile ? this.truncateFileName(exportFile.name) : 'Export file'}
                          {exportFile && (
                            <button 
                              className={styles.clearButton}
                              onClick={(e) => this.clearExportFile(e, item)}
                              title="Clear selection"
                            >
                              ✕
                            </button>
                          )}
                        </span>
                      </button>
                    </td>
                    <td className={`${styles.tableCell} ${styles.fileActionsCell}`}>
                      <button 
                        className={`${styles.fileButton} ${styles.importButton} ${importFile ? styles.hasFile : ''}`}
                        onClick={(e) => this.handleImportFile(e, item)}
                        title={importFile ? `Selected: ${importFile.name}` : "Select import file"}
                        disabled={loading}
                      >
                        <span className={styles.buttonContent}>
                          📥 {importFile ? this.truncateFileName(importFile.name) : 'Import file'}
                          {importFile && (
                            <button 
                              className={styles.clearButton}
                              onClick={(e) => this.clearImportFile(e, item)}
                              title="Clear selection"
                            >
                              ✕
                            </button>
                          )}
                        </span>
                      </button>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        )}
      </div>
    );
  }
}