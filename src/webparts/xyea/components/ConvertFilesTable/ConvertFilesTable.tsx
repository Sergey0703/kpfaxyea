// src/webparts/xyea/components/ConvertFilesTable/ConvertFilesTable.tsx

import * as React from 'react';
import styles from './ConvertFilesTable.module.scss';
import { IConvertFilesTableProps } from './IConvertFilesTableProps';
import { IConvertFile } from '../../models';

export interface IConvertFilesTableState {
  error: string | null;
}

export default class ConvertFilesTable extends React.Component<IConvertFilesTableProps, IConvertFilesTableState> {
  
  constructor(props: IConvertFilesTableProps) {
    super(props);
    this.state = {
      error: null
    };
  }

  private handleRowClick = (event: React.MouseEvent, convertFileId: number): void => {
    // Проверяем, что клик не был по кнопке действия
    const target = event.target as HTMLElement;
    if (target.closest('button')) {
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

  private isRowExpanded = (convertFileId: number): boolean => {
    return this.props.expandedRows.includes(convertFileId);
  }

  public render(): React.ReactElement<IConvertFilesTableProps> {
    const { convertFiles, loading, onAdd } = this.props;
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
                <th className={styles.headerCell}></th>
                <th className={styles.headerCell}>ID</th>
                <th className={styles.headerCell}>Title</th>
                <th className={styles.headerCell}>Created</th>
                <th className={styles.headerCell}>Modified</th>
                <th className={styles.headerCell}>Actions</th>
              </tr>
            </thead>
            <tbody className={styles.tableBody}>
              {convertFiles.map((item: IConvertFile) => (
                <tr 
                  key={item.Id}
                  className={`${styles.tableRow} ${this.isRowExpanded(item.Id) ? styles.expanded : ''}`}
                  onClick={(e) => this.handleRowClick(e, item.Id)}
                >
                  <td className={`${styles.tableCell} ${styles.expandCell}`}>
                    <span className={`${styles.expandIcon} ${this.isRowExpanded(item.Id) ? styles.expanded : ''}`}>
                      ▶
                    </span>
                    {this.isRowExpanded(item.Id) && (
                      <span style={{ fontSize: '10px', color: 'green', marginLeft: '4px' }}>
                        (opened)
                      </span>
                    )}
                  </td>
                  <td className={styles.tableCell}>{item.Id}</td>
                  <td className={`${styles.tableCell} ${styles.titleCell}`}>{item.Title}</td>
                  <td className={styles.tableCell}>
                    {item.Created ? new Date(item.Created).toLocaleDateString() : '-'}
                  </td>
                  <td className={styles.tableCell}>
                    {item.Modified ? new Date(item.Modified).toLocaleDateString() : '-'}
                  </td>
                  <td className={`${styles.tableCell} ${styles.actionsCell}`}>
                    <button 
                      className={`${styles.actionButton} ${styles.editButton}`}
                      onClick={(e) => this.handleEdit(e, item)}
                      title="Edit"
                    >
                      ✏️ Edit
                    </button>
                    <button 
                      className={`${styles.actionButton} ${styles.deleteButton}`}
                      onClick={(e) => this.handleDelete(e, item.Id)}
                      title="Delete"
                    >
                      🗑️ Delete
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    );
  }
}