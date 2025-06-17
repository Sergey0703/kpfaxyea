// src/webparts/xyea/components/ConvertFilesPropsTable/ConvertFilesPropsTable.tsx

import * as React from 'react';
import styles from './ConvertFilesPropsTable.module.scss';
import { IConvertFilesPropsTableProps } from './IConvertFilesPropsTableProps';
import { IConvertFileProps } from '../../models';
import { PriorityHelper } from '../../utils';
import ExcelImportButton, { IExcelImportData } from '../ExcelImportButton/ExcelImportButton';

export interface IConvertFilesPropsTableState {
  error: string | undefined;
}

export default class ConvertFilesPropsTable extends React.Component<IConvertFilesPropsTableProps, IConvertFilesPropsTableState> {
  
  constructor(props: IConvertFilesPropsTableProps) {
    super(props);
    this.state = {
      error: undefined
    };
  }

  private handleEdit = (item: IConvertFileProps): void => {
    this.props.onEdit(item);
  }

  private handleDelete = (id: number): void => {
    if (confirm('Are you sure you want to mark this item as deleted?')) {
      this.props.onToggleDeleted(id, true);
    }
  }

  private handleRestore = (id: number): void => {
    this.props.onToggleDeleted(id, false);
  }

  private handleMoveUp = (id: number): void => {
    this.props.onMoveUp(id);
  }

  private handleMoveDown = (id: number): void => {
    this.props.onMoveDown(id);
  }

  private handleAdd = (): void => {
    this.props.onAdd(this.props.convertFileId);
  }

  private handleExcelImport = async (data: IExcelImportData[]): Promise<void> => {
    try {
      console.log('[ConvertFilesPropsTable] Starting Excel import:', {
        convertFileId: this.props.convertFileId,
        dataCount: data.length
      });

      // Call the new import handler from props
      if (this.props.onImportFromExcel) {
        await this.props.onImportFromExcel(this.props.convertFileId, data);
      } else {
        throw new Error('Excel import is not supported');
      }

      console.log('[ConvertFilesPropsTable] Excel import completed successfully');
    } catch (error) {
      console.error('[ConvertFilesPropsTable] Excel import failed:', error);
      this.setState({ 
        error: error instanceof Error ? error.message : 'Excel import failed' 
      });
      throw error; // Re-throw to let ExcelImportButton handle it
    }
  }

  private canMoveUp = (item: IConvertFileProps): boolean => {
    return PriorityHelper.canMoveUp(this.props.allItems, item.Id, this.props.convertFileId);
  }

  private canMoveDown = (item: IConvertFileProps): boolean => {
    return PriorityHelper.canMoveDown(this.props.allItems, item.Id, this.props.convertFileId);
  }

  private getSortedItems = (): IConvertFileProps[] => {
    return PriorityHelper.sortByPriority(this.props.items);
  }

  private clearError = (): void => {
    this.setState({ error: undefined });
  }

  public render(): React.ReactElement<IConvertFilesPropsTableProps> {
    const { convertFileTitle, loading } = this.props;
    const { error } = this.state;
    const sortedItems = this.getSortedItems();

    return (
      <div className={styles.convertFilesPropsTable}>
        <div className={styles.header}>
          <h3 className={styles.title}>Properties for: {convertFileTitle}</h3>
          <div className={styles.headerActions}>
            <ExcelImportButton
              onImport={this.handleExcelImport}
              disabled={loading}
              existingItemsCount={sortedItems.length}
            />
            <button 
              className={styles.addButton}
              onClick={this.handleAdd}
              disabled={loading}
            >
              + Add Property
            </button>
          </div>
        </div>

        {error && (
          <div className={styles.error}>
            <span className={styles.errorIcon}>⚠️</span>
            <span className={styles.errorMessage}>Error: {error}</span>
            <button 
              className={styles.clearErrorButton}
              onClick={this.clearError}
              title="Clear error"
            >
              ✕
            </button>
          </div>
        )}

        {loading ? (
          <div className={styles.loading}>
            Loading properties...
          </div>
        ) : sortedItems.length === 0 ? (
          <div className={styles.empty}>
            <div className={styles.emptyMessage}>No properties found for this convert file.</div>
            <div className={styles.emptyActions}>
              <ExcelImportButton
                onImport={this.handleExcelImport}
                disabled={loading}
                existingItemsCount={0}
              />
              <button 
                className={styles.addButton}
                onClick={this.handleAdd}
              >
                Add First Property
              </button>
            </div>
          </div>
        ) : (
          <table className={styles.table}>
            <thead className={styles.tableHead}>
              <tr>
                <th className={styles.headerCell}>Priority</th>
                <th className={styles.headerCell}>Title</th>
                <th className={styles.headerCell}>Prop</th>
                <th className={styles.headerCell}>Prop2</th>
                <th className={styles.headerCell}>Status</th>
                <th className={styles.headerCell}>Created</th>
                <th className={styles.headerCell}>Actions</th>
              </tr>
            </thead>
            <tbody className={styles.tableBody}>
              {sortedItems.map((item: IConvertFileProps) => (
                <tr 
                  key={item.Id}
                  className={`${styles.tableRow} ${item.IsDeleted ? styles.deleted : ''}`}
                >
                  <td className={`${styles.tableCell} ${styles.priorityCell}`}>
                    {item.Priority}
                  </td>
                  <td className={`${styles.tableCell} ${styles.titleCell}`}>
                    {item.Title}
                  </td>
                  <td className={`${styles.tableCell} ${styles.propCell}`}>
                    {item.Prop}
                  </td>
                  <td className={`${styles.tableCell} ${styles.propCell}`}>
                    {item.Prop2}
                  </td>
                  <td className={`${styles.tableCell} ${styles.statusCell}`}>
                    <span className={`${styles.statusBadge} ${item.IsDeleted ? styles.deleted : styles.active}`}>
                      {item.IsDeleted ? 'Deleted' : 'Active'}
                    </span>
                  </td>
                  <td className={styles.tableCell}>
                    {item.Created ? new Date(item.Created).toLocaleDateString() : '-'}
                  </td>
                  <td className={`${styles.tableCell} ${styles.actionsCell}`}>
                    {/* Move buttons - now work for deleted items too */}
                    <button 
                      className={`${styles.actionButton} ${styles.moveButton}`}
                      onClick={() => this.handleMoveUp(item.Id)}
                      disabled={loading || !this.canMoveUp(item)}
                      title={item.IsDeleted ? "Move Deleted Item Up" : "Move Up"}
                    >
                      ↑
                    </button>
                    <button 
                      className={`${styles.actionButton} ${styles.moveButton}`}
                      onClick={() => this.handleMoveDown(item.Id)}
                      disabled={loading || !this.canMoveDown(item)}
                      title={item.IsDeleted ? "Move Deleted Item Down" : "Move Down"}
                    >
                      ↓
                    </button>
                    
                    {/* Edit button - only for active items */}
                    <button 
                      className={`${styles.actionButton} ${styles.editButton}`}
                      onClick={() => this.handleEdit(item)}
                      disabled={loading || item.IsDeleted}
                      title={item.IsDeleted ? "Cannot edit deleted item" : "Edit"}
                    >
                      Edit
                    </button>
                    
                    {/* Delete/Restore button */}
                    {item.IsDeleted ? (
                      <button 
                        className={`${styles.actionButton} ${styles.restoreButton}`}
                        onClick={() => this.handleRestore(item.Id)}
                        disabled={loading}
                        title="Restore"
                      >
                        Restore
                      </button>
                    ) : (
                      <button 
                        className={`${styles.actionButton} ${styles.deleteButton}`}
                        onClick={() => this.handleDelete(item.Id)}
                        disabled={loading}
                        title="Mark as Deleted"
                      >
                        Delete
                      </button>
                    )}
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