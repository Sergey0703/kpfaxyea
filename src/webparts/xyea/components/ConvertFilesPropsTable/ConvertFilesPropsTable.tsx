// src/webparts/xyea/components/ConvertFilesPropsTable/ConvertFilesPropsTable.tsx

import * as React from 'react';
import styles from './ConvertFilesPropsTable.module.scss';
import { IConvertFilesPropsTableProps } from './IConvertFilesPropsTableProps';
import { IConvertFileProps } from '../../models';
import { PriorityHelper } from '../../utils';

export interface IConvertFilesPropsTableState {
  error: string | null;
}

export default class ConvertFilesPropsTable extends React.Component<IConvertFilesPropsTableProps, IConvertFilesPropsTableState> {
  
  constructor(props: IConvertFilesPropsTableProps) {
    super(props);
    this.state = {
      error: null
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

  private canMoveUp = (item: IConvertFileProps): boolean => {
    return PriorityHelper.canMoveUp(this.props.allItems, item.Id, this.props.convertFileId);
  }

  private canMoveDown = (item: IConvertFileProps): boolean => {
    return PriorityHelper.canMoveDown(this.props.allItems, item.Id, this.props.convertFileId);
  }

  private getSortedItems = (): IConvertFileProps[] => {
    return PriorityHelper.sortByPriority(this.props.items);
  }

  public render(): React.ReactElement<IConvertFilesPropsTableProps> {
    const { convertFileTitle, loading } = this.props;
    const { error } = this.state;
    const sortedItems = this.getSortedItems();

    return (
      <div className={styles.convertFilesPropsTable}>
        <div className={styles.header}>
          <h3 className={styles.title}>Properties for: {convertFileTitle}</h3>
          <button 
            className={styles.addButton}
            onClick={this.handleAdd}
            disabled={loading}
          >
            + Add Property
          </button>
        </div>

        {error && (
          <div className={styles.error}>
            Error: {error}
          </div>
        )}

        {loading ? (
          <div className={styles.loading}>
            Loading properties...
          </div>
        ) : sortedItems.length === 0 ? (
          <div className={styles.empty}>
            <div className={styles.emptyMessage}>No properties found for this convert file.</div>
            <button 
              className={styles.addButton}
              onClick={this.handleAdd}
            >
              Add First Property
            </button>
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
                    <button 
                      className={`${styles.actionButton} ${styles.moveButton}`}
                      onClick={() => this.handleMoveUp(item.Id)}
                      disabled={loading || !this.canMoveUp(item)}
                      title="Move Up"
                    >
                      ↑
                    </button>
                    <button 
                      className={`${styles.actionButton} ${styles.moveButton}`}
                      onClick={() => this.handleMoveDown(item.Id)}
                      disabled={loading || !this.canMoveDown(item)}
                      title="Move Down"
                    >
                      ↓
                    </button>
                    <button 
                      className={`${styles.actionButton} ${styles.editButton}`}
                      onClick={() => this.handleEdit(item)}
                      disabled={loading}
                      title="Edit"
                    >
                      Edit
                    </button>
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