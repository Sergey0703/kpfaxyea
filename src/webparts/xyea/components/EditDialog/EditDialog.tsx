// src/webparts/xyea/components/EditDialog/EditDialog.tsx

import * as React from 'react';
import styles from './EditDialog.module.scss';
import { IEditDialogProps } from './IEditDialogProps';
import { ValidationHelper } from '../../utils';

export interface IEditDialogState {
  title: string;
  errors: { [key: string]: string };
  saving: boolean;
}

export default class EditDialog extends React.Component<IEditDialogProps, IEditDialogState> {
  
  constructor(props: IEditDialogProps) {
    super(props);
    
    this.state = {
      title: props.item?.Title || '',
      errors: {},
      saving: false
    };
  }

  public componentDidUpdate(prevProps: IEditDialogProps): void {
    // Обновить состояние при изменении пропсов
    if (prevProps.item !== this.props.item || prevProps.isOpen !== this.props.isOpen) {
      this.setState({
        title: this.props.item?.Title || '',
        errors: {},
        saving: false // Сбросить внутреннее состояние saving
      });
    }
  }

  private handleTitleChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const title = event.target.value;
    this.setState({ 
      title,
      errors: { ...this.state.errors, title: '' } // Очистить ошибку при изменении
    });
  }

  private validateForm = (): boolean => {
    const { title } = this.state;
    const validation = ValidationHelper.validateConvertFile(title);
    
    if (!validation.isValid) {
      this.setState({
        errors: {
          title: validation.errors[0] || 'Invalid title'
        }
      });
      return false;
    }

    this.setState({ errors: {} });
    return true;
  }

  private handleSave = (): void => {
    if (!this.validateForm()) {
      return;
    }

    const { title } = this.state;
    const sanitizedTitle = ValidationHelper.sanitizeString(title);

    this.setState({ saving: true });
    
    // Handle the promise properly
    this.props.onSave(sanitizedTitle)
      .then(() => {
        // Dialog will close through onSave callback
      })
      .catch((error) => {
        console.error('Error saving:', error);
        this.setState({
          errors: {
            general: error instanceof Error ? error.message : 'Failed to save item'
          },
          saving: false
        });
      });
  }

  private handleCancel = (): void => {
    this.setState({
      title: this.props.item?.Title || '',
      errors: {},
      saving: false
    });
    this.props.onCancel();
  }

  private handleKeyDown = (event: React.KeyboardEvent): void => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      this.handleSave();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.handleCancel();
    }
  }

  private handleOverlayClick = (event: React.MouseEvent): void => {
    // Закрыть диалог при клике на overlay
    if (event.target === event.currentTarget) {
      this.handleCancel();
    }
  }

  public render(): React.ReactElement<IEditDialogProps> | undefined { // Changed from null to undefined
    if (!this.props.isOpen) {
      return undefined;
    }

    const { title, errors, saving } = this.state;
    const { isEditMode, loading } = this.props;
    const isLoading = loading || saving;

    return (
      <div className={styles.dialogOverlay} onClick={this.handleOverlayClick}>
        <div className={styles.dialog} onKeyDown={this.handleKeyDown}>
          <div className={styles.header}>
            <h2 className={styles.title}>
              {isEditMode ? 'Edit Convert File' : 'Create New Convert File'}
            </h2>
          </div>

          <div className={styles.content}>
            {errors.general && (
              <div className={styles.field}>
                <div className={styles.errorMessage}>{errors.general}</div>
              </div>
            )}

            <div className={styles.field}>
              <label className={styles.label}>
                Title <span className={styles.required}>*</span>
              </label>
              <input
                type="text"
                className={`${styles.input} ${errors.title ? styles.error : ''}`}
                value={title}
                onChange={this.handleTitleChange}
                placeholder="Enter title..."
                disabled={isLoading}
                autoFocus
              />
              {errors.title && (
                <div className={styles.errorMessage}>{errors.title}</div>
              )}
            </div>
          </div>

          <div className={styles.footer}>
            <button
              className={`${styles.button} ${styles.secondary}`}
              onClick={this.handleCancel}
              disabled={isLoading}
            >
              Cancel
            </button>
            <button
              className={`${styles.button} ${styles.primary}`}
              onClick={this.handleSave}
              disabled={isLoading || !title.trim()}
            >
              {isLoading ? 'Saving...' : (isEditMode ? 'Update' : 'Create')}
            </button>
          </div>
        </div>
      </div>
    );
  }
}