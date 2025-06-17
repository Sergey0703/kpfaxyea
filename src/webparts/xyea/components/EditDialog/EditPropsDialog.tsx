// src/webparts/xyea/components/EditDialog/EditPropsDialog.tsx

import * as React from 'react';
import styles from './EditDialog.module.scss';
import { ValidationHelper } from '../../utils';
import { IConvertFileProps } from '../../models';

export interface IEditPropsDialogProps {
  isOpen: boolean;
  isEditMode: boolean;
  convertFileId: number;
  item?: IConvertFileProps | undefined; // Changed from null to undefined
  title: string;
  loading?: boolean;
  onSave: (convertFileId: number, title: string, prop: string, prop2: string) => Promise<void>;
  onCancel: () => void;
}

export interface IEditPropsDialogState {
  title: string;
  prop: string;
  prop2: string;
  errors: { [key: string]: string };
  saving: boolean;
}

export default class EditPropsDialog extends React.Component<IEditPropsDialogProps, IEditPropsDialogState> {
  
  constructor(props: IEditPropsDialogProps) {
    super(props);
    
    this.state = {
      title: props.item?.Title || '',
      prop: props.item?.Prop || '',
      prop2: props.item?.Prop2 || '',
      errors: {},
      saving: false
    };
  }

  public componentDidUpdate(prevProps: IEditPropsDialogProps): void {
    // Обновить состояние при изменении пропсов
    if (prevProps.item !== this.props.item || prevProps.isOpen !== this.props.isOpen) {
      this.setState({
        title: this.props.item?.Title || '',
        prop: this.props.item?.Prop || '',
        prop2: this.props.item?.Prop2 || '',
        errors: {},
        saving: false
      });
    }
  }

  private handleInputChange = (field: string): (event: React.ChangeEvent<HTMLInputElement>) => void => { // Added explicit return type
    return (event: React.ChangeEvent<HTMLInputElement>): void => {
      const value = event.target.value;
      this.setState(prevState => ({ 
        ...prevState,
        [field]: value,
        errors: { ...prevState.errors, [field]: '' } // Очистить ошибку при изменении
      }));
    };
  }

  private validateForm = (): boolean => {
    const { title, prop, prop2 } = this.state;
    const validation = ValidationHelper.validateConvertFileProps(
      title, 
      this.props.convertFileId, 
      prop, 
      prop2, 
      1 // Priority проверим отдельно
    );
    
    if (!validation.isValid) {
      const errors: { [key: string]: string } = {};
      validation.errors.forEach(error => {
        if (error.includes('Title')) errors.title = error;
        else if (error.includes('Prop2')) errors.prop2 = error;
        else if (error.includes('Prop')) errors.prop = error;
        else errors.general = error;
      });
      
      this.setState({ errors });
      return false;
    }

    this.setState({ errors: {} });
    return true;
  }

  private handleSave = (): void => {
    if (!this.validateForm()) {
      return;
    }

    const { title, prop, prop2 } = this.state;
    const sanitizedTitle = ValidationHelper.sanitizeString(title);
    const sanitizedProp = ValidationHelper.sanitizeString(prop);
    const sanitizedProp2 = ValidationHelper.sanitizeString(prop2);

    this.setState({ saving: true });
    
    // Handle the promise properly
    this.props.onSave(this.props.convertFileId, sanitizedTitle, sanitizedProp, sanitizedProp2)
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
      prop: this.props.item?.Prop || '',
      prop2: this.props.item?.Prop2 || '',
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
    if (event.target === event.currentTarget) {
      this.handleCancel();
    }
  }

  public render(): React.ReactElement<IEditPropsDialogProps> | undefined { // Changed from null to undefined
    if (!this.props.isOpen) {
      return undefined;
    }

    const { title, prop, prop2, errors, saving } = this.state;
    const { isEditMode, loading } = this.props;
    const isLoading = loading || saving;

    return (
      <div className={styles.dialogOverlay} onClick={this.handleOverlayClick}>
        <div className={styles.dialog} onKeyDown={this.handleKeyDown}>
          <div className={styles.header}>
            <h2 className={styles.title}>
              {isEditMode ? 'Edit Property' : 'Create New Property'}
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
                onChange={this.handleInputChange('title')}
                placeholder="Enter title..."
                disabled={isLoading}
                autoFocus
              />
              {errors.title && (
                <div className={styles.errorMessage}>{errors.title}</div>
              )}
            </div>

            <div className={styles.field}>
              <label className={styles.label}>
                Prop <span className={styles.required}>*</span>
              </label>
              <input
                type="text"
                className={`${styles.input} ${errors.prop ? styles.error : ''}`}
                value={prop}
                onChange={this.handleInputChange('prop')}
                placeholder="Enter prop value..."
                disabled={isLoading}
              />
              {errors.prop && (
                <div className={styles.errorMessage}>{errors.prop}</div>
              )}
            </div>

            <div className={styles.field}>
              <label className={styles.label}>
                Prop2 <span className={styles.required}>*</span>
              </label>
              <input
                type="text"
                className={`${styles.input} ${errors.prop2 ? styles.error : ''}`}
                value={prop2}
                onChange={this.handleInputChange('prop2')}
                placeholder="Enter prop2 value..."
                disabled={isLoading}
              />
              {errors.prop2 && (
                <div className={styles.errorMessage}>{errors.prop2}</div>
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
              disabled={isLoading || !title.trim() || !prop.trim() || !prop2.trim()}
            >
              {isLoading ? 'Saving...' : (isEditMode ? 'Update' : 'Create')}
            </button>
          </div>
        </div>
      </div>
    );
  }
}