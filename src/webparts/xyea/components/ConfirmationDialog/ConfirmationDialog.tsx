// src/webparts/xyea/components/ConfirmationDialog/ConfirmationDialog.tsx

import * as React from 'react';
import styles from './ConfirmationDialog.module.scss';

export interface IConfirmationDialogProps {
  isOpen: boolean;
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  onConfirm: () => void;
  onCancel: () => void;
  type?: 'warning' | 'danger' | 'info' | 'success';
  loading?: boolean;
  showIcon?: boolean;
}

export default class ConfirmationDialog extends React.Component<IConfirmationDialogProps> {

  private handleOverlayClick = (event: React.MouseEvent): void => {
    if (event.target === event.currentTarget && !this.props.loading) {
      this.props.onCancel();
    }
  }

  private handleKeyDown = (event: React.KeyboardEvent): void => {
    if (this.props.loading) return;

    if (event.key === 'Enter') {
      event.preventDefault();
      this.props.onConfirm();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      this.props.onCancel();
    }
  }

  private getIcon = (): string => {
    const { type = 'warning' } = this.props;
    switch (type) {
      case 'danger':
        return '⚠️';
      case 'info':
        return 'ℹ️';
      case 'success':
        return '✅';
      case 'warning':
      default:
        return '❓';
    }
  }

  private formatMessage = (message: string): React.ReactNode => {
    // Разбиваем сообщение на параграфы по \n\n
    const paragraphs = message.split('\n\n');
    
    return paragraphs.map((paragraph, index) => {
      // Разбиваем параграф на строки по \n
      const lines = paragraph.split('\n');
      return (
        <p key={index} className={styles.paragraph}>
          {lines.map((line, lineIndex) => (
            <React.Fragment key={lineIndex}>
              {line}
              {lineIndex < lines.length - 1 && <br />}
            </React.Fragment>
          ))}
        </p>
      );
    });
  }

  public render(): React.ReactElement<IConfirmationDialogProps> | null {
    if (!this.props.isOpen) {
      return null;
    }

    const { 
      title, 
      message, 
      confirmText = 'Confirm', 
      cancelText = 'Cancel', 
      onConfirm, 
      onCancel,
      type = 'warning',
      loading = false,
      showIcon = true
    } = this.props;

    return (
      <div className={styles.dialogOverlay} onClick={this.handleOverlayClick}>
        <div className={styles.dialog} onKeyDown={this.handleKeyDown}>
          <div className={styles.header}>
            <div className={styles.titleContainer}>
              {showIcon && (
                <span className={styles.icon}>{this.getIcon()}</span>
              )}
              <h3 className={styles.title}>{title}</h3>
            </div>
          </div>

          <div className={styles.content}>
            <div className={styles.message}>
              {this.formatMessage(message)}
            </div>
          </div>

          <div className={styles.footer}>
            <button
              className={`${styles.button} ${styles.secondary}`}
              onClick={onCancel}
              disabled={loading}
              autoFocus={type !== 'danger'} // Для опасных действий фокус НЕ на Cancel
            >
              {loading ? 'Processing...' : cancelText}
            </button>
            <button
              className={`${styles.button} ${styles.primary} ${styles[type]}`}
              onClick={onConfirm}
              disabled={loading}
              autoFocus={type === 'danger'} // Для опасных действий фокус на Confirm
            >
              {loading ? (
                <>
                  <span className={styles.spinner}></span>
                  Processing...
                </>
              ) : confirmText}
            </button>
          </div>
        </div>
      </div>
    );
  }
}