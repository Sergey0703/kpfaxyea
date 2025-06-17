// src/webparts/xyea/components/ConfirmationDialog/ConfirmationDialog.tsx

import * as React from 'react';

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

class ConfirmationDialog extends React.Component<IConfirmationDialogProps> {

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

  public render(): JSX.Element {
    if (!this.props.isOpen) {
      return <div style={{ display: 'none' }} />;
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
      <div style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'rgba(0, 0, 0, 0.4)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 1000
      }}>
        <div style={{
          background: 'white',
          borderRadius: '4px',
          boxShadow: '0 4px 23px rgba(0, 0, 0, 0.2)',
          minWidth: '400px',
          maxWidth: '500px',
          overflow: 'hidden'
        }}>
          <div style={{
            padding: '20px 24px 16px',
            borderBottom: '1px solid #edebe9'
          }}>
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '12px'
            }}>
              {showIcon && (
                <span style={{ fontSize: '24px' }}>{this.getIcon()}</span>
              )}
              <h3 style={{
                margin: 0,
                fontSize: '18px',
                fontWeight: 600,
                color: '#323130'
              }}>
                {title}
              </h3>
            </div>
          </div>

          <div style={{
            padding: '20px 24px'
          }}>
            <div style={{
              margin: 0,
              fontSize: '14px',
              lineHeight: 1.5,
              color: '#605e5c',
              whiteSpace: 'pre-line'
            }}>
              {message}
            </div>
          </div>

          <div style={{
            padding: '16px 24px 20px',
            borderTop: '1px solid #edebe9',
            backgroundColor: '#faf9f8',
            display: 'flex',
            justifyContent: 'flex-end',
            gap: '12px'
          }}>
            <button
              style={{
                padding: '8px 16px',
                border: '1px solid #edebe9',
                borderRadius: '2px',
                backgroundColor: 'white',
                color: '#323130',
                cursor: loading ? 'not-allowed' : 'pointer',
                fontSize: '14px',
                minWidth: '80px',
                opacity: loading ? 0.6 : 1
              }}
              onClick={onCancel}
              disabled={loading}
            >
              {loading ? 'Processing...' : cancelText}
            </button>
            <button
              style={{
                padding: '8px 16px',
                border: '1px solid #0078d4',
                borderRadius: '2px',
                backgroundColor: type === 'danger' ? '#d13438' : '#0078d4',
                color: 'white',
                cursor: loading ? 'not-allowed' : 'pointer',
                fontSize: '14px',
                minWidth: '80px',
                opacity: loading ? 0.6 : 1
              }}
              onClick={onConfirm}
              disabled={loading}
            >
              {loading ? 'Processing...' : confirmText}
            </button>
          </div>
        </div>
      </div>
    );
  }
}

export default ConfirmationDialog;