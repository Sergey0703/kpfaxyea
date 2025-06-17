// src/webparts/xyea/components/EditDialog/EditPropsDialog.tsx

import * as React from 'react';

export interface IEditPropsDialogProps {
  isOpen: boolean;
  isEditMode: boolean;
  convertFileId: number;
  item?: { Title?: string; Prop?: string; Prop2?: string } | undefined;
  title: string;
  loading?: boolean;
  onSave: (convertFileId: number, title: string, prop: string, prop2: string) => Promise<void>;
  onCancel: () => void;
}

interface IEditPropsDialogState {
  title: string;
  prop: string;
  prop2: string;
  titleError: string;
}

class EditPropsDialog extends React.Component<IEditPropsDialogProps, IEditPropsDialogState> {
  
  constructor(props: IEditPropsDialogProps) {
    super(props);
    this.state = {
      title: props.item?.Title || '',
      prop: props.item?.Prop || '',
      prop2: props.item?.Prop2 || '',
      titleError: ''
    };
  }

  public componentDidUpdate(prevProps: IEditPropsDialogProps): void {
    // Reset form when dialog opens with new data
    if (this.props.isOpen && !prevProps.isOpen) {
      this.setState({
        title: this.props.item?.Title || '',
        prop: this.props.item?.Prop || '',
        prop2: this.props.item?.Prop2 || '',
        titleError: ''
      });
    }
  }

  private validateTitle = (value: string): string => {
    const trimmedValue = value.trim();
    if (!trimmedValue) {
      return 'Title is required and cannot be empty';
    }
    if (trimmedValue.length > 255) {
      return 'Title cannot exceed 255 characters';
    }
    return '';
  }

  private handleTitleChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const value = e.target.value;
    const error = this.validateTitle(value);
    
    this.setState({ 
      title: value,
      titleError: error
    });
  }

  private handlePropChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ prop: e.target.value });
  }

  private handleProp2Change = (e: React.ChangeEvent<HTMLInputElement>): void => {
    this.setState({ prop2: e.target.value });
  }

  private handleSave = async (): Promise<void> => {
    const { title, prop, prop2 } = this.state;
    
    // Validate title before saving
    const titleError = this.validateTitle(title);
    if (titleError) {
      this.setState({ titleError });
      return;
    }

    try {
      await this.props.onSave(
        this.props.convertFileId, 
        title.trim(), 
        prop.trim(), 
        prop2.trim()
      );
    } catch (error) {
      console.error('Error saving property:', error);
      // You could add error state here if needed
    }
  }

  private isFormValid = (): boolean => {
    const { title } = this.state;
    return title.trim().length > 0 && title.trim().length <= 255;
  }

  public render(): JSX.Element {
    if (!this.props.isOpen) {
      return <div style={{ display: 'none' }} />;
    }

    const { loading } = this.props;
    const { title, prop, prop2, titleError } = this.state;
    const isValid = this.isFormValid();

    return (
      <div style={{ 
        position: 'fixed', 
        top: 0, 
        left: 0, 
        right: 0, 
        bottom: 0, 
        backgroundColor: 'rgba(0,0,0,0.5)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 1000
      }}>
        <div style={{ 
          backgroundColor: 'white', 
          borderRadius: '4px',
          minWidth: '400px',
          maxWidth: '600px',
          boxShadow: '0 4px 23px rgba(0, 0, 0, 0.2)',
          overflow: 'hidden'
        }}>
          {/* Header */}
          <div style={{
            padding: '20px 24px 16px',
            borderBottom: '1px solid #edebe9',
            backgroundColor: '#f8f8f8'
          }}>
            <h3 style={{
              margin: 0,
              fontSize: '18px',
              fontWeight: 600,
              color: '#323130'
            }}>
              {this.props.isEditMode ? 'Edit Property' : 'Create Property'}
            </h3>
          </div>

          {/* Content */}
          <div style={{
            padding: '24px'
          }}>
            {/* Information Message */}
            <div style={{
              backgroundColor: '#deecf9',
              border: '1px solid #c7e0f4',
              borderRadius: '4px',
              padding: '12px 16px',
              marginBottom: '20px',
              display: 'flex',
              alignItems: 'center',
              gap: '8px'
            }}>
              <span style={{ fontSize: '16px' }}>ℹ️</span>
              <div style={{ fontSize: '14px', color: '#005a9e' }}>
                <strong>Required:</strong> Only the Title field is mandatory. 
                Prop and Prop2 fields are optional and can be left empty.
              </div>
            </div>

            {/* Title Field */}
            <div style={{ marginBottom: '16px' }}>
              <label style={{
                display: 'block',
                marginBottom: '4px',
                fontWeight: 600,
                color: '#323130',
                fontSize: '14px'
              }}>
                Title <span style={{ color: '#d13438' }}>*</span>
              </label>
              <input 
                type="text" 
                value={title}
                onChange={this.handleTitleChange}
                placeholder="Enter title (required)"
                disabled={loading}
                style={{ 
                  width: '100%', 
                  padding: '8px 12px', 
                  border: titleError ? '1px solid #d13438' : '1px solid #edebe9',
                  borderRadius: '2px',
                  fontSize: '14px',
                  boxSizing: 'border-box',
                  outline: 'none'
                }}
                onFocus={(e) => {
                  if (!titleError) {
                    e.target.style.borderColor = '#0078d4';
                    e.target.style.boxShadow = '0 0 0 1px #0078d4';
                  }
                }}
                onBlur={(e) => {
                  if (!titleError) {
                    e.target.style.borderColor = '#edebe9';
                    e.target.style.boxShadow = 'none';
                  }
                }}
              />
              {titleError && (
                <div style={{
                  marginTop: '4px',
                  color: '#d13438',
                  fontSize: '12px'
                }}>
                  {titleError}
                </div>
              )}
            </div>

            {/* Prop Field */}
            <div style={{ marginBottom: '16px' }}>
              <label style={{
                display: 'block',
                marginBottom: '4px',
                fontWeight: 600,
                color: '#323130',
                fontSize: '14px'
              }}>
                Prop <span style={{ color: '#a19f9d', fontWeight: 400 }}>(optional)</span>
              </label>
              <input 
                type="text" 
                value={prop}
                onChange={this.handlePropChange}
                placeholder="Enter prop value (optional)"
                disabled={loading}
                style={{ 
                  width: '100%', 
                  padding: '8px 12px', 
                  border: '1px solid #edebe9',
                  borderRadius: '2px',
                  fontSize: '14px',
                  boxSizing: 'border-box',
                  outline: 'none'
                }}
                onFocus={(e) => {
                  e.target.style.borderColor = '#0078d4';
                  e.target.style.boxShadow = '0 0 0 1px #0078d4';
                }}
                onBlur={(e) => {
                  e.target.style.borderColor = '#edebe9';
                  e.target.style.boxShadow = 'none';
                }}
              />
            </div>

            {/* Prop2 Field */}
            <div style={{ marginBottom: '8px' }}>
              <label style={{
                display: 'block',
                marginBottom: '4px',
                fontWeight: 600,
                color: '#323130',
                fontSize: '14px'
              }}>
                Prop2 <span style={{ color: '#a19f9d', fontWeight: 400 }}>(optional)</span>
              </label>
              <input 
                type="text" 
                value={prop2}
                onChange={this.handleProp2Change}
                placeholder="Enter prop2 value (optional)"
                disabled={loading}
                style={{ 
                  width: '100%', 
                  padding: '8px 12px', 
                  border: '1px solid #edebe9',
                  borderRadius: '2px',
                  fontSize: '14px',
                  boxSizing: 'border-box',
                  outline: 'none'
                }}
                onFocus={(e) => {
                  e.target.style.borderColor = '#0078d4';
                  e.target.style.boxShadow = '0 0 0 1px #0078d4';
                }}
                onBlur={(e) => {
                  e.target.style.borderColor = '#edebe9';
                  e.target.style.boxShadow = 'none';
                }}
              />
            </div>
          </div>

          {/* Footer */}
          <div style={{
            padding: '16px 24px 20px',
            borderTop: '1px solid #edebe9',
            backgroundColor: '#f8f8f8',
            display: 'flex',
            justifyContent: 'flex-end',
            gap: '12px'
          }}>
            <button
              onClick={this.props.onCancel}
              disabled={loading}
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
            >
              {loading ? 'Processing...' : 'Cancel'}
            </button>
            <button
              onClick={() => { this.handleSave().catch(console.error); }}
              disabled={loading || !isValid}
              style={{
                padding: '8px 16px',
                border: '1px solid #0078d4',
                borderRadius: '2px',
                backgroundColor: isValid && !loading ? '#0078d4' : '#c8c6c4',
                color: 'white',
                cursor: (loading || !isValid) ? 'not-allowed' : 'pointer',
                fontSize: '14px',
                minWidth: '80px',
                opacity: loading ? 0.6 : 1
              }}
            >
              {loading ? 'Saving...' : 'Save'}
            </button>
          </div>
        </div>
      </div>
    );
  }
}

export default EditPropsDialog;