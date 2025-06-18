// src/webparts/xyea/components/EditDialog/EditPropsDialog.tsx - Updated with ConvertType support

import * as React from 'react';
import { IConvertType } from '../../models/IConvertType';

export interface IEditPropsDialogProps {
  isOpen: boolean;
  isEditMode: boolean;
  convertFileId: number;
  item?: { 
    Title?: string; 
    Prop?: string; 
    Prop2?: string;
    ConvertType?: number;
    ConvertType2?: number;
  } | undefined;
  title: string;
  loading?: boolean;
  convertTypes: IConvertType[]; // NEW: List of available convert types
  onSave: (
    convertFileId: number, 
    title: string, 
    prop: string, 
    prop2: string, 
    convertTypeId: number, 
    convertType2Id: number
  ) => Promise<void>;
  onCancel: () => void;
}

interface IEditPropsDialogState {
  title: string;
  prop: string;
  prop2: string;
  convertTypeId: number;
  convertType2Id: number;
  titleError: string;
}

class EditPropsDialog extends React.Component<IEditPropsDialogProps, IEditPropsDialogState> {
  private readonly DEFAULT_CONVERT_TYPE_ID = 1; // Строковый тип по умолчанию
  
  constructor(props: IEditPropsDialogProps) {
    super(props);
    this.state = {
      title: props.item?.Title || '',
      prop: props.item?.Prop || '',
      prop2: props.item?.Prop2 || '',
      convertTypeId: props.item?.ConvertType || this.DEFAULT_CONVERT_TYPE_ID,
      convertType2Id: props.item?.ConvertType2 || this.DEFAULT_CONVERT_TYPE_ID,
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
        convertTypeId: this.props.item?.ConvertType || this.DEFAULT_CONVERT_TYPE_ID,
        convertType2Id: this.props.item?.ConvertType2 || this.DEFAULT_CONVERT_TYPE_ID,
        titleError: ''
      });
    }

    // Update convert type selections when types are loaded
    if (prevProps.convertTypes.length === 0 && this.props.convertTypes.length > 0) {
      // If we don't have item data but types are now loaded, set defaults
      if (!this.props.item?.ConvertType) {
        this.setState({
          convertTypeId: this.DEFAULT_CONVERT_TYPE_ID,
          convertType2Id: this.DEFAULT_CONVERT_TYPE_ID
        });
      }
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

  private handleConvertTypeChange = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({ convertTypeId: parseInt(e.target.value, 10) });
  }

  private handleConvertType2Change = (e: React.ChangeEvent<HTMLSelectElement>): void => {
    this.setState({ convertType2Id: parseInt(e.target.value, 10) });
  }

  private handleSave = async (): Promise<void> => {
    const { title, prop, prop2, convertTypeId, convertType2Id } = this.state;
    
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
        prop2.trim(),
        convertTypeId,
        convertType2Id
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

  private getConvertTypeName = (id: number): string => {
    const convertType = this.props.convertTypes.find(ct => ct.Id === id);
    return convertType ? convertType.Title : `Type ${id}`;
  }

  public render(): React.ReactElement<IEditPropsDialogProps> {
    if (!this.props.isOpen) {
      return <div style={{ display: 'none' }} />;
    }

    const { loading, convertTypes } = this.props;
    const { title, prop, prop2, convertTypeId, convertType2Id, titleError } = this.state;
    const isValid = this.isFormValid();

    // Debug logging
    console.log('[EditPropsDialog] Render with convertTypes:', {
      convertTypesLength: convertTypes.length,
      convertTypes: convertTypes.slice(0, 3), // Log first 3 types
      currentConvertTypeId: convertTypeId,
      currentConvertType2Id: convertType2Id
    });

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
          minWidth: '500px',
          maxWidth: '700px',
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
                Prop and Prop2 fields are optional. Convert types default to "String" (ID=1).
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

            {/* Two-column layout for Prop fields */}
            <div style={{
              display: 'grid',
              gridTemplateColumns: '1fr 1fr',
              gap: '16px',
              marginBottom: '16px'
            }}>
              {/* Prop Field */}
              <div>
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
                  placeholder="Enter prop value"
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
              <div>
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
                  placeholder="Enter prop2 value"
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

            {/* Two-column layout for ConvertType fields */}
            <div style={{
              display: 'grid',
              gridTemplateColumns: '1fr 1fr',
              gap: '16px',
              marginBottom: '8px'
            }}>
              {/* ConvertType Field */}
              <div>
                <label style={{
                  display: 'block',
                  marginBottom: '4px',
                  fontWeight: 600,
                  color: '#323130',
                  fontSize: '14px'
                }}>
                  Convert Type for Prop
                </label>
                <select
                  value={convertTypeId}
                  onChange={this.handleConvertTypeChange}
                  disabled={loading}
                  style={{ 
                    width: '100%', 
                    padding: '8px 12px', 
                    border: '1px solid #edebe9',
                    borderRadius: '2px',
                    fontSize: '14px',
                    boxSizing: 'border-box',
                    outline: 'none',
                    backgroundColor: 'white'
                  }}
                  onFocus={(e) => {
                    e.target.style.borderColor = '#0078d4';
                    e.target.style.boxShadow = '0 0 0 1px #0078d4';
                  }}
                  onBlur={(e) => {
                    e.target.style.borderColor = '#edebe9';
                    e.target.style.boxShadow = 'none';
                  }}
                >
                  {convertTypes.length === 0 ? (
                    <option value={this.DEFAULT_CONVERT_TYPE_ID}>Loading types...</option>
                  ) : (
                    convertTypes.map(type => (
                      <option key={type.Id} value={type.Id}>
                        {type.Title}
                      </option>
                    ))
                  )}
                </select>
              </div>

              {/* ConvertType2 Field */}
              <div>
                <label style={{
                  display: 'block',
                  marginBottom: '4px',
                  fontWeight: 600,
                  color: '#323130',
                  fontSize: '14px'
                }}>
                  Convert Type for Prop2
                </label>
                <select
                  value={convertType2Id}
                  onChange={this.handleConvertType2Change}
                  disabled={loading}
                  style={{ 
                    width: '100%', 
                    padding: '8px 12px', 
                    border: '1px solid #edebe9',
                    borderRadius: '2px',
                    fontSize: '14px',
                    boxSizing: 'border-box',
                    outline: 'none',
                    backgroundColor: 'white'
                  }}
                  onFocus={(e) => {
                    e.target.style.borderColor = '#0078d4';
                    e.target.style.boxShadow = '0 0 0 1px #0078d4';
                  }}
                  onBlur={(e) => {
                    e.target.style.borderColor = '#edebe9';
                    e.target.style.boxShadow = 'none';
                  }}
                >
                  {convertTypes.length === 0 ? (
                    <option value={this.DEFAULT_CONVERT_TYPE_ID}>Loading types...</option>
                  ) : (
                    convertTypes.map(type => (
                      <option key={type.Id} value={type.Id}>
                        {type.Title}
                      </option>
                    ))
                  )}
                </select>
              </div>
            </div>

            {/* Current selections info */}
            <div style={{
              backgroundColor: '#f8f8f8',
              border: '1px solid #edebe9',
              borderRadius: '4px',
              padding: '8px 12px',
              fontSize: '12px',
              color: '#605e5c',
              marginBottom: '16px'
            }}>
              <strong>Selected types:</strong> Prop → {this.getConvertTypeName(convertTypeId)}, 
              Prop2 → {this.getConvertTypeName(convertType2Id)}
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