// src/webparts/xyea/components/EditDialog/EditPropsDialog.tsx

import * as React from 'react';

export interface IEditPropsDialogProps {
  isOpen: boolean;
  isEditMode: boolean;
  convertFileId: number;
  item?: { Title?: string; Prop?: string; Prop2?: string } | undefined; // Use specific type instead of any
  title: string;
  loading?: boolean;
  onSave: (convertFileId: number, title: string, prop: string, prop2: string) => Promise<void>;
  onCancel: () => void;
}

interface IEditPropsDialogState {
  title: string;
  prop: string;
  prop2: string;
}

class EditPropsDialog extends React.Component<IEditPropsDialogProps, IEditPropsDialogState> {
  
  constructor(props: IEditPropsDialogProps) {
    super(props);
    this.state = {
      title: props.item?.Title || '',
      prop: props.item?.Prop || '',
      prop2: props.item?.Prop2 || ''
    };
  }

  public render(): JSX.Element {
    if (!this.props.isOpen) {
      return <div style={{ display: 'none' }} />;
    }

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
          padding: '20px', 
          borderRadius: '4px',
          minWidth: '300px'
        }}>
          <h3>{this.props.isEditMode ? 'Edit Property' : 'Create Property'}</h3>
          <input 
            type="text" 
            value={this.state.title}
            onChange={(e) => this.setState({ title: e.target.value })}
            placeholder="Title"
            style={{ width: '100%', padding: '8px', marginBottom: '8px' }}
          />
          <input 
            type="text" 
            value={this.state.prop}
            onChange={(e) => this.setState({ prop: e.target.value })}
            placeholder="Prop"
            style={{ width: '100%', padding: '8px', marginBottom: '8px' }}
          />
          <input 
            type="text" 
            value={this.state.prop2}
            onChange={(e) => this.setState({ prop2: e.target.value })}
            placeholder="Prop2"
            style={{ width: '100%', padding: '8px', marginBottom: '16px' }}
          />
          <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end' }}>
            <button onClick={this.props.onCancel}>Cancel</button>
            <button 
              onClick={() => this.props.onSave(
                this.props.convertFileId, 
                this.state.title, 
                this.state.prop, 
                this.state.prop2
              ).catch(console.error)}
              disabled={!this.state.title.trim() || !this.state.prop.trim() || !this.state.prop2.trim()}
            >
              Save
            </button>
          </div>
        </div>
      </div>
    );
  }
}

export default EditPropsDialog;