// src/webparts/xyea/components/EditDialog/EditDialog.tsx

import * as React from 'react';

export interface IEditDialogProps {
  isOpen: boolean;
  isEditMode: boolean;
  item?: { Title?: string } | undefined; // Use specific type instead of any
  title: string;
  loading?: boolean;
  onSave: (title: string) => Promise<void>;
  onCancel: () => void;
}

interface IEditDialogState {
  title: string;
}

class EditDialog extends React.Component<IEditDialogProps, IEditDialogState> {
  
  constructor(props: IEditDialogProps) {
    super(props);
    this.state = {
      title: props.item?.Title || ''
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
          <h3>{this.props.isEditMode ? 'Edit' : 'Create'}</h3>
          <input 
            type="text" 
            value={this.state.title}
            onChange={(e) => this.setState({ title: e.target.value })}
            placeholder="Enter title"
            style={{ width: '100%', padding: '8px', marginBottom: '16px' }}
          />
          <div style={{ display: 'flex', gap: '8px', justifyContent: 'flex-end' }}>
            <button onClick={this.props.onCancel}>Cancel</button>
            <button 
              onClick={() => this.props.onSave(this.state.title).catch(console.error)}
              disabled={!this.state.title.trim()}
            >
              Save
            </button>
          </div>
        </div>
      </div>
    );
  }
}

export default EditDialog;