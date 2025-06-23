// src/webparts/xyea/components/Xyea.tsx - Updated with ConvertType support and Rename Files tab

import * as React from 'react';
import styles from './Xyea.module.scss';
import { IXyeaProps } from './IXyeaProps';
import { IConvertFile, IConvertFileProps } from '../models';
import { IConvertType } from '../models/IConvertType';
import { ConvertFilesService, ConvertFilesPropsService, ConvertTypesService } from '../services';
import ConvertFilesTable from './ConvertFilesTable/ConvertFilesTable';
import ConvertFilesPropsTable from './ConvertFilesPropsTable/ConvertFilesPropsTable';
import EditDialog from './EditDialog/EditDialog';
import EditPropsDialog from './EditDialog/EditPropsDialog';
import Tabs, { ITabItem } from './Tabs/Tabs';
import SeparateFilesManagement from './SeparateFilesManagement/SeparateFilesManagement';
import RenameFilesManagement from './RenameFilesManagement/RenameFilesManagement';
import { IExcelImportData } from './ExcelImportButton/ExcelImportButton';
import { ISelectedFiles } from './ConvertFilesTable/IConvertFilesTableProps';

export interface IXyeaState {
  convertFiles: IConvertFile[];
  convertFilesProps: IConvertFileProps[];
  convertTypes: IConvertType[]; // NEW: Convert types
  loading: boolean;
  error: string | undefined;
  expandedRows: number[];
  // Состояние диалога для ConvertFiles
  dialogOpen: boolean;
  dialogEditMode: boolean;
  dialogItem: IConvertFile | undefined;
  dialogLoading: boolean;
  // Состояние диалога для ConvertFilesProps
  propsDialogOpen: boolean;
  propsDialogEditMode: boolean;
  propsDialogItem: IConvertFileProps | undefined;
  propsDialogConvertFileId: number;
  propsDialogLoading: boolean;
  // Selected files state
  selectedFiles: ISelectedFiles;
}

export default class Xyea extends React.Component<IXyeaProps, IXyeaState> {
  private convertFilesService: ConvertFilesService;
  private convertFilesPropsService: ConvertFilesPropsService;
  private convertTypesService: ConvertTypesService; // NEW: Convert types service

  constructor(props: IXyeaProps) {
    super(props);
    
    this.state = {
      convertFiles: [],
      convertFilesProps: [],
      convertTypes: [], // NEW: Initialize convert types
      loading: true,
      error: undefined,
      expandedRows: [],
      // Диалог для ConvertFiles
      dialogOpen: false,
      dialogEditMode: false,
      dialogItem: undefined,
      dialogLoading: false,
      // Диалог для ConvertFilesProps
      propsDialogOpen: false,
      propsDialogEditMode: false,
      propsDialogItem: undefined,
      propsDialogConvertFileId: 0,
      propsDialogLoading: false,
      // Selected files
      selectedFiles: {}
    };

    this.convertFilesService = new ConvertFilesService(this.props.context);
    this.convertFilesPropsService = new ConvertFilesPropsService(this.props.context);
    this.convertTypesService = new ConvertTypesService(this.props.context); // NEW: Initialize service
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  private loadData = async (): Promise<void> => {
    try {
      this.setState({ loading: true, error: undefined });
      
      console.log('[Xyea] Starting data load...');
      
      // Load all data in parallel
      const [convertFiles, convertFilesProps, convertTypes] = await Promise.all([
        this.convertFilesService.getAllConvertFiles(),
        this.convertFilesPropsService.getAllConvertFilesProps(),
        this.convertTypesService.getAllConvertTypes() // This now handles errors internally
      ]);

      console.log('[Xyea] All data loaded successfully:', {
        convertFiles: convertFiles.length,
        convertFilesProps: convertFilesProps.length,
        convertTypes: convertTypes.length,
        firstConvertType: convertTypes[0]
      });

      this.setState({
        convertFiles,
        convertFilesProps,
        convertTypes,
        loading: false
      });

    } catch (error) {
      console.error('[Xyea] Error loading data:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Failed to load data'
      });
    }
  }

  private handleAddConvertFile = (): void => {
    this.setState({
      dialogOpen: true,
      dialogEditMode: false,
      dialogItem: undefined
    });
  }

  private handleEditConvertFile = (item: IConvertFile): void => {
    this.setState({
      dialogOpen: true,
      dialogEditMode: true,
      dialogItem: item
    });
  }

  private handleDialogSave = async (title: string): Promise<void> => {
    const { dialogEditMode, dialogItem } = this.state;

    try {
      this.setState({ dialogLoading: true });

      if (dialogEditMode && dialogItem) {
        await this.convertFilesService.updateConvertFile(dialogItem.Id, title);
      } else {
        await this.convertFilesService.createConvertFile(title);
      }

      this.setState({
        dialogOpen: false,
        dialogEditMode: false,
        dialogItem: undefined,
        dialogLoading: false
      });

      await this.loadData();
    } catch (error) {
      console.error('Error saving convert file:', error);
      this.setState({ dialogLoading: false });
      throw error;
    }
  }

  private handleDialogCancel = (): void => {
    this.setState({
      dialogOpen: false,
      dialogEditMode: false,
      dialogItem: undefined,
      dialogLoading: false
    });
  }

  // ConvertFilesProps methods
  private handleAddConvertFileProp = (convertFileId: number): void => {
    console.log('[Xyea] Opening add dialog with convertTypes:', {
      convertTypesLength: this.state.convertTypes.length,
      convertTypes: this.state.convertTypes.slice(0, 3)
    });
    
    this.setState({
      propsDialogOpen: true,
      propsDialogEditMode: false,
      propsDialogItem: undefined,
      propsDialogConvertFileId: convertFileId
    });
  }

  private handleEditConvertFileProp = (item: IConvertFileProps): void => {
    this.setState({
      propsDialogOpen: true,
      propsDialogEditMode: true,
      propsDialogItem: item,
      propsDialogConvertFileId: item.ConvertFilesID
    });
  }

  // NEW: Updated to handle ConvertType parameters
  private handlePropsDialogSave = async (
    convertFileId: number, 
    title: string, 
    prop: string, 
    prop2: string,
    convertTypeId: number,
    convertType2Id: number
  ): Promise<void> => {
    const { propsDialogEditMode, propsDialogItem, convertFilesProps } = this.state;

    try {
      this.setState({ propsDialogLoading: true });

      if (propsDialogEditMode && propsDialogItem) {
        // Редактирование - pass convert type IDs
        await this.convertFilesPropsService.updateConvertFileProp(
          propsDialogItem.Id, 
          title, 
          prop, 
          prop2,
          convertTypeId,
          convertType2Id
        );
      } else {
        // Создание - pass convert type IDs
        await this.convertFilesPropsService.createConvertFileProp(
          convertFileId, 
          title, 
          prop, 
          prop2,
          convertTypeId,
          convertType2Id,
          convertFilesProps
        );
      }

      this.setState({
        propsDialogOpen: false,
        propsDialogEditMode: false,
        propsDialogItem: undefined,
        propsDialogConvertFileId: 0,
        propsDialogLoading: false
      });

      await this.loadData();
    } catch (error) {
      console.error('Error saving convert file prop:', error);
      this.setState({ propsDialogLoading: false });
      throw error;
    }
  }

  private handlePropsDialogCancel = (): void => {
    this.setState({
      propsDialogOpen: false,
      propsDialogEditMode: false,
      propsDialogItem: undefined,
      propsDialogConvertFileId: 0,
      propsDialogLoading: false
    });
  }

  // Excel import for ConvertFilesProps
  private handleExcelImport = async (convertFileId: number, excelData: IExcelImportData[]): Promise<void> => {
    try {
      console.log('[Xyea] Starting Excel import:', {
        convertFileId,
        dataCount: excelData.length
      });

      this.setState({ loading: true, error: undefined });

      await this.convertFilesPropsService.importFromExcel(
        convertFileId,
        excelData,
        this.state.convertFilesProps
      );

      console.log('[Xyea] Excel import completed successfully');
      await this.loadData();

    } catch (error) {
      console.error('[Xyea] Excel import failed:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Excel import failed'
      });
      throw error;
    }
  }

  private handleToggleDeleted = async (id: number, isDeleted: boolean): Promise<void> => {
    try {
      this.setState({ loading: true });
      
      if (isDeleted) {
        await this.convertFilesPropsService.markAsDeleted(id);
      } else {
        await this.convertFilesPropsService.restoreDeleted(id);
      }
      
      await this.loadData();
    } catch (error) {
      console.error('Error toggling deleted status:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Failed to update item status'
      });
    }
  }

  private handleMoveUp = async (id: number): Promise<void> => {
    try {
      this.setState({ loading: true });
      await this.convertFilesPropsService.moveItemUp(id, this.state.convertFilesProps);
      await this.loadData();
    } catch (error) {
      console.error('Error moving item up:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Failed to move item up'
      });
    }
  }

  private handleMoveDown = async (id: number): Promise<void> => {
    try {
      this.setState({ loading: true });
      await this.convertFilesPropsService.moveItemDown(id, this.state.convertFilesProps);
      await this.loadData();
    } catch (error) {
      console.error('Error moving item down:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Failed to move item down'
      });
    }
  }

  private handleDeleteConvertFile = async (id: number): Promise<void> => {
    try {
      this.setState({ loading: true });
      await this.convertFilesService.deleteConvertFile(id);
      await this.loadData();
    } catch (error) {
      console.error('Error deleting convert file:', error);
      this.setState({
        loading: false,
        error: error instanceof Error ? error.message : 'Failed to delete convert file'
      });
    }
  }

  private handleRowClick = (convertFileId: number): void => {
    const { expandedRows } = this.state;
    
    if (expandedRows.includes(convertFileId)) {
      this.setState({ expandedRows: [] });
    } else {
      this.setState({ expandedRows: [convertFileId] });
    }
  }

  private handleSelectedFilesChange = (selectedFiles: ISelectedFiles): void => {
    this.setState({ selectedFiles });
  }

  private handleDeleteProp = (): void => {
    console.log('Delete operation handled through toggle deleted');
  }

  public render(): React.ReactElement<IXyeaProps> {
    const { 
      convertFiles, 
      convertFilesProps, 
      convertTypes,
      loading, 
      error, 
      expandedRows, 
      dialogOpen, 
      dialogEditMode, 
      dialogItem, 
      dialogLoading,
      propsDialogOpen,
      propsDialogEditMode,
      propsDialogItem,
      propsDialogConvertFileId,
      propsDialogLoading
    } = this.state;

    // Convert Files content
    const convertFilesContent = (
      <>
        {error && (
          <div className={styles.error}>
            <strong>Error:</strong> {error}
            <button 
              className={styles.retryButton}
              onClick={() => { this.loadData().catch(console.error); }}
            >
              Retry
            </button>
          </div>
        )}

        <ConvertFilesTable
          context={this.props.context}
          convertFiles={convertFiles}
          loading={loading}
          onAdd={this.handleAddConvertFile}
          onEdit={this.handleEditConvertFile}
          onDelete={this.handleDeleteConvertFile}
          onRowClick={this.handleRowClick}
          expandedRows={expandedRows}
          selectedFiles={this.state.selectedFiles}
          onSelectedFilesChange={this.handleSelectedFilesChange}
        />

        {/* Show child tables for expanded rows */}
        {expandedRows.map(convertFileId => {
          const convertFile = convertFiles.find(cf => cf.Id === convertFileId);
          const propsForFile = convertFilesProps.filter(cfp => cfp.ConvertFilesID === convertFileId);
          
          if (!convertFile) return null;

          return (
            <ConvertFilesPropsTable
              key={convertFileId}
              context={this.props.context}
              convertFileId={convertFileId}
              convertFileTitle={convertFile.Title}
              items={propsForFile}
              allItems={convertFilesProps}
              loading={loading}
              convertTypes={convertTypes} // NEW: Pass convert types
              onAdd={this.handleAddConvertFileProp}
              onEdit={this.handleEditConvertFileProp}
              onDelete={this.handleDeleteProp}
              onMoveUp={this.handleMoveUp}
              onMoveDown={this.handleMoveDown}
              onToggleDeleted={this.handleToggleDeleted}
              onImportFromExcel={this.handleExcelImport}
              selectedFiles={this.state.selectedFiles}
            />
          );
        })}
      </>
    );

    // Tab items
    const tabItems: ITabItem[] = [
      {
        key: 'convert-files',
        label: 'Convert Files Management',
        content: convertFilesContent
      },
      {
        key: 'separate-files',
        label: 'Separate Files Management',
        content: (
          <SeparateFilesManagement
            context={this.props.context}
            userDisplayName={this.props.userDisplayName}
          />
        )
      },
      {
        key: 'rename-files',
        label: 'Rename Files',
        content: (
          <RenameFilesManagement
            context={this.props.context}
            userDisplayName={this.props.userDisplayName}
          />
        )
      }
    ];

    return (
      <section className={styles.xyea}>
        <div className={styles.container}>
          <Tabs
            items={tabItems}
            defaultActiveKey="convert-files"
          />

          <EditDialog
            isOpen={dialogOpen}
            isEditMode={dialogEditMode}
            item={dialogItem}
            title={dialogEditMode ? 'Edit Convert File' : 'Create New Convert File'}
            loading={dialogLoading}
            onSave={this.handleDialogSave}
            onCancel={this.handleDialogCancel}
          />

          <EditPropsDialog
            isOpen={propsDialogOpen}
            isEditMode={propsDialogEditMode}
            convertFileId={propsDialogConvertFileId}
            item={propsDialogItem}
            title={propsDialogEditMode ? 'Edit Property' : 'Create New Property'}
            loading={propsDialogLoading}
            convertTypes={convertTypes} // NEW: Pass convert types to dialog
            onSave={this.handlePropsDialogSave}
            onCancel={this.handlePropsDialogCancel}
          />
        </div>
      </section>
    );
  }
}