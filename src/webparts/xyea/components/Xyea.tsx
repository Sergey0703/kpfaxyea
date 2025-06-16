// src/webparts/xyea/components/Xyea.tsx

import * as React from 'react';
import styles from './Xyea.module.scss';
import { IXyeaProps } from './IXyeaProps';
import { IConvertFile, IConvertFileProps } from '../models';
import { ConvertFilesService, ConvertFilesPropsService } from '../services';
import ConvertFilesTable from './ConvertFilesTable/ConvertFilesTable';
import ConvertFilesPropsTable from './ConvertFilesPropsTable/ConvertFilesPropsTable';
import EditDialog from './EditDialog/EditDialog';
import EditPropsDialog from './EditDialog/EditPropsDialog';

export interface IXyeaState {
  convertFiles: IConvertFile[];
  convertFilesProps: IConvertFileProps[];
  loading: boolean;
  error: string | null;
  expandedRows: number[];
  // Состояние диалога для ConvertFiles
  dialogOpen: boolean;
  dialogEditMode: boolean;
  dialogItem: IConvertFile | null;
  dialogLoading: boolean;
  // Состояние диалога для ConvertFilesProps
  propsDialogOpen: boolean;
  propsDialogEditMode: boolean;
  propsDialogItem: IConvertFileProps | null;
  propsDialogConvertFileId: number;
  propsDialogLoading: boolean;
}

export default class Xyea extends React.Component<IXyeaProps, IXyeaState> {
  private convertFilesService: ConvertFilesService;
  private convertFilesPropsService: ConvertFilesPropsService;

  constructor(props: IXyeaProps) {
    super(props);
    
    this.state = {
      convertFiles: [],
      convertFilesProps: [],
      loading: true,
      error: null,
      expandedRows: [],
      // Диалог для ConvertFiles
      dialogOpen: false,
      dialogEditMode: false,
      dialogItem: null,
      dialogLoading: false,
      // Диалог для ConvertFilesProps
      propsDialogOpen: false,
      propsDialogEditMode: false,
      propsDialogItem: null,
      propsDialogConvertFileId: 0,
      propsDialogLoading: false
    };

    this.convertFilesService = new ConvertFilesService(this.props.context);
    this.convertFilesPropsService = new ConvertFilesPropsService(this.props.context);
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  private loadData = async (): Promise<void> => {
    try {
      this.setState({ loading: true, error: null });
      
      const [convertFiles, convertFilesProps] = await Promise.all([
        this.convertFilesService.getAllConvertFiles(),
        this.convertFilesPropsService.getAllConvertFilesProps()
      ]);

      this.setState({
        convertFiles,
        convertFilesProps,
        loading: false
      });
    } catch (error) {
      console.error('Error loading data:', error);
      this.setState({
        loading: false,
        error: error.message || 'Failed to load data'
      });
    }
  }

  private handleAddConvertFile = (): void => {
    this.setState({
      dialogOpen: true,
      dialogEditMode: false,
      dialogItem: null
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
        // Редактирование
        await this.convertFilesService.updateConvertFile(dialogItem.Id, title);
      } else {
        // Создание
        await this.convertFilesService.createConvertFile(title);
      }

      // Закрыть диалог и полностью сбросить состояние
      this.setState({
        dialogOpen: false,
        dialogEditMode: false,
        dialogItem: null,
        dialogLoading: false // Важно! Сбросить loading состояние
      });

      // Обновить данные
      await this.loadData();
    } catch (error) {
      console.error('Error saving convert file:', error);
      // При ошибке тоже сбрасываем loading, но диалог оставляем открытым
      this.setState({ dialogLoading: false });
      throw error; // Пробросить ошибку в диалог
    }
  }

  private handleDialogCancel = (): void => {
    this.setState({
      dialogOpen: false,
      dialogEditMode: false,
      dialogItem: null,
      dialogLoading: false
    });
  }

  // Методы для работы с ConvertFilesProps
  private handleAddConvertFileProp = (convertFileId: number): void => {
    this.setState({
      propsDialogOpen: true,
      propsDialogEditMode: false,
      propsDialogItem: null,
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

  private handlePropsDialogSave = async (convertFileId: number, title: string, prop: string, prop2: string): Promise<void> => {
    const { propsDialogEditMode, propsDialogItem, convertFilesProps } = this.state;

    try {
      this.setState({ propsDialogLoading: true });

      if (propsDialogEditMode && propsDialogItem) {
        // Редактирование
        await this.convertFilesPropsService.updateConvertFileProp(propsDialogItem.Id, title, prop, prop2);
      } else {
        // Создание
        await this.convertFilesPropsService.createConvertFileProp(convertFileId, title, prop, prop2, convertFilesProps);
      }

      // Закрыть диалог и полностью сбросить состояние
      this.setState({
        propsDialogOpen: false,
        propsDialogEditMode: false,
        propsDialogItem: null,
        propsDialogConvertFileId: 0,
        propsDialogLoading: false
      });

      // Обновить данные
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
      propsDialogItem: null,
      propsDialogConvertFileId: 0,
      propsDialogLoading: false
    });
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
        error: error.message || 'Failed to update item status'
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
        error: error.message || 'Failed to move item up'
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
        error: error.message || 'Failed to move item down'
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
        error: error.message || 'Failed to delete convert file'
      });
    }
  }

  private handleRowClick = (convertFileId: number): void => {
    const { expandedRows } = this.state;
    
    if (expandedRows.includes(convertFileId)) {
      // Закрыть строку - убрать все раскрытые строки
      this.setState({
        expandedRows: []
      });
    } else {
      // Открыть только эту строку - заменить все раскрытые строки на эту одну
      this.setState({
        expandedRows: [convertFileId]
      });
    }
  }

  public render(): React.ReactElement<IXyeaProps> {
    const { 
      convertFiles, 
      convertFilesProps, 
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

    return (
      <section className={styles.xyea}>
        <div className={styles.container}>
          <h1 className={styles.title}>Convert Files Management</h1>
          
          {error && (
            <div className={styles.error}>
              <strong>Error:</strong> {error}
              <button 
                className={styles.retryButton}
                onClick={this.loadData}
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
          />

          {/* Показать подчиненные таблицы для раскрытых строк */}
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
                onAdd={this.handleAddConvertFileProp}
                onEdit={this.handleEditConvertFileProp}
                onDelete={() => {}} // Не используется, используем onToggleDeleted
                onMoveUp={this.handleMoveUp}
                onMoveDown={this.handleMoveDown}
                onToggleDeleted={this.handleToggleDeleted}
              />
            );
          })}

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
            onSave={this.handlePropsDialogSave}
            onCancel={this.handlePropsDialogCancel}
          />
        </div>
      </section>
    );
  }
}