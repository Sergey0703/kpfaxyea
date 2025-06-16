// src/webparts/xyea/components/ConfirmationDialog/IConfirmationDialogProps.ts

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

export interface IConfirmationDialogConfig {
  title: string;
  message: string;
  confirmText?: string;
  cancelText?: string;
  type?: 'warning' | 'danger' | 'info' | 'success';
  showIcon?: boolean;
}

// Хелпер для создания конфигураций диалогов
export class ConfirmationDialogHelper {
  
  /**
   * Создает конфигурацию для подтверждения удаления
   */
  public static createDeleteConfirmation(itemName: string): IConfirmationDialogConfig {
    return {
      title: 'Delete Item',
      message: `Are you sure you want to delete "${itemName}"?\n\nThis action cannot be undone.`,
      confirmText: 'Delete',
      cancelText: 'Cancel',
      type: 'danger',
      showIcon: true
    };
  }

  /**
   * Создает конфигурацию для замены данных
   */
  public static createReplaceDataConfirmation(
    currentDataInfo: string,
    replacementAction: string
  ): IConfirmationDialogConfig {
    return {
      title: 'Replace Current Data',
      message: `${currentDataInfo}\n\n${replacementAction} will replace all current data.\n\nAre you sure you want to continue?`,
      confirmText: 'Yes, Replace Data',
      cancelText: 'Cancel',
      type: 'warning',
      showIcon: true
    };
  }

  /**
   * Создает конфигурацию для потери несохраненных изменений
   */
  public static createUnsavedChangesConfirmation(): IConfirmationDialogConfig {
    return {
      title: 'Unsaved Changes',
      message: 'You have unsaved changes that will be lost.\n\nAre you sure you want to continue without saving?',
      confirmText: 'Yes, Discard Changes',
      cancelText: 'Cancel',
      type: 'warning',
      showIcon: true
    };
  }

  /**
   * Создает конфигурацию для очистки фильтров
   */
  public static createClearFiltersConfirmation(activeFiltersCount: number): IConfirmationDialogConfig {
    return {
      title: 'Clear All Filters',
      message: `You have ${activeFiltersCount} active filter${activeFiltersCount > 1 ? 's' : ''} applied.\n\nClearing filters will show all data rows. Are you sure?`,
      confirmText: 'Yes, Clear Filters',
      cancelText: 'Cancel',
      type: 'info',
      showIcon: true
    };
  }

  /**
   * Создает конфигурацию для экспорта больших файлов
   */
  public static createLargeExportConfirmation(rowCount: number): IConfirmationDialogConfig {
    return {
      title: 'Large Export Warning',
      message: `You are about to export ${rowCount.toLocaleString()} rows of data.\n\nThis may take some time and could impact browser performance. Are you sure you want to continue?`,
      confirmText: 'Yes, Export Data',
      cancelText: 'Cancel',
      type: 'warning',
      showIcon: true
    };
  }
}