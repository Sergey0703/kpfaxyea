// src/webparts/xyea/components/RenameFilesManagement/handlers/CellEditingHandler.ts

import { IRenameFilesData, IRenameTableRow, ITableCell } from '../types/RenameFilesTypes';

export class CellEditingHandler {
  constructor() {
    // No parameters needed - this class is purely functional
  }

  public updateCell(
    data: IRenameFilesData, 
    columnId: string, 
    rowIndex: number, 
    newValue: string
  ): IRenameFilesData {
    
    const updatedRows = data.rows.map(row => {
      if (row.rowIndex === rowIndex) {
        const oldCell = row.cells[columnId];
        const isNowEdited = newValue !== (oldCell?.originalValue || '');
        
        const updatedCells = {
          ...row.cells,
          [columnId]: {
            ...oldCell,
            value: newValue,
            isEdited: isNowEdited
          }
        };

        return {
          ...row,
          cells: updatedCells,
          isEdited: Object.values(updatedCells).some(cell => cell.isEdited)
        };
      }
      return row;
    });

    // Calculate edited cells count
    const editedCellsCount = updatedRows.reduce((count, row) => {
      return count + Object.values(row.cells).filter(cell => cell.isEdited).length;
    }, 0);

    return {
      ...data,
      rows: updatedRows,
      editedCellsCount
    };
  }

  public clearCell(data: IRenameFilesData, columnId: string, rowIndex: number): IRenameFilesData {
    return this.updateCell(data, columnId, rowIndex, '');
  }

  public restoreCell(data: IRenameFilesData, columnId: string, rowIndex: number): IRenameFilesData {
    const row = data.rows.find(r => r.rowIndex === rowIndex);
    const cell = row?.cells[columnId];
    
    if (cell && cell.originalValue !== undefined) {
      return this.updateCell(data, columnId, rowIndex, String(cell.originalValue));
    }
    
    return data;
  }

  public bulkUpdateCells(
    data: IRenameFilesData, 
    updates: Array<{ columnId: string; rowIndex: number; newValue: string }>
  ): IRenameFilesData {
    
    let currentData = data;
    
    updates.forEach(update => {
      currentData = this.updateCell(currentData, update.columnId, update.rowIndex, update.newValue);
    });
    
    return currentData;
  }

  public validateCellValue(value: string, dataType: string): { isValid: boolean; error?: string } {
    switch (dataType) {
      case 'number':
        const numValue = parseFloat(value);
        if (value !== '' && (isNaN(numValue) || !isFinite(numValue))) {
          return { isValid: false, error: 'Must be a valid number' };
        }
        break;
        
      case 'date':
        if (value !== '' && isNaN(Date.parse(value))) {
          return { isValid: false, error: 'Must be a valid date' };
        }
        break;
        
      case 'boolean':
        const lowerValue = value.toLowerCase();
        if (value !== '' && !['true', 'false', '1', '0', 'yes', 'no'].includes(lowerValue)) {
          return { isValid: false, error: 'Must be true/false, yes/no, or 1/0' };
        }
        break;
        
      case 'text':
      default:
        // Text values are generally always valid, but check length
        if (value.length > 1000) {
          return { isValid: false, error: 'Text too long (max 1000 characters)' };
        }
        break;
    }
    
    return { isValid: true };
  }

  public formatCellValue(value: string, dataType: string): string {
    if (!value) return '';
    
    switch (dataType) {
      case 'number':
        const numValue = parseFloat(value);
        return isNaN(numValue) ? value : numValue.toString();
        
      case 'date':
        const dateValue = new Date(value);
        return isNaN(dateValue.getTime()) ? value : dateValue.toISOString().split('T')[0];
        
      case 'boolean':
        const lowerValue = value.toLowerCase();
        if (['true', '1', 'yes'].includes(lowerValue)) return 'true';
        if (['false', '0', 'no'].includes(lowerValue)) return 'false';
        return value;
        
      case 'text':
      default:
        return value.trim();
    }
  }

  public getCellEditHistory(data: IRenameFilesData): Array<{
    columnId: string;
    rowIndex: number;
    originalValue: any;
    currentValue: any;
    columnName: string;
  }> {
    const history: Array<{
      columnId: string;
      rowIndex: number;
      originalValue: any;
      currentValue: any;
      columnName: string;
    }> = [];

    data.rows.forEach(row => {
      Object.values(row.cells).forEach(cell => {
        if (cell.isEdited) {
          const column = data.columns.find(col => col.id === cell.columnId);
          history.push({
            columnId: cell.columnId,
            rowIndex: cell.rowIndex,
            originalValue: cell.originalValue,
            currentValue: cell.value,
            columnName: column?.name || cell.columnId
          });
        }
      });
    });

    return history;
  }

  public revertAllChanges(data: IRenameFilesData): IRenameFilesData {
    const updatedRows = data.rows.map(row => {
      const updatedCells: { [columnId: string]: ITableCell } = {};
      
      Object.entries(row.cells).forEach(([columnId, cell]) => {
        updatedCells[columnId] = {
          ...cell,
          value: cell.originalValue,
          isEdited: false
        };
      });

      return {
        ...row,
        cells: updatedCells,
        isEdited: false
      };
    });

    return {
      ...data,
      rows: updatedRows,
      editedCellsCount: 0
    };
  }

  public getEditedRowsCount(data: IRenameFilesData): number {
    return data.rows.filter(row => row.isEdited).length;
  }

  public getEditedCellsInRow(row: IRenameTableRow): ITableCell[] {
    return Object.values(row.cells).filter(cell => cell.isEdited);
  }

  public hasUnsavedChanges(data: IRenameFilesData): boolean {
    return data.editedCellsCount > 0;
  }

  public exportEditedData(data: IRenameFilesData): Array<{
    rowIndex: number;
    changes: Array<{
      columnName: string;
      originalValue: any;
      newValue: any;
    }>;
  }> {
    const exportData: Array<{
      rowIndex: number;
      changes: Array<{
        columnName: string;
        originalValue: any;
        newValue: any;
      }>;
    }> = [];

    data.rows.forEach(row => {
      if (row.isEdited) {
        const changes: Array<{
          columnName: string;
          originalValue: any;
          newValue: any;
        }> = [];

        Object.values(row.cells).forEach(cell => {
          if (cell.isEdited) {
            const column = data.columns.find(col => col.id === cell.columnId);
            changes.push({
              columnName: column?.name || cell.columnId,
              originalValue: cell.originalValue,
              newValue: cell.value
            });
          }
        });

        if (changes.length > 0) {
          exportData.push({
            rowIndex: row.rowIndex,
            changes
          });
        }
      }
    });

    return exportData;
  }
}