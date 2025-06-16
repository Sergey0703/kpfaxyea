// src/webparts/xyea/services/ConvertFilesPropsService.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { IConvertFileProps } from '../models';
import { PriorityHelper } from '../utils';

export class ConvertFilesPropsService {
  private spService: SharePointService;
  private readonly LIST_NAME = 'convertfilesprops';

  constructor(context: WebPartContext) {
    this.spService = new SharePointService(context);
  }

  // Получить все свойства
  public async getAllConvertFilesProps(): Promise<IConvertFileProps[]> {
    try {
      const items = await this.spService.getListItems<any>(
        this.LIST_NAME,
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,Created,Modified',
        undefined,
        undefined,
        'ConvertFilesIDId asc, Priority asc'
      );

      // Маппинг ConvertFilesIDId в ConvertFilesID
      return items.map((item: any) => ({
        ...item,
        ConvertFilesID: item.ConvertFilesIDId
      }));
    } catch (error) {
      console.error('Error getting convert files props:', error);
      throw error;
    }
  }

  // Получить свойства по ConvertFilesID
  public async getConvertFilesPropsById(convertFilesId: number): Promise<IConvertFileProps[]> {
    try {
      const items = await this.spService.getListItems<any>(
        this.LIST_NAME,
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,Created,Modified',
        undefined,
        `ConvertFilesIDId eq ${convertFilesId}`,
        'Priority asc'
      );

      // Маппинг ConvertFilesIDId в ConvertFilesID
      return items.map((item: any) => ({
        ...item,
        ConvertFilesID: item.ConvertFilesIDId
      }));
    } catch (error) {
      console.error(`Error getting convert files props for ${convertFilesId}:`, error);
      throw error;
    }
  }

  // Получить элемент по ID
  public async getConvertFilePropById(id: number): Promise<IConvertFileProps> {
    try {
      const item = await this.spService.getListItemById<any>(
        this.LIST_NAME,
        id,
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,Created,Modified'
      );

      // Маппинг ConvertFilesIDId в ConvertFilesID
      return {
        ...item,
        ConvertFilesID: item.ConvertFilesIDId
      };
    } catch (error) {
      console.error(`Error getting convert file prop ${id}:`, error);
      throw error;
    }
  }

  // Создать новое свойство
  public async createConvertFileProp(
    convertFilesId: number,
    title: string,
    prop: string,
    prop2: string,
    allItems?: IConvertFileProps[]
  ): Promise<IConvertFileProps> {
    try {
      // Если не передан массив всех элементов, получаем их
      if (!allItems) {
        allItems = await this.getAllConvertFilesProps();
      }

      // Вычисляем следующий приоритет
      const nextPriority = PriorityHelper.getNextPriority(allItems, convertFilesId);

      // Попробуем разные варианты поля Lookup
      const newItem = {
        Title: title,
        ConvertFilesID: convertFilesId, // Попробуем без Id суффикса
        Prop: prop,
        Prop2: prop2,
        IsDeleted: false,
        Priority: nextPriority
      };

      console.log('Creating ConvertFileProps with data (attempt 1):', newItem);

      try {
        const createdItem = await this.spService.createListItem<any>(this.LIST_NAME, newItem);
        console.log('Created item response (attempt 1):', createdItem);
        return {
          ...createdItem,
          ConvertFilesID: createdItem.ConvertFilesID || convertFilesId
        };
      } catch (error1) {
        console.log('Attempt 1 failed, trying with ConvertFilesIDId...');
        
        // Попробуем с Id суффиксом
        const newItem2 = {
          Title: title,
          ConvertFilesIDId: convertFilesId,
          Prop: prop,
          Prop2: prop2,
          IsDeleted: false,
          Priority: nextPriority
        };

        console.log('Creating ConvertFileProps with data (attempt 2):', newItem2);
        
        try {
          const createdItem = await this.spService.createListItem<any>(this.LIST_NAME, newItem2);
          console.log('Created item response (attempt 2):', createdItem);
          return {
            ...createdItem,
            ConvertFilesID: createdItem.ConvertFilesIDId || convertFilesId
          };
        } catch (error2) {
          console.log('Attempt 2 failed, trying with lookup format...');
          
          // Попробуем формат {results: [id]}
          const newItem3 = {
            Title: title,
            ConvertFilesIDId: { results: [convertFilesId] },
            Prop: prop,
            Prop2: prop2,
            IsDeleted: false,
            Priority: nextPriority
          };

          console.log('Creating ConvertFileProps with data (attempt 3):', newItem3);
          const createdItem = await this.spService.createListItem<any>(this.LIST_NAME, newItem3);
          console.log('Created item response (attempt 3):', createdItem);
          return {
            ...createdItem,
            ConvertFilesID: createdItem.ConvertFilesIDId || convertFilesId
          };
        }
      }
    } catch (error) {
      console.error('Error creating convert file prop (all attempts failed):', error);
      throw error;
    }
  }

  // Обновить свойство
  public async updateConvertFileProp(
    id: number,
    title: string,
    prop: string,
    prop2: string
  ): Promise<IConvertFileProps> {
    try {
      const updateItem = {
        Title: title,
        Prop: prop,
        Prop2: prop2
      };

      return await this.spService.updateListItem<IConvertFileProps>(this.LIST_NAME, id, updateItem);
    } catch (error) {
      console.error(`Error updating convert file prop ${id}:`, error);
      throw error;
    }
  }

  // Пометить элемент как удаленный
  public async markAsDeleted(id: number): Promise<IConvertFileProps> {
    try {
      const updateItem = {
        IsDeleted: true
      };

      return await this.spService.updateListItem<IConvertFileProps>(this.LIST_NAME, id, updateItem);
    } catch (error) {
      console.error(`Error marking convert file prop ${id} as deleted:`, error);
      throw error;
    }
  }

  // Восстановить удаленный элемент
  public async restoreDeleted(id: number): Promise<IConvertFileProps> {
    try {
      const updateItem = {
        IsDeleted: false
      };

      return await this.spService.updateListItem<IConvertFileProps>(this.LIST_NAME, id, updateItem);
    } catch (error) {
      console.error(`Error restoring convert file prop ${id}:`, error);
      throw error;
    }
  }

  // Переместить элемент вверх
  public async moveItemUp(id: number, allItems: IConvertFileProps[]): Promise<IConvertFileProps[]> {
    try {
      const result = PriorityHelper.moveUp(allItems, id);
      
      // Обновляем приоритеты в SharePoint
      for (const update of result.itemsToUpdate) {
        await this.spService.updateListItem(this.LIST_NAME, update.id, { Priority: update.priority });
      }

      return result.updatedItems;
    } catch (error) {
      console.error(`Error moving item ${id} up:`, error);
      throw error;
    }
  }

  // Переместить элемент вниз
  public async moveItemDown(id: number, allItems: IConvertFileProps[]): Promise<IConvertFileProps[]> {
    try {
      const result = PriorityHelper.moveDown(allItems, id);
      
      // Обновляем приоритеты в SharePoint
      for (const update of result.itemsToUpdate) {
        await this.spService.updateListItem(this.LIST_NAME, update.id, { Priority: update.priority });
      }

      return result.updatedItems;
    } catch (error) {
      console.error(`Error moving item ${id} down:`, error);
      throw error;
    }
  }

  // Нормализовать приоритеты для ConvertFilesID
  public async normalizePriorities(convertFilesId: number, allItems: IConvertFileProps[]): Promise<IConvertFileProps[]> {
    try {
      const result = PriorityHelper.normalizePriorities(allItems, convertFilesId);
      
      // Обновляем приоритеты в SharePoint
      for (const update of result.itemsToUpdate) {
        await this.spService.updateListItem(this.LIST_NAME, update.id, { Priority: update.priority });
      }

      return result.updatedItems;
    } catch (error) {
      console.error(`Error normalizing priorities for ${convertFilesId}:`, error);
      throw error;
    }
  }

  // Проверить возможность перемещения
  public canMoveUp(itemId: number, convertFilesId: number, allItems: IConvertFileProps[]): boolean {
    return PriorityHelper.canMoveUp(allItems, itemId, convertFilesId);
  }

  public canMoveDown(itemId: number, convertFilesId: number, allItems: IConvertFileProps[]): boolean {
    return PriorityHelper.canMoveDown(allItems, itemId, convertFilesId);
  }
}