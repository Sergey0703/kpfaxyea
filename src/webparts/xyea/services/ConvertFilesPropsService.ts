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
      // Сначала создаем с основными полями
      const basicItem = {
        Title: title,
        Prop: prop,
        Prop2: prop2
      };

      console.log('Creating ConvertFileProps with basic fields only:', basicItem);
      const createdItem = await this.spService.createListItem<any>(this.LIST_NAME, basicItem);
      console.log('Created item with basic fields:', createdItem);

      // Теперь попробуем обновить с Lookup полем
      const itemId = createdItem.Id;
      
      try {
        // Попробуем разные варианты Lookup поля при обновлении
        console.log('Trying to update with ConvertFilesIDId...');
        const updateData1 = {
          ConvertFilesIDId: convertFilesId
        };
        
        await this.spService.updateListItem(this.LIST_NAME, itemId, updateData1);
        console.log('Successfully updated with ConvertFilesIDId');
        
        return {
          ...createdItem,
          ConvertFilesID: convertFilesId,
          IsDeleted: false,
          Priority: 1
        };
        
      } catch (error1) {
        console.log('ConvertFilesIDId update failed, trying ConvertFilesID...');
        
        try {
          const updateData2 = {
            ConvertFilesID: convertFilesId
          };
          
          await this.spService.updateListItem(this.LIST_NAME, itemId, updateData2);
          console.log('Successfully updated with ConvertFilesID');
          
          return {
            ...createdItem,
            ConvertFilesID: convertFilesId,
            IsDeleted: false,
            Priority: 1
          };
          
        } catch (error2) {
          console.log('Both lookup update attempts failed. Item created but without lookup.');
          
          // Возвращаем элемент без lookup
          return {
            ...createdItem,
            ConvertFilesID: 0, // Не удалось установить lookup
            IsDeleted: false,
            Priority: 1
          };
        }
      }
    } catch (error) {
      console.error('Error creating convert file prop:', error);
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