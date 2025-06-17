// src/webparts/xyea/services/ConvertFilesPropsService.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { IConvertFileProps } from '../models';
import { PriorityHelper } from '../utils';

// SharePoint API response types for better type safety
interface ISharePointConvertFilePropsResponse {
  Id: number;
  Title: string;
  ConvertFilesIDId: number;
  Prop: string;
  Prop2: string;
  IsDeleted: number; // SharePoint stores boolean as number (0/1)
  Priority: number;
  Created?: string; // ISO date string from SharePoint
  Modified?: string; // ISO date string from SharePoint
  Author?: {
    Title: string;
    Email: string;
  };
  Editor?: {
    Title: string;
    Email: string;
  };
}

interface ISharePointCreateResponse {
  Id: number;
  Title: string;
  Created: string;
  Modified: string;
}

interface ISharePointUpdateResponse {
  Id: number;
  Modified: string;
}





export class ConvertFilesPropsService {
  private spService: SharePointService;
  private readonly LIST_NAME = 'convertfilesprops';

  constructor(context: WebPartContext) {
    this.spService = new SharePointService(context);
  }

  // Получить все свойства
  public async getAllConvertFilesProps(): Promise<IConvertFileProps[]> {
    try {
      const items = await this.spService.getListItems<ISharePointConvertFilePropsResponse>(
        this.LIST_NAME,
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,Created,Modified',
        undefined,
        undefined,
        'ConvertFilesIDId asc, Priority asc'
      );

      // Маппинг ConvertFilesIDId в ConvertFilesID и IsDeleted в boolean
      return items.map((item: ISharePointConvertFilePropsResponse) => this.mapSharePointItemToModel(item));
    } catch (error) {
      console.error('Error getting convert files props:', error);
      throw error;
    }
  }

  // Получить свойства по ConvertFilesID
  public async getConvertFilesPropsById(convertFilesId: number): Promise<IConvertFileProps[]> {
    try {
      const items = await this.spService.getListItems<ISharePointConvertFilePropsResponse>(
        this.LIST_NAME,
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,Created,Modified',
        undefined,
        `ConvertFilesIDId eq ${convertFilesId}`,
        'Priority asc'
      );

      // Маппинг ConvertFilesIDId в ConvertFilesID и IsDeleted в boolean
      return items.map((item: ISharePointConvertFilePropsResponse) => this.mapSharePointItemToModel(item));
    } catch (error) {
      console.error(`Error getting convert files props for ${convertFilesId}:`, error);
      throw error;
    }
  }

  // Получить элемент по ID
  public async getConvertFilePropById(id: number): Promise<IConvertFileProps> {
    try {
      const item = await this.spService.getListItemById<ISharePointConvertFilePropsResponse>(
        this.LIST_NAME,
        id,
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,Created,Modified'
      );

      // Маппинг ConvertFilesIDId в ConvertFilesID и IsDeleted в boolean
      return this.mapSharePointItemToModel(item);
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

      // Сначала создаем с основными полями
      const basicItem = {
        Title: title,
        Prop: prop,
        Prop2: prop2,
        Priority: nextPriority,
        IsDeleted: 0  // Используем 0 для SharePoint (false)
      };

      console.log('Creating ConvertFileProps with all fields:', basicItem);
      const createdItem = await this.spService.createListItem<ISharePointCreateResponse>(this.LIST_NAME, basicItem);
      console.log('Created item:', createdItem);

      // Теперь попробуем обновить с Lookup полем
      const itemId = createdItem.Id;
      
      try {
        console.log('Trying to update with ConvertFilesIDId...');
        const updateData1 = {
          ConvertFilesIDId: convertFilesId
        };
        
        const updatedItem = await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, itemId, updateData1);
        console.log('Successfully updated with ConvertFilesIDId:', updatedItem);
        
        return {
          Id: itemId,
          Title: title,
          ConvertFilesID: convertFilesId,
          Prop: prop,
          Prop2: prop2,
          Priority: nextPriority,
          IsDeleted: false,
          Created: createdItem.Created ? new Date(createdItem.Created) : undefined,
          Modified: updatedItem?.Modified ? new Date(updatedItem.Modified) : (createdItem.Modified ? new Date(createdItem.Modified) : undefined)
        };
        
      } catch (error1) {
        console.log('ConvertFilesIDId update failed, trying ConvertFilesID...');
        
        try {
          const updateData2 = {
            ConvertFilesID: convertFilesId
          };
          
          const updatedItem = await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, itemId, updateData2);
          console.log('Successfully updated with ConvertFilesID:', updatedItem);
          
          return {
            Id: itemId,
            Title: title,
            ConvertFilesID: convertFilesId,
            Prop: prop,
            Prop2: prop2,
            Priority: nextPriority,
            IsDeleted: false,
            Created: createdItem.Created ? new Date(createdItem.Created) : undefined,
            Modified: updatedItem?.Modified ? new Date(updatedItem.Modified) : (createdItem.Modified ? new Date(createdItem.Modified) : undefined)
          };
          
        } catch (error2) {
          console.log('Both lookup update attempts failed. Item created but without lookup.');
          
          // Возвращаем элемент без lookup
          return {
            Id: itemId,
            Title: title,
            ConvertFilesID: 0, // Не удалось установить lookup
            Prop: prop,
            Prop2: prop2,
            Priority: nextPriority,
            IsDeleted: false,
            Created: createdItem.Created ? new Date(createdItem.Created) : undefined,
            Modified: createdItem.Modified ? new Date(createdItem.Modified) : undefined
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

      await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, id, updateItem);
      
      // Получаем полную информацию об обновленном элементе
      return await this.getConvertFilePropById(id);
    } catch (error) {
      console.error(`Error updating convert file prop ${id}:`, error);
      throw error;
    }
  }

  // Пометить элемент как удаленный
  public async markAsDeleted(id: number): Promise<IConvertFileProps> {
    try {
      const updateItem = {
        IsDeleted: 1  // Используем 1 для SharePoint (true)
      };

      await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, id, updateItem);
      return await this.getConvertFilePropById(id);
    } catch (error) {
      console.error(`Error marking convert file prop ${id} as deleted:`, error);
      throw error;
    }
  }

  // Восстановить удаленный элемент
  public async restoreDeleted(id: number): Promise<IConvertFileProps> {
    try {
      const updateItem = {
        IsDeleted: 0  // Используем 0 для SharePoint (false)
      };

      await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, id, updateItem);
      return await this.getConvertFilePropById(id);
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
        await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, update.id, { Priority: update.priority });
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
        await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, update.id, { Priority: update.priority });
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
        await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, update.id, { Priority: update.priority });
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

  // Helper method to map SharePoint response to model
  private mapSharePointItemToModel(item: ISharePointConvertFilePropsResponse): IConvertFileProps {
    return {
      Id: item.Id,
      Title: item.Title,
      ConvertFilesID: item.ConvertFilesIDId,
      ConvertFilesIDId: item.ConvertFilesIDId,
      Prop: item.Prop,
      Prop2: item.Prop2,
      IsDeleted: item.IsDeleted === 1,
      Priority: item.Priority,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      Author: item.Author,
      Editor: item.Editor
    };
  }
}