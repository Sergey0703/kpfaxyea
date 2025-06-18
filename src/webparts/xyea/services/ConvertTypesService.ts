// src/webparts/xyea/services/ConvertTypesService.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { IConvertType } from '../models/IConvertType';

// SharePoint API response types
interface ISharePointConvertTypeResponse {
  Id: number;
  Title: string;
  Created?: string;
  Modified?: string;
  Author?: {
    Title: string;
    Email: string;
  };
  Editor?: {
    Title: string;
    Email: string;
  };
}

export class ConvertTypesService {
  private spService: SharePointService;
  private readonly LIST_NAME = 'convertypes'; // FIXED: Remove 't' - correct name

  constructor(context: WebPartContext) {
    this.spService = new SharePointService(context);
  }

  // Получить все типы величин
  public async getAllConvertTypes(): Promise<IConvertType[]> {
    try {
      console.log('[ConvertTypesService] Loading convert types from list:', this.LIST_NAME);
      
      const items = await this.spService.getListItems<ISharePointConvertTypeResponse>(
        this.LIST_NAME,
        'Id,Title,Created,Modified',
        undefined,
        undefined,
        'Title asc',
        1000 // Limit to 1000 types
      );

      const convertTypes = items.map((item: ISharePointConvertTypeResponse) => this.mapSharePointItemToModel(item));
      
      console.log('[ConvertTypesService] Loaded convert types:', {
        count: convertTypes.length,
        types: convertTypes.slice(0, 3) // Log first 3 types
      });

      // If no types found, return a default type
      if (convertTypes.length === 0) {
        console.warn('[ConvertTypesService] No convert types found, returning default type');
        return [{
          Id: 1,
          Title: 'String',
          Created: new Date(),
          Modified: new Date()
        }];
      }

      return convertTypes;
    } catch (error) {
      console.error('[ConvertTypesService] Error getting convert types:', error);
      
      // Return default type on error
      console.warn('[ConvertTypesService] Returning default type due to error');
      return [{
        Id: 1,
        Title: 'String',
        Created: new Date(),
        Modified: new Date()
      }];
    }
  }

  // Получить тип по ID
  public async getConvertTypeById(id: number): Promise<IConvertType> {
    try {
      const item = await this.spService.getListItemById<ISharePointConvertTypeResponse>(
        this.LIST_NAME,
        id,
        'Id,Title,Created,Modified'
      );

      return this.mapSharePointItemToModel(item);
    } catch (error) {
      console.error(`Error getting convert type ${id}:`, error);
      throw error;
    }
  }

  // Создать новый тип
  public async createConvertType(title: string): Promise<IConvertType> {
    try {
      const newItem = {
        Title: title.trim()
      };

      const createdItem = await this.spService.createListItem<ISharePointConvertTypeResponse>(
        this.LIST_NAME,
        newItem
      );

      return this.mapSharePointItemToModel(createdItem);
    } catch (error) {
      console.error('Error creating convert type:', error);
      throw error;
    }
  }

  // Обновить тип
  public async updateConvertType(id: number, title: string): Promise<IConvertType> {
    try {
      const updateItem = {
        Title: title.trim()
      };

      await this.spService.updateListItem<ISharePointConvertTypeResponse>(
        this.LIST_NAME,
        id,
        updateItem
      );

      return await this.getConvertTypeById(id);
    } catch (error) {
      console.error(`Error updating convert type ${id}:`, error);
      throw error;
    }
  }

  // Удалить тип
  public async deleteConvertType(id: number): Promise<void> {
    try {
      await this.spService.deleteListItem(this.LIST_NAME, id);
    } catch (error) {
      console.error(`Error deleting convert type ${id}:`, error);
      throw error;
    }
  }

  // Проверить существование типа по заголовку
  public async checkTitleExists(title: string, excludeId?: number): Promise<boolean> {
    try {
      const filter = excludeId 
        ? `Title eq '${title}' and Id ne ${excludeId}`
        : `Title eq '${title}'`;

      const items = await this.spService.getListItems<ISharePointConvertTypeResponse>(
        this.LIST_NAME,
        'Id',
        undefined,
        filter
      );

      return items.length > 0;
    } catch (error) {
      console.error('Error checking convert type title exists:', error);
      return false;
    }
  }

  // Получить дефолтный тип (ID=1, строковый)
  public async getDefaultConvertType(): Promise<IConvertType | undefined> {
    try {
      return await this.getConvertTypeById(1);
    } catch (error) {
      console.warn('Default convert type (ID=1) not found:', error);
      return undefined;
    }
  }

  // Helper method to map SharePoint response to model
  private mapSharePointItemToModel(item: ISharePointConvertTypeResponse): IConvertType {
    return {
      Id: item.Id,
      Title: item.Title,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      Author: item.Author,
      Editor: item.Editor
    };
  }
}