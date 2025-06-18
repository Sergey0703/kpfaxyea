// src/webparts/xyea/services/ConvertFilesPropsService.ts - Updated with ConvertType support

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { IConvertFileProps } from '../models';
import { PriorityHelper } from '../utils';
import { IExcelImportData } from '../components/ExcelImportButton/ExcelImportButton';

// SharePoint API response types for better type safety
interface ISharePointConvertFilePropsResponse {
  Id: number;
  Title: string;
  ConvertFilesIDId: number;
  Prop: string;
  Prop2: string;
  IsDeleted: number; // SharePoint stores boolean as number (0/1)
  Priority: number;
  ConvertTypeId: number; // NEW: ConvertType Lookup field
  ConvertType2Id: number; // NEW: ConvertType2 Lookup field
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
  private readonly DEFAULT_CONVERT_TYPE_ID = 1; // Строковый тип по умолчанию

  constructor(context: WebPartContext) {
    this.spService = new SharePointService(context);
  }

  // Получить все свойства
  public async getAllConvertFilesProps(): Promise<IConvertFileProps[]> {
    try {
      const items = await this.spService.getListItems<ISharePointConvertFilePropsResponse>(
        this.LIST_NAME,
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,ConvertTypeId,ConvertType2Id,Created,Modified',
        undefined,
        undefined,
        'ConvertFilesIDId asc, Priority asc',
        5000 // Increase limit to 5000 items
      );

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
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,ConvertTypeId,ConvertType2Id,Created,Modified',
        undefined,
        `ConvertFilesIDId eq ${convertFilesId}`,
        'Priority asc',
        1000 // Increase limit for single ConvertFile
      );

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
        'Id,Title,ConvertFilesIDId,Prop,Prop2,IsDeleted,Priority,ConvertTypeId,ConvertType2Id,Created,Modified'
      );

      return this.mapSharePointItemToModel(item);
    } catch (error) {
      console.error(`Error getting convert file prop ${id}:`, error);
      throw error;
    }
  }

  // NEW: Bulk import from Excel data with replacement of existing items
  public async importFromExcel(
    convertFilesId: number, 
    excelData: IExcelImportData[],
    allItems: IConvertFileProps[]
  ): Promise<IConvertFileProps[]> {
    try {
      console.log(`[ConvertFilesPropsService] Starting Excel import for ConvertFiles ${convertFilesId}:`, {
        dataCount: excelData.length,
        convertFilesId
      });

      // Step 1: Delete all existing items for this ConvertFilesID
      const existingItems = allItems.filter(item => item.ConvertFilesID === convertFilesId);
      
      if (existingItems.length > 0) {
        console.log(`[ConvertFilesPropsService] Deleting ${existingItems.length} existing items`);
        
        // Delete items in parallel batches of 10 to avoid overwhelming SharePoint
        const deletePromises: Promise<void>[] = [];
        for (let i = 0; i < existingItems.length; i += 10) {
          const batch = existingItems.slice(i, i + 10);
          for (const item of batch) {
            deletePromises.push(this.deleteConvertFileProp(item.Id));
          }
          
          // Wait for this batch before starting the next
          if (deletePromises.length >= 10) {
            await Promise.all(deletePromises);
            deletePromises.length = 0; // Clear the array
            
            // Small delay between batches
            await this.delay(200);
          }
        }
        
        // Wait for any remaining deletes
        if (deletePromises.length > 0) {
          await Promise.all(deletePromises);
        }
      }

      // Step 2: Create new items from Excel data in batches
      const createdItems: IConvertFileProps[] = [];
      const batchSize = 5; // Smaller batch size for creation to ensure reliability
      
      console.log(`[ConvertFilesPropsService] Creating ${excelData.length} new items in batches of ${batchSize}`);
      
      for (let i = 0; i < excelData.length; i += batchSize) {
        const batch = excelData.slice(i, i + batchSize);
        const batchPromises: Promise<IConvertFileProps>[] = [];
        
        console.log(`[ConvertFilesPropsService] Processing batch ${Math.floor(i / batchSize) + 1}/${Math.ceil(excelData.length / batchSize)}`);
        
        for (let j = 0; j < batch.length; j++) {
          const data = batch[j];
          const priority = i + j + 1; // Sequential priority across all batches

          // Generate title from Prop value, fallback to row number
          const title = data.prop.trim() || `Row ${i + j + 1}`;

          console.log(`[ConvertFilesPropsService] Creating item ${i + j + 1}/${excelData.length}:`, {
            title,
            prop: data.prop,
            prop2: data.prop2,
            priority
          });

          batchPromises.push(
            this.createConvertFilePropDirect(
              convertFilesId,
              title,
              data.prop,
              data.prop2,
              priority,
              this.DEFAULT_CONVERT_TYPE_ID, // NEW: Default ConvertType
              this.DEFAULT_CONVERT_TYPE_ID  // NEW: Default ConvertType2
            ).catch(error => {
              console.error(`[ConvertFilesPropsService] Failed to create item ${i + j + 1}:`, error);
              // Return a placeholder item to maintain count
              return {
                Id: -1,
                Title: title,
                ConvertFilesID: convertFilesId,
                Prop: data.prop,
                Prop2: data.prop2,
                Priority: priority,
                IsDeleted: false,
                ConvertType: this.DEFAULT_CONVERT_TYPE_ID,
                ConvertType2: this.DEFAULT_CONVERT_TYPE_ID,
                Created: undefined,
                Modified: undefined
              } as IConvertFileProps;
            })
          );
        }
        
        // Wait for this batch to complete
        const batchResults = await Promise.all(batchPromises);
        
        // Add successful items to results (filter out failed ones with Id: -1)
        const successfulItems = batchResults.filter(item => item.Id !== -1);
        createdItems.push(...successfulItems);
        
        // Log batch progress
        console.log(`[ConvertFilesPropsService] Batch completed. Created ${successfulItems.length}/${batch.length} items. Total: ${createdItems.length}/${excelData.length}`);
        
        // Delay between batches to avoid overwhelming SharePoint
        if (i + batchSize < excelData.length) {
          await this.delay(500); // 500ms delay between batches
        }
      }

      console.log(`[ConvertFilesPropsService] Excel import completed. Successfully created ${createdItems.length}/${excelData.length} items`);
      
      return createdItems;

    } catch (error) {
      console.error('[ConvertFilesPropsService] Excel import failed:', error);
      throw new Error(`Excel import failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  // NEW: Utility method for delays
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  // NEW: Direct creation method for bulk operations (bypasses priority calculation)
  private async createConvertFilePropDirect(
    convertFilesId: number,
    title: string,
    prop: string = '',
    prop2: string = '',
    priority: number,
    convertTypeId: number = this.DEFAULT_CONVERT_TYPE_ID,
    convertType2Id: number = this.DEFAULT_CONVERT_TYPE_ID
  ): Promise<IConvertFileProps> {
    try {
      // Sanitize inputs
      const sanitizedTitle = title.trim();
      const sanitizedProp = prop.trim();
      const sanitizedProp2 = prop2.trim();

      // Validate that title is not empty
      if (!sanitizedTitle) {
        throw new Error('Title is required and cannot be empty');
      }

      // Create item with all fields including priority and convert types
      const basicItem = {
        Title: sanitizedTitle,
        Prop: sanitizedProp,
        Prop2: sanitizedProp2,
        Priority: priority,
        IsDeleted: 0,
        ConvertTypeId: convertTypeId,
        ConvertType2Id: convertType2Id
      };

      const createdItem = await this.spService.createListItem<ISharePointCreateResponse>(this.LIST_NAME, basicItem);
      const itemId = createdItem.Id;

      // Update with ConvertFilesID lookup
      try {
        const updateData = { ConvertFilesIDId: convertFilesId };
        const updatedItem = await this.spService.updateListItem<ISharePointUpdateResponse>(this.LIST_NAME, itemId, updateData);
        
        return {
          Id: itemId,
          Title: sanitizedTitle,
          ConvertFilesID: convertFilesId,
          Prop: sanitizedProp,
          Prop2: sanitizedProp2,
          Priority: priority,
          IsDeleted: false,
          ConvertType: convertTypeId,
          ConvertType2: convertType2Id,
          Created: createdItem.Created ? new Date(createdItem.Created) : undefined,
          Modified: updatedItem?.Modified ? new Date(updatedItem.Modified) : (createdItem.Modified ? new Date(createdItem.Modified) : undefined)
        };
      } catch (lookupError) {
        console.warn(`[ConvertFilesPropsService] Lookup update failed for item ${itemId}, but item was created`);
        
        return {
          Id: itemId,
          Title: sanitizedTitle,
          ConvertFilesID: 0, // Lookup failed
          Prop: sanitizedProp,
          Prop2: sanitizedProp2,
          Priority: priority,
          IsDeleted: false,
          ConvertType: convertTypeId,
          ConvertType2: convertType2Id,
          Created: createdItem.Created ? new Date(createdItem.Created) : undefined,
          Modified: createdItem.Modified ? new Date(createdItem.Modified) : undefined
        };
      }
    } catch (error) {
      console.error('[ConvertFilesPropsService] Error in createConvertFilePropDirect:', error);
      throw error;
    }
  }

  // NEW: Delete a convert file prop (hard delete)
  private async deleteConvertFileProp(id: number): Promise<void> {
    try {
      await this.spService.deleteListItem(this.LIST_NAME, id);
    } catch (error) {
      console.error(`[ConvertFilesPropsService] Error deleting convert file prop ${id}:`, error);
      throw error;
    }
  }

  // Создать новое свойство - updated to handle ConvertType fields
  public async createConvertFileProp(
    convertFilesId: number,
    title: string,
    prop: string = '',
    prop2: string = '',
    convertTypeId?: number,
    convertType2Id?: number,
    allItems?: IConvertFileProps[]
  ): Promise<IConvertFileProps> {
    try {
      // Если не передан массив всех элементов, получаем их
      if (!allItems) {
        allItems = await this.getAllConvertFilesProps();
      }

      // Вычисляем следующий приоритет
      const nextPriority = PriorityHelper.getNextPriority(allItems, convertFilesId);

      // Используем переданные типы или дефолтные
      const finalConvertTypeId = convertTypeId || this.DEFAULT_CONVERT_TYPE_ID;
      const finalConvertType2Id = convertType2Id || this.DEFAULT_CONVERT_TYPE_ID;

      return await this.createConvertFilePropDirect(
        convertFilesId, 
        title, 
        prop, 
        prop2, 
        nextPriority, 
        finalConvertTypeId, 
        finalConvertType2Id
      );
    } catch (error) {
      console.error('Error creating convert file prop:', error);
      throw error;
    }
  }

  // Обновить свойство - updated to handle ConvertType fields
  public async updateConvertFileProp(
    id: number,
    title: string,
    prop: string = '',
    prop2: string = '',
    convertTypeId?: number,
    convertType2Id?: number
  ): Promise<IConvertFileProps> {
    try {
      // Sanitize inputs - allow empty strings for optional fields
      const sanitizedTitle = title.trim();
      const sanitizedProp = prop.trim();
      const sanitizedProp2 = prop2.trim();

      // Validate that title is not empty
      if (!sanitizedTitle) {
        throw new Error('Title is required and cannot be empty');
      }

      const updateItem: any = {
        Title: sanitizedTitle,
        Prop: sanitizedProp,
        Prop2: sanitizedProp2
      };

      // Add ConvertType fields if provided
      if (convertTypeId !== undefined) {
        updateItem.ConvertTypeId = convertTypeId;
      }
      if (convertType2Id !== undefined) {
        updateItem.ConvertType2Id = convertType2Id;
      }

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
      Prop: item.Prop || '',
      Prop2: item.Prop2 || '',
      IsDeleted: item.IsDeleted === 1,
      Priority: item.Priority,
      ConvertType: item.ConvertTypeId || this.DEFAULT_CONVERT_TYPE_ID, // NEW: Map ConvertType
      ConvertTypeId: item.ConvertTypeId,
      ConvertType2: item.ConvertType2Id || this.DEFAULT_CONVERT_TYPE_ID, // NEW: Map ConvertType2
      ConvertType2Id: item.ConvertType2Id,
      Created: item.Created ? new Date(item.Created) : undefined,
      Modified: item.Modified ? new Date(item.Modified) : undefined,
      Author: item.Author,
      Editor: item.Editor
    };
  }
}