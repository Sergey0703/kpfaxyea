// src/webparts/xyea/services/ConvertFilesService.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointService } from './SharePointService';
import { IConvertFile } from '../models';

export class ConvertFilesService {
  private spService: SharePointService;
  private readonly LIST_NAME = 'convertfiles';

  constructor(context: WebPartContext) {
    this.spService = new SharePointService(context);
  }

  // Получить все файлы
  public async getAllConvertFiles(): Promise<IConvertFile[]> {
    try {
      return await this.spService.getListItems<IConvertFile>(
        this.LIST_NAME,
        'Id,Title,Created,Modified',
        undefined,
        undefined,
        'Title asc'
      );
    } catch (error) {
      console.error('Error getting convert files:', error);
      throw error;
    }
  }

  // Получить файл по ID
  public async getConvertFileById(id: number): Promise<IConvertFile> {
    try {
      return await this.spService.getListItemById<IConvertFile>(
        this.LIST_NAME,
        id,
        'Id,Title,Created,Modified'
      );
    } catch (error) {
      console.error(`Error getting convert file ${id}:`, error);
      throw error;
    }
  }

  // Создать новый файл
  public async createConvertFile(title: string): Promise<IConvertFile> {
    try {
      const newItem = {
        Title: title
      };

      return await this.spService.createListItem<IConvertFile>(this.LIST_NAME, newItem);
    } catch (error) {
      console.error('Error creating convert file:', error);
      throw error;
    }
  }

  // Обновить файл
  public async updateConvertFile(id: number, title: string): Promise<IConvertFile> {
    try {
      const updateItem = {
        Title: title
      };

      return await this.spService.updateListItem<IConvertFile>(this.LIST_NAME, id, updateItem);
    } catch (error) {
      console.error(`Error updating convert file ${id}:`, error);
      throw error;
    }
  }

  // Удалить файл
  public async deleteConvertFile(id: number): Promise<void> {
    try {
      await this.spService.deleteListItem(this.LIST_NAME, id);
    } catch (error) {
      console.error(`Error deleting convert file ${id}:`, error);
      throw error;
    }
  }

  // Проверить существование файла по заголовку
  public async checkTitleExists(title: string, excludeId?: number): Promise<boolean> {
    try {
      const filter = excludeId 
        ? `Title eq '${title}' and Id ne ${excludeId}`
        : `Title eq '${title}'`;

      const items = await this.spService.getListItems<IConvertFile>(
        this.LIST_NAME,
        'Id',
        undefined,
        filter
      );

      return items.length > 0;
    } catch (error) {
      console.error('Error checking title exists:', error);
      return false;
    }
  }
}