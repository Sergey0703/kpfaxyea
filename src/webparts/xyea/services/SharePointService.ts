// src/webparts/xyea/services/SharePointService.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export class SharePointService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  // Получить все элементы из списка
  public async getListItems<T>(listName: string, select?: string, expand?: string, filter?: string, orderBy?: string): Promise<T[]> {
    try {
      let url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      
      const queryParams: string[] = [];
      if (select) queryParams.push(`$select=${encodeURIComponent(select)}`);
      if (expand) queryParams.push(`$expand=${encodeURIComponent(expand)}`);
      if (filter) queryParams.push(`$filter=${encodeURIComponent(filter)}`);
      if (orderBy) queryParams.push(`$orderby=${encodeURIComponent(orderBy)}`);
      
      if (queryParams.length > 0) {
        url += `?${queryParams.join('&')}`;
      }

      console.log('SharePoint API URL:', url);

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error(`SharePoint API Error Response:`, errorText);
        throw new Error(`HTTP error! status: ${response.status}, details: ${errorText}`);
      }

      const data = await response.json();
      return data.value;
    } catch (error) {
      console.error(`Error getting items from list ${listName}:`, error);
      throw error;
    }
  }

  // Получить элемент по ID
  public async getListItemById<T>(listName: string, id: number, select?: string, expand?: string): Promise<T> {
    try {
      let url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;
      
      const queryParams: string[] = [];
      if (select) queryParams.push(`$select=${select}`);
      if (expand) queryParams.push(`$expand=${expand}`);
      
      if (queryParams.length > 0) {
        url += `?${queryParams.join('&')}`;
      }

      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error(`Error getting item ${id} from list ${listName}:`, error);
      throw error;
    }
  }

  // Создать новый элемент
  public async createListItem<T>(listName: string, item: any): Promise<T> {
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: JSON.stringify(item)
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      return await response.json();
    } catch (error) {
      console.error(`Error creating item in list ${listName}:`, error);
      throw error;
    }
  }

  // Обновить элемент
  public async updateListItem<T>(listName: string, id: number, item: any): Promise<T> {
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify(item)
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      // Возвращаем обновленный элемент
      return await this.getListItemById<T>(listName, id);
    } catch (error) {
      console.error(`Error updating item ${id} in list ${listName}:`, error);
      throw error;
    }
  }

  // Удалить элемент
  public async deleteListItem(listName: string, id: number): Promise<void> {
    try {
      const url = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;
      
      const response: SPHttpClientResponse = await this.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
    } catch (error) {
      console.error(`Error deleting item ${id} from list ${listName}:`, error);
      throw error;
    }
  }
}