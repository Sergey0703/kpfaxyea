// src/webparts/xyea/utils/PriorityHelper.ts

import { IConvertFileProps } from '../models';

export class PriorityHelper {
  
  // Получить следующий приоритет для ConvertFilesID
  public static getNextPriority(items: IConvertFileProps[], convertFilesId: number): number {
    const filteredItems = items.filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId && !item.IsDeleted);
    
    if (filteredItems.length === 0) {
      return 1;
    }

    const maxPriority = Math.max(...filteredItems.map((item: IConvertFileProps) => item.Priority));
    return maxPriority + 1;
  }

  // Отсортировать элементы по приоритету
  public static sortByPriority(items: IConvertFileProps[]): IConvertFileProps[] {
    return [...items].sort((a: IConvertFileProps, b: IConvertFileProps) => a.Priority - b.Priority);
  }

  // Переместить элемент вверх (уменьшить приоритет)
  public static moveUp(items: IConvertFileProps[], itemId: number): { 
    updatedItems: IConvertFileProps[], 
    itemsToUpdate: Array<{ id: number, priority: number }> 
  } {
    const sortedItems = this.sortByPriority(items);
    const currentIndex = sortedItems.findIndex((item: IConvertFileProps) => item.Id === itemId);
    
    if (currentIndex <= 0) {
      // Элемент уже первый или не найден
      return { updatedItems: items, itemsToUpdate: [] };
    }

    const currentItem = sortedItems[currentIndex];
    const previousItem = sortedItems[currentIndex - 1];

    // Меняем приоритеты местами
    const tempPriority = currentItem.Priority;
    currentItem.Priority = previousItem.Priority;
    previousItem.Priority = tempPriority;

    return {
      updatedItems: items,
      itemsToUpdate: [
        { id: currentItem.Id, priority: currentItem.Priority },
        { id: previousItem.Id, priority: previousItem.Priority }
      ]
    };
  }

  // Переместить элемент вниз (увеличить приоритет)
  public static moveDown(items: IConvertFileProps[], itemId: number): { 
    updatedItems: IConvertFileProps[], 
    itemsToUpdate: Array<{ id: number, priority: number }> 
  } {
    const sortedItems = this.sortByPriority(items);
    const currentIndex = sortedItems.findIndex((item: IConvertFileProps) => item.Id === itemId);
    
    if (currentIndex < 0 || currentIndex >= sortedItems.length - 1) {
      // Элемент последний или не найден
      return { updatedItems: items, itemsToUpdate: [] };
    }

    const currentItem = sortedItems[currentIndex];
    const nextItem = sortedItems[currentIndex + 1];

    // Меняем приоритеты местами
    const tempPriority = currentItem.Priority;
    currentItem.Priority = nextItem.Priority;
    nextItem.Priority = tempPriority;

    return {
      updatedItems: items,
      itemsToUpdate: [
        { id: currentItem.Id, priority: currentItem.Priority },
        { id: nextItem.Id, priority: nextItem.Priority }
      ]
    };
  }

  // Нормализовать приоритеты для ConvertFilesID (привести к последовательности 1,2,3,...)
  public static normalizePriorities(items: IConvertFileProps[], convertFilesId: number): { 
    updatedItems: IConvertFileProps[], 
    itemsToUpdate: Array<{ id: number, priority: number }> 
  } {
    const filteredItems = items
      .filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId && !item.IsDeleted)
      .sort((a: IConvertFileProps, b: IConvertFileProps) => a.Priority - b.Priority);

    const itemsToUpdate: Array<{ id: number, priority: number }> = [];
    
    filteredItems.forEach((item: IConvertFileProps, index: number) => {
      const newPriority = index + 1;
      if (item.Priority !== newPriority) {
        item.Priority = newPriority;
        itemsToUpdate.push({ id: item.Id, priority: newPriority });
      }
    });

    return {
      updatedItems: items,
      itemsToUpdate
    };
  }

  // Проверить, можно ли переместить элемент вверх
  public static canMoveUp(items: IConvertFileProps[], itemId: number, convertFilesId: number): boolean {
    const filteredItems = items
      .filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId && !item.IsDeleted)
      .sort((a: IConvertFileProps, b: IConvertFileProps) => a.Priority - b.Priority);
    
    const currentIndex = filteredItems.findIndex((item: IConvertFileProps) => item.Id === itemId);
    return currentIndex > 0;
  }

  // Проверить, можно ли переместить элемент вниз
  public static canMoveDown(items: IConvertFileProps[], itemId: number, convertFilesId: number): boolean {
    const filteredItems = items
      .filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId && !item.IsDeleted)
      .sort((a: IConvertFileProps, b: IConvertFileProps) => a.Priority - b.Priority);
    
    const currentIndex = filteredItems.findIndex((item: IConvertFileProps) => item.Id === itemId);
    return currentIndex >= 0 && currentIndex < filteredItems.length - 1;
  }
}