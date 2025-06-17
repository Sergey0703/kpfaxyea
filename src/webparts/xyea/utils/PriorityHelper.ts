// src/webparts/xyea/utils/PriorityHelper.ts

import { IConvertFileProps } from '../models';

export class PriorityHelper {
  
  // Получить следующий приоритет для ConvertFilesID - FIXED to count ALL items (including deleted)
  public static getNextPriority(items: IConvertFileProps[], convertFilesId: number): number {
    // Filter by ConvertFilesID but include ALL items (both deleted and active)
    const filteredItems = items.filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId);
    
    if (filteredItems.length === 0) {
      return 1;
    }

    // Find the maximum priority among ALL items (including deleted ones)
    const maxPriority = Math.max(...filteredItems.map((item: IConvertFileProps) => item.Priority));
    return maxPriority + 1;
  }

  // Отсортировать элементы по приоритету
  public static sortByPriority(items: IConvertFileProps[]): IConvertFileProps[] {
    return [...items].sort((a: IConvertFileProps, b: IConvertFileProps) => a.Priority - b.Priority);
  }

  // Переместить элемент вверх (уменьшить приоритет) - работает со ВСЕМИ элементами
  public static moveUp(items: IConvertFileProps[], itemId: number): { 
    updatedItems: IConvertFileProps[], 
    itemsToUpdate: Array<{ id: number, priority: number }> 
  } {
    const sortedItems = this.sortByPriority(items);
    let currentIndex = -1;
    
    for (let i = 0; i < sortedItems.length; i++) {
      if (sortedItems[i].Id === itemId) {
        currentIndex = i;
        break;
      }
    }
    
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

  // Переместить элемент вниз (увеличить приоритет) - работает со ВСЕМИ элементами
  public static moveDown(items: IConvertFileProps[], itemId: number): { 
    updatedItems: IConvertFileProps[], 
    itemsToUpdate: Array<{ id: number, priority: number }> 
  } {
    const sortedItems = this.sortByPriority(items);
    let currentIndex = -1;
    
    for (let i = 0; i < sortedItems.length; i++) {
      if (sortedItems[i].Id === itemId) {
        currentIndex = i;
        break;
      }
    }
    
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
  // This method works only with NON-DELETED items to create clean sequence
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

  // Проверить, можно ли переместить элемент вверх - работает со ВСЕМИ элементами (включая удаленные)
  public static canMoveUp(items: IConvertFileProps[], itemId: number, convertFilesId: number): boolean {
    const filteredItems = items
      .filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId)
      .sort((a: IConvertFileProps, b: IConvertFileProps) => a.Priority - b.Priority);
    
    let currentIndex = -1;
    for (let i = 0; i < filteredItems.length; i++) {
      if (filteredItems[i].Id === itemId) {
        currentIndex = i;
        break;
      }
    }
    return currentIndex > 0;
  }

  // Проверить, можно ли переместить элемент вниз - работает со ВСЕМИ элементами (включая удаленные)
  public static canMoveDown(items: IConvertFileProps[], itemId: number, convertFilesId: number): boolean {
    const filteredItems = items
      .filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId)
      .sort((a: IConvertFileProps, b: IConvertFileProps) => a.Priority - b.Priority);
    
    let currentIndex = -1;
    for (let i = 0; i < filteredItems.length; i++) {
      if (filteredItems[i].Id === itemId) {
        currentIndex = i;
        break;
      }
    }
    return currentIndex >= 0 && currentIndex < filteredItems.length - 1;
  }

  // New method: Get all used priorities for a ConvertFilesID (including deleted items)
  public static getUsedPriorities(items: IConvertFileProps[], convertFilesId: number): number[] {
    return items
      .filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId)
      .map((item: IConvertFileProps) => item.Priority)
      .sort((a, b) => a - b);
  }

  // New method: Get next available priority slot (useful for filling gaps)
  public static getNextAvailablePriority(items: IConvertFileProps[], convertFilesId: number): number {
    const usedPriorities = this.getUsedPriorities(items, convertFilesId);
    
    if (usedPriorities.length === 0) {
      return 1;
    }

    // Look for gaps in sequence
    for (let i = 1; i <= usedPriorities.length + 1; i++) {
      if (!usedPriorities.includes(i)) {
        return i;
      }
    }

    // If no gaps found, return next number
    return Math.max(...usedPriorities) + 1;
  }

  // New method: Validate priority uniqueness
  public static validatePriorityUniqueness(items: IConvertFileProps[], convertFilesId: number): {
    isValid: boolean;
    duplicates: Array<{ priority: number; itemIds: number[] }>;
  } {
    const filteredItems = items.filter((item: IConvertFileProps) => item.ConvertFilesID === convertFilesId);
    const priorityMap = new Map<number, number[]>();

    // Group items by priority
    filteredItems.forEach(item => {
      if (!priorityMap.has(item.Priority)) {
        priorityMap.set(item.Priority, []);
      }
      const existingItems = priorityMap.get(item.Priority);
      if (existingItems) {
        existingItems.push(item.Id);
      }
    });

    // Find duplicates
    const duplicates: Array<{ priority: number; itemIds: number[] }> = [];
    priorityMap.forEach((itemIds, priority) => {
      if (itemIds.length > 1) {
        duplicates.push({ priority, itemIds });
      }
    });

    return {
      isValid: duplicates.length === 0,
      duplicates
    };
  }
}