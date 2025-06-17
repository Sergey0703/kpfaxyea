// src/webparts/xyea/utils/ValidationHelper.ts

export class ValidationHelper {
  
  // Проверить, что строка не пустая и не состоит только из пробелов
  public static isNotEmpty(value: string | undefined): boolean { // Changed from null to undefined
    return value !== undefined && value.trim().length > 0;
  }

  // Проверить длину строки
  public static isValidLength(value: string, minLength: number = 1, maxLength: number = 255): boolean {
    if (!this.isNotEmpty(value)) {
      return false;
    }
    const trimmedValue = value.trim();
    return trimmedValue.length >= minLength && trimmedValue.length <= maxLength;
  }

  // Проверить, что число положительное
  public static isPositiveNumber(value: number): boolean {
    return typeof value === 'number' && value > 0 && !isNaN(value);
  }

  // Проверить валидность ID
  public static isValidId(id: number | undefined): boolean { // Changed from null to undefined
    return id !== undefined && this.isPositiveNumber(id);
  }

  // Валидация ConvertFile
  public static validateConvertFile(title: string): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];

    if (!this.isNotEmpty(title)) {
      errors.push('Title is required');
    } else if (!this.isValidLength(title, 1, 255)) {
      errors.push('Title must be between 1 and 255 characters');
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  }

  // Валидация ConvertFileProps
  public static validateConvertFileProps(
    title: string,
    convertFilesId: number,
    prop: string,
    prop2: string,
    priority: number
  ): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];

    if (!this.isNotEmpty(title)) {
      errors.push('Title is required');
    } else if (!this.isValidLength(title, 1, 255)) {
      errors.push('Title must be between 1 and 255 characters');
    }

    if (!this.isValidId(convertFilesId)) {
      errors.push('ConvertFilesID is required and must be a positive number');
    }

    if (!this.isNotEmpty(prop)) {
      errors.push('Prop is required');
    } else if (!this.isValidLength(prop, 1, 255)) {
      errors.push('Prop must be between 1 and 255 characters');
    }

    if (!this.isNotEmpty(prop2)) {
      errors.push('Prop2 is required');
    } else if (!this.isValidLength(prop2, 1, 255)) {
      errors.push('Prop2 must be between 1 and 255 characters');
    }

    if (!this.isPositiveNumber(priority)) {
      errors.push('Priority must be a positive number');
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  }

  // Очистить строку от лишних пробелов
  public static sanitizeString(value: string | undefined): string { // Changed from null to undefined
    if (!value) {
      return '';
    }
    return value.trim();
  }

  // Проверить, содержит ли строка только допустимые символы
  public static containsOnlyValidCharacters(value: string, allowedPattern?: RegExp): boolean {
    if (!this.isNotEmpty(value)) {
      return false;
    }

    // По умолчанию разрешаем буквы, цифры, пробелы и основные знаки препинания
    const defaultPattern = /^[a-zA-Z0-9\s\-_.,!?()]+$/;
    const pattern = allowedPattern || defaultPattern;
    
    return pattern.test(value);
  }
}