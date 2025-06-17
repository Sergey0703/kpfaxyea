// src/webparts/xyea/utils/ValidationHelper.ts

export class ValidationHelper {
  
  // Proверить, что строка не пустая и не состоит только из пробелов
  public static isNotEmpty(value: string | undefined): boolean {
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
  public static isValidId(id: number | undefined): boolean {
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

  // Updated validation for ConvertFileProps - only title is required
  public static validateConvertFileProps(
    title: string,
    convertFilesId: number,
    prop?: string,
    prop2?: string,
    priority?: number
  ): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];

    // Title is required
    if (!this.isNotEmpty(title)) {
      errors.push('Title is required');
    } else if (!this.isValidLength(title, 1, 255)) {
      errors.push('Title must be between 1 and 255 characters');
    }

    // ConvertFilesID is required
    if (!this.isValidId(convertFilesId)) {
      errors.push('ConvertFilesID is required and must be a positive number');
    }

    // Prop is optional, but if provided, validate length
    if (prop !== undefined && prop.trim().length > 0 && !this.isValidLength(prop, 1, 255)) {
      errors.push('Prop must be between 1 and 255 characters when provided');
    }

    // Prop2 is optional, but if provided, validate length
    if (prop2 !== undefined && prop2.trim().length > 0 && !this.isValidLength(prop2, 1, 255)) {
      errors.push('Prop2 must be between 1 and 255 characters when provided');
    }

    // Priority is optional, but if provided, must be positive
    if (priority !== undefined && !this.isPositiveNumber(priority)) {
      errors.push('Priority must be a positive number when provided');
    }

    return {
      isValid: errors.length === 0,
      errors
    };
  }

  // Clean string from extra spaces
  public static sanitizeString(value: string | undefined): string {
    if (!value) {
      return '';
    }
    return value.trim();
  }

  // Sanitize optional string - returns empty string if value is undefined or only whitespace
  public static sanitizeOptionalString(value: string | undefined): string {
    if (!value || value.trim().length === 0) {
      return '';
    }
    return value.trim();
  }

  // Check if string contains only valid characters
  public static containsOnlyValidCharacters(value: string, allowedPattern?: RegExp): boolean {
    if (!this.isNotEmpty(value)) {
      return false;
    }

    // By default allow letters, numbers, spaces and basic punctuation
    const defaultPattern = /^[a-zA-Z0-9\s\-_.,!?()]+$/;
    const pattern = allowedPattern || defaultPattern;
    
    return pattern.test(value);
  }

  // Validate and prepare ConvertFileProps data for saving
  public static prepareConvertFilePropsData(
    title: string,
    convertFilesId: number,
    prop?: string,
    prop2?: string
  ): { 
    isValid: boolean; 
    errors: string[]; 
    data?: { 
      title: string; 
      convertFilesId: number; 
      prop: string; 
      prop2: string; 
    } 
  } {
    const validation = this.validateConvertFileProps(title, convertFilesId, prop, prop2);
    
    if (!validation.isValid) {
      return {
        isValid: false,
        errors: validation.errors
      };
    }

    return {
      isValid: true,
      errors: [],
      data: {
        title: this.sanitizeString(title),
        convertFilesId,
        prop: this.sanitizeOptionalString(prop),
        prop2: this.sanitizeOptionalString(prop2)
      }
    };
  }
}