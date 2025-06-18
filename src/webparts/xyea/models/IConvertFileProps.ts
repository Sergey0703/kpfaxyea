// src/webparts/xyea/models/IConvertFileProps.ts - Updated with ConvertType fields

export interface IConvertFileProps {
  Id: number;
  Title: string;
  ConvertFilesID: number; // Для внутреннего использования
  ConvertFilesIDId?: number; // Поле из SharePoint API для Lookup
  Prop: string;
  Prop2: string;
  IsDeleted: boolean;
  Priority: number;
  
  // NEW: ConvertType Lookup fields
  ConvertType: number; // ID типа величины для Prop
  ConvertTypeId?: number; // Поле из SharePoint API для Lookup
  ConvertType2: number; // ID типа величины для Prop2
  ConvertType2Id?: number; // Поле из SharePoint API для Lookup
  
  // Дополнительные системные поля SharePoint
  Created?: Date;
  Modified?: Date;
  Author?: {
    Title: string;
    Email: string;
  };
  Editor?: {
    Title: string;
    Email: string;
  };
}