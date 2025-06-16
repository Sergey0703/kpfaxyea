// src/webparts/xyea/models/IConvertFileProps.ts

export interface IConvertFileProps {
  Id: number;
  Title: string;
  ConvertFilesID: number; // Для внутреннего использования
  ConvertFilesIDId?: number; // Поле из SharePoint API для Lookup
  Prop: string;
  Prop2: string;
  IsDeleted: boolean;
  Priority: number; // Исправлено название поля (было Prioruty)
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