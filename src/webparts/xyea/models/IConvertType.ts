// src/webparts/xyea/models/IConvertType.ts

export interface IConvertType {
  Id: number;
  Title: string;
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