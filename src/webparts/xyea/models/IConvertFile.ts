// src/webparts/xyea/models/IConvertFile.ts

export interface IConvertFile {
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