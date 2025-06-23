// src/webparts/xyea/components/RenameFilesManagement/services/FileSearchService.ts

import { IRenameTableRow } from '../types/RenameFilesTypes';
import { ExcelFileProcessor } from './ExcelFileProcessor';

export class FileSearchService {
  private context: any;
  private excelProcessor: ExcelFileProcessor;

  constructor(context: any) {
    this.context = context;
    this.excelProcessor = new ExcelFileProcessor();
  }

  public async searchFiles(
    folderPath: string,
    rows: IRenameTableRow[],
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void
  ): Promise<{ [rowIndex: number]: 'found' | 'not-found' | 'searching' }> {
    
    console.log(`Starting file search in folder: ${folderPath}`);
    
    const results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' } = {};
    
    // Initialize all rows as searching
    rows.forEach(row => {
      results[row.rowIndex] = 'searching';
      progressCallback(row.rowIndex, 'searching');
    });
    
    // Search for each file
    for (const row of rows) {
      try {
        const found = await this.searchForFile(row, folderPath);
        const result = found ? 'found' : 'not-found';
        
        results[row.rowIndex] = result;
        progressCallback(row.rowIndex, result);
        
        // Small delay to prevent overwhelming SharePoint API
        await this.delay(100);
        
      } catch (error) {
        console.error(`Error searching for file in row ${row.rowIndex}:`, error);
        results[row.rowIndex] = 'not-found';
        progressCallback(row.rowIndex, 'not-found');
      }
    }
    
    console.log('File search completed', results);
    return results;
  }

  private async searchForFile(row: IRenameTableRow, baseFolderPath: string): Promise<boolean> {
    try {
      // Extract RelativePath from the row
      const relativePath = this.excelProcessor.extractRelativePath(row);
      
      if (!relativePath) {
        console.log(`No RelativePath found for row ${row.rowIndex}`);
        return false;
      }
      
      // Extract filename from path
      const fileName = this.excelProcessor.extractFileName(relativePath);
      
      if (!fileName) {
        console.log(`No filename extracted from path: ${relativePath}`);
        return false;
      }
      
      console.log(`Searching for filename: ${fileName} in row ${row.rowIndex}`);
      
      // Search for the file recursively in the selected folder
      const found = await this.searchFileRecursively(baseFolderPath, fileName);
      
      return found;
      
    } catch (error) {
      console.error(`Error searching for file in row ${row.rowIndex}:`, error);
      return false;
    }
  }

  private async searchFileRecursively(folderPath: string, fileName: string): Promise<boolean> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      
      // Search for files in current folder
      const filesUrl = `${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/files?$select=Name`;
      
      const filesResponse = await fetch(filesUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (filesResponse.ok) {
        const filesData = await filesResponse.json();
        const files = filesData.d?.results || filesData.value || [];
        
        // Check if file exists in current folder (case-insensitive)
        const fileFound = files.some((file: any) => 
          file.Name.toLowerCase() === fileName.toLowerCase()
        );
        
        if (fileFound) {
          console.log(`File "${fileName}" found in folder: ${folderPath}`);
          return true;
        }
      }
      
      // Search in subfolders
      const foldersUrl = `${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/folders?$select=Name,ServerRelativeUrl`;
      
      const foldersResponse = await fetch(foldersUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (foldersResponse.ok) {
        const foldersData = await foldersResponse.json();
        const folders = foldersData.d?.results || foldersData.value || [];
        
        // Filter out system folders
        const userFolders = folders.filter((folder: any) => 
          !folder.Name.startsWith('_') && 
          !folder.Name.startsWith('Forms')
        );
        
        // Search recursively in each subfolder
        for (const folder of userFolders) {
          const found = await this.searchFileRecursively(folder.ServerRelativeUrl, fileName);
          if (found) {
            return true;
          }
          
          // Small delay between folder searches
          await this.delay(50);
        }
      }
      
      return false;
      
    } catch (error) {
      console.error(`Error searching in folder ${folderPath}:`, error);
      return false;
    }
  }

  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  public async searchSingleFile(folderPath: string, fileName: string): Promise<{ found: boolean; path?: string }> {
    try {
      console.log(`Searching for single file: ${fileName} in ${folderPath}`);
      
      const found = await this.searchFileRecursively(folderPath, fileName);
      
      return {
        found,
        path: found ? folderPath : undefined
      };
      
    } catch (error) {
      console.error('Error in single file search:', error);
      return { found: false };
    }
  }

  public async getFileDetails(filePath: string): Promise<any> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      
      const response = await fetch(`${webUrl}/_api/web/getFileByServerRelativeUrl('${filePath}')`, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });

      if (response.ok) {
        const data = await response.json();
        return data.d || data;
      }

      return null;
    } catch (error) {
      console.error('Error getting file details:', error);
      return null;
    }
  }

  public extractFileNameFromPath(fullPath: string): string {
    // Handle both Windows (\) and Unix (/) path separators
    const pathParts = fullPath.split(/[\\\/]/);
    return pathParts[pathParts.length - 1] || '';
  }

  public validateFileName(fileName: string): boolean {
    // Check for invalid characters in SharePoint file names
    const invalidChars = /[<>:"/\\|?*]/;
    return !invalidChars.test(fileName) && fileName.trim().length > 0;
  }
}