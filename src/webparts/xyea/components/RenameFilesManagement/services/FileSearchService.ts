// src/webparts/xyea/components/RenameFilesManagement/services/FileSearchService.ts

import { IRenameTableRow } from '../types/RenameFilesTypes';

export class FileSearchService {
  private context: any;
  private isCancelled: boolean = false;
  private currentSearchId: string | null = null;

  constructor(context: any) {
    this.context = context;
  }

  public async searchFiles(
    folderPath: string,
    rows: IRenameTableRow[],
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    statusCallback?: (currentRow: number, totalRows: number, fileName: string) => void
  ): Promise<{ [rowIndex: number]: 'found' | 'not-found' | 'searching' }> {
    
    // Generate unique search ID
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`Starting file search in folder: ${folderPath} (Search ID: ${searchId})`);
    
    const results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' } = {};
    
    // Initialize all rows as searching
    rows.forEach(row => {
      results[row.rowIndex] = 'searching';
      progressCallback(row.rowIndex, 'searching');
    });
    
    // Search for each file with progress tracking
    for (let i = 0; i < rows.length; i++) {
      // Check if search was cancelled
      if (this.isCancelled || this.currentSearchId !== searchId) {
        console.log('Search cancelled');
        break;
      }
      
      const row = rows[i];
      
      try {
        // Get filename directly from the first column (custom_0) which was populated during Excel processing
        const fileName = String(row.cells['custom_0']?.value || '');
        
        console.log(`[FileSearchService] Row ${row.rowIndex + 1}: Using filename from first column: "${fileName}"`);
        
        // Update status
        if (statusCallback) {
          statusCallback(i + 1, rows.length, fileName || 'Unknown file');
        }
        
        const found = await this.searchForFile(row, folderPath);
        
        // Check again if cancelled during search
        if (this.isCancelled || this.currentSearchId !== searchId) {
          console.log('Search cancelled during file search');
          break;
        }
        
        const result = found ? 'found' : 'not-found';
        results[row.rowIndex] = result;
        progressCallback(row.rowIndex, result);
        
        // Shorter delay for better performance
        await this.delay(50);
        
      } catch (error) {
        console.error(`Error searching for file in row ${row.rowIndex}:`, error);
        
        // Don't fail the entire search for one error
        if (!this.isCancelled && this.currentSearchId === searchId) {
          results[row.rowIndex] = 'not-found';
          progressCallback(row.rowIndex, 'not-found');
        }
      }
    }
    
    if (this.isCancelled || this.currentSearchId !== searchId) {
      console.log('Search was cancelled');
    } else {
      console.log('File search completed', results);
    }
    
    return results;
  }

  public cancelSearch(): void {
    console.log('Cancelling file search...');
    this.isCancelled = true;
    this.currentSearchId = null;
  }

  public isSearchActive(): boolean {
    return this.currentSearchId !== null && !this.isCancelled;
  }

  private async searchForFile(row: IRenameTableRow, baseFolderPath: string): Promise<boolean> {
    // Check cancellation before starting
    if (this.isCancelled) return false;
    
    try {
      // Get filename directly from the first column (custom_0)
      const fileName = String(row.cells['custom_0']?.value || '');
      
      if (!fileName) {
        console.log(`[FileSearchService] No filename found in first column for row ${row.rowIndex}`);
        return false;
      }
      
      console.log(`[FileSearchService] Searching for filename: "${fileName}" in row ${row.rowIndex}`);
      
      // Search for the file recursively in the selected folder
      const found = await this.searchFileRecursively(baseFolderPath, fileName);
      
      return found;
      
    } catch (error) {
      console.error(`Error searching for file in row ${row.rowIndex}:`, error);
      return false;
    }
  }

  private async searchFileRecursively(folderPath: string, fileName: string): Promise<boolean> {
    // Check cancellation before each API call
    if (this.isCancelled) return false;
    
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      
      // Search for files in current folder
      const filesUrl = `${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/files?$select=Name&$top=1000`;
      
      const filesResponse = await fetch(filesUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (this.isCancelled) return false;
      
      if (filesResponse.ok) {
        const filesData = await filesResponse.json();
        const files = filesData.d?.results || filesData.value || [];
        
        // Check if file exists in current folder (case-insensitive)
        const fileFound = files.some((file: any) => 
          file.Name.toLowerCase() === fileName.toLowerCase()
        );
        
        if (fileFound) {
          console.log(`[FileSearchService] File "${fileName}" found in folder: ${folderPath}`);
          return true;
        }
      }
      
      // Only search subfolders if not cancelled and file not found
      if (!this.isCancelled) {
        // Search in subfolders (limit depth to prevent infinite recursion)
        const foldersUrl = `${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/folders?$select=Name,ServerRelativeUrl&$top=50`;
        
        const foldersResponse = await fetch(foldersUrl, {
          method: 'GET',
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        });
        
        if (this.isCancelled) return false;
        
        if (foldersResponse.ok) {
          const foldersData = await foldersResponse.json();
          const folders = foldersData.d?.results || foldersData.value || [];
          
          // Filter out system folders
          const userFolders = folders.filter((folder: any) => 
            !folder.Name.startsWith('_') && 
            !folder.Name.startsWith('Forms')
          );
          
          // Search recursively in each subfolder (but check cancellation frequently)
          for (const folder of userFolders) {
            if (this.isCancelled) return false;
            
            const found = await this.searchFileRecursively(folder.ServerRelativeUrl, fileName);
            if (found) {
              return true;
            }
            
            // Very small delay between folders
            await this.delay(10);
          }
        }
      }
      
      return false;
      
    } catch (error) {
      console.error(`[FileSearchService] Error searching in folder ${folderPath}:`, error);
      return false;
    }
  }

  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  public async searchSingleFile(folderPath: string, fileName: string): Promise<{ found: boolean; path?: string }> {
    try {
      console.log(`[FileSearchService] Searching for single file: ${fileName} in ${folderPath}`);
      
      const found = await this.searchFileRecursively(folderPath, fileName);
      
      return {
        found,
        path: found ? folderPath : undefined
      };
      
    } catch (error) {
      console.error('[FileSearchService] Error in single file search:', error);
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
      console.error('[FileSearchService] Error getting file details:', error);
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

  // Legacy method for backward compatibility - now uses filename from first column
  public getFileNameFromRow(row: IRenameTableRow): string {
    // Get filename directly from the first column (custom_0)
    const fileName = String(row.cells['custom_0']?.value || '');
    
    console.log(`[FileSearchService] getFileNameFromRow for row ${row.rowIndex}: "${fileName}"`);
    
    return fileName;
  }
}