// src/webparts/xyea/components/RenameFilesManagement/services/FileSearchService.ts

import { IRenameTableRow } from '../types/RenameFilesTypes';
import { SharePointFolderService } from './SharePointFolderService';
import { ExcelFileProcessor } from './ExcelFileProcessor';

export interface ISearchPlan {
  totalRows: number;
  validRows: number;
  invalidRows: number;
  directoryGroups: Array<{
    directoryPath: string;
    fullSharePointPath: string;
    exists: boolean;
    fileCount: number;
    rowIndexes: number[];
  }>;
}

export class FileSearchService {
  private context: any;
  private folderService: SharePointFolderService;
  private excelProcessor: ExcelFileProcessor;
  private isCancelled: boolean = false;
  private currentSearchId: string | null = null;

  constructor(context: any) {
    this.context = context;
    this.folderService = new SharePointFolderService(context);
    this.excelProcessor = new ExcelFileProcessor();
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
    
    console.log(`[FileSearchService] Starting optimized file search in folder: ${folderPath} (Search ID: ${searchId})`);
    
    const results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' } = {};
    
    try {
      // Step 1: Load all subfolders if not already loaded
      console.log(`[FileSearchService] Step 1: Loading all subfolders...`);
      if (statusCallback) {
        statusCallback(0, rows.length, 'Loading folder structure...');
      }
      
      await this.folderService.loadAllSubfolders(folderPath, (currentPath, foldersLoaded) => {
        if (statusCallback) {
          statusCallback(0, rows.length, `Loading folders... (${foldersLoaded} loaded)`);
        }
      });

      // Step 2: Create search plan - analyze directory structure
      console.log(`[FileSearchService] Step 2: Analyzing directory structure...`);
      if (statusCallback) {
        statusCallback(0, rows.length, 'Analyzing file paths...');
      }
      
      const searchPlan = this.createSearchPlan(rows, folderPath);
      this.logSearchPlan(searchPlan);

      // Step 3: Initialize all rows as searching
      rows.forEach(row => {
        results[row.rowIndex] = 'searching';
        progressCallback(row.rowIndex, 'searching');
      });

      // Step 4: Process by directory groups for efficiency
      console.log(`[FileSearchService] Step 3: Processing ${searchPlan.directoryGroups.length} directory groups...`);
      
      let processedRows = 0;
      
      for (const dirGroup of searchPlan.directoryGroups) {
        if (this.isCancelled || this.currentSearchId !== searchId) {
          console.log('[FileSearchService] Search cancelled during directory processing');
          break;
        }

        console.log(`[FileSearchService] Processing directory: ${dirGroup.directoryPath} (${dirGroup.fileCount} files)`);
        
        if (!dirGroup.exists) {
          // Skip this entire directory - mark all files as not found
          console.log(`[FileSearchService] Directory doesn't exist, skipping ${dirGroup.fileCount} files`);
          
          for (const rowIndex of dirGroup.rowIndexes) {
            if (!this.isCancelled && this.currentSearchId === searchId) {
              results[rowIndex] = 'not-found';
              progressCallback(rowIndex, 'not-found');
              processedRows++;
              
              if (statusCallback) {
                statusCallback(processedRows, rows.length, `Skipped: Directory not found`);
              }
            }
          }
        } else {
          // Directory exists - search for files in this specific directory
          await this.searchInSpecificDirectory(
            dirGroup,
            rows,
            results,
            progressCallback,
            (current, total, fileName) => {
              processedRows++;
              if (statusCallback) {
                statusCallback(processedRows, rows.length, fileName);
              }
            }
          );
        }

        // Small delay between directories
        await this.delay(25);
      }

      // Step 5: Handle any remaining rows that weren't in valid directories
      const processedRowIndexes = new Set<number>();
      searchPlan.directoryGroups.forEach(g => {
        g.rowIndexes.forEach(index => processedRowIndexes.add(index));
      });
      const unprocessedRows = rows.filter(row => !processedRowIndexes.has(row.rowIndex));
      
      if (unprocessedRows.length > 0) {
        console.log(`[FileSearchService] Processing ${unprocessedRows.length} rows without valid directories`);
        
        for (const row of unprocessedRows) {
          if (this.isCancelled || this.currentSearchId !== searchId) break;
          
          results[row.rowIndex] = 'not-found';
          progressCallback(row.rowIndex, 'not-found');
          processedRows++;
          
          if (statusCallback) {
            statusCallback(processedRows, rows.length, 'Invalid path structure');
          }
        }
      }

    } catch (error) {
      console.error('[FileSearchService] Error during optimized search:', error);
      
      // Mark all unprocessed rows as not found
      rows.forEach(row => {
        if (results[row.rowIndex] === 'searching') {
          results[row.rowIndex] = 'not-found';
          progressCallback(row.rowIndex, 'not-found');
        }
      });
    }
    
    if (this.isCancelled || this.currentSearchId !== searchId) {
      console.log('[FileSearchService] Search was cancelled');
    } else {
      console.log('[FileSearchService] Optimized file search completed', results);
    }
    
    return results;
  }

  private createSearchPlan(rows: IRenameTableRow[], baseFolderPath: string): ISearchPlan {
    console.log(`[FileSearchService] Creating search plan for ${rows.length} rows`);
    
    // Validate directory structure for all rows
    const validationResults = rows.map(row => {
      const fileName = String(row.cells['custom_0']?.value || '');
      const directoryPath = this.excelProcessor.extractDirectoryPathFromRow(row);
      const hasValidPath = directoryPath !== '' && fileName !== '';
      
      return {
        rowIndex: row.rowIndex,
        fileName,
        directoryPath,
        hasValidPath
      };
    });

    // Group by directory
    const directoryMap = new Map<string, number[]>();
    
    validationResults.forEach(result => {
      if (result.hasValidPath && result.directoryPath) {
        const existing = directoryMap.get(result.directoryPath);
        if (existing) {
          existing.push(result.rowIndex);
        } else {
          directoryMap.set(result.directoryPath, [result.rowIndex]);
        }
      }
    });

    // Create directory groups with existence check
    const directoryGroups = Array.from(directoryMap.entries()).map(([directoryPath, rowIndexes]) => {
      const fullSharePointPath = this.folderService.getFullDirectoryPath(directoryPath, baseFolderPath);
      const exists = this.folderService.checkDirectoryExists(fullSharePointPath);
      
      return {
        directoryPath,
        fullSharePointPath,
        exists,
        fileCount: rowIndexes.length,
        rowIndexes
      };
    });

    // Sort by file count descending (process directories with more files first)
    directoryGroups.sort((a, b) => b.fileCount - a.fileCount);

    const validRows = validationResults.filter(r => r.hasValidPath).length;
    const invalidRows = rows.length - validRows;

    return {
      totalRows: rows.length,
      validRows,
      invalidRows,
      directoryGroups
    };
  }

  private logSearchPlan(plan: ISearchPlan): void {
    console.log(`[FileSearchService] Search Plan Summary:`);
    console.log(`  Total rows: ${plan.totalRows}`);
    console.log(`  Valid rows: ${plan.validRows}`);
    console.log(`  Invalid rows: ${plan.invalidRows}`);
    console.log(`  Directory groups: ${plan.directoryGroups.length}`);
    
    console.log(`[FileSearchService] Directory Groups:`);
    plan.directoryGroups.forEach((group, index) => {
      console.log(`  ${index + 1}. "${group.directoryPath}" -> ${group.exists ? 'EXISTS' : 'NOT FOUND'} (${group.fileCount} files)`);
    });

    const existingDirs = plan.directoryGroups.filter(g => g.exists).length;
    const missingDirs = plan.directoryGroups.length - existingDirs;
    console.log(`[FileSearchService] Directories: ${existingDirs} exist, ${missingDirs} missing`);
  }

  private async searchInSpecificDirectory(
    dirGroup: { directoryPath: string; fullSharePointPath: string; fileCount: number; rowIndexes: number[] },
    allRows: IRenameTableRow[],
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    statusCallback?: (current: number, total: number, fileName: string) => void
  ): Promise<void> {
    
    console.log(`[FileSearchService] Searching in directory: ${dirGroup.fullSharePointPath}`);
    
    try {
      // Get all files in this directory
      const folderContents = await this.folderService.getFolderContents(dirGroup.fullSharePointPath);
      const files = folderContents.files;
      
      console.log(`[FileSearchService] Found ${files.length} files in directory "${dirGroup.directoryPath}"`);
      
      // Create a map of filenames for fast lookup (case-insensitive)
      const fileMap = new Map<string, any>();
      files.forEach(file => {
        fileMap.set(file.Name.toLowerCase(), file);
      });
      
      // Check each file in this directory group
      for (const rowIndex of dirGroup.rowIndexes) {
        if (this.isCancelled) break;
        
        const row = allRows.find(r => r.rowIndex === rowIndex);
        if (!row) continue;
        
        const fileName = String(row.cells['custom_0']?.value || '');
        
        if (statusCallback) {
          statusCallback(0, 0, fileName); // Current/total will be updated by caller
        }
        
        // Check if file exists (case-insensitive)
        const fileExists = fileMap.has(fileName.toLowerCase());
        
        const result = fileExists ? 'found' : 'not-found';
        results[rowIndex] = result;
        progressCallback(rowIndex, result);
        
        console.log(`[FileSearchService] File "${fileName}" in "${dirGroup.directoryPath}": ${result.toUpperCase()}`);
        
        // Small delay between files
        await this.delay(10);
      }
      
    } catch (error) {
      console.error(`[FileSearchService] Error searching in directory ${dirGroup.fullSharePointPath}:`, error);
      
      // Mark all files in this directory as not found
      for (const rowIndex of dirGroup.rowIndexes) {
        if (!this.isCancelled) {
          results[rowIndex] = 'not-found';
          progressCallback(rowIndex, 'not-found');
        }
      }
    }
  }

  public cancelSearch(): void {
    console.log('[FileSearchService] Cancelling file search...');
    this.isCancelled = true;
    this.currentSearchId = null;
  }

  public isSearchActive(): boolean {
    return this.currentSearchId !== null && !this.isCancelled;
  }

  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  public async searchSingleFile(folderPath: string, fileName: string): Promise<{ found: boolean; path?: string }> {
    try {
      console.log(`[FileSearchService] Searching for single file: ${fileName} in ${folderPath}`);
      
      const folderContents = await this.folderService.getFolderContents(folderPath);
      const files = folderContents.files;
      
      const fileFound = files.some((file: any) => 
        file.Name.toLowerCase() === fileName.toLowerCase()
      );
      
      return {
        found: fileFound,
        path: fileFound ? folderPath : undefined
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

  public getFileNameFromRow(row: IRenameTableRow): string {
    // Get filename directly from the first column (custom_0)
    const fileName = String(row.cells['custom_0']?.value || '');
    
    console.log(`[FileSearchService] getFileNameFromRow for row ${row.rowIndex}: "${fileName}"`);
    
    return fileName;
  }
}