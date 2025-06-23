// src/webparts/xyea/components/RenameFilesManagement/services/FileSearchService.ts

import { 
  IRenameTableRow, 
  SearchStage, 
  ISearchProgress, 
  IDirectoryAnalysis, 
  ISearchPlan,
  SearchProgressHelper 
} from '../types/RenameFilesTypes';
import { SharePointFolderService } from './SharePointFolderService';
import { ExcelFileProcessor } from './ExcelFileProcessor';

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

  /**
   * NEW: Three-stage file search with detailed progress tracking
   */
  public async searchFiles(
    folderPath: string,
    rows: IRenameTableRow[],
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<{ [rowIndex: number]: 'found' | 'not-found' | 'searching' }> {
    
    // Generate unique search ID
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] Starting three-stage file search (Search ID: ${searchId})`);
    
    const results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' } = {};
    let currentProgress = SearchProgressHelper.createInitialProgress();
    
    try {
      // Initialize all rows as searching
      rows.forEach(row => {
        results[row.rowIndex] = 'searching';
        progressCallback(row.rowIndex, 'searching');
      });

      // STAGE 1: ANALYZING DIRECTORIES (0-25%)
      currentProgress = await this.executeStage1_AnalyzeDirectories(
        rows, 
        folderPath, 
        currentProgress, 
        statusCallback
      );
      
      if (this.isCancelled || this.currentSearchId !== searchId) {
        return this.handleCancellation(results, rows, progressCallback);
      }

      // STAGE 2: CHECKING DIRECTORY EXISTENCE (25-50%)
      currentProgress = await this.executeStage2_CheckDirectoryExistence(
        currentProgress,
        statusCallback
      );
      
      if (this.isCancelled || this.currentSearchId !== searchId) {
        return this.handleCancellation(results, rows, progressCallback);
      }

      // STAGE 3: SEARCHING FILES (50-100%)
      await this.executeStage3_SearchFiles(
        currentProgress,
        rows,
        results,
        progressCallback,
        statusCallback
      );

      // Mark completion
      if (!this.isCancelled && this.currentSearchId === searchId) {
        const finalProgress = SearchProgressHelper.transitionToStage(
          currentProgress,
          SearchStage.COMPLETED,
          {
            currentFileName: 'Search completed successfully',
            overallProgress: 100
          }
        );
        statusCallback?.(finalProgress);
      }

    } catch (error) {
      console.error('[FileSearchService] Error during search:', error);
      
      // Mark all unprocessed rows as not found
      rows.forEach(row => {
        if (results[row.rowIndex] === 'searching') {
          results[row.rowIndex] = 'not-found';
          progressCallback(row.rowIndex, 'not-found');
        }
      });

      // Report error
      const errorProgress = SearchProgressHelper.transitionToStage(
        currentProgress,
        SearchStage.ERROR,
        {
          currentFileName: 'Search failed',
          errors: [error instanceof Error ? error.message : 'Unknown error']
        }
      );
      statusCallback?.(errorProgress);
    }
    
    console.log('[FileSearchService] Search completed:', results);
    return results;
  }

  /**
   * STAGE 1: Analyze directories and create search plan (0-25%) - OPTIMIZED
   */
  private async executeStage1_AnalyzeDirectories(
    rows: IRenameTableRow[],
    baseFolderPath: string,
    currentProgress: ISearchProgress,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<ISearchProgress> {
    
    console.log('[FileSearchService] STAGE 1: Analyzing directories (OPTIMIZED)...');
    
    // Transition to stage 1
    let progress = SearchProgressHelper.transitionToStage(
      currentProgress,
      SearchStage.ANALYZING_DIRECTORIES,
      {
        totalRows: rows.length,
        currentFileName: 'Extracting unique directories...'
      }
    );
    statusCallback?.(progress);

    // OPTIMIZED: Fast extraction of unique directories from Directory column
    const uniqueDirectories = new Set<string>();
    const directoryToRows = new Map<string, number[]>();
    let validRows = 0;

    console.log('[FileSearchService] Fast extraction from Directory column...');

    // Quick pass without progress updates - extract from Directory column directly
    rows.forEach(row => {
      // Get directory directly from Directory column (custom_1)
      const directoryCell = row.cells['custom_1'];
      let directoryPath = '';
      
      if (directoryCell && directoryCell.value) {
        directoryPath = String(directoryCell.value).trim();
      } else {
        // Fallback: extract from RelativePath if Directory column is empty
        directoryPath = this.excelProcessor.extractDirectoryPathFromRow(row);
      }
      
      if (directoryPath) {
        uniqueDirectories.add(directoryPath);
        
        if (!directoryToRows.has(directoryPath)) {
          directoryToRows.set(directoryPath, []);
        }
        directoryToRows.get(directoryPath)!.push(row.rowIndex);
        validRows++;
      }
    });

    console.log('[FileSearchService] Unique directories extracted:', {
      totalRows: rows.length,
      uniqueDirectories: uniqueDirectories.size,
      validRows
    });

    // Update progress for unique directory processing
    progress = SearchProgressHelper.updateStageProgress(
      progress,
      50,
      {
        currentFileName: `Found ${uniqueDirectories.size} unique directories from ${rows.length} rows`,
        directoriesAnalyzed: uniqueDirectories.size,
        totalDirectories: uniqueDirectories.size
      }
    );
    statusCallback?.(progress);

    // Create directory analysis results - UPDATE AFTER EACH DIRECTORY
    const directoryGroups: IDirectoryAnalysis[] = [];
    const directoryArray = Array.from(uniqueDirectories);
    let processedDirectories = 0;

    for (const directoryPath of directoryArray) {
      if (this.isCancelled) break;

      const rowIndexes = directoryToRows.get(directoryPath) || [];
      const fullSharePointPath = this.folderService.getFullDirectoryPath(directoryPath, baseFolderPath);
      
      directoryGroups.push({
        directoryPath,
        normalizedPath: this.normalizePath(directoryPath),
        exists: false, // Will be determined in stage 2
        fileCount: rowIndexes.length,
        rowIndexes,
        fullSharePointPath,
        hasValidPath: directoryPath.length > 0
      });

      processedDirectories++;

      // UPDATE AFTER EACH DIRECTORY as requested
      const stageProgress = 50 + ((processedDirectories / directoryArray.length) * 50);
      progress = SearchProgressHelper.updateStageProgress(
        progress,
        stageProgress,
        {
          currentFileName: `Processing directory ${processedDirectories}/${directoryArray.length}: ${directoryPath}`,
          directoriesAnalyzed: processedDirectories
        }
      );
      statusCallback?.(progress);

      // Small delay
      await this.delay(10);
    }

    // Sort by file count descending (process directories with more files first)
    directoryGroups.sort((a, b) => b.fileCount - a.fileCount);

    // Create search plan
    const searchPlan: ISearchPlan = {
      totalRows: rows.length,
      validRows,
      invalidRows: rows.length - validRows,
      totalDirectories: directoryGroups.length,
      existingDirectories: 0, // Will be determined in stage 2
      missingDirectories: 0, // Will be determined in stage 2
      directoryGroups,
      estimatedDuration: directoryGroups.length * 2 // Rough estimate: 2 seconds per directory
    };

    console.log('[FileSearchService] STAGE 1 completed (OPTIMIZED):', {
      uniqueDirectories: directoryGroups.length,
      totalFiles: searchPlan.validRows,
      invalidRows: searchPlan.invalidRows,
      avgFilesPerDirectory: validRows / directoryGroups.length,
      topDirectories: directoryGroups.slice(0, 3).map(d => ({ path: d.directoryPath, files: d.fileCount }))
    });

    // Final stage 1 progress
    progress = SearchProgressHelper.updateStageProgress(
      progress,
      100,
      {
        currentFileName: `Analyzed ${directoryGroups.length} unique directories with ${validRows} files`,
        directoriesAnalyzed: directoryGroups.length,
        totalDirectories: directoryGroups.length,
        searchPlan
      }
    );
    
    statusCallback?.(progress);
    
    return progress;
  }

  /**
   * STAGE 2: Check directory existence in SharePoint (25-50%) - UPDATE EACH DIRECTORY
   */
  private async executeStage2_CheckDirectoryExistence(
    currentProgress: ISearchProgress,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<ISearchProgress> {
    
    console.log('[FileSearchService] STAGE 2: Checking directory existence...');
    
    // Transition to stage 2
    let progress = SearchProgressHelper.transitionToStage(
      currentProgress,
      SearchStage.CHECKING_EXISTENCE,
      {
        currentFileName: 'Loading SharePoint folder structure...'
      }
    );
    statusCallback?.(progress);

    const searchPlan = currentProgress.searchPlan;
    if (!searchPlan) {
      throw new Error('Search plan not found from Stage 1');
    }

    // Load SharePoint folders if not already loaded (single operation)
    try {
      await this.folderService.loadAllSubfolders(
        searchPlan.directoryGroups[0]?.fullSharePointPath?.split('/').slice(0, -1).join('/') || '',
        (currentPath, foldersLoaded) => {
          if (statusCallback) {
            const loadProgress = Math.min(20, (foldersLoaded / 100) * 20); // Limit to 20% of stage 2
            const stageProgress = SearchProgressHelper.updateStageProgress(
              progress,
              loadProgress,
              {
                currentFileName: `Loading folders... (${foldersLoaded} loaded)`
              }
            );
            statusCallback(stageProgress);
          }
        }
      );
    } catch (error) {
      console.warn('[FileSearchService] Error loading folders, continuing with basic checks:', error);
    }

    // Check existence of each directory - UPDATE AFTER EACH as requested
    let checkedDirectories = 0;
    let existingDirectories = 0;

    for (const dirGroup of searchPlan.directoryGroups) {
      if (this.isCancelled) break;

      // Check if directory exists
      dirGroup.exists = this.folderService.checkDirectoryExists(dirGroup.fullSharePointPath);
      
      if (dirGroup.exists) {
        existingDirectories++;
      }

      checkedDirectories++;

      // UPDATE AFTER EACH DIRECTORY as requested
      const stageProgress = 20 + ((checkedDirectories / searchPlan.directoryGroups.length) * 80);
      progress = SearchProgressHelper.updateStageProgress(
        progress,
        stageProgress,
        {
          currentFileName: `Checking ${dirGroup.directoryPath}... (${dirGroup.exists ? 'EXISTS' : 'NOT FOUND'})`,
          directoriesChecked: checkedDirectories,
          existingDirectories
        }
      );
      
      statusCallback?.(progress);

      // Small delay
      await this.delay(50);
    }

    // Update search plan
    const updatedSearchPlan: ISearchPlan = {
      ...searchPlan,
      existingDirectories,
      missingDirectories: searchPlan.totalDirectories - existingDirectories
    };

    console.log('[FileSearchService] STAGE 2 completed:', {
      totalDirectories: searchPlan.totalDirectories,
      existingDirectories,
      missingDirectories: updatedSearchPlan.missingDirectories
    });

    // Final stage 2 progress
    progress = SearchProgressHelper.updateStageProgress(
      progress,
      100,
      {
        currentFileName: `${existingDirectories}/${searchPlan.totalDirectories} directories exist`,
        searchPlan: updatedSearchPlan
      }
    );
    
    statusCallback?.(progress);
    
    return progress;
  }

  /**
   * STAGE 3: Search for files in existing directories (50-100%) - BATCH OPTIMIZED
   */
  private async executeStage3_SearchFiles(
    currentProgress: ISearchProgress,
    rows: IRenameTableRow[],
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<void> {
    
    console.log('[FileSearchService] STAGE 3: Searching for files (BATCH OPTIMIZED)...');
    
    // Transition to stage 3
    let progress = SearchProgressHelper.transitionToStage(
      currentProgress,
      SearchStage.SEARCHING_FILES,
      {
        currentFileName: 'Starting file search...'
      }
    );
    statusCallback?.(progress);

    const searchPlan = currentProgress.searchPlan;
    if (!searchPlan) {
      throw new Error('Search plan not found');
    }

    let processedRows = 0;
    let filesFound = 0;
    const totalRowsToProcess = searchPlan.validRows;

    // BATCH TRACKING for progress updates
    const BATCH_SIZE = 50; // Update progress every 50 files
    const TIME_BATCH = 3000; // Or every 3 seconds
    let filesProcessedSinceUpdate = 0;
    let lastUpdateTime = Date.now();

    // Helper function for batch progress updates
    const updateProgressIfNeeded = (forceUpdate = false) => {
      const timeSinceUpdate = Date.now() - lastUpdateTime;
      const shouldUpdate = forceUpdate || 
                          filesProcessedSinceUpdate >= BATCH_SIZE || 
                          timeSinceUpdate >= TIME_BATCH;

      if (shouldUpdate) {
        const stageProgress = (processedRows / totalRowsToProcess) * 100;
        progress = SearchProgressHelper.updateStageProgress(
          progress,
          stageProgress,
          {
            currentRow: processedRows,
            currentFileName: `Processed ${processedRows}/${totalRowsToProcess} files`,
            filesSearched: processedRows,
            filesFound
          }
        );
        statusCallback?.(progress);
        
        filesProcessedSinceUpdate = 0;
        lastUpdateTime = Date.now();
      }
    };

    // Process each directory group
    for (const dirGroup of searchPlan.directoryGroups) {
      if (this.isCancelled) break;

      if (!dirGroup.exists) {
        // OPTIMIZED: Skip this entire directory with single progress update
        console.log(`[FileSearchService] Skipping directory (not found): ${dirGroup.directoryPath} (${dirGroup.fileCount} files)`);
        
        // Mass update all files in this directory as not found
        for (const rowIndex of dirGroup.rowIndexes) {
          if (!this.isCancelled) {
            results[rowIndex] = 'not-found';
            progressCallback(rowIndex, 'not-found');
            processedRows++;
            filesProcessedSinceUpdate++;
          }
        }
        
        // Update directory-level progress
        const stageProgress = (processedRows / totalRowsToProcess) * 100;
        progress = SearchProgressHelper.updateStageProgress(
          progress,
          stageProgress,
          {
            currentRow: processedRows,
            currentFileName: `Skipped directory: ${dirGroup.directoryPath} (${dirGroup.fileCount} files)`,
            currentDirectory: dirGroup.directoryPath,
            filesSearched: processedRows,
            filesFound
          }
        );
        statusCallback?.(progress);
        
      } else {
        // Directory exists - search for files with BATCH UPDATES
        console.log(`[FileSearchService] Searching in directory: ${dirGroup.directoryPath} (${dirGroup.fileCount} files)`);
        
        const foundFiles = await this.searchInSpecificDirectoryBatch(
          dirGroup,
          rows,
          results,
          progressCallback,
          (batchProcessed, batchFound, currentFile) => {
            processedRows += batchProcessed;
            filesFound += batchFound;
            filesProcessedSinceUpdate += batchProcessed;
            
            // Update current file info in progress
            progress = SearchProgressHelper.updateStageProgress(
              progress,
              (processedRows / totalRowsToProcess) * 100,
              {
                currentRow: processedRows,
                currentFileName: currentFile,
                currentDirectory: dirGroup.directoryPath,
                filesSearched: processedRows,
                filesFound
              }
            );
            
            // Check if batch update is needed
            updateProgressIfNeeded();
          }
        );
        
        filesFound += foundFiles;
      }

      // Small delay between directories
      await this.delay(25);
    }

    // Handle any remaining rows that weren't in valid directories
    const processedRowIndexes = new Set<number>();
    searchPlan.directoryGroups.forEach((g: IDirectoryAnalysis) => {
      g.rowIndexes.forEach((index: number) => processedRowIndexes.add(index));
    });
    
    const unprocessedRows = rows.filter(row => !processedRowIndexes.has(row.rowIndex));
    
    if (unprocessedRows.length > 0) {
      console.log(`[FileSearchService] Processing ${unprocessedRows.length} rows without valid directories`);
      
      for (const row of unprocessedRows) {
        if (this.isCancelled) break;
        
        results[row.rowIndex] = 'not-found';
        progressCallback(row.rowIndex, 'not-found');
        processedRows++;
        filesProcessedSinceUpdate++;
        
        // Batch update for unprocessed rows
        updateProgressIfNeeded();
      }
    }

    // Final progress update
    updateProgressIfNeeded(true);

    console.log('[FileSearchService] STAGE 3 completed (BATCH OPTIMIZED):', {
      processedRows,
      filesFound,
      successRate: processedRows > 0 ? (filesFound / processedRows * 100).toFixed(1) + '%' : '0%',
      batchSize: BATCH_SIZE,
      timeBasedBatches: TIME_BATCH + 'ms'
    });
  } 100).toFixed(1) + '%' : '0%'
    });
  }

  /**
   * NEW: Search for files in a specific directory with BATCH UPDATES
   */
  private async searchInSpecificDirectoryBatch(
    dirGroup: IDirectoryAnalysis,
    allRows: IRenameTableRow[],
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    batchCallback?: (batchProcessed: number, batchFound: number, currentFile: string) => void
  ): Promise<number> {
    
    console.log(`[FileSearchService] Batch searching in directory: ${dirGroup.fullSharePointPath}`);
    
    let totalFilesFound = 0;
    
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
      
      // Process files in batches
      const BATCH_SIZE = 25; // Process 25 files before callback
      let batchProcessed = 0;
      let batchFound = 0;
      let currentFileName = '';
      
      // Check each file in this directory group
      for (let i = 0; i < dirGroup.rowIndexes.length; i++) {
        if (this.isCancelled) break;
        
        const rowIndex = dirGroup.rowIndexes[i];
        const row = allRows.find(r => r.rowIndex === rowIndex);
        if (!row) continue;
        
        const fileName = String(row.cells['custom_0']?.value || '');
        currentFileName = fileName;
        
        // Check if file exists (case-insensitive)
        const fileExists = fileMap.has(fileName.toLowerCase());
        
        const result = fileExists ? 'found' : 'not-found';
        results[rowIndex] = result;
        progressCallback(rowIndex, result);
        
        if (fileExists) {
          batchFound++;
          totalFilesFound++;
        }
        
        batchProcessed++;
        
        console.log(`[FileSearchService] File "${fileName}" in "${dirGroup.directoryPath}": ${result.toUpperCase()}`);
        
        // Batch callback every BATCH_SIZE files or at the end
        if (batchProcessed >= BATCH_SIZE || i === dirGroup.rowIndexes.length - 1) {
          batchCallback?.(batchProcessed, batchFound, currentFileName);
          batchProcessed = 0;
          batchFound = 0;
          
          // Small delay between batches within directory
          if (i < dirGroup.rowIndexes.length - 1) {
            await this.delay(10);
          }
        }
      }
      
    } catch (error) {
      console.error(`[FileSearchService] Error searching in directory ${dirGroup.fullSharePointPath}:`, error);
      
      // Mark all files in this directory as not found
      for (const rowIndex of dirGroup.rowIndexes) {
        if (!this.isCancelled) {
          const row = allRows.find(r => r.rowIndex === rowIndex);
          const fileName = row ? String(row.cells['custom_0']?.value || '') : 'Unknown';
          
          results[rowIndex] = 'not-found';
          progressCallback(rowIndex, 'not-found');
        }
      }
      
      // Error callback
      batchCallback?.(dirGroup.rowIndexes.length, 0, `Error in ${dirGroup.directoryPath}`);
    }
    
    return totalFilesFound;
  }
  private async searchInSpecificDirectory(
    dirGroup: IDirectoryAnalysis,
    allRows: IRenameTableRow[],
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    fileCallback?: (fileName: string, found: boolean) => void
  ): Promise<number> {
    
    console.log(`[FileSearchService] Searching in directory: ${dirGroup.fullSharePointPath}`);
    
    let filesFound = 0;
    
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
        
        // Check if file exists (case-insensitive)
        const fileExists = fileMap.has(fileName.toLowerCase());
        
        const result = fileExists ? 'found' : 'not-found';
        results[rowIndex] = result;
        progressCallback(rowIndex, result);
        
        if (fileExists) {
          filesFound++;
        }
        
        console.log(`[FileSearchService] File "${fileName}" in "${dirGroup.directoryPath}": ${result.toUpperCase()}`);
        
        // Notify callback
        fileCallback?.(fileName, fileExists);
        
        // Small delay between files
        await this.delay(10);
      }
      
    } catch (error) {
      console.error(`[FileSearchService] Error searching in directory ${dirGroup.fullSharePointPath}:`, error);
      
      // Mark all files in this directory as not found
      for (const rowIndex of dirGroup.rowIndexes) {
        if (!this.isCancelled) {
          const row = allRows.find(r => r.rowIndex === rowIndex);
          const fileName = row ? String(row.cells['custom_0']?.value || '') : 'Unknown';
          
          results[rowIndex] = 'not-found';
          progressCallback(rowIndex, 'not-found');
          
          fileCallback?.(fileName, false);
        }
      }
    }
    
    return filesFound;
  }

  /**
   * Handle search cancellation
   */
  private handleCancellation(
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    rows: IRenameTableRow[],
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void
  ): { [rowIndex: number]: 'found' | 'not-found' | 'searching' } {
    
    console.log('[FileSearchService] Search was cancelled');
    
    // Mark all searching rows as not found
    rows.forEach(row => {
      if (results[row.rowIndex] === 'searching') {
        results[row.rowIndex] = 'not-found';
        progressCallback(row.rowIndex, 'not-found');
      }
    });
    
    return results;
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

  private normalizePath(path: string): string {
    return path
      .replace(/\\/g, '/')           // Convert backslashes to forward slashes
      .replace(/\/+/g, '/')          // Remove duplicate slashes
      .toLowerCase()                 // Case insensitive
      .replace(/\/$/, '');           // Remove trailing slash
  }

  // Keep existing methods for compatibility
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

  public getFileNameFromRow(row: IRenameTableRow): string {
    // Get filename directly from the first column (custom_0)
    const fileName = String(row.cells['custom_0']?.value || '');
    
    console.log(`[FileSearchService] getFileNameFromRow for row ${row.rowIndex}: "${fileName}"`);
    
    return fileName;
  }
}