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

  // AGGRESSIVE: Much shorter timeouts to prevent hanging
  private readonly DIRECTORY_CHECK_TIMEOUT = 3000; // 3 seconds per directory
  private readonly FOLDER_LOAD_TIMEOUT = 8000; // 8 seconds for folder loading
  private readonly BATCH_TIMEOUT = 2000; // 2 seconds per batch

  constructor(context: any) {
    this.context = context;
    this.folderService = new SharePointFolderService(context);
    this.excelProcessor = new ExcelFileProcessor();
  }

  /**
   * NEW: Analyze directories and check existence (Stages 1-2)
   */
  public async analyzeDirectories(
    folderPath: string,
    rows: IRenameTableRow[],
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<ISearchProgress> {
    
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] Starting directory analysis (Search ID: ${searchId})`);
    
    let currentProgress = SearchProgressHelper.createInitialProgress();
    
    try {
      // STAGE 1: ANALYZING DIRECTORIES (0-50%)
      currentProgress = await this.executeStage1_AnalyzeDirectories(
        rows, 
        folderPath, 
        currentProgress, 
        statusCallback
      );
      
      if (this.isCancelled || this.currentSearchId !== searchId) {
        throw new Error('Analysis was cancelled');
      }

      // STAGE 2: CHECKING DIRECTORY EXISTENCE (50-100%)
      currentProgress = await this.executeStage2_CheckDirectoryExistence(
        currentProgress,
        statusCallback
      );
      
      if (this.isCancelled || this.currentSearchId !== searchId) {
        throw new Error('Analysis was cancelled');
      }

      console.log('[FileSearchService] Directory analysis completed successfully');
      return currentProgress;

    } catch (error) {
      console.error('[FileSearchService] Error during directory analysis:', error);
      
      const errorProgress = SearchProgressHelper.transitionToStage(
        currentProgress,
        SearchStage.ERROR,
        {
          currentFileName: 'Directory analysis failed',
          errors: [error instanceof Error ? error.message : 'Unknown error']
        }
      );
      statusCallback?.(errorProgress);
      throw error;
    }
  }

  /**
   * NEW: Search for files in analyzed directories (Stage 3 only)
   */
  public async searchFilesInDirectories(
    searchProgress: ISearchProgress,
    rows: IRenameTableRow[],
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<{ [rowIndex: number]: 'found' | 'not-found' | 'searching' }> {
    
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] Starting file search (Search ID: ${searchId})`);
    
    const results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' } = {};
    
    try {
      // Initialize all rows as searching
      rows.forEach(row => {
        results[row.rowIndex] = 'searching';
        progressCallback(row.rowIndex, 'searching');
      });

      // STAGE 3: SEARCHING FILES (0-100%)
      await this.executeStage3_SearchFiles(
        searchProgress,
        rows,
        results,
        progressCallback,
        statusCallback
      );

      // Mark completion
      if (!this.isCancelled && this.currentSearchId === searchId) {
        const finalProgress = SearchProgressHelper.transitionToStage(
          searchProgress,
          SearchStage.COMPLETED,
          {
            currentFileName: 'File search completed successfully',
            overallProgress: 100
          }
        );
        statusCallback?.(finalProgress);
      }

      console.log('[FileSearchService] File search completed:', results);
      return results;

    } catch (error) {
      console.error('[FileSearchService] Error during file search:', error);
      
      // Mark all unprocessed rows as not found
      rows.forEach(row => {
        if (results[row.rowIndex] === 'searching') {
          results[row.rowIndex] = 'not-found';
          progressCallback(row.rowIndex, 'not-found');
        }
      });

      const errorProgress = SearchProgressHelper.transitionToStage(
        searchProgress,
        SearchStage.ERROR,
        {
          currentFileName: 'File search failed',
          errors: [error instanceof Error ? error.message : 'Unknown error']
        }
      );
      statusCallback?.(errorProgress);
      
      return results;
    }
  }

  /**
   * STAGE 1: Analyze directories with timeout protection
   */
  private async executeStage1_AnalyzeDirectories(
    rows: IRenameTableRow[],
    baseFolderPath: string,
    currentProgress: ISearchProgress,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<ISearchProgress> {
    
    console.log('[FileSearchService] STAGE 1: Analyzing directories with timeout protection...');
    
    let progress = SearchProgressHelper.transitionToStage(
      currentProgress,
      SearchStage.ANALYZING_DIRECTORIES,
      {
        totalRows: rows.length,
        currentFileName: 'Extracting unique directories...'
      }
    );
    statusCallback?.(progress);

    // Fast extraction of unique directories
    const uniqueDirectories = new Set<string>();
    const directoryToRows = new Map<string, number[]>();
    let validRows = 0;

    rows.forEach(row => {
      const directoryCell = row.cells['custom_1'];
      let directoryPath = '';
      
      if (directoryCell && directoryCell.value) {
        directoryPath = String(directoryCell.value).trim();
      } else {
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

    progress = SearchProgressHelper.updateStageProgress(
      progress,
      50,
      {
        currentFileName: `Found ${uniqueDirectories.size} unique directories`,
        directoriesAnalyzed: uniqueDirectories.size,
        totalDirectories: uniqueDirectories.size
      }
    );
    statusCallback?.(progress);

    // Create directory analysis results
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
        exists: false,
        fileCount: rowIndexes.length,
        rowIndexes,
        fullSharePointPath,
        hasValidPath: directoryPath.length > 0
      });

      processedDirectories++;

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

      await this.delay(5); // Small delay
    }

    directoryGroups.sort((a, b) => b.fileCount - a.fileCount);

    const searchPlan: ISearchPlan = {
      totalRows: rows.length,
      validRows,
      invalidRows: rows.length - validRows,
      totalDirectories: directoryGroups.length,
      existingDirectories: 0,
      missingDirectories: 0,
      directoryGroups,
      estimatedDuration: directoryGroups.length * 2
    };

    progress = SearchProgressHelper.updateStageProgress(
      progress,
      100,
      {
        currentFileName: `Analyzed ${directoryGroups.length} unique directories`,
        searchPlan
      }
    );
    
    statusCallback?.(progress);
    return progress;
  }

  /**
   * STAGE 2: Check directory existence with timeout protection
   */
  private async executeStage2_CheckDirectoryExistence(
    currentProgress: ISearchProgress,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<ISearchProgress> {
    
    console.log('[FileSearchService] STAGE 2: Checking directory existence with timeouts...');
    
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

    // Load SharePoint folders with timeout
    try {
      const folderLoadPromise = this.folderService.loadAllSubfolders(
        searchPlan.directoryGroups[0]?.fullSharePointPath?.split('/').slice(0, -1).join('/') || '',
        (currentPath, foldersLoaded) => {
          if (statusCallback) {
            const loadProgress = Math.min(20, (foldersLoaded / 100) * 20);
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

      // Apply timeout to folder loading
      await Promise.race([
        folderLoadPromise,
        this.createTimeoutPromise(this.FOLDER_LOAD_TIMEOUT, 'Folder loading timeout')
      ]);

    } catch (error) {
      console.warn('[FileSearchService] Folder loading failed or timed out:', error);
      // Continue with basic directory checks
    }

    // Check existence of each directory with individual timeouts
    let checkedDirectories = 0;
    let existingDirectories = 0;

    for (const dirGroup of searchPlan.directoryGroups) {
      if (this.isCancelled) break;

      try {
        // Apply timeout to directory existence check
        const checkPromise = Promise.resolve(
          this.folderService.checkDirectoryExists(dirGroup.fullSharePointPath)
        );
        
        dirGroup.exists = await Promise.race([
          checkPromise,
          this.createTimeoutPromise(this.DIRECTORY_CHECK_TIMEOUT, false) // Return false on timeout
        ]) as boolean;
        
        if (dirGroup.exists) {
          existingDirectories++;
        }

      } catch (error) {
        console.warn(`[FileSearchService] Directory check failed for ${dirGroup.directoryPath}:`, error);
        dirGroup.exists = false; // Assume not exists on error
      }

      checkedDirectories++;

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
      await this.delay(50);
    }

    const updatedSearchPlan: ISearchPlan = {
      ...searchPlan,
      existingDirectories,
      missingDirectories: searchPlan.totalDirectories - existingDirectories
    };

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
   * STAGE 3: Search files with CORRECT LOGIC - only search files belonging to each directory
   */
  private async executeStage3_SearchFiles(
    currentProgress: ISearchProgress,
    rows: IRenameTableRow[],
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<void> {
    
    console.log('[FileSearchService] STAGE 3: Searching files with CORRECTED LOGIC...');
    
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

    const BATCH_SIZE = 50; // Increased since we're processing correctly now
    const TIME_BATCH = 2000;
    let filesProcessedSinceUpdate = 0;
    let lastUpdateTime = Date.now();

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
        
        console.log(`[FileSearchService] ⏱️ Progress update: ${processedRows}/${totalRowsToProcess} files (${filesFound} found)`);
      }
    };

    // NEW: Log all directories before processing with CORRECT file counts
    console.log(`[FileSearchService] 📁 CORRECTED DIRECTORY PROCESSING PLAN:`);
    searchPlan.directoryGroups.forEach((dirGroup, index) => {
      console.log(`  ${index + 1}. "${dirGroup.directoryPath}" -> ${dirGroup.exists ? '✅ EXISTS' : '❌ MISSING'} (${dirGroup.fileCount} files from Excel)`);
      console.log(`     Full path: "${dirGroup.fullSharePointPath}"`);
      console.log(`     Row indexes: [${dirGroup.rowIndexes.slice(0, 5).join(', ')}${dirGroup.rowIndexes.length > 5 ? '...' : ''}]`);
    });

    // Process each directory group with CORRECTED logic
    for (let dirIndex = 0; dirIndex < searchPlan.directoryGroups.length; dirIndex++) {
      const dirGroup = searchPlan.directoryGroups[dirIndex];
      
      if (this.isCancelled) break;

      console.log(`[FileSearchService] 🚀 STARTING DIRECTORY ${dirIndex + 1}/${searchPlan.directoryGroups.length}:`);
      console.log(`   Path: "${dirGroup.directoryPath}"`);
      console.log(`   Full SharePoint path: "${dirGroup.fullSharePointPath}"`);
      console.log(`   Exists: ${dirGroup.exists ? '✅ YES' : '❌ NO'}`);
      console.log(`   Excel rows for this directory: ${dirGroup.fileCount} (rows: ${dirGroup.rowIndexes.slice(0, 3).join(', ')}...)`);

      if (!dirGroup.exists) {
        // Skip non-existing directory - mark ONLY files from this directory
        console.log(`[FileSearchService] ⏭️ SKIPPING non-existing directory: ${dirGroup.directoryPath}`);
        console.log(`   Marking ${dirGroup.rowIndexes.length} files as not-found (only files from this directory)`);
        
        for (const rowIndex of dirGroup.rowIndexes) {
          if (!this.isCancelled) {
            results[rowIndex] = 'not-found';
            progressCallback(rowIndex, 'not-found');
            processedRows++;
            filesProcessedSinceUpdate++;
          }
        }
        
        updateProgressIfNeeded();
        console.log(`[FileSearchService] ✅ COMPLETED skipping directory: ${dirGroup.directoryPath}`);
        
      } else {
        // Search in existing directory - ONLY for files belonging to this directory
        console.log(`[FileSearchService] 🔍 SEARCHING in existing directory: ${dirGroup.directoryPath}`);
        console.log(`   Will process ${dirGroup.rowIndexes.length} Excel rows that belong to this directory`);
        console.log(`📞 About to call getFolderContents for: "${dirGroup.fullSharePointPath}"`);
        
        const directoryStartTime = Date.now();
        
        try {
          const foundInDirectory = await this.searchInSpecificDirectoryDetailed(
            dirGroup,
            rows,
            results,
            progressCallback,
            (batchProcessed: number, batchFound: number, currentFile: string) => {
              processedRows += batchProcessed;
              filesFound += batchFound;
              filesProcessedSinceUpdate += batchProcessed;
              
              console.log(`[FileSearchService] 📦 Batch completed in ${dirGroup.directoryPath}: +${batchProcessed} files, +${batchFound} found. Current file: "${currentFile}"`);
              
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
              
              updateProgressIfNeeded();
            }
          );

          const directoryEndTime = Date.now();
          const directoryDuration = directoryEndTime - directoryStartTime;
          console.log(`[FileSearchService] ✅ COMPLETED directory: ${dirGroup.directoryPath} in ${directoryDuration}ms`);
          console.log(`   Found ${foundInDirectory} files out of ${dirGroup.rowIndexes.length} Excel rows for this directory`);

        } catch (error) {
          const directoryEndTime = Date.now();
          const directoryDuration = directoryEndTime - directoryStartTime;
          
          console.error(`[FileSearchService] ❌ FAILED directory: ${dirGroup.directoryPath} after ${directoryDuration}ms`);
          console.error(`[FileSearchService] Error details:`, error);
          
          // Mark remaining files in this directory as not found - ONLY files from this directory
          let markedCount = 0;
          for (const rowIndex of dirGroup.rowIndexes) {
            if (results[rowIndex] === 'searching') {
              results[rowIndex] = 'not-found';
              progressCallback(rowIndex, 'not-found');
              processedRows++;
              filesProcessedSinceUpdate++;
              markedCount++;
            }
          }
          
          console.log(`[FileSearchService] 🔄 Marked ${markedCount} files from this directory as not-found due to error`);
          updateProgressIfNeeded();
        }
      }

      console.log(`[FileSearchService] 📊 DIRECTORY ${dirIndex + 1} SUMMARY:`);
      console.log(`   Total processed so far: ${processedRows}/${totalRowsToProcess}`);
      console.log(`   Total found so far: ${filesFound}`);
      console.log(`   Remaining directories: ${searchPlan.directoryGroups.length - dirIndex - 1}`);
      console.log(`[FileSearchService] ➡️ Moving to next directory...\n`);

      await this.delay(100);
    }

    // Handle unprocessed rows - these should be minimal now with correct logic
    const processedRowIndexes = new Set<number>();
    searchPlan.directoryGroups.forEach((g: IDirectoryAnalysis) => {
      g.rowIndexes.forEach((index: number) => processedRowIndexes.add(index));
    });
    
    const unprocessedRows = rows.filter(row => !processedRowIndexes.has(row.rowIndex));
    
    if (unprocessedRows.length > 0) {
      console.log(`[FileSearchService] ⚠️ Found ${unprocessedRows.length} unprocessed rows (should be minimal with correct logic)`);
      
      for (const row of unprocessedRows) {
        if (this.isCancelled) break;
        
        results[row.rowIndex] = 'not-found';
        progressCallback(row.rowIndex, 'not-found');
        processedRows++;
        filesProcessedSinceUpdate++;
        
        updateProgressIfNeeded();
      }
    }

    updateProgressIfNeeded(true);

    console.log('[FileSearchService] 🎯 STAGE 3 FINAL SUMMARY (CORRECTED LOGIC):');
    console.log(`   Total files processed: ${processedRows}`);
    console.log(`   Total files found: ${filesFound}`);
    console.log(`   Success rate: ${processedRows > 0 ? (filesFound / processedRows * 100).toFixed(1) + '%' : '0%'}`);
    console.log(`   Directories processed: ${searchPlan.directoryGroups.length}`);
  }

  /**
   * Helper method to create timeout promise
   */
  private createTimeoutPromise<T>(timeoutMs: number, errorMessage: string | T): Promise<T> {
    return new Promise((_, reject) => {
      setTimeout(() => {
        if (typeof errorMessage === 'string') {
          reject(new Error(errorMessage));
        } else {
          // For boolean returns, resolve with the fallback value
          reject(errorMessage);
        }
      }, timeoutMs);
    });
  }

  /**
   * NEW: Search in specific directory with DETAILED logging
   */
  private async searchInSpecificDirectoryDetailed(
    dirGroup: IDirectoryAnalysis,
    allRows: IRenameTableRow[],
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    batchCallback?: (batchProcessed: number, batchFound: number, currentFile: string) => void
  ): Promise<number> {
    
    console.log(`[FileSearchService] 🔍 DETAILED batch search starting for: "${dirGroup.fullSharePointPath}"`);
    
    let totalFilesFound = 0;
    
    try {
      console.log(`[FileSearchService] 📞 Calling folderService.getFolderContents("${dirGroup.fullSharePointPath}")...`);
      const contentStartTime = Date.now();
      
      // Get folder contents with timeout
      const contentPromise = this.folderService.getFolderContents(dirGroup.fullSharePointPath);
      const folderContents = await Promise.race([
        contentPromise,
        this.createTimeoutPromise(this.BATCH_TIMEOUT, { files: [], folders: [] })
      ]) as {files: any[], folders: any[]};
      
      const contentEndTime = Date.now();
      const contentDuration = contentEndTime - contentStartTime;
      
      const files = folderContents.files;
      
      console.log(`[FileSearchService] ✅ getFolderContents completed in ${contentDuration}ms`);
      console.log(`[FileSearchService] 📁 Found ${files.length} files in SharePoint directory "${dirGroup.directoryPath}"`);
      console.log(`[FileSearchService] 📊 Need to check ${dirGroup.rowIndexes.length} Excel rows against these ${files.length} SharePoint files`);
      
      // Create a map of filenames for fast lookup (case-insensitive)
      const fileMap = new Map<string, any>();
      files.forEach(file => {
        fileMap.set(file.Name.toLowerCase(), file);
      });
      
      console.log(`[FileSearchService] 🗂️ Created file lookup map with ${fileMap.size} entries`);
      
      // Process files in smaller batches with detailed logging
      const BATCH_SIZE = 10;
      let batchProcessed = 0;
      let batchFound = 0;
      let currentFileName = '';
      
      console.log(`[FileSearchService] 🔄 Starting to process ${dirGroup.rowIndexes.length} files in batches of ${BATCH_SIZE}...`);
      
      for (let i = 0; i < dirGroup.rowIndexes.length; i++) {
        if (this.isCancelled) break;
        
        const rowIndex = dirGroup.rowIndexes[i];
        const row = allRows.find(r => r.rowIndex === rowIndex);
        if (!row) {
          console.warn(`[FileSearchService] ⚠️ Row not found for index ${rowIndex}`);
          continue;
        }
        
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
          console.log(`[FileSearchService] ✅ FOUND: "${fileName}" (row ${rowIndex + 1})`);
        } else {
          console.log(`[FileSearchService] ❌ NOT FOUND: "${fileName}" (row ${rowIndex + 1})`);
        }
        
        batchProcessed++;
        
        // Batch callback every BATCH_SIZE files OR at the end
        if (batchProcessed >= BATCH_SIZE || i === dirGroup.rowIndexes.length - 1) {
          console.log(`[FileSearchService] 📦 Batch ${Math.floor(i / BATCH_SIZE) + 1} completed: ${batchProcessed} processed, ${batchFound} found`);
          
          batchCallback?.(batchProcessed, batchFound, currentFileName);
          batchProcessed = 0;
          batchFound = 0;
          
          // Shorter delay between batches
          if (i < dirGroup.rowIndexes.length - 1) {
            await this.delay(50);
          }
        }
      }
      
    } catch (error) {
      console.error(`[FileSearchService] ❌ CRITICAL ERROR in directory ${dirGroup.fullSharePointPath}:`, error);
      console.error(`[FileSearchService] Error type: ${error?.constructor?.name || 'Unknown'}`);
      console.error(`[FileSearchService] Error message: ${error instanceof Error ? error.message : String(error)}`);
      
      // Mark all files in this directory as not found
      let errorMarkedCount = 0;
      for (const rowIndex of dirGroup.rowIndexes) {
        if (!this.isCancelled) {
          results[rowIndex] = 'not-found';
          progressCallback(rowIndex, 'not-found');
          errorMarkedCount++;
        }
      }
      
      console.log(`[FileSearchService] 🔄 Marked ${errorMarkedCount} files as not-found due to error`);
      
      // Error callback
      batchCallback?.(dirGroup.rowIndexes.length, 0, `Error in ${dirGroup.directoryPath}`);
    }
    
    console.log(`[FileSearchService] ✅ Completed search in ${dirGroup.directoryPath}: ${totalFilesFound} files found out of ${dirGroup.rowIndexes.length} checked`);
    return totalFilesFound;
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
      .replace(/\\/g, '/')
      .replace(/\/+/g, '/')
      .toLowerCase()
      .replace(/\/$/, '');
  }

  /**
   * Keep existing methods for compatibility
   */
  public async searchSingleFile(folderPath: string, fileName: string): Promise<{ found: boolean; path?: string }> {
    try {
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
    const fileName = String(row.cells['custom_0']?.value || '');
    console.log(`[FileSearchService] getFileNameFromRow for row ${row.rowIndex}: "${fileName}"`);
    return fileName;
  }
}