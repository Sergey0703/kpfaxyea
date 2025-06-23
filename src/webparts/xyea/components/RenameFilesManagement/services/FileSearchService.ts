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
   * OPTIMIZED: Search for files in analyzed directories (Stage 3 only)
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
    
    console.log(`[FileSearchService] 🚀 STARTING OPTIMIZED FILE SEARCH (Search ID: ${searchId})`);
    
    const results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' } = {};
    
    try {
      // Initialize all rows as searching
      rows.forEach(row => {
        results[row.rowIndex] = 'searching';
        progressCallback(row.rowIndex, 'searching');
      });

      // OPTIMIZED STAGE 3: Search files with MINIMAL API calls
      await this.executeOptimizedStage3_SearchFiles(
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
   * OPTIMIZED STAGE 3: Search files with CORRECTED LOGIC and MINIMAL API calls
   */
  private async executeOptimizedStage3_SearchFiles(
    currentProgress: ISearchProgress,
    rows: IRenameTableRow[],
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<void> {
    
    console.log('[FileSearchService] 🚀 OPTIMIZED STAGE 3: Searching files with MINIMAL API calls...');
    
    let progress = SearchProgressHelper.transitionToStage(
      currentProgress,
      SearchStage.SEARCHING_FILES,
      {
        currentFileName: 'Building optimized search plan...'
      }
    );
    statusCallback?.(progress);

    const searchPlan = currentProgress.searchPlan;
    if (!searchPlan) {
      throw new Error('Search plan not found');
    }

    // STEP 1: Build directory-to-files mapping
    const directoryToFilesMap = this.buildDirectoryToFilesMap(rows, searchPlan);
    
    console.log(`[FileSearchService] 📊 Built directory mapping:`);
    Object.entries(directoryToFilesMap).forEach(([dir, files]) => {
      console.log(`  📁 "${dir}" -> ${files.length} files: [${files.slice(0, 3).map(f => f.fileName).join(', ')}...]`);
    });

    let processedFiles = 0;
    let foundFiles = 0;
    const totalFiles = rows.length;
    const directories = Object.keys(directoryToFilesMap);

    // STEP 2: Process each directory with ONE API call
    for (let dirIndex = 0; dirIndex < directories.length; dirIndex++) {
      const directoryPath = directories[dirIndex];
      const filesFromExcel = directoryToFilesMap[directoryPath];
      
      if (this.isCancelled) break;

      console.log(`[FileSearchService] 🔍 DIRECTORY ${dirIndex + 1}/${directories.length}: "${directoryPath}"`);
      console.log(`[FileSearchService] 📋 Looking for ${filesFromExcel.length} Excel files in this directory`);

      // Update progress
      progress = SearchProgressHelper.updateStageProgress(
        progress,
        (dirIndex / directories.length) * 100,
        {
          currentDirectory: directoryPath,
          currentFileName: `Loading directory contents...`,
          filesSearched: processedFiles,
          filesFound: foundFiles
        }
      );
      statusCallback?.(progress);

      try {
        // ONE API CALL to get directory contents
        console.log(`[FileSearchService] 📞 API call: getFolderContents("${directoryPath}")`);
        const startTime = Date.now();
        
        const folderContentsPromise = this.folderService.getFolderContents(directoryPath);
        const folderContents = await Promise.race([
          folderContentsPromise,
          this.createTimeoutPromise(this.BATCH_TIMEOUT, { files: [], folders: [] })
        ]) as {files: any[], folders: any[]};
        
        const endTime = Date.now();
        console.log(`[FileSearchService] ✅ API response received in ${endTime - startTime}ms, found ${folderContents.files.length} files`);

        // Create SharePoint files map (case-insensitive)
        const sharePointFilesMap = new Map<string, any>();
        folderContents.files.forEach(file => {
          sharePointFilesMap.set(file.Name.toLowerCase(), file);
        });

        console.log(`[FileSearchService] 🗂️ Created SharePoint files map: ${sharePointFilesMap.size} files`);

        // CHECK each Excel file against SharePoint files
        for (let fileIndex = 0; fileIndex < filesFromExcel.length; fileIndex++) {
          const excelFile = filesFromExcel[fileIndex];
          
          if (this.isCancelled) break;

          const fileExists = sharePointFilesMap.has(excelFile.fileName.toLowerCase());
          const result = fileExists ? 'found' : 'not-found';
          
          results[excelFile.rowIndex] = result;
          progressCallback(excelFile.rowIndex, result);
          
          if (fileExists) {
            foundFiles++;
            console.log(`[FileSearchService] ✅ FOUND: "${excelFile.fileName}" (row ${excelFile.rowIndex + 1})`);
          } else {
            console.log(`[FileSearchService] ❌ NOT FOUND: "${excelFile.fileName}" (row ${excelFile.rowIndex + 1})`);
          }
          
          processedFiles++;

          // Update progress every 10 files
          if (fileIndex % 10 === 0 || fileIndex === filesFromExcel.length - 1) {
            progress = SearchProgressHelper.updateStageProgress(
              progress,
              ((dirIndex + (fileIndex / filesFromExcel.length)) / directories.length) * 100,
              {
                currentDirectory: directoryPath,
                currentFileName: excelFile.fileName,
                filesSearched: processedFiles,
                filesFound: foundFiles
              }
            );
            statusCallback?.(progress);
          }
        }

      } catch (error) {
        console.error(`[FileSearchService] ❌ ERROR in directory "${directoryPath}":`, error);
        
        // Mark all files in this directory as not found
        filesFromExcel.forEach(excelFile => {
          if (!this.isCancelled) {
            results[excelFile.rowIndex] = 'not-found';
            progressCallback(excelFile.rowIndex, 'not-found');
            processedFiles++;
          }
        });
      }

      // Delay between directories to avoid throttling
      await this.delay(200);
      
      console.log(`[FileSearchService] 📊 Progress: ${processedFiles}/${totalFiles} files, ${foundFiles} found`);
    }

    console.log(`[FileSearchService] 🎯 OPTIMIZED SEARCH COMPLETED:`);
    console.log(`  📊 Files processed: ${processedFiles}/${totalFiles}`);
    console.log(`  ✅ Files found: ${foundFiles}`);
    console.log(`  📈 Success rate: ${processedFiles > 0 ? (foundFiles / processedFiles * 100).toFixed(1) + '%' : '0%'}`);
    console.log(`  🏗️ API calls made: ${directories.length} (instead of ${totalFiles})`);
    console.log(`  ⚡ Performance improvement: ${totalFiles > 0 ? Math.round(totalFiles / directories.length) : 0}x fewer API calls`);
  }

  /**
   * OPTIMIZATION: Build directory-to-files mapping for efficient processing
   */
  private buildDirectoryToFilesMap(
    rows: IRenameTableRow[], 
    searchPlan: ISearchPlan
  ): { [directoryPath: string]: Array<{ fileName: string; rowIndex: number }> } {
    
    console.log(`[FileSearchService] 🏗️ Building directory-to-files mapping...`);
    
    const directoryToFilesMap: { [directoryPath: string]: Array<{ fileName: string; rowIndex: number }> } = {};
    
    // Use searchPlan for efficient grouping
    searchPlan.directoryGroups.forEach(dirGroup => {
      if (!dirGroup.exists) {
        console.log(`[FileSearchService] ⏭️ Skipping non-existing directory: "${dirGroup.directoryPath}"`);
        return; // Skip non-existing directories
      }

      const filesInDirectory: Array<{ fileName: string; rowIndex: number }> = [];
      
      dirGroup.rowIndexes.forEach(rowIndex => {
        const row = rows.find(r => r.rowIndex === rowIndex);
        if (row) {
          const fileName = String(row.cells['custom_0']?.value || '').trim();
          if (fileName) {
            filesInDirectory.push({ fileName, rowIndex });
          }
        }
      });

      if (filesInDirectory.length > 0) {
        directoryToFilesMap[dirGroup.fullSharePointPath] = filesInDirectory;
        console.log(`[FileSearchService] 📁 Directory "${dirGroup.directoryPath}" -> ${filesInDirectory.length} files`);
      }
    });

    const totalDirectories = Object.keys(directoryToFilesMap).length;
    const totalFiles = Object.values(directoryToFilesMap).reduce((sum, files) => sum + files.length, 0);
    
    console.log(`[FileSearchService] 📊 Mapping created: ${totalDirectories} directories, ${totalFiles} files`);
    
    return directoryToFilesMap;
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