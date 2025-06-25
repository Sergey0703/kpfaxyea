// src/webparts/xyea/components/RenameFilesManagement/services/FileSearchService.ts

import { 
  IRenameTableRow, 
  SearchStage, 
  ISearchProgress, 
  IDirectoryAnalysis, 
  ISearchPlan,
  SearchProgressHelper,
  DirectoryStatus,
  FileSearchStatus,
  DirectoryStatusCallback,
  FileSearchResultCallback
} from '../types/RenameFilesTypes';
import { SharePointFolderService } from './SharePointFolderService';
import { ExcelFileProcessor } from './ExcelFileProcessor';

// FIXED: Define specific interfaces instead of using 'any'
interface IWebPartContext {
  pageContext: {
    user: {
      displayName: string;
    };
    web: {
      absoluteUrl: string;
      serverRelativeUrl: string;
    };
  };
  aadTokenProviderFactory?: unknown;
}

interface ISharePointFolder {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  TimeCreated: string;
  TimeLastModified: string;
}

export class FileSearchService {
  private context: IWebPartContext; // FIXED: specific type instead of any
  private folderService: SharePointFolderService;
  private excelProcessor: ExcelFileProcessor;
  private isCancelled: boolean = false;
  private currentSearchId: string | undefined = undefined; // FIXED: undefined instead of null

  // AGGRESSIVE: Much shorter timeouts to prevent hanging
  private readonly DIRECTORY_CHECK_TIMEOUT = 3000; // 3 seconds per directory
  private readonly FOLDER_LOAD_TIMEOUT = 8000; // 8 seconds for folder loading

  constructor(context: IWebPartContext) { // FIXED: specific type instead of any
    this.context = context;
    this.folderService = new SharePointFolderService(context);
    this.excelProcessor = new ExcelFileProcessor();
  }

  /**
   * Calculate adaptive timeout based on file count
   */
  private calculateTimeout(fileCount: number): number {
    const baseTimeout = 2000;
    const additionalTime = Math.min(fileCount * 50, 15000); // max 15 seconds
    const adaptiveTimeout = baseTimeout + additionalTime;
    
    console.log(`[FileSearchService] 📊 Adaptive timeout for ${fileCount} files: ${adaptiveTimeout}ms`);
    return adaptiveTimeout;
  }

  /**
   * UPDATED: Analyze directories and check existence (Stages 1-2) with directory status callback
   */
  public async analyzeDirectories(
    folderPath: string,
    rows: IRenameTableRow[],
    statusCallback?: (progress: ISearchProgress) => void,
    directoryStatusCallback?: DirectoryStatusCallback // NEW: Directory status callback
  ): Promise<ISearchProgress> {
    
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] Starting directory analysis with status callback (Search ID: ${searchId})`);
    
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

      // STAGE 2: CHECKING DIRECTORY EXISTENCE (50-100%) - OPTIMIZED: Check directories, not rows
      currentProgress = await this.executeStage2_CheckDirectoryExistence_OPTIMIZED(
        currentProgress,
        statusCallback,
        directoryStatusCallback // NEW: Pass directory callback to Stage 2
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
   * UPDATED: Search for files in analyzed directories (Stage 3 only) with file search callback
   */
  public async searchFilesInDirectories(
    searchProgress: ISearchProgress,
    rows: IRenameTableRow[],
    progressCallback: FileSearchResultCallback, // UPDATED: Use typed callback
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<{ [rowIndex: number]: FileSearchStatus }> { // UPDATED: Return type
    
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] 🚀 STARTING OPTIMIZED FILE SEARCH (Search ID: ${searchId})`);
    
    const results: { [rowIndex: number]: FileSearchStatus } = {}; // UPDATED: Use typed results
    
    try {
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
      
      // NEW: Mark all unprocessed rows with appropriate status
      rows.forEach(row => {
        if (results[row.rowIndex] === 'searching') {
          // Determine if this is due to missing directory or actual search failure
          const directoryPath = this.getDirectoryFromRow(row);
          const directoryExists = this.checkDirectoryExistsInPlan(directoryPath, searchProgress.searchPlan);
          
          if (!directoryExists) {
            results[row.rowIndex] = 'directory-missing';
            progressCallback(row.rowIndex, 'directory-missing');
          } else {
            results[row.rowIndex] = 'not-found';
            progressCallback(row.rowIndex, 'not-found');
          }
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
      const directoryCell = row.cells.custom_1; // FIXED: dot notation
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
        directoryToRows.get(directoryPath)?.push(row.rowIndex); // FIXED: optional chaining instead of non-null assertion
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

      await this.delay(5);
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
   * OPTIMIZED STAGE 2: Check directories ONCE, not per row - FIXED IMPLEMENTATION
   */
private async executeStage2_CheckDirectoryExistence_OPTIMIZED(
  currentProgress: ISearchProgress,
  statusCallback?: (progress: ISearchProgress) => void,
  directoryStatusCallback?: DirectoryStatusCallback
): Promise<ISearchProgress> {
  
  console.log('[FileSearchService] 🚀 OPTIMIZED STAGE 2: Checking directories ONCE, not per row...');
  
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

  console.log(`[FileSearchService] 📊 EFFICIENCY GAINED: Checking ${searchPlan.directoryGroups.length} directories instead of ${searchPlan.totalRows} rows!`);
  console.log(`[FileSearchService] 🎯 API calls reduced from ${searchPlan.totalRows} to ${searchPlan.directoryGroups.length} (${Math.round(searchPlan.totalRows / searchPlan.directoryGroups.length)}x improvement)`);

  // NEW: Initialize all rows as 'checking' (bulk operation)
  if (directoryStatusCallback) {
    console.log('[FileSearchService] 🔄 Bulk initializing all rows as "checking"...');
    searchPlan.directoryGroups.forEach(dirGroup => {
      directoryStatusCallback(dirGroup.rowIndexes, 'checking');
    });
  }

  // Load SharePoint folders with timeout (optional optimization)
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

  // OPTIMIZED: Check each UNIQUE directory ONCE (not per row)
  let checkedDirectories = 0;
  let existingDirectories = 0;

  for (const dirGroup of searchPlan.directoryGroups) {
    if (this.isCancelled) break;

    console.log(`[FileSearchService] 🔍 Checking UNIQUE directory ${checkedDirectories + 1}/${searchPlan.directoryGroups.length}:`);
    console.log(`[FileSearchService] 📁 Path: "${dirGroup.directoryPath}"`);
    console.log(`[FileSearchService] 📊 Will update ${dirGroup.rowIndexes.length} rows with result`);

    try {
      // ONE API call per directory (not per row)
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

      // BULK UPDATE: Update ALL rows for this directory with the SAME result
      if (directoryStatusCallback) {
        const directoryStatus: DirectoryStatus = dirGroup.exists ? 'exists' : 'not-exists';
        
        console.log(`[FileSearchService] 📂 Directory "${dirGroup.directoryPath}" -> ${directoryStatus}`);
        console.log(`[FileSearchService] 🔄 Bulk updating ${dirGroup.rowIndexes.length} rows with status: ${directoryStatus}`);
        
        // NEW (fast) - bulk callback
        directoryStatusCallback(dirGroup.rowIndexes, directoryStatus);
        
        console.log(`[FileSearchService] ✅ Updated rows: ${dirGroup.rowIndexes.slice(0, 5).map(r => r + 1).join(', ')}${dirGroup.rowIndexes.length > 5 ? `, ... and ${dirGroup.rowIndexes.length - 5} more` : ''}`);
      }

    } catch (error) {
      console.warn(`[FileSearchService] Directory check failed for ${dirGroup.directoryPath}:`, error);
      dirGroup.exists = false; // Assume not exists on error

      // BULK UPDATE: Mark all rows for this directory as error
      if (directoryStatusCallback) {
        console.log(`[FileSearchService] 📂⚠️ Directory "${dirGroup.directoryPath}" -> error`);
        console.log(`[FileSearchService] 🔄 Bulk updating ${dirGroup.rowIndexes.length} rows with status: error`);
        
        directoryStatusCallback(dirGroup.rowIndexes, 'error');
      }
    }

    checkedDirectories++;

    // Progress based on DIRECTORIES, not rows
    const stageProgress = 20 + ((checkedDirectories / searchPlan.directoryGroups.length) * 80);
    progress = SearchProgressHelper.updateStageProgress(
      progress,
      stageProgress,
      {
        currentFileName: `Checked directory ${checkedDirectories}/${searchPlan.directoryGroups.length}: ${dirGroup.directoryPath} (${dirGroup.exists ? 'EXISTS' : 'NOT FOUND'})`,
        directoriesChecked: checkedDirectories,
        existingDirectories
      }
    );
    
    statusCallback?.(progress);
    await this.delay(50); // Small delay between directories
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
      currentFileName: `Optimization complete: ${existingDirectories}/${searchPlan.totalDirectories} directories exist`,
      searchPlan: updatedSearchPlan
    }
  );
  
  statusCallback?.(progress);
  
  console.log(`[FileSearchService] 🎯 STAGE 2 OPTIMIZATION COMPLETE:`);
  console.log(`[FileSearchService] ✅ Directories checked: ${checkedDirectories} (instead of ${searchPlan.totalRows} rows)`);
  console.log(`[FileSearchService] 📈 Performance improvement: ${Math.round(searchPlan.totalRows / checkedDirectories)}x fewer API calls`);
  console.log(`[FileSearchService] 📊 Results: ${existingDirectories} exist, ${searchPlan.totalDirectories - existingDirectories} missing`);
  
  return progress;
}

  /**
   * UPDATED: OPTIMIZED STAGE 3: Search files with CORRECTED LOGIC and NEW status handling
   */
  private async executeOptimizedStage3_SearchFiles(
    currentProgress: ISearchProgress,
    rows: IRenameTableRow[],
    results: { [rowIndex: number]: FileSearchStatus }, // UPDATED: Use typed results
    progressCallback: FileSearchResultCallback, // UPDATED: Use typed callback
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<void> {
    
    console.log('[FileSearchService] 🚀 OPTIMIZED STAGE 3 with NEW STATUS LOGIC...');
    
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

    // STEP 1: Initialize ALL rows with appropriate status FIRST
    this.initializeAllRowsWithCorrectStatus(rows, searchPlan, results, progressCallback);

    // STEP 2: Build directory-to-files mapping for EXISTING directories only
    const directoryToFilesMap = this.buildDirectoryToFilesMap(rows, searchPlan);
    
    console.log(`[FileSearchService] 📊 Built directory mapping:`);
    Object.entries(directoryToFilesMap).forEach(([dir, files]) => {
      console.log(`  📁 "${dir}" -> ${files.length} files to search`);
    });

    let processedFiles = 0;
    let foundFiles = 0;
    const directories = Object.keys(directoryToFilesMap);
    const totalFilesToSearch = Object.values(directoryToFilesMap).reduce((sum, files) => sum + files.length, 0);

    console.log(`[FileSearchService] 🎯 STARTING SEARCH: ${totalFilesToSearch} files in ${directories.length} EXISTING directories`);

    // STEP 3: Process each EXISTING directory with ONE API call
    for (let dirIndex = 0; dirIndex < directories.length; dirIndex++) {
      const directoryPath = directories[dirIndex];
      const filesFromExcel = directoryToFilesMap[directoryPath];
      
      if (this.isCancelled) break;

      console.log(`[FileSearchService] 🔍 DIRECTORY ${dirIndex + 1}/${directories.length}: "${directoryPath}"`);
      console.log(`[FileSearchService] 📋 Looking for ${filesFromExcel.length} Excel files in this EXISTING directory`);

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
        // ONE API CALL to get directory contents with adaptive timeout
        console.log(`[FileSearchService] 📞 API call: getFolderContents("${directoryPath}")`);
        const startTime = Date.now();
        
        const adaptiveTimeout = this.calculateTimeout(filesFromExcel.length);
        const folderContentsPromise = this.folderService.getFolderContents(directoryPath);
        const folderContents = await Promise.race([
          folderContentsPromise,
          this.createTimeoutPromise(adaptiveTimeout, { files: [], folders: [] })
        ]) as {files: ISharePointFolder[], folders: ISharePointFolder[]}; // FIXED: specific type instead of any
        
        const endTime = Date.now();
        console.log(`[FileSearchService] ✅ API response received in ${endTime - startTime}ms`);
        console.log(`[FileSearchService] 📄 SharePoint files found: ${folderContents.files.length}`);
        console.log(`[FileSearchService] 📁 SharePoint folders found: ${folderContents.folders.length}`);

        // IMPROVED: Handle empty directories gracefully
        if (folderContents.files.length === 0) {
          console.log(`[FileSearchService] ⚠️ Directory is empty or doesn't exist: "${directoryPath}"`);
          console.log(`[FileSearchService] 📝 Marking all ${filesFromExcel.length} files as NOT FOUND`);
          
          // Mark all files in this directory as not found
          filesFromExcel.forEach(excelFile => {
            if (!this.isCancelled) {
              results[excelFile.rowIndex] = 'not-found';
              progressCallback(excelFile.rowIndex, 'not-found');
              processedFiles++;
            }
          });
          
          console.log(`[FileSearchService] 📁 DIRECTORY SUMMARY "${directoryPath}": 0/${filesFromExcel.length} files found (empty directory)`);
          continue; // Skip to next directory
        }

        // Show sample of SharePoint files
        console.log(`[FileSearchService] 📋 Sample SharePoint files:`, 
          folderContents.files.slice(0, 5).map(f => `"${f.Name}"`).join(', ')
        );

        // Show sample of Excel files we're looking for
        console.log(`[FileSearchService] 🔍 Sample Excel files to find:`, 
          filesFromExcel.slice(0, 5).map(f => `"${f.fileName}"`).join(', ')
        );

        // Create SharePoint files map (case-insensitive)
        const sharePointFilesMap = new Map<string, ISharePointFolder>(); // FIXED: specific type instead of any
        folderContents.files.forEach(file => {
          sharePointFilesMap.set(file.Name.toLowerCase(), file);
        });

        console.log(`[FileSearchService] 🗂️ Created SharePoint files lookup map: ${sharePointFilesMap.size} entries`);

        // CHECK each Excel file against SharePoint files with DETAILED LOGGING
        let directoryFoundCount = 0;
        const BATCH_SIZE = 20; // Process in batches of 20 for logging

        for (let fileIndex = 0; fileIndex < filesFromExcel.length; fileIndex++) {
          const excelFile = filesFromExcel[fileIndex];
          
          if (this.isCancelled) break;

          const fileExists = sharePointFilesMap.has(excelFile.fileName.toLowerCase());
          const result: FileSearchStatus = fileExists ? 'found' : 'not-found'; // UPDATED: Use typed result
          
          results[excelFile.rowIndex] = result;
          progressCallback(excelFile.rowIndex, result);
          
          if (fileExists) {
            foundFiles++;
            directoryFoundCount++;
            console.log(`[FileSearchService] ✅ FOUND ${foundFiles}: "${excelFile.fileName}" (row ${excelFile.rowIndex + 1})`);
          } else {
            console.log(`[FileSearchService] ❌ NOT FOUND: "${excelFile.fileName}" (row ${excelFile.rowIndex + 1})`);
          }
          
          processedFiles++;

          // Batch progress logging
          if ((fileIndex + 1) % BATCH_SIZE === 0 || fileIndex === filesFromExcel.length - 1) {
            console.log(`[FileSearchService] 📦 BATCH PROGRESS: Processed ${fileIndex + 1}/${filesFromExcel.length} files in this directory`);
            console.log(`[FileSearchService] 📊 Current totals: ${foundFiles} found out of ${processedFiles} processed`);
            
            // Update progress every batch
            progress = SearchProgressHelper.updateStageProgress(
              progress,
              ((dirIndex + ((fileIndex + 1) / filesFromExcel.length)) / directories.length) * 100,
              {
                currentDirectory: directoryPath,
                currentFileName: excelFile.fileName,
                filesSearched: processedFiles,
                filesFound: foundFiles
              }
            );
            statusCallback?.(progress);

            // Small delay to prevent UI freezing
            await this.delay(50);
          }
        }

        console.log(`[FileSearchService] 📁 DIRECTORY SUMMARY "${directoryPath}":`);
        console.log(`  ✅ Found: ${directoryFoundCount}/${filesFromExcel.length} files`);
        console.log(`  📊 Success rate: ${filesFromExcel.length > 0 ? (directoryFoundCount / filesFromExcel.length * 100).toFixed(1) + '%' : '0%'}`);

      } catch (error) {
        console.error(`[FileSearchService] Error type: ${error?.constructor?.name || 'Unknown'}`);
        console.error(`[FileSearchService] Error message: ${error instanceof Error ? error.message : String(error)}`);
        
        // IMPROVED: Better error handling for non-existent directories
        if (error instanceof Error && (error.message.includes('404') || error.message.includes('Not Found'))) {
          console.log(`[FileSearchService] 📝 Directory doesn't exist, marking ${filesFromExcel.length} files as NOT FOUND`);
        } else {
          console.log(`[FileSearchService] 📝 API error, marking ${filesFromExcel.length} files as NOT FOUND`);
        }
        
        // Mark all files in this directory as not found
        filesFromExcel.forEach(excelFile => {
          if (!this.isCancelled) {
            results[excelFile.rowIndex] = 'not-found';
            progressCallback(excelFile.rowIndex, 'not-found');
            processedFiles++;
          }
        });
        
        console.log(`[FileSearchService] 📁 DIRECTORY SUMMARY "${directoryPath}": 0/${filesFromExcel.length} files found (error/not exist)`);
      }

      // Delay between directories to avoid throttling
      await this.delay(200);
      
      console.log(`[FileSearchService] 📊 OVERALL PROGRESS: ${processedFiles}/${totalFilesToSearch} files searched, ${foundFiles} found`);
      console.log(`[FileSearchService] ➡️ Moving to next directory...\n`);
    }

    console.log(`[FileSearchService] 🎯 OPTIMIZED SEARCH COMPLETED:`);
    console.log(`  📊 Files processed: ${processedFiles}/${totalFilesToSearch}`);
    console.log(`  ✅ Files found: ${foundFiles}`);
    console.log(`  📈 Success rate: ${processedFiles > 0 ? (foundFiles / processedFiles * 100).toFixed(1) + '%' : '0%'}`);
    console.log(`  🏗️ API calls made: ${directories.length} (instead of total files)`);
    console.log(`  ⚡ Performance improvement: ${totalFilesToSearch > 0 ? Math.round(totalFilesToSearch / directories.length) : 0}x fewer API calls`);
  }

  /**
   * NEW: Initialize all rows with correct status based on directory existence
   */
  private initializeAllRowsWithCorrectStatus(
    rows: IRenameTableRow[],
    searchPlan: ISearchPlan,
    results: { [rowIndex: number]: FileSearchStatus },
    progressCallback: FileSearchResultCallback
  ): void {
    console.log('[FileSearchService] 📋 Initializing all rows with correct status...');
    
    let directoryMissingCount = 0;
    let searchingCount = 0;
    
    rows.forEach(row => {
      const directoryPath = this.getDirectoryFromRow(row);
      const directoryGroup = searchPlan.directoryGroups.find(group => 
        group.rowIndexes.includes(row.rowIndex)
      );
      
      if (!directoryGroup || !directoryGroup.exists) {
        // Files in missing directories get 'directory-missing' status
        results[row.rowIndex] = 'directory-missing';
        progressCallback(row.rowIndex, 'directory-missing');
        directoryMissingCount++;
        console.log(`[FileSearchService] 📁❌ Row ${row.rowIndex + 1}: directory-missing (${directoryPath})`);
      } else {
        // Files in existing directories get 'searching' status (will be updated during search)
        results[row.rowIndex] = 'searching';
        progressCallback(row.rowIndex, 'searching');
        searchingCount++;
      }
    });
    
    console.log(`[FileSearchService] ✅ Status initialization complete:`);
    console.log(`  📁❌ Directory missing: ${directoryMissingCount} files`);
    console.log(`  🔍 Ready to search: ${searchingCount} files`);
  }

  /**
   * NEW: Helper method to get directory path from row
   */
  private getDirectoryFromRow(row: IRenameTableRow): string {
    const directoryCell = row.cells.custom_1; // FIXED: dot notation
    if (directoryCell && directoryCell.value) {
      return String(directoryCell.value).trim();
    }
    return this.excelProcessor.extractDirectoryPathFromRow(row);
  }

  /**
   * NEW: Check if directory exists in search plan
   */
  private checkDirectoryExistsInPlan(directoryPath: string, searchPlan?: ISearchPlan): boolean {
    if (!searchPlan) return false;
    
    const directoryGroup = searchPlan.directoryGroups.find(group => 
      group.directoryPath === directoryPath
    );
    
    return directoryGroup?.exists || false;
  }

  /**
   * OPTIMIZATION: Build directory-to-files mapping for efficient processing
   * FIXED: Only process directories that exist (exists: true)
   */
  private buildDirectoryToFilesMap(
    rows: IRenameTableRow[], 
    searchPlan: ISearchPlan
  ): { [directoryPath: string]: Array<{ fileName: string; rowIndex: number }> } {
    
    console.log(`[FileSearchService] 🏗️ Building directory-to-files mapping...`);
    console.log(`[FileSearchService] 📊 Total directories in plan: ${searchPlan.directoryGroups.length}`);
    console.log(`[FileSearchService] ✅ Existing directories: ${searchPlan.existingDirectories}`);
    console.log(`[FileSearchService] ❌ Missing directories: ${searchPlan.missingDirectories}`);
    
    const directoryToFilesMap: { [directoryPath: string]: Array<{ fileName: string; rowIndex: number }> } = {};
    
    // CORRECTED: Only process directories that exist (exists: true)
    searchPlan.directoryGroups.forEach(dirGroup => {
      if (!dirGroup.exists) {
        console.log(`[FileSearchService] ⏭️ Skipping non-existing directory: "${dirGroup.directoryPath}" (${dirGroup.fileCount} files skipped)`);
        return; // Skip non-existing directories
      }

      console.log(`[FileSearchService] ✅ Processing existing directory: "${dirGroup.directoryPath}" (exists: ${dirGroup.exists})`);

      const filesInDirectory: Array<{ fileName: string; rowIndex: number }> = [];
      
      dirGroup.rowIndexes.forEach(rowIndex => {
        const row = rows.find(r => r.rowIndex === rowIndex);
        if (row) {
          const fileName = String(row.cells.custom_0?.value || '').trim(); // FIXED: dot notation
          if (fileName) {
            filesInDirectory.push({ fileName, rowIndex });
          }
        }
      });

      if (filesInDirectory.length > 0) {
        directoryToFilesMap[dirGroup.fullSharePointPath] = filesInDirectory;
        console.log(`[FileSearchService] ✅ Added existing directory "${dirGroup.directoryPath}" -> ${filesInDirectory.length} files`);
        console.log(`[FileSearchService] 📋 Sample files: [${filesInDirectory.slice(0, 3).map(f => `"${f.fileName}"`).join(', ')}...]`);
      } else {
        console.log(`[FileSearchService] ⚠️ No files found for existing directory "${dirGroup.directoryPath}"`);
      }
    });

    const totalDirectories = Object.keys(directoryToFilesMap).length;
    const totalFiles = Object.values(directoryToFilesMap).reduce((sum, files) => sum + files.length, 0);
    const skippedDirectories = searchPlan.directoryGroups.length - totalDirectories;
    
    console.log(`[FileSearchService] 📊 FINAL mapping created:`);
    console.log(`[FileSearchService]   ✅ Existing directories to search: ${totalDirectories}`);
    console.log(`[FileSearchService]   ⏭️ Skipped non-existing directories: ${skippedDirectories}`);
    console.log(`[FileSearchService]   📄 Total files to search: ${totalFiles}`);
    console.log(`[FileSearchService] 📁 Directories to process: ${Object.keys(directoryToFilesMap).map(path => path.split('/').slice(-3).join('/')).join(', ')}`);
    
    return directoryToFilesMap;
  }

  /**
   * Rename found files with staffID prefix - UPDATED: Skip if target exists
   */
  public async renameFoundFiles(
    rows: IRenameTableRow[],
    fileSearchResults: { [rowIndex: number]: FileSearchStatus }, // UPDATED: Use typed results
    baseFolderPath: string,
    progressCallback: (rowIndex: number, status: 'renaming' | 'renamed' | 'error' | 'skipped') => void,
    statusCallback?: (progress: { 
      current: number; 
      total: number; 
      fileName: string; 
      success: number; 
      errors: number; 
      skipped: number;
    }) => void
  ): Promise<{ 
    success: number; 
    errors: number; 
    skipped: number;
    errorDetails: string[];
    skippedDetails: string[];
  }> {
    
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] 🏷️ STARTING FILE RENAME (Search ID: ${searchId})`);
    
    // Подготовка файлов для переименования
    const filesToRename: Array<{
      rowIndex: number;
      originalFileName: string;
      staffID: string;
      directoryPath: string;
      fullOriginalPath: string;
      fullNewPath: string;
      newFileName: string;
    }> = [];

    // Собираем и валидируем файлы для переименования
    rows.forEach(row => {
      const searchResult = fileSearchResults[row.rowIndex];
      
      if (searchResult === 'found') {
        const originalFileName = String(row.cells.custom_0?.value || '').trim(); // FIXED: dot notation
        const directoryPath = String(row.cells.custom_1?.value || '').trim(); // FIXED: dot notation
        
        // Поиск staffID в разных возможных колонках
        let staffID = '';
        
        const staffIDColumns = ['staffID', 'staffid', 'StaffID', 'staff_id', 'ID', 'id'];
        for (const columnName of staffIDColumns) {
          const cellValue = String(row.cells[columnName]?.value || '').trim();
          if (cellValue) {
            staffID = cellValue;
            break;
          }
        }
        
        if (!staffID) {
          const excelColumns = Object.keys(row.cells).filter(key => key.startsWith('excel_'));
          for (const columnId of excelColumns) {
            const cellValue = String(row.cells[columnId]?.value || '').trim();
            if (cellValue && /^[0-9A-Za-z]{1,10}$/.test(cellValue)) {
              staffID = cellValue;
              console.log(`[FileSearchService] 📋 Found staffID "${staffID}" in column ${columnId} for row ${row.rowIndex}`);
              break;
            }
          }
        }
        
        if (originalFileName && staffID && directoryPath) {
          const directorySharePointPath = this.buildDirectoryPath(directoryPath, baseFolderPath);
          const fullOriginalPath = `${directorySharePointPath}/${originalFileName}`;
          
          const newFileName = this.generateSafeFileName(originalFileName, staffID, directorySharePointPath);
          const fullNewPath = `${directorySharePointPath}/${newFileName}`;
          
          filesToRename.push({
            rowIndex: row.rowIndex,
            originalFileName,
            staffID,
            directoryPath,
            fullOriginalPath,
            fullNewPath,
            newFileName
          });
          
          console.log(`[FileSearchService] 📝 Prepared rename: "${originalFileName}" -> "${newFileName}"`);
        } else {
          console.warn(`[FileSearchService] ⚠️ Missing data for row ${row.rowIndex}`);
        }
      }
    });

    console.log(`[FileSearchService] 📊 Prepared ${filesToRename.length} files for renaming`);

    if (filesToRename.length === 0) {
      console.warn(`[FileSearchService] ⚠️ No files prepared for renaming`);
      return { 
        success: 0, 
        errors: 0, 
        skipped: 0, 
        errorDetails: ['No files prepared for renaming'], 
        skippedDetails: [] 
      };
    }

    let processedFiles = 0;
    let successCount = 0;
    let errorCount = 0;
    let skippedCount = 0; // Counter for skipped files
    const errorDetails: string[] = [];
    const skippedDetails: string[] = []; // Details for skipped files

    try {
      let requestDigest = await this.getRequestDigest(); // FIXED: changed to let
      const BATCH_SIZE = 1;
      
      for (let i = 0; i < filesToRename.length; i += BATCH_SIZE) {
        if (this.isCancelled || this.currentSearchId !== searchId) {
          console.log('[FileSearchService] ❌ Rename operation cancelled');
          break;
        }

        // Обновляем Request Digest каждые 100 файлов для длительных операций
        if (processedFiles % 100 === 0) {
          console.log(`[FileSearchService] 🔄 Refreshing request digest after ${processedFiles} files...`);
          requestDigest = await this.getRequestDigest();
          console.log(`[FileSearchService] ✅ Request digest refreshed`);
        }

        const batch = filesToRename.slice(i, i + BATCH_SIZE);
        console.log(`[FileSearchService] 📦 Processing file ${i + 1}/${filesToRename.length}`);

        for (const fileInfo of batch) {
          if (this.isCancelled) break;

          try {
            progressCallback(fileInfo.rowIndex, 'renaming');
            
            statusCallback?.({
              current: processedFiles + 1,
              total: filesToRename.length,
              fileName: fileInfo.originalFileName,
              success: successCount,
              errors: errorCount,
              skipped: skippedCount
            });

            console.log(`[FileSearchService] 🔄 Processing file ${processedFiles + 1}/${filesToRename.length}:`);
            console.log(`  Original: "${fileInfo.originalFileName}"`);
            console.log(`  New: "${fileInfo.newFileName}"`);
            console.log(`  StaffID: "${fileInfo.staffID}"`);

            await this.renameSingleFile(fileInfo.fullOriginalPath, fileInfo.fullNewPath, requestDigest);
            
            successCount++;
            progressCallback(fileInfo.rowIndex, 'renamed');
            console.log(`[FileSearchService] ✅ SUCCESS: "${fileInfo.originalFileName}" -> "${fileInfo.newFileName}"`);
            
          } catch (error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            
            // Check if this is a "file already exists" error
            if (errorMessage.startsWith('FILE_ALREADY_EXISTS:')) {
              // File already exists - skip it
              skippedCount++;
              const skippedMessage = `Row ${fileInfo.rowIndex + 1} - ${fileInfo.originalFileName}: Target file already exists, skipped to avoid overwrite`;
              skippedDetails.push(skippedMessage);
              progressCallback(fileInfo.rowIndex, 'skipped');
              console.log(`[FileSearchService] ⏭️ SKIPPED: "${fileInfo.originalFileName}" (target exists)`);
            } else {
              // Other error - count as error
              errorCount++;
              const detailedError = `Row ${fileInfo.rowIndex + 1} - ${fileInfo.originalFileName}: ${errorMessage}`;
              errorDetails.push(detailedError);
              progressCallback(fileInfo.rowIndex, 'error');
              console.error(`[FileSearchService] ❌ ERROR: "${fileInfo.originalFileName}": ${errorMessage}`);
            }
          }
          
          processedFiles++;
          await this.delay(2000);
        }
      }

      console.log(`[FileSearchService] 🎯 Rename completed:`);
      console.log(`  📊 Total files: ${filesToRename.length}`);
      console.log(`  ✅ Successful: ${successCount}`);
      console.log(`  ❌ Failed: ${errorCount}`);
      console.log(`  ⏭️ Skipped: ${skippedCount}`);
      console.log(`  📈 Success rate: ${filesToRename.length > 0 ? (successCount / filesToRename.length * 100).toFixed(1) + '%' : '0%'}`);

      // Log skipped files
      if (skippedDetails.length > 0) {
        console.log(`[FileSearchService] 📋 Skipped files (target already exists):`);
        skippedDetails.slice(0, 3).forEach((skipped, index) => {
          console.log(`  ${index + 1}. ${skipped}`);
        });
      }

      if (errorDetails.length > 0) {
        console.error(`[FileSearchService] 📋 Error details:`);
        errorDetails.slice(0, 3).forEach((error, index) => {
          console.error(`  ${index + 1}. ${error}`);
        });
      }

      return { 
        success: successCount, 
        errors: errorCount, 
        skipped: skippedCount,
        errorDetails, 
        skippedDetails
      };

    } catch (error) {
      console.error('[FileSearchService] ❌ Critical error in rename operation:', error);
      
      const errorMessage = error instanceof Error ? error.message : String(error);
      errorDetails.push(`Critical error: ${errorMessage}`);
      
      return { 
        success: successCount, 
        errors: filesToRename.length - successCount - skippedCount, 
        skipped: skippedCount,
        errorDetails, 
        skippedDetails
      };
    }
  }

  /**
   * Generate safe filename with staffID prefix
   */
  private generateSafeFileName(originalFileName: string, staffID: string, directoryPath: string): string {
    const cleanStaffID = staffID.replace(/[<>:"/\\|?*]/g, '').trim();
    
    if (originalFileName.toLowerCase().startsWith(cleanStaffID.toLowerCase())) {
      console.log(`[FileSearchService] ⚠️ File already starts with staffID: "${originalFileName}"`);
      return originalFileName;
    }
    
    const newFileName = `${cleanStaffID} ${originalFileName}`;
    
    const fullPath = `${directoryPath}/${newFileName}`;
    if (fullPath.length > 380) {
      console.warn(`[FileSearchService] ⚠️ Path too long, truncating filename`);
      
      const extension = originalFileName.split('.').pop();
      const baseName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
      const maxBaseLength = 200 - cleanStaffID.length - (extension?.length || 0) - 3; // FIXED: optional chaining
      const truncatedBase = baseName.substring(0, maxBaseLength);
      
      return `${cleanStaffID} ${truncatedBase}.${extension}`;
    }
    
    return newFileName;
  }

  /**
   * Clean and normalize SharePoint paths
   */
  private cleanSharePointPath(path: string): string {
    let cleanPath = path.trim().replace(/\\/g, '/');
    cleanPath = cleanPath.replace(/\/+/g, '/');
    cleanPath = cleanPath.replace(/\/$/, '');
    
    if (!cleanPath.startsWith('/')) {
      cleanPath = '/' + cleanPath;
    }
    
    console.log(`[FileSearchService] Path cleaning: "${path}" -> "${cleanPath}"`);
    return cleanPath;
  }

  /**
   * Check if file exists at given path
   */
  private async checkFileExists(filePath: string): Promise<{ exists: boolean; error?: string }> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const checkUrl = `${webUrl}/_api/web/getFileByServerRelativeUrl('${encodeURIComponent(filePath)}')`;
      
      console.log(`[FileSearchService] 🔍 Checking file existence: ${checkUrl}`);
      
      const response = await fetch(checkUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (response.ok) {
        console.log(`[FileSearchService] ✅ File exists: "${filePath}"`);
        return { exists: true };
      } else if (response.status === 404) {
        console.log(`[FileSearchService] ❌ File does not exist: "${filePath}"`);
        return { exists: false };
      } else {
        console.log(`[FileSearchService] ⚠️ Unknown status ${response.status} for file: "${filePath}"`);
        return { exists: false, error: `HTTP ${response.status}` };
      }
    } catch (error) {
      console.log(`[FileSearchService] ⚠️ Error checking file existence: ${error}`);
      return { exists: false, error: String(error) };
    }
  }

  /**
   * Try simple MoveTo API with proper encoding
   */
  private async trySimpleMoveTo(originalPath: string, newPath: string, requestDigest: string): Promise<boolean> {
    try {
      console.log(`[FileSearchService] 🔄 Trying simple MoveTo API`);
      
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const moveToUrl = `${webUrl}/_api/web/getFileByServerRelativeUrl('${originalPath}')/MoveTo(newurl='${newPath}',flags=1)`;
      
      console.log(`[FileSearchService] 📞 Simple MoveTo URL:`, moveToUrl);
      
      const response = await fetch(moveToUrl, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': requestDigest
        }
      });
      
      if (response.ok) {
        console.log(`[FileSearchService] ✅ Simple MoveTo succeeded`);
        return true;
      } else {
        const errorText = await response.text();
        console.log(`[FileSearchService] ❌ Simple MoveTo failed (${response.status}): ${errorText}`);
        return false;
      }
    } catch (error) {
      console.log(`[FileSearchService] ❌ Simple MoveTo exception:`, error);
      return false;
    }
  }

  /**
   * Try modern Move API with correct parameters
   */
  private async tryModernMoveAPI(originalPath: string, newPath: string, requestDigest: string): Promise<void> {
    console.log(`[FileSearchService] 🔄 Trying modern SP.MoveCopyUtil.MoveFileByPath API`);
    
    const webUrl = this.context.pageContext.web.absoluteUrl;
    const moveApiUrl = `${webUrl}/_api/SP.MoveCopyUtil.MoveFileByPath`;
    
    const movePayload = {
      srcPath: {
        __metadata: { type: "SP.ResourcePath" },
        DecodedUrl: originalPath
      },
      destPath: {
        __metadata: { type: "SP.ResourcePath" },
        DecodedUrl: newPath
      },
      options: {
        __metadata: { type: "SP.MoveCopyOptions" },
        KeepBoth: false,
        ResetAuthorAndCreatedOnCopy: false,
        ShouldBypassSharedLocks: true
      }
    };
    
    console.log(`[FileSearchService] 📞 Modern API payload:`, JSON.stringify(movePayload, null, 2));
    
    const response = await fetch(moveApiUrl, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': requestDigest
      },
      body: JSON.stringify(movePayload)
    });
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error(`[FileSearchService] ❌ Modern API failed (${response.status}):`, errorText);
      throw new Error(`Modern API failed: HTTP ${response.status}: ${errorText}`);
    }
    
    console.log(`[FileSearchService] ✅ Modern API succeeded`);
  }

  /**
   * Rename a single file using SharePoint REST API - UPDATED: Skip if target exists
   */
  private async renameSingleFile(originalPath: string, newPath: string, requestDigest: string): Promise<void> {
    console.log(`[FileSearchService] 🔄 CHECKING AND RENAMING file:`);
    console.log(`  From: "${originalPath}"`);
    console.log(`  To: "${newPath}"`);
    
    const cleanOriginalPath = this.cleanSharePointPath(originalPath);
    const cleanNewPath = this.cleanSharePointPath(newPath);
    
    console.log(`[FileSearchService] 🧹 Cleaned paths:`);
    console.log(`  Clean from: "${cleanOriginalPath}"`);
    console.log(`  Clean to: "${cleanNewPath}"`);
    
    try {
      // Check if file with new name already exists
      const checkResult = await this.checkFileExists(cleanNewPath);
      if (checkResult.exists) {
        // Don't create unique name, throw special error instead
        const message = `File already exists with target name. Skipping rename to avoid overwrite.`;
        console.log(`[FileSearchService] ⚠️ TARGET FILE EXISTS: "${cleanNewPath}"`);
        console.log(`[FileSearchService] ⏭️ SKIPPING RENAME to avoid overwrite`);
        throw new Error(`FILE_ALREADY_EXISTS: ${message}`);
      }
      
      console.log(`[FileSearchService] ✅ Target path is free, proceeding with rename...`);
      
      // Try simple MoveTo API first
      const success = await this.trySimpleMoveTo(cleanOriginalPath, cleanNewPath, requestDigest);
      if (success) {
        console.log(`[FileSearchService] ✅ File renamed successfully using simple MoveTo`);
        return;
      }
      
      // If simple doesn't work, try modern API
      await this.tryModernMoveAPI(cleanOriginalPath, cleanNewPath, requestDigest);
      console.log(`[FileSearchService] ✅ File renamed successfully using modern API`);
      
    } catch (error) {
      // Check if this is a "file already exists" error
      if (error instanceof Error && error.message.startsWith('FILE_ALREADY_EXISTS:')) {
        // Re-throw the special error as is
        throw error;
      }
      
      console.error(`[FileSearchService] ❌ All rename methods failed:`, error);
      throw error;
    }
  }

  /**
   * Get SharePoint request digest for authenticated requests
   */
  private async getRequestDigest(): Promise<string> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const response = await fetch(`${webUrl}/_api/contextinfo`, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (response.ok) {
        const data = await response.json();
        return data.d.GetContextWebInformation.FormDigestValue;
      } else {
        throw new Error(`Failed to get request digest: ${response.status}`);
      }
    } catch (error) {
      console.error('[FileSearchService] Error getting request digest:', error);
      throw error;
    }
  }

  /**
   * Build directory SharePoint path
   */
  private buildDirectoryPath(relativePath: string, basePath: string): string {
    const normalizedRelative = relativePath.replace(/\\/g, '/');
    const fullPath = `${basePath}/${normalizedRelative}`;
    return fullPath.replace(/\/+/g, '/').replace(/\/$/, '');
  }

  /**
   * Helper method to create timeout promise
   */
  private createTimeoutPromise<T>(timeoutMs: number, errorMessage: string | T): Promise<T> {
    return new Promise((resolve, reject) => { // FIXED: parameter names
      setTimeout(() => {
        if (typeof errorMessage === 'string') {
          reject(new Error(errorMessage));
        } else {
          // For boolean returns, resolve with the fallback value
          resolve(errorMessage);
        }
      }, timeoutMs);
    });
  }

  public cancelSearch(): void {
    console.log('[FileSearchService] Cancelling file search...');
    this.isCancelled = true;
    this.currentSearchId = undefined;
  }

  public isSearchActive(): boolean {
    return this.currentSearchId !== undefined && !this.isCancelled;
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
      
      const fileFound = files.some((file: ISharePointFolder) => // FIXED: specific type instead of any
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

  public async getFileDetails(filePath: string): Promise<ISharePointFolder | undefined> { // FIXED: specific type instead of any
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

      return undefined;
    } catch (error) {
      console.error('[FileSearchService] Error getting file details:', error);
      return undefined;
    }
  }

  public getFileNameFromRow(row: IRenameTableRow): string {
    const fileName = String(row.cells.custom_0?.value || ''); // FIXED: dot notation
    console.log(`[FileSearchService] getFileNameFromRow for row ${row.rowIndex}: "${fileName}"`);
    return fileName;
  }
}