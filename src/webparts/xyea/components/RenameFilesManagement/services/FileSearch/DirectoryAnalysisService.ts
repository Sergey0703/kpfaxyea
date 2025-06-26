// src/webparts/xyea/components/RenameFilesManagement/services/FileSearch/DirectoryAnalysisService.ts

import { 
  IRenameTableRow, 
  SearchStage, 
  ISearchProgress, 
  IDirectoryAnalysis, 
  ISearchPlan,
  SearchProgressHelper,
  DirectoryStatus,
  DirectoryStatusCallback
} from '../../types/RenameFilesTypes';
import { SharePointFolderService } from '../SharePointFolderService';
import { ExcelFileProcessor } from '../ExcelFileProcessor';
import { FileSearchConfigService } from './FileSearchConfigService';

export class DirectoryAnalysisService {
  private folderService: SharePointFolderService;
  private excelProcessor: ExcelFileProcessor;
  private configService: FileSearchConfigService;

  constructor(
    folderService: SharePointFolderService,
    excelProcessor: ExcelFileProcessor,
    configService: FileSearchConfigService
  ) {
    this.folderService = folderService;
    this.excelProcessor = excelProcessor;
    this.configService = configService;
  }

  /**
   * STAGE 1: Analyze directories with timeout protection
   */
  public async executeStage1_AnalyzeDirectories(
    rows: IRenameTableRow[],
    baseFolderPath: string,
    currentProgress: ISearchProgress,
    statusCallback?: (progress: ISearchProgress) => void,
    isCancelled?: () => boolean
  ): Promise<ISearchProgress> {
    
    console.log('[DirectoryAnalysisService] STAGE 1: Analyzing directories with timeout protection...');
    
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
    const validRows = 0;

    rows.forEach(row => {
      const directoryCell = row.cells.custom_1;
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
        directoryToRows.get(directoryPath)?.push(row.rowIndex);
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
      if (isCancelled?.()) break;

      const rowIndexes = directoryToRows.get(directoryPath) || [];
      const fullSharePointPath = this.folderService.getFullDirectoryPath(directoryPath, baseFolderPath);
      
      directoryGroups.push({
        directoryPath,
        normalizedPath: this.configService.normalizePath(directoryPath),
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

      await this.configService.delay(5);
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
  public async executeStage2_CheckDirectoryExistence_OPTIMIZED(
    currentProgress: ISearchProgress,
    statusCallback?: (progress: ISearchProgress) => void,
    directoryStatusCallback?: DirectoryStatusCallback,
    isCancelled?: () => boolean
  ): Promise<ISearchProgress> {
    
    console.log('[DirectoryAnalysisService] 🚀 OPTIMIZED STAGE 2: Checking directories ONCE, not per row...');
    
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

    console.log(`[DirectoryAnalysisService] 📊 EFFICIENCY GAINED: Checking ${searchPlan.directoryGroups.length} directories instead of ${searchPlan.totalRows} rows!`);
    console.log(`[DirectoryAnalysisService] 🎯 API calls reduced from ${searchPlan.totalRows} to ${searchPlan.directoryGroups.length} (${Math.round(searchPlan.totalRows / searchPlan.directoryGroups.length)}x improvement)`);

    // Initialize all rows as 'checking' (bulk operation)
    if (directoryStatusCallback) {
      console.log('[DirectoryAnalysisService] 🔄 Bulk initializing all rows as "checking"...');
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
        this.configService.createTimeoutPromise(this.configService.FOLDER_LOAD_TIMEOUT, 'Folder loading timeout')
      ]);

    } catch (error) {
      console.warn('[DirectoryAnalysisService] Folder loading failed or timed out:', error);
      // Continue with basic directory checks
    }

    // OPTIMIZED: Check each UNIQUE directory ONCE (not per row)
    let checkedDirectories = 0;
    let existingDirectories = 0;

    for (const dirGroup of searchPlan.directoryGroups) {
      if (isCancelled?.()) break;

      console.log(`[DirectoryAnalysisService] 🔍 Checking UNIQUE directory ${checkedDirectories + 1}/${searchPlan.directoryGroups.length}:`);
      console.log(`[DirectoryAnalysisService] 📁 Path: "${dirGroup.directoryPath}"`);
      console.log(`[DirectoryAnalysisService] 📊 Will update ${dirGroup.rowIndexes.length} rows with result`);

      try {
        // ONE API call per directory (not per row)
        const checkPromise = Promise.resolve(
          this.folderService.checkDirectoryExists(dirGroup.fullSharePointPath)
        );
        
        dirGroup.exists = await Promise.race([
          checkPromise,
          this.configService.createTimeoutPromise(this.configService.DIRECTORY_CHECK_TIMEOUT, false)
        ]) as boolean;
        
        if (dirGroup.exists) {
          existingDirectories++;
        }

        // BULK UPDATE: Update ALL rows for this directory with the SAME result
        if (directoryStatusCallback) {
          const directoryStatus: DirectoryStatus = dirGroup.exists ? 'exists' : 'not-exists';
          
          console.log(`[DirectoryAnalysisService] 📂 Directory "${dirGroup.directoryPath}" -> ${directoryStatus}`);
          console.log(`[DirectoryAnalysisService] 🔄 Bulk updating ${dirGroup.rowIndexes.length} rows with status: ${directoryStatus}`);
          
          // Bulk callback
          directoryStatusCallback(dirGroup.rowIndexes, directoryStatus);
          
          console.log(`[DirectoryAnalysisService] ✅ Updated rows: ${dirGroup.rowIndexes.slice(0, 5).map(r => r + 1).join(', ')}${dirGroup.rowIndexes.length > 5 ? `, ... and ${dirGroup.rowIndexes.length - 5} more` : ''}`);
        }

      } catch (error) {
        console.warn(`[DirectoryAnalysisService] Directory check failed for ${dirGroup.directoryPath}:`, error);
        dirGroup.exists = false;

        // BULK UPDATE: Mark all rows for this directory as error
        if (directoryStatusCallback) {
          console.log(`[DirectoryAnalysisService] 📂⚠️ Directory "${dirGroup.directoryPath}" -> error`);
          console.log(`[DirectoryAnalysisService] 🔄 Bulk updating ${dirGroup.rowIndexes.length} rows with status: error`);
          
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
      await this.configService.delay(50);
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
    
    console.log(`[DirectoryAnalysisService] 🎯 STAGE 2 OPTIMIZATION COMPLETE:`);
    console.log(`[DirectoryAnalysisService] ✅ Directories checked: ${checkedDirectories} (instead of ${searchPlan.totalRows} rows)`);
    console.log(`[DirectoryAnalysisService] 📈 Performance improvement: ${Math.round(searchPlan.totalRows / checkedDirectories)}x fewer API calls`);
    console.log(`[DirectoryAnalysisService] 📊 Results: ${existingDirectories} exist, ${searchPlan.totalDirectories - existingDirectories} missing`);
    
    return progress;
  }
}