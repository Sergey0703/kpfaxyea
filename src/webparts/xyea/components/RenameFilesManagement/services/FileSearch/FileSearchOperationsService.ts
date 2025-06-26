// src/webparts/xyea/components/RenameFilesManagement/services/FileSearch/FileSearchOperationsService.ts

import { 
  IRenameTableRow, 
  SearchStage, 
  ISearchProgress, 
  ISearchPlan,
  SearchProgressHelper,
  FileSearchStatus,
  FileSearchResultCallback
} from '../../types/RenameFilesTypes';
import { SharePointFolderService } from '../SharePointFolderService';
import { FileSearchConfigService } from './FileSearchConfigService';

interface ISharePointFolder {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  TimeCreated: string;
  TimeLastModified: string;
}

export class FileSearchOperationsService {
  private folderService: SharePointFolderService;
  private configService: FileSearchConfigService;

  constructor(
    folderService: SharePointFolderService,
    configService: FileSearchConfigService
  ) {
    this.folderService = folderService;
    this.configService = configService;
  }

  /**
   * UPDATED: OPTIMIZED STAGE 3: Search files with CORRECTED LOGIC and NEW status handling
   */
  public async executeOptimizedStage3_SearchFiles(
    currentProgress: ISearchProgress,
    rows: IRenameTableRow[],
    progressCallback: FileSearchResultCallback,
    statusCallback?: (progress: ISearchProgress) => void,
    isCancelled?: () => boolean
  ): Promise<{ [rowIndex: number]: FileSearchStatus }> {
    
    console.log('[FileSearchOperationsService] 🚀 OPTIMIZED STAGE 3 with NEW STATUS LOGIC...');
    
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

    const results: { [rowIndex: number]: FileSearchStatus } = {};

    // STEP 1: Initialize ALL rows with appropriate status FIRST
    this.initializeAllRowsWithCorrectStatus(rows, searchPlan, results, progressCallback);

    // STEP 2: Build directory-to-files mapping for EXISTING directories only
    const directoryToFilesMap = this.buildDirectoryToFilesMap(rows, searchPlan);
    
    console.log(`[FileSearchOperationsService] 📊 Built directory mapping:`);
    Object.entries(directoryToFilesMap).forEach(([dir, files]) => {
      console.log(`  📁 "${dir}" -> ${files.length} files to search`);
    });

    let processedFiles = 0;
    let foundFiles = 0;
    const directories = Object.keys(directoryToFilesMap);
    const totalFilesToSearch = Object.values(directoryToFilesMap).reduce((sum, files) => sum + files.length, 0);

    console.log(`[FileSearchOperationsService] 🎯 STARTING SEARCH: ${totalFilesToSearch} files in ${directories.length} EXISTING directories`);

    // STEP 3: Process each EXISTING directory with ONE API call
    for (let dirIndex = 0; dirIndex < directories.length; dirIndex++) {
      const directoryPath = directories[dirIndex];
      const filesFromExcel = directoryToFilesMap[directoryPath];
      
      if (isCancelled?.()) break;

      console.log(`[FileSearchOperationsService] 🔍 DIRECTORY ${dirIndex + 1}/${directories.length}: "${directoryPath}"`);
      console.log(`[FileSearchOperationsService] 📋 Looking for ${filesFromExcel.length} Excel files in this EXISTING directory`);

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
        console.log(`[FileSearchOperationsService] 📞 API call: getFolderContents("${directoryPath}")`);
        const startTime = Date.now();
        
        const adaptiveTimeout = this.configService.calculateTimeout(filesFromExcel.length);
        const folderContentsPromise = this.folderService.getFolderContents(directoryPath);
        const folderContents = await Promise.race([
          folderContentsPromise,
          this.configService.createTimeoutPromise(adaptiveTimeout, { files: [], folders: [] })
        ]) as {files: ISharePointFolder[], folders: ISharePointFolder[]};
        
        const endTime = Date.now();
        console.log(`[FileSearchOperationsService] ✅ API response received in ${endTime - startTime}ms`);
        console.log(`[FileSearchOperationsService] 📄 SharePoint files found: ${folderContents.files.length}`);
        console.log(`[FileSearchOperationsService] 📁 SharePoint folders found: ${folderContents.folders.length}`);

        // IMPROVED: Handle empty directories gracefully
        if (folderContents.files.length === 0) {
          console.log(`[FileSearchOperationsService] ⚠️ Directory is empty or doesn't exist: "${directoryPath}"`);
          console.log(`[FileSearchOperationsService] 📝 Marking all ${filesFromExcel.length} files as NOT FOUND`);
          
          // Mark all files in this directory as not found
          filesFromExcel.forEach(excelFile => {
            if (!isCancelled?.()) {
              results[excelFile.rowIndex] = 'not-found';
              progressCallback(excelFile.rowIndex, 'not-found');
              processedFiles++;
            }
          });
          
          console.log(`[FileSearchOperationsService] 📁 DIRECTORY SUMMARY "${directoryPath}": 0/${filesFromExcel.length} files found (empty directory)`);
          continue;
        }

        // Show sample of SharePoint files
        console.log(`[FileSearchOperationsService] 📋 Sample SharePoint files:`, 
          folderContents.files.slice(0, 5).map(f => `"${f.Name}"`).join(', ')
        );

        // Show sample of Excel files we're looking for
        console.log(`[FileSearchOperationsService] 🔍 Sample Excel files to find:`, 
          filesFromExcel.slice(0, 5).map(f => `"${f.fileName}"`).join(', ')
        );

        // Create SharePoint files map (case-insensitive)
        const sharePointFilesMap = new Map<string, ISharePointFolder>();
        folderContents.files.forEach(file => {
          sharePointFilesMap.set(file.Name.toLowerCase(), file);
        });

        console.log(`[FileSearchOperationsService] 🗂️ Created SharePoint files lookup map: ${sharePointFilesMap.size} entries`);

        // CHECK each Excel file against SharePoint files with DETAILED LOGGING
        let directoryFoundCount = 0;
        const BATCH_SIZE = 20; // Process in batches of 20 for logging

        for (let fileIndex = 0; fileIndex < filesFromExcel.length; fileIndex++) {
          const excelFile = filesFromExcel[fileIndex];
          
          if (isCancelled?.()) break;

          const fileExists = sharePointFilesMap.has(excelFile.fileName.toLowerCase());
          
          // FIXED: Use atomic update function to completely avoid race condition
          const updateFileResult = (rowIndex: number, status: FileSearchStatus): void => {
            results[rowIndex] = status;
            progressCallback(rowIndex, status);
          };
          
          // Determine result and update atomically
          if (fileExists) {
            updateFileResult(excelFile.rowIndex, 'found');
          } else {
            updateFileResult(excelFile.rowIndex, 'not-found');
          }
          
          if (fileExists) {
            foundFiles++;
            directoryFoundCount++;
            console.log(`[FileSearchOperationsService] ✅ FOUND ${foundFiles}: "${excelFile.fileName}" (row ${excelFile.rowIndex + 1})`);
          } else {
            console.log(`[FileSearchOperationsService] ❌ NOT FOUND: "${excelFile.fileName}" (row ${excelFile.rowIndex + 1})`);
          }
          
          processedFiles++;

          // Batch progress logging
          if ((fileIndex + 1) % BATCH_SIZE === 0 || fileIndex === filesFromExcel.length - 1) {
            console.log(`[FileSearchOperationsService] 📦 BATCH PROGRESS: Processed ${fileIndex + 1}/${filesFromExcel.length} files in this directory`);
            console.log(`[FileSearchOperationsService] 📊 Current totals: ${foundFiles} found out of ${processedFiles} processed`);
            
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
            await this.configService.delay(50);
          }
        }

        console.log(`[FileSearchOperationsService] 📁 DIRECTORY SUMMARY "${directoryPath}":`);
        console.log(`  ✅ Found: ${directoryFoundCount}/${filesFromExcel.length} files`);
        console.log(`  📊 Success rate: ${filesFromExcel.length > 0 ? (directoryFoundCount / filesFromExcel.length * 100).toFixed(1) + '%' : '0%'}`);

      } catch (error) {
        console.error(`[FileSearchOperationsService] Error type: ${error?.constructor?.name || 'Unknown'}`);
        console.error(`[FileSearchOperationsService] Error message: ${error instanceof Error ? error.message : String(error)}`);
        
        // IMPROVED: Better error handling for non-existent directories
        if (error instanceof Error && (error.message.includes('404') || error.message.includes('Not Found'))) {
          console.log(`[FileSearchOperationsService] 📝 Directory doesn't exist, marking ${filesFromExcel.length} files as NOT FOUND`);
        } else {
          console.log(`[FileSearchOperationsService] 📝 API error, marking ${filesFromExcel.length} files as NOT FOUND`);
        }
        
        // Mark all files in this directory as not found
        filesFromExcel.forEach(excelFile => {
          if (!isCancelled?.()) {
            const notFoundResult: FileSearchStatus = 'not-found';
            results[excelFile.rowIndex] = notFoundResult;
            progressCallback(excelFile.rowIndex, notFoundResult);
            processedFiles++;
          }
        });
        
        console.log(`[FileSearchOperationsService] 📁 DIRECTORY SUMMARY "${directoryPath}": 0/${filesFromExcel.length} files found (error/not exist)`);
      }

      // Delay between directories to avoid throttling
      await this.configService.delay(this.configService.DELAY_BETWEEN_BATCHES);
      
      console.log(`[FileSearchOperationsService] 📊 OVERALL PROGRESS: ${processedFiles}/${totalFilesToSearch} files searched, ${foundFiles} found`);
      console.log(`[FileSearchOperationsService] ➡️ Moving to next directory...\n`);
    }

    console.log(`[FileSearchOperationsService] 🎯 OPTIMIZED SEARCH COMPLETED:`);
    console.log(`  📊 Files processed: ${processedFiles}/${totalFilesToSearch}`);
    console.log(`  ✅ Files found: ${foundFiles}`);
    console.log(`  📈 Success rate: ${processedFiles > 0 ? (foundFiles / processedFiles * 100).toFixed(1) + '%' : '0%'}`);
    console.log(`  🏗️ API calls made: ${directories.length} (instead of total files)`);
    console.log(`  ⚡ Performance improvement: ${totalFilesToSearch > 0 ? Math.round(totalFilesToSearch / directories.length) : 0}x fewer API calls`);

    // Mark completion and transition to COMPLETED stage
    if (!isCancelled?.()) {
      const finalProgress = SearchProgressHelper.transitionToStage(
        progress,
        SearchStage.COMPLETED,
        {
          currentFileName: 'File search completed successfully',
          overallProgress: 100,
          filesFound: foundFiles,
          filesSearched: processedFiles,
          stageProgress: 100
        }
      );
      
      console.log(`[FileSearchOperationsService] 🏁 Transitioning to COMPLETED stage with ${foundFiles} files found`);
      statusCallback?.(finalProgress);
    }

    return results;
  }

  /**
   * Initialize all rows with correct status based on directory existence
   */
  private initializeAllRowsWithCorrectStatus(
    rows: IRenameTableRow[],
    searchPlan: ISearchPlan,
    results: { [rowIndex: number]: FileSearchStatus },
    progressCallback: FileSearchResultCallback
  ): void {
    console.log('[FileSearchOperationsService] 📋 Initializing all rows with correct status...');
    
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
        console.log(`[FileSearchOperationsService] 📁❌ Row ${row.rowIndex + 1}: directory-missing (${directoryPath})`);
      } else {
        // Files in existing directories get 'searching' status (will be updated during search)
        results[row.rowIndex] = 'searching';
        progressCallback(row.rowIndex, 'searching');
        searchingCount++;
      }
    });
    
    console.log(`[FileSearchOperationsService] ✅ Status initialization complete:`);
    console.log(`  📁❌ Directory missing: ${directoryMissingCount} files`);
    console.log(`  🔍 Ready to search: ${searchingCount} files`);
  }

  /**
   * Helper method to get directory path from row
   */
  private getDirectoryFromRow(row: IRenameTableRow): string {
    const directoryCell = row.cells.custom_1;
    if (directoryCell && directoryCell.value) {
      return String(directoryCell.value).trim();
    }
    // Fallback would require ExcelFileProcessor, but we'll keep it simple for now
    return '';
  }

  /**
   * OPTIMIZATION: Build directory-to-files mapping for efficient processing
   * FIXED: Only process directories that exist (exists: true)
   */
  private buildDirectoryToFilesMap(
    rows: IRenameTableRow[], 
    searchPlan: ISearchPlan
  ): { [directoryPath: string]: Array<{ fileName: string; rowIndex: number }> } {
    
    console.log(`[FileSearchOperationsService] 🏗️ Building directory-to-files mapping...`);
    console.log(`[FileSearchOperationsService] 📊 Total directories in plan: ${searchPlan.directoryGroups.length}`);
    console.log(`[FileSearchOperationsService] ✅ Existing directories: ${searchPlan.existingDirectories}`);
    console.log(`[FileSearchOperationsService] ❌ Missing directories: ${searchPlan.missingDirectories}`);
    
    const directoryToFilesMap: { [directoryPath: string]: Array<{ fileName: string; rowIndex: number }> } = {};
    
    // CORRECTED: Only process directories that exist (exists: true)
    searchPlan.directoryGroups.forEach(dirGroup => {
      if (!dirGroup.exists) {
        console.log(`[FileSearchOperationsService] ⏭️ Skipping non-existing directory: "${dirGroup.directoryPath}" (${dirGroup.fileCount} files skipped)`);
        return; // Skip non-existing directories
      }

      console.log(`[FileSearchOperationsService] ✅ Processing existing directory: "${dirGroup.directoryPath}" (exists: ${dirGroup.exists})`);

      const filesInDirectory: Array<{ fileName: string; rowIndex: number }> = [];
      
      dirGroup.rowIndexes.forEach(rowIndex => {
        const row = rows.find(r => r.rowIndex === rowIndex);
        if (row) {
          const fileName = String(row.cells.custom_0?.value || '').trim();
          if (fileName) {
            filesInDirectory.push({ fileName, rowIndex });
          }
        }
      });

      if (filesInDirectory.length > 0) {
        directoryToFilesMap[dirGroup.fullSharePointPath] = filesInDirectory;
        console.log(`[FileSearchOperationsService] ✅ Added existing directory "${dirGroup.directoryPath}" -> ${filesInDirectory.length} files`);
        console.log(`[FileSearchOperationsService] 📋 Sample files: [${filesInDirectory.slice(0, 3).map(f => `"${f.fileName}"`).join(', ')}...]`);
      } else {
        console.log(`[FileSearchOperationsService] ⚠️ No files found for existing directory "${dirGroup.directoryPath}"`);
      }
    });

    const totalDirectories = Object.keys(directoryToFilesMap).length;
    const totalFiles = Object.values(directoryToFilesMap).reduce((sum, files) => sum + files.length, 0);
    const skippedDirectories = searchPlan.directoryGroups.length - totalDirectories;
    
    console.log(`[FileSearchOperationsService] 📊 FINAL mapping created:`);
    console.log(`[FileSearchOperationsService]   ✅ Existing directories to search: ${totalDirectories}`);
    console.log(`[FileSearchOperationsService]   ⏭️ Skipped non-existing directories: ${skippedDirectories}`);
    console.log(`[FileSearchOperationsService]   📄 Total files to search: ${totalFiles}`);
    console.log(`[FileSearchOperationsService] 📁 Directories to process: ${Object.keys(directoryToFilesMap).map(path => path.split('/').slice(-3).join('/')).join(', ')}`);
    
    return directoryToFilesMap;
  }
}