// src/webparts/xyea/components/RenameFilesManagement/services/FileSearch/FileRenameService.ts

import { 
  IRenameTableRow, 
  FileSearchStatus
} from '../../types/RenameFilesTypes';
import { FileSearchConfigService } from './FileSearchConfigService';
import { SharePointFolderService } from '../SharePointFolderService'; // NEW: Add folder service

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

interface IRenameFileInfo {
  rowIndex: number;
  originalFileName: string;
  staffID: string;
  directoryPath: string;
  fullOriginalPath: string;
  fullNewPath: string;
  newFileName: string;
}

interface IRenameResult {
  success: number;
  errors: number;
  skipped: number;
  errorDetails: string[];
  skippedDetails: string[];
}

interface IRenameProgress {
  current: number;
  total: number;
  fileName: string;
  success: number;
  errors: number;
  skipped: number;
}

type RenameStatus = 'renaming' | 'renamed' | 'error' | 'skipped';

export class FileRenameService {
  private context: IWebPartContext;
  private configService: FileSearchConfigService;
  private folderService: SharePointFolderService; // NEW: Add folder service

  constructor(context: IWebPartContext, configService: FileSearchConfigService) {
    this.context = context;
    this.configService = configService;
    this.folderService = new SharePointFolderService(context); // NEW: Initialize folder service
  }

  /**
   * Rename found files with staffID prefix - OPTIMIZED: Batch existence checks
   */
  public async renameFoundFiles(
    rows: IRenameTableRow[],
    fileSearchResults: { [rowIndex: number]: FileSearchStatus },
    baseFolderPath: string,
    progressCallback: (rowIndex: number, status: RenameStatus) => void,
    statusCallback?: (progress: IRenameProgress) => void,
    isCancelled?: () => boolean
  ): Promise<IRenameResult> {
    
    console.log(`[FileRenameService] 🏷️ STARTING OPTIMIZED FILE RENAME WITH 634 REPLACEMENT LOGIC`);
    
    // Prepare files for renaming
    const filesToRename = this.prepareFilesForRename(rows, fileSearchResults, baseFolderPath);

    console.log(`[FileRenameService] 📊 Prepared ${filesToRename.length} files for renaming`);

    if (filesToRename.length === 0) {
      console.warn(`[FileRenameService] ⚠️ No files prepared for renaming`);
      return { 
        success: 0, 
        errors: 0, 
        skipped: 0, 
        errorDetails: ['No files prepared for renaming'], 
        skippedDetails: [] 
      };
    }

    // NEW: Batch load all directories before starting rename operations
    await this.preloadDirectoriesForExistenceChecks(filesToRename);

    return this.executeRenameOperations(filesToRename, progressCallback, statusCallback, isCancelled);
  }

  /**
   * NEW: Preload all directories for batch existence checks
   */
  private async preloadDirectoriesForExistenceChecks(filesToRename: IRenameFileInfo[]): Promise<void> {
    console.log(`[FileRenameService] 🚀 OPTIMIZATION: Preloading directories for batch existence checks`);
    
    // Get unique directories from all files to rename
    const uniqueDirectories = this.getUniqueDirectories(filesToRename);
    
    console.log(`[FileRenameService] 📁 Found ${uniqueDirectories.length} unique directories to preload`);
    console.log(`[FileRenameService] 📈 Performance gain: ${filesToRename.length} API calls → ${uniqueDirectories.length} API calls`);
    
    try {
      // Batch load all directory contents
      await this.folderService.batchLoadDirectoryContents(
        uniqueDirectories,
        (loaded: number, total: number, currentPath: string) => {
          console.log(`[FileRenameService] 📦 Preloading directories: ${loaded}/${total} - ${currentPath}`);
        }
      );
      
      console.log(`[FileRenameService] ✅ Successfully preloaded ${uniqueDirectories.length} directories`);
      
      // Log performance statistics
      const stats = this.folderService.getPerformanceStats();
      console.log(`[FileRenameService] 📊 Folder service performance:`, {
        directoriesCached: stats.directoriesCached,
        totalFilesCached: stats.totalFilesCached,
        batchLoadsExecuted: stats.batchLoadsExecuted
      });
      
    } catch (error) {
      console.error(`[FileRenameService] ❌ Error preloading directories:`, error);
      console.warn(`[FileRenameService] ⚠️ Continuing with individual existence checks as fallback`);
    }
  }

  /**
   * NEW: Get unique directories from files to rename
   */
  private getUniqueDirectories(filesToRename: IRenameFileInfo[]): string[] {
    const uniqueDirectories = new Set<string>();
    
    filesToRename.forEach(fileInfo => {
      const directoryPath = this.getDirectoryFromPath(fileInfo.fullNewPath);
      if (directoryPath) {
        uniqueDirectories.add(directoryPath);
      }
    });
    
    const directories = Array.from(uniqueDirectories);
    
    console.log(`[FileRenameService] 🗂️ Unique directories extracted:`, {
      totalFiles: filesToRename.length,
      uniqueDirectories: directories.length,
      optimizationRatio: Math.round(filesToRename.length / directories.length),
      sampleDirectories: directories.slice(0, 3)
    });
    
    return directories;
  }

  /**
   * NEW: Extract directory path from full file path
   */
  private getDirectoryFromPath(fullPath: string): string {
    if (!fullPath) return '';
    
    const pathParts = fullPath.split('/');
    // Remove the filename (last part) to get directory path
    const directoryParts = pathParts.slice(0, -1);
    const directoryPath = directoryParts.join('/');
    
    return this.configService.cleanSharePointPath(directoryPath);
  }

  /**
   * NEW: Extract filename from full file path
   */
  private getFileNameFromPath(fullPath: string): string {
    if (!fullPath) return '';
    
    const pathParts = fullPath.split('/');
    return pathParts[pathParts.length - 1] || '';
  }

  /**
   * Prepare files for renaming by collecting and validating data
   */
  private prepareFilesForRename(
    rows: IRenameTableRow[],
    fileSearchResults: { [rowIndex: number]: FileSearchStatus },
    baseFolderPath: string
  ): IRenameFileInfo[] {
    
    const filesToRename: IRenameFileInfo[] = [];

    rows.forEach(row => {
      const searchResult = fileSearchResults[row.rowIndex];
      
      if (searchResult === 'found') {
        const fileInfo = this.extractFileInfoFromRow(row, baseFolderPath);
        
        if (fileInfo) {
          filesToRename.push(fileInfo);
          console.log(`[FileRenameService] 📝 Prepared rename: "${fileInfo.originalFileName}" -> "${fileInfo.newFileName}"`);
        } else {
          console.warn(`[FileRenameService] ⚠️ Missing data for row ${row.rowIndex}`);
        }
      }
    });

    return filesToRename;
  }

  /**
   * Extract file information from a row
   */
  private extractFileInfoFromRow(row: IRenameTableRow, baseFolderPath: string): IRenameFileInfo | null {
    const originalFileName = String(row.cells.custom_0?.value || '').trim();
    const directoryPath = String(row.cells.custom_1?.value || '').trim();
    
    // Find staffID in different possible columns
    const staffID = this.findStaffIDInRow(row);
    
    if (!originalFileName || !staffID || !directoryPath) {
      return null;
    }

    const directorySharePointPath = this.configService.buildDirectoryPath(directoryPath, baseFolderPath);
    const fullOriginalPath = `${directorySharePointPath}/${originalFileName}`;
    
    // UPDATED: Use new filename generation logic with 634 replacement
    const newFileName = this.generateNewFileNameWith634Replacement(originalFileName, staffID);
    const fullNewPath = `${directorySharePointPath}/${newFileName}`;
    
    return {
      rowIndex: row.rowIndex,
      originalFileName,
      staffID,
      directoryPath,
      fullOriginalPath,
      fullNewPath,
      newFileName
    };
  }

  /**
   * NEW: Generate new filename with 634 replacement logic
   * If filename starts with "634", replace it with staffID
   * If filename doesn't start with "634", add staffID at the beginning
   */
  private generateNewFileNameWith634Replacement(originalFileName: string, staffID: string): string {
    const cleanStaffID = staffID.replace(/[<>:"/\\|?*]/g, '').trim();
    
    console.log(`[FileRenameService] 🔄 Generating new filename:`);
    console.log(`  Original: "${originalFileName}"`);
    console.log(`  StaffID: "${cleanStaffID}"`);
    
    // Check if filename starts with "634"
    if (originalFileName.startsWith('634')) {
      console.log(`[FileRenameService] 🔄 File starts with "634", replacing...`);
      
      // Replace "634" at the beginning with staffID
      const newFileName = cleanStaffID + originalFileName.substring(3);
      
      console.log(`[FileRenameService] ✅ 634 replacement: "${originalFileName}" -> "${newFileName}"`);
      return newFileName;
    } else {
      console.log(`[FileRenameService] 🔄 File doesn't start with "634", adding staffID prefix...`);
      
      // Check if file already starts with the staffID
      if (originalFileName.toLowerCase().startsWith(cleanStaffID.toLowerCase())) {
        console.log(`[FileRenameService] ⚠️ File already starts with staffID: "${originalFileName}"`);
        return originalFileName;
      }
      
      // Add staffID at the beginning (original behavior)
      const newFileName = `${cleanStaffID} ${originalFileName}`;
      
      console.log(`[FileRenameService] ✅ StaffID prefix added: "${originalFileName}" -> "${newFileName}"`);
      return newFileName;
    }
  }

  /**
   * Find staffID in various possible columns
   */
  private findStaffIDInRow(row: IRenameTableRow): string {
    // First, try known staffID column names
    const staffIDColumns = ['staffID', 'staffid', 'StaffID', 'staff_id', 'ID', 'id'];
    
    for (const columnName of staffIDColumns) {
      const cellValue = String(row.cells[columnName]?.value || '').trim();
      if (cellValue) {
        console.log(`[FileRenameService] 📋 Found staffID "${cellValue}" in column ${columnName} for row ${row.rowIndex}`);
        return cellValue;
      }
    }
    
    // If not found, search through Excel columns for ID-like values
    const excelColumns = Object.keys(row.cells).filter(key => key.startsWith('excel_'));
    for (const columnId of excelColumns) {
      const cellValue = String(row.cells[columnId]?.value || '').trim();
      if (cellValue && /^[0-9A-Za-z]{1,10}$/.test(cellValue)) {
        console.log(`[FileRenameService] 📋 Found staffID "${cellValue}" in column ${columnId} for row ${row.rowIndex}`);
        return cellValue;
      }
    }
    
    return '';
  }

  /**
   * Execute the actual rename operations
   */
  private async executeRenameOperations(
    filesToRename: IRenameFileInfo[],
    progressCallback: (rowIndex: number, status: RenameStatus) => void,
    statusCallback?: (progress: IRenameProgress) => void,
    isCancelled?: () => boolean
  ): Promise<IRenameResult> {
    
    let processedFiles = 0;
    let successCount = 0;
    let errorCount = 0;
    let skippedCount = 0;
    const errorDetails: string[] = [];
    const skippedDetails: string[] = [];

    try {
      let requestDigest = await this.getRequestDigest();
      const BATCH_SIZE = 1;
      
      for (let i = 0; i < filesToRename.length; i += BATCH_SIZE) {
        if (isCancelled?.()) {
          console.log('[FileRenameService] ❌ Rename operation cancelled');
          break;
        }

        // Refresh Request Digest every 100 files for long operations
        if (processedFiles % 100 === 0 && processedFiles > 0) {
          console.log(`[FileRenameService] 🔄 Refreshing request digest after ${processedFiles} files...`);
          requestDigest = await this.getRequestDigest();
          console.log(`[FileRenameService] ✅ Request digest refreshed`);
        }

        const batch = filesToRename.slice(i, i + BATCH_SIZE);
        console.log(`[FileRenameService] 📦 Processing file ${i + 1}/${filesToRename.length}`);

        for (const fileInfo of batch) {
          if (isCancelled?.()) break;

          const result = await this.renameSingleFileWithHandling(fileInfo, requestDigest, progressCallback);
          
          // Update counters based on result
          switch (result.status) {
            case 'success':
              successCount++;
              break;
            case 'error':
              errorCount++;
              errorDetails.push(result.message);
              break;
            case 'skipped':
              skippedCount++;
              skippedDetails.push(result.message);
              break;
          }
          
          processedFiles++;

          // Update status callback
          statusCallback?.({
            current: processedFiles,
            total: filesToRename.length,
            fileName: fileInfo.originalFileName,
            success: successCount,
            errors: errorCount,
            skipped: skippedCount
          });

          // Delay between renames to avoid overwhelming SharePoint
          await this.configService.delay(this.configService.DELAY_BETWEEN_RENAMES);
        }
      }

      this.logRenameResults(filesToRename.length, successCount, errorCount, skippedCount, errorDetails, skippedDetails);

      return { 
        success: successCount, 
        errors: errorCount, 
        skipped: skippedCount,
        errorDetails, 
        skippedDetails
      };

    } catch (error) {
      console.error('[FileRenameService] ❌ Critical error in rename operation:', error);
      
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
   * Rename a single file with comprehensive error handling
   */
  private async renameSingleFileWithHandling(
    fileInfo: IRenameFileInfo,
    requestDigest: string,
    progressCallback: (rowIndex: number, status: RenameStatus) => void
  ): Promise<{ status: 'success' | 'error' | 'skipped'; message: string }> {
    
    try {
      progressCallback(fileInfo.rowIndex, 'renaming');
      
      console.log(`[FileRenameService] 🔄 Processing file with 634 replacement logic:`);
      console.log(`  Original: "${fileInfo.originalFileName}"`);
      console.log(`  New: "${fileInfo.newFileName}"`);
      console.log(`  StaffID: "${fileInfo.staffID}"`);

      await this.renameSingleFile(fileInfo.fullOriginalPath, fileInfo.fullNewPath, requestDigest);
      
      progressCallback(fileInfo.rowIndex, 'renamed');
      console.log(`[FileRenameService] ✅ SUCCESS: "${fileInfo.originalFileName}" -> "${fileInfo.newFileName}"`);
      
      return {
        status: 'success',
        message: `Successfully renamed "${fileInfo.originalFileName}"`
      };
      
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      
      // Check if this is a "file already exists" error
      if (errorMessage.startsWith('FILE_ALREADY_EXISTS:')) {
        const skippedMessage = `Row ${fileInfo.rowIndex + 1} - ${fileInfo.originalFileName}: Target file already exists, skipped to avoid overwrite`;
        progressCallback(fileInfo.rowIndex, 'skipped');
        console.log(`[FileRenameService] ⏭️ SKIPPED: "${fileInfo.originalFileName}" (target exists)`);
        
        return {
          status: 'skipped',
          message: skippedMessage
        };
      } else {
        const detailedError = `Row ${fileInfo.rowIndex + 1} - ${fileInfo.originalFileName}: ${errorMessage}`;
        progressCallback(fileInfo.rowIndex, 'error');
        console.error(`[FileRenameService] ❌ ERROR: "${fileInfo.originalFileName}": ${errorMessage}`);
        
        return {
          status: 'error',
          message: detailedError
        };
      }
    }
  }

  /**
   * Log comprehensive rename results
   */
  private logRenameResults(
    totalFiles: number,
    successCount: number,
    errorCount: number,
    skippedCount: number,
    errorDetails: string[],
    skippedDetails: string[]
  ): void {
    
    console.log(`[FileRenameService] 🎯 Rename completed with 634 replacement logic:`);
    console.log(`  📊 Total files: ${totalFiles}`);
    console.log(`  ✅ Successful: ${successCount}`);
    console.log(`  ❌ Failed: ${errorCount}`);
    console.log(`  ⏭️ Skipped: ${skippedCount}`);
    console.log(`  📈 Success rate: ${totalFiles > 0 ? (successCount / totalFiles * 100).toFixed(1) + '%' : '0%'}`);

    // Log skipped files
    if (skippedDetails.length > 0) {
      console.log(`[FileRenameService] 📋 Skipped files (target already exists):`);
      skippedDetails.slice(0, 3).forEach((skipped, index) => {
        console.log(`  ${index + 1}. ${skipped}`);
      });
      if (skippedDetails.length > 3) {
        console.log(`  ... and ${skippedDetails.length - 3} more skipped files`);
      }
    }

    // Log error details
    if (errorDetails.length > 0) {
      console.error(`[FileRenameService] 📋 Error details:`);
      errorDetails.slice(0, 3).forEach((error, index) => {
        console.error(`  ${index + 1}. ${error}`);
      });
      if (errorDetails.length > 3) {
        console.error(`  ... and ${errorDetails.length - 3} more errors`);
      }
    }
  }

  /**
   * OPTIMIZED: Check if file exists using cached directory data
   */
  private async checkFileExistsOptimized(filePath: string): Promise<{ exists: boolean; error?: string }> {
    try {
      const directoryPath = this.getDirectoryFromPath(filePath);
      const fileName = this.getFileNameFromPath(filePath);
      
      console.log(`[FileRenameService] 🔍 OPTIMIZED existence check: "${fileName}" in "${directoryPath}"`);
      
      // Try to use cached data first
      if (this.folderService.isDirectoryCached(directoryPath)) {
        const exists = this.folderService.checkFileExistsInCache(directoryPath, fileName);
        console.log(`[FileRenameService] ⚡ Cache hit: "${fileName}" ${exists ? 'EXISTS' : 'NOT FOUND'}`);
        return { exists };
      }
      
      // Fallback to direct API check if not cached
      console.log(`[FileRenameService] 📞 Cache miss, using direct API check for: "${filePath}"`);
      return await this.checkFileExistsDirect(filePath);
      
    } catch (error) {
      console.log(`[FileRenameService] ⚠️ Error in optimized existence check: ${error}`);
      return { exists: false, error: String(error) };
    }
  }

  /**
   * Direct API check for file existence (fallback)
   */
  private async checkFileExistsDirect(filePath: string): Promise<{ exists: boolean; error?: string }> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const checkUrl = `${webUrl}/_api/web/getFileByServerRelativeUrl('${encodeURIComponent(filePath)}')`;
      
      console.log(`[FileRenameService] 🔍 Direct API file existence check: ${checkUrl}`);
      
      const response = await fetch(checkUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (response.ok) {
        console.log(`[FileRenameService] ✅ File exists: "${filePath}"`);
        return { exists: true };
      } else if (response.status === 404) {
        console.log(`[FileRenameService] ❌ File does not exist: "${filePath}"`);
        return { exists: false };
      } else {
        console.log(`[FileRenameService] ⚠️ Unknown status ${response.status} for file: "${filePath}"`);
        return { exists: false, error: `HTTP ${response.status}` };
      }
    } catch (error) {
      console.log(`[FileRenameService] ⚠️ Error checking file existence: ${error}`);
      return { exists: false, error: String(error) };
    }
  }

  /**
   * Try simple MoveTo API with proper encoding
   */
  private async trySimpleMoveTo(originalPath: string, newPath: string, requestDigest: string): Promise<boolean> {
    try {
      console.log(`[FileRenameService] 🔄 Trying simple MoveTo API`);
      
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const moveToUrl = `${webUrl}/_api/web/getFileByServerRelativeUrl('${originalPath}')/MoveTo(newurl='${newPath}',flags=1)`;
      
      console.log(`[FileRenameService] 📞 Simple MoveTo URL:`, moveToUrl);
      
      const response = await fetch(moveToUrl, {
        method: 'POST',
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-RequestDigest': requestDigest
        }
      });
      
      if (response.ok) {
        console.log(`[FileRenameService] ✅ Simple MoveTo succeeded`);
        return true;
      } else {
        const errorText = await response.text();
        console.log(`[FileRenameService] ❌ Simple MoveTo failed (${response.status}): ${errorText}`);
        return false;
      }
    } catch (error) {
      console.log(`[FileRenameService] ❌ Simple MoveTo exception:`, error);
      return false;
    }
  }

  /**
   * Try modern Move API with correct parameters
   */
  private async tryModernMoveAPI(originalPath: string, newPath: string, requestDigest: string): Promise<void> {
    console.log(`[FileRenameService] 🔄 Trying modern SP.MoveCopyUtil.MoveFileByPath API`);
    
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
    
    console.log(`[FileRenameService] 📞 Modern API payload:`, JSON.stringify(movePayload, null, 2));
    
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
      console.error(`[FileRenameService] ❌ Modern API failed (${response.status}):`, errorText);
      throw new Error(`Modern API failed: HTTP ${response.status}: ${errorText}`);
    }
    
    console.log(`[FileRenameService] ✅ Modern API succeeded`);
  }

  /**
   * Rename a single file using SharePoint REST API - OPTIMIZED: Use cached existence checks
   */
  private async renameSingleFile(originalPath: string, newPath: string, requestDigest: string): Promise<void> {
    console.log(`[FileRenameService] 🔄 OPTIMIZED CHECKING AND RENAMING file:`);
    console.log(`  From: "${originalPath}"`);
    console.log(`  To: "${newPath}"`);
    
    const cleanOriginalPath = this.configService.cleanSharePointPath(originalPath);
    const cleanNewPath = this.configService.cleanSharePointPath(newPath);
    
    console.log(`[FileRenameService] 🧹 Cleaned paths:`);
    console.log(`  Clean from: "${cleanOriginalPath}"`);
    console.log(`  Clean to: "${cleanNewPath}"`);
    
    try {
      // OPTIMIZED: Check if file with new name already exists using cached data
      const checkResult = await this.checkFileExistsOptimized(cleanNewPath);
      if (checkResult.exists) {
        // Don't create unique name, throw special error instead
        const message = `File already exists with target name. Skipping rename to avoid overwrite.`;
        console.log(`[FileRenameService] ⚠️ TARGET FILE EXISTS (cached check): "${cleanNewPath}"`);
        console.log(`[FileRenameService] ⏭️ SKIPPING RENAME to avoid overwrite`);
        throw new Error(`FILE_ALREADY_EXISTS: ${message}`);
      }
      
      console.log(`[FileRenameService] ✅ Target path is free (cached check), proceeding with rename...`);
      
      // Try simple MoveTo API first
      const success = await this.trySimpleMoveTo(cleanOriginalPath, cleanNewPath, requestDigest);
      if (success) {
        console.log(`[FileRenameService] ✅ File renamed successfully using simple MoveTo`);
        return;
      }
      
      // If simple doesn't work, try modern API
      await this.tryModernMoveAPI(cleanOriginalPath, cleanNewPath, requestDigest);
      console.log(`[FileRenameService] ✅ File renamed successfully using modern API`);
      
    } catch (error) {
      // Check if this is a "file already exists" error
      if (error instanceof Error && error.message.startsWith('FILE_ALREADY_EXISTS:')) {
        // Re-throw the special error as is
        throw error;
      }
      
      console.error(`[FileRenameService] ❌ All rename methods failed:`, error);
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
      console.error('[FileRenameService] Error getting request digest:', error);
      throw error;
    }
  }

  /**
   * Validate rename operation prerequisites
   */
  public validateRenamePrerequisites(
    rows: IRenameTableRow[],
    fileSearchResults: { [rowIndex: number]: FileSearchStatus }
  ): { isValid: boolean; errors: string[]; warnings: string[]; foundFilesCount: number } {
    
    const errors: string[] = [];
    const warnings: string[] = [];
    let foundFilesCount = 0;
    let missingStaffIDCount = 0;
    let missingFileNameCount = 0;

    rows.forEach(row => {
      const searchResult = fileSearchResults[row.rowIndex];
      
      if (searchResult === 'found') {
        foundFilesCount++;
        
        const originalFileName = String(row.cells.custom_0?.value || '').trim();
        const staffID = this.findStaffIDInRow(row);
        
        if (!originalFileName) {
          missingFileNameCount++;
          errors.push(`Row ${row.rowIndex + 1}: Missing filename`);
        }
        
        if (!staffID) {
          missingStaffIDCount++;
          errors.push(`Row ${row.rowIndex + 1}: Missing Staff ID`);
        }
        
        // Validate filename
        if (originalFileName && !this.configService.isValidFileName(originalFileName)) {
          warnings.push(`Row ${row.rowIndex + 1}: Filename "${originalFileName}" contains invalid characters`);
        }
      }
    });

    if (foundFilesCount === 0) {
      errors.push('No files found to rename. Please run file search first.');
    }

    if (missingStaffIDCount > 0) {
      warnings.push(`${missingStaffIDCount} files are missing Staff ID and will be skipped`);
    }

    if (missingFileNameCount > 0) {
      warnings.push(`${missingFileNameCount} files are missing filename and will be skipped`);
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
      foundFilesCount
    };
  }

  /**
   * Get rename statistics for reporting
   */
  public getRenameStatistics(result: IRenameResult): {
    totalAttempted: number;
    successRate: number;
    errorRate: number;
    skipRate: number;
    summary: string;
  } {
    
    const totalAttempted = result.success + result.errors + result.skipped;
    const successRate = totalAttempted > 0 ? (result.success / totalAttempted) * 100 : 0;
    const errorRate = totalAttempted > 0 ? (result.errors / totalAttempted) * 100 : 0;
    const skipRate = totalAttempted > 0 ? (result.skipped / totalAttempted) * 100 : 0;

    let summary = `Renamed ${result.success} files successfully`;
    if (result.skipped > 0) {
      summary += `, skipped ${result.skipped} files`;
    }
    if (result.errors > 0) {
      summary += `, ${result.errors} errors occurred`;
    }

    return {
      totalAttempted,
      successRate: Math.round(successRate * 100) / 100,
      errorRate: Math.round(errorRate * 100) / 100,
      skipRate: Math.round(skipRate * 100) / 100,
      summary
    };
  }
}