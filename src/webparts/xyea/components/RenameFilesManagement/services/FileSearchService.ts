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

  constructor(context: any) {
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
   * NEW: Rename found files with staffID prefix - ИСПРАВЛЕННАЯ ВЕРСИЯ
   */
  public async renameFoundFiles(
    rows: IRenameTableRow[],
    fileSearchResults: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    baseFolderPath: string,
    progressCallback: (rowIndex: number, status: 'renaming' | 'renamed' | 'error') => void,
    statusCallback?: (progress: { current: number; total: number; fileName: string; success: number; errors: number }) => void
  ): Promise<{ success: number; errors: number; errorDetails: string[] }> {
    
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] 🏷️ STARTING FILE RENAME (Search ID: ${searchId})`);
    
    // ИСПРАВЛЕНИЕ 1: Более безопасная подготовка файлов для переименования
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
        const originalFileName = String(row.cells['custom_0']?.value || '').trim();
        const directoryPath = String(row.cells['custom_1']?.value || '').trim();
        
        // ИСПРАВЛЕНИЕ 2: Более гибкий поиск staffID в разных колонках
        let staffID = '';
        
        // Пробуем найти staffID в разных возможных колонках
        const staffIDColumns = ['staffID', 'staffid', 'StaffID', 'staff_id', 'ID', 'id'];
        for (const columnName of staffIDColumns) {
          const cellValue = String(row.cells[columnName]?.value || '').trim();
          if (cellValue) {
            staffID = cellValue;
            break;
          }
        }
        
        // Если не нашли в именованных колонках, ищем в Excel колонках
        if (!staffID) {
          const excelColumns = Object.keys(row.cells).filter(key => key.startsWith('excel_'));
          for (const columnId of excelColumns) {
            const cellValue = String(row.cells[columnId]?.value || '').trim();
            // Проверяем, похоже ли значение на ID (число или короткая строка)
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
          
          // ИСПРАВЛЕНИЕ 3: Используем улучшенную генерацию имени
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
          console.warn(`[FileSearchService] ⚠️ Missing data for row ${row.rowIndex}:`);
          console.warn(`  fileName: "${originalFileName}"`);
          console.warn(`  staffID: "${staffID}"`);
          console.warn(`  directoryPath: "${directoryPath}"`);
          console.warn(`  Available columns:`, Object.keys(row.cells));
        }
      }
    });

    console.log(`[FileSearchService] 📊 Prepared ${filesToRename.length} files for renaming`);

    if (filesToRename.length === 0) {
      console.warn(`[FileSearchService] ⚠️ No files prepared for renaming. Check staffID column mapping.`);
      return { success: 0, errors: 0, errorDetails: ['No files prepared for renaming. Check staffID column mapping.'] };
    }

    let processedFiles = 0;
    let successCount = 0;
    let errorCount = 0;
    const errorDetails: string[] = [];

    try {
      // Get SharePoint request digest once
      const requestDigest = await this.getRequestDigest();
      
      // ИСПРАВЛЕНИЕ 4: Еще меньший batch size и больше задержек для стабильности
      const BATCH_SIZE = 2; // Очень маленький batch size
      
      for (let i = 0; i < filesToRename.length; i += BATCH_SIZE) {
        if (this.isCancelled || this.currentSearchId !== searchId) {
          console.log('[FileSearchService] ❌ Rename operation cancelled');
          break;
        }

        const batch = filesToRename.slice(i, i + BATCH_SIZE);
        console.log(`[FileSearchService] 📦 Processing batch ${Math.floor(i / BATCH_SIZE) + 1}: ${batch.length} files`);

        // Process batch sequentially (not parallel to avoid conflicts)
        for (const fileInfo of batch) {
          if (this.isCancelled) break;

          try {
            progressCallback(fileInfo.rowIndex, 'renaming');
            
            statusCallback?.({
              current: processedFiles + 1,
              total: filesToRename.length,
              fileName: fileInfo.originalFileName,
              success: successCount,
              errors: errorCount
            });

            // ИСПРАВЛЕНИЕ 5: Добавляем дополнительные проверки перед переименованием
            console.log(`[FileSearchService] 🔄 Processing file ${processedFiles + 1}/${filesToRename.length}:`);
            console.log(`  Original: "${fileInfo.originalFileName}"`);
            console.log(`  New: "${fileInfo.newFileName}"`);
            console.log(`  StaffID: "${fileInfo.staffID}"`);

            await this.renameSingleFile(fileInfo.fullOriginalPath, fileInfo.fullNewPath, requestDigest);
            
            successCount++;
            progressCallback(fileInfo.rowIndex, 'renamed');
            console.log(`[FileSearchService] ✅ SUCCESS: "${fileInfo.originalFileName}" -> "${fileInfo.newFileName}"`);
            
          } catch (error) {
            errorCount++;
            const errorMessage = error instanceof Error ? error.message : String(error);
            const detailedError = `Row ${fileInfo.rowIndex + 1} - ${fileInfo.originalFileName}: ${errorMessage}`;
            errorDetails.push(detailedError);
            progressCallback(fileInfo.rowIndex, 'error');
            
            console.error(`[FileSearchService] ❌ ERROR: "${fileInfo.originalFileName}": ${errorMessage}`);
            
            // ИСПРАВЛЕНИЕ 6: Логируем дополнительную диагностическую информацию
            console.error(`[FileSearchService] 🔍 Error details:`);
            console.error(`  Full original path: "${fileInfo.fullOriginalPath}"`);
            console.error(`  Full new path: "${fileInfo.fullNewPath}"`);
            console.error(`  Directory: "${fileInfo.directoryPath}"`);
          }
          
          processedFiles++;
          
          // ИСПРАВЛЕНИЕ 7: Увеличенная задержка между файлами
          await this.delay(1000); // 1 секунда между файлами
        }

        // ИСПРАВЛЕНИЕ 8: Увеличенная задержка между batch'ами
        if (i + BATCH_SIZE < filesToRename.length) {
          console.log(`[FileSearchService] ⏳ Waiting 2 seconds before next batch...`);
          await this.delay(2000); // 2 секунды между batch'ами
        }
      }

      console.log(`[FileSearchService] 🎯 Rename completed:`);
      console.log(`  📊 Total files: ${filesToRename.length}`);
      console.log(`  ✅ Successful: ${successCount}`);
      console.log(`  ❌ Failed: ${errorCount}`);
      console.log(`  📈 Success rate: ${filesToRename.length > 0 ? (successCount / filesToRename.length * 100).toFixed(1) + '%' : '0%'}`);

      // ИСПРАВЛЕНИЕ 9: Показываем первые несколько ошибок в консоли для диагностики
      if (errorDetails.length > 0) {
        console.error(`[FileSearchService] 📋 First few errors:`);
        errorDetails.slice(0, 3).forEach((error, index) => {
          console.error(`  ${index + 1}. ${error}`);
        });
      }

      return { success: successCount, errors: errorCount, errorDetails };

    } catch (error) {
      console.error('[FileSearchService] ❌ Critical error in rename operation:', error);
      
      const errorMessage = error instanceof Error ? error.message : String(error);
      errorDetails.push(`Critical error: ${errorMessage}`);
      
      return { 
        success: successCount, 
        errors: filesToRename.length - successCount, 
        errorDetails 
      };
    }
  }

  /**
   * ИСПРАВЛЕНИЕ 4: Улучшенная генерация нового имени файла
   */
  private generateSafeFileName(originalFileName: string, staffID: string, directoryPath: string): string {
    // Очищаем staffID от недопустимых символов
    const cleanStaffID = staffID.replace(/[<>:"/\\|?*]/g, '').trim();
    
    // Проверяем, не начинается ли уже файл с этого staffID
    if (originalFileName.toLowerCase().startsWith(cleanStaffID.toLowerCase())) {
      console.log(`[FileSearchService] ⚠️ File already starts with staffID: "${originalFileName}"`);
      return originalFileName; // Не добавляем префикс повторно
    }
    
    // Добавляем префикс с разделителем
    const newFileName = `${cleanStaffID} ${originalFileName}`;
    
    // Проверяем длину пути (SharePoint ограничение ~400 символов)
    const fullPath = `${directoryPath}/${newFileName}`;
    if (fullPath.length > 380) {
      console.warn(`[FileSearchService] ⚠️ Path too long, truncating filename`);
      
      // Сокращаем имя файла
      const extension = originalFileName.split('.').pop();
      const baseName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
      const maxBaseLength = 200 - cleanStaffID.length - extension!.length - 3; // 3 for " " and "."
      const truncatedBase = baseName.substring(0, maxBaseLength);
      
      return `${cleanStaffID} ${truncatedBase}.${extension}`;
    }
    
    return newFileName;
  }

  /**
   * Rename a single file using SharePoint REST API with better error handling
   */
  private async renameSingleFile(originalPath: string, newPath: string, requestDigest: string): Promise<void> {
    console.log(`[FileSearchService] 🔄 Renaming file:`);
    console.log(`  From: "${originalPath}"`);
    console.log(`  To: "${newPath}"`);
    
    const webUrl = this.context.pageContext.web.absoluteUrl;
    
    try {
      // ИСПРАВЛЕНИЕ 1: Проверим, существует ли файл с новым именем
      const checkNewFileUrl = `${webUrl}/_api/web/getFileByServerRelativeUrl('${encodeURIComponent(newPath)}')`;
      const checkResponse = await fetch(checkNewFileUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });
      
      if (checkResponse.ok) {
        // Файл с таким именем уже существует - добавим суффикс
        const pathParts = newPath.split('/');
        const fileName = pathParts[pathParts.length - 1];
        const directory = pathParts.slice(0, -1).join('/');
        const fileNameParts = fileName.split('.');
        const extension = fileNameParts.pop();
        const baseName = fileNameParts.join('.');
        
        // Создаем уникальное имя с timestamp
        const timestamp = new Date().getTime();
        const uniqueFileName = `${baseName}_${timestamp}.${extension}`;
        newPath = `${directory}/${uniqueFileName}`;
        
        console.log(`[FileSearchService] ⚠️ File exists, using unique name: "${newPath}"`);
      }
      
      // ИСПРАВЛЕНИЕ 2: Используем более современный SP.MoveCopyUtil API
      const moveApiUrl = `${webUrl}/_api/SP.MoveCopyUtil.MoveFileByPath()`;
      
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
          KeepBoth: false, // Не сохранять оба файла
          ResetAuthorAndCreatedOnCopy: false,
          ShouldBypassSharedLocks: true
        }
      };
      
      console.log(`[FileSearchService] 📞 Using SP.MoveCopyUtil.MoveFileByPath API`);
      
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
        console.error(`[FileSearchService] Modern API failed, trying fallback. Error:`, errorText);
        
        // FALLBACK: Используем старый MoveTo API с флагом overwrite
        await this.renameSingleFileWithMoveTo(originalPath, newPath, requestDigest);
      } else {
        console.log(`[FileSearchService] ✅ File renamed successfully using modern API`);
      }
      
    } catch (error) {
      console.error(`[FileSearchService] Error in modern API, trying fallback:`, error);
      
      // FALLBACK: Используем старый MoveTo API
      await this.renameSingleFileWithMoveTo(originalPath, newPath, requestDigest);
    }
  }

  /**
   * Fallback method using old MoveTo API
   */
  private async renameSingleFileWithMoveTo(originalPath: string, newPath: string, requestDigest: string): Promise<void> {
    console.log(`[FileSearchService] 🔄 Using fallback MoveTo API`);
    
    const webUrl = this.context.pageContext.web.absoluteUrl;
    
    // ИСПРАВЛЕНИЕ 3: Используем правильное кодирование URL и флаг overwrite
    const encodedOriginalPath = encodeURIComponent(originalPath);
    const encodedNewPath = encodeURIComponent(newPath);
    
    // flags=1 означает overwrite existing file
    const moveToUrl = `${webUrl}/_api/web/getFileByServerRelativeUrl('${encodedOriginalPath}')/MoveTo(newurl='${encodedNewPath}',flags=1)`;
    
    console.log(`[FileSearchService] 📞 MoveTo URL:`, moveToUrl);
    
    const response = await fetch(moveToUrl, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'X-RequestDigest': requestDigest
      }
    });
    
    if (!response.ok) {
      const errorText = await response.text();
      console.error(`[FileSearchService] ❌ MoveTo API also failed:`, errorText);
      
      // Попробуем распарсить ошибку
      let errorMessage = errorText;
      try {
        const errorData = JSON.parse(errorText);
        if (errorData.error && errorData.error.message) {
          errorMessage = errorData.error.message.value || errorData.error.message;
        }
      } catch (e) {
        // Игнорируем ошибки парсинга
      }
      
      throw new Error(`HTTP ${response.status}: ${errorMessage}`);
    }
    
    console.log(`[FileSearchService] ✅ File renamed successfully using MoveTo API`);
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
   * OPTIMIZED STAGE 3: Search files with CORRECTED LOGIC and MINIMAL API calls + DETAILED LOGGING
   */
  private async executeOptimizedStage3_SearchFiles(
    currentProgress: ISearchProgress,
    rows: IRenameTableRow[],
    results: { [rowIndex: number]: 'found' | 'not-found' | 'searching' },
    progressCallback: (rowIndex: number, result: 'found' | 'not-found' | 'searching') => void,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<void> {
    
    console.log('[FileSearchService] 🚀 OPTIMIZED STAGE 3 with DETAILED LOGGING...');
    
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
      console.log(`  📁 "${dir}" -> ${files.length} files to search`);
    });

    let processedFiles = 0;
    let foundFiles = 0;
    const totalFiles = rows.length;
    const directories = Object.keys(directoryToFilesMap);

    console.log(`[FileSearchService] 🎯 STARTING SEARCH: ${totalFiles} total files in ${directories.length} directories`);

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
        // ONE API CALL to get directory contents with adaptive timeout
        console.log(`[FileSearchService] 📞 API call: getFolderContents("${directoryPath}")`);
        const startTime = Date.now();
        
        const adaptiveTimeout = this.calculateTimeout(filesFromExcel.length);
        const folderContentsPromise = this.folderService.getFolderContents(directoryPath);
        const folderContents = await Promise.race([
          folderContentsPromise,
          this.createTimeoutPromise(adaptiveTimeout, { files: [], folders: [] })
        ]) as {files: any[], folders: any[]};
        
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
        const sharePointFilesMap = new Map<string, any>();
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
          const result = fileExists ? 'found' : 'not-found';
          
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
      
      console.log(`[FileSearchService] 📊 OVERALL PROGRESS: ${processedFiles}/${totalFiles} files, ${foundFiles} found`);
      console.log(`[FileSearchService] ➡️ Moving to next directory...\n`);
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
          const fileName = String(row.cells['custom_0']?.value || '').trim();
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