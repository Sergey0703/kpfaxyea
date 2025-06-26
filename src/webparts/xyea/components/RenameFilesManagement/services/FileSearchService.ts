// src/webparts/xyea/components/RenameFilesManagement/services/FileSearchService.ts

import { 
  IRenameTableRow, 
  SearchStage, 
  ISearchProgress, 
  ISearchPlan,
  SearchProgressHelper,
  FileSearchStatus,
  DirectoryStatusCallback,
  FileSearchResultCallback
} from '../types/RenameFilesTypes';
import { SharePointFolderService } from './SharePointFolderService';
import { ExcelFileProcessor } from './ExcelFileProcessor';
import { DirectoryAnalysisService } from './FileSearch/DirectoryAnalysisService';
import { FileSearchOperationsService } from './FileSearch/FileSearchOperationsService';
import { FileRenameService } from './FileSearch/FileRenameService';
import { FileSearchConfigService } from './FileSearch/FileSearchConfigService';

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
  private context: IWebPartContext;
  private folderService: SharePointFolderService;
  private excelProcessor: ExcelFileProcessor;
  private directoryAnalysisService: DirectoryAnalysisService;
  private fileSearchOperationsService: FileSearchOperationsService;
  private fileRenameService: FileRenameService;
  private configService: FileSearchConfigService;
  
  private isCancelled: boolean = false;
  private currentSearchId: string | undefined = undefined;

  constructor(context: IWebPartContext) {
    this.context = context;
    this.folderService = new SharePointFolderService(context);
    this.excelProcessor = new ExcelFileProcessor();
    
    // Initialize service modules
    this.configService = new FileSearchConfigService();
    this.directoryAnalysisService = new DirectoryAnalysisService(this.folderService, this.excelProcessor, this.configService);
    this.fileSearchOperationsService = new FileSearchOperationsService(this.folderService, this.configService);
    this.fileRenameService = new FileRenameService(context, this.configService);
  }

  /**
   * UPDATED: Analyze directories and check existence (Stages 1-2) with directory status callback
   */
  public async analyzeDirectories(
    folderPath: string,
    rows: IRenameTableRow[],
    statusCallback?: (progress: ISearchProgress) => void,
    directoryStatusCallback?: DirectoryStatusCallback
  ): Promise<ISearchProgress> {
    
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] Starting directory analysis with status callback (Search ID: ${searchId})`);
    
    let currentProgress = SearchProgressHelper.createInitialProgress();
    
    try {
      // STAGE 1: ANALYZING DIRECTORIES (0-50%)
      currentProgress = await this.directoryAnalysisService.executeStage1_AnalyzeDirectories(
        rows, 
        folderPath, 
        currentProgress, 
        statusCallback,
        () => this.isCancelled || this.currentSearchId !== searchId
      );
      
      if (this.isCancelled || this.currentSearchId !== searchId) {
        throw new Error('Analysis was cancelled');
      }

      // STAGE 2: CHECKING DIRECTORY EXISTENCE (50-100%)
      currentProgress = await this.directoryAnalysisService.executeStage2_CheckDirectoryExistence_OPTIMIZED(
        currentProgress,
        statusCallback,
        directoryStatusCallback,
        () => this.isCancelled || this.currentSearchId !== searchId
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
    progressCallback: FileSearchResultCallback,
    statusCallback?: (progress: ISearchProgress) => void
  ): Promise<{ [rowIndex: number]: FileSearchStatus }> {
    
    this.currentSearchId = Date.now().toString();
    const searchId = this.currentSearchId;
    this.isCancelled = false;
    
    console.log(`[FileSearchService] 🚀 STARTING OPTIMIZED FILE SEARCH (Search ID: ${searchId})`);
    
    try {
      const results = await this.fileSearchOperationsService.executeOptimizedStage3_SearchFiles(
        searchProgress,
        rows,
        progressCallback,
        statusCallback,
        () => this.isCancelled || this.currentSearchId !== searchId
      );

      console.log('[FileSearchService] File search completed:', results);
      return results;

    } catch (error) {
      console.error('[FileSearchService] Error during file search:', error);
      
      const errorProgress = SearchProgressHelper.transitionToStage(
        searchProgress,
        SearchStage.ERROR,
        {
          currentFileName: 'File search failed',
          errors: [error instanceof Error ? error.message : 'Unknown error']
        }
      );
      statusCallback?.(errorProgress);
      
      // Mark unprocessed rows with appropriate status
      const results: { [rowIndex: number]: FileSearchStatus } = {};
      rows.forEach(row => {
        const directoryPath = this.getDirectoryFromRow(row);
        const directoryExists = this.checkDirectoryExistsInPlan(directoryPath, searchProgress.searchPlan);
        
        if (!directoryExists) {
          results[row.rowIndex] = 'directory-missing';
          progressCallback(row.rowIndex, 'directory-missing');
        } else {
          results[row.rowIndex] = 'not-found';
          progressCallback(row.rowIndex, 'not-found');
        }
      });
      
      return results;
    }
  }

  /**
   * Rename found files with staffID prefix - UPDATED: Skip if target exists
   */
  public async renameFoundFiles(
    rows: IRenameTableRow[],
    fileSearchResults: { [rowIndex: number]: FileSearchStatus },
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
    
    return this.fileRenameService.renameFoundFiles(
      rows,
      fileSearchResults,
      baseFolderPath,
      progressCallback,
      statusCallback,
      () => this.isCancelled || this.currentSearchId !== searchId
    );
  }

  /**
   * Helper method to get directory path from row
   */
  private getDirectoryFromRow(row: IRenameTableRow): string {
    const directoryCell = row.cells.custom_1;
    if (directoryCell && directoryCell.value) {
      return String(directoryCell.value).trim();
    }
    return this.excelProcessor.extractDirectoryPathFromRow(row);
  }

  /**
   * Check if directory exists in search plan
   */
  private checkDirectoryExistsInPlan(directoryPath: string, searchPlan?: ISearchPlan): boolean {
    if (!searchPlan) return false;
    
    const directoryGroup = searchPlan.directoryGroups.find(group => 
      group.directoryPath === directoryPath
    );
    
    return directoryGroup?.exists || false;
  }

  public cancelSearch(): void {
    console.log('[FileSearchService] Cancelling file search...');
    this.isCancelled = true;
    this.currentSearchId = undefined;
  }

  public isSearchActive(): boolean {
    return this.currentSearchId !== undefined && !this.isCancelled;
  }

  /**
   * Keep existing methods for compatibility
   */
  public async searchSingleFile(folderPath: string, fileName: string): Promise<{ found: boolean; path?: string }> {
    try {
      const folderContents = await this.folderService.getFolderContents(folderPath);
      const files = folderContents.files;
      
      const fileFound = files.some((file: ISharePointFolder) =>
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

  public async getFileDetails(filePath: string): Promise<ISharePointFolder | undefined> {
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
    const fileName = String(row.cells.custom_0?.value || '');
    console.log(`[FileSearchService] getFileNameFromRow for row ${row.rowIndex}: "${fileName}"`);
    return fileName;
  }
}