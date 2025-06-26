// src/webparts/xyea/components/RenameFilesManagement/services/FileSearch/FileSearchConfigService.ts

export class FileSearchConfigService {
  // AGGRESSIVE: Much shorter timeouts to prevent hanging
  public readonly DIRECTORY_CHECK_TIMEOUT = 3000; // 3 seconds per directory
  public readonly FOLDER_LOAD_TIMEOUT = 8000; // 8 seconds for folder loading
  public readonly FILE_SEARCH_TIMEOUT = 5000; // 5 seconds per file search
  public readonly FILE_RENAME_TIMEOUT = 10000; // 10 seconds per file rename
  
  // Batch processing settings
  public readonly BATCH_SIZE = 5; // Process directories in batches of 5
  public readonly DELAY_BETWEEN_OPERATIONS = 50; // 50ms delay between operations
  public readonly DELAY_BETWEEN_BATCHES = 200; // 200ms delay between batches
  public readonly DELAY_BETWEEN_RENAMES = 2000; // 2 seconds delay between renames

  /**
   * Calculate adaptive timeout based on file count
   */
  public calculateTimeout(fileCount: number): number {
    const baseTimeout = 2000;
    const additionalTime = Math.min(fileCount * 50, 15000); // max 15 seconds
    const adaptiveTimeout = baseTimeout + additionalTime;
    
    console.log(`[FileSearchConfigService] 📊 Adaptive timeout for ${fileCount} files: ${adaptiveTimeout}ms`);
    return adaptiveTimeout;
  }

  /**
   * Helper method to create timeout promise
   */
  public createTimeoutPromise<T>(timeoutMs: number, errorMessage: string | T): Promise<T> {
    return new Promise((resolve, reject) => {
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

  /**
   * Create a delay promise
   */
  public delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Normalize path for comparison
   */
  public normalizePath(path: string): string {
    return path
      .replace(/\\/g, '/')
      .replace(/\/+/g, '/')
      .toLowerCase()
      .replace(/\/$/, '');
  }

  /**
   * Check if file looks like a valid file path
   */
  public looksLikeValidFilePath(value: string): boolean {
    // Must contain at least one directory separator
    if (!value.includes('\\') && !value.includes('/')) {
      return false;
    }
    
    // Must have a file extension at the end
    const parts = value.split(/[\\//]/);
    const lastPart = parts[parts.length - 1];
    if (!lastPart || !lastPart.includes('.')) {
      return false;
    }
    
    // The file extension should be reasonable (2-5 characters after the last dot)
    const extensionMatch = lastPart.match(/\.([a-zA-Z0-9]{2,5})$/);
    if (!extensionMatch) {
      return false;
    }
    
    // Should have multiple path components (not just a filename)
    if (parts.length < 2) {
      return false;
    }
    
    // Must be reasonably long for a file path
    if (value.length < 10) {
      return false;
    }
    
    // Reject if it looks like a person's name or other non-path content
    if (value.includes('(') && value.includes(')') && 
        (value.includes('INST') || value.includes('DRV') || value.includes('MGR'))) {
      return false;
    }
    
    return true;
  }

  /**
   * Validate filename for SharePoint
   */
  public isValidFileName(fileName: string): boolean {
    const invalidChars = /[<>:"/\\|?*]/;
    return !invalidChars.test(fileName) &&
           fileName.trim().length > 0 &&
           fileName.length <= 255;
  }

  /**
   * Clean SharePoint path
   */
  public cleanSharePointPath(path: string): string {
    let cleanPath = path.trim().replace(/\\/g, '/');
    cleanPath = cleanPath.replace(/\/+/g, '/');
    cleanPath = cleanPath.replace(/\/$/, '');
    
    if (!cleanPath.startsWith('/')) {
      cleanPath = '/' + cleanPath;
    }
    
    console.log(`[FileSearchConfigService] Path cleaning: "${path}" -> "${cleanPath}"`);
    return cleanPath;
  }

  /**
   * Generate safe filename with staffID prefix
   */
  public generateSafeFileName(originalFileName: string, staffID: string, directoryPath: string): string {
    const cleanStaffID = staffID.replace(/[<>:"/\\|?*]/g, '').trim();
    
    if (originalFileName.toLowerCase().startsWith(cleanStaffID.toLowerCase())) {
      console.log(`[FileSearchConfigService] ⚠️ File already starts with staffID: "${originalFileName}"`);
      return originalFileName;
    }
    
    const newFileName = `${cleanStaffID} ${originalFileName}`;
    
    const fullPath = `${directoryPath}/${newFileName}`;
    if (fullPath.length > 380) {
      console.warn(`[FileSearchConfigService] ⚠️ Path too long, truncating filename`);
      
      const extension = originalFileName.split('.').pop();
      const baseName = originalFileName.substring(0, originalFileName.lastIndexOf('.'));
      const maxBaseLength = 200 - cleanStaffID.length - (extension?.length || 0) - 3;
      const truncatedBase = baseName.substring(0, maxBaseLength);
      
      return `${cleanStaffID} ${truncatedBase}.${extension}`;
    }
    
    return newFileName;
  }

  /**
   * Build directory SharePoint path
   */
  public buildDirectoryPath(relativePath: string, basePath: string): string {
    const normalizedRelative = relativePath.replace(/\\/g, '/');
    const fullPath = `${basePath}/${normalizedRelative}`;
    return fullPath.replace(/\/+/g, '/').replace(/\/$/, '');
  }

  /**
   * SharePoint API endpoints
   */
  public getSharePointEndpoints(): {
    CONTEXT_INFO: string;
    GET_FILE: string;
    GET_FOLDER: string;
    MOVE_FILE_SIMPLE: string;
    MOVE_FILE_MODERN: string;
  } {
    return {
      CONTEXT_INFO: '/_api/contextinfo',
      GET_FILE: '/_api/web/getFileByServerRelativeUrl',
      GET_FOLDER: '/_api/web/getFolderByServerRelativeUrl',
      MOVE_FILE_SIMPLE: '/MoveTo',
      MOVE_FILE_MODERN: '/_api/SP.MoveCopyUtil.MoveFileByPath'
    };
  }

  /**
   * Get performance logging configuration
   */
  public getLoggingConfig(): {
    enableDetailedLogging: boolean;
    logBatchOperations: boolean;
    logApiCalls: boolean;
    logTimingInfo: boolean;
  } {
    return {
      enableDetailedLogging: process.env.NODE_ENV === 'development',
      logBatchOperations: true,
      logApiCalls: true,
      logTimingInfo: true
    };
  }
}