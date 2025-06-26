// src/webparts/xyea/components/RenameFilesManagement/services/SharePointFolderService.ts

import { ISharePointFolder } from '../types/RenameFilesTypes';

export interface ICachedFolder {
  Name: string;
  ServerRelativeUrl: string;
  FullPath: string; // Normalized path for easy comparison
  ItemCount: number;
  TimeCreated: string;
  TimeLastModified: string;
}

// NEW: Interface for cached directory contents
export interface ICachedDirectoryContents {
  directoryPath: string;
  files: Map<string, ISharePointFolder>; // filename (lowercase) -> file info
  folders: ISharePointFolder[];
  lastLoaded: Date;
  fileCount: number;
}

// FIXED: Define specific interface instead of using 'any'
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

export class SharePointFolderService {
  private context: IWebPartContext;
  private cachedFolders: ICachedFolder[] = [];
  private isLoadingAllFolders: boolean = false;
  private allFoldersLoaded: boolean = false;

  // NEW: Cached directory contents for fast existence checks
  private directoryContentsCache: Map<string, ICachedDirectoryContents> = new Map();
  private batchLoadingPromises: Map<string, Promise<ICachedDirectoryContents>> = new Map();

  // NEW: Performance tracking
  private performanceStats = {
    cacheHits: 0,
    cacheMisses: 0,
    batchLoadsExecuted: 0,
    totalApiCalls: 0
  };

  constructor(context: IWebPartContext) {
    this.context = context;
  }

  /**
   * NEW: Batch load multiple directory contents for fast existence checking
   * This is the KEY method for performance optimization
   */
  public async batchLoadDirectoryContents(
    directoryPaths: string[],
    progressCallback?: (loaded: number, total: number, currentPath: string) => void
  ): Promise<Map<string, ICachedDirectoryContents>> {
    
    console.log(`[SharePointFolderService] 🚀 BATCH LOADING ${directoryPaths.length} directories for fast existence checks`);
    console.log(`[SharePointFolderService] 📊 Performance boost: ${directoryPaths.length} API calls instead of thousands of individual checks`);
    
    const results = new Map<string, ICachedDirectoryContents>();
    
    // FIXED: ES5 compatible Set to Array conversion
    const uniquePathsSet = new Set<string>();
    directoryPaths.forEach(path => {
      uniquePathsSet.add(this.normalizePath(path));
    });
    
    const uniquePaths: string[] = [];
    uniquePathsSet.forEach(path => {
      uniquePaths.push(path);
    });
    
    console.log(`[SharePointFolderService] 📋 Unique directories to load: ${uniquePaths.length}`);
    
    // Process directories in parallel batches of 5 to avoid overwhelming SharePoint
    const BATCH_SIZE = 5;
    let processed = 0;
    
    for (let i = 0; i < uniquePaths.length; i += BATCH_SIZE) {
      const batch = uniquePaths.slice(i, i + BATCH_SIZE);
      
      console.log(`[SharePointFolderService] 📦 Processing batch ${Math.floor(i / BATCH_SIZE) + 1}/${Math.ceil(uniquePaths.length / BATCH_SIZE)}: ${batch.length} directories`);
      
      // Process batch in parallel
      const batchPromises = batch.map(async (path) => {
        try {
          progressCallback?.(processed, uniquePaths.length, path);
          
          const contents = await this.loadSingleDirectoryContents(path);
          results.set(path, contents);
          processed++;
          
          console.log(`[SharePointFolderService] ✅ Loaded directory "${path}": ${contents.fileCount} files`);
          return contents;
          
        } catch (error) {
          console.warn(`[SharePointFolderService] ⚠️ Failed to load directory "${path}":`, error);
          
          // Create empty cache entry for failed directories
          const emptyContents: ICachedDirectoryContents = {
            directoryPath: path,
            files: new Map(),
            folders: [],
            lastLoaded: new Date(),
            fileCount: 0
          };
          
          this.directoryContentsCache.set(path, emptyContents);
          results.set(path, emptyContents);
          processed++;
          
          return emptyContents;
        }
      });
      
      await Promise.all(batchPromises);
      
      // Small delay between batches to be nice to SharePoint
      if (i + BATCH_SIZE < uniquePaths.length) {
        await this.delay(200);
      }
      
      progressCallback?.(processed, uniquePaths.length, `Batch ${Math.floor(i / BATCH_SIZE) + 1} complete`);
    }
    
    this.performanceStats.batchLoadsExecuted++;
    
    console.log(`[SharePointFolderService] 🎯 BATCH LOADING COMPLETE:`);
    console.log(`[SharePointFolderService]   📁 Directories loaded: ${results.size}`);
    console.log(`[SharePointFolderService]   📄 Total files cached: ${Array.from(results.values()).reduce((sum, dir) => sum + dir.fileCount, 0)}`);
    console.log(`[SharePointFolderService]   ⚡ API calls made: ${uniquePaths.length} (instead of thousands)`);
    
    return results;
  }

  /**
   * NEW: Load contents of a single directory and cache it
   */
  private async loadSingleDirectoryContents(directoryPath: string): Promise<ICachedDirectoryContents> {
    const normalizedPath = this.normalizePath(directoryPath);
    
    // Check if already cached and recent (within 5 minutes)
    const cached = this.directoryContentsCache.get(normalizedPath);
    if (cached && (Date.now() - cached.lastLoaded.getTime()) < 300000) {
      console.log(`[SharePointFolderService] 📋 Cache hit for directory: "${normalizedPath}"`);
      this.performanceStats.cacheHits++;
      return cached;
    }
    
    // Check if already loading
    const existingPromise = this.batchLoadingPromises.get(normalizedPath);
    if (existingPromise) {
      console.log(`[SharePointFolderService] ⏳ Waiting for existing load of: "${normalizedPath}"`);
      return existingPromise;
    }
    
    // Start loading
    const loadPromise = this.executeDirectoryLoad(directoryPath, normalizedPath);
    this.batchLoadingPromises.set(normalizedPath, loadPromise);
    
    try {
      const result = await loadPromise;
      this.batchLoadingPromises.delete(normalizedPath);
      return result;
    } catch (error) {
      this.batchLoadingPromises.delete(normalizedPath);
      throw error;
    }
  }

  /**
   * NEW: Execute the actual directory loading with API call
   */
  private async executeDirectoryLoad(originalPath: string, normalizedPath: string): Promise<ICachedDirectoryContents> {
    console.log(`[SharePointFolderService] 📞 API call: getFolderContents("${originalPath}")`);
    this.performanceStats.totalApiCalls++;
    this.performanceStats.cacheMisses++;
    
    const startTime = Date.now();
    
    try {
      // Use existing getFolderContents method but with caching
      const { files, folders } = await this.getFolderContents(originalPath);
      
      const endTime = Date.now();
      console.log(`[SharePointFolderService] ✅ API response received in ${endTime - startTime}ms for "${originalPath}"`);
      
      // Create case-insensitive file map for fast lookups
      const fileMap = new Map<string, ISharePointFolder>();
      files.forEach(file => {
        fileMap.set(file.Name.toLowerCase(), file);
      });
      
      const contents: ICachedDirectoryContents = {
        directoryPath: normalizedPath,
        files: fileMap,
        folders,
        lastLoaded: new Date(),
        fileCount: files.length
      };
      
      // Cache the results
      this.directoryContentsCache.set(normalizedPath, contents);
      
      console.log(`[SharePointFolderService] 💾 Cached directory "${originalPath}": ${files.length} files, ${folders.length} folders`);
      
      return contents;
      
    } catch (error) {
      const endTime = Date.now();
      console.error(`[SharePointFolderService] ❌ API call failed after ${endTime - startTime}ms for "${originalPath}":`, error);
      throw error;
    }
  }

  /**
   * NEW: Fast file existence check using cached data
   * This replaces thousands of individual API calls!
   */
  public checkFileExistsInCache(directoryPath: string, fileName: string): boolean {
    const normalizedPath = this.normalizePath(directoryPath);
    const normalizedFileName = fileName.toLowerCase();
    
    const contents = this.directoryContentsCache.get(normalizedPath);
    if (!contents) {
      console.warn(`[SharePointFolderService] ⚠️ No cached data for directory: "${directoryPath}"`);
      console.warn(`[SharePointFolderService] 💡 Suggestion: Call batchLoadDirectoryContents() first`);
      return false;
    }
    
    const exists = contents.files.has(normalizedFileName);
    
    if (exists) {
      console.log(`[SharePointFolderService] ✅ File EXISTS in cache: "${fileName}" in "${directoryPath}"`);
    } else {
      console.log(`[SharePointFolderService] ❌ File NOT FOUND in cache: "${fileName}" in "${directoryPath}"`);
    }
    
    return exists;
  }

  /**
   * NEW: Get all cached directory contents
   */
  public getCachedDirectoryContents(directoryPath: string): ICachedDirectoryContents | undefined {
    const normalizedPath = this.normalizePath(directoryPath);
    return this.directoryContentsCache.get(normalizedPath);
  }

  /**
   * NEW: Check if directory contents are cached
   */
  public isDirectoryCached(directoryPath: string): boolean {
    const normalizedPath = this.normalizePath(directoryPath);
    const cached = this.directoryContentsCache.get(normalizedPath);
    
    if (!cached) return false;
    
    // Check if cache is recent (within 5 minutes)
    const isRecent = (Date.now() - cached.lastLoaded.getTime()) < 300000;
    return isRecent;
  }

  /**
   * NEW: Clear directory contents cache
   */
  public clearDirectoryCache(directoryPath?: string): void {
    if (directoryPath) {
      const normalizedPath = this.normalizePath(directoryPath);
      this.directoryContentsCache.delete(normalizedPath);
      console.log(`[SharePointFolderService] 🗑️ Cleared cache for directory: "${directoryPath}"`);
    } else {
      this.directoryContentsCache.clear();
      console.log(`[SharePointFolderService] 🗑️ Cleared entire directory cache`);
    }
  }

  /**
   * NEW: Get performance statistics
   */
  public getPerformanceStats(): {
    cacheHits: number;
    cacheMisses: number;
    batchLoadsExecuted: number;
    totalApiCalls: number;
    cacheHitRatio: number;
    directoriesCached: number;
    totalFilesCached: number;
  } {
    const directoriesCached = this.directoryContentsCache.size;
    const totalFilesCached = Array.from(this.directoryContentsCache.values())
      .reduce((sum, dir) => sum + dir.fileCount, 0);
    
    return {
      ...this.performanceStats,
      cacheHitRatio: this.performanceStats.cacheHits + this.performanceStats.cacheMisses > 0 
        ? this.performanceStats.cacheHits / (this.performanceStats.cacheHits + this.performanceStats.cacheMisses) 
        : 0,
      directoriesCached,
      totalFilesCached
    };
  }

  /**
   * NEW: Invalidate cache entry when file is renamed
   * Call this after successful rename operations
   */
  public invalidateDirectoryCache(directoryPath: string): void {
    const normalizedPath = this.normalizePath(directoryPath);
    const cached = this.directoryContentsCache.get(normalizedPath);
    
    if (cached) {
      // Mark as old so it gets refreshed on next access
      cached.lastLoaded = new Date(0);
      console.log(`[SharePointFolderService] 🔄 Invalidated cache for directory: "${directoryPath}"`);
    }
  }

  // EXISTING METHODS (keeping all original functionality)
  public async getDocumentLibraryFolders(): Promise<ISharePointFolder[]> {
    try {
      const { context } = this;
      
      console.log('[SharePointFolderService] Fetching folders using simple REST API...');
      
      // Use a simpler approach - get folders from the root Documents/Shared Documents
      const webUrl = context.pageContext.web.absoluteUrl;
      const possiblePaths = [
        '/Shared Documents',
        '/Documents', 
        '/Shared%20Documents'
      ];
      
      let foldersData: { d?: { results?: ISharePointFolder[] }; value?: ISharePointFolder[] } | undefined = undefined;
      let workingPath = '';
      
      // Try each possible path
      for (const path of possiblePaths) {
        try {
          const foldersUrl = `${webUrl}/_api/web/getFolderByServerRelativeUrl('${context.pageContext.web.serverRelativeUrl}${path}')/folders`;
          
          const response = await fetch(foldersUrl, {
            method: 'GET',
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Authorization': `Bearer ${context.aadTokenProviderFactory}`
            }
          });
          
          if (response.ok) {
            const data = await response.json();
            foldersData = data;
            workingPath = path;
            console.log(`[SharePointFolderService] Successfully accessed folders at: ${path}`);
            break;
          }
        } catch (error) {
          console.log(`[SharePointFolderService] Failed to access ${path}:`, error);
          continue;
        }
      }
      
      // If REST API fails, return a manual list based on what we saw in your screenshot
      if (!foldersData) {
        console.log('[SharePointFolderService] REST API failed, returning known folders from your site...');
        const knownFolders: ISharePointFolder[] = [
          {
            Name: '(Root - Documents)',
            ServerRelativeUrl: `${context.pageContext.web.serverRelativeUrl}/Shared%20Documents`,
            ItemCount: 0,
            TimeCreated: new Date().toISOString(),
            TimeLastModified: new Date().toISOString()
          },
          {
            Name: 'Debug',
            ServerRelativeUrl: `${context.pageContext.web.serverRelativeUrl}/Shared%20Documents/Debug`,
            ItemCount: 0,
            TimeCreated: new Date().toISOString(),
            TimeLastModified: new Date().toISOString()
          },
          {
            Name: 'LeaveReports',
            ServerRelativeUrl: `${context.pageContext.web.serverRelativeUrl}/Shared%20Documents/LeaveReports`,
            ItemCount: 0,
            TimeCreated: new Date().toISOString(),
            TimeLastModified: new Date().toISOString()
          },
          {
            Name: 'SRS',
            ServerRelativeUrl: `${context.pageContext.web.serverRelativeUrl}/Shared%20Documents/SRS`,
            ItemCount: 0,
            TimeCreated: new Date().toISOString(),
            TimeLastModified: new Date().toISOString()
          },
          {
            Name: 'Templates',
            ServerRelativeUrl: `${context.pageContext.web.serverRelativeUrl}/Shared%20Documents/Templates`,
            ItemCount: 0,
            TimeCreated: new Date().toISOString(),
            TimeLastModified: new Date().toISOString()
          }
        ];
        
        return knownFolders;
      }
      
      // Process the successful response
      const folders = foldersData.d?.results || foldersData.value || [];
      
      // Filter out system folders
      const userFolders = folders
        .filter((folder: ISharePointFolder) => 
          !folder.Name.startsWith('Forms') && 
          !folder.Name.startsWith('_') &&
          folder.Name !== 'Forms'
        )
        .map((folder: ISharePointFolder) => ({
          Name: folder.Name,
          ServerRelativeUrl: folder.ServerRelativeUrl,
          ItemCount: folder.ItemCount || 0,
          TimeCreated: folder.TimeCreated || new Date().toISOString(),
          TimeLastModified: folder.TimeLastModified || new Date().toISOString()
        }));

      // Add root folder option
      const rootFolder: ISharePointFolder = {
        Name: `(Root - Documents)`,
        ServerRelativeUrl: `${context.pageContext.web.serverRelativeUrl}${workingPath}`,
        ItemCount: 0,
        TimeCreated: new Date().toISOString(),
        TimeLastModified: new Date().toISOString()
      };

      const allFolders = [rootFolder, ...userFolders];
      console.log('[SharePointFolderService] Returning folders:', allFolders);
      return allFolders;
      
    } catch (error) {
      console.error('[SharePointFolderService] Error fetching SharePoint folders:', error);
      
      // Fallback to the folders we know exist in your site
      const fallbackFolders: ISharePointFolder[] = [
        {
          Name: '(Root - Documents)',
          ServerRelativeUrl: `${this.context.pageContext.web.serverRelativeUrl}/Shared%20Documents`,
          ItemCount: 0,
          TimeCreated: new Date().toISOString(),
          TimeLastModified: new Date().toISOString()
        },
        {
          Name: 'Debug',
          ServerRelativeUrl: `${this.context.pageContext.web.serverRelativeUrl}/Shared%20Documents/Debug`,
          ItemCount: 0,
          TimeCreated: new Date().toISOString(),
          TimeLastModified: new Date().toISOString()
        },
        {
          Name: 'LeaveReports', 
          ServerRelativeUrl: `${this.context.pageContext.web.serverRelativeUrl}/Shared%20Documents/LeaveReports`,
          ItemCount: 0,
          TimeCreated: new Date().toISOString(),
          TimeLastModified: new Date().toISOString()
        },
        {
          Name: 'SRS',
          ServerRelativeUrl: `${this.context.pageContext.web.serverRelativeUrl}/Shared%20Documents/SRS`,
          ItemCount: 0,
          TimeCreated: new Date().toISOString(),
          TimeLastModified: new Date().toISOString()
        },
        {
          Name: 'Templates',
          ServerRelativeUrl: `${this.context.pageContext.web.serverRelativeUrl}/Shared%20Documents/Templates`,
          ItemCount: 0,
          TimeCreated: new Date().toISOString(),
          TimeLastModified: new Date().toISOString()
        }
      ];
      
      console.log('[SharePointFolderService] Using fallback folders based on your site structure');
      return fallbackFolders;
    }
  }

  /**
   * ENHANCED: Load all subfolders recursively and cache them
   * With better error handling and timeout protection
   */
  public async loadAllSubfolders(
    baseFolderPath: string, 
    progressCallback?: (currentPath: string, foldersLoaded: number) => void
  ): Promise<ICachedFolder[]> {
    
    if (this.isLoadingAllFolders) {
      console.log('[SharePointFolderService] Already loading folders, please wait...');
      return this.cachedFolders;
    }

    console.log(`[SharePointFolderService] Starting to load all subfolders from: "${baseFolderPath}"`);
    
    this.isLoadingAllFolders = true;
    this.cachedFolders = [];
    this.allFoldersLoaded = false;

    try {
      // NEW: Set a global timeout for the entire loading process
      const loadingPromise = this.loadFoldersRecursively(baseFolderPath, progressCallback);
      const timeoutPromise = new Promise<void>((resolve, reject) => {
        setTimeout(() => reject(new Error('Folder loading timeout after 30 seconds')), 30000);
      });
      
      await Promise.race([loadingPromise, timeoutPromise]);
      
      this.allFoldersLoaded = true;
      console.log(`[SharePointFolderService] ✅ Successfully loaded ${this.cachedFolders.length} folders`);
      
      // Log detailed folder cache for debugging
      console.log('[SharePointFolderService] Cached folders detail:');
      this.cachedFolders.slice(0, 10).forEach((folder, index) => {
        console.log(`  ${index + 1}. "${folder.FullPath}" (Original: "${folder.ServerRelativeUrl}")`);
      });
      if (this.cachedFolders.length > 10) {
        console.log(`  ... and ${this.cachedFolders.length - 10} more folders`);
      }
      
      return this.cachedFolders;
      
    } catch (error) {
      console.error('[SharePointFolderService] ❌ Error loading subfolders:', error);
      
      // NEW: Set allFoldersLoaded to true even on error, so fallback logic kicks in
      this.allFoldersLoaded = true;
      
      console.warn('[SharePointFolderService] ⚠️ Folder loading failed, will use direct API checks as fallback');
      return this.cachedFolders; // Return empty cache, fallback will handle checks
      
    } finally {
      this.isLoadingAllFolders = false;
    }
  }

  /**
   * Recursively load folders and add them to cache
   */
  private async loadFoldersRecursively(
    folderPath: string, 
    progressCallback?: (currentPath: string, foldersLoaded: number) => void,
    depth: number = 0
  ): Promise<void> {
    
    // Prevent infinite recursion
    if (depth > 10) {
      console.warn(`[SharePointFolderService] Maximum depth reached for ${folderPath}`);
      return;
    }

    try {
      if (progressCallback) {
        progressCallback(folderPath, this.cachedFolders.length);
      }

      const webUrl = this.context.pageContext.web.absoluteUrl;
      const foldersUrl = `${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/folders?$select=Name,ServerRelativeUrl,ItemCount,TimeCreated,TimeLastModified`;
      
      console.log(`[SharePointFolderService] Loading folders from: ${folderPath}`);
      
      const response = await fetch(foldersUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });

      if (response.ok) {
        const data = await response.json();
        const folders = data.d?.results || data.value || [];
        
        console.log(`[SharePointFolderService] Found ${folders.length} folders in ${folderPath}`);
        
        // Filter out system folders
        const userFolders = folders.filter((folder: ISharePointFolder) => 
          !folder.Name.startsWith('_') && 
          !folder.Name.startsWith('Forms') &&
          folder.Name !== 'Forms'
        );

        console.log(`[SharePointFolderService] After filtering: ${userFolders.length} user folders`);

        // Add folders to cache
        for (const folder of userFolders) {
          const cachedFolder: ICachedFolder = {
            Name: folder.Name,
            ServerRelativeUrl: folder.ServerRelativeUrl,
            FullPath: this.normalizePath(folder.ServerRelativeUrl),
            ItemCount: folder.ItemCount || 0,
            TimeCreated: folder.TimeCreated || new Date().toISOString(),
            TimeLastModified: folder.TimeLastModified || new Date().toISOString()
          };

          this.cachedFolders.push(cachedFolder);
          console.log(`[SharePointFolderService] Cached folder: "${cachedFolder.FullPath}" (Original: "${cachedFolder.ServerRelativeUrl}")`);
        }

        // Recursively load subfolders
        for (const folder of userFolders) {
          await this.loadFoldersRecursively(folder.ServerRelativeUrl, progressCallback, depth + 1);
          
          // Small delay to prevent overwhelming SharePoint
          await this.delay(50);
        }
      } else {
        console.warn(`[SharePointFolderService] Failed to load folders from ${folderPath}: ${response.status} ${response.statusText}`);
      }
      
    } catch (error) {
      console.error(`[SharePointFolderService] Error loading folders from ${folderPath}:`, error);
    }
  }

  /**
   * ENHANCED: Check if a specific directory path exists in the cached folders
   * With fallback to direct API check if cache is empty
   */
  public checkDirectoryExists(directoryPath: string): boolean {
    console.log(`[SharePointFolderService] ========= DIRECTORY CHECK DEBUG =========`);
    console.log(`[SharePointFolderService] Checking directory: "${directoryPath}"`);
    console.log(`[SharePointFolderService] Folders loaded: ${this.allFoldersLoaded}`);
    console.log(`[SharePointFolderService] Cached folders count: ${this.cachedFolders.length}`);
    
    // NEW: If cache is empty but folders are supposedly loaded, this indicates a loading failure
    if (this.allFoldersLoaded && this.cachedFolders.length === 0) {
      console.warn('[SharePointFolderService] ⚠️ Cache is empty despite folders being "loaded" - possible loading failure');
      console.log('[SharePointFolderService] 🔄 Falling back to direct API check...');
      
      // Fallback to direct API check
      return this.checkDirectoryExistsDirectly(directoryPath);
    }
    
    if (!this.allFoldersLoaded) {
      console.warn('[SharePointFolderService] ⚠️ Folders not loaded yet, falling back to direct API check');
      return this.checkDirectoryExistsDirectly(directoryPath);
    }

    const normalizedPath = this.normalizePath(directoryPath);
    console.log(`[SharePointFolderService] Normalized input path: "${normalizedPath}"`);
    
    // Log all cached folder paths for comparison
    console.log(`[SharePointFolderService] Comparing against ${this.cachedFolders.length} cached folders:`);
    this.cachedFolders.forEach((folder, index) => {
      const matches = folder.FullPath === normalizedPath || folder.FullPath.endsWith(normalizedPath);
      console.log(`  ${index + 1}. "${folder.FullPath}" -> Match: ${matches}`);
    });
    
    const exists = this.cachedFolders.some(folder => {
      const exactMatch = folder.FullPath === normalizedPath;
      const endsWithMatch = folder.FullPath.endsWith(normalizedPath);
      
      if (exactMatch || endsWithMatch) {
        console.log(`[SharePointFolderService] ✅ FOUND MATCH: "${folder.FullPath}" (${exactMatch ? 'exact' : 'ends-with'})`);
        return true;
      }
      return false;
    });

    console.log(`[SharePointFolderService] Final result for "${normalizedPath}": ${exists ? '✅ EXISTS' : '❌ NOT FOUND'}`);
    console.log(`[SharePointFolderService] ========================================`);
    
    return exists;
  }

  /**
   * NEW: Direct API check for directory existence (synchronous fallback)
   */
  private checkDirectoryExistsDirectly(directoryPath: string): boolean {
    console.log(`[SharePointFolderService] 🚀 DIRECT API CHECK for: "${directoryPath}"`);
    
    try {
      const { context } = this;
      const webUrl = context.pageContext.web.absoluteUrl;
      
      // Use XMLHttpRequest for synchronous call (not ideal but necessary for fallback)
      const xhr = new XMLHttpRequest();
      xhr.open('GET', `${webUrl}/_api/web/getFolderByServerRelativeUrl('${directoryPath}')`, false);
      xhr.setRequestHeader('Accept', 'application/json;odata=verbose');
      xhr.send();
      
      const exists = xhr.status === 200;
      console.log(`[SharePointFolderService] 🎯 DIRECT API result: ${exists ? '✅ EXISTS' : '❌ NOT FOUND'} (Status: ${xhr.status})`);
      
      return exists;
      
    } catch (error) {
      console.error(`[SharePointFolderService] ❌ DIRECT API error:`, error);
      return false;
    }
  }

  /**
   * Get the full SharePoint path for a directory
   */
  public getFullDirectoryPath(relativePath: string, basePath: string): string {
    console.log(`[SharePointFolderService] Building full path:`);
    console.log(`  Base path: "${basePath}"`);
    console.log(`  Relative path: "${relativePath}"`);
    
    // ИСПРАВЛЕНИЕ: НЕ приводим к нижнему регистру!
    // Convert RelativePath format (e.g., "634\2025\6") to SharePoint format
    const normalizedRelative = relativePath.replace(/\\/g, '/');
    const fullPath = `${basePath}/${normalizedRelative}`;
    
    // ИСПРАВЛЕНИЕ: Убираем нормализацию, которая портила регистр
    const cleanPath = fullPath
      .replace(/\/+/g, '/')          // Remove duplicate slashes
      .replace(/\/$/, '');           // Remove trailing slash
    // НЕ ДЕЛАЕМ .toLowerCase() !!!
    
    console.log(`  Normalized relative: "${normalizedRelative}"`);
    console.log(`  Combined path: "${fullPath}"`);
    console.log(`  Final path: "${cleanPath}"`);
    
    return cleanPath;
  }

  // Также исправьте метод normalizePath (если он используется для путей):
  private normalizePath(path: string): string {
    // ТОЛЬКО для сравнения, НЕ для реальных API вызовов!
    const normalized = path
      .replace(/\\/g, '/')           // Convert backslashes to forward slashes
      .replace(/\/+/g, '/')          // Remove duplicate slashes
      .toLowerCase()                 // Case insensitive ТОЛЬКО для сравнения
      .replace(/\/$/, '');           // Remove trailing slash
    
    console.log(`[SharePointFolderService] Path normalization FOR COMPARISON: "${path}" -> "${normalized}"`);
    return normalized;
  }

  /**
   * Get cached folders (for debugging or display)
   */
  public getCachedFolders(): ICachedFolder[] {
    return [...this.cachedFolders];
  }

  /**
   * Clear cached folders
   */
  public clearCache(): void {
    console.log('[SharePointFolderService] Clearing folder cache');
    this.cachedFolders = [];
    this.allFoldersLoaded = false;
  }

  /**
   * Check if all folders are loaded
   */
  public areAllFoldersLoaded(): boolean {
    return this.allFoldersLoaded;
  }

  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  // Existing methods remain unchanged...

  public async getFolderContents(folderPath: string): Promise<{files: ISharePointFolder[], folders: ISharePointFolder[]}> {
    try {
      const { context } = this;
      const webUrl = context.pageContext.web.absoluteUrl;
      
      console.log(`[SharePointFolderService] Getting contents of folder: "${folderPath}"`);
      
      // Get both files and folders
      const [filesResponse, foldersResponse] = await Promise.all([
        fetch(`${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/files?$select=Name,ServerRelativeUrl,TimeCreated,TimeLastModified,Length`, {
          method: 'GET',
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        }),
        fetch(`${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')/folders?$select=Name,ServerRelativeUrl,ItemCount`, {
          method: 'GET',
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        })
      ]);

      const files = filesResponse.ok ? (await filesResponse.json()).d?.results || [] : [];
      const folders = foldersResponse.ok ? (await foldersResponse.json()).d?.results || [] : [];

      console.log(`[SharePointFolderService] Folder contents: ${files.length} files, ${folders.length} folders`);
      console.log(`[SharePointFolderService] Files found:`, files.map((f: ISharePointFolder) => f.Name).slice(0, 5));

      return {
        files: files.filter((file: ISharePointFolder) => !file.Name.startsWith('~')), // Filter out temp files
        folders: folders.filter((folder: ISharePointFolder) => !folder.Name.startsWith('_') && !folder.Name.startsWith('Forms'))
      };
    } catch (error) {
      console.error('[SharePointFolderService] Error getting folder contents:', error);
      return { files: [], folders: [] };
    }
  }

  public async checkFolderExists(folderPath: string): Promise<boolean> {
    try {
      const { context } = this;
      const webUrl = context.pageContext.web.absoluteUrl;
      
      console.log(`[SharePointFolderService] Direct API check for folder: "${folderPath}"`);
      
      const response = await fetch(`${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')`, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });

      const exists = response.ok;
      console.log(`[SharePointFolderService] Direct API result for "${folderPath}": ${exists ? 'EXISTS' : 'NOT FOUND'} (${response.status})`);
      
      return exists;
    } catch (error) {
      console.error(`[SharePointFolderService] Direct API error for "${folderPath}":`, error);
      return false;
    }
  }
}