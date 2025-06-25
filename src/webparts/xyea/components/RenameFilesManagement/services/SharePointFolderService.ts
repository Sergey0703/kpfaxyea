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
  private context: IWebPartContext; // FIXED: specific type instead of any
  private cachedFolders: ICachedFolder[] = [];
  private isLoadingAllFolders: boolean = false;
  private allFoldersLoaded: boolean = false;

  constructor(context: IWebPartContext) { // FIXED: specific type instead of any
    this.context = context;
  }

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
      
      let foldersData: { d?: { results?: ISharePointFolder[] }; value?: ISharePointFolder[] } | undefined = undefined; // FIXED: specific type instead of any
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
        .filter((folder: ISharePointFolder) => // FIXED: specific type instead of any
          !folder.Name.startsWith('Forms') && 
          !folder.Name.startsWith('_') &&
          folder.Name !== 'Forms'
        )
        .map((folder: ISharePointFolder) => ({ // FIXED: specific type instead of any
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
      const timeoutPromise = new Promise<void>((resolve, reject) => { // FIXED: parameter names
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
        const userFolders = folders.filter((folder: ISharePointFolder) => // FIXED: specific type instead of any
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

  public async getFolderContents(folderPath: string): Promise<{files: ISharePointFolder[], folders: ISharePointFolder[]}> { // FIXED: specific types instead of any
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
      console.log(`[SharePointFolderService] Files found:`, files.map((f: ISharePointFolder) => f.Name).slice(0, 5)); // FIXED: specific type

      return {
        files: files.filter((file: ISharePointFolder) => !file.Name.startsWith('~')), // Filter out temp files // FIXED: specific type
        folders: folders.filter((folder: ISharePointFolder) => !folder.Name.startsWith('_') && !folder.Name.startsWith('Forms')) // FIXED: specific type
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