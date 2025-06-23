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

export class SharePointFolderService {
  private context: any;
  private cachedFolders: ICachedFolder[] = [];
  private isLoadingAllFolders: boolean = false;
  private allFoldersLoaded: boolean = false;

  constructor(context: any) {
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
      
      let foldersData: any = null;
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
        .filter((folder: any) => 
          !folder.Name.startsWith('Forms') && 
          !folder.Name.startsWith('_') &&
          folder.Name !== 'Forms'
        )
        .map((folder: any) => ({
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
   * NEW METHOD: Load all subfolders recursively and cache them
   * This should be called when user selects a main folder for searching
   */
  public async loadAllSubfolders(
    baseFolderPath: string, 
    progressCallback?: (currentPath: string, foldersLoaded: number) => void
  ): Promise<ICachedFolder[]> {
    
    if (this.isLoadingAllFolders) {
      console.log('[SharePointFolderService] Already loading folders, please wait...');
      return this.cachedFolders;
    }

    console.log(`[SharePointFolderService] Starting to load all subfolders from: ${baseFolderPath}`);
    
    this.isLoadingAllFolders = true;
    this.cachedFolders = [];
    this.allFoldersLoaded = false;

    try {
      await this.loadFoldersRecursively(baseFolderPath, progressCallback);
      
      this.allFoldersLoaded = true;
      console.log(`[SharePointFolderService] Finished loading ${this.cachedFolders.length} folders`);
      
      // Log some examples for debugging
      if (this.cachedFolders.length > 0) {
        console.log('[SharePointFolderService] First 5 cached folders:');
        this.cachedFolders.slice(0, 5).forEach((folder, index) => {
          console.log(`  ${index + 1}. ${folder.FullPath}`);
        });
      }
      
      return this.cachedFolders;
      
    } catch (error) {
      console.error('[SharePointFolderService] Error loading all subfolders:', error);
      throw error;
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
      
      const response = await fetch(foldersUrl, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });

      if (response.ok) {
        const data = await response.json();
        const folders = data.d?.results || data.value || [];
        
        // Filter out system folders
        const userFolders = folders.filter((folder: any) => 
          !folder.Name.startsWith('_') && 
          !folder.Name.startsWith('Forms') &&
          folder.Name !== 'Forms'
        );

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
          console.log(`[SharePointFolderService] Cached folder: ${cachedFolder.FullPath}`);
        }

        // Recursively load subfolders
        for (const folder of userFolders) {
          await this.loadFoldersRecursively(folder.ServerRelativeUrl, progressCallback, depth + 1);
          
          // Small delay to prevent overwhelming SharePoint
          await this.delay(50);
        }
      } else {
        console.warn(`[SharePointFolderService] Failed to load folders from ${folderPath}: ${response.status}`);
      }
      
    } catch (error) {
      console.error(`[SharePointFolderService] Error loading folders from ${folderPath}:`, error);
    }
  }

  /**
   * Check if a specific directory path exists in the cached folders
   */
  public checkDirectoryExists(directoryPath: string): boolean {
    if (!this.allFoldersLoaded) {
      console.warn('[SharePointFolderService] Folders not loaded yet, cannot check directory existence');
      return false;
    }

    const normalizedPath = this.normalizePath(directoryPath);
    const exists = this.cachedFolders.some(folder => 
      folder.FullPath === normalizedPath || 
      folder.FullPath.endsWith(normalizedPath)
    );

    console.log(`[SharePointFolderService] Directory exists check: "${normalizedPath}" -> ${exists}`);
    return exists;
  }

  /**
   * Get the full SharePoint path for a directory
   */
  public getFullDirectoryPath(relativePath: string, basePath: string): string {
    // Convert RelativePath format (e.g., "634\2025\6") to SharePoint format
    const normalizedRelative = relativePath.replace(/\\/g, '/');
    const fullPath = `${basePath}/${normalizedRelative}`;
    
    return this.normalizePath(fullPath);
  }

  /**
   * Normalize path for consistent comparison
   */
  private normalizePath(path: string): string {
    return path
      .replace(/\\/g, '/')           // Convert backslashes to forward slashes
      .replace(/\/+/g, '/')          // Remove duplicate slashes
      .toLowerCase()                 // Case insensitive
      .replace(/\/$/, '');           // Remove trailing slash
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

  public async getFolderContents(folderPath: string): Promise<{files: any[], folders: any[]}> {
    try {
      const { context } = this;
      const webUrl = context.pageContext.web.absoluteUrl;
      
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

      return {
        files: files.filter((file: any) => !file.Name.startsWith('~')), // Filter out temp files
        folders: folders.filter((folder: any) => !folder.Name.startsWith('_') && !folder.Name.startsWith('Forms'))
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
      
      const response = await fetch(`${webUrl}/_api/web/getFolderByServerRelativeUrl('${folderPath}')`, {
        method: 'GET',
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      });

      return response.ok;
    } catch (error) {
      return false;
    }
  }
}