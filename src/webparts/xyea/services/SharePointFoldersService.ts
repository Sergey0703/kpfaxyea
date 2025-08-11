// src/webparts/xyea/services/SharePointFoldersService.ts

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/files";
import "@pnp/sp/folders";

export interface ISharePointFolder {
  Name: string;
  ServerRelativeUrl: string;
  ItemCount: number;
  TimeCreated: string;
  TimeLastModified: string;
  Exists: boolean;
}

export class SharePointFoldersService {
  private sp: SPFI;
  private context: WebPartContext;
  private readonly REQUEST_DELAY = 100; // Delay between requests in ms

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Get complete recursive folder structure with NO DEPTH LIMIT
   */
  public async getFoldersInPath(folderPath: string): Promise<ISharePointFolder[]> {
    try {
      console.log('[SharePointFoldersService] Starting COMPLETE recursive scan for path:', folderPath);
      console.log('[SharePointFoldersService] Current web context:', {
        absoluteUrl: this.context.pageContext.web.absoluteUrl,
        serverRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
        title: this.context.pageContext.web.title
      });

      const cleanPath = this.cleanFolderPath(folderPath);
      console.log('[SharePointFoldersService] Cleaned path:', cleanPath);

      const structure: ISharePointFolder[] = [];
      
      // Start UNLIMITED recursive exploration - NO DEPTH LIMIT!
      await this.exploreDirectoryRecursiveUnlimited(cleanPath, 0, structure);
      
      console.log('[SharePointFoldersService] COMPLETE recursive exploration finished:', {
        totalItems: structure.length,
        files: structure.filter(item => (item as any).IsFile === true).length,
        folders: structure.filter(item => (item as any).IsFile === false).length,
        maxLevelReached: Math.max(...structure.map(item => (item as any).Level || 0))
      });
      
      return structure;

    } catch (error) {
      console.error('[SharePointFoldersService] Error getting complete recursive structure:', error);
      throw new Error(`Failed to retrieve complete folder structure: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * UNLIMITED recursive function to explore ALL directory levels (no depth limit)
   */
  private async exploreDirectoryRecursiveUnlimited(
    currentPath: string,
    currentLevel: number,
    structure: ISharePointFolder[]
  ): Promise<void> {
    
    try {
      console.log(`[SharePointFoldersService] Level ${currentLevel}: Exploring ${currentPath}`);
      
      // Add a small delay to prevent overwhelming SharePoint
      if (currentLevel > 0) {
        await this.delay(this.REQUEST_DELAY);
      }
      
      // Get folders and files in current directory
      const [folders, files] = await Promise.all([
        this.getFoldersOnly(currentPath),
        this.getFilesOnly(currentPath)
      ]);
      
      console.log(`[SharePointFoldersService] Level ${currentLevel}: ${currentPath} - ${files.length} files, ${folders.length} folders`);
      
      // 1. FIRST: Add all FILES (files first, like directory listing)
      files.forEach(file => {
        structure.push({
          Name: file.Name,
          ServerRelativeUrl: file.ServerRelativeUrl,
          ItemCount: file.ItemCount,
          TimeCreated: file.TimeCreated,
          TimeLastModified: file.TimeLastModified,
          Exists: true,
          // Add metadata for tree display
          ['Level' as any]: currentLevel,
          ['IsFile' as any]: true
        } as any);
      });
      
      // 2. SECOND: Add all FOLDERS and recurse into them (NO DEPTH LIMIT!)
      for (let i = 0; i < folders.length; i++) {
        const folder = folders[i];
        
        // Add folder to structure with '/' suffix
        structure.push({
          Name: folder.Name + '/',
          ServerRelativeUrl: folder.ServerRelativeUrl,
          ItemCount: folder.ItemCount,
          TimeCreated: folder.TimeCreated,
          TimeLastModified: folder.TimeLastModified,
          Exists: folder.Exists,
          ['Level' as any]: currentLevel,
          ['IsFile' as any]: false
        } as any);
        
        // Recurse into subdirectory - NO DEPTH LIMIT!
        try {
          await this.exploreDirectoryRecursiveUnlimited(
            folder.ServerRelativeUrl,
            currentLevel + 1,
            structure
          );
        } catch (subDirError) {
          console.warn(`[SharePointFoldersService] Failed to access subdirectory at level ${currentLevel + 1}: ${folder.ServerRelativeUrl}`, subDirError);
          
          // Add permission denied indicator
          structure.push({
            Name: '[Permission Denied]',
            ServerRelativeUrl: folder.ServerRelativeUrl + '/[denied]',
            ItemCount: 0,
            TimeCreated: new Date().toISOString(),
            TimeLastModified: new Date().toISOString(),
            Exists: false,
            ['Level' as any]: currentLevel + 1,
            ['IsFile' as any]: true
          } as any);
        }
        
        // Add delay between processing folders to prevent rate limiting
        if (i < folders.length - 1) {
          await this.delay(this.REQUEST_DELAY / 2);
        }
      }
      
    } catch (error) {
      console.error(`[SharePointFoldersService] Error exploring level ${currentLevel} at ${currentPath}:`, error);
      
      // Add error indicator to structure
      structure.push({
        Name: `[Error: ${error instanceof Error ? error.message : 'Unknown'}]`,
        ServerRelativeUrl: currentPath + '/[error]',
        ItemCount: 0,
        TimeCreated: new Date().toISOString(),
        TimeLastModified: new Date().toISOString(),
        Exists: false,
        ['Level' as any]: currentLevel,
        ['IsFile' as any]: true
      } as any);
    }
  }

  /**
   * Get only folders from a path with multiple fallback methods
   */
  private async getFoldersOnly(folderPath: string): Promise<any[]> {
    try {
      // Method 1: Try direct folder access
      try {
        const folders = await this.getFoldersMethod1(folderPath);
        if (folders.length >= 0) return folders;
      } catch (error) {
        console.warn('[SharePointFoldersService] Method 1 failed for folders:', error);
      }
      
      // Method 2: Try cross-site access
      try {
        const folders = await this.getFoldersMethod2(folderPath);
        if (folders.length >= 0) return folders;
      } catch (error) {
        console.warn('[SharePointFoldersService] Method 2 failed for folders:', error);
      }
      
      // Method 3: Try relative path resolution
      try {
        const folders = await this.getFoldersMethod3(folderPath);
        return folders;
      } catch (error) {
        console.warn('[SharePointFoldersService] All methods failed for folders:', error);
        return [];
      }
    } catch (error) {
      console.warn(`[SharePointFoldersService] Could not get folders for path: ${folderPath}`, error);
      return [];
    }
  }

  /**
   * Get only files from a path with fallback methods
   */
  private async getFilesOnly(folderPath: string): Promise<any[]> {
    try {
      // Use current SPFI instance first
      let files = await this.getFilesWithSp(this.sp, folderPath);
      if (files.length >= 0) return files;
      
      // If failed, try with cross-site access
      const tenantUrl = this.context.pageContext.web.absoluteUrl.split('/sites/')[0];
      if (!folderPath.startsWith(this.context.pageContext.web.serverRelativeUrl)) {
        const targetSp = spfi(tenantUrl).using(SPFx(this.context));
        files = await this.getFilesWithSp(targetSp, folderPath);
      }
      
      return files;
    } catch (error) {
      console.warn(`[SharePointFoldersService] Could not get files for path: ${folderPath}`, error);
      return [];
    }
  }

  /**
   * Helper to get files with specific SP instance
   */
  private async getFilesWithSp(sp: SPFI, folderPath: string): Promise<any[]> {
    try {
      const files = await sp.web.getFolderByServerRelativePath(folderPath)
        .files
        .select("Name", "ServerRelativeUrl", "TimeCreated", "TimeLastModified", "Length")
        .orderBy("Name")();

      return files
        .filter(file => !file.Name.startsWith('.') && !file.Name.startsWith('~')) // Filter hidden/temp files
        .map(file => ({
          Name: file.Name,
          ServerRelativeUrl: file.ServerRelativeUrl,
          ItemCount: file.Length || 0,
          TimeCreated: file.TimeCreated,
          TimeLastModified: file.TimeLastModified,
          Exists: true
        }));
    } catch (error) {
      throw error;
    }
  }

  /**
   * Method 1: Direct folder access using PnPjs
   */
  private async getFoldersMethod1(folderPath: string): Promise<any[]> {
    try {
      console.log('[SharePointFoldersService] Method 1 - Direct folder access:', folderPath);
      
      const subfolders = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .folders
        .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")
        .orderBy("Name")();

      return this.mapAndFilterFolders(subfolders);
    } catch (error) {
      throw error;
    }
  }

  /**
   * Method 2: Cross-site access for different sites
   */
  private async getFoldersMethod2(folderPath: string): Promise<any[]> {
    try {
      console.log('[SharePointFoldersService] Method 2 - Cross-site access:', folderPath);
      
      const currentSiteUrl = this.context.pageContext.web.serverRelativeUrl;
      const tenantUrl = this.context.pageContext.web.absoluteUrl.split('/sites/')[0];
      
      if (!folderPath.startsWith(currentSiteUrl)) {
        let targetSiteAbsoluteUrl: string;
        
        const pathParts = folderPath.split('/');
        if (pathParts.length >= 3 && pathParts[1] === 'sites') {
          const targetSiteName = pathParts[2];
          targetSiteAbsoluteUrl = `${tenantUrl}/sites/${targetSiteName}`;
        } 
        else if (folderPath.startsWith('/Shared Documents') || folderPath.startsWith('/Documents')) {
          targetSiteAbsoluteUrl = tenantUrl; // Root site collection
        }
        else if (!folderPath.includes('/sites/')) {
          targetSiteAbsoluteUrl = tenantUrl; // Root site collection
        }
        else {
          throw new Error('Invalid cross-site path format');
        }
        
        const targetSp = spfi(targetSiteAbsoluteUrl).using(SPFx(this.context));
        
        const subfolders = await targetSp.web.getFolderByServerRelativePath(folderPath)
          .folders
          .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")
          .orderBy("Name")();

        return this.mapAndFilterFolders(subfolders);
      }
      
      throw new Error('Not a cross-site request');
    } catch (error) {
      throw error;
    }
  }

  /**
   * Method 3: Relative path resolution and root site collection access
   */
  private async getFoldersMethod3(folderPath: string): Promise<any[]> {
    try {
      console.log('[SharePointFoldersService] Method 3 - Relative path resolution:', folderPath);
      
      const currentSiteUrl = this.context.pageContext.web.serverRelativeUrl;
      const tenantUrl = this.context.pageContext.web.absoluteUrl.split('/sites/')[0];
      
      const pathsToTry = [
        folderPath,
        `${currentSiteUrl}${folderPath}`,
        `${currentSiteUrl}/${folderPath.replace(/^\/+/, '')}`,
        folderPath.replace(currentSiteUrl, ''),
        `/Shared Documents`
      ];

      // First try with current site instance
      for (const tryPath of pathsToTry.slice(0, -1)) {
        if (!tryPath || tryPath === currentSiteUrl) continue;
        
        try {
          const subfolders = await this.sp.web.getFolderByServerRelativePath(tryPath)
            .folders
            .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")
            .orderBy("Name")();

          return this.mapAndFilterFolders(subfolders);
        } catch (pathError) {
          continue;
        }
      }
      
      // If current site failed, try with root site collection instance
      try {
        const rootSp = spfi(tenantUrl).using(SPFx(this.context));
        
        const rootPathsToTry = [
          folderPath,
          `/Shared Documents`,
          `/Documents`,
          `/sites/root/Shared Documents`
        ];
        
        for (const rootPath of rootPathsToTry) {
          if (!rootPath) continue;
          
          try {
            const subfolders = await rootSp.web.getFolderByServerRelativePath(rootPath)
              .folders
              .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")
              .orderBy("Name")();

            return this.mapAndFilterFolders(subfolders);
          } catch (rootPathError) {
            continue;
          }
        }
      } catch (rootError) {
        console.error('[SharePointFoldersService] Root site access failed:', rootError);
      }
      
      throw new Error('All path combinations failed');
    } catch (error) {
      throw error;
    }
  }

  /**
   * Utility method for delays to prevent rate limiting
   */
  private delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Clean and normalize the folder path
   */
  private cleanFolderPath(path: string): string {
    if (!path || typeof path !== 'string') {
      throw new Error('Folder path is required and must be a string');
    }

    let cleanPath = path.trim();
    cleanPath = cleanPath.replace(/^\/+|\/+$/g, '');
    
    if (!cleanPath) {
      throw new Error('Folder path cannot be empty');
    }

    cleanPath = '/' + cleanPath;
    return cleanPath;
  }

  /**
   * Map SharePoint response to our interface and filter system folders
   */
  private mapAndFilterFolders(folders: any[]): any[] {
    return folders
      .filter(folder => this.shouldIncludeFolder(folder))
      .map(folder => ({
        Name: folder.Name || '',
        ServerRelativeUrl: folder.ServerRelativeUrl || '',
        ItemCount: folder.ItemCount || 0,
        TimeCreated: folder.TimeCreated || new Date().toISOString(),
        TimeLastModified: folder.TimeLastModified || new Date().toISOString(),
        Exists: folder.Exists !== false
      }))
      .sort((a, b) => a.Name.localeCompare(b.Name));
  }

  /**
   * Determine if a folder should be included in the results
   */
  private shouldIncludeFolder(folder: any): boolean {
    if (!folder || !folder.Name) {
      return false;
    }

    const systemFolders = [
      'Forms',
      '_private',
      '_catalogs',
      '_layouts',
      'Style Library',
      'Site Assets',
      'Site Pages',
      'Master Page Gallery',
      'Theme Gallery',
      'Web Part Gallery',
      'List Template Gallery',
      'Solution Gallery'
    ];

    if (folder.Name.startsWith('_')) {
      return false;
    }

    if (systemFolders.includes(folder.Name)) {
      return false;
    }

    if (folder.Exists === false) {
      return false;
    }

    return true;
  }

  /**
   * Check if a folder exists at the given path
   */
  public async checkFolderExists(folderPath: string): Promise<boolean> {
    try {
      const cleanPath = this.cleanFolderPath(folderPath);
      const folderInfo = await this.sp.web.getFolderByServerRelativePath(cleanPath)
        .select("Exists")();
      return folderInfo.Exists === true;
    } catch (error) {
      console.warn('[SharePointFoldersService] Error checking folder existence:', error);
      return false;
    }
  }

  /**
   * Get folder details for a specific path
   */
  public async getFolderDetails(folderPath: string): Promise<ISharePointFolder | null> {
    try {
      const cleanPath = this.cleanFolderPath(folderPath);
      const folder = await this.sp.web.getFolderByServerRelativePath(cleanPath)
        .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")();
      
      return {
        Name: folder.Name,
        ServerRelativeUrl: folder.ServerRelativeUrl,
        ItemCount: folder.ItemCount || 0,
        TimeCreated: folder.TimeCreated,
        TimeLastModified: folder.TimeLastModified,
        Exists: folder.Exists !== false
      };
    } catch (error) {
      console.error('[SharePointFoldersService] Error getting folder details:', error);
      return null;
    }
  }

  /**
   * Get the current web's server relative URL for path resolution
   */
  public getCurrentWebPath(): string {
    return this.context.pageContext.web.serverRelativeUrl;
  }

  /**
   * Get current web information for debugging
   */
  public getWebInfo(): { absoluteUrl: string; serverRelativeUrl: string; title: string } {
    return {
      absoluteUrl: this.context.pageContext.web.absoluteUrl,
      serverRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
      title: this.context.pageContext.web.title
    };
  }

  /**
   * Validate if a path looks like a valid SharePoint path
   */
  public validateSharePointPath(path: string): { isValid: boolean; error?: string } {
    if (!path || typeof path !== 'string') {
      return { isValid: false, error: 'Path is required' };
    }

    const trimmedPath = path.trim();
    
    if (!trimmedPath) {
      return { isValid: false, error: 'Path cannot be empty' };
    }

    const invalidChars = /[<>:"|?*]/;
    if (invalidChars.test(trimmedPath)) {
      return { isValid: false, error: 'Path contains invalid characters: < > : " | ? *' };
    }

    if (trimmedPath.includes('//')) {
      return { isValid: false, error: 'Path cannot contain double slashes' };
    }

    return { isValid: true };
  }
}