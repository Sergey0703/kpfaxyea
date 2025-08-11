// src/webparts/xyea/services/SharePointFoldersService.ts - Updated to use PnPjs

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
  FolderChildCount: number;
  FileChildCount: number;
  Exists: boolean;
}

export class SharePointFoldersService {
  private sp: SPFI;
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Get all subfolders in the specified SharePoint folder path
   */
  public async getFoldersInPath(folderPath: string): Promise<ISharePointFolder[]> {
    try {
      console.log('[SharePointFoldersService] Getting folders for path:', folderPath);
      console.log('[SharePointFoldersService] Current web context:', {
        absoluteUrl: this.context.pageContext.web.absoluteUrl,
        serverRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
        title: this.context.pageContext.web.title
      });

      // Clean and validate the path
      const cleanPath = this.cleanFolderPath(folderPath);
      console.log('[SharePointFoldersService] Cleaned path:', cleanPath);

      let folders: ISharePointFolder[] = [];

      // Method 1: Try direct path access
      try {
        folders = await this.getFoldersMethod1(cleanPath);
        if (folders.length >= 0) { // 0 is valid (empty folder)
          console.log('[SharePointFoldersService] Method 1 successful, found:', folders.length, 'folders');
          return folders;
        }
      } catch (error) {
        console.warn('[SharePointFoldersService] Method 1 failed:', error);
      }

      // Method 2: Try cross-site access if path contains different site
      try {
        folders = await this.getFoldersMethod2(cleanPath);
        if (folders.length >= 0) {
          console.log('[SharePointFoldersService] Method 2 successful, found:', folders.length, 'folders');
          return folders;
        }
      } catch (error) {
        console.warn('[SharePointFoldersService] Method 2 failed:', error);
      }

      // Method 3: Try relative path resolution
      try {
        folders = await this.getFoldersMethod3(cleanPath);
        console.log('[SharePointFoldersService] Method 3 successful, found:', folders.length, 'folders');
        return folders;
      } catch (error) {
        console.error('[SharePointFoldersService] All methods failed. Last error:', error);
        throw new Error(`Failed to access folder "${cleanPath}". Please check the path and permissions.`);
      }

    } catch (error) {
      console.error('[SharePointFoldersService] Error getting folders:', error);
      throw new Error(`Failed to retrieve folders: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Method 1: Direct folder access using PnPjs
   */
  private async getFoldersMethod1(folderPath: string): Promise<ISharePointFolder[]> {
    try {
      console.log('[SharePointFoldersService] Method 1 - Direct folder access:', folderPath);
      
      // Use PnPjs to get folder by server relative path
      const folderInfo = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")();
      
      console.log('[SharePointFoldersService] Method 1 - Folder exists:', folderInfo);

      // Get subfolders using PnPjs - remove FolderChildCount and FileChildCount as they're not available
      const subfolders = await this.sp.web.getFolderByServerRelativePath(folderPath)
        .folders
        .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")
        .orderBy("Name")();

      console.log('[SharePointFoldersService] Method 1 - Raw subfolders:', subfolders);

      return this.mapAndFilterFolders(subfolders);
    } catch (error) {
      console.error('[SharePointFoldersService] Method 1 error:', error);
      throw error;
    }
  }

  /**
   * Method 2: Cross-site access for different sites (including root site collection)
   */
  private async getFoldersMethod2(folderPath: string): Promise<ISharePointFolder[]> {
    try {
      console.log('[SharePointFoldersService] Method 2 - Cross-site access:', folderPath);
      
      const currentSiteUrl = this.context.pageContext.web.serverRelativeUrl;
      const tenantUrl = this.context.pageContext.web.absoluteUrl.split('/sites/')[0];
      
      console.log('[SharePointFoldersService] Method 2 - URLs:', {
        currentSiteUrl,
        tenantUrl,
        folderPath
      });
      
      // Check if this is a cross-site request or root site collection access
      let targetSiteAbsoluteUrl: string;
      
      if (!folderPath.startsWith(currentSiteUrl)) {
        // Check if it's a path to a different subsite
        const pathParts = folderPath.split('/');
        if (pathParts.length >= 3 && pathParts[1] === 'sites') {
          const targetSiteName = pathParts[2];
          targetSiteAbsoluteUrl = `${tenantUrl}/sites/${targetSiteName}`;
        } 
        // Check if it's a root site collection path (like /Shared Documents)
        else if (folderPath.startsWith('/Shared Documents') || folderPath.startsWith('/Documents')) {
          targetSiteAbsoluteUrl = tenantUrl; // Root site collection
          console.log('[SharePointFoldersService] Method 2 - Detected root site collection access');
        }
        // Check if it's any root-level path
        else if (!folderPath.includes('/sites/')) {
          targetSiteAbsoluteUrl = tenantUrl; // Root site collection
          console.log('[SharePointFoldersService] Method 2 - Detected root-level path');
        }
        else {
          throw new Error('Invalid cross-site path format');
        }
        
        console.log('[SharePointFoldersService] Method 2 - Target site URL:', targetSiteAbsoluteUrl);
        
        // Create new SPFI instance for the target site
        const targetSp = spfi(targetSiteAbsoluteUrl).using(SPFx(this.context));
        
        // Get subfolders from target site
        const subfolders = await targetSp.web.getFolderByServerRelativePath(folderPath)
          .folders
          .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")
          .orderBy("Name")();

        console.log('[SharePointFoldersService] Method 2 - Cross-site subfolders:', subfolders);
        return this.mapAndFilterFolders(subfolders);
      }
      
      throw new Error('Not a cross-site request');
    } catch (error) {
      console.error('[SharePointFoldersService] Method 2 error:', error);
      throw error;
    }
  }

  /**
   * Method 3: Relative path resolution and root site collection access
   */
  private async getFoldersMethod3(folderPath: string): Promise<ISharePointFolder[]> {
    try {
      console.log('[SharePointFoldersService] Method 3 - Relative path resolution:', folderPath);
      
      const currentSiteUrl = this.context.pageContext.web.serverRelativeUrl;
      const tenantUrl = this.context.pageContext.web.absoluteUrl.split('/sites/')[0];
      
      // Try different path combinations including root site collection
      const pathsToTry = [
        folderPath,
        `${currentSiteUrl}${folderPath}`,
        `${currentSiteUrl}/${folderPath.replace(/^\/+/, '')}`,
        folderPath.replace(currentSiteUrl, ''),
        `/Shared Documents`, // Root site collection default
        `/Documents`, // Alternative root document library name
        // For root site collection, try with different SPFI instance
      ];

      // First try with current site instance
      for (const tryPath of pathsToTry.slice(0, -2)) {
        if (!tryPath || tryPath === currentSiteUrl) continue;
        
        try {
          console.log('[SharePointFoldersService] Method 3 - Trying path with current site:', tryPath);
          
          const subfolders = await this.sp.web.getFolderByServerRelativePath(tryPath)
            .folders
            .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")
            .orderBy("Name")();

          console.log('[SharePointFoldersService] Method 3 - Success with current site path:', tryPath, 'Found:', subfolders.length, 'folders');
          return this.mapAndFilterFolders(subfolders);
        } catch (pathError) {
          console.log('[SharePointFoldersService] Method 3 - Current site path failed:', tryPath, pathError.message);
          continue;
        }
      }
      
      // If current site failed, try with root site collection instance
      console.log('[SharePointFoldersService] Method 3 - Trying root site collection access');
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
            console.log('[SharePointFoldersService] Method 3 - Trying root site path:', rootPath);
            
            const subfolders = await rootSp.web.getFolderByServerRelativePath(rootPath)
              .folders
              .select("Name", "ServerRelativeUrl", "ItemCount", "TimeCreated", "TimeLastModified", "Exists")
              .orderBy("Name")();

            console.log('[SharePointFoldersService] Method 3 - Success with root site path:', rootPath, 'Found:', subfolders.length, 'folders');
            return this.mapAndFilterFolders(subfolders);
          } catch (rootPathError) {
            console.log('[SharePointFoldersService] Method 3 - Root site path failed:', rootPath, rootPathError.message);
            continue;
          }
        }
      } catch (rootError) {
        console.error('[SharePointFoldersService] Method 3 - Root site access failed:', rootError);
      }
      
      throw new Error('All path combinations failed');
    } catch (error) {
      console.error('[SharePointFoldersService] Method 3 error:', error);
      throw error;
    }
  }

  /**
   * Clean and normalize the folder path
   */
  private cleanFolderPath(path: string): string {
    if (!path || typeof path !== 'string') {
      throw new Error('Folder path is required and must be a string');
    }

    let cleanPath = path.trim();

    // Remove leading and trailing slashes for consistency, then add back leading slash
    cleanPath = cleanPath.replace(/^\/+|\/+$/g, '');
    
    if (!cleanPath) {
      throw new Error('Folder path cannot be empty');
    }

    // Add leading slash back
    cleanPath = '/' + cleanPath;

    return cleanPath;
  }

  /**
   * Map SharePoint response to our interface and filter system folders
   */
  private mapAndFilterFolders(folders: any[]): ISharePointFolder[] {
    console.log('[SharePointFoldersService] Mapping folders:', folders);
    
    return folders
      .filter(folder => this.shouldIncludeFolder(folder))
      .map(folder => ({
        Name: folder.Name || '',
        ServerRelativeUrl: folder.ServerRelativeUrl || '',
        ItemCount: folder.ItemCount || 0,
        TimeCreated: folder.TimeCreated || new Date().toISOString(),
        TimeLastModified: folder.TimeLastModified || new Date().toISOString(),
        FolderChildCount: 0, // PnPjs doesn't provide this property, set to 0
        FileChildCount: 0,   // PnPjs doesn't provide this property, set to 0
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

    // Filter out folders that start with underscore (system folders)
    if (folder.Name.startsWith('_')) {
      return false;
    }

    // Filter out known system folders
    if (systemFolders.includes(folder.Name)) {
      return false;
    }

    // Only include folders that exist
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
        FolderChildCount: 0, // Not available in PnPjs, set to 0
        FileChildCount: 0,   // Not available in PnPjs, set to 0
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

    // Check for invalid characters
    const invalidChars = /[<>:"|?*]/;
    if (invalidChars.test(trimmedPath)) {
      return { isValid: false, error: 'Path contains invalid characters: < > : " | ? *' };
    }

    // Check for double slashes (except at the beginning)
    if (trimmedPath.includes('//')) {
      return { isValid: false, error: 'Path cannot contain double slashes' };
    }

    return { isValid: true };
  }
}