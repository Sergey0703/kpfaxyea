// src/webparts/xyea/services/FolderStructureExportService.ts

import * as XLSX from 'xlsx';
import { ISharePointFolder } from './SharePointFoldersService';

// Type definitions for better type safety
type CellValue = string | number | boolean | Date | undefined;

interface IFolderStructureExportSettings {
  fileName: string;
  includeHeaders: boolean;
  includeHierarchy: boolean; // Show indentation/levels
  includeMetadata: boolean; // Include size, dates, etc.
  fileFormat: 'xlsx' | 'csv';
  maxLevels?: number; // Optional depth limit for export
}

interface IFolderStructureExportStatistics {
  totalItems: number;
  totalFolders: number;
  totalFiles: number;
  maxDepth: number;
  estimatedFileSize: string;
  canExport: boolean;
}

interface IColumnWidth {
  wch: number;
}

export class FolderStructureExportService {

  /**
   * Export folder structure to Excel/CSV
   */
  public static async exportFolderStructure(
    folderPath: string,
    folders: ISharePointFolder[],
    settings: IFolderStructureExportSettings
  ): Promise<{ success: boolean; fileName?: string; error?: string }> {
    try {
      console.log('[FolderStructureExportService] Starting folder structure export:', {
        folderPath,
        totalItems: folders.length,
        settings
      });

      if (folders.length === 0) {
        return {
          success: false,
          error: 'No folder structure data to export.'
        };
      }

      // Generate filename
      const fileName = this.generateExportFileName(folderPath, settings);

      // Prepare export data
      const exportData = this.prepareFolderStructureExportData(folders, settings);

      // Create and download file
      const blob = await this.createExportFile(exportData, settings);
      this.downloadFile(blob, fileName);

      console.log('[FolderStructureExportService] Export completed:', {
        fileName,
        rowsExported: exportData.length - (settings.includeHeaders ? 1 : 0)
      });

      return {
        success: true,
        fileName
      };

    } catch (error) {
      console.error('[FolderStructureExportService] Export failed:', error);
      return {
        success: false,
        error: error instanceof Error ? error.message : 'Export failed'
      };
    }
  }

  /**
   * Get export statistics
   */
  public static getFolderStructureExportStatistics(
    folders: ISharePointFolder[]
  ): IFolderStructureExportStatistics {
    
    const totalFiles = folders.filter(item => (item as any).IsFile === true).length;
    const totalFolders = folders.filter(item => (item as any).IsFile === false).length;
    const maxDepth = folders.length > 0 ? Math.max(...folders.map(item => (item as any).Level || 0)) : 0;
    
    // Estimate file size (approximate)
    const avgCellSize = 25; // bytes per cell
    const estimatedBytes = folders.length * 6 * avgCellSize; // 6 columns average
    const estimatedFileSize = this.formatFileSize(estimatedBytes);

    return {
      totalItems: folders.length,
      totalFolders,
      totalFiles,
      maxDepth: maxDepth + 1, // +1 because levels are 0-based
      estimatedFileSize,
      canExport: folders.length > 0
    };
  }

  /**
   * Generate export filename
   */
  private static generateExportFileName(
    folderPath: string, 
    settings: IFolderStructureExportSettings
  ): string {
    
    // Clean the folder path for filename
    const pathForFilename = folderPath
      .replace(/[^a-zA-Z0-9]/g, '_')
      .replace(/_{2,}/g, '_')
      .replace(/^_|_$/g, '');
    
    // Base name
    let baseName = settings.fileName;
    if (pathForFilename) {
      baseName += `_${pathForFilename}`;
    }
    
    // Add timestamp
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, '_');
    const cleanName = baseName.replace(/[^a-zA-Z0-9_-]/g, '_');
    
    // Determine extension
    const extension = settings.fileFormat === 'csv' ? 'csv' : 'xlsx';
    
    return `${cleanName}_${timestamp}.${extension}`;
  }

  /**
   * Prepare folder structure data for export
   */
  private static prepareFolderStructureExportData(
    folders: ISharePointFolder[],
    settings: IFolderStructureExportSettings
  ): CellValue[][] {
    
    const exportData: CellValue[][] = [];
    
    // Prepare headers
    if (settings.includeHeaders) {
      const headers: string[] = [];
      
      if (settings.includeHierarchy) {
        headers.push('Level');
        headers.push('Hierarchy');
      }
      
      headers.push('Name');
      headers.push('Type');
      headers.push('Path');
      
      if (settings.includeMetadata) {
        headers.push('Size/Item Count');
        headers.push('Created');
        headers.push('Modified');
      }
      
      exportData.push(headers);
    }
    
    // Process each folder/file item
    folders.forEach(item => {
      const level = (item as any).Level || 0;
      const isFile = (item as any).IsFile === true;
      
      // Apply level filter if specified
      if (settings.maxLevels && level >= settings.maxLevels) {
        return; // Skip items beyond max level
      }
      
      const rowData: CellValue[] = [];
      
      // Level and hierarchy
      if (settings.includeHierarchy) {
        rowData.push(level);
        
        // Create visual hierarchy with indentation
        const indent = '  '.repeat(level); // 2 spaces per level
        const icon = isFile ? '📄' : '📁';
        const displayName = `${indent}${icon} ${item.Name}`;
        rowData.push(displayName);
      }
      
      // Basic info
      rowData.push(item.Name);
      rowData.push(isFile ? 'File' : 'Folder');
      rowData.push(item.ServerRelativeUrl);
      
      // Metadata
      if (settings.includeMetadata) {
        // Size/Item Count
        if (isFile && item.ItemCount > 0) {
          rowData.push(this.formatFileSize(item.ItemCount));
        } else if (!isFile) {
          rowData.push(`${item.ItemCount || 0} items`);
        } else {
          rowData.push('');
        }
        
        // Dates
        rowData.push(this.formatDate(item.TimeCreated));
        rowData.push(this.formatDate(item.TimeLastModified));
      }
      
      exportData.push(rowData);
    });
    
    return exportData;
  }

  /**
   * Create export file (Excel or CSV)
   */
  private static async createExportFile(
    data: CellValue[][],
    settings: IFolderStructureExportSettings
  ): Promise<Blob> {
    
    if (settings.fileFormat === 'csv') {
      return this.createCSVBlob(data);
    } else {
      return this.createExcelBlob(data, 'Folder Structure');
    }
  }

  /**
   * Create CSV blob from data array
   */
  private static createCSVBlob(data: CellValue[][]): Blob {
    const csvContent = data
      .map(row => 
        row.map(cell => {
          const cellValue = String(cell || '');
          // Escape quotes and wrap in quotes if contains comma, quote, or newline
          if (cellValue.includes(',') || cellValue.includes('"') || cellValue.includes('\n')) {
            return `"${cellValue.replace(/"/g, '""')}"`;
          }
          return cellValue;
        }).join(',')
      )
      .join('\n');
    
    return new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  }

  /**
   * Create Excel blob from data array
   */
  private static createExcelBlob(data: CellValue[][], sheetName: string = 'Folder Structure'): Blob {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    
    // Auto-adjust column widths
    const columnWidths = this.calculateColumnWidths(data);
    worksheet['!cols'] = columnWidths;
    
    // Add some basic styling for hierarchy column if present
    if (data.length > 0 && data[0].length > 1) {
      // Set font family for hierarchy column to monospace for better alignment
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
      for (let row = 0; row <= range.e.r; row++) {
        const cellRef = XLSX.utils.encode_cell({ r: row, c: 1 }); // Second column (hierarchy)
        if (worksheet[cellRef]) {
          worksheet[cellRef].s = {
            font: { name: 'Consolas' }
          };
        }
      }
    }
    
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    
    const excelBuffer = XLSX.write(workbook, { 
      bookType: 'xlsx', 
      type: 'array',
      compression: true
    });
    
    return new Blob([excelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  /**
   * Calculate optimal column widths
   */
  private static calculateColumnWidths(data: CellValue[][]): IColumnWidth[] {
    if (data.length === 0) return [];
    
    const columnCount = data[0].length;
    const widths: IColumnWidth[] = [];
    
    for (let col = 0; col < columnCount; col++) {
      let maxWidth = 10; // Minimum width
      
      data.forEach((row, rowIndex) => {
        if (row[col] !== undefined && row[col] !== null) {
          const cellLength = String(row[col]).length;
          
          // Special handling for hierarchy column (usually column 1)
          if (col === 1 && rowIndex > 0) { // Skip header row
            // Hierarchy column needs extra width for indentation
            maxWidth = Math.max(maxWidth, Math.min(cellLength + 5, 80));
          } else {
            maxWidth = Math.max(maxWidth, Math.min(cellLength + 2, 50));
          }
        }
      });
      
      widths.push({ wch: maxWidth });
    }
    
    return widths;
  }

  /**
   * Download file helper method
   */
  private static downloadFile(blob: Blob, fileName: string): void {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.style.display = 'none';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    
    // Clean up the URL object
    setTimeout(() => {
      window.URL.revokeObjectURL(url);
    }, 100);
  }

  /**
   * Format file size
   */
  private static formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  /**
   * Format date for display
   */
  private static formatDate(dateString: string): string {
    try {
      return new Date(dateString).toLocaleDateString();
    } catch {
      return dateString || '';
    }
  }

  /**
   * Create default export settings
   */
  public static createDefaultExportSettings(baseName?: string): IFolderStructureExportSettings {
    return {
      fileName: baseName || 'folder_structure_export',
      includeHeaders: true,
      includeHierarchy: true,
      includeMetadata: true,
      fileFormat: 'xlsx'
    };
  }

  /**
   * Validate export settings
   */
  public static validateExportSettings(
    settings: IFolderStructureExportSettings,
    itemCount?: number
  ): { isValid: boolean; errors: string[]; warnings: string[] } {
    
    const errors: string[] = [];
    const warnings: string[] = [];

    // Check filename
    if (!settings.fileName || settings.fileName.trim().length === 0) {
      errors.push('File name is required');
    } else if (settings.fileName.length > 200) {
      errors.push('File name is too long (maximum 200 characters)');
    }

    // Check item count
    if (itemCount !== undefined) {
      if (itemCount === 0) {
        errors.push('No data to export');
      } else if (itemCount > 50000) {
        warnings.push('Large dataset detected. Export may take some time.');
      }
    }

    // Check max levels
    if (settings.maxLevels !== undefined && settings.maxLevels < 1) {
      errors.push('Maximum levels must be at least 1');
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings
    };
  }

  /**
   * Get preview of export data (first few rows)
   */
  public static getExportPreview(
    folders: ISharePointFolder[],
    settings: IFolderStructureExportSettings,
    previewRows: number = 10
  ): { headers: string[]; sampleRows: CellValue[][]; totalRows: number; hasMoreData: boolean } {
    
    const exportData = this.prepareFolderStructureExportData(folders, settings);
    
    let headers: string[] = [];
    let dataRows: CellValue[][] = [];
    
    if (settings.includeHeaders && exportData.length > 0) {
      headers = exportData[0].map(h => String(h));
      dataRows = exportData.slice(1);
    } else {
      dataRows = exportData;
    }
    
    const sampleRows = dataRows.slice(0, previewRows);
    
    return {
      headers,
      sampleRows,
      totalRows: dataRows.length,
      hasMoreData: dataRows.length > previewRows
    };
  }
}