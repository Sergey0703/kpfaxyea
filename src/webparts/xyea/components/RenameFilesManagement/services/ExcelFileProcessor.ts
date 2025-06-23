// src/webparts/xyea/components/RenameFilesManagement/services/ExcelFileProcessor.ts

import * as XLSX from 'xlsx';
import { 
  IRenameFilesData, 
  IColumnConfiguration, 
  IRenameTableRow, 
  ITableCell, 
  ICustomColumn 
} from '../types/RenameFilesTypes';
import { IExcelFile, IExcelSheet } from '../../../interfaces/ExcelInterfaces';

export class ExcelFileProcessor {
  
  public async processFile(
    file: File, 
    progressCallback: (stage: string, progress: number, message: string) => void
  ): Promise<IRenameFilesData> {
    try {
      // Validate file format
      const fileName = file.name.toLowerCase();
      if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
        throw new Error('Please select a valid Excel file (.xlsx or .xls)');
      }

      // Validate file size (max 5MB)
      if (file.size > 5 * 1024 * 1024) {
        throw new Error('File size is too large. Please select a file smaller than 5MB.');
      }

      progressCallback('uploading', 25, 'Reading file...');

      // Read file
      const arrayBuffer = await this.readFileAsArrayBuffer(file);
      
      progressCallback('parsing', 50, 'Parsing Excel data...');

      // Parse Excel
      const workbook = XLSX.read(arrayBuffer, { 
        type: 'array',
        cellDates: true,
        dateNF: 'yyyy-mm-dd'
      });

      // Get first sheet
      const firstSheetName = workbook.SheetNames[0];
      if (!firstSheetName) {
        throw new Error('No sheets found in the Excel file');
      }

      const worksheet = workbook.Sheets[firstSheetName];
      
      // Convert to array format
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        defval: '',
        raw: false
      }) as (string | number | boolean | undefined)[][];

      if (jsonData.length === 0) {
        throw new Error('Excel file is empty');
      }

      progressCallback('parsing', 75, 'Processing data...');

      // Create Excel file structure
      const excelFile: IExcelFile = {
        name: file.name,
        size: file.size,
        lastModified: new Date(file.lastModified),
        data: arrayBuffer,
        sheets: []
      };

      // Process data
      const processedData = this.processExcelData(jsonData, excelFile, firstSheetName);

      progressCallback('complete', 100, 'Processing complete!');

      return processedData;

    } catch (error) {
      console.error('Error processing Excel file:', error);
      throw error;
    }
  }

  private readFileAsArrayBuffer(file: File): Promise<ArrayBuffer> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (event) => {
        if (event.target?.result instanceof ArrayBuffer) {
          resolve(event.target.result);
        } else {
          reject(new Error('Failed to read file as ArrayBuffer'));
        }
      };
      
      reader.onerror = () => reject(new Error('FileReader error'));
      reader.readAsArrayBuffer(file);
    });
  }

  private processExcelData(
    jsonData: (string | number | boolean | undefined)[][], 
    excelFile: IExcelFile, 
    sheetName: string
  ): IRenameFilesData {
    const headers = jsonData[0] || [];
    const dataRows = jsonData.slice(1);

    console.log(`[ExcelFileProcessor] Starting to process ${dataRows.length} data rows...`);
    console.log(`[ExcelFileProcessor] Headers found:`, headers);

    // Create filename column as the first custom column
    const filenameColumn: ICustomColumn = {
      id: 'custom_0',
      name: 'Filename',
      isEditable: true,
      defaultValue: '',
      width: 200
    };

    // Create column configurations - first the filename column, then Excel columns
    const columns: IColumnConfiguration[] = [
      {
        id: filenameColumn.id,
        name: filenameColumn.name,
        currentIndex: 0,
        isVisible: true,
        isCustom: true,
        isEditable: true,
        width: 200,
        dataType: 'text'
      }
    ];

    // Add Excel columns
    headers.forEach((header, index) => {
      columns.push({
        id: `excel_${index}`,
        name: String(header || `Column ${index + 1}`),
        originalIndex: index,
        currentIndex: index + 1, // +1 because of filename column at position 0
        isVisible: true,
        isCustom: false,
        isEditable: true,
        dataType: this.detectDataType(dataRows, index)
      });
    });

    // Create table rows
    const rows: IRenameTableRow[] = dataRows.map((row, rowIndex) => {
      const cells: { [columnId: string]: ITableCell } = {};

      // Extract filename from the row data and populate the filename column
      const relativePath = this.extractRelativePathFromRowData(row, headers, rowIndex);
      const fileName = relativePath ? this.extractFileName(relativePath, rowIndex) || '' : '';
      
      console.log(`[ExcelFileProcessor] Row ${rowIndex + 1}: RelativePath="${relativePath}" -> Filename="${fileName}"`);

      cells[filenameColumn.id] = {
        value: fileName,
        isEdited: false,
        originalValue: fileName,
        columnId: filenameColumn.id,
        rowIndex
      };

      // Add Excel data cells
      headers.forEach((_, colIndex) => {
        const columnId = `excel_${colIndex}`;
        const value = row[colIndex];
        
        cells[columnId] = {
          value: value,
          isEdited: false,
          originalValue: value,
          columnId,
          rowIndex
        };
      });

      return {
        rowIndex,
        cells,
        isVisible: true,
        isEdited: false
      };
    });

    // Create sheet structure
    const sheet: IExcelSheet = {
      name: sheetName,
      headers: headers.map(h => String(h || '')),
      data: [], // We're using our own row structure
      totalRows: dataRows.length,
      isValid: true
    };

    excelFile.sheets = [sheet];

    return {
      originalFile: excelFile,
      currentSheet: sheet,
      columns,
      rows,
      customColumns: [filenameColumn],
      totalRows: dataRows.length,
      editedCellsCount: 0
    };
  }

  private extractRelativePathFromRowData(
    row: (string | number | boolean | undefined)[], 
    headers: (string | number | boolean | undefined)[],
    rowIndex: number
  ): string {
    console.log(`[ExcelFileProcessor] Processing row ${rowIndex + 1} with ${row.length} cells:`);
    
    // Log all headers for debugging
    headers.forEach((header, index) => {
      const cellValue = String(row[index] || '');
      console.log(`  Header[${index}]: "${header}" -> Value: "${cellValue}"`);
    });
    
    // FIRST: Look specifically for RelativePath column by header name
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i] || '').toLowerCase();
      const cellValue = String(row[i] || '');
      
      if (header.includes('relativepath') || 
          header.includes('relative_path')) {
        console.log(`[ExcelFileProcessor] Found RelativePath by header "${header}" in column ${i}: "${cellValue}"`);
        return cellValue;
      }
    }
    
    // SECOND: Look for other path-related headers
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i] || '').toLowerCase();
      const cellValue = String(row[i] || '');
      
      if (header.includes('path') || header.includes('filepath') || header.includes('file_path')) {
        console.log(`[ExcelFileProcessor] Found path by header "${header}" in column ${i}: "${cellValue}"`);
        return cellValue;
      }
    }
    
    // THIRD: Only if no path headers found, look for content that looks like a file path
    // But be more strict - must have file extension and proper path structure
    for (let i = 0; i < row.length; i++) {
      const cellValue = String(row[i] || '');
      if (cellValue && 
          (cellValue.includes('\\') || cellValue.includes('/')) && 
          cellValue.includes('.') && // must have file extension
          cellValue.length > 10 && // reasonable minimum length for a file path
          this.looksLikeFilePath(cellValue)) { // additional validation
        console.log(`[ExcelFileProcessor] Found path by content pattern in column ${i}: "${cellValue}"`);
        return cellValue;
      }
    }
    
    console.log(`[ExcelFileProcessor] No path found in row ${rowIndex + 1}`);
    return '';
  }

  private looksLikeFilePath(value: string): boolean {
    // Additional validation to ensure this really looks like a file path
    // and not just text that happens to contain slashes
    
    // Must contain at least one directory separator
    if (!value.includes('\\') && !value.includes('/')) {
      return false;
    }
    
    // Must have a file extension at the end
    const parts = value.split(/[\\\/]/);
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
    
    // Reject if it looks like a person's name (contains parentheses with common abbreviations)
    if (value.includes('(') && value.includes(')') && 
        (value.includes('INST') || value.includes('DRV') || value.includes('MGR'))) {
      return false;
    }
    
    return true;
  }

  private detectDataType(dataRows: (string | number | boolean | undefined)[][], columnIndex: number): 'text' | 'number' | 'date' | 'boolean' {
    const sample = dataRows.slice(0, 10).map(row => row[columnIndex]);
    
    let numberCount = 0;
    let dateCount = 0;
    let booleanCount = 0;
    
    for (const value of sample) {
      if (value === undefined || value === null || value === '') continue;
      
      if (typeof value === 'number') {
        numberCount++;
      } else if (typeof value === 'boolean') {
        booleanCount++;
      } else if (typeof value === 'string') {
        // Check if it looks like a date
        if (this.isDateString(value)) {
          dateCount++;
        }
      }
    }
    
    if (booleanCount > sample.length * 0.5) return 'boolean';
    if (numberCount > sample.length * 0.5) return 'number';
    if (dateCount > sample.length * 0.5) return 'date';
    
    return 'text';
  }

  private isDateString(value: string): boolean {
    const datePatterns = [
      /^\d{4}-\d{2}-\d{2}$/,
      /^\d{2}\/\d{2}\/\d{4}$/,
      /^\d{2}-\d{2}-\d{4}$/
    ];
    
    return datePatterns.some(pattern => pattern.test(value));
  }

  public extractFileName(relativePath: string, rowIndex?: number): string {
    const logPrefix = rowIndex !== undefined ? `[ExcelFileProcessor] Row ${rowIndex + 1}` : '[ExcelFileProcessor]';
    
    if (!relativePath) {
      console.log(`${logPrefix}: No relative path provided`);
      return '';
    }
    
    console.log(`${logPrefix}: Extracting filename from path: "${relativePath}"`);
    
    // Extract filename from path (e.g., "634\2022\10\634-AS Food Hygiene.pdf" -> "634-AS Food Hygiene.pdf")
    const pathParts = relativePath.split(/[\\\/]/);
    const fileName = pathParts[pathParts.length - 1] || '';
    
    console.log(`${logPrefix}: Path split into ${pathParts.length} parts:`, pathParts);
    console.log(`${logPrefix}: Extracted filename: "${fileName}"`);
    
    return fileName;
  }

  public extractDirectoryPath(relativePath: string): string {
    if (!relativePath) return '';
    
    console.log(`[ExcelFileProcessor] Extracting directory from path: "${relativePath}"`);
    
    // Split the path by both backslashes and forward slashes
    const pathParts = relativePath.split(/[\\\/]/);
    
    // Remove the filename (last part) to get directory path
    const directoryParts = pathParts.slice(0, -1);
    
    // Join back with forward slashes (SharePoint format)
    const directoryPath = directoryParts.join('/');
    
    console.log(`[ExcelFileProcessor] Path parts:`, pathParts);
    console.log(`[ExcelFileProcessor] Directory parts:`, directoryParts);
    console.log(`[ExcelFileProcessor] Extracted directory: "${directoryPath}"`);
    
    return directoryPath;
  }

  public extractDirectoryPathFromRow(row: IRenameTableRow): string {
    console.log(`[ExcelFileProcessor] Extracting directory path from row ${row.rowIndex}`);
    
    // Log all cells in the row for debugging
    Object.entries(row.cells).forEach(([columnId, cell]) => {
      console.log(`  Column ${columnId}: "${cell.value}"`);
    });
    
    // LOOK FOR RelativePath DATA - check both column ID and cell content
    const relativePathCell = Object.values(row.cells).find(cell => {
      const columnIdLower = cell.columnId.toLowerCase();
      const cellValue = String(cell.value || '');
      
      // Method 1: Check if column ID contains "relativepath"
      if (columnIdLower.includes('relativepath') || columnIdLower.includes('relative_path')) {
        return true;
      }
      
      // Method 2: Check if this cell contains a path-like value (backup method)
      if (cellValue && 
          (cellValue.includes('\\') || cellValue.includes('/')) && 
          cellValue.includes('.') && // has file extension
          cellValue.length > 10 && // reasonable length
          this.looksLikeFilePath(cellValue)) {
        console.log(`[ExcelFileProcessor] Found path-like content in column ${cell.columnId}: "${cellValue}"`);
        return true;
      }
      
      return false;
    });
    
    if (!relativePathCell || !relativePathCell.value) {
      console.log(`[ExcelFileProcessor] No RelativePath found in row ${row.rowIndex}`);
      console.log(`[ExcelFileProcessor] Available columns:`, Object.keys(row.cells));
      return '';
    }
    
    const relativePath = String(relativePathCell.value);
    console.log(`[ExcelFileProcessor] Found RelativePath in row ${row.rowIndex}: "${relativePath}"`);
    
    // Only proceed if this actually looks like a file path
    if (!relativePath.includes('\\') && !relativePath.includes('/')) {
      console.log(`[ExcelFileProcessor] RelativePath doesn't contain path separators: "${relativePath}"`);
      return '';
    }
    
    const directoryPath = this.extractDirectoryPath(relativePath);
    
    console.log(`[ExcelFileProcessor] Row ${row.rowIndex}: RelativePath="${relativePath}" -> Directory="${directoryPath}"`);
    
    return directoryPath;
  }

  public validateDirectoryStructure(data: IRenameFilesData): Array<{
    rowIndex: number;
    fileName: string;
    relativePath: string;
    directoryPath: string;
    hasValidPath: boolean;
  }> {
    console.log(`[ExcelFileProcessor] Validating directory structure for ${data.rows.length} rows`);
    
    const validationResults: Array<{
      rowIndex: number;
      fileName: string;
      relativePath: string;
      directoryPath: string;
      hasValidPath: boolean;
    }> = [];

    data.rows.forEach(row => {
      // Get filename from first column
      const fileName = String(row.cells['custom_0']?.value || '');
      
      // Extract RelativePath and directory
      const relativePath = this.extractRelativePath(row);
      const directoryPath = this.extractDirectoryPathFromRow(row);
      
      // Check if path structure is valid
      const hasValidPath = relativePath !== '' && directoryPath !== '' && fileName !== '';
      
      const result = {
        rowIndex: row.rowIndex,
        fileName,
        relativePath,
        directoryPath,
        hasValidPath
      };
      
      validationResults.push(result);
      
      console.log(`[ExcelFileProcessor] Row ${row.rowIndex}: FileName="${fileName}", Directory="${directoryPath}", Valid=${hasValidPath}`);
    });

    const validRows = validationResults.filter(r => r.hasValidPath).length;
    console.log(`[ExcelFileProcessor] Validation complete: ${validRows}/${validationResults.length} rows have valid paths`);
    
    return validationResults;
  }

  public getDirectoryStatistics(data: IRenameFilesData): Array<{
    directoryPath: string;
    fileCount: number;
    rowIndexes: number[];
  }> {
    console.log(`[ExcelFileProcessor] Calculating directory statistics`);
    
    const directoryMap = new Map<string, { count: number; rowIndexes: number[] }>();
    
    data.rows.forEach(row => {
      const directoryPath = this.extractDirectoryPathFromRow(row);
      
      if (directoryPath) {
        const existing = directoryMap.get(directoryPath);
        if (existing) {
          existing.count++;
          existing.rowIndexes.push(row.rowIndex);
        } else {
          directoryMap.set(directoryPath, {
            count: 1,
            rowIndexes: [row.rowIndex]
          });
        }
      }
    });
    
    const statistics = Array.from(directoryMap.entries()).map(([directoryPath, data]) => ({
      directoryPath,
      fileCount: data.count,
      rowIndexes: data.rowIndexes
    }));
    
    // Sort by file count descending
    statistics.sort((a, b) => b.fileCount - a.fileCount);
    
    console.log(`[ExcelFileProcessor] Directory statistics:`, statistics);
    
    return statistics;
  }

  public splitPathAndFilename(fullPath: string): { directoryPath: string; fileName: string } {
    if (!fullPath) {
      return { directoryPath: '', fileName: '' };
    }
    
    const pathParts = fullPath.split(/[\\\/]/);
    const fileName = pathParts[pathParts.length - 1] || '';
    const directoryParts = pathParts.slice(0, -1);
    const directoryPath = directoryParts.join('/');
    
    console.log(`[ExcelFileProcessor] Split "${fullPath}" -> Directory: "${directoryPath}", File: "${fileName}"`);
    
    return { directoryPath, fileName };
  }

  public addCustomColumn(data: IRenameFilesData): IRenameFilesData {
    const newColumnId = `custom_${data.customColumns.length}`;
    
    const newCustomColumn: ICustomColumn = {
      id: newColumnId,
      name: `New Column ${data.customColumns.length + 1}`,
      isEditable: true,
      defaultValue: '',
      width: 150
    };

    const newColumnConfig: IColumnConfiguration = {
      id: newColumnId,
      name: newCustomColumn.name,
      currentIndex: data.columns.length,
      isVisible: true,
      isCustom: true,
      isEditable: true,
      width: 150,
      dataType: 'text'
    };

    // Add empty cells for the new column in all rows
    const updatedRows = data.rows.map(row => ({
      ...row,
      cells: {
        ...row.cells,
        [newColumnId]: {
          value: '',
          isEdited: false,
          columnId: newColumnId,
          rowIndex: row.rowIndex
        }
      }
    }));

    return {
      ...data,
      columns: [...data.columns, newColumnConfig],
      customColumns: [...data.customColumns, newCustomColumn],
      rows: updatedRows
    };
  }

  public updateColumnWidth(data: IRenameFilesData, columnId: string, newWidth: number): IRenameFilesData {
    const updatedColumns = data.columns.map(column => 
      column.id === columnId 
        ? { ...column, width: newWidth }
        : column
    );

    // Also update custom columns if applicable
    const updatedCustomColumns = data.customColumns.map(column =>
      column.id === columnId
        ? { ...column, width: newWidth }
        : column
    );

    return {
      ...data,
      columns: updatedColumns,
      customColumns: updatedCustomColumns
    };
  }

  public extractRelativePath(row: IRenameTableRow): string {
    // Find cell that contains RelativePath data
    const relativePathCell = Object.values(row.cells).find(cell => {
      const cellValue = String(cell.value || '');
      return (
        cell.columnId.toLowerCase().includes('relativepath') ||
        cell.columnId.toLowerCase().includes('relative_path') ||
        cell.columnId.toLowerCase().includes('path') ||
        cellValue.includes('\\') || 
        cellValue.includes('/')
      );
    });
    
    if (!relativePathCell || !relativePathCell.value) {
      return '';
    }
    
    return String(relativePathCell.value);
  }
}