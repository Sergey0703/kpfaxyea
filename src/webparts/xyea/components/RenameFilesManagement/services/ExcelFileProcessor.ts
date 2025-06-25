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

    console.log(`[ExcelFileProcessor] 🚀 STARTING DATA PROCESSING`);
    console.log(`[ExcelFileProcessor] Processing ${dataRows.length} data rows...`);
    console.log(`[ExcelFileProcessor] Headers found (${headers.length}):`, headers.map((h, i) => `${i}: "${h}"`));

    // ОТЛАДКА: Показываем первые несколько строк данных
    console.log(`[ExcelFileProcessor] 📋 SAMPLE DATA (first 3 rows):`);
    dataRows.slice(0, 3).forEach((row, index) => {
      console.log(`  Row ${index + 1}:`, row.map((cell, i) => `${i}:"${cell}"`));
    });

    // NEW: Create both filename and directory custom columns
    const filenameColumn: ICustomColumn = {
      id: 'custom_0',
      name: 'Filename',
      isEditable: true,
      defaultValue: '',
      width: 200
    };

    const directoryColumn: ICustomColumn = {
      id: 'custom_1',
      name: 'Directory',
      isEditable: true,
      defaultValue: '',
      width: 250
    };

    // Create column configurations - filename, directory, then Excel columns
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
      },
      {
        id: directoryColumn.id,
        name: directoryColumn.name,
        currentIndex: 1,
        isVisible: true,
        isCustom: true,
        isEditable: true,
        width: 250,
        dataType: 'text'
      }
    ];

    // Add Excel columns
    headers.forEach((header, index) => {
      columns.push({
        id: `excel_${index}`,
        name: String(header || `Column ${index + 1}`),
        originalIndex: index,
        currentIndex: index + 2, // +2 because of filename and directory columns
        isVisible: true,
        isCustom: false,
        isEditable: true,
        dataType: this.detectDataType(dataRows, index)
      });
    });

    // ОТЛАДКА: Показываем созданные колонки
    console.log(`[ExcelFileProcessor] 📊 CREATED COLUMNS (${columns.length}):`);
    columns.forEach(col => {
      console.log(`  ${col.id}: "${col.name}" (index: ${col.currentIndex}, custom: ${col.isCustom})`);
    });

    // Create table rows
    const rows: IRenameTableRow[] = dataRows.map((row, rowIndex) => {
      const cells: { [columnId: string]: ITableCell } = {};

      // ОТЛАДКА: Анализируем каждую строку
      console.log(`[ExcelFileProcessor] 🔍 ANALYZING ROW ${rowIndex + 1}:`);
      console.log(`  All headers (${headers.length}):`, headers.map((h, i) => `${i}:"${h}"`));
      console.log(`  All values (${row.length}):`, row.map((v, i) => `${i}:"${v}"`));

      // Ищем возможные staffID колонки в текущей строке
      const possibleStaffIDData: Array<{index: number, header: string, value: string | number | boolean | undefined}> = []; // FIXED: specific type instead of any
      headers.forEach((header, index) => {
        const headerLower = String(header || '').toLowerCase();
        const cellValue = row[index];
        
        if (headerLower.includes('staff') || 
            headerLower.includes('id') || 
            headerLower === 'id' ||
            /^(staff|id|employee|emp)(_?id)?$/i.test(headerLower)) {
          possibleStaffIDData.push({
            index,
            header: String(header),
            value: cellValue
          });
          console.log(`[ExcelFileProcessor] 🎯 Potential staffID column found:`);
          console.log(`    Header[${index}]: "${header}" -> Value: "${cellValue}"`);
        }
      });

      if (possibleStaffIDData.length === 0) {
        console.log(`[ExcelFileProcessor] ⚠️ No obvious staffID columns found in headers. Checking for numeric/short string values...`);
        
        // Ищем колонки с числовыми значениями или короткими строками (возможные ID)
        headers.forEach((header, index) => {
          const cellValue = row[index];
          const cellStr = String(cellValue || '').trim();
          
          if (cellStr && /^[0-9A-Za-z]{1,10}$/.test(cellStr)) {
            possibleStaffIDData.push({
              index,
              header: String(header),
              value: cellValue
            });
            console.log(`[ExcelFileProcessor] 🔍 Possible ID-like value found:`);
            console.log(`    Header[${index}]: "${header}" -> Value: "${cellValue}" (looks like ID)`);
          }
        });
      }

      // NEW: Extract both filename and directory from the row data
      const pathAnalysis = this.analyzeRowPaths(row, headers, rowIndex);
      
      console.log(`[ExcelFileProcessor] Row ${rowIndex + 1} path analysis:`, pathAnalysis);
      console.log(`[ExcelFileProcessor] Row ${rowIndex + 1} staffID candidates:`, possibleStaffIDData);

      // Populate filename column
      cells[filenameColumn.id] = {
        value: pathAnalysis.fileName,
        isEdited: false,
        originalValue: pathAnalysis.fileName,
        columnId: filenameColumn.id,
        rowIndex
      };

      // NEW: Populate directory column
      cells[directoryColumn.id] = {
        value: pathAnalysis.directoryPath,
        isEdited: false,
        originalValue: pathAnalysis.directoryPath,
        columnId: directoryColumn.id,
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

    // ОТЛАДКА: Показываем примеры обработанных данных
    console.log(`[ExcelFileProcessor] 📋 SAMPLE PROCESSED DATA (first row):`);
    if (rows.length > 0) {
      const firstRow = rows[0];
      Object.entries(firstRow.cells).forEach(([columnId, cell]) => {
        const column = columns.find(c => c.id === columnId);
        console.log(`  ${columnId} (${column?.name}): "${cell.value}"`);
      });
    }

    // ОТЛАДКА: Ищем и показываем все возможные staffID колонки
    console.log(`[ExcelFileProcessor] 🔍 STAFFID COLUMN ANALYSIS:`);
    const possibleStaffIDColumns = columns.filter(col => 
      col.name.toLowerCase().includes('staff') || 
      col.name.toLowerCase().includes('id') ||
      col.id.toLowerCase().includes('staff') ||
      col.id.toLowerCase().includes('id') ||
      /^(staff|id|employee|emp)(_?id)?$/i.test(col.name.toLowerCase())
    );

    if (possibleStaffIDColumns.length > 0) {
      console.log(`Found ${possibleStaffIDColumns.length} possible staffID columns:`);
      possibleStaffIDColumns.forEach(col => {
        console.log(`  - ${col.id} (${col.name})`);
        
        // Показываем примеры значений из первых 3 строк
        const sampleValues = rows.slice(0, 3).map(row => row.cells[col.id]?.value);
        console.log(`    Sample values: [${sampleValues.map(v => `"${v}"`).join(', ')}]`);
      });
    } else {
      console.warn(`[ExcelFileProcessor] ⚠️ NO staffID columns found!`);
      console.log(`Available columns:`, columns.map(c => `${c.id}(${c.name})`).join(', '));
      
      // Попробуем найти колонки с числовыми значениями
      console.log(`[ExcelFileProcessor] 🔍 Looking for numeric/ID-like columns...`);
      columns.forEach(col => {
        if (col.isCustom) return; // Skip custom columns
        
        const sampleValues = rows.slice(0, 5).map(row => row.cells[col.id]?.value);
        const numericCount = sampleValues.filter(v => {
          const str = String(v || '').trim();
          return str && /^[0-9A-Za-z]{1,10}$/.test(str);
        }).length;
        
        if (numericCount >= 3) { // At least 3 out of 5 samples look like IDs
          console.log(`  Potential ID column: ${col.id} (${col.name})`);
          console.log(`    Sample values: [${sampleValues.map(v => `"${v}"`).join(', ')}]`);
        }
      });
    }

    // Create sheet structure
    const sheet: IExcelSheet = {
      name: sheetName,
      headers: headers.map(h => String(h || '')),
      data: [], // We're using our own row structure
      totalRows: dataRows.length,
      isValid: true
    };

    excelFile.sheets = [sheet];

    const finalData = {
      originalFile: excelFile,
      currentSheet: sheet,
      columns,
      rows,
      customColumns: [filenameColumn, directoryColumn], // Both custom columns
      totalRows: dataRows.length,
      editedCellsCount: 0
    };

    // ФИНАЛЬНАЯ ОТЛАДКА
    console.log(`[ExcelFileProcessor] 🎯 FINAL PROCESSING SUMMARY:`);
    console.log(`  📊 Total rows: ${finalData.totalRows}`);
    console.log(`  📋 Total columns: ${finalData.columns.length}`);
    console.log(`  🔧 Custom columns: ${finalData.customColumns.length}`);
    console.log(`  📁 Files with paths: ${rows.filter(r => r.cells.custom_0?.value).length}`); // FIXED: dot notation
    console.log(`  📂 Files with directories: ${rows.filter(r => r.cells.custom_1?.value).length}`); // FIXED: dot notation

    return finalData;
  }

  /**
   * NEW: Comprehensive path analysis for each row
   */
  private analyzeRowPaths(
    row: (string | number | boolean | undefined)[], 
    headers: (string | number | boolean | undefined)[],
    rowIndex: number
  ): { fileName: string; directoryPath: string; relativePath: string; pathFound: boolean } {
    
    console.log(`[ExcelFileProcessor] 🔍 ANALYZING PATHS for row ${rowIndex + 1}`);
    
    // Step 1: Find RelativePath column or path-like content
    const relativePath = this.extractRelativePathFromRowData(row, headers, rowIndex);
    
    if (!relativePath) {
      console.log(`[ExcelFileProcessor] ❌ No path found in row ${rowIndex + 1}`);
      return {
        fileName: '',
        directoryPath: '',
        relativePath: '',
        pathFound: false
      };
    }
    
    // Step 2: Split the path into directory and filename
    const pathComponents = this.splitPathAndFilename(relativePath);
    
    console.log(`[ExcelFileProcessor] ✅ Row ${rowIndex + 1} path analysis complete:`);
    console.log(`  📄 Original path: "${relativePath}"`);
    console.log(`  📁 Directory: "${pathComponents.directoryPath}"`);
    console.log(`  📄 Filename: "${pathComponents.fileName}"`);
    
    return {
      fileName: pathComponents.fileName,
      directoryPath: pathComponents.directoryPath,
      relativePath,
      pathFound: true
    };
  }

  private extractRelativePathFromRowData(
    row: (string | number | boolean | undefined)[], 
    headers: (string | number | boolean | undefined)[],
    rowIndex: number
  ): string {
    console.log(`[ExcelFileProcessor] 🔍 Searching for RelativePath in row ${rowIndex + 1}`);
    
    // FIRST: Look specifically for RelativePath column by header name
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i] || '').toLowerCase();
      const cellValue = String(row[i] || '');
      
      if (header.includes('relativepath') || 
          header.includes('relative_path') ||
          header.includes('relative path')) {
        console.log(`[ExcelFileProcessor] ✅ Found RelativePath by header "${header}" in column ${i}: "${cellValue}"`);
        return cellValue;
      }
    }
    
    // SECOND: Look for other path-related headers
    for (let i = 0; i < headers.length; i++) {
      const header = String(headers[i] || '').toLowerCase();
      const cellValue = String(row[i] || '');
      
      if (header.includes('path') || 
          header.includes('filepath') || 
          header.includes('file_path') ||
          header.includes('file path')) {
        console.log(`[ExcelFileProcessor] ✅ Found path by header "${header}" in column ${i}: "${cellValue}"`);
        return cellValue;
      }
    }
    
    // THIRD: Look for content that looks like a file path (more strict validation)
    for (let i = 0; i < row.length; i++) {
      const cellValue = String(row[i] || '');
      if (cellValue && this.looksLikeValidFilePath(cellValue)) {
        console.log(`[ExcelFileProcessor] ✅ Found path by content pattern in column ${i}: "${cellValue}"`);
        return cellValue;
      }
    }
    
    console.log(`[ExcelFileProcessor] ❌ No valid path found in row ${rowIndex + 1}`);
    return '';
  }

  /**
   * NEW: More strict validation for file paths
   */
  private looksLikeValidFilePath(value: string): boolean {
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
      /^\d{2}\/\d{2}\/\d{4}$/, // FIXED: removed unnecessary escape
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
    const pathParts = relativePath.split(/[\\//]/);
    const fileName = pathParts[pathParts.length - 1] || '';
    
    console.log(`${logPrefix}: Path split into ${pathParts.length} parts:`, pathParts);
    console.log(`${logPrefix}: Extracted filename: "${fileName}"`);
    
    return fileName;
  }

  public extractDirectoryPath(relativePath: string): string {
    if (!relativePath) return '';
    
    console.log(`[ExcelFileProcessor] Extracting directory from path: "${relativePath}"`);
    
    // Split the path by both backslashes and forward slashes
    const pathParts = relativePath.split(/[\\//]/);
    
    // Remove the filename (last part) to get directory path
    const directoryParts = pathParts.slice(0, -1);
    
    // Join back with forward slashes (SharePoint format)
    const directoryPath = directoryParts.join('/');
    
    console.log(`[ExcelFileProcessor] Path parts:`, pathParts);
    console.log(`[ExcelFileProcessor] Directory parts:`, directoryParts);
    console.log(`[ExcelFileProcessor] Extracted directory: "${directoryPath}"`);
    
    return directoryPath;
  }

  /**
   * NEW: Enhanced method to get directory path from row using the Directory column
   */
  public extractDirectoryPathFromRow(row: IRenameTableRow): string {
    console.log(`[ExcelFileProcessor] Extracting directory path from row ${row.rowIndex}`);
    
    // FIRST: Try to get directory from the Directory column (custom_1)
    const directoryCell = row.cells.custom_1; // FIXED: dot notation
    if (directoryCell && directoryCell.value) {
      const directoryPath = String(directoryCell.value).trim();
      if (directoryPath) {
        console.log(`[ExcelFileProcessor] Found directory in Directory column: "${directoryPath}"`);
        return directoryPath;
      }
    }
    
    // FALLBACK: Use the old method to extract from RelativePath
    console.log(`[ExcelFileProcessor] Directory column empty, trying to extract from RelativePath...`);
    
    // Look for RelativePath data in Excel columns
    const relativePathCell = Object.values(row.cells).find(cell => {
      const columnIdLower = cell.columnId.toLowerCase();
      const cellValue = String(cell.value || '');
      
      // Method 1: Check if column ID suggests it's a path column
      if (columnIdLower.includes('relativepath') || 
          columnIdLower.includes('relative_path') || 
          columnIdLower.includes('path')) {
        return true;
      }
      
      // Method 2: Check if this cell contains a path-like value
      if (cellValue && this.looksLikeValidFilePath(cellValue)) {
        console.log(`[ExcelFileProcessor] Found path-like content in column ${cell.columnId}: "${cellValue}"`);
        return true;
      }
      
      return false;
    });
    
    if (!relativePathCell || !relativePathCell.value) {
      console.log(`[ExcelFileProcessor] No RelativePath found in row ${row.rowIndex}`);
      return '';
    }
    
    const relativePath = String(relativePathCell.value);
    const directoryPath = this.extractDirectoryPath(relativePath);
    
    console.log(`[ExcelFileProcessor] Row ${row.rowIndex}: RelativePath="${relativePath}" -> Directory="${directoryPath}"`);
    
    return directoryPath;
  }

  public splitPathAndFilename(fullPath: string): { directoryPath: string; fileName: string } {
    if (!fullPath) {
      return { directoryPath: '', fileName: '' };
    }
    
    const pathParts = fullPath.split(/[\\//]/);
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