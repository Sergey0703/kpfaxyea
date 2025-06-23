// src/webparts/xyea/components/RenameFilesManagement/handlers/ColumnResizeHandler.ts

export class ColumnResizeHandler {
  private resizingColumn: string | null = null;
  private startX: number = 0;
  private startWidth: number = 0;
  private onColumnResize: (columnId: string, newWidth: number) => void;

  constructor(onColumnResize: (columnId: string, newWidth: number) => void) {
    this.onColumnResize = onColumnResize;
    
    // Bind methods to preserve 'this' context
    this.handleMouseMove = this.handleMouseMove.bind(this);
    this.handleMouseUp = this.handleMouseUp.bind(this);
  }

  public addEventListeners(): void {
    document.addEventListener('mousemove', this.handleMouseMove);
    document.addEventListener('mouseup', this.handleMouseUp);
  }

  public removeEventListeners(): void {
    document.removeEventListener('mousemove', this.handleMouseMove);
    document.removeEventListener('mouseup', this.handleMouseUp);
  }

  public handleResizeStart = (columnId: string, event: React.MouseEvent): void => {
    event.preventDefault();
    event.stopPropagation();
    
    this.resizingColumn = columnId;
    this.startX = event.clientX;
    
    // Get current width from the column header element
    const headerElement = event.currentTarget.closest('th') as HTMLElement;
    if (headerElement) {
      this.startWidth = headerElement.offsetWidth;
    } else {
      this.startWidth = 150; // fallback width
    }
    
    // Set cursor styles
    document.body.style.cursor = 'col-resize';
    document.body.style.userSelect = 'none';
    
    console.log(`Started resizing column ${columnId}, start width: ${this.startWidth}`);
  }

  private handleMouseMove(event: MouseEvent): void {
    if (!this.resizingColumn) return;
    
    const deltaX = event.clientX - this.startX;
    const newWidth = Math.max(80, this.startWidth + deltaX); // Minimum width of 80px
    
    // Apply the width change immediately for visual feedback
    this.onColumnResize(this.resizingColumn, newWidth);
  }

  private handleMouseUp(): void {
    if (!this.resizingColumn) return;
    
    console.log(`Finished resizing column ${this.resizingColumn}`);
    
    this.resizingColumn = null;
    document.body.style.cursor = '';
    document.body.style.userSelect = '';
  }

  public isResizing(): boolean {
    return this.resizingColumn !== null;
  }

  public getCurrentResizingColumn(): string | null {
    return this.resizingColumn;
  }

  // Utility method to set column width programmatically
  public setColumnWidth(columnId: string, width: number): void {
    const clampedWidth = Math.max(80, Math.min(800, width)); // Between 80px and 800px
    this.onColumnResize(columnId, clampedWidth);
  }

  // Method to auto-resize column to fit content
  public autoResizeColumn(columnId: string, tableElement: HTMLTableElement): void {
    try {
      // Find all cells in this column
      const columnIndex = this.getColumnIndex(columnId, tableElement);
      if (columnIndex === -1) return;

      let maxWidth = 80; // minimum width
      
      // Check header width
      const headerCell = tableElement.querySelector(`thead th:nth-child(${columnIndex + 1})`) as HTMLElement;
      if (headerCell) {
        const headerWidth = this.getTextWidth(headerCell.textContent || '', getComputedStyle(headerCell));
        maxWidth = Math.max(maxWidth, headerWidth + 24); // Add padding
      }

      // Check content cells width (sample first 10 rows for performance)
      const contentCells = tableElement.querySelectorAll(`tbody tr:nth-child(-n+10) td:nth-child(${columnIndex + 1})`);
      contentCells.forEach(cell => {
        const cellElement = cell as HTMLElement;
        const cellWidth = this.getTextWidth(cellElement.textContent || '', getComputedStyle(cellElement));
        maxWidth = Math.max(maxWidth, cellWidth + 16); // Add padding
      });

      // Cap at reasonable maximum
      maxWidth = Math.min(maxWidth, 400);
      
      this.setColumnWidth(columnId, maxWidth);
    } catch (error) {
      console.error('Error auto-resizing column:', error);
    }
  }

  private getColumnIndex(columnId: string, tableElement: HTMLTableElement): number {
    const headers = tableElement.querySelectorAll('thead th');
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i] as HTMLElement;
      if (header.dataset.columnId === columnId) {
        return i;
      }
    }
    return -1;
  }

  private getTextWidth(text: string, style: CSSStyleDeclaration): number {
    const canvas = document.createElement('canvas');
    const context = canvas.getContext('2d');
    
    if (!context) return text.length * 8; // fallback
    
    context.font = `${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    return context.measureText(text).width;
  }

  // Save/restore column widths to localStorage
  public saveColumnWidths(tableId: string, columns: Array<{id: string, width?: number}>): void {
    try {
      const widthsData = columns.reduce((acc, col) => {
        if (col.width) {
          acc[col.id] = col.width;
        }
        return acc;
      }, {} as {[key: string]: number});

      localStorage.setItem(`columnWidths_${tableId}`, JSON.stringify(widthsData));
    } catch (error) {
      console.warn('Could not save column widths to localStorage:', error);
    }
  }

  public loadColumnWidths(tableId: string): {[key: string]: number} {
    try {
      const saved = localStorage.getItem(`columnWidths_${tableId}`);
      return saved ? JSON.parse(saved) : {};
    } catch (error) {
      console.warn('Could not load column widths from localStorage:', error);
      return {};
    }
  }
}