// src/webparts/xyea/components/RenameFilesManagement/constants/RenameFilesConstants.ts

export const RENAME_FILES_CONSTANTS = {
  // File upload constraints
  MAX_FILE_SIZE: 5 * 1024 * 1024, // 5MB
  SUPPORTED_FILE_TYPES: ['.xlsx', '.xls'],
  
  // Column defaults
  DEFAULT_COLUMN_WIDTH: 150,
  MIN_COLUMN_WIDTH: 80,
  MAX_COLUMN_WIDTH: 800,
  
  // Custom column defaults
  DEFAULT_CUSTOM_COLUMN_NAME: 'New Column',
  CUSTOM_COLUMN_PREFIX: 'custom_',
  EXCEL_COLUMN_PREFIX: 'excel_',
  
  // Table display
  MAX_CELL_LENGTH: 1000,
  DEFAULT_ROWS_PER_PAGE: 50,
  
  // Search settings
  SEARCH_DELAY_MS: 100,
  FOLDER_SEARCH_DELAY_MS: 50,
  MAX_CONCURRENT_SEARCHES: 5,
  
  // Export settings
  DEFAULT_EXPORT_FILENAME: 'renamed_files',
  EXPORT_FORMATS: ['xlsx', 'csv'] as const,
  
  // SharePoint paths
  SHAREPOINT_DOCUMENTS_PATHS: [
    '/Shared Documents',
    '/Documents',
    '/Shared%20Documents'
  ],
  
  // Data types
  SUPPORTED_DATA_TYPES: ['text', 'number', 'date', 'boolean'] as const,
  
  // File search patterns
  RELATIVE_PATH_INDICATORS: ['\\', '/', 'relativepath', 'relative_path'],
  PATH_SEPARATORS: ['\\', '/'],
  
  // UI constants
  UPLOAD_PROGRESS_CLEAR_DELAY: 2000,
  RESIZE_HANDLE_WIDTH: 4,
  DIALOG_ANIMATION_DURATION: 300,
  
  // Validation patterns
  INVALID_FILENAME_CHARS: /[<>:"/\\|?*]/,
  DATE_PATTERNS: [
    /^\d{4}-\d{2}-\d{2}$/,           // YYYY-MM-DD
    /^\d{2}\/\d{2}\/\d{4}$/,         // MM/DD/YYYY
    /^\d{2}-\d{2}-\d{4}$/,           // MM-DD-YYYY
    /^\d{1,2}\/\d{1,2}\/\d{4}$/,     // M/D/YYYY
    /^\d{4}\/\d{2}\/\d{2}$/          // YYYY/MM/DD
  ],
  EMAIL_PATTERN: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
  URL_PATTERN: /^https?:\/\/.+/,
  
  // Boolean values
  BOOLEAN_TRUE_VALUES: ['true', '1', 'yes', 'y', 'on', 'enabled'],
  BOOLEAN_FALSE_VALUES: ['false', '0', 'no', 'n', 'off', 'disabled'],
  
  // Error messages
  ERROR_MESSAGES: {
    FILE_TOO_LARGE: 'File size is too large. Please select a file smaller than 5MB.',
    INVALID_FILE_TYPE: 'Please select a valid Excel file (.xlsx or .xls)',
    EMPTY_FILE: 'Excel file is empty',
    NO_SHEETS: 'No sheets found in the Excel file',
    NO_FOLDER_SELECTED: 'Please select a SharePoint folder first',
    NO_DATA_TO_SEARCH: 'No data rows to search',
    FAILED_TO_LOAD_FOLDERS: 'Failed to load folders from Documents library',
    FAILED_TO_SEARCH_FILES: 'Failed to search files',
    CELL_TOO_LONG: 'Text too long (max 1000 characters)',
    INVALID_NUMBER: 'Must be a valid number',
    INVALID_DATE: 'Must be a valid date',
    INVALID_BOOLEAN: 'Must be true/false, yes/no, or 1/0',
    INVALID_EMAIL: 'Must be a valid email address',
    INVALID_URL: 'Must be a valid URL starting with http:// or https://',
    NETWORK_ERROR: 'Network error occurred. Please check your connection.',
    PERMISSION_DENIED: 'Permission denied. Please check your access rights.',
    TIMEOUT_ERROR: 'Request timed out. Please try again.',
    UNKNOWN_ERROR: 'An unknown error occurred. Please try again.',
    NO_RELATIVE_PATH: 'No RelativePath column found in the data',
    NO_FILENAME_EXTRACTED: 'Could not extract filename from path',
    SHAREPOINT_API_ERROR: 'SharePoint API error occurred'
  },
  
  // Success messages
  SUCCESS_MESSAGES: {
    FILE_LOADED: 'File loaded successfully!',
    SEARCH_COMPLETED: 'File search completed',
    COLUMN_ADDED: 'Column added successfully',
    CHANGES_SAVED: 'Changes saved successfully',
    FOLDER_SELECTED: 'Folder selected successfully',
    FILES_FOUND: 'Files found in SharePoint',
    EXPORT_COMPLETED: 'Export completed successfully',
    SETTINGS_SAVED: 'Settings saved successfully'
  },
  
  // Info messages
  INFO_MESSAGES: {
    PROCESSING_FILE: 'Processing Excel file...',
    SEARCHING_FILES: 'Searching for files...',
    LOADING_FOLDERS: 'Loading SharePoint folders...',
    PREPARING_EXPORT: 'Preparing export...',
    VALIDATING_DATA: 'Validating data...',
    APPLYING_CHANGES: 'Applying changes...'
  },
  
  // SharePoint folder icons
  FOLDER_ICONS: {
    ROOT: '📂',
    REGULAR: '📁',
    SYSTEM: '🔒',
    SHARED: '🗂️',
    TEMPLATES: '📋',
    ARCHIVE: '📦'
  },
  
  // Search result icons
  SEARCH_ICONS: {
    SEARCHING: '🔍',
    FOUND: '✅',
    NOT_FOUND: '❌',
    ERROR: '⚠️',
    PARTIAL: '🔶',
    LOADING: '⏳'
  },
  
  // File type icons
  FILE_TYPE_ICONS: {
    EXCEL: '📊',
    CSV: '📋',
    PDF: '📄',
    WORD: '📝',
    IMAGE: '🖼️',
    UNKNOWN: '📎'
  },
  
  // Local storage keys
  STORAGE_KEYS: {
    COLUMN_WIDTHS: 'renameFiles_columnWidths',
    USER_PREFERENCES: 'renameFiles_userPreferences',
    RECENT_FOLDERS: 'renameFiles_recentFolders',
    RECENT_FILES: 'renameFiles_recentFiles',
    SEARCH_HISTORY: 'renameFiles_searchHistory'
  },
  
  // Animation durations (in ms)
  ANIMATION_DURATIONS: {
    DIALOG_SLIDE: 300,
    FADE_IN: 200,
    FADE_OUT: 150,
    SPINNER: 1000,
    PULSE: 1500,
    BOUNCE: 600,
    SLIDE_UP: 250,
    SLIDE_DOWN: 250
  },
  
  // CSS class names
  CSS_CLASSES: {
    SEARCHING: 'searching',
    FOUND: 'found',
    NOT_FOUND: 'not-found',
    EDITED: 'edited',
    CUSTOM_COLUMN: 'custom-column',
    EXCEL_COLUMN: 'excel-column',
    HIGHLIGHTED: 'highlighted',
    SELECTED: 'selected',
    DISABLED: 'disabled',
    LOADING: 'loading',
    ERROR: 'error',
    SUCCESS: 'success'
  },
  
  // Keyboard shortcuts
  KEYBOARD_SHORTCUTS: {
    SAVE: 'Ctrl+S',
    UNDO: 'Ctrl+Z',
    REDO: 'Ctrl+Y',
    SELECT_ALL: 'Ctrl+A',
    COPY: 'Ctrl+C',
    PASTE: 'Ctrl+V',
    FIND: 'Ctrl+F',
    ESCAPE: 'Escape',
    ENTER: 'Enter',
    DELETE: 'Delete'
  },
  
  // Table limits
  TABLE_LIMITS: {
    MAX_COLUMNS: 50,
    MAX_ROWS: 10000,
    MAX_VISIBLE_ROWS: 1000, // For performance
    MAX_CELL_LENGTH: 1000,
    MIN_COLUMN_WIDTH: 80,
    MAX_COLUMN_WIDTH: 800
  },
  
  // Performance settings
  PERFORMANCE: {
    DEBOUNCE_DELAY: 300,
    THROTTLE_DELAY: 100,
    VIRTUAL_SCROLL_THRESHOLD: 100,
    BATCH_SIZE: 50,
    MAX_CONCURRENT_REQUESTS: 3
  },
  
  // File naming patterns
  FILE_PATTERNS: {
    TEMP_FILE_PREFIX: '~',
    SYSTEM_FILE_PREFIXES: ['_', '.'],
    BACKUP_SUFFIX: '.backup',
    VERSION_PATTERN: /\(\d+\)$/,
    INVALID_CHARS: ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
  },
  
  // SharePoint specific
  SHAREPOINT: {
    MAX_PATH_LENGTH: 400,
    MAX_FILENAME_LENGTH: 255,
    FORBIDDEN_NAMES: ['CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'],
    SYSTEM_FOLDERS: ['Forms', '_catalogs', '_private', '_vti_'],
    API_ENDPOINTS: {
      LISTS: '/_api/web/lists',
      FILES: '/_api/web/getFileByServerRelativeUrl',
      FOLDERS: '/_api/web/getFolderByServerRelativeUrl'
    }
  },
  
  // Default folder structure (fallback for your site)
  DEFAULT_FOLDERS: [
    { 
      name: '(Root - Documents)', 
      isRoot: true, 
      icon: '📂',
      description: 'Root Documents folder'
    },
    { 
      name: 'Debug', 
      isRoot: false, 
      icon: '🐛',
      description: 'Debug files and logs'
    },
    { 
      name: 'LeaveReports', 
      isRoot: false, 
      icon: '📊',
      description: 'Leave management reports'
    },
    { 
      name: 'SRS', 
      isRoot: false, 
      icon: '📋',
      description: 'Software Requirements Specifications'
    },
    { 
      name: 'Templates', 
      isRoot: false, 
      icon: '📄',
      description: 'Document templates'
    }
  ],
  
  // Column type mappings
  COLUMN_TYPES: {
    TEXT: {
      name: 'Text',
      icon: '📝',
      defaultWidth: 150,
      validation: (value: string) => value.length <= 1000
    },
    NUMBER: {
      name: 'Number',
      icon: '🔢',
      defaultWidth: 100,
      validation: (value: string) => !isNaN(parseFloat(value))
    },
    DATE: {
      name: 'Date',
      icon: '📅',
      defaultWidth: 120,
      validation: (value: string) => !isNaN(Date.parse(value))
    },
    BOOLEAN: {
      name: 'Boolean',
      icon: '☑️',
      defaultWidth: 80,
      validation: (value: string) => ['true', 'false', '1', '0'].includes(value.toLowerCase())
    }
  },
  
  // Export configurations
  EXPORT_CONFIG: {
    XLSX: {
      extension: '.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      maxRows: 1048576,
      maxColumns: 16384
    },
    CSV: {
      extension: '.csv',
      mimeType: 'text/csv',
      delimiter: ',',
      encoding: 'UTF-8'
    }
  },
  
  // Theme colors
  COLORS: {
    PRIMARY: '#0078d4',
    SUCCESS: '#107c10',
    WARNING: '#ffb900',
    ERROR: '#d13438',
    INFO: '#0078d4',
    NEUTRAL: '#605e5c',
    BACKGROUND: '#ffffff',
    SURFACE: '#f8f9fa',
    BORDER: '#edebe9'
  },
  
  // Accessibility
  ACCESSIBILITY: {
    MIN_CONTRAST_RATIO: 4.5,
    FOCUS_OUTLINE_WIDTH: 2,
    MIN_TOUCH_TARGET_SIZE: 44, // pixels
    ARIA_LABELS: {
      RESIZE_HANDLE: 'Drag to resize column',
      SEARCH_BUTTON: 'Search for files in SharePoint',
      FOLDER_SELECT: 'Select SharePoint folder',
      FILE_UPLOAD: 'Upload Excel file',
      COLUMN_ADD: 'Add new column',
      CELL_EDIT: 'Edit cell value'
    }
  }
};

// Type definitions for constants
export type ExportFormat = typeof RENAME_FILES_CONSTANTS.EXPORT_FORMATS[number];
export type DataType = typeof RENAME_FILES_CONSTANTS.SUPPORTED_DATA_TYPES[number];
export type ErrorMessageKey = keyof typeof RENAME_FILES_CONSTANTS.ERROR_MESSAGES;
export type SuccessMessageKey = keyof typeof RENAME_FILES_CONSTANTS.SUCCESS_MESSAGES;
export type InfoMessageKey = keyof typeof RENAME_FILES_CONSTANTS.INFO_MESSAGES;
export type SearchResultType = 'searching' | 'found' | 'not-found' | 'error';
export type UploadStage = 'idle' | 'uploading' | 'parsing' | 'complete' | 'error';

// Helper functions for constants
export const RENAME_FILES_HELPERS = {
  // Get error message by key
  getErrorMessage: (key: ErrorMessageKey): string => {
    return RENAME_FILES_CONSTANTS.ERROR_MESSAGES[key];
  },
  
  // Get success message by key
  getSuccessMessage: (key: SuccessMessageKey): string => {
    return RENAME_FILES_CONSTANTS.SUCCESS_MESSAGES[key];
  },
  
  // Get info message by key
  getInfoMessage: (key: InfoMessageKey): string => {
    return RENAME_FILES_CONSTANTS.INFO_MESSAGES[key];
  },
  
  // Check if file type is supported
  isSupportedFileType: (fileName: string): boolean => {
    const lowerName = fileName.toLowerCase();
    return RENAME_FILES_CONSTANTS.SUPPORTED_FILE_TYPES.some(type => 
      lowerName.endsWith(type)
    );
  },
  
  // Check if file size is valid
  isValidFileSize: (size: number): boolean => {
    return size <= RENAME_FILES_CONSTANTS.MAX_FILE_SIZE;
  },
  
  // Get icon for search result
  getSearchIcon: (result: SearchResultType): string => {
    switch (result) {
      case 'searching': return RENAME_FILES_CONSTANTS.SEARCH_ICONS.SEARCHING;
      case 'found': return RENAME_FILES_CONSTANTS.SEARCH_ICONS.FOUND;
      case 'not-found': return RENAME_FILES_CONSTANTS.SEARCH_ICONS.NOT_FOUND;
      case 'error': return RENAME_FILES_CONSTANTS.SEARCH_ICONS.ERROR;
      default: return RENAME_FILES_CONSTANTS.SEARCH_ICONS.LOADING;
    }
  },
  
  // Validate filename
  isValidFileName: (fileName: string): boolean => {
    return !RENAME_FILES_CONSTANTS.INVALID_FILENAME_CHARS.test(fileName) &&
           fileName.trim().length > 0 &&
           fileName.length <= RENAME_FILES_CONSTANTS.SHAREPOINT.MAX_FILENAME_LENGTH;
  },
  
  // Format file size
  formatFileSize: (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  },
  
  // Get file extension
  getFileExtension: (fileName: string): string => {
    return fileName.slice((fileName.lastIndexOf('.') - 1 >>> 0) + 2);
  },
  
  // Get base filename without extension
  getBaseName: (fileName: string): string => {
    return fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
  }
};