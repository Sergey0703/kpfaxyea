// src/webparts/xyea/components/SeparateFilesManagement/FileUploader.tsx

import * as React from 'react';
import styles from './FileUploader.module.scss';
import { IUploadProgress, UploadStage } from '../../interfaces/ExcelInterfaces';

export interface IFileUploaderProps {
  onFileSelect: (file: File) => void;
  loading: boolean;
  progress?: IUploadProgress;
  disabled?: boolean;
  acceptedFormats?: string[];
  maxFileSize?: number; // в MB
}

export interface IFileUploaderState {
  isDragOver: boolean;
  error: string | undefined; // Changed from null to undefined
}

export default class FileUploader extends React.Component<IFileUploaderProps, IFileUploaderState> {
  private fileInputRef: React.RefObject<HTMLInputElement>;

  constructor(props: IFileUploaderProps) {
    super(props);
    
    this.state = {
      isDragOver: false,
      error: undefined // Changed from null to undefined
    };

    this.fileInputRef = React.createRef<HTMLInputElement>();
  }

  private handleDragEnter = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    if (!this.props.disabled) {
      this.setState({ isDragOver: true });
    }
  }

  private handleDragLeave = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    this.setState({ isDragOver: false });
  }

  private handleDragOver = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
  }

  private handleDrop = (e: React.DragEvent): void => {
    e.preventDefault();
    e.stopPropagation();
    this.setState({ isDragOver: false });

    if (this.props.disabled) {
      return;
    }

    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
      this.processFile(files[0]);
    }
  }

  private handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const files = e.target.files;
    if (files && files.length > 0) {
      this.processFile(files[0]);
    }
  }

  private processFile = (file: File): void => {
    this.setState({ error: undefined });

    // Валидация файла
    const validation = this.validateFile(file);
    if (!validation.isValid) {
      this.setState({ error: validation.error || undefined });
      return;
    }

    // Передаем файл родительскому компоненту
    this.props.onFileSelect(file);
  }

  private validateFile = (file: File): { isValid: boolean; error?: string } => {
    const { acceptedFormats = ['.xlsx', '.xls', '.csv'], maxFileSize = 10 } = this.props;

    // Проверка формата
    const nameParts = file.name.split('.');
    const fileExtension = nameParts.length > 1 ? '.' + nameParts[nameParts.length - 1].toLowerCase() : '';
    
    if (!acceptedFormats.some(format => format.toLowerCase() === fileExtension)) {
      return {
        isValid: false,
        error: `Unsupported file format. Accepted formats: ${acceptedFormats.join(', ')}`
      };
    }

    // Проверка размера
    const maxSizeBytes = maxFileSize * 1024 * 1024;
    if (file.size > maxSizeBytes) {
      return {
        isValid: false,
        error: `File size (${this.formatFileSize(file.size)}) exceeds maximum allowed size (${maxFileSize}MB)`
      };
    }

    return { isValid: true };
  }

  private handleBrowseClick = (): void => {
    if (this.fileInputRef.current && !this.props.disabled) {
      this.fileInputRef.current.click();
    }
  }

  private formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  private getProgressMessage = (): string => {
    const { progress } = this.props;
    if (!progress) return '';

    switch (progress.stage) {
      case UploadStage.UPLOADING:
        return 'Uploading file...';
      case UploadStage.PARSING:
        return 'Parsing Excel data...';
      case UploadStage.VALIDATING:
        return 'Validating file format...';
      case UploadStage.ANALYZING:
        return 'Analyzing data structure...';
      case UploadStage.COMPLETE:
        return 'Upload complete!';
      case UploadStage.ERROR:
        return 'Upload failed';
      default:
        return progress.message || 'Processing...';
    }
  }

  public render(): React.ReactElement<IFileUploaderProps> {
    const { loading, progress, disabled, acceptedFormats = ['.xlsx', '.xls', '.csv'], maxFileSize = 10 } = this.props;
    const { isDragOver, error } = this.state;

    const isActive = isDragOver && !disabled;
    const isProcessing = loading || (progress && progress.stage !== UploadStage.IDLE && progress.stage !== UploadStage.COMPLETE);

    return (
      <div className={styles.fileUploader}>
        <div 
          className={`${styles.dropZone} ${isActive ? styles.active : ''} ${disabled ? styles.disabled : ''} ${isProcessing ? styles.processing : ''}`}
          onDragEnter={this.handleDragEnter}
          onDragLeave={this.handleDragLeave}
          onDragOver={this.handleDragOver}
          onDrop={this.handleDrop}
        >
          <input
            ref={this.fileInputRef}
            type="file"
            accept={acceptedFormats.join(',')}
            onChange={this.handleFileInputChange}
            style={{ display: 'none' }}
            disabled={disabled}
          />

          {isProcessing ? (
            <div className={styles.progressContainer}>
              <div className={styles.progressIcon}>📊</div>
              <div className={styles.progressText}>
                <div className={styles.progressMessage}>{this.getProgressMessage()}</div>
                {progress && (
                  <div className={styles.progressBar}>
                    <div 
                      className={styles.progressFill}
                      style={{ width: `${progress.progress}%` }}
                    />
                  </div>
                )}
                {progress && (
                  <div className={styles.progressPercent}>{progress.progress}%</div>
                )}
              </div>
            </div>
          ) : (
            <div className={styles.uploadContent}>
              <div className={styles.uploadIcon}>📄</div>
              <div className={styles.uploadText}>
                <div className={styles.primaryText}>
                  {isActive ? 'Drop your file here' : 'Drag & drop your Excel file here'}
                </div>
                <div className={styles.secondaryText}>
                  or <button className={styles.browseButton} onClick={this.handleBrowseClick} disabled={disabled}>
                    browse files
                  </button>
                </div>
              </div>
              <div className={styles.fileInfo}>
                <div className={styles.supportedFormats}>
                  Supported formats: {acceptedFormats.join(', ')}
                </div>
                <div className={styles.maxSize}>
                  Maximum file size: {maxFileSize}MB
                </div>
              </div>
            </div>
          )}
        </div>

        {error && (
          <div className={styles.error}>
            <span className={styles.errorIcon}>⚠️</span>
            <span className={styles.errorMessage}>{error}</span>
          </div>
        )}

        {progress && progress.hasError && progress.error && (
          <div className={styles.error}>
            <span className={styles.errorIcon}>❌</span>
            <span className={styles.errorMessage}>{progress.error}</span>
          </div>
        )}
      </div>
    );
  }
}