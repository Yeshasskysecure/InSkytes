export type ExtractionMethod =
  | 'local-text'
  | 'local-docx'
  | 'local-xlsx'
  | 'local-pdf-text'
  | 'local-pptx'
  | 'document-intelligence'
  | 'skip-media'
  | 'skip-unsupported';

export interface ExtractionPlan {
  method: ExtractionMethod;
  reason: string;
}

const normalizeExtension = (fileNameOrExtension: string): string =>
  fileNameOrExtension.split('.').pop()?.toLowerCase().trim() || '';

export const planExtraction = (fileNameOrExtension: string): ExtractionPlan => {
  const extension = normalizeExtension(fileNameOrExtension);

  if (['txt', 'csv', 'md', 'markdown', 'html', 'htm'].includes(extension)) {
    return { method: 'local-text', reason: 'Plain text and markup can be parsed without OCR.' };
  }

  if (extension === 'docx') {
    return { method: 'local-docx', reason: 'Born-digital DOCX text is efficient to extract locally.' };
  }

  if (['xlsx', 'xlsm', 'xls', 'csv'].includes(extension)) {
    return { method: 'local-xlsx', reason: 'Spreadsheet text and cells can be extracted locally first.' };
  }

  if (extension === 'pdf') {
    return {
      method: 'local-pdf-text',
      reason: 'Try embedded PDF text first; fall back to Document Intelligence for scanned PDFs.'
    };
  }

  if (['pptx', 'ppt'].includes(extension)) {
    return {
      method: 'local-pptx',
      reason: 'Try slide text locally first; use Document Intelligence if layout or OCR is needed.'
    };
  }

  if (['png', 'jpg', 'jpeg', 'tif', 'tiff', 'bmp', 'heif'].includes(extension)) {
    return { method: 'document-intelligence', reason: 'Image files require OCR.' };
  }

  if (['mp4', 'mov', 'avi', 'mkv', 'wmv', 'mp3', 'wav', 'm4a'].includes(extension)) {
    return { method: 'skip-media', reason: 'Media requires a transcript pipeline, which is phase 2.' };
  }

  return { method: 'skip-unsupported', reason: `Unsupported file extension: ${extension || 'unknown'}.` };
};
