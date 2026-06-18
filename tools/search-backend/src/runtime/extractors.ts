import mammoth = require('mammoth');
import XLSX = require('xlsx');
import { requireValue, SearchEnv } from './env';
import { sleep } from './retry';

export interface ExtractionResult {
  text: string;
  method: string;
  reason?: string;
}

const stripHtml = (value: string): string => String(value || '')
  .replace(/<script[\s\S]*?<\/script>/gi, ' ')
  .replace(/<style[\s\S]*?<\/style>/gi, ' ')
  .replace(/<[^>]+>/g, ' ')
  .replace(/&nbsp;/g, ' ')
  .replace(/&amp;/g, '&')
  .replace(/&lt;/g, '<')
  .replace(/&gt;/g, '>')
  .replace(/\s+/g, ' ')
  .trim();

export const extensionFromName = (fileName: string): string => {
  const match = String(fileName || '').match(/\.([^.]+)$/);
  return match ? match[1].toLowerCase() : '';
};

const startsWithAscii = (buffer: Buffer, value: string): boolean =>
  buffer.slice(0, value.length).toString() === value;

const bufferIncludesAscii = (buffer: Buffer, value: string): boolean =>
  buffer.toString().includes(value);

const looksLikeHtml = (buffer: Buffer): boolean => {
  const head = buffer.slice(0, 512).toString().trimStart().toLowerCase();
  return head.startsWith('<!doctype html')
    || head.startsWith('<html')
    || head.startsWith('<style')
    || head.startsWith('<div')
    || head.startsWith('<p');
};

const detectOfficePackageKind = (buffer: Buffer): 'docx' | 'xlsx' | 'pptx' | undefined => {
  if (!startsWithAscii(buffer, 'PK')) {
    return undefined;
  }

  if (bufferIncludesAscii(buffer, 'word/document.xml')) {
    return 'docx';
  }
  if (bufferIncludesAscii(buffer, 'xl/workbook.xml')) {
    return 'xlsx';
  }
  if (bufferIncludesAscii(buffer, 'ppt/presentation.xml')) {
    return 'pptx';
  }

  return undefined;
};

const detectDocumentMimeType = (
  fileName: string,
  mimeType: string | undefined,
  buffer: Buffer
): { mimeType?: string; detectedKind?: string } => {
  if (startsWithAscii(buffer, '%PDF-')) {
    return { mimeType: 'application/pdf', detectedKind: 'pdf' };
  }

  if (looksLikeHtml(buffer)) {
    return { mimeType: 'text/html', detectedKind: 'html' };
  }

  const officeKind = detectOfficePackageKind(buffer);
  if (officeKind === 'docx') {
    return {
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      detectedKind: 'docx'
    };
  }
  if (officeKind === 'xlsx') {
    return {
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      detectedKind: 'xlsx'
    };
  }
  if (officeKind === 'pptx') {
    return {
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      detectedKind: 'pptx'
    };
  }

  return { mimeType, detectedKind: extensionFromName(fileName) || undefined };
};

export const analyzeWithDocumentIntelligence = async (
  env: SearchEnv,
  buffer: Buffer,
  mimeType?: string
): Promise<ExtractionResult> => {
  if (!env.AZURE_DOC_INTEL_ENDPOINT || !env.AZURE_DOC_INTEL_KEY) {
    return { text: '', method: 'document-intelligence-unconfigured' };
  }

  const endpoint = env.AZURE_DOC_INTEL_ENDPOINT.replace(/\/$/, '');
  const apiVersion = env.AZURE_DOC_INTEL_API_VERSION || '2024-11-30';
  const model = env.AZURE_DOC_INTEL_MODEL || 'prebuilt-read';
  const analyzeUrl = `${endpoint}/documentintelligence/documentModels/${encodeURIComponent(model)}:analyze?_overload=analyzeDocument&api-version=${apiVersion}`;
  const response = await fetch(analyzeUrl, {
    method: 'POST',
    headers: {
      'Content-Type': mimeType || 'application/octet-stream',
      'Ocp-Apim-Subscription-Key': requireValue(env, 'AZURE_DOC_INTEL_KEY')
    },
    body: buffer
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Document Intelligence analyze failed ${response.status}: ${text.slice(0, 700)}`);
  }

  const operationLocation = response.headers.get('operation-location');
  if (!operationLocation) {
    throw new Error('Document Intelligence did not return an operation-location header.');
  }

  for (let attempt = 0; attempt < 60; attempt += 1) {
    await sleep(1500);
    const poll = await fetch(operationLocation, {
      headers: { 'Ocp-Apim-Subscription-Key': requireValue(env, 'AZURE_DOC_INTEL_KEY') }
    });
    if (!poll.ok) {
      const text = await poll.text();
      throw new Error(`Document Intelligence polling failed ${poll.status}: ${text.slice(0, 700)}`);
    }
    const json = await poll.json();
    if (json.status === 'succeeded') {
      return {
        text: (json.analyzeResult && json.analyzeResult.content) || '',
        method: `document-intelligence:${model}`
      };
    }
    if (json.status === 'failed') {
      throw new Error(`Document Intelligence analysis failed: ${JSON.stringify(json.error || {}).slice(0, 700)}`);
    }
  }

  throw new Error('Document Intelligence polling timed out.');
};

const extractLocalText = async (fileName: string, buffer: Buffer, detectedKind?: string): Promise<ExtractionResult> => {
  const extension = extensionFromName(fileName);
  const kind = detectedKind || extension;

  if (['txt', 'csv', 'md', 'markdown'].includes(kind)) {
    return { text: buffer.toString('utf8'), method: 'local-text' };
  }

  if (['html', 'htm', 'mhtml', 'mht'].includes(kind)) {
    return { text: stripHtml(buffer.toString('utf8')), method: 'local-html' };
  }

  if (kind === 'docx') {
    const result = await mammoth.extractRawText({ buffer });
    return { text: result.value || '', method: extension === 'docx' ? 'local-docx:mammoth' : `local-detected-docx:mammoth:${extension || 'no-extension'}` };
  }

  if (['xlsx', 'xlsm', 'xls', 'xlsb'].includes(kind)) {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const text = workbook.SheetNames.map((sheetName: string) => {
      const sheet = workbook.Sheets[sheetName];
      return `Sheet: ${sheetName}\n${XLSX.utils.sheet_to_csv(sheet)}`;
    }).join('\n\n');
    return { text, method: extension === kind ? 'local-xlsx:xlsx' : `local-detected-xlsx:xlsx:${extension || 'no-extension'}` };
  }

  return { text: '', method: 'local-unsupported' };
};

export const extractDocumentText = async (
  env: SearchEnv,
  fileName: string,
  mimeType: string | undefined,
  buffer: Buffer
): Promise<ExtractionResult> => {
  const extension = extensionFromName(fileName);
  const skippedMediaExtensions = String(env.EXTRACTION_SKIP_MEDIA_EXTENSIONS || 'mp4,m4v,mov,avi,mkv,wmv,mp3,wav,m4a')
    .split(',')
    .map((item) => item.trim().toLowerCase())
    .filter(Boolean);

  if (skippedMediaExtensions.includes(extension)) {
    return { text: '', method: 'skipped', reason: 'media_pipeline_not_enabled' };
  }

  const maxBytes = Number(env.EXTRACTION_MAX_FILE_MB || 100) * 1024 * 1024;
  if (buffer.byteLength > maxBytes) {
    return { text: '', method: 'skipped', reason: 'exceeds_size_limit' };
  }

  try {
    const detected = detectDocumentMimeType(fileName, mimeType, buffer);
    const local = await extractLocalText(fileName, buffer, detected.detectedKind);
    if (local.text && local.text.trim().length >= 50) {
      return { ...local, text: local.text.replace(/\s+/g, ' ').trim() };
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn(`Local extraction failed for ${fileName}: ${message}`);
  }

  try {
    const detected = detectDocumentMimeType(fileName, mimeType, buffer);
    const docIntel = await analyzeWithDocumentIntelligence(env, buffer, detected.mimeType);
    if (docIntel.text && docIntel.text.trim().length > 0) {
      return { ...docIntel, text: docIntel.text.replace(/\s+/g, ' ').trim() };
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return { text: '', method: 'failed', reason: message };
  }

  return { text: '', method: 'skipped', reason: 'no_extractable_text' };
};
