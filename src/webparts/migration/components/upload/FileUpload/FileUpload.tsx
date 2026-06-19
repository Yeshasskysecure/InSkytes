import * as React from 'react';
import { IFileUploadProps } from './IFileUploadProps';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './FileUpload.module.scss';
import { MetadataForm } from '../../metadata/MetadataForm/MetadataForm';
import { IKShellFooter, IKShellHeader } from '../../shell/IKShellChrome';
import { NAV_PATHS } from '../../../services/permalinkService';
import { DocumentParser } from '../../../services/DocumentParser';
import { AzureOpenAIService } from '../../../services/AzureOpenAIService';
import {
  AZURE_OPENAI_API_KEY,
  AZURE_OPENAI_ENDPOINT,
  AZURE_OPENAI_DEPLOYMENT
} from '../../../services/SearchConfig';
import { KnowledgeSearchApiClient } from '../../../services/KnowledgeSearchApiClient';
import { VideoAnalysis } from '../VideoAnalysis/VideoAnalysis';
import {
  fetchAllTaxonomyOptions,
  ITaxonomyFieldOptions,
  ITaxonomyTerm,
  matchTaxonomyTerm,
  TAXONOMY_FIELD_CONFIGS
} from '../../../services/TaxonomyService';
import {
  buildBuDepartmentFieldUpdates,
  KM_DATA_HUB_BU_FIELD_INTERNAL_NAME,
  KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME
} from '../../../utils/buDepartmentSelections';
import { fetchDocumentTypeTerms, Term } from '../../../utils/termStore';
import {
  configurePermalinkService,
  buildPrettyUrl,
  buildAssetUrl,
} from '../../../services/permalinkService';
import { buildZeroKMDocumentMetrics } from '../../../services/kmMetricsService';
import { emitDocumentDataChanged } from '../../../services/documentChangeEvents';
import { getKmsUsers } from '../../../services/KmsUsersService';
import {
  checkoutFile,
  checkinMajor,
  saveMetadataAndCheckin,
  undoCheckout,
  uploadNewFileVersion,
  updateWithoutVersion
} from '../../../services/kmVersionControl';
import { CACHE_KEYS, COLUMN_NAMES, LIBRARY_NAMES, LIST_NAMES } from '../../../config/appConfig';

const MAX_FILES = 5;
const LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;
const SHAREPOINT_SHORT_TEXT_MAX_LENGTH = 255;
const MAX_SENSITIVE_TERMS_LENGTH = 1000;
const OPEN_ROUTED_DOCUMENT_EVENT = 'ikn:open-routed-document';
const UNSUPPORTED_FILE_FORMAT_MESSAGE = 'Unsupported file format. Please upload a valid document';
const SUPPORTED_DOCUMENT_EXTENSIONS = [
  'pdf',
  'doc',
  'docx',
  'ppt',
  'pptx',
  'xls',
  'xlsx',
  'xlsm',
  'xlsb',
  'xltx',
  'xltm',
  'csv',
  'txt',
  'mhtml',
  'mht',
  'mpp'
];
const SUPPORTED_MEDIA_EXTENSIONS = [
  'mp4',
  'mp3',
  'wav',
  'mov',
  'avi',
  'mkv',
  'webm',
  'm4a',
  'aac',
  'flac',
  'ogg'
];
const IMAGE_UPLOAD_EXTENSIONS = [
  'jpg',
  'jpeg',
  'png',
  'gif',
  'bmp',
  'webp',
  'svg',
  'tif',
  'tiff',
  'heic',
  'heif'
];
const DOCUMENT_UPLOAD_ACCEPT = ".pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.xlsm,.xlsb,.xltx,.xltm,.csv,.txt,.mhtml,.mht,.mpp,text/csv,text/plain";
const REPLACE_UPLOAD_ACCEPT = `${DOCUMENT_UPLOAD_ACCEPT},.mp4,.mp3,.wav,.mov,.avi,.mkv,.webm,.m4a,.aac,.flac,.ogg,video/*,audio/*`;
const MEDIA_UPLOAD_ACCEPT = "video/*,audio/*";

const isDebugLoggingEnabled = (): boolean => {
  try {
    return typeof window !== 'undefined' &&
      window.localStorage?.getItem('IKNOWLEDGE_DEBUG_LOGS') === 'true';
  } catch {
    return false;
  }
};

const debugLog = (...args: unknown[]): void => {
  if (isDebugLoggingEnabled()) {
    console.log(...args);
  }
};

interface FileData {
  file: File;
  itemId: number;
  metadata: Record<string, any>;
}

interface IKMDataHubFieldMap {
  status: string;
  published: string;
  title: string;
  description: string;
  author: string;
  documentType: string;
  documentTypeNote: string;
  client: string;
  clientNote: string;
  geography: string;
  geographyNote: string;
  diseaseArea: string;
  diseaseAreaNote: string;
  therapyArea: string;
  therapyAreaNote: string;
  sensitiveTerms: string;
  views: string;
  likes: string;
  comments: string;
  downloads: string;
  follow: string;
  share: string;
  bookmark: string;
  projectId: string;
  edited: string;
  editedBy: string;
}

interface IKMDataHubItemSnapshot {
  itemId: number;
  fileRef: string;
  item: Record<string, any>;
  textValues: Record<string, any>;
}

interface ISuccessDocumentLink {
  id: string;
  title: string;
  url: string;
}

interface ISessionUploadedDoc {
  id: number;
  name: string;
  time: string;
  serverRelativeUrl?: string;
}

interface IRecentUploadListItem {
  ID: number;
  FileLeafRef: string;
  Created: string;
  FileRef: string;
}

type TSuccessNotificationType = 'success' | 'success-media' | 'success-update' | 'error' | 'no-content';

interface IUploadSuccessRestorePayload {
  variant: 'default' | 'kmArtifact';
  successNotification: {
    message: string;
    type: TSuccessNotificationType;
  };
  successDocumentUrl: string | null;
  lastUploadedDocId: string | null;
  successDocumentLinks: ISuccessDocumentLink[];
}

type TProcessingPhase = 'idle' | 'analysis' | 'upload' | 'generic';

const MAX_MEDIA_FILES = 1;
const DOCUMENT_UPLOAD_REJECTS_MEDIA_MESSAGE = "Media Upload not supported here\nUse 'Upload Media' to add media files";
const MEDIA_UPLOAD_REJECTS_DOCUMENT_MESSAGE = "Document upload not supported here\nUse 'Upload document' to add document files";
const getMediaUploadLimitMessage = (): string =>
  `Upload limit exceeded\nYou can upload a maximum of ${MAX_MEDIA_FILES} media file at a time`;
const getDocumentUploadLimitMessage = (): string =>
  `Upload limit exceeded\nYou can upload a maximum of ${MAX_FILES} files at a time`;
const getDuplicateUploadMessage = (assetType: 'document' | 'media', hasTitleDuplicate: boolean, hasFileNameDuplicate: boolean): string => {
  if (hasTitleDuplicate && hasFileNameDuplicate) {
    return `A ${assetType} with the same title and file name already exists. Please update them and try again.`;
  }

  if (hasTitleDuplicate) {
    return `A ${assetType} with this title already exists. Please change the title and try again.`;
  }

  return `A ${assetType} with this file name already exists. Please rename the file and try again.`;
};
const escapeODataString = (value: string): string => (value || '').replace(/'/g, "''");
const normalizeKqlPhraseValue = (value: string): string => (value || '').replace(/"/g, ' ').trim();
const escapeXmlValue = (value: string): string =>
  (value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
const formatSharePointDateFieldValue = (date: Date = new Date()): string => {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  }).formatToParts(date);
  const getPart = (type: string): string =>
    parts.find((part) => part.type === type)?.value || '';

  return `${getPart('month')}/${getPart('day')}/${getPart('year')} ${getPart('hour')}:${getPart('minute')} ${getPart('dayPeriod')}`;
};
const UPLOAD_SUCCESS_RESTORE_STORAGE_KEY = CACHE_KEYS.uploadSuccessRestore;

const getFolderServerRelativeUrl = (fileRef: string): string => {
  const normalizedFileRef = (fileRef || '').trim();
  const lastSlashIndex = normalizedFileRef.lastIndexOf('/');
  return lastSlashIndex > 0 ? normalizedFileRef.slice(0, lastSlashIndex) : LIBRARY_NAME;
};

const getLeafFileName = (fileRef: string, fallbackFileName: string): string => {
  const normalizedFileRef = (fileRef || '').trim();
  const leafName = normalizedFileRef.split('/').pop() || '';
  return leafName || fallbackFileName;
};

const normalizeComparisonValue = (value: string | null | undefined): string => (value || '').trim().toLowerCase();

const getFileExtension = (fileName: string): string => {
  const normalizedName = (fileName || '').trim().toLowerCase();
  const extensionStartIndex = normalizedName.lastIndexOf('.');
  return extensionStartIndex >= 0 ? normalizedName.slice(extensionStartIndex + 1) : '';
};

const isImageUploadFile = (file: File): boolean =>
  (file.type || '').toLowerCase().startsWith('image/') ||
  IMAGE_UPLOAD_EXTENSIONS.indexOf(getFileExtension(file.name)) !== -1;

const isSupportedDocumentUploadFile = (file: File): boolean =>
  !isImageUploadFile(file) &&
  SUPPORTED_DOCUMENT_EXTENSIONS.indexOf(getFileExtension(file.name)) !== -1;

const isSupportedMediaUploadFile = (file: File): boolean =>
  SUPPORTED_MEDIA_EXTENSIONS.indexOf(getFileExtension(file.name)) !== -1;

const isSupportedReplacementUploadFile = (file: File): boolean =>
  isSupportedDocumentUploadFile(file) || isSupportedMediaUploadFile(file);

const normalizeMultiValueComparison = (value: string | null | undefined): string =>
  (value || '')
    .split(/[;,#\r\n]+/)
    .map((entry) => entry.trim().toLowerCase())
    .filter(Boolean)
    .sort()
    .join('|');

const normalizeTaxonomyTermsForComparison = (terms: ITaxonomyTerm[] | null | undefined): string =>
  (terms || [])
    .map((term) => (term?.label || '').trim().toLowerCase())
    .filter(Boolean)
    .sort()
    .join('|');

const buildClaimsValues = (upns: string[] | null | undefined): string =>
  JSON.stringify(
    (upns || [])
      .map((upn) => (upn || '').trim())
      .filter(Boolean)
      .map((upn) => ({ Key: `i:0#.f|membership|${upn}` }))
  );

const joinUniqueValues = (values: Array<string | null | undefined>): string => {
  const uniqueValues = values
    .reduce((accumulator: string[], value) => accumulator.concat((value || '').split(/[\n,;]+/)), [])
    .map((value) => value.trim())
    .filter((value, index, array) => value.length > 0 && array.indexOf(value) === index);

  return uniqueValues.join(', ');
};

const cleanTextField = (value: string, maxLength?: number): string => {
  if (!value) return '';

  let cleaned = value
    // Remove special characters
    .replace(/[$%@*+?!#&^~`|\\<>]/g, '')
    // Remove file extensions
    .replace(/\.(pdf|docx|doc|pptx|ppt|xlsx|xls|txt|csv|zip|png|jpg|jpeg|mp4|mp3|wav|mov|avi|mkv|webm|m4a|aac|flac|ogg)$/gi, '')
    // Remove version patterns like v1, v2.0, v3, -v1, V1.0
    .replace(/[\s-]?v\d+(\.\d+)*[\s_-]?/gi, ' ')
    // Remove version words like *final, draft, copy
    .replace(/[\s-]?(final|draft|copy|revised|updated|new|old|backup|temp)[\s-]?/gi, ' ')
    // Remove date patterns like 20240101, -2024-01-01, 20240101
    .replace(/[\s-]?\d{4}[-*]?\d{2}[-*]?\d{2}[\s*-]?/g, ' ')
    // Remove standalone IDs like -001, *123, #456
    .replace(/[\s*-]#?\d{2,6}[\s_-]?/g, ' ')
    // Remove multiple spaces
    .replace(/\s{2,}/g, ' ')
    // Trim
    .trim();

  if (maxLength && cleaned.length > maxLength) {
    cleaned = cleaned.substring(0, maxLength).trim();
  }

  return cleaned;
};

const logTaxonomyAvailability = (allTaxonomyOptions?: ITaxonomyFieldOptions | null): void => {
  const taxonomyAvailability = {
    therapyAreaTerms: (allTaxonomyOptions as any)?.therapyAreaTerms || allTaxonomyOptions?.therapyArea,
    diseaseAreaTerms: (allTaxonomyOptions as any)?.diseaseAreaTerms || allTaxonomyOptions?.diseaseArea
  };

  debugLog('Taxonomy loaded - therapy terms available:', taxonomyAvailability?.therapyAreaTerms?.length || 0);
  debugLog('Taxonomy loaded - disease terms available:', taxonomyAvailability?.diseaseAreaTerms?.length || 0);
};

const normalizeSensitiveTermsValue = (value: string | null | undefined): string =>
  (value || '').trim().slice(0, MAX_SENSITIVE_TERMS_LENGTH);

const buildClientTermsPromptSection = (taxonomyOptions?: ITaxonomyFieldOptions | null): string => {
  const maxClientTermsPromptChars = 24000;
  const clientTerms = (taxonomyOptions?.client || [])
    .map((term) => (term?.label || '').trim())
    .filter((label, index, array) => !!label && array.indexOf(label) === index)
    .sort((left, right) => left.localeCompare(right));

  const lines: string[] = [];
  let totalLength = 0;

  for (let index = 0; index < clientTerms.length; index++) {
    const line = `- "${clientTerms[index]}"`;
    if (totalLength + line.length + 1 > maxClientTermsPromptChars) {
      break;
    }

    lines.push(line);
    totalLength += line.length + 1;
  }

  return lines.join('\n');
};

const openSuccessDocumentInApp = (itemId: string, url: string): void => {
  window.history.pushState({ sameTabNav: true }, document.title, url);
  window.dispatchEvent(
    new CustomEvent(OPEN_ROUTED_DOCUMENT_EVENT, {
      detail: {
        documentId: Number(itemId)
      }
    })
  );
};

const clearUploadSuccessRestorePayload = (): void => {
  try {
    sessionStorage.removeItem(UPLOAD_SUCCESS_RESTORE_STORAGE_KEY);
  } catch (error) {
    console.warn('Unable to clear upload success restore payload.', error);
  }
};

const readUploadSuccessRestorePayload = (): IUploadSuccessRestorePayload | null => {
  try {
    const rawValue = sessionStorage.getItem(UPLOAD_SUCCESS_RESTORE_STORAGE_KEY);
    if (!rawValue) {
      return null;
    }

    const parsedValue = JSON.parse(rawValue) as IUploadSuccessRestorePayload;
    if (!parsedValue?.successNotification || !Array.isArray(parsedValue.successDocumentLinks)) {
      clearUploadSuccessRestorePayload();
      return null;
    }

    return parsedValue;
  } catch (error) {
    console.warn('Unable to read upload success restore payload.', error);
    clearUploadSuccessRestorePayload();
    return null;
  }
};

const uniqueTaxonomyTerms = (terms: ITaxonomyTerm[]): ITaxonomyTerm[] => {
  const seen: Record<string, boolean> = {};
  const uniqueTerms: ITaxonomyTerm[] = [];

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index];
    if (!term?.id || seen[term.id]) {
      continue;
    }

    seen[term.id] = true;
    uniqueTerms.push(term);
  }

  return uniqueTerms;
};

const normalizeTaxonomyTerms = (
  rawValue: any,
  fallbackTerm: ITaxonomyTerm | null | undefined
): ITaxonomyTerm[] => {
  if (Array.isArray(rawValue)) {
    return uniqueTaxonomyTerms(rawValue.filter(Boolean));
  }

  if (fallbackTerm) {
    return [fallbackTerm];
  }

  return [];
};

const normalizeSelectedBuDepartmentTerms = (
  selectedTerms: Array<{ id?: string }> | null | undefined,
  taxonomyOptions?: ITaxonomyFieldOptions
): ITaxonomyTerm[] => {
  if (!taxonomyOptions || !Array.isArray(selectedTerms) || selectedTerms.length === 0) {
    return [];
  }

  const matchedTerms = selectedTerms
    .map((selectedTerm) => {
      const selectedTermId = String(selectedTerm?.id || '').trim().toLowerCase();
      if (!selectedTermId) {
        return null;
      }

      return taxonomyOptions.buDepartment.find((option) => option.id.toLowerCase() === selectedTermId) || null;
    })
    .filter((term): term is ITaxonomyTerm => !!term);

  return uniqueTaxonomyTerms(matchedTerms);
};

const normalizeBuDepartmentTerms = (
  metadata: Record<string, any>,
  taxonomyOptions?: ITaxonomyFieldOptions
): ITaxonomyTerm[] => {
  const selectedTerms = Array.isArray(metadata.selectedDepts) && metadata.selectedDepts.length > 0
    ? metadata.selectedDepts
    : metadata.selectedBUs;
  const selectedBuDepartmentTerms = normalizeSelectedBuDepartmentTerms(selectedTerms, taxonomyOptions);
  if (selectedBuDepartmentTerms.length > 0) {
    return selectedBuDepartmentTerms;
  }

  return [];
};

const getTaxonomyValue = (term: ITaxonomyTerm | null | undefined): string =>
  term ? `${term.label}|${term.id};` : '';

const getMultiTaxonomyValue = (terms: ITaxonomyTerm[] | null | undefined): string =>
  (terms || [])
    .filter((term) => !!term?.id)
    .map((term) => `${term.label}|${term.id}`)
    .join(';');

const getTaxonomyNoteValue = (term: ITaxonomyTerm | null | undefined): string =>
  term ? `-1;#${term.label}|${term.id}` : '';

const getMultiTaxonomyNoteValue = (terms: ITaxonomyTerm[] | null | undefined): string =>
  (terms || [])
    .filter((term) => !!term?.id)
    .map((term) => `-1;#${term.label}|${term.id}`)
    .join(';#');

const getTaxonomyLabel = (term: ITaxonomyTerm | ITaxonomyTerm[] | null | undefined, fallback?: string): string => {
  if (Array.isArray(term)) {
    return term.map((item) => item.label).filter((label) => !!label).join(', ') || fallback || '';
  }

  return term?.label || fallback || '';
};

const getSelectedTermNames = (terms: Array<{ name?: string }> | null | undefined): string =>
  (terms || [])
    .map((term) => term?.name || '')
    .filter((name) => !!name)
    .join(', ');

const SPECIAL_CLIENT_VALUES = ['Multi-Client', 'Others', 'Not Applicable'];
const SPECIAL_GEOGRAPHY_VALUES = ['General', 'Not Applicable'];
const SPECIAL_THERAPY_VALUES = ['Cross-Therapy', 'Not Applicable'];
const SPECIAL_DISEASE_VALUES = ['Cross-Disease', 'Not Applicable'];

const normalizeMissingClassificationValue = (value: any, fallback: string = 'Not Applicable'): string => {
  const text = typeof value === 'string' ? value.trim() : '';
  const normalized = text.toLowerCase();

  if (!text || normalized === 'not found' || normalized === 'n/a' || normalized === 'na' || normalized === 'none' || normalized === 'unknown') {
    return fallback;
  }

  return text;
};

const isSpecialSelectionValue = (
  value: string | null | undefined,
  specialValues: string[]
): boolean => {
  const normalizedValue = normalizeComparisonValue(value);
  return !!normalizedValue && specialValues.some((specialValue) => normalizedValue === normalizeComparisonValue(specialValue));
};

const findSpecialTaxonomyTerm = (
  value: string | null | undefined,
  terms: ITaxonomyTerm[] | null | undefined
): ITaxonomyTerm | null => {
  const normalizedValue = normalizeComparisonValue(value);
  if (!normalizedValue || !terms || terms.length === 0) {
    return null;
  }

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index] as ITaxonomyTerm & { name?: string };
    if (
      normalizeComparisonValue(term.label) === normalizedValue ||
      normalizeComparisonValue(term.name) === normalizedValue
    ) {
      return term;
    }
  }

  return null;
};

const findExactTaxonomyTerm = (
  value: string | null | undefined,
  terms: ITaxonomyTerm[] | null | undefined
): ITaxonomyTerm | null => {
  const normalizedValue = normalizeComparisonValue(value);
  if (!normalizedValue || !terms || terms.length === 0) {
    return null;
  }

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index] as ITaxonomyTerm & { name?: string };
    if (
      normalizeComparisonValue(term.label) === normalizedValue ||
      normalizeComparisonValue(term.name) === normalizedValue ||
      normalizeComparisonValue(term.path) === normalizedValue
    ) {
      return term;
    }
  }

  return null;
};

const inferTaxonomyTermFromContent = (
  content: string,
  terms: ITaxonomyTerm[]
): ITaxonomyTerm | null => {
  const normalizedContent = (content || '').toLowerCase();
  if (!normalizedContent.trim() || !terms || terms.length === 0) {
    return null;
  }

  let bestTerm: ITaxonomyTerm | null = null;
  let bestScore = 0;

  const contentTokens = normalizedContent
    .replace(/[^a-z0-9]+/g, ' ')
    .split(/\s+/)
    .filter((token) => token.length > 1);

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index];
    const label = (term.label || '').toLowerCase().trim();
    const path = (term.path || '').toLowerCase().trim();

    if (!label) {
      continue;
    }

    let score = 0;

    if (normalizedContent.indexOf(label) !== -1) {
      score = Math.max(score, label.length + 10);
    }

    if (path && normalizedContent.indexOf(path) !== -1) {
      score = Math.max(score, path.length + 20);
    }

    const labelTokens = label
      .replace(/[^a-z0-9]+/g, ' ')
      .split(/\s+/)
      .filter((token) => token.length > 1);

    let tokenScore = 0;
    for (let tokenIndex = 0; tokenIndex < labelTokens.length; tokenIndex++) {
      const labelToken = labelTokens[tokenIndex];
      for (let contentIndex = 0; contentIndex < contentTokens.length; contentIndex++) {
        const contentToken = contentTokens[contentIndex];
        if (labelToken === contentToken) {
          tokenScore += 10;
        } else if (labelToken.indexOf(contentToken) !== -1 || contentToken.indexOf(labelToken) !== -1) {
          tokenScore += 5;
        }
      }
    }

    if (tokenScore > 0) {
      score = Math.max(score, tokenScore);
    }

    const pathSegments = path.split(' > ').map((segment) => segment.trim()).filter((segment) => segment.length > 0);
    let matchedSegments = 0;
    for (let segmentIndex = 0; segmentIndex < pathSegments.length; segmentIndex++) {
      if (normalizedContent.indexOf(pathSegments[segmentIndex]) !== -1) {
        matchedSegments++;
      }
    }

    if (matchedSegments > 0) {
      score = Math.max(score, matchedSegments * 8 + label.length);
    }

    if (score > bestScore) {
      bestScore = score;
      bestTerm = term;
    }
  }

  return bestScore >= 10 ? bestTerm : null;
};

const normalizeMetadataForKMDataHub = (
  metadata: Record<string, any> = {},
  taxonomyOptions?: ITaxonomyFieldOptions
): Record<string, any> => {
  const aiResult = metadata;
  const buDepartmentTerms = normalizeBuDepartmentTerms(metadata, taxonomyOptions);
  const documentType = metadata.documentType || '';
  const client = normalizeMissingClassificationValue(metadata.client);
  const geography = normalizeMissingClassificationValue(metadata.geography || metadata.region);
  const diseaseArea = normalizeMissingClassificationValue(metadata.diseaseArea);
  const therapyArea = normalizeMissingClassificationValue(metadata.therapyArea);
  const clientAIValue = client;
  const geographyAIValue = geography;
  const diseaseAIValue = diseaseArea;
  const therapyAIValue = therapyArea;
  const therapyTerms = taxonomyOptions?.therapyArea;
  const isSpecialClientValue = isSpecialSelectionValue(clientAIValue, SPECIAL_CLIENT_VALUES);
  const isSpecialGeographyValue = isSpecialSelectionValue(geographyAIValue, SPECIAL_GEOGRAPHY_VALUES);
  const isSpecialTherapyValue = isSpecialSelectionValue(therapyAIValue, SPECIAL_THERAPY_VALUES);
  const isSpecialDiseaseValue = isSpecialSelectionValue(diseaseAIValue, SPECIAL_DISEASE_VALUES);

  debugLog('Client AI value:', clientAIValue, '| Is special:', isSpecialClientValue);
  debugLog('Geography AI value:', geographyAIValue, '| Is special:', isSpecialGeographyValue);
  debugLog('Therapy AI value:', therapyAIValue, '| Is special:', isSpecialTherapyValue);
  debugLog('Disease AI value:', diseaseAIValue, '| Is special:', isSpecialDiseaseValue);
  debugLog('Therapy area AI value:', therapyAIValue);
  debugLog('Is special therapy value:', SPECIAL_THERAPY_VALUES.some(v => therapyAIValue.toLowerCase() === v.toLowerCase()));
  debugLog('Available therapy terms count:', therapyTerms?.length);
  debugLog('Sample therapy terms:', therapyTerms?.slice(0, 5).map((t: any) => t.label || t.name));

  const specialClientTerm = isSpecialClientValue
    ? findSpecialTaxonomyTerm(clientAIValue, taxonomyOptions?.client)
    : null;
  const specialGeographyTerm = isSpecialGeographyValue
    ? findSpecialTaxonomyTerm(geographyAIValue, taxonomyOptions?.geography)
    : null;
  const specialTherapyAreaTerm = isSpecialTherapyValue
    ? findSpecialTaxonomyTerm(therapyAIValue, taxonomyOptions?.therapyArea)
    : null;
  const specialDiseaseAreaTerm = isSpecialDiseaseValue
    ? findSpecialTaxonomyTerm(diseaseAIValue, taxonomyOptions?.diseaseArea)
    : null;
  const inferenceContent = [
    metadata.title,
    metadata.description,
    metadata.abstract,
    metadata.documentType,
    !isSpecialClientValue ? metadata.client : '',
    !isSpecialGeographyValue ? metadata.geography : '',
    !isSpecialGeographyValue ? metadata.region : '',
    !isSpecialDiseaseValue ? metadata.diseaseArea : '',
    !isSpecialTherapyValue ? metadata.therapyArea : ''
  ].filter((value) => typeof value === 'string' && value.trim().length > 0).join(' | ');

  const inferredDocumentTypeTerm = taxonomyOptions
    ? inferTaxonomyTermFromContent(inferenceContent, taxonomyOptions.documentType)
    : null;
  const fallbackClientTerm = findSpecialTaxonomyTerm('Not Applicable', taxonomyOptions?.client);
  const inferredGeographyTerm = taxonomyOptions
    ? inferTaxonomyTermFromContent(inferenceContent, taxonomyOptions.geography)
    : null;
  const inferredDiseaseAreaTerm = taxonomyOptions
    ? inferTaxonomyTermFromContent(inferenceContent, taxonomyOptions.diseaseArea)
    : null;
  const inferredTherapyAreaTerm = taxonomyOptions
    ? inferTaxonomyTermFromContent(inferenceContent, taxonomyOptions.therapyArea)
    : null;
  const resolvedClientTerm = metadata.clientTerm ||
    specialClientTerm ||
    (!isSpecialClientValue && taxonomyOptions ? findExactTaxonomyTerm(client, taxonomyOptions.client) : null) ||
    fallbackClientTerm;
  const resolvedGeographyTerm = metadata.geographyTerm ||
    specialGeographyTerm ||
    (!isSpecialGeographyValue && taxonomyOptions ? matchTaxonomyTerm(geography, taxonomyOptions.geography) : null) ||
    (!isSpecialGeographyValue ? inferredGeographyTerm : null);
  const resolvedDiseaseAreaTerm = metadata.diseaseAreaTerm ||
    specialDiseaseAreaTerm ||
    (!isSpecialDiseaseValue && taxonomyOptions ? matchTaxonomyTerm(diseaseArea, taxonomyOptions.diseaseArea) : null) ||
    (!isSpecialDiseaseValue ? inferredDiseaseAreaTerm : null);
  const resolvedTherapyAreaTerm = metadata.therapyAreaTerm ||
    specialTherapyAreaTerm ||
    (!isSpecialTherapyValue && taxonomyOptions ? matchTaxonomyTerm(therapyArea, taxonomyOptions.therapyArea) : null) ||
    (!isSpecialTherapyValue ? inferredTherapyAreaTerm : null);

  const normalizedData = {
    ...metadata,
    title: cleanTextField(String(aiResult.title || ''), 255),
    description: cleanTextField(String(aiResult.description || aiResult.abstract || ''), SHAREPOINT_SHORT_TEXT_MAX_LENGTH),
    selectedAuthorUpns: Array.isArray(metadata.selectedAuthorUpns)
      ? metadata.selectedAuthorUpns.filter(Boolean)
      : metadata.selectedAuthorUpn
        ? [metadata.selectedAuthorUpn].filter(Boolean)
        : [],
    selectedBUs: Array.isArray(metadata.selectedBUs) ? metadata.selectedBUs.filter(Boolean) : [],
    selectedDepts: Array.isArray(metadata.selectedDepts) ? metadata.selectedDepts.filter(Boolean) : [],
    sensitiveTerms: normalizeSensitiveTermsValue(
      metadata.sensitiveTerms || joinUniqueValues([
        metadata.emails,
        metadata.phones,
        metadata.ids,
        metadata.pricing,
        metadata.sensitive
      ])
    ),
    bu: '',
    department: '',
    subDepartment: '',
    buDepartment: '',
    buDepartmentTerms,
    buDepartmentTerm: buDepartmentTerms[0] || null,
    documentTypeTerm: metadata.documentTypeTerm || (taxonomyOptions ? matchTaxonomyTerm(documentType, taxonomyOptions.documentType) : null) || inferredDocumentTypeTerm,
    clientTerm: resolvedClientTerm,
    geographyTerm: resolvedGeographyTerm,
    diseaseAreaTerm: resolvedDiseaseAreaTerm,
    therapyAreaTerm: resolvedTherapyAreaTerm,
    documentTypeTerms: normalizeTaxonomyTerms(metadata.documentTypeTerms, metadata.documentTypeTerm || (taxonomyOptions ? matchTaxonomyTerm(documentType, taxonomyOptions.documentType) : null) || inferredDocumentTypeTerm),
    clientTerms: normalizeTaxonomyTerms(metadata.clientTerms, resolvedClientTerm),
    geographyTerms: normalizeTaxonomyTerms(metadata.geographyTerms, resolvedGeographyTerm),
    diseaseAreaTerms: normalizeTaxonomyTerms(metadata.diseaseAreaTerms, resolvedDiseaseAreaTerm),
    therapyAreaTerms: normalizeTaxonomyTerms(metadata.therapyAreaTerms, resolvedTherapyAreaTerm)
  };

  debugLog('Raw AI title:', aiResult.title);
  debugLog('Cleaned title:', normalizedData.title);
  debugLog('Raw AI description:', aiResult.description);
  debugLog('Cleaned description:', normalizedData.description);

  return normalizedData;
};

const buildKMDataHubTaxonomyFormValues = (
  metadata: Record<string, any>,
  fieldMap: IKMDataHubFieldMap
): Array<{ FieldName: string; FieldValue: string }> => [
  ...buildBuDepartmentFieldUpdates(metadata.selectedBUs || [], metadata.selectedDepts || []),
  { FieldName: fieldMap.documentType, FieldValue: getMultiTaxonomyValue(metadata.documentTypeTerms) },
  { FieldName: fieldMap.client, FieldValue: getMultiTaxonomyValue(metadata.clientTerms) },
  { FieldName: fieldMap.geography, FieldValue: getMultiTaxonomyValue(metadata.geographyTerms) },
  { FieldName: fieldMap.diseaseArea, FieldValue: getMultiTaxonomyValue(metadata.diseaseAreaTerms) },
  { FieldName: fieldMap.therapyArea, FieldValue: getMultiTaxonomyValue(metadata.therapyAreaTerms) }
].filter((field) => !!field.FieldValue);

export const FileUpload: React.FC<IFileUploadProps> = (props) => {
  const effectiveSidebarOffset = Math.max(0, props.sidebarOffset ?? 0);
  const isKmArtifactUpload = props.variant === 'kmArtifact';
  const isReplaceMode = props.variant === 'replace';
  const isCompactFlow = props.flowDensity === 'compact';
  const overlayStyle = React.useMemo<React.CSSProperties>(
    () => ({
      ['--upload-sidebar-offset' as string]: `${effectiveSidebarOffset}px`
    }),
    [effectiveSidebarOffset]
  );

  const [dragOver, setDragOver] = React.useState(false);
  const [uploadedFiles, setUploadedFiles] = React.useState<File[]>([]);
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [processingProgress, setProcessingProgress] = React.useState<string>('');
  const [processingFileCount, setProcessingFileCount] = React.useState<number>(0);
  const [filesData, setFilesData] = React.useState<FileData[]>([]);
  const [processingError, setProcessingError] = React.useState<string | null>(null);
  const [processingPhase, setProcessingPhase] = React.useState<TProcessingPhase>('idle');
  const [showForm, setShowForm] = React.useState(false);
  const [activeMetadataFileIndex, setActiveMetadataFileIndex] = React.useState<number>(0);
  const [multiMetadataDrafts, setMultiMetadataDrafts] = React.useState<Record<number, Record<string, any>>>({});
  const [multiMetadataValidity, setMultiMetadataValidity] = React.useState<Record<number, boolean>>({});
  const [mediaFile, setMediaFile] = React.useState<File | null>(null);
  const [taxonomyOptions, setTaxonomyOptions] = React.useState<ITaxonomyFieldOptions | null>(null);
  const [projectName, setProjectName] = React.useState<string | null>(null);
  const [projectBU, setProjectBU] = React.useState<string | null>(null);
  const [isProjectValid, setIsProjectValid] = React.useState<boolean>(false);
  const [sessionUploadedDocs, setSessionUploadedDocs] = React.useState<ISessionUploadedDoc[]>([]);
  const [showSuccessModal, setShowSuccessModal] = React.useState<boolean>(false);
  const [lastUploadedDocId, setLastUploadedDocId] = React.useState<string | null>(null);
  const [successNotification, setSuccessNotification] = React.useState<{ message: string; type: TSuccessNotificationType } | null>(null);
  const [successDocumentUrl, setSuccessDocumentUrl] = React.useState<string | null>(null);
  const [successDocumentLinks, setSuccessDocumentLinks] = React.useState<ISuccessDocumentLink[]>([]);
  const [showConfirmation, setShowConfirmation] = React.useState<boolean>(false);
  const [projectStatus, setProjectStatus] = React.useState<string | null>(null);
  const [projectCount, setProjectCount] = React.useState<number>(0);
  const [currentProjectId, setCurrentProjectId] = React.useState<string>(props.projectId || '');

  const fileInputRef = React.useRef<HTMLInputElement | null>(null);
  const openAIService = React.useRef(new AzureOpenAIService(undefined, props.context));
  const knowledgeSearchApiClient = React.useMemo(
    () => new KnowledgeSearchApiClient(undefined, props.context),
    [props.context]
  );
  const isSubmittingRef = React.useRef(false);
  const taxonomyOptionsRef = React.useRef<ITaxonomyFieldOptions | null>(null);
  const kmDataHubFieldMapRef = React.useRef<IKMDataHubFieldMap | null>(null);
  const hasHydratedRestorePayloadRef = React.useRef(false);

  const mediaInputRef = React.useRef<HTMLInputElement | null>(null);
  const shouldShowReviewerFields = false;

  const showUploadValidationMessage = React.useCallback((message: string): void => {
    setSuccessNotification({ message, type: 'error' });
  }, []);

  const isMediaUploadFile = React.useCallback((file: File): boolean =>
    isSupportedMediaUploadFile(file),
  []);

  const resetFileInputs = React.useCallback((): void => {
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }

    if (mediaInputRef.current) {
      mediaInputRef.current.value = '';
    }
  }, []);

  const removeMetadataFileAtIndex = React.useCallback((removeIndex: number): void => {
    setFilesData((currentFilesData) => {
      const nextFilesData = currentFilesData.filter((_fileData, index) => index !== removeIndex);

      if (nextFilesData.length === 0) {
        setUploadedFiles([]);
        setShowForm(false);
      } else {
        setUploadedFiles((currentUploadedFiles) => currentUploadedFiles.filter((_file, index) => index !== removeIndex));
      }

      return nextFilesData;
    });

    setMultiMetadataDrafts((currentDrafts) => {
      const nextDrafts: Record<number, Record<string, any>> = {};
      Object.entries(currentDrafts).forEach(([key, value]) => {
        const draftIndex = Number(key);
        if (!Number.isFinite(draftIndex) || draftIndex === removeIndex) {
          return;
        }

        nextDrafts[draftIndex > removeIndex ? draftIndex - 1 : draftIndex] = value;
      });
      return nextDrafts;
    });

    setMultiMetadataValidity((currentValidity) => {
      const nextValidity: Record<number, boolean> = {};
      Object.entries(currentValidity).forEach(([key, value]) => {
        const validityIndex = Number(key);
        if (!Number.isFinite(validityIndex) || validityIndex === removeIndex) {
          return;
        }

        nextValidity[validityIndex > removeIndex ? validityIndex - 1 : validityIndex] = value;
      });
      return nextValidity;
    });

    setActiveMetadataFileIndex((currentIndex) => {
      if (currentIndex === removeIndex) {
        return Math.max(0, removeIndex - 1);
      }

      return currentIndex > removeIndex ? currentIndex - 1 : currentIndex;
    });
  }, []);

  const handleMetadataDraftChange = React.useCallback((index: number, data: Record<string, any>, isValid: boolean): void => {
    setMultiMetadataDrafts((currentDrafts) => {
      const currentDraft = currentDrafts[index];
      if (JSON.stringify(currentDraft || {}) === JSON.stringify(data || {})) {
        return currentDrafts;
      }

      return {
        ...currentDrafts,
        [index]: data
      };
    });

    setMultiMetadataValidity((currentValidity) => {
      if (currentValidity[index] === isValid) {
        return currentValidity;
      }

      return {
        ...currentValidity,
        [index]: isValid
      };
    });
  }, []);

  const processingTitle = React.useMemo((): string => {
    if (processingPhase === 'analysis') {
      return mediaFile ? 'Analyzing Media...' : 'Analyzing Documents...';
    }

    if (processingPhase === 'upload') {
      return 'Uploading Files...';
    }

    return processingProgress || 'Processing...';
  }, [mediaFile, processingPhase, processingProgress]);

  const processingMessage = React.useMemo((): string => {
    if (processingPhase === 'analysis') {
      return 'Your file(s) is being processed by AI';
    }

    if (processingPhase === 'upload') {
      const fileCount = processingFileCount || uploadedFiles.length || filesData.length || 1;
      return `Your ${fileCount} file(s) is being processed to SharePoint...`;
    }

    return processingProgress || 'Your file(s) is being processed by AI';
  }, [filesData.length, processingFileCount, processingPhase, processingProgress, uploadedFiles.length]);

  React.useEffect(() => {
    if (props.projectId) {
      setCurrentProjectId(props.projectId);
    }
  }, [props.projectId]);

  React.useEffect(() => {
    if (filesData.length === 0) {
      setActiveMetadataFileIndex(0);
      setMultiMetadataDrafts({});
      setMultiMetadataValidity({});
      return;
    }

    setActiveMetadataFileIndex((currentIndex) => Math.min(currentIndex, filesData.length - 1));
  }, [filesData.length]);

  React.useEffect(() => {
    if (hasHydratedRestorePayloadRef.current) {
      return;
    }

    if (isReplaceMode) {
      clearUploadSuccessRestorePayload();
      hasHydratedRestorePayloadRef.current = true;
      return;
    }

    const restorePayload = readUploadSuccessRestorePayload();
    const expectedVariant: 'default' | 'kmArtifact' = isKmArtifactUpload ? 'kmArtifact' : 'default';
    if (!restorePayload || restorePayload.variant !== expectedVariant) {
      return;
    }

    hasHydratedRestorePayloadRef.current = true;
    setSuccessNotification(restorePayload.successNotification);
    setSuccessDocumentUrl(restorePayload.successDocumentUrl);
    setLastUploadedDocId(restorePayload.lastUploadedDocId);
    setSuccessDocumentLinks(restorePayload.successDocumentLinks);
    clearUploadSuccessRestorePayload();
  }, [isKmArtifactUpload, isReplaceMode]);

  const persistSuccessStateForReturn = React.useCallback((): void => {
    if (!successNotification || (successNotification.type !== 'success' && successNotification.type !== 'success-media')) {
      return;
    }

    const payload: IUploadSuccessRestorePayload = {
      variant: isKmArtifactUpload ? 'kmArtifact' : 'default',
      successNotification,
      successDocumentUrl,
      lastUploadedDocId,
      successDocumentLinks
    };

    try {
      sessionStorage.setItem(UPLOAD_SUCCESS_RESTORE_STORAGE_KEY, JSON.stringify(payload));
    } catch (error) {
      console.warn('Unable to persist upload success restore payload.', error);
    }
  }, [
    isKmArtifactUpload,
    lastUploadedDocId,
    successDocumentLinks,
    successDocumentUrl,
    successNotification
  ]);

  const loadTaxonomyOptions = React.useCallback(async (): Promise<ITaxonomyFieldOptions> => {
    if (taxonomyOptionsRef.current) {
      return taxonomyOptionsRef.current;
    }

    const options = await fetchAllTaxonomyOptions(props.context);
    taxonomyOptionsRef.current = options;
    setTaxonomyOptions(options);
    return options;
  }, [props.context]);

  React.useEffect(() => {
    configurePermalinkService(props.context.spHttpClient, props.context.pageContext.web.absoluteUrl);
  }, [props.context]);

  const storeUrlOnSubmit = React.useCallback(async (
    itemId: number,
    title: string | null | undefined,
    isCheckedOut: boolean = false
  ): Promise<string | null> => {
    const documentTitle = (title || `Document ${itemId}`).trim();

    // Generate URL immediately on submit regardless of status
    try {
      const storedUrl = buildAssetUrl(itemId);
      const siteUrl = props.context.pageContext.web.absoluteUrl;

      debugLog('Generating URL on submit for itemId:', itemId);
      debugLog('Generated URL:', storedUrl);

      // Store URL directly via ValidateUpdateListItem
      // bypassing the Active-only check in updateItemUrlField
      const response = await props.context.spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${LIBRARY_NAMES.kmDataHub}')/items(${itemId})/ValidateUpdateListItem`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
          },
          body: JSON.stringify({
            formValues: [
              {
                FieldName: COLUMN_NAMES.url,
                FieldValue: `${storedUrl}, ${documentTitle}`
              }
            ],
            bNewDocumentUpdate: !isCheckedOut
          })
        }
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`URL field update failed: ${response.status} ${errorText}`);
      }

      debugLog('URL stored successfully on submit:', storedUrl);

      // Set the success URL so the success message link works
      setSuccessDocumentUrl(buildAssetUrl(itemId));
      return storedUrl;
    } catch (urlError) {
      console.warn('URL generation on submit failed (non-blocking):', urlError);
      // Do NOT throw - URL failure should not block the success flow
      return null;
    }
  }, [props.context]);

  React.useEffect(() => {
    loadTaxonomyOptions().catch((error) => {
      console.error('Failed to load taxonomy options:', error);
      setProcessingError(error instanceof Error ? error.message : 'Failed to load taxonomy terms.');
    });

    if (currentProjectId) {
      validateProjectId(currentProjectId);
    }
  }, [loadTaxonomyOptions, currentProjectId]);

  React.useEffect(() => {
    const urlParams = new URLSearchParams(window.location.search);
    const action = urlParams.get('action');
    if (action === 'noContent' && currentProjectId && isProjectValid && projectStatus !== null) {
      debugLog(`[BOT] Auto-triggering "No Content" update for project: ${currentProjectId}`);
      // Remove the action param from URL so it doesn't re-trigger
      const newUrl = window.location.protocol + "//" + window.location.host + window.location.pathname + "?projectId=" + currentProjectId;
      window.history.replaceState({ path: newUrl }, '', newUrl);

      handleNoArtifacts();
    }
  }, [isProjectValid, currentProjectId, projectStatus, projectCount]);

  const fetchRecentlyUploadedDocs = async (pId: string) => {
    try {
      const fieldMap = await getKMDataHubFieldMap();
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const queryUrl = `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items?$select=ID,FileLeafRef,Created,FileRef&$filter=${fieldMap.projectId} eq '${pId}'&$orderby=Created desc&$top=10`;

      let items: IRecentUploadListItem[] = [];
      const response = await props.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1);
      if (response.ok) {
        const json = await response.json() as { value?: IRecentUploadListItem[] };
        items = Array.isArray(json.value) ? json.value : [];
      } else if (response.status === 500) {
        const siteOrigin = webUrl.split('/sites/')[0];
        const rootResp = await props.context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/RootFolder?$select=ServerRelativeUrl`,
          SPHttpClient.configurations.v1
        );
        if (rootResp.ok) {
          const rootData = await rootResp.json();
          const libraryPath = `${siteOrigin}${rootData.ServerRelativeUrl}`;
          const searchUrl =
            `${webUrl}/_api/search/query` +
            `?querytext='${encodeURIComponent(`Path:"${libraryPath}/*"`)}'` +
            `&selectproperties='ListItemID,Path,FileLeafRef,Title,Write'` +
            `&rowlimit=10` +
            `&sortlist='Write:descending'`;
          const searchResp = await props.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
          if (searchResp.ok) {
            const searchData = await searchResp.json();
            const rows = searchData?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];
            items = rows.map((row: any) => {
              const obj: Record<string, string> = {};
              (row.Cells || []).forEach((cell: any) => { obj[cell.Key] = cell.Value; });
              return {
                ID: Number(obj.ListItemID),
                FileLeafRef: obj.FileLeafRef || obj.Title || '',
                Created: obj.Write || new Date().toISOString(),
                FileRef: obj.Path || ''
              };
            }).filter((item: IRecentUploadListItem) => item.ID > 0);
          }
        }
      }

      if (items.length > 0) {
        const formattedDocs: ISessionUploadedDoc[] = items.map((item) => ({
          id: item.ID,
          name: item.FileLeafRef,
          time: new Date(item.Created).toLocaleDateString() + ' ' + new Date(item.Created).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          serverRelativeUrl: item.FileRef
        }));

        setSessionUploadedDocs((prevDocs) => {
          const sessionIds = new Set(prevDocs.map((doc) => doc.id));
          const filteredPrevious = formattedDocs.filter((doc) => !sessionIds.has(doc.id));
          return [...prevDocs, ...filteredPrevious];
        });
      }
    } catch (err) {
      console.error("[BOT] Error fetching recently uploaded docs:", err);
    }
  };

  const handleDeleteDoc = async (itemId: number, fileRef: string) => {
    try {
      setIsProcessing(true);
      setProcessingProgress("Deleting artifact from KM Data Hub...");
      const webUrl = props.context.pageContext.web.absoluteUrl;

      // 1. Delete the file from the library if we have the fileRef
      if (fileRef) {
        debugLog(`[BOT] Deleting file at: ${fileRef}`);
        await props.context.spHttpClient.post(
          `${webUrl}/_api/web/getFileByServerRelativeUrl('${fileRef}')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'X-HTTP-Method': 'DELETE',
              'IF-MATCH': '*'
            }
          }
        );
      } else {
        // Fallback to deleting the item by ID
        debugLog(`[BOT] Deleting list item by ID: ${itemId}`);
        await props.context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'X-HTTP-Method': 'DELETE',
              'IF-MATCH': '*'
            }
          }
        );
      }

      // 2. Update the count in Harvest Hub (decrement by 1)
      if (currentProjectId) {
        setProcessingProgress("Updating project artifacts count...");
        await updateHarvestTracker(currentProjectId, -1);
      }

      // 3. Update local UI state
      setSessionUploadedDocs(prev => prev.filter(d => d.id !== itemId));

      setIsProcessing(false);
      setProcessingProgress("");
      debugLog(`[BOT] Successfully deleted artifact ${itemId}`);

    } catch (err) {
      console.error("🔴 Deletion Error:", err);
      setIsProcessing(false);
      setProcessingProgress("");
      alert("Failed to delete artifact. It might have already been removed or you lack permissions.");
    }
  };

  const validateProjectId = async (pId: string) => {
    try {
      const cleanId = (pId || '').trim();
      debugLog(`[BOT] Validating Project ID: "${cleanId}"`);

      const webUrl = props.context.pageContext.web.absoluteUrl;

      // Discover Field Internal Names dynamically
      const fieldsResp = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('KM Harvest Hub')/fields?$select=InternalName,Title`,
        SPHttpClient.configurations.v1
      );
      const fields = (await fieldsResp.json()).value || [];

      const projectIdField = fields.find((f: any) =>
        ["ProjectId", "ProjectID", "Project_x0020_Id", "Project_x0020_ID"].indexOf(f.InternalName) !== -1 ||
        ["Project ID", "ProjectID", "Project Id"].indexOf(f.Title) !== -1
      )?.InternalName || "ProjectId";

      const projectNameField = fields.find((f: any) => f.InternalName === "ProjectName")?.InternalName ||
        fields.find((f: any) =>
          f.InternalName === "Project_x0020_Name" ||
          f.InternalName === "ProjectNameOWSTEXT" ||
          f.InternalName === "Title" ||
          f.Title === "Project Name" ||
          f.Title === "Project" ||
          f.Title === "Project Title" ||
          (f.Title && f.Title.toLowerCase().indexOf("project") !== -1 && f.Title.toLowerCase().indexOf("name") !== -1)
        )?.InternalName || "ProjectName";

      const buField = fields.find((f: any) =>
        f.InternalName === "BU" || f.Title === "BU" || f.Title === "Business Unit"
      )?.InternalName || "BU";

      const countField = fields.find((f: any) => 
          (f.InternalName && f.InternalName === "CountofDocuments") ||
          (f.InternalName && f.InternalName === "Count_x0020_of_x0020_Documents") ||
          (f.Title && f.Title === "Count of Documents") ||
          (f.Title && f.Title === "CountofDocuments")
        )?.InternalName || 
        fields.find((f: any) => 
          (f.Title && f.Title.toLowerCase().replace(/\s/g, '') === "countofdocuments") ||
          (f.InternalName && f.InternalName.toLowerCase().indexOf("count") !== -1)
        )?.InternalName || "CountofDocuments";

      debugLog(`[BOT] Discovery Results -> ID: ${projectIdField}, Name: ${projectNameField}, BU: ${buField}, Count: ${countField}`);
      debugLog(`[BOT] Available Fields:`, fields.map((f: any) => `${f.Title} (${f.InternalName})`).join(', '));

      // Temporary alert to verify discovery for the user
      // alert(`Discovered Fields: ID=${projectIdField}, Name=${projectNameField}, BU=${buField}`);

      const queryUrl = `${webUrl}/_api/web/lists/getbytitle('KM Harvest Hub')/items?$select=*,${projectNameField},${countField}&$filter=${projectIdField} eq '${cleanId}'`;
      debugLog(`[BOT] Validation Query: ${queryUrl}`);

      const response = await props.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        debugLog(`[BOT] Validation Response:`, data);
        if (data.value && data.value.length > 0) {
          const item = data.value[0];

          // DIAGNOSTIC POPUP: Show all columns to the user so we can find the name
          const allKeys = Object.keys(item).join(", ");
          // alert(`Found Project! Available Columns: ${allKeys}`);

          const statusField = fields.find((f: any) =>
            f.InternalName === "HarvestingStatus" ||
            f.InternalName === "Harvesting_x0020_Status" ||
            f.Title === "Harvesting Status" ||
            (f.Title && f.Title.toLowerCase().indexOf("harvesting") !== -1)
          )?.InternalName || fields.find((f: any) =>
            (f.Title && f.Title.toLowerCase().indexOf("status") !== -1 && f.Title !== "Title")
          )?.InternalName || "HarvestingStatus";

          const currentStatus = item[statusField] || "Draft";
          const currentCount = Number(item[countField]) || 0;
          setProjectName(item[projectNameField]?.Title || item[projectNameField] || item.Title || "Unnamed Project");
          setProjectBU(item[buField] || "Standard");
          setProjectStatus(currentStatus);
          setProjectCount(currentCount);
          setIsProjectValid(true);
          debugLog(`[BOT] Project ID "${cleanId}" is VALID. Status: ${currentStatus}, Count: ${currentCount}`);

          // Fetch existing docs for this project
          void fetchRecentlyUploadedDocs(cleanId);
        } else {
          console.warn(`[BOT] Project ID "${cleanId}" was NOT FOUND in KM Harvest Hub.`);
          setIsProjectValid(false);
          setProjectStatus('');
          showUploadValidationMessage(`Project ID ${cleanId} was not found in KM Harvest Hub.`);
        }
      } else {
        if (response.status === 500) {
          console.warn(
            'KM Harvest Hub threshold exceeded. ' +
            'Ask IT to add index on ProjectId field in KM Harvest Hub list settings.'
          );
          setProjectStatus('Draft');
          setProjectCount(0);
          setIsProjectValid(true);
          return;
        }
        const errText = await response.text();
        console.warn(`[BOT] Project ID validation request failed: ${errText}`);
        setProjectStatus('Draft');
        setProjectCount(0);
        setIsProjectValid(true);
      }
    } catch (err) {
      console.warn("[BOT] KM Harvest Hub lookup failed - continuing upload", err);
      setProjectStatus('Draft');
      setProjectCount(0);
      setIsProjectValid(true);
    }
  };

  const getKMDataHubFieldMap = React.useCallback(async (): Promise<IKMDataHubFieldMap> => {
    if (kmDataHubFieldMapRef.current) {
      return kmDataHubFieldMapRef.current;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const response = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/fields?$select=Title,InternalName,Hidden&$top=5000`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal'
        }
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to read KM Data Hub fields: ${response.status} ${errorText}`);
    }

    const json = await response.json();
    const fields = (json?.d?.results || json?.value || []) as Array<{
      Title?: string;
      InternalName?: string;
      Hidden?: boolean;
      Id?: string;
      TextField?: string;
      TypeAsString?: string;
    }>;
    const normalize = (value?: string): string => (value || '').trim().toLowerCase();
    const normalizeGuid = (value?: string): string => (value || '').replace(/[{}]/g, '').trim().toLowerCase();
    const findInternalName = (displayName: string, fallback: string): string => {
      const matchedField = fields.find((field) =>
        normalize(field.Title) === normalize(displayName) ||
        normalize(field.InternalName) === normalize(fallback)
      );

      return matchedField?.InternalName || fallback;
    };
    const findOptionalInternalName = (displayName: string, fallback: string): string | undefined => {
      const matchedField = fields.find((field) =>
        normalize(field.Title) === normalize(displayName) ||
        normalize(field.InternalName) === normalize(fallback)
      );

      return matchedField?.InternalName;
    };
    const findField = (internalName: string) =>
      fields.find((field) => normalize(field.InternalName) === normalize(internalName));
    const findTaxonomyNoteInternalName = (taxonomyInternalName: string): string => {
      const taxonomyField = findField(taxonomyInternalName);
      const textFieldId = normalizeGuid(taxonomyField?.TextField);
      if (!textFieldId) {
        return '';
      }

      const noteField = fields.find((field) => normalizeGuid(field.Id) === textFieldId);
      return noteField?.InternalName || '';
    };

    const fieldMap: IKMDataHubFieldMap = {
      status: findInternalName(COLUMN_NAMES.status, COLUMN_NAMES.status),
      published: findInternalName(COLUMN_NAMES.published, COLUMN_NAMES.published),
      title: findInternalName(COLUMN_NAMES.title, COLUMN_NAMES.title),
      description: findInternalName(COLUMN_NAMES.description, COLUMN_NAMES.description),
      author: findInternalName('Author', COLUMN_NAMES.author),
      documentType: findInternalName('Document Type', TAXONOMY_FIELD_CONFIGS.documentType.fieldInternalName),
      documentTypeNote: findTaxonomyNoteInternalName(TAXONOMY_FIELD_CONFIGS.documentType.fieldInternalName),
      client: findInternalName('Client', TAXONOMY_FIELD_CONFIGS.client.fieldInternalName),
      clientNote: findTaxonomyNoteInternalName(TAXONOMY_FIELD_CONFIGS.client.fieldInternalName),
      geography: findInternalName('Geography', TAXONOMY_FIELD_CONFIGS.geography.fieldInternalName),
      geographyNote: findTaxonomyNoteInternalName(TAXONOMY_FIELD_CONFIGS.geography.fieldInternalName),
      diseaseArea: findInternalName('Disease Area', TAXONOMY_FIELD_CONFIGS.diseaseArea.fieldInternalName),
      diseaseAreaNote: findTaxonomyNoteInternalName(TAXONOMY_FIELD_CONFIGS.diseaseArea.fieldInternalName),
      therapyArea: findInternalName('Therapy Area', TAXONOMY_FIELD_CONFIGS.therapyArea.fieldInternalName),
      therapyAreaNote: findTaxonomyNoteInternalName(TAXONOMY_FIELD_CONFIGS.therapyArea.fieldInternalName),
      sensitiveTerms: findInternalName('Sensitive Terms', COLUMN_NAMES.sensitiveTerms),
      views: findInternalName(COLUMN_NAMES.views, COLUMN_NAMES.views),
      likes: findInternalName(COLUMN_NAMES.likes, COLUMN_NAMES.likes),
      comments: findInternalName(COLUMN_NAMES.comments, COLUMN_NAMES.comments),
      downloads: findInternalName(COLUMN_NAMES.downloads, COLUMN_NAMES.downloads),
      follow: findInternalName(COLUMN_NAMES.follow, COLUMN_NAMES.follow),
      share: findInternalName(COLUMN_NAMES.share, COLUMN_NAMES.share),
      bookmark: findInternalName(COLUMN_NAMES.bookmark, COLUMN_NAMES.bookmark),
      projectId: findInternalName('Project ID', COLUMN_NAMES.projectId),
      edited: findInternalName('Edited', COLUMN_NAMES.edited),
      editedBy: findInternalName('Editor', COLUMN_NAMES.editedBy)
    };

    kmDataHubFieldMapRef.current = fieldMap;
    return fieldMap;
  }, [props.context]);

  const ensureSiteUser = React.useCallback(async (upn: string): Promise<number> => {
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const response = await props.context.spHttpClient.post(
      `${webUrl}/_api/web/ensureuser`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: JSON.stringify({
          logonName: `i:0#.f|membership|${upn}`
        })
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to ensure site user ${upn}: ${response.status} ${errorText}`);
    }

    const json = await response.json();
    const ensuredUserId = json?.d?.Id || json?.Id;

    if (!ensuredUserId) {
      throw new Error(`SharePoint did not return a user id for ${upn}.`);
    }

    return ensuredUserId;
  }, [props.context]);

  const buildKMReviewAuditFormValues = React.useCallback((): Array<{ FieldName: string; FieldValue: string }> => {
    const currentUserLogin =
      props.userEmail ||
      props.context?.pageContext.user.email ||
      props.context?.pageContext.user.loginName ||
      '';
    const formValues = [
      {
        FieldName: COLUMN_NAMES.edited,
        FieldValue: formatSharePointDateFieldValue()
      }
    ];

    if (currentUserLogin) {
      formValues.push({
        FieldName: COLUMN_NAMES.editedBy,
        FieldValue: buildClaimsValues([currentUserLogin])
      });
    }

    return formValues;
  }, [props.context, props.userEmail]);

  const appendKMReviewAuditFormValues = React.useCallback((
    formValues: Array<{ FieldName: string; FieldValue: string }>
  ): Array<{ FieldName: string; FieldValue: string }> => {
    const auditFormValues = buildKMReviewAuditFormValues();
    const auditFieldNames = new Set(auditFormValues.map((entry) => entry.FieldName.toLowerCase()));

    return [
      ...formValues.filter((entry) => !auditFieldNames.has(entry.FieldName.toLowerCase())),
      ...auditFormValues
    ];
  }, [buildKMReviewAuditFormValues]);

  const submitKMDataHubFormValues = React.useCallback(async (
    itemId: number,
    formValues: Array<{ FieldName: string; FieldValue: string }>,
    bNewDocumentUpdate: boolean,
    metadata?: Record<string, any>,
    skipCheckout?: boolean
  ): Promise<void> => {
    if (formValues.length === 0) {
      return;
    }

    const values = metadata || {};
    const formValuesWithAudit = appendKMReviewAuditFormValues(formValues);
    debugLog('BU terms:', values.selectedBUs);
    debugLog('Dept terms:', values.selectedDepts);
    debugLog('BU string:', (values.selectedBUs || []).map((t: any) => `${t.name}|${t.id}`).join(';'));
    debugLog('Dept string:', (values.selectedDepts || []).map((t: any) => `${t.name}|${t.id}`).join(';'));
    const selectedDeptIds = new Set(
      (values.selectedDepts || []).map((t: any) => t.id)
    );

    const leafDepts = (values.selectedDepts || []).filter((term: any) => {
      const hasSelectedChild = (term.children || []).some((child: any) =>
        selectedDeptIds.has(child.id)
      );
      return !hasSelectedChild;
    });

    debugLog('All selected depts:',
      (values.selectedDepts || []).map((t: any) => t.name));
    debugLog('Leaf depts only (what gets stored to SharePoint):',
      leafDepts.map((t: any) => t.name));
    debugLog('Dept FieldValue string:',
      leafDepts.map((t: any) => `${t.name}|${t.id}`).join(';'));

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const isExistingItem = typeof props.targetItemId === 'number' && props.targetItemId > 0;
    const effectiveBNewDocumentUpdate = (isExistingItem || skipCheckout) ? false : bNewDocumentUpdate;
    let response: SPHttpClientResponse | undefined;
    const submitMetadataUpdate = async (): Promise<void> => {
      response = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})/ValidateUpdateListItem`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: JSON.stringify({
            formValues: formValuesWithAudit,
            bNewDocumentUpdate: effectiveBNewDocumentUpdate
          })
        }
      );
    };

    if (isExistingItem && !skipCheckout) {
      await updateWithoutVersion(props.context.spHttpClient, webUrl, itemId, submitMetadataUpdate);
    } else {
      await submitMetadataUpdate();
    }

    if (!response || !response.ok) {
      const errorText = response ? await response.text() : '';
      throw new Error(`Failed to update metadata: ${response?.status || 'unknown'} ${errorText}`);
    }

    const responseJson = await response.json();
    debugLog('KM Data Hub ValidateUpdateListItem response:', responseJson);
    const fieldResults = responseJson?.value || responseJson?.d?.ValidateUpdateListItem?.results || responseJson?.d?.results || [];
    const fieldErrors = fieldResults.filter((entry: any) => entry.HasException);
    if (fieldErrors.length > 0) {
      console.error('KM Data Hub field validation errors:', fieldErrors);
      throw new Error(`KM Data Hub validation failed: ${fieldErrors.map((entry: any) => `${entry.FieldName}: ${entry.ErrorMessage}`).join('; ')}`);
    }
  }, [appendKMReviewAuditFormValues, props.context]);

  const fetchKMDataHubItemSnapshot = React.useCallback(async (
    itemId: number,
    fieldMap: IKMDataHubFieldMap
  ): Promise<IKMDataHubItemSnapshot> => {
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const itemResponse = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})?$select=Id,FileRef,${fieldMap.status},${fieldMap.published},${fieldMap.title},${fieldMap.description},${fieldMap.sensitiveTerms},${fieldMap.author}/Id,${fieldMap.author}/EMail&$expand=${fieldMap.author}`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal'
        }
      }
    );

    if (!itemResponse.ok) {
      const errorText = await itemResponse.text();
      throw new Error(`Failed to read KM Data Hub item ${itemId}: ${itemResponse.status} ${errorText}`);
    }

    const itemJson = await itemResponse.json();
    const item = itemJson?.d || itemJson || {};

    const textValuesResponse = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})/FieldValuesAsText`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal'
        }
      }
    );

    if (!textValuesResponse.ok) {
      const errorText = await textValuesResponse.text();
      throw new Error(`Failed to read KM Data Hub field values for item ${itemId}: ${textValuesResponse.status} ${errorText}`);
    }

    const textValuesJson = await textValuesResponse.json();
    const textValues = textValuesJson?.d || textValuesJson || {};

    return {
      itemId,
      fileRef: item.FileRef || '',
      item,
      textValues
    };
  }, [props.context]);

  const findExistingKMDataHubItemByFileName = React.useCallback(async (
    fileName: string,
    fieldMap: IKMDataHubFieldMap
  ): Promise<IKMDataHubItemSnapshot | null> => {
    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const siteOrigin = webUrl.split('/sites/')[0];
      const rootResp = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/RootFolder?$select=ServerRelativeUrl`,
        SPHttpClient.configurations.v1
      );
      if (!rootResp.ok) {
        return null;
      }

      const rootData = await rootResp.json();
      const libraryPath = `${siteOrigin}${rootData.ServerRelativeUrl}`;
      const escapedName = fileName.replace(/'/g, "''");
      const query = `FileLeafRef:"${escapedName}" AND Path:"${libraryPath}/*"`;
      const searchUrl =
        `${webUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(query)}'` +
        `&selectproperties='ListItemID,Path,Title'` +
        `&rowlimit=1`;

      const searchResp = await props.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
      if (!searchResp.ok) {
        return null;
      }

      const searchData = await searchResp.json();
      const rows = searchData?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];
      if (rows.length === 0) {
        return null;
      }

      const cells = rows[0].Cells || [];
      const obj: Record<string, string> = {};
      cells.forEach((cell: any) => { obj[cell.Key] = cell.Value; });
      const itemId = Number(obj.ListItemID);
      if (!itemId || itemId <= 0) {
        return null;
      }

      return fetchKMDataHubItemSnapshot(itemId, fieldMap);
    } catch {
      return null;
    }
  }, [fetchKMDataHubItemSnapshot, props.context]);

  const findKMDataHubDuplicateId = React.useCallback(async (
    managedProperty: 'Title' | 'FileLeafRef',
    value: string,
    excludeItemId?: number
  ): Promise<number | null> => {
    const normalizedValue = normalizeKqlPhraseValue(value);

    if (!normalizedValue) {
      return null;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const siteOrigin = webUrl.split('/sites/')[0];
    const rootResp = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/RootFolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1
    );

    if (!rootResp.ok) {
      return null;
    }

    const rootData = await rootResp.json();
    const serverRelativeUrl = rootData?.ServerRelativeUrl || rootData?.d?.ServerRelativeUrl || '';
    const libraryPath = `${siteOrigin}${serverRelativeUrl}`;
    const query = `${managedProperty}:"${normalizedValue}" AND Path:"${libraryPath}/*"`;
    const searchUrl =
      `${webUrl}/_api/search/query` +
      `?querytext='${encodeURIComponent(query)}'` +
      `&selectproperties='ListItemID,Title,FileLeafRef,Path'` +
      `&rowlimit=5`;

    const searchResp = await props.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);

    if (!searchResp.ok) {
      return null;
    }

    const searchData = await searchResp.json();
    const rows = searchData?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];

    for (const row of rows) {
      const cells = row.Cells || [];
      const result: Record<string, string> = {};
      cells.forEach((cell: any) => { result[cell.Key] = cell.Value; });

      const itemId = Number(result.ListItemID || 0);
      if (!itemId || itemId <= 0 || itemId === excludeItemId) {
        continue;
      }

      const resultValue = managedProperty === 'Title' ? result.Title : result.FileLeafRef;
      if ((resultValue || '').trim().toLowerCase() === normalizedValue.toLowerCase()) {
        return itemId;
      }
    }

    return null;
  }, [props.context]);

  const findKMDataHubDuplicateIdByTitle = React.useCallback(async (
    title: string,
    excludeItemId?: number
  ): Promise<number | null> => {
    const normalizedTitle = String(title || '').trim();

    if (!normalizedTitle) {
      return null;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const viewXml = [
      '<View Scope="RecursiveAll">',
      '<ViewFields><FieldRef Name="ID" /><FieldRef Name="Title" /></ViewFields>',
      '<Query><Where><Eq>',
      '<FieldRef Name="Title" />',
      `<Value Type="Text">${escapeXmlValue(normalizedTitle)}</Value>`,
      '</Eq></Where></Query>',
      '<RowLimit Paged="FALSE">5</RowLimit>',
      '</View>'
    ].join('');

    try {
      const response = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/RenderListDataAsStream`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata.metadata=minimal',
            'Content-Type': 'application/json;odata.metadata=minimal'
          },
          body: JSON.stringify({
            parameters: {
              ViewXml: viewXml,
              RenderOptions: 2
            }
          })
        }
      );

      if (response.ok) {
        const data = await response.json();
        const rows = data?.Row || data?.ListData?.Row || [];

        for (const row of rows) {
          const itemId = Number(row.ID || row.Id || 0);
          const rowTitle = String(row.Title || '').trim();

          if (
            itemId > 0 &&
            itemId !== excludeItemId &&
            rowTitle.toLowerCase() === normalizedTitle.toLowerCase()
          ) {
            return itemId;
          }
        }
      } else {
        console.warn('Immediate title duplicate lookup failed; falling back to search.', response.status);
      }
    } catch (error) {
      console.warn('Immediate title duplicate lookup failed; falling back to search.', error);
    }

    return findKMDataHubDuplicateId('Title', normalizedTitle, excludeItemId);
  }, [findKMDataHubDuplicateId, props.context]);

  const getKMDataHubRootServerRelativeUrl = React.useCallback(async (): Promise<string> => {
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const rootResp = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/RootFolder?$select=ServerRelativeUrl`,
      SPHttpClient.configurations.v1
    );

    if (!rootResp.ok) {
      return '';
    }

    const rootData = await rootResp.json();
    return rootData?.ServerRelativeUrl || rootData?.d?.ServerRelativeUrl || '';
  }, [props.context]);

  const findKMDataHubDuplicateIdByFileName = React.useCallback(async (
    fileName: string,
    excludeItemId?: number
  ): Promise<number | null> => {
    const rootServerRelativeUrl = await getKMDataHubRootServerRelativeUrl();

    if (!rootServerRelativeUrl) {
      return null;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const fileServerRelativeUrl = `${rootServerRelativeUrl.replace(/\/$/, '')}/${fileName}`;
    try {
      const response = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/GetFileByServerRelativeUrl('${escapeODataString(fileServerRelativeUrl)}')/ListItemAllFields?$select=Id`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata.metadata=minimal'
          }
        }
      );

      if (response.ok) {
        const data = await response.json();
        const itemId = Number(data?.Id || data?.d?.Id || 0);

        if (itemId > 0 && itemId !== excludeItemId) {
          return itemId;
        }
      }
    } catch (error) {
      console.warn('Immediate file name duplicate lookup failed; falling back to list data.', error);
    }

    const viewXml = [
      '<View Scope="RecursiveAll">',
      '<ViewFields><FieldRef Name="ID" /><FieldRef Name="FileLeafRef" /></ViewFields>',
      '<Query><Where><Eq>',
      '<FieldRef Name="FileLeafRef" />',
      `<Value Type="File">${escapeXmlValue(fileName)}</Value>`,
      '</Eq></Where></Query>',
      '<RowLimit Paged="FALSE">5</RowLimit>',
      '</View>'
    ].join('');

    try {
      const response = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/RenderListDataAsStream`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata.metadata=minimal',
            'Content-Type': 'application/json;odata.metadata=minimal'
          },
          body: JSON.stringify({
            parameters: {
              ViewXml: viewXml,
              RenderOptions: 2
            }
          })
        }
      );

      if (!response.ok) {
        return null;
      }

      const data = await response.json();
      const rows = data?.Row || data?.ListData?.Row || [];

      for (const row of rows) {
        const itemId = Number(row.ID || row.Id || 0);
        const rowFileName = String(row.FileLeafRef || '').trim();

        if (
          itemId > 0 &&
          itemId !== excludeItemId &&
          rowFileName.toLowerCase() === fileName.toLowerCase()
        ) {
          return itemId;
        }
      }
    } catch (error) {
      console.warn('List data file name duplicate lookup failed.', error);
    }

    return null;
  }, [getKMDataHubRootServerRelativeUrl, props.context]);

  const validateUniqueKMDataHubUpload = React.useCallback(async (
    file: File,
    title: string,
    fieldMap: IKMDataHubFieldMap,
    assetType: 'document' | 'media'
  ): Promise<void> => {
    if (isReplaceMode || props.targetItemId) {
      return;
    }

    const normalizedTitle = String(title || '').trim();
    const duplicateChecks: Array<Promise<number | null>> = [
      normalizedTitle
        ? findKMDataHubDuplicateIdByTitle(normalizedTitle)
        : Promise.resolve(null),
      findKMDataHubDuplicateIdByFileName(file.name)
    ];
    const [titleDuplicateId, fileNameDuplicateId] = await Promise.all(duplicateChecks);

    if (titleDuplicateId || fileNameDuplicateId) {
      throw new Error(getDuplicateUploadMessage(assetType, !!titleDuplicateId, !!fileNameDuplicateId));
    }
  }, [findKMDataHubDuplicateIdByFileName, findKMDataHubDuplicateIdByTitle, isReplaceMode, props.targetItemId]);

  const resolveTargetKMDataHubItem = React.useCallback(async (
    fileName: string,
    fieldMap: IKMDataHubFieldMap
  ): Promise<IKMDataHubItemSnapshot | null> => {
    if (typeof props.targetItemId === 'number' && props.targetItemId > 0) {
      return fetchKMDataHubItemSnapshot(props.targetItemId, fieldMap);
    }

    return findExistingKMDataHubItemByFileName(fileName, fieldMap);
  }, [fetchKMDataHubItemSnapshot, findExistingKMDataHubItemByFileName, props.targetItemId]);

  const uploadFileContentToTarget = React.useCallback(async (
    file: File,
    targetFileRef?: string
  ): Promise<string> => {
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const preferredFileRef = (targetFileRef || props.targetFileRef || '').trim();
    const targetFolder = escapeODataString(getFolderServerRelativeUrl(preferredFileRef));
    const targetFileName = escapeODataString(getLeafFileName(preferredFileRef, file.name));

    if (file.size <= 10485760) {
      const uploadResp = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${targetFolder}')/Files/add(url='${targetFileName}',overwrite=true)`,
        SPHttpClient.configurations.v1,
        { body: file }
      );

      if (!uploadResp.ok) {
        const txt = await uploadResp.text();
        throw new Error(`File upload failed: ${uploadResp.status} ${txt}`);
      }

      const uploadJson = await uploadResp.json();
      return uploadJson.ServerRelativeUrl;
    }

    const chunkSize = 10485760;
    const addResp = await props.context.spHttpClient.post(
      `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${targetFolder}')/Files/add(url='${targetFileName}',overwrite=true)`,
      SPHttpClient.configurations.v1,
      { body: ' ' }
    );
    if (!addResp.ok) {
      throw new Error('Chunk upload prep failed');
    }

    const serverRelativeUrl = (await addResp.json()).ServerRelativeUrl;
    const uploadId = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      const r = Math.random() * 16 | 0;
      const v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });

    let pointer = 0;
    let chunkIndex = 0;

    while (pointer < file.size) {
      const chunk = file.slice(pointer, pointer + chunkSize);
      const isFirst = chunkIndex === 0;
      const isLast = (pointer + chunkSize) >= file.size;

      let endpoint = '';
      if (isFirst) {
        endpoint = `${webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/startupload(uploadId=guid'${uploadId}')`;
      } else if (isLast) {
        endpoint = `${webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/finishupload(uploadId=guid'${uploadId}',fileOffset=${pointer})`;
      } else {
        endpoint = `${webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/continueupload(uploadId=guid'${uploadId}',fileOffset=${pointer})`;
      }

      const res = await props.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { body: chunk });
      if (!res.ok) {
        throw new Error(`Chunk upload failed at offset ${pointer}`);
      }

      pointer += chunkSize;
      chunkIndex++;
    }

    return serverRelativeUrl;
  }, [props.context, props.targetFileRef]);

  const getFileContentHash = React.useCallback(async (content: Blob | ArrayBuffer): Promise<string> => {
    const buffer = content instanceof ArrayBuffer ? content : await content.arrayBuffer();
    const digest = await crypto.subtle.digest('SHA-256', buffer);
    return Array.from(new Uint8Array(digest))
      .map((value) => value.toString(16).padStart(2, '0'))
      .join('');
  }, []);

  const hasDifferentFileContent = React.useCallback(async (
    file: File,
    fileRef: string
  ): Promise<boolean> => {
    if (!fileRef) {
      return true;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const response = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/GetFileByServerRelativeUrl('${escapeODataString(fileRef)}')/$value`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return true;
    }

    const [submittedHash, existingHash] = await Promise.all([
      getFileContentHash(file),
      response.arrayBuffer().then((buffer) => getFileContentHash(buffer))
    ]);

    return submittedHash !== existingHash;
  }, [getFileContentHash, props.context]);

  const buildKMDataHubFormValues = React.useCallback(async (
    metadata: Record<string, any>,
    fieldMap: IKMDataHubFieldMap,
    options?: {
      existingItem?: IKMDataHubItemSnapshot | null;
      initializeMetrics?: boolean;
      includeAllFields?: boolean;
    }
  ): Promise<Array<{ FieldName: string; FieldValue: string }>> => {
    const existingItem = options?.existingItem || null;
    const includeAllFields = options?.includeAllFields === true;
    const effectiveStatus = String(metadata.status || 'Under Review').trim() || 'Under Review';
    const isPublishedStatus = ['active', 'approved'].indexOf(effectiveStatus.toLowerCase()) !== -1;
    const effectiveTitle = metadata.title || '-';
    const effectiveDescription = cleanTextField(metadata.description || '-', SHAREPOINT_SHORT_TEXT_MAX_LENGTH) || '-';
    const effectiveSensitiveTerms = metadata.sensitiveTerms || '';
    const formValues: Array<{ FieldName: string; FieldValue: string }> = [];

    const pushIfChanged = (fieldName: string, fieldValue: string, currentValue?: string): void => {
      if (includeAllFields || normalizeComparisonValue(fieldValue) !== normalizeComparisonValue(currentValue)) {
        formValues.push({ FieldName: fieldName, FieldValue: fieldValue });
      }
    };

    const pushTaxonomyIfChanged = (fieldName: string, terms: ITaxonomyTerm[] | null | undefined, currentValue?: string): void => {
      const fieldValue = getMultiTaxonomyValue(terms);
      if (includeAllFields || normalizeTaxonomyTermsForComparison(terms) !== normalizeMultiValueComparison(currentValue)) {
        formValues.push({ FieldName: fieldName, FieldValue: fieldValue });
      }
    };

    pushIfChanged(fieldMap.status, effectiveStatus, existingItem?.textValues?.[fieldMap.status]);
    if (fieldMap.published && !isPublishedStatus) {
      pushIfChanged(fieldMap.published, '', existingItem?.textValues?.[fieldMap.published]);
    }
    pushIfChanged(fieldMap.title, effectiveTitle, existingItem?.textValues?.[fieldMap.title]);
    pushIfChanged(fieldMap.description, effectiveDescription, existingItem?.textValues?.[fieldMap.description]);
    pushIfChanged(fieldMap.sensitiveTerms, effectiveSensitiveTerms, existingItem?.textValues?.[fieldMap.sensitiveTerms]);

    if (props.projectId) {
      pushIfChanged(fieldMap.projectId, props.projectId, existingItem?.textValues?.[fieldMap.projectId]);
    }

    const normalizedSelectedAuthorUpns = Array.isArray(metadata.selectedAuthorUpns)
      ? metadata.selectedAuthorUpns.filter(Boolean)
      : metadata.selectedAuthorUpn
        ? [metadata.selectedAuthorUpn].filter(Boolean)
        : [];

    if (normalizedSelectedAuthorUpns.length > 0) {
      const currentAuthorEmails = Array.isArray(existingItem?.item?.[fieldMap.author])
        ? existingItem.item[fieldMap.author].map((author: any) => author?.EMail || '').filter(Boolean)
        : existingItem?.item?.[fieldMap.author]?.EMail
          ? [existingItem.item[fieldMap.author].EMail]
          : [];
      if (includeAllFields || normalizeMultiValueComparison(normalizedSelectedAuthorUpns.join(';')) !== normalizeMultiValueComparison(currentAuthorEmails.join(';'))) {
        formValues.push({
          FieldName: fieldMap.author,
          FieldValue: buildClaimsValues(normalizedSelectedAuthorUpns)
        });
      }
    }

    buildBuDepartmentFieldUpdates(
      metadata.selectedBUs || [],
      metadata.selectedDepts || [],
      { includeEmpty: includeAllFields }
    ).forEach((fieldUpdate) => {
      pushIfChanged(fieldUpdate.FieldName, fieldUpdate.FieldValue, existingItem?.textValues?.[fieldUpdate.FieldName]);
    });

    pushTaxonomyIfChanged(fieldMap.documentType, metadata.documentTypeTerms, existingItem?.textValues?.[fieldMap.documentType]);
    pushTaxonomyIfChanged(fieldMap.client, metadata.clientTerms, existingItem?.textValues?.[fieldMap.client]);
    pushTaxonomyIfChanged(fieldMap.geography, metadata.geographyTerms, existingItem?.textValues?.[fieldMap.geography]);
    pushTaxonomyIfChanged(fieldMap.diseaseArea, metadata.diseaseAreaTerms, existingItem?.textValues?.[fieldMap.diseaseArea]);
    pushTaxonomyIfChanged(fieldMap.therapyArea, metadata.therapyAreaTerms, existingItem?.textValues?.[fieldMap.therapyArea]);

    if (options?.initializeMetrics) {
      const metricValues = buildZeroKMDocumentMetrics();
      formValues.push({ FieldName: fieldMap.views, FieldValue: String(metricValues.views) });
      formValues.push({ FieldName: fieldMap.likes, FieldValue: String(metricValues.likes) });
      formValues.push({ FieldName: fieldMap.comments, FieldValue: String(metricValues.comments) });
      formValues.push({ FieldName: fieldMap.downloads, FieldValue: String(metricValues.downloads) });
      formValues.push({ FieldName: fieldMap.follow, FieldValue: String(metricValues.follow) });
      formValues.push({ FieldName: fieldMap.share, FieldValue: String(metricValues.share) });
      formValues.push({ FieldName: fieldMap.bookmark, FieldValue: String(metricValues.bookmark) });
    }

    return formValues;
  }, []);

  const updateKMDataHubTextFields = React.useCallback(async (
    itemId: number,
    metadata: Record<string, any>,
    fieldMap: IKMDataHubFieldMap
  ): Promise<void> => {
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const body: Record<string, any> = {
      [fieldMap.status]: metadata.status || 'Under Review',
      [fieldMap.title]: metadata.title || '-',
      [fieldMap.description]: cleanTextField(metadata.description || '-', SHAREPOINT_SHORT_TEXT_MAX_LENGTH) || '-',
      [fieldMap.sensitiveTerms]: metadata.sensitiveTerms || ''
    };

    if (props.projectId) {
      body[fieldMap.projectId] = props.projectId;
    }

    if (fieldMap.documentTypeNote && metadata.documentTypeTerms?.length) {
      body[fieldMap.documentTypeNote] = getMultiTaxonomyNoteValue(metadata.documentTypeTerms);
    }
    if (fieldMap.clientNote && metadata.clientTerms?.length) {
      body[fieldMap.clientNote] = getMultiTaxonomyNoteValue(metadata.clientTerms);
    }
    if (fieldMap.geographyNote && metadata.geographyTerms?.length) {
      body[fieldMap.geographyNote] = getMultiTaxonomyNoteValue(metadata.geographyTerms);
    }
    if (fieldMap.diseaseAreaNote && metadata.diseaseAreaTerms?.length) {
      body[fieldMap.diseaseAreaNote] = getMultiTaxonomyNoteValue(metadata.diseaseAreaTerms);
    }
    if (fieldMap.therapyAreaNote && metadata.therapyAreaTerms?.length) {
      body[fieldMap.therapyAreaNote] = getMultiTaxonomyNoteValue(metadata.therapyAreaTerms);
    }
    if (metadata.initializeMetrics === true) {
      const metricValues = buildZeroKMDocumentMetrics();
      body[fieldMap.views] = metricValues.views;
      body[fieldMap.likes] = metricValues.likes;
      body[fieldMap.comments] = metricValues.comments;
      body[fieldMap.downloads] = metricValues.downloads;
      body[fieldMap.follow] = metricValues.follow;
      body[fieldMap.share] = metricValues.share;
      body[fieldMap.bookmark] = metricValues.bookmark;
    }

    debugLog('KM Data Hub text and note field payload:', body);

    const response = await props.context.spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal',
          'Content-Type': 'application/json;odata.metadata=minimal',
          'X-HTTP-Method': 'MERGE',
          'IF-MATCH': '*',
          'odata-version': ''
        },
        body: JSON.stringify(body)
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to update KM Data Hub text fields: ${response.status} ${errorText}`);
    }
  }, [ensureSiteUser, props.context]);

  const logKMDataHubStoredValues = React.useCallback(async (
    itemId: number,
    fieldMap: IKMDataHubFieldMap
  ): Promise<void> => {
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const response = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})/FieldValuesAsText`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal'
        }
      }
    );

    if (!response.ok) {
      return;
    }

    const json = await response.json();
    const values = json?.d || json || {};
    debugLog('KM Data Hub stored field values:', {
      status: values[fieldMap.status],
      title: values[fieldMap.title],
      description: values[fieldMap.description],
      author: values[fieldMap.author],
      businessUnit: values[KM_DATA_HUB_BU_FIELD_INTERNAL_NAME],
      department: values[KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME],
      documentType: values[fieldMap.documentType],
      documentTypeNote: fieldMap.documentTypeNote ? values[fieldMap.documentTypeNote] : '',
      client: values[fieldMap.client],
      clientNote: fieldMap.clientNote ? values[fieldMap.clientNote] : '',
      geography: values[fieldMap.geography],
      geographyNote: fieldMap.geographyNote ? values[fieldMap.geographyNote] : '',
      diseaseArea: values[fieldMap.diseaseArea],
      diseaseAreaNote: fieldMap.diseaseAreaNote ? values[fieldMap.diseaseAreaNote] : '',
      therapyArea: values[fieldMap.therapyArea],
      therapyAreaNote: fieldMap.therapyAreaNote ? values[fieldMap.therapyAreaNote] : '',
      sensitiveTerms: values[fieldMap.sensitiveTerms]
    });
  }, [props.context]);

  const onBrowse = () => {
    fileInputRef.current?.click();
  };

  const onBrowseMedia = (e?: React.MouseEvent) => {
    e?.stopPropagation();
    mediaInputRef.current?.click();
  };

  const handleMediaUpload = async (f?: FileList | null) => {
    if (!f || f.length === 0) return;

    const filesArray = Array.from(f);
    const hasDocumentFile = filesArray.some((file) => isSupportedDocumentUploadFile(file));

    if (hasDocumentFile) {
      showUploadValidationMessage(MEDIA_UPLOAD_REJECTS_DOCUMENT_MESSAGE);
      resetFileInputs();
      return;
    }

    const hasUnsupportedFile = filesArray.some((file) => !isSupportedMediaUploadFile(file));

    if (hasUnsupportedFile) {
      showUploadValidationMessage(UNSUPPORTED_FILE_FORMAT_MESSAGE);
      resetFileInputs();
      return;
    }

    if (filesArray.length > MAX_MEDIA_FILES) {
      showUploadValidationMessage(getMediaUploadLimitMessage());
      resetFileInputs();
      return;
    }

    setMediaFile(filesArray[0]);
    resetFileInputs();
  };

  const onFileSelected = async (f?: FileList | null, allowMedia: boolean = false) => {
    if (!f || f.length === 0) return;

    const selectedFiles = Array.from(f);

    if (isReplaceMode) {
      if (f.length > 1) {
        alert('Please select only one file to replace the document.');
      }
      const file = selectedFiles[0];
      if (file && !isSupportedReplacementUploadFile(file)) {
        showUploadValidationMessage(UNSUPPORTED_FILE_FORMAT_MESSAGE);
        resetFileInputs();
        return;
      }
      if (file) {
        void handleReplaceFile(file);
      }
      return;
    }

    const hasMediaFile = selectedFiles.some((file) => isMediaUploadFile(file));
    const hasDocumentFile = selectedFiles.some((file) => isSupportedDocumentUploadFile(file));

    if (hasMediaFile && !allowMedia) {
      showUploadValidationMessage(DOCUMENT_UPLOAD_REJECTS_MEDIA_MESSAGE);
      resetFileInputs();
      return;
    }

    const hasUnsupportedFile = selectedFiles.some((file) =>
      allowMedia
        ? !isSupportedMediaUploadFile(file)
        : !isSupportedDocumentUploadFile(file)
    );

    if (hasUnsupportedFile) {
      showUploadValidationMessage(UNSUPPORTED_FILE_FORMAT_MESSAGE);
      resetFileInputs();
      return;
    }

    if (allowMedia && hasDocumentFile) {
      showUploadValidationMessage(MEDIA_UPLOAD_REJECTS_DOCUMENT_MESSAGE);
      resetFileInputs();
      return;
    }

    if (allowMedia && selectedFiles.length > MAX_MEDIA_FILES) {
      showUploadValidationMessage(getMediaUploadLimitMessage());
      resetFileInputs();
      return;
    }

    const filesArray = selectedFiles.slice(0, MAX_FILES);

    if (!allowMedia && selectedFiles.length > MAX_FILES) {
      showUploadValidationMessage(getDocumentUploadLimitMessage());
      resetFileInputs();
      return;
    }

    setUploadedFiles(filesArray);
    setProcessingFileCount(filesArray.length);
    setShowForm(false);
    setIsProcessing(true);
    setProcessingError(null);
    setProcessingPhase('analysis');
    setFilesData([]);
    setActiveMetadataFileIndex(0);
    setMultiMetadataDrafts({});
    setMultiMetadataValidity({});
    setProcessingProgress('');
    resetFileInputs();

    try {
      const currentTaxonomyOptions = await loadTaxonomyOptions();
      const spHttpClient = props.context.spHttpClient;
      const siteUrl = props.context.pageContext.web.absoluteUrl;
      const docTypeTerms = await fetchDocumentTypeTerms(spHttpClient, siteUrl);

      const flatDocTypeTerms: { name: string; description: string }[] = [];
      const flattenTerms = (terms: Term[]) => {
        terms.forEach((t) => {
          flatDocTypeTerms.push({
            name: t.name,
            description: t.description || ''
          });
          if (t.children?.length) flattenTerms(t.children);
        });
      };
      flattenTerms(docTypeTerms);

      const docTypePromptSection = flatDocTypeTerms
        .map((t) => (
          t.description
            ? `- "${t.name}": ${t.description}`
            : `- "${t.name}"`
        ))
        .join('\n');
      const clientTermsPromptSection = buildClientTermsPromptSection(currentTaxonomyOptions);

      debugLog('=== DOCUMENT TYPE TERMS FROM TERM STORE ===');
      debugLog('Total doc type terms fetched:', flatDocTypeTerms.length);
      debugLog('Terms with descriptions:', flatDocTypeTerms.filter(t => t.description).length);
      debugLog('Sample terms with descriptions:', JSON.stringify(flatDocTypeTerms.slice(0, 5)));
      debugLog('Full docTypePromptSection sent to AI (first 1000 chars):', docTypePromptSection.substring(0, 1000));
      debugLog('==========================================');
      debugLog('Doc type terms count for AI:', flatDocTypeTerms.length);

      // Process all files with AI in parallel
      setProcessingProgress(`Processing ${filesArray.length} file(s) with AI...`);

      const processingPromises = filesArray.map(async (file, index) => {
        debugLog(`Starting parallel processing for file ${index + 1}: ${file.name}`);
        const metadata = await processFileWithAI(file, docTypePromptSection, clientTermsPromptSection);
        const isMediaFile = file.type.startsWith("video/") || file.type.startsWith("audio/");
        const metadataForNormalization = isMediaFile
          ? {
              ...metadata,
              title: metadata?.title || '',
              description: cleanTextField(metadata?.description || '', SHAREPOINT_SHORT_TEXT_MAX_LENGTH),
              documentType: metadata?.documentType || '',
              client: metadata?.client || '',
              geography: metadata?.geography || '',
              therapyArea: metadata?.therapyArea || '',
              diseaseArea: metadata?.diseaseArea || '',
              bu: metadata?.bu || '',
              department: metadata?.department || '',
              subDepartment: '',
              sensitiveTerms: metadata?.sensitiveTerms || ''
            }
          : (metadata || {});

        if (isMediaFile) {
          debugLog('Object passed to normalizeMetadataForKMDataHub:', JSON.stringify(metadataForNormalization));
        }
        logTaxonomyAvailability(currentTaxonomyOptions);

        return {
          file: file,
          itemId: -1, // Placeholder - will be set when uploaded on submit
          metadata: normalizeMetadataForKMDataHub(metadataForNormalization, currentTaxonomyOptions)
        };
      });

      // Wait for all files to be processed
      const processedFilesData = await Promise.all(processingPromises);
      const fieldMap = await getKMDataHubFieldMap();
      const seenUploadTitles = new Set<string>();
      const seenUploadFileNames = new Set<string>();

      processedFilesData.forEach((fileData) => {
        const assetType = isSupportedMediaUploadFile(fileData.file) ? 'media' : 'document';
        const normalizedTitle = String(fileData.metadata?.title || fileData.file.name || '').trim().toLowerCase();
        const normalizedFileName = fileData.file.name.trim().toLowerCase();
        const hasTitleDuplicate = !!normalizedTitle && seenUploadTitles.has(normalizedTitle);
        const hasFileNameDuplicate = !!normalizedFileName && seenUploadFileNames.has(normalizedFileName);

        if (hasTitleDuplicate || hasFileNameDuplicate) {
          throw new Error(getDuplicateUploadMessage(assetType, hasTitleDuplicate, hasFileNameDuplicate));
        }

        if (normalizedTitle) {
          seenUploadTitles.add(normalizedTitle);
        }

        if (normalizedFileName) {
          seenUploadFileNames.add(normalizedFileName);
        }
      });

      await Promise.all(processedFilesData.map((fileData) =>
        validateUniqueKMDataHubUpload(
          fileData.file,
          fileData.metadata?.title || fileData.file.name,
          fieldMap,
          isSupportedMediaUploadFile(fileData.file) ? 'media' : 'document'
        )
      ));

      setFilesData(processedFilesData);
      setActiveMetadataFileIndex(0);
      setMultiMetadataDrafts({});
      setMultiMetadataValidity({});
      setIsProcessing(false);
      setProcessingPhase('idle');
      setProcessingProgress('');

      if (processedFilesData.length > 0) {
        setShowForm(true);
      } else {
        setProcessingError('No files were successfully processed.');
      }
    } catch (error) {
      console.error('Error processing files:', error);
      setUploadedFiles([]);
      setFilesData([]);
      setProcessingError(null);
      setIsProcessing(false);
      setProcessingPhase('idle');
      setProcessingProgress('');
      setProcessingFileCount(0);
      setSuccessNotification({
        message: error instanceof Error ? error.message : 'An error occurred while processing files.',
        type: 'error'
      });
    }
  };


  const processFileWithAI = async (
    file: File,
    docTypePromptSection?: string,
    clientTermsPromptSection?: string
  ): Promise<Record<string, any> | null> => {
    debugLog('=== STARTING FILE PROCESSING ===');
    debugLog('File:', file.name);
    debugLog('File type:', file.type);
    debugLog('File size:', file.size, 'bytes');

    const isMedia = file.type.startsWith("video/") || file.type.startsWith("audio/");

    try {
      if (isMedia) {
        debugLog('Step 1: Transcribing media...');
        setProcessingProgress(`Transcribing ${file.name}...`);
        const { transcriptText } = await openAIService.current.transcribeMedia(file);

        let mediaDocTypeTermsSection = '';
        const flatMediaDocTypeTerms: { name: string; description: string }[] = [];
        try {
          const mediaDocTypeTerms = await fetchDocumentTypeTerms(
            props.context.spHttpClient,
            props.context.pageContext.web.absoluteUrl
          );
          const flattenMediaTerms = (terms: Term[]) => {
            terms.forEach((t) => {
              flatMediaDocTypeTerms.push({
                name: t.name,
                description: t.description || ''
              });
              if (t.children?.length) flattenMediaTerms(t.children);
            });
          };
          flattenMediaTerms(mediaDocTypeTerms);
          mediaDocTypeTermsSection = flatMediaDocTypeTerms
            .map((t) => (
              t.description
                ? `- "${t.name}": ${t.description}`
                : `- "${t.name}"`
            ))
            .join('\n');
        } catch (e) {
          console.warn('Could not fetch doc type terms for media:', e);
          mediaDocTypeTermsSection = docTypePromptSection || '';
        }

        debugLog('=== MEDIA DOCUMENT TYPE TERMS FROM TERM STORE ===');
        debugLog('Total media doc type terms fetched:', flatMediaDocTypeTerms.length);
        debugLog('Terms with descriptions:', flatMediaDocTypeTerms.filter(t => t.description).length);
        debugLog('Sample terms with descriptions:', JSON.stringify(flatMediaDocTypeTerms.slice(0, 5)));
        debugLog('Full media docTypePromptSection (first 1000 chars):', mediaDocTypeTermsSection.substring(0, 1000));
        debugLog('=================================================');

        debugLog('Step 2: Generating abstract and metadata from transcript...');
        setProcessingProgress(`Analyzing media content: ${file.name}...`);
        const mediaAnalysis = await openAIService.current.generateAbstractFromTranscript(
          file.name,
          transcriptText,
          mediaDocTypeTermsSection,
          clientTermsPromptSection
        );
        const cleanedMediaTitle = cleanTextField(
          (mediaAnalysis.title || file.name)
            .replace(/\.(mp4|mp3|wav|mov|avi|mkv|webm|m4a|aac|flac|ogg)$/gi, '')
            .replace(/[_-]/g, ' ')
            .trim(),
          255
        );
        const mediaMetadataForNormalization = {
          title: cleanedMediaTitle,
          description: cleanTextField(mediaAnalysis.abstract || '', SHAREPOINT_SHORT_TEXT_MAX_LENGTH),
          documentType: mediaAnalysis.documentType || '',
          client: mediaAnalysis.client || '',
          geography: mediaAnalysis.geography || '',
          therapyArea: mediaAnalysis.therapyArea || '',
          diseaseArea: mediaAnalysis.diseaseArea || '',
          bu: mediaAnalysis.businessUnit || '',
          department: mediaAnalysis.department || '',
          subDepartment: '',
          sensitiveTerms: mediaAnalysis.sensitiveTerms || ''
        };

        debugLog('Batch media raw result:', JSON.stringify(mediaAnalysis));
        debugLog('Batch therapyArea:', mediaAnalysis.therapyArea);
        debugLog('Media AI raw result:', mediaAnalysis);
        debugLog('Media cleaned title:', cleanedMediaTitle);
        debugLog('Object passed to normalizeMetadataForKMDataHub:', JSON.stringify(mediaMetadataForNormalization));

        debugLog('=== MEDIA PROCESSING COMPLETE ===');
        return mediaMetadataForNormalization;
      }

      // Step 1: Parse the document to extract text
      debugLog('Step 1: Parsing document...');
      const parseResult = await DocumentParser.parseFile(file);

      debugLog('Parse result:', {
        success: parseResult.success,
        textLength: parseResult.text?.length || 0,
        error: parseResult.error
      });

      if (!parseResult.success) {
        const errorMsg = parseResult.error || 'Failed to parse document';
        console.error('Document parsing failed:', errorMsg);
        return null;
      }

      if (!parseResult.text || parseResult.text.trim().length === 0) {
        console.warn('No text content found in the document. The document might be image-based or empty.');
        return null;
      }

      // Log parsed text for debugging
      debugLog('=== FILE PARSED SUCCESSFULLY ===');
      debugLog('File name:', file.name);
      debugLog('File size:', file.size, 'bytes');
      debugLog('Extracted text length:', parseResult.text.length, 'characters');
      debugLog('Sample text (first 500 chars):', parseResult.text.substring(0, 500));

      // Step 2: Extract metadata using Azure OpenAI
      debugLog('Step 2: Extracting metadata with AI...');
      const metadata = await openAIService.current.extractMetadata(
        parseResult.text,
        docTypePromptSection,
        clientTermsPromptSection
      );

      debugLog('=== FILE PROCESSING COMPLETE ===');
      return metadata;
    } catch (error) {
      console.error('=== ERROR PROCESSING FILE ===');
      console.error('Error type:', typeof error);
      console.error('Error details:', error);
      console.error('Error message:', error instanceof Error ? error.message : String(error));
      console.error('Error stack:', error instanceof Error ? error.stack : 'N/A');

      return null;
    }
  };

  /**
   * Create an Audit Log item in SharePoint list.
   * Uses the provided list GUID and the internal field names seen in the URLs.
   */
  const createAuditLogItem = async (file: File, action: string, documentTitle?: string | null): Promise<number> => {

    if (!props.context) {
      throw new Error('SPFx context not provided to FileUpload component.');
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;

    // Get current user id
    const currentUserResp = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/currentuser`,
      SPHttpClient.configurations.v1
    );
    if (!currentUserResp.ok) {
      const txt = await currentUserResp.text();
      throw new Error(`Failed to get current user: ${currentUserResp.status} ${txt}`);
    }
    const currentUser = await currentUserResp.json();
    const userId = currentUser.Id;

    // Build payload using the exact internal field names from the Audit Log list.
    const auditTitle = (documentTitle || '').trim() || file.name;

    // - Title      = KM Review Hub title (the 'Title' column always exists)
    // - FileName   = file name (Audit Log "File Name" column internal name)
    // - Action     = choice field (Draft / Under Review / Published)
    // - PerformedById = person lookup (Id suffix = SharePoint user id)
    // - TimeStamp  = date/time field
    // NOTE: Do NOT include 'UserId' or 'FileName' – those columns don't exist
    //       and cause a 400 error that silently blocks the whole item creation.
    const body: any = {
      "__metadata": { "type": "SP.Data.Audit_x0020_LogListItem" },
      Title: auditTitle,
      FileName: file.name,
      Action: action,
      PerformedById: userId,
      TimeStamp: new Date().toISOString()
    };

    debugLog('Creating audit log item with payload:', body);

    const postResp = await props.context.spHttpClient.post(
      `${webUrl}/_api/web/lists/GetByTitle('${LIST_NAMES.auditLog}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Accept": "application/json;odata=verbose",
          "Content-Type": "application/json;odata=verbose",
          "odata-version": ""
        },
        body: JSON.stringify(body)
      }
    );

    const respText = await postResp.text();
    if (!postResp.ok) {
      let spErrorMsg = respText;
      try {
        const errorJson = JSON.parse(respText);
        if (errorJson?.error?.message?.value) {
          spErrorMsg = errorJson.error.message.value;
        }
      } catch (e) {
        // Not JSON
      }
      console.error('Create audit item failed. Status:', postResp.status, 'Error Details:', spErrorMsg);
      console.error('Payload sent was:', JSON.stringify(body, null, 2));
      throw new Error(`Failed to create audit item: ${postResp.status} ${spErrorMsg}`);
    }

    // If we get here, the server created the item. Try parsing JSON if returned.
    let created: any = null;
    try {
      created = respText ? JSON.parse(respText) : null;
    } catch (e) {
      // If the response isn't JSON (depending on OData settings), log raw text
      console.warn('Could not parse create response as JSON; raw response:', respText);
    }

    const createdId = created && created.Id ? created.Id : null;
    debugLog('Audit log item created. ID (if returned):', createdId, 'rawResponse:', respText);
    return createdId;
  };

  /**
   * Upload file to SharePoint with KM Data Hub metadata
   */
  const uploadFileToSharePoint = async (file: File, formData: any): Promise<{itemId: number; isNew: boolean}> => {
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const allTaxonomyOptions = await loadTaxonomyOptions();
    logTaxonomyAvailability(allTaxonomyOptions);
    const normalizedFormData = normalizeMetadataForKMDataHub(formData, allTaxonomyOptions);
    const fieldMap = await getKMDataHubFieldMap();
    await validateUniqueKMDataHubUpload(file, normalizedFormData.title || file.name, fieldMap, 'document');
    const existingItem = await resolveTargetKMDataHubItem(file.name, fieldMap);

    if (existingItem) {
      const [fileChanged, changedFormValues] = await Promise.all([
        hasDifferentFileContent(file, existingItem.fileRef),
        buildKMDataHubFormValues(normalizedFormData, fieldMap, { existingItem })
      ]);

      if (!fileChanged && changedFormValues.length === 0) {
        debugLog(`Skipping SharePoint update for ${file.name}; file content and metadata are unchanged.`);
        await updateWithoutVersion(
          props.context.spHttpClient,
          webUrl,
          existingItem.itemId,
          async () => {
            await storeUrlOnSubmit(existingItem.itemId, normalizedFormData.title || file.name, true);
          }
        );
        return { itemId: existingItem.itemId, isNew: false };
      }

      let itemId = existingItem.itemId;
      const effectiveItemId = props.targetItemId || existingItem.itemId;
      if (fileChanged) {
        let serverRelativeUrl = '';
        const uploadExistingFile = async (): Promise<void> => {
          serverRelativeUrl = await uploadFileContentToTarget(file, existingItem.fileRef);
        };

        // Use props.targetItemId if provided (DocumentDetailPage update)
        // OR use existingItem.itemId if found by filename (kmArtifact upload).
        // Both cases need proper versioning.
        if (effectiveItemId) {
          await uploadNewFileVersion(
            props.context.spHttpClient,
            webUrl,
            effectiveItemId,
            uploadExistingFile
          );
        } else {
          await uploadExistingFile();
        }

        await new Promise(r => setTimeout(r, 500));

        const listItemResp = await props.context.spHttpClient.get(
          `${webUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')/ListItemAllFields?$select=Id`,
          SPHttpClient.configurations.v1
        );

        if (!listItemResp.ok) {
          const txt = await listItemResp.text();
          throw new Error(`Failed to get list item: ${listItemResp.status} ${txt}`);
        }

        itemId = (await listItemResp.json()).Id;
      }

      const existingItemFormValues = fileChanged
        ? [
          ...changedFormValues,
          {
            FieldName: COLUMN_NAMES.versionFileName,
            FieldValue: file.name
          },
          {
            FieldName: COLUMN_NAMES.versionFileType,
            FieldValue: file.name.split('.').pop() || ''
          }
        ]
        : changedFormValues;

      await updateWithoutVersion(
        props.context.spHttpClient,
        webUrl,
        effectiveItemId,
        async () => {
          await storeUrlOnSubmit(itemId, normalizedFormData.title || file.name, true);

          if (existingItemFormValues.length > 0) {
            debugLog('KM Data Hub changed form values:', existingItemFormValues);
            await submitKMDataHubFormValues(itemId, existingItemFormValues, true, normalizedFormData, true);
          }
        }
      );

      await logKMDataHubStoredValues(itemId, fieldMap);
      emitDocumentDataChanged({
        documentIds: [itemId],
        reason: fileChanged ? 'upload' : 'metadata'
      });
      debugLog('Existing KM Data Hub item updated only where changes were detected. itemId:', itemId);
      return { itemId, isNew: false };
    }

    let serverRelativeUrl = "";
    if (file.size <= 10485760) {
      const uploadResp = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${LIBRARY_NAME}')/Files/add(url='${file.name}',overwrite=true)`,
        SPHttpClient.configurations.v1,
        { body: file }
      );

      if (!uploadResp.ok) {
        const txt = await uploadResp.text();
        throw new Error(`File upload failed: ${uploadResp.status} ${txt}`);
      }

      const uploadJson = await uploadResp.json();
      serverRelativeUrl = uploadJson.ServerRelativeUrl;
    } else {
      const chunkSize = 10485760; // 10MB
      const addResp = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${LIBRARY_NAME}')/Files/add(url='${file.name}',overwrite=true)`,
        SPHttpClient.configurations.v1,
        { body: " " }
      );
      if (!addResp.ok) throw new Error("Chunk upload prep failed");
      serverRelativeUrl = (await addResp.json()).ServerRelativeUrl;

      const uploadId = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
      });

      let pointer = 0;
      let chunkIndex = 0;

      while (pointer < file.size) {
        const chunk = file.slice(pointer, pointer + chunkSize);
        const isFirst = chunkIndex === 0;
        const isLast = (pointer + chunkSize) >= file.size;

        let endpoint = "";
        if (isFirst) {
          endpoint = `${webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/startupload(uploadId=guid'${uploadId}')`;
        } else if (isLast) {
          endpoint = `${webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/finishupload(uploadId=guid'${uploadId}',fileOffset=${pointer})`;
        } else {
          endpoint = `${webUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/continueupload(uploadId=guid'${uploadId}',fileOffset=${pointer})`;
        }

        const res = await props.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, { body: chunk });
        if (!res.ok) throw new Error(`Chunk upload failed at offset ${pointer}`);

        pointer += chunkSize;
        chunkIndex++;
      }
    }

    await new Promise(r => setTimeout(r, 500));

    const listItemResp = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')/ListItemAllFields?$select=Id`,
      SPHttpClient.configurations.v1
    );

    if (!listItemResp.ok) {
      const txt = await listItemResp.text();
      throw new Error(`Failed to get list item: ${listItemResp.status} ${txt}`);
    }

    const itemId = (await listItemResp.json()).Id;
    const newFileRef = serverRelativeUrl;

    const initialFormValues = await buildKMDataHubFormValues(
      {
        ...normalizedFormData,
        title: normalizedFormData.title || file.name
      },
      fieldMap,
      {
        includeAllFields: true,
        initializeMetrics: true
      }
    );
    const versionFileFields = [
      {
        FieldName: COLUMN_NAMES.versionFileName,
        FieldValue: file.name
      },
      {
        FieldName: COLUMN_NAMES.versionFileType,
        FieldValue: file.name.split('.').pop() || ''
      },
      {
        FieldName: COLUMN_NAMES.contentRefreshDate,
        FieldValue: formatSharePointDateFieldValue()
      }
    ];
    const initialFormValuesWithVersionFields = [
      ...initialFormValues,
      ...versionFileFields
    ];

    debugLog('KM Data Hub initial form values:', initialFormValuesWithVersionFields);
    await saveMetadataAndCheckin(
      props.context.spHttpClient,
      webUrl,
      newFileRef,
      async () => {
        await storeUrlOnSubmit(itemId, normalizedFormData.title || file.name, true);
        await submitKMDataHubFormValues(itemId, initialFormValuesWithVersionFields, true, {
          ...normalizedFormData,
          title: normalizedFormData.title || file.name
        });
      }
    );
    await logKMDataHubStoredValues(itemId, fieldMap);
    emitDocumentDataChanged({
      documentIds: [itemId],
      reason: 'upload'
    });
    debugLog("File uploaded to SharePoint KM Data Hub with Status='Under Review', itemId:", itemId);
    return { itemId, isNew: true };
  };
  const updateKMDataHubWithFormData = async (itemId: number, data: any) => {
    const allTaxonomyOptions = await loadTaxonomyOptions();
    logTaxonomyAvailability(allTaxonomyOptions);
    const normalizedData = normalizeMetadataForKMDataHub(data, allTaxonomyOptions);

    debugLog("=== DEBUG: updateKMDataHubWithFormData ===");
    debugLog("ItemId:", itemId);
    debugLog("Form data received:", JSON.stringify(normalizedData, null, 2));

    const fieldMap = await getKMDataHubFieldMap();
    const existingItem = await fetchKMDataHubItemSnapshot(itemId, fieldMap);
    const formValues = await buildKMDataHubFormValues(normalizedData, fieldMap, { existingItem });

    if (formValues.length > 0) {
      debugLog('KM Data Hub changed form values:', formValues);
      await submitKMDataHubFormValues(itemId, formValues, true, normalizedData);
      emitDocumentDataChanged({
        documentIds: [itemId],
        reason: 'metadata'
      });
    } else {
      debugLog(`Skipping KM Data Hub metadata update for item ${itemId}; no field changes detected.`);
    }

    debugLog("=== DEBUG: Update successful ===");
    debugLog("KM Data Hub updated with form data and Status='Under Review', itemId:", itemId);
    debugLog("SensitiveTerms value that was sent:", normalizedData.sensitiveTerms || "");
    await logKMDataHubStoredValues(itemId, fieldMap);
  };


  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length) {
      onFileSelected(e.dataTransfer.files);
    }
  };

  const onFormSubmit = async (allFilesData: FileData[]) => {
    debugLog("onFormSubmit triggered with", allFilesData);

    try {
      if (!allFilesData || allFilesData.length === 0) {
        throw new Error("No file data found for submission.");
      }

      isSubmittingRef.current = true;
      setSuccessDocumentLinks([]);
      setSuccessDocumentUrl(null);
      setLastUploadedDocId(null);
      setIsProcessing(true);
      setProcessingError(null);
      setProcessingPhase('upload');
      setProcessingFileCount(allFilesData.length);
      setProcessingProgress(`Your ${allFilesData.length} file(s) is being processed to SharePoint...`);

      // Upload files to SharePoint with KM Data Hub metadata
      setProcessingProgress(`Your ${allFilesData.length} file(s) is being processed to SharePoint...`);
      const allTaxonomyOptions = await loadTaxonomyOptions();

      const requiresAuthorLookup = allFilesData.some((fileData) => {
        logTaxonomyAvailability(allTaxonomyOptions);
        const normalizedMetadata = normalizeMetadataForKMDataHub(fileData.metadata, allTaxonomyOptions);
        return Array.isArray(normalizedMetadata.selectedAuthorUpns) && normalizedMetadata.selectedAuthorUpns.length > 0;
      });
      const kmsUsers = requiresAuthorLookup ? await getKmsUsers(props.context) : [];

      // Process all files in parallel - upload and set metadata in one go
      const uploadPromises = allFilesData.map(async (fileData) => {
        try {
          debugLog(`Uploading file ${fileData.file.name} to SharePoint...`);

          logTaxonomyAvailability(allTaxonomyOptions);
          const normalizedMetadata = normalizeMetadataForKMDataHub(fileData.metadata, allTaxonomyOptions);

          // 1. Upload file to SharePoint with KM Data Hub metadata
          const uploadResult = await uploadFileToSharePoint(fileData.file, normalizedMetadata);
          const itemId = uploadResult.itemId;
          const isNewItem = uploadResult.isNew;

          if (!itemId || itemId === -1) {
            throw new Error('Invalid item ID. File may not have been uploaded correctly.');
          }

          // 2. Create audit log item
          try {
            await createAuditLogItem(fileData.file, "Under Review", normalizedMetadata.title || fileData.file.name);
          } catch (err) {
            console.warn(`Audit log creation failed for ${fileData.file.name}:`, err);
            // Continue even if audit log fails
          }

          debugLog(`File uploaded successfully: ${fileData.file.name} (itemId: ${itemId})`);

          return {
            success: true,
            fileName: fileData.file.name,
            newItemId: itemId,
            isNew: isNewItem,
            submittedTitle: normalizedMetadata.title || fileData.file.name
          };
        } catch (err) {
          return {
            success: false,
            fileName: fileData.file.name,
            error: err instanceof Error ? err.message : String(err)
          };
        }
      });

      // Wait for all uploads to complete
      const results = await Promise.all(uploadPromises);
      const successfulUploads = results.filter(r => r.success);
      const successfulItemId = successfulUploads.length > 0 ? (successfulUploads[0] as any).newItemId : null;

      if (successfulUploads.length === 0) {
        const firstError = results.find(r => !r.success && r.error)?.error || "Unknown error";
        throw new Error(`Upload failed: ${firstError}`);
      }

      if (knowledgeSearchApiClient.isConfigured()) {
        // TODO: Restore Easy Auth/token enforcement when backend sync auth is enabled.
        knowledgeSearchApiClient
          .triggerSync('frontend_upload_or_metadata_update')
          .catch((searchErr) => console.warn('Search backend sync trigger skipped/failed after upload:', searchErr));
      }

      // Store in session tracker
      const newUploads = successfulUploads.map(r => ({
        id: (r as any).newItemId || -1,
        name: r.fileName,
        time: new Date().toLocaleTimeString()
      }));

      setFilesData([]); // Clear the form data
      setUploadedFiles([]);
      setProcessingFileCount(0);
      setShowForm(false);
      setSessionUploadedDocs(prev => {
        // Deduplicate: remove old entry with same id so re-uploads move to top without duplicating
        const prevFiltered = prev.filter(p => !newUploads.some(n => n.id === p.id));
        return [...newUploads, ...prevFiltered];
      });

      if (successfulItemId) {
        const uploadedDocumentLinks = successfulUploads
          .map((uploadResult: any) => {
            const itemId = Number(uploadResult?.newItemId);
            if (!Number.isFinite(itemId) || itemId <= 0) {
              return null;
            }

            const title = uploadResult?.submittedTitle || uploadResult?.fileName;
            return {
              id: String(itemId),
              url: buildAssetUrl(itemId),
              title
            };
          })
          .filter((link): link is ISuccessDocumentLink => !!link);
        const successUrl = buildAssetUrl(Number(successfulItemId));

        setLastUploadedDocId(String(successfulItemId));
        setSuccessDocumentUrl(successUrl);
        setSuccessDocumentLinks(uploadedDocumentLinks);
        debugLog('Success URL set:', successUrl);
        debugLog('Success document links:', uploadedDocumentLinks);
        if (allFilesData[0]?.file) {
          props.onUploaded?.(allFilesData[0].file);
        }
      } else {
        setLastUploadedDocId(null);
        setSuccessDocumentUrl(null);
        setSuccessDocumentLinks([]);
      }

      if (props.projectId && successfulItemId) {
        try {
          const isAnyMedia = allFilesData.some(f => f.file.type.startsWith('video/') || f.file.type.startsWith('audio/'));
          // Only count genuinely new uploads — not overwrites of existing documents
          const newUploadsCount = successfulUploads.filter(r => (r as any).isNew === true).length;
          await updateHarvestTracker(props.projectId, newUploadsCount, 'Submitted', isAnyMedia);
        } catch (e) {
          console.warn("Tracker update failed:", e);
        }

        setSuccessNotification({
          message: 'Your contribution has been submitted.',
          type: 'success'
        });
        setIsProcessing(false);
        setProcessingPhase('idle');
        setProcessingFileCount(0);
        isSubmittingRef.current = false;
      } else {
        setSuccessNotification({
          message: 'Your contribution has been submitted.',
          type: 'success'
        });
        setProcessingError(null);
      setProcessingProgress('');
      setProcessingFileCount(0);
      isSubmittingRef.current = false;
        setIsProcessing(false);
        setProcessingPhase('idle');
      }
    } catch (err) {
      console.error("Error during form submission:", err);
      setProcessingError(err instanceof Error ? err.message : 'Error submitting files');
      isSubmittingRef.current = false;
      setIsProcessing(false);
      setProcessingPhase('idle');
      setProcessingFileCount(0);
      alert("Submission Error: " + (err instanceof Error ? err.message : String(err)));
    }
  };


  const updateHarvestTracker = async (pId: string, countIncrement: number = 1, newStatus: string = 'Submitted', isMedia: boolean = false) => {
    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;

      // 0. Discover field names for Harvest Hub (Smart Discovery)
      const fieldsResp = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('KM Harvest Hub')/fields?$select=InternalName,Title`,
        SPHttpClient.configurations.v1
      );
      const fields = (await fieldsResp.json()).value || [];

      const projectIdField = fields.find((f: any) =>
        ["ProjectId", "ProjectID", "Project_x0020_Id", "Project_x0020_ID"].indexOf(f.InternalName) !== -1 ||
        ["Project ID", "ProjectID"].indexOf(f.Title) !== -1
      )?.InternalName || "ProjectId";

      const countField = fields.find((f: any) => 
          (f.InternalName && f.InternalName === "CountofDocuments") ||
          (f.InternalName && f.InternalName === "Count_x0020_of_x0020_Documents") ||
          (f.Title && f.Title === "Count of Documents") ||
          (f.Title && f.Title === "CountofDocuments")
        )?.InternalName || 
        fields.find((f: any) => 
          (f.Title && f.Title.toLowerCase().replace(/\s/g, '') === "countofdocuments") ||
          (f.InternalName && f.InternalName.toLowerCase().indexOf("count") !== -1)
        )?.InternalName || "CountofDocuments";

      const statusField = fields.find((f: any) => 
        f.InternalName === "HarvestingStatus" || 
        f.InternalName === "Harvesting_x0020_Status" ||
        f.Title === "Harvesting Status" ||
        (f.Title && f.Title.toLowerCase().indexOf("harvesting") !== -1)
      )?.InternalName || fields.find((f: any) => 
        (f.Title && f.Title.toLowerCase().indexOf("status") !== -1 && 
         f.Title.toLowerCase().indexOf("project") === -1 && 
         f.Title !== "Title")
      )?.InternalName || "HarvestingStatus";

      debugLog(`[BOT] Discovery Results -> ProjectID: ${projectIdField}, Count: ${countField}, Status: ${statusField}`);
      debugLog(`[BOT] Available Fields in List:`, fields.map((f: any) => `${f.Title} (${f.InternalName})`).join(', '));

      // 1. Get the current item
      const queryUrl = `${webUrl}/_api/web/lists/getbytitle('KM Harvest Hub')/items?$select=Id,${countField}&$filter=${projectIdField} eq '${pId}'`;
      debugLog(`[BOT] Querying Harvest Hub: ${queryUrl}`);

      const response = await props.context.spHttpClient.get(queryUrl, SPHttpClient.configurations.v1);

      if (response.ok) {
        const data = await response.json();
        debugLog(`[BOT] Harvest Hub Results:`, data.value);
        if (data.value && data.value.length > 0) {
          const item = data.value[0];
          const currentCount = Number(item[countField]) || 0;
          const itemId = item.ID || item.Id;
          const nextCount = Math.max(0, currentCount + countIncrement);
          
          debugLog(`[BOT] Found Record! ID=${itemId}, CurrentCount=${currentCount}`);
          debugLog(`[BOT] Payload (Count handled by PowerAutomate):`, { [statusField]: newStatus });

          // 2. Update status (Count is now handled by Power Automate)
          const updateResp = await props.context.spHttpClient.post(
            `${webUrl}/_api/web/lists/getbytitle('KM Harvest Hub')/items(${itemId})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata.metadata=minimal',
                'Content-Type': 'application/json;odata.metadata=minimal',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
              },
              body: JSON.stringify({
                [statusField]: newStatus
              })
            }
          );

          if (updateResp.ok) {
            debugLog(`[BOT] Successfully updated Harvest Hub for project ${pId}`);
            if (countIncrement === 0) {
              setSuccessNotification({
                message: `Project status successfully updated to 'No Content' in KM Harvest Hub for project ${pId}`,
                type: 'no-content'
              });
            } else {
              setSuccessNotification({
                message: `Knowledge Artefact submitted!`,
                type: isMedia ? 'success-media' : 'success'
              });
            }
          } else {
            const errDetails = await updateResp.text();
            console.error(`[BOT] Failed to update Harvest Hub: ${errDetails}`);
            setSuccessNotification({
              message: `SharePoint Update Failed: ${errDetails}`,
              type: 'error'
            });
          }
          debugLog(`Successfully updated count (+${countIncrement}) for Project ${pId} in KM Harvest Hub.`);
        } else {
          console.warn(`[BOT] Project ${pId} not found in KM Harvest Hub list.`);
          if (newStatus === 'No Content') {
            setSuccessNotification({
              message: `Project ID ${pId} was not found in KM Harvest Hub.`,
              type: 'error'
            });
          }
        }
      } else {
        if (response.status === 500) {
          console.warn(
            'KM Harvest Hub threshold exceeded. ' +
            'Ask IT to add index on ProjectId field in KM Harvest Hub list settings.'
          );
          if (newStatus === 'No Content') {
            setSuccessNotification({
              message: 'Unable to update KM Harvest Hub because the Project ID lookup hit a SharePoint threshold. Please ask IT to index the Project ID field in KM Harvest Hub.',
              type: 'error'
            });
          }
          return;
        }
        console.warn('KM Harvest Hub query failed:', response.status);
        if (newStatus === 'No Content') {
          setSuccessNotification({
            message: `Unable to update KM Harvest Hub. SharePoint returned status ${response.status}.`,
            type: 'error'
          });
        }
        return;
      }
    } catch (err) {
      console.warn("KM Harvest Hub lookup failed - continuing upload", err);
      if (newStatus === 'No Content') {
        setSuccessNotification({
          message: 'Unable to update KM Harvest Hub. Please try again or contact IT.',
          type: 'error'
        });
      }
    }
  };


  const handleMediaSubmit = async (file: File, metadata: any) => {
    isSubmittingRef.current = true;
    setSuccessDocumentLinks([]);
    setSuccessDocumentUrl(null);
    setLastUploadedDocId(null);
    setIsProcessing(true);
    setProcessingError(null);
    setProcessingPhase('upload');
    setProcessingFileCount(1);
    setProcessingProgress('Uploading media to SharePoint...');

    try {
      const allTaxonomyOptions = await loadTaxonomyOptions();
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const fieldMap = await getKMDataHubFieldMap();
      logTaxonomyAvailability(allTaxonomyOptions);
      const normalizedMetadata = normalizeMetadataForKMDataHub(metadata, allTaxonomyOptions);
      await validateUniqueKMDataHubUpload(file, normalizedMetadata.title || file.name, fieldMap, 'media');
      const existingItem = await resolveTargetKMDataHubItem(file.name, fieldMap);

      if (existingItem) {
        const [fileChanged, changedFormValues] = await Promise.all([
          hasDifferentFileContent(file, existingItem.fileRef),
          buildKMDataHubFormValues(normalizedMetadata, fieldMap, { existingItem })
        ]);

        if (!fileChanged && changedFormValues.length === 0) {
          debugLog(`Skipping media overwrite for ${file.name}; file content and metadata are unchanged.`);
          setProcessingProgress('Media Uploaded Successfully');
          await updateWithoutVersion(
            props.context.spHttpClient,
            webUrl,
            existingItem.itemId,
            async () => {
              await storeUrlOnSubmit(existingItem.itemId, normalizedMetadata.title || file.name, true);
            }
          );
          const successUrl = buildAssetUrl(existingItem.itemId);
          const successDocumentLinks = [{
            id: String(existingItem.itemId),
            url: buildAssetUrl(existingItem.itemId),
            title: normalizedMetadata.title || file.name
          }];
          setLastUploadedDocId(String(existingItem.itemId));
          setSuccessDocumentUrl(successUrl);
          setSuccessDocumentLinks(successDocumentLinks);
          debugLog('Success URL set:', successUrl);
          debugLog('Success document links:', successDocumentLinks);
          if (props.projectId) {
            try { await updateHarvestTracker(props.projectId, 1, 'Submitted', true); } catch (e) {
              console.warn('Media tracker update failed');
              setSuccessNotification({ message: 'Media file submitted to KM Data Hub.', type: 'success-media' });
            }
          } else {
            setSuccessNotification({ message: 'Media file submitted successfully!', type: 'success-media' });
          }
          setTimeout(() => {
            setMediaFile(null);
            setProcessingProgress('');
            setProcessingFileCount(0);
            setIsProcessing(false);
            setProcessingPhase('idle');
            isSubmittingRef.current = false;
            props.onUploaded?.(file);
          }, 1500);
          return;
        }

        if (!fileChanged && changedFormValues.length > 0) {
          setProcessingProgress('Saving metadata...');
          await updateWithoutVersion(
            props.context.spHttpClient,
            webUrl,
            existingItem.itemId,
            async () => {
              await submitKMDataHubFormValues(existingItem.itemId, changedFormValues, true, normalizedMetadata, true);
              await storeUrlOnSubmit(existingItem.itemId, normalizedMetadata.title || file.name, true);
            }
          );
          await logKMDataHubStoredValues(existingItem.itemId, fieldMap);

          try {
            await createAuditLogItem(file, 'Under Review', normalizedMetadata.title || file.name);
          } catch (err) {
            console.warn(`Audit log creation failed for media ${file.name}:`, err);
          }

          setProcessingProgress('Media Uploaded Successfully');
          const successUrl = buildAssetUrl(existingItem.itemId);
          const successDocumentLinks = [{
            id: String(existingItem.itemId),
            url: buildAssetUrl(existingItem.itemId),
            title: normalizedMetadata.title || file.name
          }];
          setLastUploadedDocId(String(existingItem.itemId));
          setSuccessDocumentUrl(successUrl);
          setSuccessDocumentLinks(successDocumentLinks);
          debugLog('Success URL set:', successUrl);
          debugLog('Success document links:', successDocumentLinks);
          if (props.projectId) {
            try { await updateHarvestTracker(props.projectId, 1, 'Submitted', true); } catch (e) {
              console.warn('Media tracker update failed');
              setSuccessNotification({ message: 'Media file submitted to KM Data Hub.', type: 'success-media' });
            }
          } else {
            setSuccessNotification({ message: 'Media file submitted successfully!', type: 'success-media' });
          }
          setTimeout(() => {
            setMediaFile(null);
            setProcessingProgress('');
            setProcessingFileCount(0);
            setIsProcessing(false);
            setProcessingPhase('idle');
            isSubmittingRef.current = false;
            props.onUploaded?.(file);
          }, 1500);
          return;
        }
      }

      let fileUrl = "";

      // 1) Chunked or Simple Upload
      const uploadMediaFile = async (): Promise<void> => {
        fileUrl = await uploadFileContentToTarget(file, existingItem?.fileRef);
      };

      if (file.size <= 10485760) {
        setProcessingProgress('Uploading media file...');
      } else {
        setProcessingProgress('Initializing upload for large media file...');
      }

      const effectiveMediaItemId = props.targetItemId ||
        (existingItem ? existingItem.itemId : undefined);

      if (existingItem && effectiveMediaItemId) {
        await uploadNewFileVersion(
          props.context.spHttpClient,
          webUrl,
          effectiveMediaItemId,
          uploadMediaFile
        );
      } else {
        await uploadMediaFile();
      }

      setProcessingProgress('Saving metadata...');

      // Wait briefly
      await new Promise(r => setTimeout(r, 500));

      // 2) Get Item ID
      const itemResp = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/GetFileByServerRelativeUrl('${fileUrl}')/ListItemAllFields?$select=Id`,
        SPHttpClient.configurations.v1
      );
      if (!itemResp.ok) throw new Error("Failed to get list item");
      const itemId = (await itemResp.json()).Id;
      const mediaFileRef = fileUrl;
      if (existingItem && effectiveMediaItemId) {
        const mediaExistingItem = await fetchKMDataHubItemSnapshot(itemId, fieldMap);
        const mediaFormValues = await buildKMDataHubFormValues(normalizedMetadata, fieldMap, {
          existingItem: mediaExistingItem
        });
        const mediaExistingItemFormValues = [
          ...mediaFormValues,
          {
            FieldName: COLUMN_NAMES.versionFileName,
            FieldValue: file.name
          },
          {
            FieldName: COLUMN_NAMES.versionFileType,
            FieldValue: file.name.split('.').pop() || ''
          }
        ];

        await updateWithoutVersion(
          props.context.spHttpClient,
          webUrl,
          effectiveMediaItemId,
          async () => {
            await storeUrlOnSubmit(itemId, normalizedMetadata.title || file.name, true);
            if (mediaExistingItemFormValues.length > 0) {
              debugLog('KM Data Hub changed form values:', mediaExistingItemFormValues);
              await submitKMDataHubFormValues(itemId, mediaExistingItemFormValues, true, normalizedMetadata, true);
            }
          }
        );
      } else {
        await saveMetadataAndCheckin(
          props.context.spHttpClient,
          webUrl,
          mediaFileRef,
          async () => {
            await storeUrlOnSubmit(itemId, normalizedMetadata.title || file.name);
            await updateKMDataHubWithFormData(itemId, normalizedMetadata);
          }
        );
      }
      const successUrl = buildAssetUrl(itemId);
      const successDocumentLinks = [{
        id: String(itemId),
        url: buildAssetUrl(itemId),
        title: normalizedMetadata.title || file.name
      }];
      setLastUploadedDocId(String(itemId));
      setSuccessDocumentUrl(successUrl);
      setSuccessDocumentLinks(successDocumentLinks);
      debugLog('Success URL set:', successUrl);
      debugLog('Success document links:', successDocumentLinks);

      // Store in session tracker
      setSessionUploadedDocs(prev => [
        {
          id: itemId,
          name: file.name,
          time: new Date().toLocaleTimeString()
        },
        ...prev
      ]);

      // Create audit log item for media
      try {
        await createAuditLogItem(file, "Under Review", normalizedMetadata.title || file.name);
      } catch (err) {
        console.warn(`Audit log creation failed for media ${file.name}:`, err);
      }

      if (props.projectId) {
         try {
           await updateHarvestTracker(props.projectId, 1, 'Submitted', true);
         } catch (e) {
           console.warn("Media tracker update failed");
           setSuccessNotification({ message: 'Media file submitted to KM Data Hub.', type: 'success-media' });
         }
      } else {
        setSuccessNotification({ message: 'Media file submitted successfully!', type: 'success-media' });
      }

      setProcessingProgress('Media Uploaded Successfully');

      setTimeout(() => {
        setMediaFile(null);
        setProcessingProgress('');
        setProcessingFileCount(0);
        setIsProcessing(false);
        setProcessingPhase('idle');
        isSubmittingRef.current = false;
        props.onUploaded?.(file);
      }, 1500);

    } catch (err) {
      console.error(err);
      setProcessingError(err instanceof Error ? err.message : 'Error uploading media file');
      isSubmittingRef.current = false;
      // NOTE: NOT calling setIsProcessing(false) here so the error stays visible
      // in the spinner UI. The "Go Back" button allows returning to VideoAnalysis.
    }
  };

  const handleNoArtifacts = () => {
    if (projectStatus?.trim() === 'Submitted' || projectCount > 0) {
      showUploadValidationMessage('This project has already been submitted and cannot be updated.');
      return;
    }
    if (!props.projectId) return;
    setShowConfirmation(true);
  };

  const handleReplaceFile = async (file: File): Promise<void> => {
    if (!props.context || !props.targetItemId || !props.targetFileRef) {
      console.error('Replace mode requires targetItemId and targetFileRef');
      return;
    }

    setIsProcessing(true);
    setProcessingError(null);
    setProcessingPhase('upload');
    setProcessingProgress('Uploading new version...');

    const webUrl = props.context.pageContext.web.absoluteUrl;

    try {
      const existingFileRef = props.targetFileRef;
      const existingFolder = getFolderServerRelativeUrl(existingFileRef);
      const existingFileName = existingFileRef.split('/').pop() || '';
      const newFileName = file.name;
      const newFileRef = `${existingFolder}/${newFileName}`;
      const isSameFileName = newFileName.toLowerCase() === existingFileName.toLowerCase();
      let activeFileRef = existingFileRef;
      let existingTitle = '';

      debugLog('Replace: existing file:', existingFileRef);
      debugLog('Replace: new file name:', newFileName);
      debugLog('Replace: new file ref:', newFileRef);
      debugLog('Replace: file size:', file.size, 'bytes');
      debugLog('Replace: file type:', file.type);

      try {
        const titleResp = await props.context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
          `/items(${props.targetItemId})?$select=Title`,
          SPHttpClient.configurations.v1
        );
        if (titleResp.ok) {
          const titleData = await titleResp.json();
          existingTitle = titleData?.d?.Title || titleData?.Title || '';
          debugLog('Replace: existing Title saved before file change:', existingTitle);
        }
      } catch (titleErr) {
        console.warn('Replace: could not fetch existing Title before file change:', titleErr);
      }

      await checkoutFile(props.context.spHttpClient, webUrl, existingFileRef);

      try {
        if (!isSameFileName) {
          const moveResp = await props.context.spHttpClient.post(
            `${webUrl}/_api/web/GetFileByServerRelativeUrl('${escapeODataString(existingFileRef)}')` +
            `/MoveTo(newUrl='${escapeODataString(newFileRef)}',flags=1)`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: 'application/json;odata=nometadata'
              }
            }
          );

          if (!moveResp.ok) {
            const moveText = await moveResp.text();
            throw new Error(`File rename failed: ${moveResp.status} ${moveText}`);
          }

          activeFileRef = newFileRef;
          debugLog('Replace: file renamed to:', activeFileRef);
        }

        debugLog('Replace: uploading new bytes for file:', file.name);
        debugLog('Replace: target fileRef:', activeFileRef);

        const uploadedFileRef = await uploadFileContentToTarget(file, activeFileRef);
        debugLog('Replace: upload complete, uploadedFileRef:', uploadedFileRef);
        debugLog('Replace: original targetFileRef:', props.targetFileRef);

        await checkinMajor(
          props.context.spHttpClient,
          webUrl,
          activeFileRef,
          'New file version uploaded'
        );
        debugLog('Replace: new version checked in');
      } catch (uploadErr) {
        try {
          await undoCheckout(props.context.spHttpClient, webUrl, activeFileRef);
        } catch (undoErr) {
          console.warn('Replace: undo checkout failed after upload error:', undoErr);
        }
        throw uploadErr;
      }

      const fieldMap = await getKMDataHubFieldMap();
      if (fieldMap.status) {
        const newFileExtension = newFileName.split('.').pop() || '';

        await updateWithoutVersion(
          props.context.spHttpClient,
          webUrl,
          props.targetItemId,
          async () => {
            const response = await props.context.spHttpClient.post(
              `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${props.targetItemId})/ValidateUpdateListItem`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept: 'application/json;odata=verbose',
                  'Content-Type': 'application/json;odata=verbose',
                  'odata-version': ''
                },
                body: JSON.stringify({
                  formValues: appendKMReviewAuditFormValues([
                    {
                      FieldName: COLUMN_NAMES.title,
                      FieldValue: existingTitle || ''
                    },
                    {
                      FieldName: fieldMap.status,
                      FieldValue: 'Under Review'
                    },
                    {
                      FieldName: COLUMN_NAMES.published,
                      FieldValue: ''
                    },
                    {
                      FieldName: COLUMN_NAMES.versionFileName,
                      FieldValue: newFileName
                    },
                    {
                      FieldName: COLUMN_NAMES.versionFileType,
                      FieldValue: newFileExtension
                    },
                    {
                      FieldName: COLUMN_NAMES.contentRefreshDate,
                      FieldValue: formatSharePointDateFieldValue()
                    }
                  ]),
                  bNewDocumentUpdate: false
                })
              }
            );

            if (!response.ok) {
              const errorText = await response.text();
              console.warn('Status update to Under Review failed:', response.status, errorText);
            } else {
              debugLog('Replace: Title → preserved as:', existingTitle);
              debugLog('Replace: Status → Under Review ✅');
              debugLog('Replace: Published → cleared ✅');
              debugLog('Replace: FileLeafRef auto-updated by SP ✅');
              debugLog('Replace: DocIcon auto-updated by SP ✅');
            }
          }
        );
      }

      try {
        await createAuditLogItem(file, 'Under Review', existingTitle || file.name);
      } catch (auditErr) {
        console.warn('Audit log failed on replace:', auditErr);
      }

      setProcessingProgress('Document updated successfully!');

      await new Promise(r => setTimeout(r, 1000));

      setIsProcessing(false);
      setProcessingPhase('idle');
      setProcessingProgress('');
      setProcessingFileCount(0);
      setLastUploadedDocId(String(props.targetItemId));
      const updatedDocumentUrl = buildAssetUrl(props.targetItemId);
      setSuccessDocumentUrl(updatedDocumentUrl);
      setSuccessDocumentLinks([{
        id: String(props.targetItemId),
        url: updatedDocumentUrl,
        title: existingTitle || file.name
      }]);
      setSuccessNotification({
        message: 'Thank you for new version Update',
        type: 'success-update'
      });
      props.onUploaded?.(file);
    } catch (err) {
      console.error('File replace failed:', err);
      setProcessingError(err instanceof Error ? err.message : 'File replace failed. Please try again.');
      setIsProcessing(false);
      setProcessingPhase('idle');
    }
  };

  const confirmNoArtifacts = async () => {
    if (projectStatus?.trim() === 'Submitted' || projectCount > 0) {
      setShowConfirmation(false);
      showUploadValidationMessage('This project has already been submitted and cannot be updated.');
      return;
    }
    setShowConfirmation(false);
    setIsProcessing(true);
    setProcessingPhase('generic');
    setProcessingProgress('Updating project status to No Content...');

    try {
      await updateHarvestTracker(props.projectId!, 0, 'No Content');
      setIsProcessing(false);
      setProcessingPhase('idle');
      setProcessingProgress('');
      // The success notification is set inside updateHarvestTracker
    } catch (err) {
      console.error(err);
      setSuccessNotification({
        message: 'Failed to update project status.',
        type: 'error'
      });
      setIsProcessing(false);
      setProcessingPhase('idle');
    }
  };

  const renderManualUploadCard = (): React.ReactElement => (
    <div
      className={`${styles.uploadCard} ${isReplaceMode ? styles.replaceUploadCard : ''}`}
    >
      <div
        className={styles.mainContent}
      >
        <h3 className={styles.title}>{isReplaceMode ? 'Update Document' : 'Upload Asset'}</h3>
        <div
          className={`${styles.dropZone} ${dragOver ? styles.dropZoneHover : ''}`}
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          onClick={(e) => {
            if ((e.target as HTMLElement).tagName !== 'BUTTON') {
              onBrowse();
            }
          }}
          role="button"
          aria-label="Upload file"
        >
          <div className={styles.uploadIconWrapper}>
            <svg className={styles.uploadIcon} viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" aria-hidden>
              <path d="M19 15v4a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
              <path d="M7 9l5-5 5 5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
              <path d="M12 4v12" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </div>
          <div className={styles.hintText}>Drag & drop files here</div>
          <div className={styles.instructions}>
            {isReplaceMode
              ? 'Drop the new file here. Version will increment and status will change to Under Review.'
              : 'You can upload Word, Ppt, Excel, PDF, TXT, Video and Audio files.'}
          </div>
          <div className={styles.fileTypes}>
            {isReplaceMode ? 'Select one replacement file.' : 'A maximum of 5 files can be uploaded at once.'}
          </div>

          <div className={styles.uploadButtonRow}>
            <button type="button" className={styles.browseBtn} onClick={(e) => { e.stopPropagation(); onBrowse(); }}>
              <svg className={styles.uploadButtonIcon} width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M12 16V4" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
                <path d="M7 9l5-5 5 5" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
                <path d="M5 20h14" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" />
              </svg>
              {isReplaceMode ? 'Browse file' : 'Upload Document'}
            </button>
            {!isReplaceMode && (
              <button type="button" className={styles.browseBtn} onClick={(e) => { e.stopPropagation(); onBrowseMedia(); }}>
                <svg className={styles.uploadButtonIcon} width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <path d="M12 16V4" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M7 9l5-5 5 5" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M5 20h14" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" />
                </svg>
                Upload Media
              </button>
            )}
          </div>
        </div>
        {props.onClose && (
          <button
            type="button"
            className={styles.uploadCardCloseBtn}
            onClick={(e) => {
              e.stopPropagation();
              props.onClose?.();
            }}
          >
            Close
          </button>
        )}
      </div>
    </div>
  );

  const renderInitialUploadCard = (): React.ReactElement => {
    const shouldShowProjectValidation =
      isNoContentProjectAction &&
      !showConfirmation &&
      !successNotification;

    if (shouldShowProjectValidation) {
      return (
        <div className={styles.processingContainer}>
          <div className={styles.processingSpinner}></div>
          <h3 className={styles.processingTitle}>Validating Project...</h3>
        </div>
      );
    }

    return renderManualUploadCard();
  };

  const shellUserName = props.userName || props.context?.pageContext.user.displayName;
  const shellUserEmail = props.userEmail || props.context?.pageContext.user.email || props.context?.pageContext.user.loginName;
  const handleUploadHomeOpen = React.useCallback((): void => {
    const url = new URL(window.location.href);
    const projectId = url.searchParams.get('projectId');
    url.pathname = NAV_PATHS.home;
    url.search = '';
    if (projectId) {
      url.searchParams.set('projectId', projectId);
    }
    url.hash = '';
    window.history.pushState({ iknShowHome: true, sameTabNav: true }, document.title, url.toString());
    window.dispatchEvent(new PopStateEvent('popstate'));
    props.onClose?.();
  }, [props.onClose]);
  const shellHeaderProps = {
    userName: shellUserName,
    userEmail: shellUserEmail,
    userPhotoUrl: props.userPhotoUrl,
    compact: isReplaceMode,
    onLogout: props.onLogout,
    onHomeOpen: handleUploadHomeOpen,
    onAllDocumentsOpen: props.onAllDocumentsOpen,
    onBusinessUnitsOpen: props.onBusinessUnitsOpen,
    onBookmarksOpen: props.onBookmarksOpen,
    onDocumentsOpen: props.onDocumentsOpen,
    onContactOpen: props.onContactOpen,
    onAuditLogOpen: props.onAuditLogOpen,
    onAnalyticsOpen: props.onAnalyticsOpen,
    showReviewerNav: props.showReviewerNav,
    hideDocumentsNav: props.hideDocumentsNav
  };
  const uploadValidationError = successNotification?.type === 'error' ? successNotification : null;
  const uploadValidationMessageLines = uploadValidationError?.message
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean) || [];
  const uploadValidationTitle = uploadValidationMessageLines[0] || 'Something went wrong';
  const uploadValidationBody = uploadValidationMessageLines.slice(1).join('\n');
  const isNoContentProjectAction =
    isKmArtifactUpload &&
    new URLSearchParams(window.location.search).get('action') === 'noContent';
  const isProjectUploadSuccess =
    Boolean(props.projectId) &&
    Boolean(successNotification) &&
    (successNotification?.type === 'success' || successNotification?.type === 'success-media' || successNotification?.type === 'success-update');

  return (
    <div className={`${styles.overlay} ${isCompactFlow ? styles.compactFlow : ''}`} style={overlayStyle}>
      <div className={styles.modal} role="dialog" aria-modal="true">
        {/* State 1: Processing / Spinner */}
        {isProcessing ? (
          <div className={`${styles.processingPage} ${isReplaceMode ? styles.replaceProcessingPage : ''}`}>
            <div className={styles.uploadShellOverlayInner}>
              <IKShellHeader {...shellHeaderProps} />

              <div className={styles.processingBody}>
              <div className={styles.processingContainer}>
                <div className={styles.processingSpinner}></div>
                <h3 className={styles.processingTitle}>
                  {processingTitle}
                </h3>
                <p className={styles.processingMessage}>
                  {processingMessage}
                </p>
              </div>
            </div>

            <IKShellFooter compact={isReplaceMode} className={styles.uploadFlowFooter} onBackHome={handleUploadHomeOpen} />
            {processingError && (
              <div className={styles.errorMessage}>
                {processingError}
                <button
                  className={styles.browseBtn}
                  onClick={() => { setIsProcessing(false); setProcessingError(null); }}
                  style={{ display: 'block', marginTop: '10px', margin: 'auto', marginLeft: 'auto', marginRight: 'auto' }}
                >
                  Go Back
                </button>
              </div>
            )}
          </div>
        </div>
        ) : mediaFile ? (
          /* State 2: Media Analysis */
          <div className={styles.mediaAnalysisPage}>
            <div className={styles.uploadShellOverlayInner}>
              <IKShellHeader {...shellHeaderProps} />
              <div className={styles.mediaAnalysisBody}>
                <VideoAnalysis
                  file={mediaFile}
                  context={props.context}
                  onClose={() => setMediaFile(null)}
                  onAnalysisComplete={async (data) => {
                    try {
                      const allTaxonomyOptions = await loadTaxonomyOptions();
                      const mediaMetadataForNormalization = {
                        ...data,
                        title: data?.title || '',
                        description: cleanTextField(data?.description || '', SHAREPOINT_SHORT_TEXT_MAX_LENGTH),
                        documentType: data?.documentType || '',
                        client: data?.client || '',
                        geography: data?.geography || '',
                        therapyArea: data?.therapyArea || '',
                        diseaseArea: data?.diseaseArea || '',
                        bu: data?.bu || '',
                        department: data?.department || '',
                        subDepartment: '',
                        sensitiveTerms: data?.sensitiveTerms || ''
                      };
                      debugLog('Object passed to normalizeMetadataForKMDataHub:', JSON.stringify(mediaMetadataForNormalization));
                      logTaxonomyAvailability(allTaxonomyOptions);
                      const normalizedMediaMetadata = normalizeMetadataForKMDataHub(mediaMetadataForNormalization, allTaxonomyOptions);
                      const fieldMap = await getKMDataHubFieldMap();

                      if (mediaFile) {
                        await validateUniqueKMDataHubUpload(
                          mediaFile,
                          normalizedMediaMetadata.title || mediaFile.name,
                          fieldMap,
                          'media'
                        );
                      }

                      setMediaFile(null);
                      setUploadedFiles(mediaFile ? [mediaFile] : []);
                      setFilesData([{
                        file: mediaFile!,
                        itemId: -1,
                        metadata: normalizedMediaMetadata
                      }]);
                      setShowForm(true);
                    } catch (error) {
                      console.error('Media duplicate validation failed:', error);
                      setMediaFile(null);
                      setUploadedFiles([]);
                      setFilesData([]);
                      setShowForm(false);
                      showUploadValidationMessage(error instanceof Error ? error.message : 'Unable to validate media uniqueness.');
                    }
                  }}
                />
              </div>
              <IKShellFooter compact={isReplaceMode} className={styles.uploadFlowFooter} onBackHome={handleUploadHomeOpen} />
            </div>
          </div>
        ) : (showForm && taxonomyOptions && filesData.length > 0) ? (
          /* State 3: Metadata Form (Single or Multi) */
          <div className={styles.metadataPage}>
            <div className={styles.uploadShellOverlayInner}>
              <IKShellHeader {...shellHeaderProps} />
              <div className={styles.metadataPageBody}>
                {filesData.length > 1 ? (
                  <div className={styles.metadataAccordion}>
                    <div className={styles.metadataAccordionList}>
                      {filesData.map((fileData, index) => {
                        const isOpen = activeMetadataFileIndex === index;
                        const isDraftSubmitted = multiMetadataValidity[index] === true;

                        return (
                          <div className={`${styles.metadataAccordionItem} ${isOpen ? styles.metadataAccordionItemOpen : ''}`} key={`${fileData.file.name}-${index}`}>
                            <div className={styles.metadataAccordionRow}>
                              <button
                                type="button"
                                className={styles.metadataAccordionToggle}
                                aria-expanded={isOpen}
                                onClick={() => setActiveMetadataFileIndex(isOpen ? -1 : index)}
                              >
                                <span className={styles.metadataAccordionTitle}>File {index + 1} - {fileData.file.name}</span>
                                {isDraftSubmitted && <span className={styles.metadataFormSaved}>Ready</span>}
                                <span className={styles.metadataAccordionChevron} aria-hidden="true">
                                  <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.3" strokeLinecap="round" strokeLinejoin="round">
                                    <path d="M6 9l6 6 6-6" />
                                  </svg>
                                </span>
                              </button>
                              <button
                                type="button"
                                className={styles.metadataAccordionDelete}
                                aria-label={`Remove ${fileData.file.name}`}
                                title={`Remove ${fileData.file.name}`}
                                onClick={() => removeMetadataFileAtIndex(index)}
                              >
                                <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                                  <path d="M3 6h18" />
                                  <path d="M8 6V4h8v2" />
                                  <path d="M19 6l-1 14H6L5 6" />
                                  <path d="M10 11v5" />
                                  <path d="M14 11v5" />
                                </svg>
                              </button>
                            </div>
                            {isOpen && (
                              <div className={styles.metadataAccordionPanel}>
                                <MetadataForm
                                  context={props.context}
                                  onClose={() => {
                                    setUploadedFiles([]);
                                    setFilesData([]);
                                    setShowForm(false);
                                    setActiveMetadataFileIndex(0);
                                    setMultiMetadataDrafts({});
                                    setMultiMetadataValidity({});
                                    isSubmittingRef.current = false;
                                    props.onClose && props.onClose();
                                  }}
                                  initialValues={multiMetadataDrafts[index] || fileData.metadata || undefined}
                                  taxonomyOptions={taxonomyOptions}
                                  showReviewerFields={shouldShowReviewerFields}
                                  hideActions={true}
                                  density={isCompactFlow ? 'compact' : 'default'}
                                  onDraftChange={(data, isValid) => handleMetadataDraftChange(index, data, isValid)}
                                />
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>
                    <div className={styles.metadataSubmitAllBar}>
                      <button
                        type="button"
                        className={styles.metadataSubmitAllBtn}
                        disabled={Object.entries(multiMetadataValidity).filter(([key, value]) => {
                          const numericKey = Number(key);
                          return value === true && numericKey >= 0 && numericKey < filesData.length;
                        }).length < filesData.length}
                        onClick={() => {
                          const firstInvalidIndex = filesData.findIndex((_fileData, index) => multiMetadataValidity[index] !== true);
                          if (firstInvalidIndex !== -1) {
                            setActiveMetadataFileIndex(firstInvalidIndex);
                            return;
                          }

                          const submittedFilesData = filesData.map((fileData, index) => ({
                            ...fileData,
                            metadata: multiMetadataDrafts[index] || fileData.metadata || {}
                          }));
                          onFormSubmit(submittedFilesData);
                        }}
                      >
                        Submit All ({filesData.length} files)
                      </button>
                      <button
                        type="button"
                        className={styles.metadataCloseAllBtn}
                        onClick={() => {
                          setUploadedFiles([]);
                          setFilesData([]);
                          setShowForm(false);
                          setActiveMetadataFileIndex(0);
                          setMultiMetadataDrafts({});
                          setMultiMetadataValidity({});
                          isSubmittingRef.current = false;
                          props.onClose && props.onClose();
                        }}
                      >
                        Close
                      </button>
                    </div>
                  </div>
                ) : (
                  <MetadataForm
                    context={props.context}
                    onSubmit={(data) => {
                      onFormSubmit([{
                        file: filesData[0].file,
                        itemId: filesData[0].itemId,
                        metadata: data
                      }]);
                    }}
                    onClose={() => {
                      setUploadedFiles([]);
                      setFilesData([]);
                      setShowForm(false);
                      isSubmittingRef.current = false;
                      props.onClose && props.onClose();
                    }}
                    initialValues={filesData[0].metadata || undefined}
                    taxonomyOptions={taxonomyOptions}
                    showReviewerFields={shouldShowReviewerFields}
                    density={isCompactFlow ? 'compact' : 'default'}
                  />
                )}
              </div>
              <IKShellFooter compact={isReplaceMode} className={styles.uploadFlowFooter} onBackHome={handleUploadHomeOpen} />
            </div>
          </div>
        ) : (uploadedFiles.length === 0 && !uploadValidationError) ? (
          <div className={`${styles.processingPage} ${isNoContentProjectAction ? styles.projectValidationPage : ''}`}>
            <div className={styles.uploadShellOverlayInner}>
              <IKShellHeader {...shellHeaderProps} />
              <div className={`${styles.uploadInitialBody} ${isReplaceMode ? styles.replaceUploadInitialBody : ''}`}>
                {renderInitialUploadCard()}
              </div>
              <IKShellFooter compact={isReplaceMode} className={styles.uploadFlowFooter} onBackHome={handleUploadHomeOpen} />
            </div>
          </div>
        ) : null}

        {/* Global Overlays */}
        {showConfirmation && (
          <div className={`${styles.uploadShellOverlay} ${styles.confirmSubmissionPage}`} style={overlayStyle}>
            <div className={styles.uploadShellOverlayInner}>
              <div className={styles.uploadFlowHeader}>
                <IKShellHeader {...shellHeaderProps} />
              </div>
              <div className={styles.uploadInitialBody}>
                <div className={styles.confirmSubmissionCard}>
                  <h3 className={styles.confirmSubmissionTitle}>Confirm Submission</h3>
                  <p className={styles.confirmSubmissionText}>
                    Are you sure you have no artifacts to share for Project <strong>{props.projectId}</strong>?
                  </p>
                  <div className={styles.confirmSubmissionActions}>
                    <button
                      className={styles.confirmSubmissionPrimary}
                      onClick={confirmNoArtifacts}
                    >
                      Yes, Confirm
                    </button>
                    <button
                      className={styles.confirmSubmissionSecondary}
                      onClick={() => setShowConfirmation(false)}
                    >
                      Cancel
                    </button>
                  </div>
                </div>
              </div>
              <IKShellFooter compact={isReplaceMode} className={styles.uploadFlowFooter} onBackHome={handleUploadHomeOpen} />
            </div>
          </div>
        )}

        {uploadValidationError && (
          <div className={`${styles.processingPage} ${styles.uploadValidationErrorPage} ${props.projectId ? styles.projectValidationErrorPage : ''}`} style={overlayStyle}>
            <div className={styles.uploadShellOverlayInner}>
              <div className={styles.uploadFlowHeader}>
                <IKShellHeader {...shellHeaderProps} />
              </div>
              <div className={styles.uploadInitialBody}>
                <div className={styles.successNotificationContent}>
                  <button
                    className={styles.modalCloseX}
                    onClick={() => {
                      clearUploadSuccessRestorePayload();
                      setSuccessNotification(null);
                      setSuccessDocumentUrl(null);
                      setLastUploadedDocId(null);
                      setSuccessDocumentLinks([]);
                    }}
                  >✕</button>

                  <div className={styles.validationErrorIconCircle}>
                    <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M6 6L18 18M18 6L6 18" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  </div>

                  <h3 className={styles.successTitle}>{uploadValidationTitle}</h3>

                  {uploadValidationBody && (
                    <div className={styles.successText}>
                      <p style={{ whiteSpace: 'pre-line' }}>{uploadValidationBody}</p>
                    </div>
                  )}

                  <div className={styles.successButtonGroup}>
                    <button
                      className={styles.closeActionBtn}
                      onClick={() => {
                        clearUploadSuccessRestorePayload();
                        setSuccessNotification(null);
                        setSuccessDocumentUrl(null);
                        setLastUploadedDocId(null);
                        setSuccessDocumentLinks([]);
                      }}
                    >
                      Close
                    </button>
                  </div>
                </div>
              </div>
              <IKShellFooter compact={isReplaceMode} className={styles.uploadFlowFooter} onBackHome={handleUploadHomeOpen} />
            </div>
          </div>
        )}

        {successNotification && successNotification.type !== 'error' && (
          <div
            className={`${styles.processingPage} ${
              (successNotification.type === 'success' || successNotification.type === 'success-media' || successNotification.type === 'success-update')
                ? `${styles.successPage} ${isProjectUploadSuccess ? styles.projectSuccessPage : ''} ${successNotification.type === 'success-update' ? styles.updateFlowTypography : ''}`
                : successNotification.type === 'no-content'
                  ? styles.noContentThanksPage
                  : ''
            }`}
            style={overlayStyle}
          >
            <div className={styles.uploadShellOverlayInner}>
              <div className={styles.uploadFlowHeader}>
                <IKShellHeader {...shellHeaderProps} />
              </div>
              <div className={styles.uploadInitialBody}>
                <div className={styles.successNotificationContent}>
                  {!(successNotification.type === 'success' || successNotification.type === 'success-media' || successNotification.type === 'success-update') && (
                    <button
                      className={styles.modalCloseX}
                      onClick={() => {
                        clearUploadSuccessRestorePayload();
                        setSuccessNotification(null);
                        setSuccessDocumentUrl(null);
                        setLastUploadedDocId(null);
                        setSuccessDocumentLinks([]);
                      }}
                    >✕</button>
                  )}

                  {(successNotification.type === 'success' || successNotification.type === 'success-media' || successNotification.type === 'success-update') ? (
                    <div className={styles.successSmileIcon} aria-hidden="true">
                      <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <circle cx="12" cy="12" r="9" stroke="currentColor" strokeWidth="1.8" />
                        <circle cx="9" cy="10" r="1" fill="currentColor" />
                        <circle cx="15" cy="10" r="1" fill="currentColor" />
                        <path d="M8.5 14.2C9.4 15.5 10.5 16.1 12 16.1C13.5 16.1 14.6 15.5 15.5 14.2" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
                      </svg>
                    </div>
                  ) : (
                    <div className={styles.successIconCircle}>
                      <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M20 6L9 17L4 12" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                    </div>
                  )}

                  <h3 className={`${styles.successTitle} ${successNotification.type === 'success-update' ? styles.updateSuccessTitle : ''}`}>
                    {successNotification.type === 'no-content'
                      ? 'Thanks for confirmation'
                      : successNotification.type === 'success-update'
                        ? 'Thank you for new version Update'
                        : (successNotification.type === 'success' || successNotification.type === 'success-media')
                          ? 'Thank you for your submission!!'
                        : 'Something went wrong'}
                  </h3>

                  <div className={styles.successText}>
                    {(successNotification.type === 'success' || successNotification.type === 'success-media' || successNotification.type === 'success-update') ? (
                      <>
                        <p className={styles.successSubtitle}>This is now under KM review</p>
                        {successNotification.type !== 'success-update' && (successDocumentLinks.length > 0 || successDocumentUrl) ? (
                          <>
                            <p>You can access your contribution(s) here</p>
                            <div className={styles.successLinks}>
                              {successDocumentLinks.length > 0 ? successDocumentLinks.map((documentLink, index) => (
                                <a
                                  key={documentLink.id}
                                  href={documentLink.url}
                                  className={styles.successLink}
                                  onClick={(event) => {
                                    event.preventDefault();
                                    persistSuccessStateForReturn();
                                    openSuccessDocumentInApp(documentLink.id, documentLink.url);
                                  }}
                                >
                                  {index + 1}. {documentLink.title}
                                </a>
                              )) : successDocumentUrl ? (
                                <a
                                  href={successDocumentUrl}
                                  className={styles.successLink}
                                  onClick={(event) => {
                                    event.preventDefault();
                                    if (lastUploadedDocId) {
                                      persistSuccessStateForReturn();
                                      openSuccessDocumentInApp(lastUploadedDocId, successDocumentUrl);
                                    } else {
                                      window.history.pushState({ sameTabNav: true }, document.title, successDocumentUrl);
                                      window.dispatchEvent(new PopStateEvent('popstate'));
                                    }
                                  }}
                                >
                                  1. View contribution
                                </a>
                              ) : null}
                            </div>
                          </>
                        ) : null}
                      </>
                    ) : successNotification.type === 'no-content' ? (
                      <>
                        <p>Your Project ID <span style={{ color: '#2563eb', fontWeight: 800 }}>{props.projectId}</span></p>
                        <p>has been updated in <strong>KM Harvest Hub</strong>.</p>
                      </>
                    ) : (
                      <p>{successNotification.message}</p>
                    )}
                  </div>

                  <div className={styles.successButtonGroup}>
                    {(successNotification.type === 'success' || successNotification.type === 'success-media') && (
                      <button
                        className={styles.uploadAnotherBtn}
                        onClick={() => {
                          clearUploadSuccessRestorePayload();
                          setSuccessNotification(null);
                          setSuccessDocumentUrl(null);
                          setLastUploadedDocId(null);
                          setSuccessDocumentLinks([]);
                        }}
                      >
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ width: '18px', marginRight: '8px' }}>
                          <path d="M21 15v4a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12" />
                        </svg>
                        {successNotification.type === 'success-media' ? 'Upload another media' : 'Upload another document'}
                      </button>
                    )}

                    <button
                      className={styles.closeActionBtn}
                      onClick={() => {
                        clearUploadSuccessRestorePayload();
                        setSuccessNotification(null);
                        setSuccessDocumentUrl(null);
                        setLastUploadedDocId(null);
                        setSuccessDocumentLinks([]);
                        if (successNotification.type === 'success' || successNotification.type === 'success-media' || successNotification.type === 'success-update' || successNotification.type === 'no-content') {
                          props.onClose && props.onClose();
                        }
                      }}
                    >
                      {(successNotification.type === 'success' || successNotification.type === 'success-media' || successNotification.type === 'success-update') && (
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
                          <path d="M18 6L6 18M6 6l12 12" strokeLinecap="round" strokeLinejoin="round" />
                        </svg>
                      )}
                      Close
                    </button>
                  </div>
                </div>
              </div>
              <IKShellFooter compact={isReplaceMode} className={styles.uploadFlowFooter} onBackHome={handleUploadHomeOpen} />
            </div>
          </div>
        )}

        {showSuccessModal && (
          <div className={styles.uploadShellOverlay}>
            <div className={styles.uploadShellOverlayInner}>
              <IKShellHeader {...shellHeaderProps} />
              <div className={styles.uploadInitialBody}>
                <div className={styles.successModal}>
                  <div className={styles.successIconOuter}>
                    <div className={styles.successIconInner}>✓</div>
                  </div>
                  <h3 className={styles.successTitle}>Thank you for your Contribution!</h3>
                  <p className={styles.successMessage}>
                    Your Document ID is <strong>{lastUploadedDocId}</strong><br />
                    and it is now in KM Review.
                  </p>
                  <div className={styles.successActions}>
                    <button
                      className={styles.uploadAnotherBtn}
                      onClick={() => {
                        setShowSuccessModal(false);
                        setFilesData([]);
                        setUploadedFiles([]);
                        setShowForm(false);
                      }}
                    >
                      ↑ Upload Another Document
                    </button>
                    <button className={styles.successCloseBtn} onClick={() => props.onClose && props.onClose()}>✕ Close</button>
                  </div>
                </div>
              </div>
              <IKShellFooter compact={isReplaceMode} className={styles.uploadFlowFooter} onBackHome={handleUploadHomeOpen} />
            </div>
          </div>
        )}

        {/* Hidden inputs */}
        <input
          ref={fileInputRef}
          type="file"
          multiple={!isReplaceMode}
          style={{ display: 'none' }}
          accept={isReplaceMode ? REPLACE_UPLOAD_ACCEPT : DOCUMENT_UPLOAD_ACCEPT}
          onChange={(e) => onFileSelected(e.target.files)}
        />
        <input
          ref={mediaInputRef}
          type="file"
          style={{ display: 'none' }}
          accept={MEDIA_UPLOAD_ACCEPT}
          onChange={(e) => handleMediaUpload(e.target.files)}
        />
      </div>
    </div>
  );
};

export default FileUpload;
