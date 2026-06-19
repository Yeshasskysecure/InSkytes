import * as React from 'react';
import { CommandBar, Icon, ICommandBarItemProps, initializeIcons, Spinner, SpinnerSize } from '@fluentui/react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { IDocumentDetailPageProps } from './IDocumentDetailPageProps';
import { MetadataForm } from '../../components/metadata/MetadataForm/MetadataForm';
import { FileUpload } from '../../components/upload/FileUpload/FileUpload';
import { IKShellFooter, IKShellHeader } from '../../components/shell/IKShellChrome';
import { BUDeptDisplay } from '../../components/document/BUDeptDisplay';
import styles from './DocumentDetailPage.module.scss';
import {
  getDocumentMetrics,
  getUserInteractionState,
  createDocumentMetricsRow,
  incrementMetricWithRetry,
  recordUserEvent,
  toggleUserInteraction
} from '../../services/kmMetricsService';
import { NAV_PATHS, pushPageUrl, updateItemUrlFieldRaw } from '../../services/permalinkService';
import {
  fetchCurrentVersionDetail,
  fetchVersionDetail,
  getVersionHistory,
  IVersionDetail,
  IVersionEntry,
  updateWithoutVersion
} from '../../services/kmVersionControl';
import {
  buildKMDataHubItemQuery,
  fetchKMDataHubReadFieldMap,
  formatDocumentPublishedDate,
  resolveDocumentAuthor,
  resolveDocumentPublishedValue
} from '../../utils/documentMetadata';
import { normalizeSharePointUrl, makeServerRelativeUrl } from '../../utils/urlUtils';
import { downloadSharePointFile } from '../../utils/fileDownload';
import {
  fetchAllTaxonomyOptions,
  ITaxonomyFieldOptions,
  ITaxonomyTerm,
  matchTaxonomyTerm,
  TAXONOMY_FIELD_CONFIGS
} from '../../services/TaxonomyService';
import { getKmsUsers, IKmsUser } from '../../services/KmsUsersService';
import { emitDocumentDataChanged, subscribeToDocumentDataChanged } from '../../services/documentChangeEvents';
import { KnowledgeSearchApiClient } from '../../services/KnowledgeSearchApiClient';
import {
  buildBuDepartmentFieldUpdates,
  KM_DATA_HUB_BU_FIELD_INTERNAL_NAME,
  KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME
} from '../../utils/buDepartmentSelections';
import { COLUMN_NAMES, LIBRARY_NAMES, LIST_NAMES, PAGE_URLS, SITE_PAGES, SITE_URL } from '../../config/appConfig';
import { touchRecentlyPublishedCacheForDocument } from '../../services/recentlyPublishedCache';

const PdfViewer = React.lazy(() =>
  import('../../components/pdfViewer/PdfViewer').then((module) => ({
    default: module.PdfViewer
  }))
);

initializeIcons(undefined, { disableWarnings: true });

const ImageViewer = React.lazy(() =>
  import('../../components/imageViewer/ImageViewer').then((module) => ({
    default: module.ImageViewer
  }))
);

// Session-level caches to track which lists have been verified and store user titles
const VERIFIED_LISTS_DETAIL_PAGE: Set<string> = new Set();
const AUTHOR_CACHE: Record<number, string> = {};

const makeIconOnlyToolbarItem = (
  item: ICommandBarItemProps,
  label: string
): ICommandBarItemProps => ({
  ...item,
  text: '',
  ariaLabel: label,
  title: label,
  iconOnly: true
});

interface DocumentDetail {
  id: number;
  name: string;
  fileName: string;
  abstract: string;
  fileType: string;
  author: string;
  createdBy: string;
  createdByEmail?: string;
  authorUserIds: number[];
  authorEmails: string[];
  creatorUserId: number;
  date: string;
  createdDate: string;
  fileSize: string;
  serverRelativeUrl: string;
  fileRef: string;
  fileUniqueId?: string;
  status?: string;
  contentRefreshDate?: string;
  views?: number;
  likes?: number;
  comments?: number;
  downloads?: number;
  follow?: boolean;
  share?: boolean;
  bookmark?: boolean;
  sensitiveTerms?: string[];
  reviewerComments?: string;
  projectId?: string;
  versionFileName?: string;
  versionFileType?: string;
  modifiedBy?: {
    title: string;
    email: string;
    id: number;
  };
  docIcon?: string;
}

type TDetailVersionPreview = {
  url: string;
  kind: 'office' | 'iframe' | 'image' | 'video' | 'audio' | 'unsupported';
  isBlob: boolean;
  message?: string;
};

interface IKMDataHubEditFieldMap {
  status: string;
  published: string;
  title: string;
  description: string;
  author: string;
  bu: string;
  department: string;
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
  reviewerComments?: string;
  reviewerCorrections?: string;
  contentRefreshDate?: string;
  views?: string;
  likes?: string;
  comments?: string;
  downloads?: string;
  follow?: string;
  share?: string;
  bookmark?: string;
  projectId?: string;
  versionFileName?: string;
  versionFileType?: string;
  edited?: string;
  editedBy?: string;
  modifiedBy?: string;
  docIcon?: string;
  fileLeafRef?: string;
}

interface IMetadataSnippetRow {
  label: string;
  value: string;
}

const EMPTY_TAXONOMY_OPTIONS: ITaxonomyFieldOptions = {
  buDepartment: [],
  documentType: [],
  client: [],
  geography: [],
  diseaseArea: [],
  therapyArea: []
};

const createInitialEditValues = (documentDetail: DocumentDetail | null): Record<string, any> | null => {
  if (!documentDetail) {
    return null;
  }

  return {
    title: documentDetail.name || '',
    description: documentDetail.abstract || '',
    selectedAuthorUpns: [],
    buDepartmentTerms: [],
    buDepartmentTerm: null,
    documentTypeTerms: [],
    documentTypeTerm: null,
    clientTerms: [],
    clientTerm: null,
    geographyTerms: [],
    geographyTerm: null,
    diseaseAreaTerms: [],
    diseaseAreaTerm: null,
    therapyAreaTerms: [],
    therapyAreaTerm: null,
    sensitiveTerms: '',
    reviewerComments: '',
    reviewerCorrections: '',
    status: documentDetail.status || 'Under Review'
  };
};

const sanitizeDownloadFileName = (value: string): string => (
  value
    .replace(/[<>:"/\\|?*]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '')
);

const getFileExtension = (fileName: string, fileType: string): string => {
  const fileNameMatch = /\.([^./\\]+)$/.exec(fileName || '');
  if (fileNameMatch?.[1]) {
    return fileNameMatch[1];
  }

  return (fileType || '').replace(/^\./, '').trim();
};

const buildTitleDownloadFileName = (title: string, fileName: string, fileType: string): string => {
  const extension = getFileExtension(fileName, fileType);
  const sanitizedFileName = sanitizeDownloadFileName(fileName || title || 'Document') || 'Document';

  if (!extension) {
    return sanitizedFileName;
  }

  const extensionPattern = new RegExp(`\\.${extension.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, 'i');
  if (extensionPattern.test(sanitizedFileName)) {
    return sanitizedFileName;
  }

  return `${sanitizedFileName}.${extension}`;
};

interface Comment {
  id: number;
  userName: string;
  comment: string;
  created: Date;
  userId: number;
}

type PdfScrollMode = 'page' | 'single' | 'vertical' | 'horizontal' | 'wrapped';
type PdfSidebarPanel = 'thumbnails' | 'bookmarks' | 'attachments' | null;
type DocumentToggleAction = 'follow' | 'bookmark' | 'flag';
type DocumentCountAction = DocumentToggleAction | 'share';
type ShareTooltipKey = 'close' | 'copy' | 'send';
type SharePerson = {
  name: string;
  email: string;
  accountName: string;
  subtitle?: string;
  entityType?: string;
};
type ShareSiteUser = {
  Id?: number;
  Title?: string;
  Email?: string;
  LoginName?: string;
  PrincipalType?: number;
  IsHiddenInUI?: boolean;
};
type OfficePrintPreviewPage = {
  kind: 'image' | 'iframe';
  src: string;
};
type ViewerFeatureKey =
  | 'thumbnails'
  | 'zoom'
  | 'pageNavigation'
  | 'windowMode'
  | 'rotate'
  | 'layout';

type ViewerAdapterKind = 'pdf' | 'office' | 'image' | 'text' | 'video' | 'audio' | 'unsupported';

interface IViewerToolbarCapabilities {
  canToggleThumbnails: boolean;
  canZoom: boolean;
  canNavigatePages: boolean;
  canWindowMode: boolean;
  canRotate: boolean;
  canChangeLayout: boolean;
  pageUnitLabel: string;
  pageUnitLabelPlural: string;
}

interface IViewerAdapter {
  kind: ViewerAdapterKind;
  capabilities: IViewerToolbarCapabilities;
}
const VIDEO_FILE_TYPES = ['mp4', 'webm', 'ogg', 'mov', 'avi', 'm4v', 'wmv', 'mkv'];
const AUDIO_FILE_TYPES = ['mp3', 'wav', 'ogg', 'm4a', 'aac', 'flac', 'wma', 'opus'];
const IMAGE_FILE_TYPES = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg', 'avif', 'tif', 'tiff'];
const OFFICE_FILE_TYPES = ['docx', 'doc', 'docm', 'pptx', 'ppt', 'pptm', 'ppsx', 'potx', 'xlsx', 'xls', 'xlsm', 'xlsb', 'xltx', 'xltm'];
const TEXT_FILE_TYPES = ['txt', 'text', 'csv', 'tsv', 'md', 'markdown', 'log', 'json', 'xml', 'yml', 'yaml', 'ini', 'cfg', 'sql', 'rtf'];
const PRESENTATION_FILE_TYPES = ['pptx', 'ppt', 'pptm', 'ppsx', 'pps', 'potx', 'potm'];
const SPREADSHEET_FILE_TYPES = ['xls', 'xlsx', 'xlsm', 'xlsb', 'xltx', 'xltm'];
const CSV_SPREADSHEET_FILE_TYPES = ['csv'];
const WORD_PROCESSING_FILE_TYPES = ['doc', 'docx', 'docm'];
const ZOOM_PRESETS = [25, 50, 75, 100, 125, 150, 175, 200, 225, 250];
const PDF_THUMBNAIL_SCALE = 0.25;
const PDF_VIEWER_FRAGMENT_OPTIONS: Record<string, string> = {
  toolbar: '0',
  navpanes: '0',
  scrollbar: '1',
  zoom: 'page-fit'
};
const SHOW_DOCUMENT_FOLLOW_ACTION = false;
const DEFAULT_LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;
const FLAG_REVIEW_MAX_LENGTH = 2000;
const REVIEW_FLAG_REASONS = [
  'Redundant information',
  'Incorrect information',
  'Outdated information',
  'Attachment issue',
  'Other'
] as const;

const normalizeGuid = (value?: string): string => (value || '').replace(/[{}]/g, '').trim().toLowerCase();

const isSharedUseLockError = (error: unknown): boolean => {
  const message = error instanceof Error ? error.message : String(error || '');
  const normalizedMessage = message.toLowerCase();

  return normalizedMessage.indexOf('locked for shared use') !== -1 ||
    (normalizedMessage.indexOf('423') !== -1 && normalizedMessage.indexOf('locked') !== -1);
};

const getMultiTaxonomyValue = (terms: ITaxonomyTerm[] | null | undefined): string =>
  (terms || [])
    .filter((term) => !!term?.id)
    .map((term) => `${term.label}|${term.id}`)
    .join(';');

const getMultiTaxonomyNoteValue = (terms: ITaxonomyTerm[] | null | undefined): string =>
  (terms || [])
    .filter((term) => !!term?.id)
    .map((term) => `-1;#${term.label}|${term.id}`)
    .join(';#');

const uniqueTerms = (terms: ITaxonomyTerm[]): ITaxonomyTerm[] => {
  const seen: Record<string, boolean> = {};
  return terms.filter((term) => {
    if (!term?.id || seen[term.id]) {
      return false;
    }

    seen[term.id] = true;
    return true;
  });
};

const extractTermsFromValue = (rawValue: any, allTerms: ITaxonomyTerm[]): ITaxonomyTerm[] => {
  if (!rawValue) {
    return [];
  }

  if (Array.isArray(rawValue)) {
    const arrayTerms = rawValue
      .map((entry) => {
        if (entry?.TermGuid) {
          return allTerms.find((term) => normalizeGuid(term.id) === normalizeGuid(entry.TermGuid)) || null;
        }
        if (entry?.Label) {
          return matchTaxonomyTerm(String(entry.Label), allTerms);
        }
        if (typeof entry === 'string') {
          return matchTaxonomyTerm(entry, allTerms);
        }
        return null;
      })
      .filter((term): term is ITaxonomyTerm => !!term);

    return uniqueTerms(arrayTerms);
  }

  if (typeof rawValue === 'object') {
    const nestedValues = rawValue.results || rawValue.value || rawValue.Values;
    if (Array.isArray(nestedValues)) {
      return extractTermsFromValue(nestedValues, allTerms);
    }

    if (rawValue.TermGuid) {
      const matchedTerm = allTerms.find((term) => normalizeGuid(term.id) === normalizeGuid(rawValue.TermGuid));
      return matchedTerm ? [matchedTerm] : [];
    }

    if (rawValue.Label) {
      const matchedTerm = matchTaxonomyTerm(String(rawValue.Label), allTerms);
      return matchedTerm ? [matchedTerm] : [];
    }
  }

  if (typeof rawValue !== 'string') {
    return [];
  }

  const segments = rawValue
    .split(/;#|;|,/)
    .map((segment) => segment.trim())
    .filter((segment) => segment.length > 0);

  const resolvedTerms = segments
    .map((segment) => {
      const guidCandidate = segment.includes('|') ? segment.split('|').pop() || '' : '';
      if (guidCandidate) {
        const matchedById = allTerms.find((term) => normalizeGuid(term.id) === normalizeGuid(guidCandidate));
        if (matchedById) {
          return matchedById;
        }
      }

      const labelCandidate = segment.includes('|') ? segment.split('|')[0] : segment;
      return matchTaxonomyTerm(labelCandidate, allTerms);
    })
    .filter((term): term is ITaxonomyTerm => !!term);

  return uniqueTerms(resolvedTerms);
};

const extractPersonEntries = (value: any): any[] => {
  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.reduce<any[]>((entries, entry) => entries.concat(extractPersonEntries(entry)), []);
  }

  const nestedValues = value.results || value.value || value.Values;
  if (Array.isArray(nestedValues)) {
    return extractPersonEntries(nestedValues);
  }

  return [value];
};

const uniqueStrings = (values: string[]): string[] => {
  const seen: Record<string, boolean> = {};
  return values.filter((value) => {
    const normalized = String(value || '').trim().toLowerCase();
    if (!normalized || seen[normalized]) {
      return false;
    }

    seen[normalized] = true;
    return true;
  });
};

const resolveBuDepartmentTermsFromSeparateFields = (
  item: Record<string, any>,
  textValues: Record<string, any>,
  allTerms: ITaxonomyTerm[]
): ITaxonomyTerm[] => {
  const resolveTerms = (internalName: string): ITaxonomyTerm[] => {
    const itemTerms = extractTermsFromValue(item[internalName], allTerms);
    return itemTerms.length > 0
      ? uniqueTerms(itemTerms)
      : uniqueTerms(extractTermsFromValue(textValues[internalName], allTerms));
  };

  return uniqueTerms([
    ...resolveTerms(KM_DATA_HUB_BU_FIELD_INTERNAL_NAME),
    ...resolveTerms(KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME)
  ]);
};

const buildKMDataHubTaxonomyFormValues = (
  metadata: Record<string, any>,
  fieldMap: IKMDataHubEditFieldMap
): Array<{ FieldName: string; FieldValue: string }> => {
  const selectedAuthorUpns = Array.isArray(metadata.selectedAuthorUpns)
    ? metadata.selectedAuthorUpns.filter(Boolean)
    : metadata.selectedAuthorUpn
      ? [metadata.selectedAuthorUpn].filter(Boolean)
      : [];

  const buildClaimsValues = (upns: string[]): string =>
    JSON.stringify(
      upns.map((upn) => ({
        Key: `i:0#.f|membership|${(upn || '').trim()}`
      }))
    );

  return [
    { FieldName: fieldMap.author, FieldValue: selectedAuthorUpns.length > 0 ? buildClaimsValues(selectedAuthorUpns) : '' },
    { FieldName: fieldMap.bu, FieldValue: (metadata.selectedBUs || []).map((term: any) => `${term.name}|${term.id}`).join(';') },
    {
      FieldName: fieldMap.department,
      FieldValue: ((metadata.selectedDepts || []).filter((term: any, _index: number, array: any[]) => {
        const selectedDeptIds = new Set(array.map((item: any) => item.id));
        const hasSelectedChild = (term.children || []).some((child: any) => selectedDeptIds.has(child.id));
        return !hasSelectedChild;
      })).map((term: any) => `${term.name}|${term.id}`).join(';')
    },
    ...buildBuDepartmentFieldUpdates(metadata.selectedBUs || [], metadata.selectedDepts || []),
    { FieldName: fieldMap.documentType, FieldValue: getMultiTaxonomyValue(metadata.documentTypeTerms) },
    { FieldName: fieldMap.client, FieldValue: getMultiTaxonomyValue(metadata.clientTerms) },
    { FieldName: fieldMap.geography, FieldValue: getMultiTaxonomyValue(metadata.geographyTerms) },
    { FieldName: fieldMap.diseaseArea, FieldValue: getMultiTaxonomyValue(metadata.diseaseAreaTerms) },
    { FieldName: fieldMap.therapyArea, FieldValue: getMultiTaxonomyValue(metadata.therapyAreaTerms) }
  ].filter((field) => !!field.FieldName && !!field.FieldValue);
};

const formatTaxonomyDisplayValue = (value: unknown): string => {
  if (!value) {
    return '';
  }

  const stringifyEntry = (entry: any): string => {
    if (!entry) {
      return '';
    }

    if (typeof entry === 'string') {
      return entry;
    }

    return entry.Label || entry.Title || entry.TermLabel || entry.Name || '';
  };

  const rawValue = Array.isArray(value)
    ? value.map(stringifyEntry).filter(Boolean).join('; ')
    : stringifyEntry(value);

  return rawValue
    .split(/;#|;/)
    .map((segment) => segment.trim())
    .filter((segment) => segment && segment !== '-1')
    .map((segment) => {
      const withoutLookupPrefix = segment.includes('#') ? segment.split('#').pop() || segment : segment;
      return withoutLookupPrefix.includes('|') ? withoutLookupPrefix.split('|')[0].trim() : withoutLookupPrefix.trim();
    })
    .filter((segment, index, allSegments) => segment && allSegments.indexOf(segment) === index)
    .join(', ');
};

const formatAuthorDisplayValue = (value: string): string => {
  const normalizedValue = (value || '').trim();
  if (!normalizedValue) {
    return 'Internal';
  }

  return normalizedValue
    .split(/;#|;/)
    .map((segment) => segment.trim())
    .filter((segment) => segment && !/^-?\d+$/.test(segment))
    .filter((segment, index, allSegments) => allSegments.indexOf(segment) === index)
    .join(', ') || normalizedValue;
};

const splitDocumentInfoValues = (value: string): string[] =>
  (value || '')
    .split(/\s*,\s*/)
    .map((entry) => entry.trim())
    .filter(Boolean);

const normalizeDescriptionText = (value: unknown): string => {
  const text = formatTaxonomyDisplayValue(value);
  return !text || text === '-' ? '' : text;
};

const normalizeDocumentNumberField = (value: unknown): number | undefined => {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const numericValue = Number(formatTaxonomyDisplayValue(value));
  return Number.isFinite(numericValue) ? numericValue : undefined;
};

const normalizeDocumentBooleanField = (value: unknown): boolean | undefined => {
  if (typeof value === 'boolean') {
    return value;
  }

  const normalizedValue = formatTaxonomyDisplayValue(value).toLowerCase();
  if (['true', 'yes', '1'].indexOf(normalizedValue) !== -1) {
    return true;
  }
  if (['false', 'no', '0'].indexOf(normalizedValue) !== -1) {
    return false;
  }
  return undefined;
};

const normalizeDocumentStringArrayField = (value: unknown): string[] =>
  formatTaxonomyDisplayValue(value)
    .split(/[;,\n]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);

const getStatusBadgeClassName = (status?: string): string => {
  const normalizedStatus = (status || '').trim().toLowerCase();

  if (normalizedStatus === 'active' || normalizedStatus === 'approved') {
    return `${styles.badge} ${styles.badgeActive}`;
  }

  if (normalizedStatus === 'under review') {
    return `${styles.badge} ${styles.badgeReview}`;
  }

  if (normalizedStatus === 'rejected') {
    return `${styles.badge} ${styles.badgeRejected}`;
  }

  return `${styles.badge} ${styles.badgeDefault}`;
};

const buildPreviewImageUrl = (webUrl: string, serverRelativeUrl: string): string => {
  const fileUrl = `${webUrl}${serverRelativeUrl}`;
  return `${webUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(fileUrl)}&resolution=2`;
};

const buildGuidPreviewImageUrl = (
  webUrl: string,
  siteId: string,
  webId: string,
  fileUniqueId: string
): string => `${webUrl}/_layouts/15/getpreview.ashx?guidSite=${encodeURIComponent(siteId)}&guidWeb=${encodeURIComponent(webId)}&guidFile=${encodeURIComponent(fileUniqueId)}&resolution=2`;

const buildIndexedPreviewImageUrl = (
  webUrl: string,
  serverRelativeUrl: string,
  pageNumber: number,
  fileUniqueId?: string,
  siteId?: string,
  webId?: string
): string => {
  const index = Math.max(pageNumber, 1);

  if (fileUniqueId && siteId && webId) {
    return `${webUrl}/_layouts/15/getpreview.ashx?guidSite=${encodeURIComponent(siteId)}&guidWeb=${encodeURIComponent(webId)}&guidFile=${encodeURIComponent(fileUniqueId)}&resolution=2&index=${index}&force=1`;
  }

  const fileUrl = `${webUrl}${serverRelativeUrl}`;
  return `${webUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(fileUrl)}&resolution=2&index=${index}&force=1`;
};

const buildRawIndexedPreviewImageUrl = (
  webUrl: string,
  serverRelativeUrl: string,
  rawIndex: number,
  fileUniqueId?: string,
  siteId?: string,
  webId?: string
): string => {
  const index = Math.max(rawIndex, 0);

  if (fileUniqueId && siteId && webId) {
    return `${webUrl}/_layouts/15/getpreview.ashx?guidSite=${encodeURIComponent(siteId)}&guidWeb=${encodeURIComponent(webId)}&guidFile=${encodeURIComponent(fileUniqueId)}&resolution=2&index=${index}&force=1`;
  }

  const fileUrl = `${webUrl}${serverRelativeUrl}`;
  return `${webUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(fileUrl)}&resolution=2&index=${index}&force=1`;
};

const buildSlideImagePreviewUrl = (
  webUrl: string,
  serverRelativeUrl: string,
  pageNumber: number,
  fileUniqueId?: string,
  siteId?: string,
  webId?: string
): string => buildIndexedPreviewImageUrl(webUrl, serverRelativeUrl, Math.max(pageNumber, 1), fileUniqueId, siteId, webId);

const buildSlideImagePreviewCandidates = (
  webUrl: string,
  serverRelativeUrl: string,
  pageNumber: number,
  fileUniqueId?: string,
  siteId?: string,
  webId?: string
): string[] => {
  const safePageNumber = Math.max(pageNumber, 1);
  const rawIndex = Math.max(safePageNumber - 1, 0);
  const candidates = [
    buildIndexedPreviewImageUrl(webUrl, serverRelativeUrl, safePageNumber, fileUniqueId, siteId, webId),
    buildRawIndexedPreviewImageUrl(webUrl, serverRelativeUrl, safePageNumber, fileUniqueId, siteId, webId),
    buildRawIndexedPreviewImageUrl(webUrl, serverRelativeUrl, rawIndex, fileUniqueId, siteId, webId),
    safePageNumber === 1 ? buildPreviewImageUrl(webUrl, serverRelativeUrl) : ''
  ];

  return Array.from(new Set(candidates.filter(Boolean)));
};

const buildWopiFramePreviewUrl = (webUrl: string, serverRelativeUrl: string): string =>
  `${webUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(serverRelativeUrl)}&action=imagepreview`;

const buildWopiFrameImagePreviewUrl = (
  webUrl: string,
  serverRelativeUrl: string,
  fileExtension: string,
  pageNumber: number
): string => {
  const normalizedExtension = fileExtension.toLowerCase();
  const previewUrl = buildWopiFramePreviewUrl(webUrl, serverRelativeUrl);
  const safePageNumber = Math.max(pageNumber, 1);

  if (PRESENTATION_FILE_TYPES.includes(normalizedExtension)) {
    return `${previewUrl}&wdSlideIndex=${safePageNumber}`;
  }

  if (SPREADSHEET_FILE_TYPES.includes(normalizedExtension)) {
    return `${previewUrl}&wdSheetIndex=${safePageNumber}`;
  }

  if (WORD_PROCESSING_FILE_TYPES.includes(normalizedExtension)) {
    return `${previewUrl}&wdPageIndex=${safePageNumber}`;
  }

  return previewUrl;
};

const buildOfficePreviewUrl = (
  webUrl: string,
  serverRelativeUrl: string,
  fileExtension: string,
  pageNumber: number = 1,
  includePageReference: boolean = true
): string => {
  const normalizedExtension = fileExtension.toLowerCase();
  const encodedServerUrl = encodeURIComponent(serverRelativeUrl);
  const officeAction =
    PRESENTATION_FILE_TYPES.includes(normalizedExtension)
      ? 'interactivepreview'
      : SPREADSHEET_FILE_TYPES.includes(normalizedExtension)
        ? 'embedview'
        : 'interactivepreview';

  const previewUrl = `${webUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodedServerUrl}&action=${officeAction}`;

  if (!includePageReference) {
    return previewUrl;
  }

  if (PRESENTATION_FILE_TYPES.includes(normalizedExtension)) {
    return `${previewUrl}&wdSlideIndex=${Math.max(pageNumber, 1)}`;
  }

  if (SPREADSHEET_FILE_TYPES.includes(normalizedExtension)) {
    return `${previewUrl}&wdSheetIndex=${Math.max(pageNumber, 1)}&ActiveCell=A1&wdActiveCell=A1&wdAllowInteractivity=True`;
  }

  if (WORD_PROCESSING_FILE_TYPES.includes(normalizedExtension)) {
    return `${previewUrl}&wdPageIndex=${Math.max(pageNumber, 1)}`;
  }

  return previewUrl;
};

const buildOfficePrintUrl = (
  webUrl: string,
  serverRelativeUrl: string,
  fileExtension: string
): string => {
  const normalizedExtension = fileExtension.toLowerCase();
  const encodedServerUrl = encodeURIComponent(serverRelativeUrl);
  const officeAction =
    ['xls', 'xlsx', 'xlsm', 'xlsb', 'xltx', 'xltm'].includes(normalizedExtension)
      ? 'view'
      : 'default';

  return `${webUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodedServerUrl}&action=${officeAction}`;
};

const normalizeServerRelativeUrl = (serverRelativeUrl: string): string =>
  serverRelativeUrl.startsWith('/') ? serverRelativeUrl : `/${serverRelativeUrl}`;

const escapeODataPathValue = (serverRelativeUrl: string): string =>
  normalizeServerRelativeUrl(serverRelativeUrl).replace(/'/g, "''");

const buildFileApiUrl = (
  webUrl: string,
  serverRelativeUrl: string,
  suffix: string = '/$value'
): string => `${webUrl}/_api/web/GetFileByServerRelativeUrl('${escapeODataPathValue(serverRelativeUrl)}')${suffix}`;

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

const buildPersonClaimsFieldValue = (upn: string): string =>
  JSON.stringify([{ Key: `i:0#.f|membership|${(upn || '').trim()}` }]);

const appendFormValues = (
  baseFormValues: Array<{ FieldName: string; FieldValue: string }>,
  extraFormValues: Array<{ FieldName: string; FieldValue: string }>
): Array<{ FieldName: string; FieldValue: string }> => {
  const extraFieldNames = new Set(extraFormValues.map((entry) => entry.FieldName.toLowerCase()));

  return [
    ...baseFormValues.filter((entry) => !extraFieldNames.has(entry.FieldName.toLowerCase())),
    ...extraFormValues
  ];
};

const isPublishedStatus = (status?: string): boolean => {
  const normalizedStatus = (status || '').trim().toLowerCase();
  return normalizedStatus === 'active' || normalizedStatus === 'approved';
};

const shouldClearPublishedDateForStatus = (status?: string): boolean => {
  const normalizedStatus = (status || '').trim().toLowerCase();
  return [
    'under review',
    'rejected',
    'reject',
    'archived',
    'archive',
    'duplicate',
    'unpublish',
    'unpublished',
    'draft'
  ].indexOf(normalizedStatus) !== -1;
};

const appendPublishedDateUpdate = (
  body: Record<string, any>,
  fieldMap: IKMDataHubEditFieldMap,
  nextStatus?: string,
  previousStatus?: string
): void => {
  const hasNextStatus = String(nextStatus || '').trim().length > 0;
  if (!fieldMap.published || !hasNextStatus) {
    return;
  }

  const wasPublished = isPublishedStatus(previousStatus);
  const isNowPublished = isPublishedStatus(nextStatus);

  if (isNowPublished && !wasPublished) {
    body[fieldMap.published] = formatSharePointDateFieldValue();
  } else if (!isNowPublished && shouldClearPublishedDateForStatus(nextStatus)) {
    body[fieldMap.published] = '';
  }
};

const getDetailVersionPreviewKind = (fileType: string): TDetailVersionPreview['kind'] => {
  const ext = (fileType || '').toLowerCase();
  if (['pptx', 'ppt', 'docx', 'doc', 'xlsx', 'xls', 'xlsm', 'xlsb'].indexOf(ext) !== -1) {
    return 'office';
  }
  if (ext === 'pdf' || ['txt', 'csv', 'mht', 'mhtml'].indexOf(ext) !== -1) {
    return 'iframe';
  }
  if (['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'].indexOf(ext) !== -1) {
    return 'image';
  }
  if (['mp4', 'mov', 'webm', 'mkv', 'avi'].indexOf(ext) !== -1) {
    return 'video';
  }
  if (['mp3', 'wav', 'm4a', 'aac', 'flac', 'ogg'].indexOf(ext) !== -1) {
    return 'audio';
  }
  return 'unsupported';
};

const getDetailVersionMimeType = (fileType: string): string => {
  const ext = (fileType || '').toLowerCase();
  switch (ext) {
    case 'pptx': return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    case 'ppt': return 'application/vnd.ms-powerpoint';
    case 'docx': return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    case 'doc': return 'application/msword';
    case 'xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    case 'xls': return 'application/vnd.ms-excel';
    case 'xlsm': return 'application/vnd.ms-excel.sheet.macroEnabled.12';
    case 'xlsb': return 'application/vnd.ms-excel.sheet.binary.macroEnabled.12';
    case 'pdf': return 'application/pdf';
    case 'txt': return 'text/plain';
    case 'csv': return 'text/csv';
    case 'mht':
    case 'mhtml': return 'message/rfc822';
    case 'png': return 'image/png';
    case 'jpg':
    case 'jpeg': return 'image/jpeg';
    case 'gif': return 'image/gif';
    case 'bmp': return 'image/bmp';
    case 'webp': return 'image/webp';
    case 'svg': return 'image/svg+xml';
    case 'mp4': return 'video/mp4';
    case 'mov': return 'video/quicktime';
    case 'webm': return 'video/webm';
    case 'mkv': return 'video/x-matroska';
    case 'avi': return 'video/x-msvideo';
    case 'mp3': return 'audio/mpeg';
    case 'wav': return 'audio/wav';
    case 'm4a': return 'audio/mp4';
    case 'aac': return 'audio/aac';
    case 'flac': return 'audio/flac';
    case 'ogg': return 'audio/ogg';
    default: return 'application/octet-stream';
  }
};

const buildAbsoluteFileUrl = (webUrl: string, serverRelativeUrl: string): string => {
  const normalizedServerRelativeUrl = normalizeServerRelativeUrl(serverRelativeUrl);
  try {
    return `${new URL(webUrl).origin}${normalizedServerRelativeUrl}`;
  } catch {
    return `${window.location.origin}${normalizedServerRelativeUrl}`;
  }
};

const appendCacheBustParameter = (url: string, cacheBustToken?: string): string => {
  if (!cacheBustToken || !url || url.startsWith('blob:') || url === 'loaded') {
    return url;
  }

  const separator = url.indexOf('?') >= 0 ? '&' : '?';
  return `${url}${separator}v=${encodeURIComponent(cacheBustToken)}`;
};

const buildGraphSharingToken = (absoluteUrl: string): string => {
  const bytes = new TextEncoder().encode(absoluteUrl);
  let binaryValue = '';
  bytes.forEach((byte) => {
    binaryValue += String.fromCharCode(byte);
  });

  return `u!${btoa(binaryValue).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '')}`;
};

const getPowerPointSlideCountFromZip = async (fileBuffer: ArrayBuffer): Promise<number> => {
  const JSZipModule = await import('jszip');
  const JSZip = (JSZipModule as any).default || JSZipModule;
  const zip = await JSZip.loadAsync(fileBuffer);
  const slideFiles = Object.keys(zip.files)
    .filter((relativePath) => /^ppt\/slides\/slide\d+\.xml$/i.test(relativePath))
    .sort((leftPath, rightPath) => {
      const leftNumber = parseInt(leftPath.match(/slide(\d+)\.xml/i)?.[1] || '0', 10);
      const rightNumber = parseInt(rightPath.match(/slide(\d+)\.xml/i)?.[1] || '0', 10);
      return leftNumber - rightNumber;
    });

  const appXml = await zip.file('docProps/app.xml')?.async('string');
  const slidesMatch = appXml?.match(/<Slides>(\d+)<\/Slides>/i);
  return Math.max(slideFiles.length, Number(slidesMatch?.[1] || 0), 1);
};

const fetchFileBlobUrl = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  serverRelativeUrl: string,
  fallbackUrl: string,
  mimeType?: string
): Promise<string> => {
  try {
    const response = await spHttpClient.get(
      buildFileApiUrl(webUrl, serverRelativeUrl),
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return fallbackUrl;
    }

    const blob = await response.blob();
    return window.URL.createObjectURL(
      mimeType ? new Blob([blob], { type: mimeType }) : blob
    );
  } catch {
    return fallbackUrl;
  }
};

const fetchFileBytes = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  serverRelativeUrl: string
): Promise<Uint8Array | null> => {
  try {
    const response = await spHttpClient.get(
      buildFileApiUrl(webUrl, serverRelativeUrl),
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return null;
    }

    return new Uint8Array(await response.arrayBuffer());
  } catch {
    return null;
  }
};

const loadPdfBytes = async (pdfUrl: string): Promise<ArrayBuffer> => {
  const response = await fetch(pdfUrl, {
    credentials: 'include',
    headers: {
      Accept: 'application/pdf'
    }
  });

  if (!response.ok) {
    throw new Error(`PDF fetch failed: ${response.status}`);
  }

  return response.arrayBuffer();
};

const loadPdfBytesFromCandidates = async (pdfUrls: string[]): Promise<ArrayBuffer> => {
  const errors: Error[] = [];

  for (const pdfUrl of pdfUrls) {
    try {
      const pdfBytes = await loadPdfBytes(pdfUrl);
      validatePdfBytes(pdfBytes);
      return pdfBytes;
    } catch (error) {
      errors.push(error as Error);
    }
  }

  throw errors[errors.length - 1] || new Error('PDF fetch failed');
};

const validatePdfBytes = (pdfBytes: ArrayBuffer): void => {
  const signature = new TextDecoder('ascii').decode(pdfBytes.slice(0, 5));

  if (signature !== '%PDF-') {
    throw new Error(`PDF fetch did not return a PDF document: ${signature || 'empty response'}`);
  }
};

const buildPrintPopupFeatures = (): string => {
  const popupWidth = Math.min(window.screen.availWidth - 80, 1400);
  const popupHeight = Math.min(window.screen.availHeight - 80, 980);
  const popupLeft = Math.max(Math.round((window.screen.availWidth - popupWidth) / 2), 0);
  const popupTop = Math.max(Math.round((window.screen.availHeight - popupHeight) / 2), 0);

  return [
    `width=${popupWidth}`,
    `height=${popupHeight}`,
    `left=${popupLeft}`,
    `top=${popupTop}`,
    'resizable=yes',
    'scrollbars=yes',
    'toolbar=no',
    'menubar=no',
    'location=yes',
    'status=no'
  ].join(',');
};

const escapeHtml = (value: string): string =>
  value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const buildPreviewCandidates = (
  webUrl: string,
  serverRelativeUrl: string,
  siteId: string,
  webId: string,
  fileUniqueId?: string
): string[] => {
  const candidates: string[] = [];

  if (fileUniqueId) {
    candidates.push(buildGuidPreviewImageUrl(webUrl, siteId, webId, fileUniqueId));
  }

  candidates.push(buildPreviewImageUrl(webUrl, serverRelativeUrl));
  candidates.push(buildWopiFramePreviewUrl(webUrl, serverRelativeUrl));

  return candidates;
};

const extractNumericIds = (value: any): number[] => {
  if (value === null || value === undefined) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.reduce<number[]>((entries, entry) => {
      entries.push(...extractNumericIds(entry));
      return entries;
    }, []);
  }

  if (typeof value === 'number') {
    return value > 0 ? [value] : [];
  }

  if (typeof value === 'string') {
    return value
      .split(/;#|[;,]/)
      .map((entry) => Number(entry.trim()))
      .filter((entry) => Number.isFinite(entry) && entry > 0);
  }

  if (typeof value === 'object') {
    if (Array.isArray(value.results)) {
      return extractNumericIds(value.results);
    }

    const candidateId = Number(value.Id || value.ID || value.id);
    return Number.isFinite(candidateId) && candidateId > 0 ? [candidateId] : [];
  }

  return [];
};

const normalizeIdentityValue = (value?: string): string => {
  const rawValue = String(value || '').trim().toLowerCase();
  if (!rawValue) {
    return '';
  }

  return rawValue.indexOf('|') >= 0
    ? rawValue.split('|').pop() || rawValue
    : rawValue;
};

const extractIdentityCandidates = (value: any): string[] => {
  if (value === null || value === undefined) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.reduce<string[]>((entries, entry) => {
      entries.push(...extractIdentityCandidates(entry));
      return entries;
    }, []);
  }

  if (typeof value === 'string') {
    return value
      .split(/;#|[;,]/)
      .map(normalizeIdentityValue)
      .filter((entry) => entry.indexOf('@') > 0);
  }

  if (typeof value === 'object') {
    if (Array.isArray(value.results)) {
      return extractIdentityCandidates(value.results);
    }

    return [
      value.EMail,
      value.Email,
      value.email,
      value.mail,
      value.UserPrincipalName,
      value.userPrincipalName,
      value.LoginName,
      value.Name,
      value.Key,
      value.upn
    ].map(normalizeIdentityValue).filter((entry) => entry.indexOf('@') > 0);
  }

  return [];
};

const uniqueIdentityCandidates = (values: string[]): string[] => Array.from(new Set(values.map(normalizeIdentityValue).filter(Boolean)));

const identityValuesMatch = (left?: string, right?: string): boolean => {
  const normalizedLeft = normalizeIdentityValue(left);
  const normalizedRight = normalizeIdentityValue(right);
  return !!normalizedLeft && !!normalizedRight && normalizedLeft === normalizedRight;
};

interface IPdfRenderedThumbnail {
  pageNumber: number;
  dataUrl: string;
}

interface IPdfCanvasThumbnailStripProps {
  pdfDocument: any;
  pageCount: number;
  activePageNumber: number;
  onSelectPage: (pageNumber: number) => void;
}

const PdfCanvasThumbnailStrip: React.FunctionComponent<IPdfCanvasThumbnailStripProps> = ({
  pdfDocument,
  pageCount,
  activePageNumber,
  onSelectPage
}) => {
  const [thumbnails, setThumbnails] = React.useState<IPdfRenderedThumbnail[]>([]);
  const [hasError, setHasError] = React.useState<boolean>(false);

  React.useEffect(() => {
    let isCancelled = false;
    const renderTasks: any[] = [];

    setThumbnails([]);
    setHasError(false);

    const renderThumbnails = async (): Promise<void> => {
      try {
        const totalPages = Math.max(Math.min(pageCount || pdfDocument.numPages || 1, pdfDocument.numPages || pageCount || 1), 1);
        const renderedThumbnails: IPdfRenderedThumbnail[] = [];

        for (let pageNo = 1; pageNo <= totalPages; pageNo++) {
          const page = await pdfDocument.getPage(pageNo);
          const viewport = page.getViewport({ scale: PDF_THUMBNAIL_SCALE });
          const canvas = document.createElement('canvas');
          const context = canvas.getContext('2d');

          if (!context) {
            throw new Error('Canvas context unavailable');
          }

          canvas.width = Math.ceil(viewport.width);
          canvas.height = Math.ceil(viewport.height);

          const renderTask = page.render({
            canvasContext: context,
            viewport
          });
          renderTasks.push(renderTask);
          await renderTask.promise;

          if (isCancelled) {
            return;
          }

          renderedThumbnails.push({
            pageNumber: pageNo,
            dataUrl: canvas.toDataURL('image/png')
          });
          setThumbnails([...renderedThumbnails]);
        }
      } catch (error) {
        if (!isCancelled) {
          console.warn('Unable to render PDF thumbnail:', error);
          setHasError(true);
        }
      }
    };

    renderThumbnails();

    return () => {
      isCancelled = true;
      renderTasks.forEach((renderTask) => renderTask?.cancel?.());
    };
  }, [pageCount, pdfDocument]);

  if (hasError) {
    return <div className={styles.pdfSidebarEmpty}>Failed to load page thumbnails.</div>;
  }

  if (!thumbnails.length) {
    return <div className={styles.pdfSidebarEmpty}>Loading pages...</div>;
  }

  return (
    <div className={styles.pdfThumbnailStrip}>
      {thumbnails.map((thumbnail) => (
        <button
          key={`thumb-${thumbnail.pageNumber}`}
          type="button"
          className={`${styles.pdfThumbnailButton} ${activePageNumber === thumbnail.pageNumber ? styles.pdfThumbnailButtonActive : ''}`}
          onClick={() => onSelectPage(thumbnail.pageNumber)}
        >
          <span className={styles.pdfThumbnailIndex}>{thumbnail.pageNumber}</span>
          <span className={styles.pdfThumbnailCardBody}>
            <span className={styles.pdfThumbnailPreview} aria-hidden="true">
              <img
                src={thumbnail.dataUrl}
                alt=""
                className={styles.pdfThumbnailImage}
                loading="lazy"
                decoding="async"
              />
            </span>
            <span className={styles.pdfThumbnailLabel}>Page {thumbnail.pageNumber}</span>
          </span>
        </button>
      ))}
    </div>
  );
};

interface IPdfImageFallbackThumbnailStripProps {
  pageCount: number;
  activePageNumber: number;
  onSelectPage: (pageNumber: number) => void;
  getThumbnailUrl: (pageNumber: number) => string | null;
  getLoadingMode: (pageNumber: number) => 'eager' | 'lazy';
  onThumbnailError: (pageNumber: number) => void;
  hasMoreCandidates: (pageNumber: number) => boolean;
}

const PdfImageFallbackThumbnailStrip: React.FunctionComponent<IPdfImageFallbackThumbnailStripProps> = ({
  pageCount,
  activePageNumber,
  onSelectPage,
  getThumbnailUrl,
  getLoadingMode,
  onThumbnailError,
  hasMoreCandidates
}) => (
  <div className={styles.pdfThumbnailStrip}>
    {Array.from({ length: pageCount }, (_, index) => index + 1).map((pageNumber) => {
      const thumbnailUrl = getThumbnailUrl(pageNumber);

      return (
        <button
          key={`pdf-fallback-thumb-${pageNumber}`}
          type="button"
          className={`${styles.pdfThumbnailButton} ${activePageNumber === pageNumber ? styles.pdfThumbnailButtonActive : ''}`}
          onClick={() => onSelectPage(pageNumber)}
        >
          <span className={styles.pdfThumbnailIndex}>{pageNumber}</span>
          <span className={styles.pdfThumbnailCardBody}>
            <span className={styles.pdfThumbnailPreview} aria-hidden="true">
              {thumbnailUrl ? (
                <img
                  src={thumbnailUrl}
                  alt=""
                  className={styles.pdfThumbnailImage}
                  loading={getLoadingMode(pageNumber)}
                  decoding="async"
                  onError={() => onThumbnailError(pageNumber)}
                />
              ) : hasMoreCandidates(pageNumber) ? (
                <span className={styles.pdfThumbnailPreviewInner} />
              ) : (
                <span className={styles.pdfSidebarEmpty}>Preview unavailable</span>
              )}
            </span>
            <span className={styles.pdfThumbnailLabel}>Page {pageNumber}</span>
          </span>
        </button>
      );
    })}
  </div>
);

export const DocumentDetailPage: React.FunctionComponent<IDocumentDetailPageProps> = (props) => {
  const knowledgeSearchApiClient = React.useMemo(
    () => new KnowledgeSearchApiClient(undefined, props.context),
    [props.context]
  );
  const [document, setDocument] = React.useState<DocumentDetail | null>(null);
  const [showShareDialog, setShowShareDialog] = React.useState<boolean>(false);
  const [sharePeopleQuery, setSharePeopleQuery] = React.useState('');
  const [sharePeopleResults, setSharePeopleResults] = React.useState<SharePerson[]>([]);
  const [shareSelectedPeople, setShareSelectedPeople] = React.useState<SharePerson[]>([]);
  const [shareMessage, setShareMessage] = React.useState('');
  const [shareSending, setShareSending] = React.useState(false);
  const [shareSent, setShareSent] = React.useState(false);
  const [shareCopied, setShareCopied] = React.useState(false);
  const [sharePeopleLoading, setSharePeopleLoading] = React.useState(false);
  const [shareValidationError, setShareValidationError] = React.useState('');
  const [shareFocusedField, setShareFocusedField] = React.useState<'people' | 'message' | null>(null);
  const [shareTooltip, setShareTooltip] = React.useState<ShareTooltipKey | null>(null);
  const [shareActivePersonIndex, setShareActivePersonIndex] = React.useState(0);
  const [shareHoveredPersonIndex, setShareHoveredPersonIndex] = React.useState<number | null>(null);
  const sharePeopleInputRef = React.useRef<HTMLInputElement>(null);
  const sharePeopleListRef = React.useRef<HTMLDivElement>(null);
  const shareMessageInputRef = React.useRef<HTMLTextAreaElement>(null);
  const shareSiteUsersRef = React.useRef<SharePerson[] | null>(null);
  const shareSearchRequestIdRef = React.useRef(0);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [fetchError, setFetchError] = React.useState<string | null>(null);
  const [previewUrl, setPreviewUrl] = React.useState<string>('');
  const [isPreviewLoading, setIsPreviewLoading] = React.useState<boolean>(false);
  const [isConvertedOfficePdfPreview, setIsConvertedOfficePdfPreview] = React.useState<boolean>(false);
  const [previewImageUrls, setPreviewImageUrls] = React.useState<string[]>([]);
  const [previewError, setPreviewError] = React.useState<string>('');
  const [likeCount, setLikeCount] = React.useState<number>(0);
  const [isLiked, setIsLiked] = React.useState<boolean>(false);
  const [commentCount, setCommentCount] = React.useState<number>(0);
  const [comments, setComments] = React.useState<Comment[]>([]);
  const [showComments, setShowComments] = React.useState<boolean>(false);
  const [newComment, setNewComment] = React.useState<string>('');
  const commentsPanelRef = React.useRef<HTMLDivElement>(null);
  const commentButtonRef = React.useRef<HTMLButtonElement>(null);
  const shouldScrollToCommentsRef = React.useRef<boolean>(false);
  const [currentUserId, setCurrentUserId] = React.useState<number>(0);
  const [currentUserName, setCurrentUserName] = React.useState<string>('');
  const [currentUserEmail, setCurrentUserEmail] = React.useState<string>('');
  const [isKmAdmin, setIsKmAdmin] = React.useState<boolean>(false);
  const [isApprover, setIsApprover] = React.useState<boolean>(false);
  const [isContributor, setIsContributor] = React.useState<boolean>(false);
  const [isLearner, setIsLearner] = React.useState<boolean>(false);
  const [isFollowing, setIsFollowing] = React.useState<boolean>(false);
  const [isBookmarked, setIsBookmarked] = React.useState<boolean>(false);
  const [isFlaggedForReview, setIsFlaggedForReview] = React.useState<boolean>(false);
  const [actionMessage, setActionMessage] = React.useState<string>('');
  const [isFlagReviewModalOpen, setIsFlagReviewModalOpen] = React.useState<boolean>(false);
  const [flagReviewReason, setFlagReviewReason] = React.useState<string>(REVIEW_FLAG_REASONS[REVIEW_FLAG_REASONS.length - 1]);
  const [flagReviewDetails, setFlagReviewDetails] = React.useState<string>('');
  const [flagReviewError, setFlagReviewError] = React.useState<string>('');
  const [isSubmittingFlagReview, setIsSubmittingFlagReview] = React.useState<boolean>(false);
  const [followCount, setFollowCount] = React.useState<number>(0);
  const [bookmarkCount, setBookmarkCount] = React.useState<number>(0);
  const [flagCount, setFlagCount] = React.useState<number>(0);
  const [shareCount, setShareCount] = React.useState<number>(0);
  const [viewCount, setViewCount] = React.useState<number>(0);
  const [downloadCount, setDownloadCount] = React.useState<number>(0);
  const [isDownloadStarting, setIsDownloadStarting] = React.useState<boolean>(false);
  const [metadataSnippetRows, setMetadataSnippetRows] = React.useState<IMetadataSnippetRow[]>([]);
  const [taxonomyOptions, setTaxonomyOptions] = React.useState<ITaxonomyFieldOptions | null>(null);
  const [editDialogOpen, setEditDialogOpen] = React.useState<boolean>(false);
  const [editInitialValues, setEditInitialValues] = React.useState<Record<string, any> | null>(null);
  const [editLoading, setEditLoading] = React.useState<boolean>(false);
  const [editSaving, setEditSaving] = React.useState<boolean>(false);
  const [editError, setEditError] = React.useState<string | null>(null);
  const [isUpdateUploadOpen, setIsUpdateUploadOpen] = React.useState<boolean>(false);
  const [isVersionHistoryOpen, setIsVersionHistoryOpen] = React.useState<boolean>(false);
  const [versionHistoryItems, setVersionHistoryItems] = React.useState<IVersionEntry[]>([]);
  const [isVersionHistoryLoading, setIsVersionHistoryLoading] = React.useState<boolean>(false);
  const [versionHistoryError, setVersionHistoryError] = React.useState<string>('');
  const [viewingVersion, setViewingVersion] = React.useState<IVersionDetail | null>(null);
  const [viewingVersionPreview, setViewingVersionPreview] = React.useState<TDetailVersionPreview | null>(null);
  const [isVersionPreviewLoading, setIsVersionPreviewLoading] = React.useState<boolean>(false);
  const [downloadingVersionKeys, setDownloadingVersionKeys] = React.useState<Record<string, boolean>>({});
  const downloadingVersionKeysRef = React.useRef<Set<string>>(new Set());
  const editPrefetchInFlightRef = React.useRef<Promise<void> | null>(null);
  const editPrefetchDocumentIdRef = React.useRef<number | null>(null);
  const editMetadataBodyRef = React.useRef<HTMLDivElement>(null);
  const [pdfPageCount, setPdfPageCount] = React.useState<number>(0);
  const [currentPdfPage, setCurrentPdfPage] = React.useState<number>(1);
  const [pdfZoom, setPdfZoom] = React.useState<number>(1);
  const [pdfRotation, setPdfRotation] = React.useState<number>(0);
  const [pdfScrollMode, setPdfScrollMode] = React.useState<PdfScrollMode>('single');
  const [pdfThumbnailData, setPdfThumbnailData] = React.useState<Uint8Array | null>(null);
  const [pdfThumbnailDocument, setPdfThumbnailDocument] = React.useState<any>(null);
  const [pdfThumbnailError, setPdfThumbnailError] = React.useState<boolean>(false);
  const [thumbnailCandidateIndexes, setThumbnailCandidateIndexes] = React.useState<Record<string, number>>({});
  const [previewCollapsed, setPreviewCollapsed] = React.useState<boolean>(false);
  const [previewViewportWidth, setPreviewViewportWidth] = React.useState<number>(0);
  const resolvedCurrentUserId = React.useMemo(
    () => currentUserId || Number((props.context?.pageContext as any)?.legacyPageContext?.userId || 0),
    [currentUserId, props.context]
  );
  const resolvedCurrentUserEmail = React.useMemo(() => (
    normalizeIdentityValue(
      currentUserEmail ||
      props.userEmail ||
      props.context?.pageContext?.user?.email ||
      props.context?.pageContext?.user?.loginName ||
      ''
    )
  ), [currentUserEmail, props.context, props.userEmail]);
  const socialInteractionInProgressRef = React.useRef<Set<string>>(new Set());
  const [pendingSocialActions, setPendingSocialActions] = React.useState<Record<string, boolean>>({});

  const setSocialInteractionPending = React.useCallback((key: string, isPending: boolean): void => {
    if (isPending) {
      socialInteractionInProgressRef.current.add(key);
    } else {
      socialInteractionInProgressRef.current.delete(key);
    }

    setPendingSocialActions((previous) => {
      if (isPending) {
        return { ...previous, [key]: true };
      }

      const next = { ...previous };
      delete next[key];
      return next;
    });
  }, []);

  React.useEffect(() => {
    const handleBeforeUnload = (event: BeforeUnloadEvent): void => {
      if (socialInteractionInProgressRef.current.size === 0) {
        return;
      }

      event.preventDefault();
      event.returnValue = '';
    };

    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, []);

  React.useEffect(() => {
    return () => {
      if (viewingVersionPreview?.isBlob && viewingVersionPreview.url) {
        URL.revokeObjectURL(viewingVersionPreview.url);
      }
    };
  }, [viewingVersionPreview]);

  const [previewViewportHeight, setPreviewViewportHeight] = React.useState<number>(0);
  const [pdfSidebarPanel, setPdfSidebarPanel] = React.useState<PdfSidebarPanel>(null);
  const [isZoomMenuOpen, setIsZoomMenuOpen] = React.useState<boolean>(false);
  const [zoomInputValue, setZoomInputValue] = React.useState<string>('100');
  const [isEditingZoom, setIsEditingZoom] = React.useState<boolean>(false);
  const [isWindowMode, setIsWindowMode] = React.useState<boolean>(false);
  const [isExcelViewerFullscreen, setIsExcelViewerFullscreen] = React.useState<boolean>(false);
  const [isFitToContainerMode, setIsFitToContainerMode] = React.useState<boolean>(false);
  const [embeddedZoom, setEmbeddedZoom] = React.useState<number>(1);
  const [officeToolbarZoom, setOfficeToolbarZoom] = React.useState<number>(1);
  const [embeddedRotation, setEmbeddedRotation] = React.useState<number>(0);
  const [embeddedPageCount, setEmbeddedPageCount] = React.useState<number>(1);
  const [currentEmbeddedPage, setCurrentEmbeddedPage] = React.useState<number>(1);
  const [viewerNotice, setViewerNotice] = React.useState<string>('');
  const [textContent, setTextContent] = React.useState<string>('');
  const [viewerType, setViewerType] = React.useState<'document' | 'video' | 'audio' | 'text'>('document');
  const [isTextPrintMode, setIsTextPrintMode] = React.useState<boolean>(false);
  const [isPrintInProgress, setIsPrintInProgress] = React.useState<boolean>(false);
  const iframeRef = React.useRef<HTMLIFrameElement>(null);
  const previewContainerRef = React.useRef<HTMLDivElement>(null);
  const previewSectionRef = React.useRef<HTMLDivElement>(null);
  const excelPreviewShellRef = React.useRef<HTMLDivElement>(null);
  const officePageSyncTimeoutsRef = React.useRef<number[]>([]);

  React.useEffect(() => {
    if (!showComments) {
      return;
    }

    const handlePointerDown = (event: MouseEvent): void => {
      const target = event.target as Node;

      if (
        commentsPanelRef.current?.contains(target) ||
        commentButtonRef.current?.contains(target)
      ) {
        return;
      }

      setShowComments(false);
    };

    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        setShowComments(false);
      }
    };

    window.document.addEventListener('mousedown', handlePointerDown);
    window.document.addEventListener('keydown', handleKeyDown);

    return () => {
      window.document.removeEventListener('mousedown', handlePointerDown);
      window.document.removeEventListener('keydown', handleKeyDown);
    };
  }, [showComments]);

  const scrollCommentsPanelIntoView = React.useCallback((): void => {
    const commentsPanel = commentsPanelRef.current;

    if (!commentsPanel) {
      return;
    }

    const prefersReducedMotion = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches;

    commentsPanel.scrollIntoView({
      behavior: prefersReducedMotion ? 'auto' : 'smooth',
      block: 'start'
    });
  }, []);

  React.useEffect(() => {
    if (!showComments || !shouldScrollToCommentsRef.current) {
      return;
    }

    const animationFrameId = window.requestAnimationFrame(() => {
      shouldScrollToCommentsRef.current = false;
      scrollCommentsPanelIntoView();
    });

    return () => {
      window.cancelAnimationFrame(animationFrameId);
    };
  }, [showComments, scrollCommentsPanelIntoView]);

  const handleCommentButtonClick = React.useCallback((): void => {
    shouldScrollToCommentsRef.current = true;

    if (showComments) {
      shouldScrollToCommentsRef.current = false;
      window.requestAnimationFrame(scrollCommentsPanelIntoView);
      return;
    }

    setShowComments(true);
  }, [scrollCommentsPanelIntoView, showComments]);
  const textPrintResetTimeoutRef = React.useRef<number | null>(null);
  const printResetTimeoutRef = React.useRef<number | null>(null);
  const zoomMenuRef = React.useRef<HTMLDivElement>(null);
  const lastViewerWheelActionRef = React.useRef<number>(0);
  const lastTouchpadZoomActionRef = React.useRef<number>(0);
  const gestureStartToolbarZoomRef = React.useRef<number>(1);
  const officeIframeWheelCleanupRef = React.useRef<(() => void) | null>(null);
  const recordedViewKeyRef = React.useRef<string | null>(null);
  const lastMetricSyncKeyRef = React.useRef<string | null>(null);
  const initializedMetricsKeyRef = React.useRef<string | null>(null);
  const pendingViewKeyRef = React.useRef<string | null>(null);
  const officePreviewBeforeEditRef = React.useRef<string | null>(null);
  const officePreviewReleasedForEditRef = React.useRef<boolean>(false);
  const kmDataHubEditFieldMapRef = React.useRef<IKMDataHubEditFieldMap | null>(null);
  const kmsUsersRef = React.useRef<IKmsUser[] | null>(null);
  const metadataSnippetValueMap = React.useMemo<Record<string, string>>(() => (
    metadataSnippetRows.reduce<Record<string, string>>((accumulator, row) => {
      accumulator[row.label] = row.value;
      return accumulator;
    }, {})
  ), [metadataSnippetRows]);
  const documentInfoFields = React.useMemo(() => ([
    { label: 'Document Type', value: metadataSnippetValueMap['Document Type'] || '', supportsMultipleValues: true },
    { label: 'Client', value: metadataSnippetValueMap.Client || '', supportsMultipleValues: true },
    { label: 'Geography', value: metadataSnippetValueMap.Geography || '', supportsMultipleValues: true },
    { label: 'Disease Area', value: metadataSnippetValueMap['Disease Area'] || '', supportsMultipleValues: true },
    { label: 'Therapy Area', value: metadataSnippetValueMap['Therapy Area'] || '', supportsMultipleValues: true },
    { label: 'File Size', value: document?.fileSize || '' },
    { label: 'Uploaded By', value: document?.createdBy || '' }
  ]), [document, metadataSnippetValueMap]);
  const documentStatusValue = document?.status || '';
  const normalizedFileType = document?.fileType.toLowerCase() || '';
  const isCurrentUserCreator = !!document && (
    (
      resolvedCurrentUserId > 0 &&
      document.creatorUserId > 0 &&
      resolvedCurrentUserId === document.creatorUserId
    ) ||
    identityValuesMatch(document.createdByEmail, resolvedCurrentUserEmail)
  );
  const isCurrentUserDocumentAuthor = !!document && (
    (
      resolvedCurrentUserId > 0 &&
      Array.isArray(document.authorUserIds) &&
      document.authorUserIds.indexOf(resolvedCurrentUserId) >= 0
    ) ||
    (document.authorEmails || []).some((authorEmail) => identityValuesMatch(authorEmail, resolvedCurrentUserEmail))
  );
  const canEditOrUpdateDocument = isKmAdmin || isApprover || isCurrentUserCreator || isCurrentUserDocumentAuthor;
  const shouldHideDownloadActions = isLearner;
  const isPdfDocument = normalizedFileType === 'pdf';
  const isVideoDocument = VIDEO_FILE_TYPES.includes(normalizedFileType);
  const isAudioDocument = AUDIO_FILE_TYPES.includes(normalizedFileType);
  const isMediaDocument = isVideoDocument || isAudioDocument;
  const isImageDocument = IMAGE_FILE_TYPES.includes(normalizedFileType);
  const isOfficeDocument = OFFICE_FILE_TYPES.includes(normalizedFileType);
  const isTextDocument = TEXT_FILE_TYPES.includes(normalizedFileType);
  const isPresentationDocument = PRESENTATION_FILE_TYPES.includes(normalizedFileType);
  const isSpreadsheetDocument = SPREADSHEET_FILE_TYPES.includes(normalizedFileType);
  const isExcelLikeDocument =
    isSpreadsheetDocument ||
    CSV_SPREADSHEET_FILE_TYPES.includes(normalizedFileType);
  const isSimpleOfficeStyleToolbarDocument = isExcelLikeDocument || isTextDocument;
  const isExcelStyleToolbarDocument = isSimpleOfficeStyleToolbarDocument;
  const isPdfLikeToolbarDocument = isImageDocument;
  const usesLineIconToolbar = isExcelStyleToolbarDocument || isPdfLikeToolbarDocument;
  const isWordProcessingDocument = WORD_PROCESSING_FILE_TYPES.includes(normalizedFileType);
  const shouldUsePdfPreviewViewer =
    isPdfDocument ||
    ((isWordProcessingDocument || isPresentationDocument) && isConvertedOfficePdfPreview);
  const isPreviewSupportedDocument =
    isPdfDocument ||
    isOfficeDocument ||
    isImageDocument ||
    isTextDocument ||
    isVideoDocument ||
    isAudioDocument;
  const isUnsupportedPreviewDocument = !!document && !isPreviewSupportedDocument;

  React.useEffect(() => {
    setPdfThumbnailDocument(null);
    setPdfThumbnailError(false);
  }, [document?.id, isPdfDocument]);

  const releaseOfficePreviewForEdit = React.useCallback(async (): Promise<boolean> => {
    if (!isOfficeDocument || !previewUrl) {
      return false;
    }

    if (!officePreviewReleasedForEditRef.current) {
      officePreviewBeforeEditRef.current = previewUrl;
      officePreviewReleasedForEditRef.current = true;
      setPreviewUrl('');
      setPreviewError('');
      await new Promise((resolve) => window.setTimeout(resolve, 1200));
    }

    return true;
  }, [isOfficeDocument, previewUrl]);

  const restoreOfficePreviewAfterEdit = React.useCallback((): void => {
    if (!officePreviewReleasedForEditRef.current) {
      return;
    }

    const restoreUrl = officePreviewBeforeEditRef.current;
    officePreviewReleasedForEditRef.current = false;
    officePreviewBeforeEditRef.current = null;

    if (restoreUrl) {
      setPreviewUrl((currentUrl) => currentUrl || restoreUrl);
    }
  }, []);

  const getOfficePageFromUrl = React.useCallback((url: string): number | null => {
    if (!url) {
      return null;
    }

    try {
      const parsedUrl = new URL(url, window.location.origin);

      if (PRESENTATION_FILE_TYPES.includes(normalizedFileType)) {
        const slideIndex = Number(parsedUrl.searchParams.get('wdSlideIndex') || '');
        return Number.isFinite(slideIndex) && slideIndex > 0 ? slideIndex : null;
      }

      if (SPREADSHEET_FILE_TYPES.includes(normalizedFileType)) {
        const sheetIndex = Number(parsedUrl.searchParams.get('wdSheetIndex') || '');
        return Number.isFinite(sheetIndex) && sheetIndex > 0 ? sheetIndex : null;
      }

      if (WORD_PROCESSING_FILE_TYPES.includes(normalizedFileType)) {
        const pageIndex = Number(parsedUrl.searchParams.get('wdPageIndex') || '');
        return Number.isFinite(pageIndex) && pageIndex > 0 ? pageIndex : null;
      }
    } catch (error) {
      console.warn('Unable to parse Office page from preview URL:', error);
    }

    return null;
  }, [normalizedFileType]);

  const getRequestedOfficePageFromUrl = React.useCallback((url: string): number => {
    return getOfficePageFromUrl(url) || 1;
  }, [getOfficePageFromUrl]);

  const extractOfficeViewerPageState = React.useCallback((frame: HTMLIFrameElement): { page: number | null; total: number | null } => {
    let frameText = '';

    try {
      const frameDocument = frame.contentDocument || frame.contentWindow?.document || null;
      frameText =
        frameDocument?.body?.innerText ||
        frameDocument?.documentElement?.innerText ||
        '';
    } catch (error) {
      console.warn('Unable to read Office iframe text content:', error);
    }

    const normalizedText = frameText.replace(/\s+/g, ' ').trim();

    if (WORD_PROCESSING_FILE_TYPES.includes(normalizedFileType)) {
      const docMatch = normalizedText.match(/PAGE\s+(\d+)\s+OF\s+(\d+)/i);
      if (docMatch) {
        return {
          page: Math.max(Number(docMatch[1] || '1'), 1),
          total: Math.max(Number(docMatch[2] || '1'), 1)
        };
      }
    }

    if (PRESENTATION_FILE_TYPES.includes(normalizedFileType)) {
      const slideMatch =
        normalizedText.match(/SLIDE\s+(\d+)\s+OF\s+(\d+)/i) ||
        normalizedText.match(/\b(\d+)\s+OF\s+(\d+)\b/i) ||
        normalizedText.match(/\b(\d+)\s*\/\s*(\d+)\b/i);

      if (slideMatch) {
        return {
          page: Math.max(Number(slideMatch[1] || '1'), 1),
          total: Math.max(Number(slideMatch[2] || '1'), 1)
        };
      }
    }

    if (SPREADSHEET_FILE_TYPES.includes(normalizedFileType)) {
      const sheetMatch =
        normalizedText.match(/SHEET\s+(\d+)\s+OF\s+(\d+)/i) ||
        normalizedText.match(/\b(\d+)\s+OF\s+(\d+)\b/i);

      if (sheetMatch) {
        return {
          page: Math.max(Number(sheetMatch[1] || '1'), 1),
          total: Math.max(Number(sheetMatch[2] || '1'), 1)
        };
      }
    }

    return { page: null, total: null };
  }, [normalizedFileType]);

  const syncOfficeViewerPageState = React.useCallback((frame: HTMLIFrameElement | null) => {
    officePageSyncTimeoutsRef.current.forEach((timeoutId) => window.clearTimeout(timeoutId));
    officePageSyncTimeoutsRef.current = [];

    if (!frame || !isOfficeDocument) {
      return;
    }

    const requestedPage = getRequestedOfficePageFromUrl(previewUrl);

    const syncFromFrame = () => {
      let detectedPage: number | null = null;

      try {
        detectedPage = getOfficePageFromUrl(frame.contentWindow?.location?.href || '');
      } catch (error) {
        console.warn('Unable to read Office iframe location:', error);
      }

      const viewerState = extractOfficeViewerPageState(frame);

      if (viewerState.page !== null) {
        detectedPage = viewerState.page;
      }

      if (detectedPage !== null) {
        setCurrentEmbeddedPage(detectedPage);
      } else {
        setCurrentEmbeddedPage(requestedPage);
      }

      if (viewerState.total !== null && viewerState.total !== embeddedPageCount) {
        setEmbeddedPageCount(viewerState.total);
      }

      if (
        detectedPage === null &&
        viewerState.total === null &&
        (isPresentationDocument || isSpreadsheetDocument)
      ) {
        const frameSrcPage = getOfficePageFromUrl(frame.getAttribute('src') || '');
        if (frameSrcPage !== null) {
          setCurrentEmbeddedPage(frameSrcPage);
        }
      }
    };

    [0, 300, 900, 1800].forEach((delay) => {
      const timeoutId = window.setTimeout(syncFromFrame, delay);
      officePageSyncTimeoutsRef.current.push(timeoutId);
    });
  }, [embeddedPageCount, extractOfficeViewerPageState, getOfficePageFromUrl, getRequestedOfficePageFromUrl, isOfficeDocument, normalizedFileType, previewUrl]);

  const resetSpreadsheetViewerPosition = React.useCallback((frame: HTMLIFrameElement | null): void => {
    if (!frame || !isSpreadsheetDocument) {
      return;
    }

    const resetFrameScroll = (): void => {
      try {
        const frameWindow = frame.contentWindow;
        const frameDocument = frame.contentDocument || frameWindow?.document;

        frameWindow?.scrollTo(0, 0);
        if (!frameDocument) {
          return;
        }

        frameDocument.documentElement.scrollLeft = 0;
        frameDocument.documentElement.scrollTop = 0;
        frameDocument.body.scrollLeft = 0;
        frameDocument.body.scrollTop = 0;

        Array.from(frameDocument.querySelectorAll<HTMLElement>('*')).forEach((element) => {
          if (element.scrollHeight > element.clientHeight || element.scrollWidth > element.clientWidth) {
            element.scrollLeft = 0;
            element.scrollTop = 0;
          }
        });
      } catch (error) {
        // Office iframes can briefly block access while bootstrapping; delayed retries handle the accessible case.
      }
    };

    [0, 250, 800, 1600, 2600].forEach((delay) => {
      window.setTimeout(resetFrameScroll, delay);
    });
  }, [isSpreadsheetDocument]);

  const resetPdfViewerState = React.useCallback(() => {
    setPdfPageCount(0);
    setCurrentPdfPage(1);
    setPdfZoom(1);
    setPdfRotation(0);
    setPdfScrollMode('single');
    setPdfThumbnailData(null);
    setPdfThumbnailDocument(null);
    setPdfThumbnailError(false);
    setPdfSidebarPanel(null);
    setIsZoomMenuOpen(false);
    setIsWindowMode(false);
    setIsFitToContainerMode(false);
    setEmbeddedZoom(1);
    setOfficeToolbarZoom(1);
    setEmbeddedRotation(0);
    setEmbeddedPageCount(1);
    setCurrentEmbeddedPage(1);
    setViewerNotice('');
    setPreviewCollapsed(false);
    setPreviewError('');
    setPreviewImageUrls([]);
    setPreviewUrl('');
    setIsPreviewLoading(false);
    setIsConvertedOfficePdfPreview(false);
    setTextContent('');
  }, []);

  // Cleanup blob URLs on unmount
  React.useEffect(() => {
    return () => {
      officePageSyncTimeoutsRef.current.forEach((timeoutId) => window.clearTimeout(timeoutId));
      officePageSyncTimeoutsRef.current = [];
      if (textPrintResetTimeoutRef.current) {
        window.clearTimeout(textPrintResetTimeoutRef.current);
        textPrintResetTimeoutRef.current = null;
      }
      if (printResetTimeoutRef.current) {
        window.clearTimeout(printResetTimeoutRef.current);
        printResetTimeoutRef.current = null;
      }
      if (previewUrl && previewUrl.startsWith('blob:')) {
        window.URL.revokeObjectURL(previewUrl);
      }
    };
  }, [previewUrl]);

  React.useEffect(() => {
    resetPdfViewerState();
    recordedViewKeyRef.current = null;
    lastMetricSyncKeyRef.current = null;
    initializedMetricsKeyRef.current = null;
    pendingViewKeyRef.current = null;
  }, [props.documentId, resetPdfViewerState]);

  React.useEffect(() => {
    if (pdfPageCount > 0 && currentPdfPage > pdfPageCount) {
      setCurrentPdfPage(pdfPageCount);
    }
  }, [currentPdfPage, pdfPageCount]);

  React.useEffect(() => {
    if (embeddedPageCount > 0 && currentEmbeddedPage > embeddedPageCount) {
      setCurrentEmbeddedPage(embeddedPageCount);
    }
  }, [currentEmbeddedPage, embeddedPageCount]);

  React.useEffect(() => {
    if (!viewerNotice) {
      return;
    }

    const timeoutId = window.setTimeout(() => setViewerNotice(''), 3200);
    return () => window.clearTimeout(timeoutId);
  }, [viewerNotice]);

  React.useEffect(() => {
    const handlePointerDown = (event: MouseEvent) => {
      const target = event.target as Node | null;
      if (zoomMenuRef.current && !zoomMenuRef.current.contains(target)) {
        setIsZoomMenuOpen(false);
      }
    };

    window.document.addEventListener('mousedown', handlePointerDown);
    return () => window.document.removeEventListener('mousedown', handlePointerDown);
  }, []);

  React.useEffect(() => {
    const container = previewContainerRef.current;
    if (!container) {
      return;
    }

    const updateDimensions = () => {
      setPreviewViewportWidth(container.clientWidth);
      setPreviewViewportHeight(container.clientHeight);
    };

    updateDimensions();

    if (typeof ResizeObserver !== 'undefined') {
      const observer = new ResizeObserver(() => updateDimensions());
      observer.observe(container);

      return () => observer.disconnect();
    }

    window.addEventListener('resize', updateDimensions);
    return () => window.removeEventListener('resize', updateDimensions);
  }, [previewCollapsed, document?.id]);

  React.useEffect(() => {
    if (!previewContainerRef.current) {
      return;
    }

    const timeoutId = window.setTimeout(() => {
      const container = previewContainerRef.current;
      if (!container) {
        return;
      }

      container.scrollLeft = 0;
      container.scrollTop = 0;
    }, 0);

    return () => window.clearTimeout(timeoutId);
  }, [embeddedRotation, embeddedZoom, previewViewportWidth, previewViewportHeight, previewUrl]);

  React.useEffect(() => {
    const handleFullscreenChange = (): void => {
      setIsExcelViewerFullscreen(
        window.document.fullscreenElement === excelPreviewShellRef.current ||
        window.document.fullscreenElement === previewSectionRef.current
      );
    };

    window.document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => window.document.removeEventListener('fullscreenchange', handleFullscreenChange);
  }, []);

  React.useEffect(() => {
    setLoading(true);
    setDocument(null);
    setFetchError('');
    sessionStorage.removeItem(`documentDisableSocial_${props.documentId}`);
    fetchDocumentDetails();
    fetchCurrentUser();
  }, [props.documentId]);

  React.useEffect(() => {
    setIsKmAdmin(true);
    setIsApprover(true);
    setIsContributor(true);
    setIsLearner(false);
  }, [props.context]);

  React.useEffect(() => {
    return subscribeToDocumentDataChanged((detail) => {
      const changedIds = detail.documentIds || [];
      if (changedIds.length > 0 && changedIds.indexOf(props.documentId) === -1) {
        return;
      }

      fetchDocumentDetails();
    });
  }, [props.documentId]);

  React.useEffect(() => {
    if (!actionMessage) {
      return;
    }

    const timeoutId = window.setTimeout(() => setActionMessage(''), 2800);
    return () => window.clearTimeout(timeoutId);
  }, [actionMessage]);

  React.useEffect(() => {
    if (!isTextPrintMode) {
      return;
    }

    const clearTextPrintMode = () => {
      setIsTextPrintMode(false);
      if (textPrintResetTimeoutRef.current) {
        window.clearTimeout(textPrintResetTimeoutRef.current);
        textPrintResetTimeoutRef.current = null;
      }
    };

    window.addEventListener('afterprint', clearTextPrintMode, { once: true });
    textPrintResetTimeoutRef.current = window.setTimeout(clearTextPrintMode, 30000);

    return () => {
      window.removeEventListener('afterprint', clearTextPrintMode);
      if (textPrintResetTimeoutRef.current) {
        window.clearTimeout(textPrintResetTimeoutRef.current);
        textPrintResetTimeoutRef.current = null;
      }
    };
  }, [isTextPrintMode]);

  React.useEffect(() => {
    if (!isPrintInProgress) {
      return;
    }

    const clearPrintingState = () => setIsPrintInProgress(false);
    window.addEventListener('afterprint', clearPrintingState, { once: true });

    return () => {
      window.removeEventListener('afterprint', clearPrintingState);
    };
  }, [isPrintInProgress]);

  React.useEffect(() => {
    if (document && resolvedCurrentUserId && props.context) {
      const metricsKey = `${document.id}-${resolvedCurrentUserId}`;
      if (initializedMetricsKeyRef.current === metricsKey) {
        return;
      }

      const initializeMetrics = async () => {
        const [metrics, interactionState] = await Promise.all([
          getDocumentMetrics(props.context, document.id),
          getUserInteractionState(props.context, document.id, resolvedCurrentUserId)
        ]);

        if (initializedMetricsKeyRef.current === metricsKey) {
          return;
        }

        setViewCount(metrics.views);
        setLikeCount(metrics.likes);
        setCommentCount(metrics.comments);
        setDownloadCount(metrics.downloads);
        setFollowCount(metrics.follow);
        setShareCount(metrics.share);
        setBookmarkCount(metrics.bookmark);
        setIsLiked(interactionState.isLiked);
        setIsBookmarked(interactionState.isBookmarked);
        setIsFollowing(interactionState.isFollowed);
        initializedMetricsKeyRef.current = metricsKey;

        fetchComments();
        fetchDocumentActionMetrics('flag');

      };
      initializeMetrics().catch((error) => {
        console.error('Error initializing document metrics:', error);
      });
    }
  }, [document?.id, resolvedCurrentUserId, props.context]);

  React.useEffect(() => {
    if (document) {
      if (isVideoDocument) {
        setViewerType('video');
      } else if (isAudioDocument) {
        setViewerType('audio');
      } else if (isTextDocument) {
        setViewerType('text');
      } else {
        setViewerType('document');
      }
    }
  }, [document, isVideoDocument, isAudioDocument, isTextDocument]);

  const fetchDocumentDetails = async () => {
    if (!props.context) {
      setLoading(false);
      return;
    }

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const libraryName = props.listTitle || DEFAULT_LIBRARY_NAME;
      const fieldMap = await fetchKMDataHubReadFieldMap(props.context.spHttpClient, webUrl, libraryName);

      console.log('DocumentDetailPage: Fetching document with ID:', props.documentId);
      console.log('DocumentDetailPage: Web URL:', webUrl);

      const baseSelect = [
        'ID',
        fieldMap.title || COLUMN_NAMES.title,
        fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef,
        fieldMap.fileRef || COLUMN_NAMES.fileRef,
        fieldMap.description || COLUMN_NAMES.description,
        fieldMap.edited || COLUMN_NAMES.edited,
        COLUMN_NAMES.modified,
        COLUMN_NAMES.created,
        fieldMap.status || COLUMN_NAMES.status,
        `${fieldMap.editedBy || COLUMN_NAMES.editedBy}/Title`,
        `${fieldMap.editedBy || COLUMN_NAMES.editedBy}/EMail`,
        `${fieldMap.editedBy || COLUMN_NAMES.editedBy}/Id`,
        'Author/Title',
        'Author/Name',
        'Author/EMail',
        'Author/Id',
        fieldMap.contentRefreshDate || COLUMN_NAMES.contentRefreshDate,
        fieldMap.views || COLUMN_NAMES.views,
        fieldMap.likes || COLUMN_NAMES.likes,
        fieldMap.comments || COLUMN_NAMES.comments,
        fieldMap.downloads || COLUMN_NAMES.downloads,
        fieldMap.follow || COLUMN_NAMES.follow,
        fieldMap.share || COLUMN_NAMES.share,
        fieldMap.bookmark || COLUMN_NAMES.bookmark,
        fieldMap.sensitiveTerms || COLUMN_NAMES.sensitiveTerms,
        fieldMap.reviewerComments || COLUMN_NAMES.reviewerComments,
        fieldMap.projectId || COLUMN_NAMES.projectId,
        fieldMap.versionFileName || COLUMN_NAMES.versionFileName,
        fieldMap.versionFileType || COLUMN_NAMES.versionFileType,
        fieldMap.docIcon || COLUMN_NAMES.docIcon
      ];
      const queryParts = buildKMDataHubItemQuery(fieldMap, baseSelect, [fieldMap.editedBy || COLUMN_NAMES.editedBy, 'Author']);

      // Ensure File/UniqueId is expanded separately and avoid 'Abstract' in select
      const finalSelect = `${queryParts.select},File/UniqueId,File/Name,File/ServerRelativeUrl`;
      const finalExpand = `${queryParts.expand},File`;

      // Use filter instead of direct ID access for better stability
      const listEndpoint = props.listId
        ? `lists(guid'${props.listId}')`
        : `lists/getbytitle('${libraryName}')`;
      const apiUrl = `${webUrl}/_api/web/${listEndpoint}/items?$select=${finalSelect}&$expand=${finalExpand}&$filter=Id eq ${props.documentId}`;

      console.log('DocumentDetailPage: Fetching document with ID:', props.documentId);
      console.log('DocumentDetailPage: API URL:', apiUrl);

      const response: SPHttpClientResponse = await props.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const errorText = await response.text();
        console.error('DocumentDetailPage: API Error Response Status:', response.status);
        console.error('DocumentDetailPage: API Error Response Text:', errorText);

        setFetchError(`Failed to fetch document: ${response.status} - ${errorText}`);
        setLoading(false);
        return;
      }

      const data: any = await response.json();
      const item = data.value && data.value.length > 0 ? data.value[0] : null;

      if (!item) {
        setFetchError(`No document found with ID ${props.documentId} in '${libraryName}'.`);
        setLoading(false);
        return;
      }

      console.log('DocumentDetailPage: Document retrieved successfully');

      const modifiedByField = fieldMap.editedBy || COLUMN_NAMES.editedBy;
      const fileLeafRefField = fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef;
      const fileRefField = fieldMap.fileRef || COLUMN_NAMES.fileRef;
      const serverRelativeUrl = item.File?.ServerRelativeUrl || item[fileRefField] || item.FileRef || '';
      const absoluteUrlPath = normalizeSharePointUrl(webUrl, serverRelativeUrl);
      const fileName = item.File?.Name || item[fileLeafRefField] || item[COLUMN_NAMES.fileLeafRef] || item[fieldMap.title || COLUMN_NAMES.title] || '';
      const fileExtension = fileName.split('.').pop()?.toUpperCase() || '';
      const displayName = item[fieldMap.title || COLUMN_NAMES.title] || item[COLUMN_NAMES.title] || fileName || `Document ${item.Id}`;
      const abstract = normalizeDescriptionText(
        item[fieldMap.description || COLUMN_NAMES.description] || item[COLUMN_NAMES.description] || item.Abstract || ''
      );

      console.log('DocumentDetailPage: File details:', {
        fileName,
        fileExtension,
        displayName,
        serverRelativeUrl
      });

      // Prefer KM Data Hub's Author0 field, then fall back to legacy fields
      const authorId = Number(item.AuthorId || item.Author?.Id || item.Author?.ID || 0);
      const shouldResolveAuthorById =
        typeof authorId === 'number' &&
        authorId > 0 &&
        !item.Author0 &&
        !item.Author?.Title &&
        !item.Author?.Name;

      // AUTHOR RESOLUTION: Use the schema-aware resolver
      const author = resolveDocumentAuthor(item, 'Internal', fieldMap.author);
      let createdBy = item.Author?.Title || item.Author?.Name || '';
      const createdByIdentityCandidates = uniqueIdentityCandidates(extractIdentityCandidates(item.Author));
      const createdByEmail = createdByIdentityCandidates[0] || item.Author?.EMail || item.Author?.Email || item.Author?.Name || '';
      const creatorUserId = authorId;
      const authorUserIds = Array.from(
        new Set<number>([
          ...extractNumericIds(item?.[`${fieldMap.author}Id`]),
          ...extractNumericIds(item?.[fieldMap.author]),
          ...extractNumericIds(item?.Author0),
          ...extractNumericIds(item?.Author)
        ])
      );
      const authorEmails = uniqueIdentityCandidates([
        ...extractIdentityCandidates(item?.[fieldMap.author]),
        ...extractIdentityCandidates(item?.Author0)
      ]);

      if (shouldResolveAuthorById) {
        if (AUTHOR_CACHE[authorId]) {
          createdBy = createdBy || AUTHOR_CACHE[authorId];
        } else {
          // If we only have AuthorId, fetch the user details
          console.log('Have AuthorId:', authorId, '- will fetch user info');
          try {
            const userResp = await props.context.spHttpClient.get(
              `${webUrl}/_api/web/getuserbyid(${authorId})`,
              SPHttpClient.configurations.v1
            );
            if (userResp.ok) {
              const user = await userResp.json();
              AUTHOR_CACHE[authorId] = user.Title || user.LoginName || 'Unknown';
              createdBy = createdBy || AUTHOR_CACHE[authorId];
              console.log('Fetched author from AuthorId:', AUTHOR_CACHE[authorId]);
            }
          } catch (userError) {
            console.warn('Could not fetch author user info:', userError);
          }
        }
      }

      const formattedAuthor = formatAuthorDisplayValue(author);
      const formattedCreatedBy = createdBy ? formatAuthorDisplayValue(createdBy) : '';

      console.log('DocumentDetailPage: Author resolved to:', formattedAuthor);

      // Prefer KM Data Hub's Published field for published date
      const formattedDate = formatDocumentPublishedDate(resolveDocumentPublishedValue(item, fieldMap.published), 'en-GB', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });

      // Get file size
      const fileSize = '';
      // File size will be fetched separately if needed

      console.log('Document retrieval - FileRef/serverRelativeUrl:', {
        FileRef: item.FileRef,
        serverRelativeUrl: serverRelativeUrl,
        fileExtension: fileExtension
      });

      // Fetch file size asynchronously
      if (serverRelativeUrl && props.context) {
        fetchFileSizeAsync(serverRelativeUrl).catch(err => {
          console.error('Error fetching file size:', err);
        });
      }
      const contentRefreshValue = item[fieldMap.contentRefreshDate || COLUMN_NAMES.contentRefreshDate];
      const modifiedByValue = item[modifiedByField] || item[COLUMN_NAMES.editedBy] || item[COLUMN_NAMES.modifiedBy];
      const documentDetail: DocumentDetail = {
        id: item.Id,
        name: displayName,
        fileName: fileName,
        abstract: abstract,
        fileType: fileExtension,
        author: formattedAuthor,
        createdBy: formattedCreatedBy,
        createdByEmail,
        authorUserIds,
        authorEmails,
        creatorUserId,
        date: formattedDate,
        createdDate: String(item.Created || ''),
        fileSize: fileSize,
        serverRelativeUrl: serverRelativeUrl,
        fileRef: item[fileRefField] || item.FileRef || serverRelativeUrl,
        fileUniqueId: item.File?.UniqueId || '',
        status: item[fieldMap.status || COLUMN_NAMES.status] || item[COLUMN_NAMES.status] || '',
        contentRefreshDate: formatDocumentPublishedDate(String(contentRefreshValue || ''), 'en-GB', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        }) || formatTaxonomyDisplayValue(contentRefreshValue),
        views: normalizeDocumentNumberField(item[fieldMap.views || COLUMN_NAMES.views]),
        likes: normalizeDocumentNumberField(item[fieldMap.likes || COLUMN_NAMES.likes]),
        comments: normalizeDocumentNumberField(item[fieldMap.comments || COLUMN_NAMES.comments]),
        downloads: normalizeDocumentNumberField(item[fieldMap.downloads || COLUMN_NAMES.downloads]),
        follow: normalizeDocumentBooleanField(item[fieldMap.follow || COLUMN_NAMES.follow]),
        share: normalizeDocumentBooleanField(item[fieldMap.share || COLUMN_NAMES.share]),
        bookmark: normalizeDocumentBooleanField(item[fieldMap.bookmark || COLUMN_NAMES.bookmark]),
        sensitiveTerms: normalizeDocumentStringArrayField(item[fieldMap.sensitiveTerms || COLUMN_NAMES.sensitiveTerms]),
        reviewerComments: formatTaxonomyDisplayValue(item[fieldMap.reviewerComments || COLUMN_NAMES.reviewerComments]),
        projectId: formatTaxonomyDisplayValue(item[fieldMap.projectId || COLUMN_NAMES.projectId]),
        versionFileName: formatTaxonomyDisplayValue(item[fieldMap.versionFileName || COLUMN_NAMES.versionFileName]),
        versionFileType: formatTaxonomyDisplayValue(item[fieldMap.versionFileType || COLUMN_NAMES.versionFileType]),
        modifiedBy: modifiedByValue
          ? {
            title: formatTaxonomyDisplayValue(modifiedByValue.Title),
            email: formatTaxonomyDisplayValue(modifiedByValue.EMail),
            id: Number(modifiedByValue.Id || 0)
          }
          : undefined,
        docIcon: formatTaxonomyDisplayValue(item[fieldMap.docIcon || COLUMN_NAMES.docIcon])
      };

      setDocument(documentDetail);
      void fetchMetadataSnippetRows(item.Id);

      // Show page immediately - don't wait for tags or preview
      setLoading(false);

      // Start the preview before secondary metadata work so the document frame gets a URL immediately.
      if (serverRelativeUrl) {
        setIsPreviewLoading(true);
        const previewCacheBustToken = String(new Date(item[fieldMap.edited || COLUMN_NAMES.edited] || item.Modified || Date.now()).getTime());
        generatePreviewUrlAsync(serverRelativeUrl, fileExtension, previewCacheBustToken, documentDetail.fileUniqueId)
          .catch((previewGenerationError) => {
            console.error('Error generating preview URL:', previewGenerationError);
            setPreviewError('Failed to load document preview');
            setPreviewUrl('');
            setIsPreviewLoading(false);
          });
      } else {
        setIsPreviewLoading(false);
      }

    } catch (error) {
      console.error('Error fetching document details:', error);
      setLoading(false);
    }
  };

  const fetchFileSizeAsync = async (serverRelativeUrl: string) => {
    if (!props.context) return;

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const normalizedServerRelativeUrl = normalizeServerRelativeUrl(serverRelativeUrl);

      // Get file properties including size
      const response = await props.context.spHttpClient.get(
        buildFileApiUrl(webUrl, normalizedServerRelativeUrl, '?$select=Length'),
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const fileProps = await response.json();
        const bytes = fileProps.Length;

        if (bytes) {
          let formattedSize = '';
          if (bytes < 1024) {
            formattedSize = `${bytes} B`;
          } else if (bytes < 1024 * 1024) {
            formattedSize = `${(bytes / 1024).toFixed(1)} KB`;
          } else {
            formattedSize = `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
          }

          // Update document with file size
          setDocument(prevDoc => {
            if (prevDoc) {
              return { ...prevDoc, fileSize: formattedSize };
            }
            return prevDoc;
          });
        }
      }
    } catch (error) {
      console.error('Error fetching file size:', error);
      // Continue without file size - it's not critical
    }
  };

  const generatePreviewUrlAsync = async (
    serverRelativeUrl: string,
    fileExtension: string,
    cacheBustToken?: string,
    fileUniqueId?: string
  ) => {
    if (!props.context) {
      console.warn('No context available for preview generation');
      setIsPreviewLoading(false);
      return;
    }

    console.log('generatePreviewUrlAsync called with:', {
      serverRelativeUrl,
      fileExtension
    });

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const fileUrl = normalizeServerRelativeUrl(serverRelativeUrl);
    const fullFileUrl = buildAbsoluteFileUrl(webUrl, fileUrl);

    console.log('Preview URLs:', {
      fileUrl,
      fullFileUrl,
      fileExtension: fileExtension.toLowerCase()
    });
    setPreviewError('');
    setTextContent('');
    setIsConvertedOfficePdfPreview(false);

    // For Office documents, keep the preview inside the SharePoint WOPI iframe.
    if (OFFICE_FILE_TYPES.indexOf(fileExtension.toLowerCase()) !== -1) {
      setPreviewImageUrls([]);
      const normalizedExtension = fileExtension.toLowerCase();
      setCurrentEmbeddedPage(1);
      setEmbeddedPageCount(0);

      if (
        WORD_PROCESSING_FILE_TYPES.includes(normalizedExtension) ||
        PRESENTATION_FILE_TYPES.includes(normalizedExtension)
      ) {
        const wopiFallbackUrl = appendCacheBustParameter(
          buildOfficePreviewUrl(webUrl, fileUrl, normalizedExtension, 1),
          cacheBustToken
        );

        try {
          if (!props.context.msGraphClientFactory) {
            throw new Error('Microsoft Graph client is not available.');
          }

          setPreviewUrl('');
          const graphClient = await props.context.msGraphClientFactory.getClient('3');
          const sharingToken = buildGraphSharingToken(fullFileUrl);
          const convertedPdfBlob = await graphClient
            .api(`/shares/${sharingToken}/driveItem/content?format=pdf`)
            .version('v1.0')
            .responseType(ResponseType.BLOB)
            .get();

          if (!(convertedPdfBlob instanceof Blob) || convertedPdfBlob.size === 0) {
            throw new Error('Microsoft Graph returned an empty Office PDF preview.');
          }

          const convertedPdfUrl = window.URL.createObjectURL(
            convertedPdfBlob.type === 'application/pdf'
              ? convertedPdfBlob
              : new Blob([convertedPdfBlob], { type: 'application/pdf' })
          );

          setPdfPageCount(1);
          setCurrentPdfPage(1);
          setIsConvertedOfficePdfPreview(true);
          setPreviewImageUrls([]);
          setPreviewUrl(convertedPdfUrl);
          setEmbeddedPageCount(1);
          setIsPreviewLoading(false);
          return;
        } catch (officePdfPreviewError) {
          console.warn('Unable to convert Office document to PDF preview. Falling back to Office preview:', officePdfPreviewError);
          setIsConvertedOfficePdfPreview(false);
          setPreviewUrl(wopiFallbackUrl);
          setIsPreviewLoading(false);
        }
      }

      if (PRESENTATION_FILE_TYPES.includes(normalizedExtension)) {
        const siteId = String(props.context.pageContext.site.id || '');
        const webId = String(props.context.pageContext.web.id || '');
        const wopiUrl = buildOfficePreviewUrl(webUrl, fileUrl, normalizedExtension, 1);
        const slidePreviewCandidates = buildSlideImagePreviewCandidates(
          webUrl,
          fileUrl,
          1,
          fileUniqueId || document?.fileUniqueId,
          siteId,
          webId
        );
        console.log('Setting presentation image preview URLs:', slidePreviewCandidates);
        setPreviewImageUrls(slidePreviewCandidates);
        setPreviewUrl(appendCacheBustParameter(wopiUrl, cacheBustToken));
        setIsPreviewLoading(false);
      } else {
        const wopiUrl = buildOfficePreviewUrl(webUrl, fileUrl, normalizedExtension, 1);
        console.log('Setting Office iframe preview URL:', wopiUrl);
        setPreviewUrl(appendCacheBustParameter(wopiUrl, cacheBustToken));
        setIsPreviewLoading(false);
      }

      await new Promise<void>((resolve) => {
        window.requestAnimationFrame(() => window.setTimeout(resolve, 0));
      });

      try {
        const downloadUrl = buildFileApiUrl(webUrl, fileUrl);
        const fileResponse = await props.context.spHttpClient.get(
          downloadUrl,
          SPHttpClient.configurations.v1
        );

        let fileBuffer: ArrayBuffer | null = null;

        if (fileResponse.ok) {
          fileBuffer = await fileResponse.arrayBuffer();
        } else if (props.context.msGraphClientFactory) {
          const graphClient = await props.context.msGraphClientFactory.getClient('3');
          const sharingToken = buildGraphSharingToken(fullFileUrl);
          const fileBlob = await graphClient
            .api(`/shares/${sharingToken}/driveItem/content`)
            .version('v1.0')
            .responseType(ResponseType.BLOB)
            .get();

          if (fileBlob instanceof Blob && fileBlob.size > 0) {
            fileBuffer = await fileBlob.arrayBuffer();
          }
        }

        if (!fileBuffer) {
          return;
        }

        if (SPREADSHEET_FILE_TYPES.includes(normalizedExtension)) {
          try {
            const workbookXml = await (await import('jszip')).loadAsync(fileBuffer).then((zip: any) => zip.file('xl/workbook.xml')?.async('string'));
            const sheetMatches = workbookXml?.match(/<sheet\b/gi);
            setEmbeddedPageCount(Math.max(sheetMatches?.length || 1, 1));
          } catch (spreadsheetPreviewError) {
            console.warn('Unable to derive spreadsheet page count:', spreadsheetPreviewError);
          }

          return;
        }

        const JSZipModule = await import('jszip');
        const JSZip = (JSZipModule as any).default || JSZipModule;
        const zip = await JSZip.loadAsync(fileBuffer);

        if (PRESENTATION_FILE_TYPES.includes(normalizedExtension)) {
          setEmbeddedPageCount(await getPowerPointSlideCountFromZip(fileBuffer));
        } else if (WORD_PROCESSING_FILE_TYPES.includes(normalizedExtension)) {
          const appXml = await zip.file('docProps/app.xml')?.async('string');
          const pagesMatch = appXml?.match(/<Pages>(\d+)<\/Pages>/i);
          setEmbeddedPageCount(Math.max(Number(pagesMatch?.[1] || 1), 1));
        } else if (SPREADSHEET_FILE_TYPES.includes(normalizedExtension)) {
          const workbookXml = await zip.file('xl/workbook.xml')?.async('string');
          const sheetMatches = workbookXml?.match(/<sheet\b/gi);
          setEmbeddedPageCount(Math.max(sheetMatches?.length || 1, 1));
        }
      } catch (officePreviewError) {
        console.warn('Unable to derive embedded page count:', officePreviewError);
        if (PRESENTATION_FILE_TYPES.includes(normalizedExtension) && props.context.msGraphClientFactory) {
          try {
            const graphClient = await props.context.msGraphClientFactory.getClient('3');
            const sharingToken = buildGraphSharingToken(fullFileUrl);
            const fileBlob = await graphClient
              .api(`/shares/${sharingToken}/driveItem/content`)
              .version('v1.0')
              .responseType(ResponseType.BLOB)
              .get();

            if (fileBlob instanceof Blob && fileBlob.size > 0) {
              setEmbeddedPageCount(await getPowerPointSlideCountFromZip(await fileBlob.arrayBuffer()));
            }
          } catch (presentationCountFallbackError) {
            console.warn('Unable to derive PowerPoint slide count from Graph fallback:', presentationCountFallbackError);
          }
        }
      }
    } else if (fileExtension.toLowerCase() === 'pdf') {
      setPdfPageCount(1);
      setCurrentPdfPage(1);
      setPreviewImageUrls([]);
      setPdfThumbnailData(null);
      setPdfThumbnailDocument(null);
      setPdfThumbnailError(false);
      // Revert to direct file URL - WOPI does not support PDF.
      setPreviewUrl(appendCacheBustParameter(fullFileUrl, cacheBustToken));
      setIsPreviewLoading(false);
    } else if (IMAGE_FILE_TYPES.includes(fileExtension.toLowerCase())) {
      setPreviewImageUrls([]);
      setPreviewUrl(appendCacheBustParameter(fullFileUrl, cacheBustToken));
      setIsPreviewLoading(false);
    } else if (TEXT_FILE_TYPES.includes(fileExtension.toLowerCase())) {
      setPreviewImageUrls([]);
      try {
        const downloadUrl = buildFileApiUrl(webUrl, fileUrl);
        const fileResponse = await props.context.spHttpClient.get(
          downloadUrl,
          SPHttpClient.configurations.v1
        );
        if (fileResponse.ok) {
          const text = await fileResponse.text();
          setTextContent(text);
          setPreviewUrl('loaded');
          setIsPreviewLoading(false);
        } else {
          setPreviewUrl('loaded');
          setPreviewError('Failed to load text content');
          setIsPreviewLoading(false);
        }
      } catch (error) {
        console.error('Error fetching text content:', error);
        setPreviewUrl('loaded');
        setPreviewError('Failed to load text content');
        setIsPreviewLoading(false);
      }
    } else if (VIDEO_FILE_TYPES.includes(fileExtension.toLowerCase())) {
      setPreviewImageUrls([]);
      setPreviewUrl(appendCacheBustParameter(fullFileUrl, cacheBustToken));
      setIsPreviewLoading(false);
    } else if (AUDIO_FILE_TYPES.includes(fileExtension.toLowerCase())) {
      setPreviewImageUrls([]);
      setPreviewUrl(appendCacheBustParameter(fullFileUrl, cacheBustToken));
      setIsPreviewLoading(false);
    } else {
      setPreviewImageUrls([]);
      setPreviewUrl('');
      setPreviewError('');
      setIsPreviewLoading(false);
    }
  };

  const getFileTypeIcon = (fileType: string): string => {
    const type = fileType.toLowerCase();
    if (type === 'pdf') {
      return '📄';
    } else if (type === 'pptx' || type === 'ppt') {
      return '📊';
    } else if (type === 'docx' || type === 'doc') {
      return '📝';
    } else if (type === 'xlsx' || type === 'xls') {
      return '📈';
    }
    return '📎';
  };

  const loadTaxonomyOptions = React.useCallback(async (): Promise<ITaxonomyFieldOptions | null> => {
    if (!props.context) {
      return null;
    }

    if (taxonomyOptions) {
      return taxonomyOptions;
    }

    try {
      const loadedOptions = await fetchAllTaxonomyOptions(props.context);
      setTaxonomyOptions(loadedOptions);
      return loadedOptions;
    } catch (error) {
      console.error('Failed to load taxonomy options:', error);
      return null;
    }
  }, [props.context, taxonomyOptions]);

  const loadKmsUsers = React.useCallback(async (): Promise<IKmsUser[]> => {
    if (!props.context) {
      return [];
    }

    if (kmsUsersRef.current) {
      return kmsUsersRef.current;
    }

    try {
      const users = await getKmsUsers(props.context);
      kmsUsersRef.current = users;
      return users;
    } catch (error) {
      console.error('Failed to load KMS users:', error);
      return [];
    }
  }, [props.context]);

  const getKMDataHubEditFieldMap = React.useCallback(async (): Promise<IKMDataHubEditFieldMap> => {
    if (!props.context) {
      throw new Error('Context is not available.');
    }

    if (kmDataHubEditFieldMapRef.current) {
      return kmDataHubEditFieldMapRef.current;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const effectiveLibraryName = props.listTitle || DEFAULT_LIBRARY_NAME;
    const response = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/fields?$select=Title,InternalName,Hidden,Id,TextField&$top=5000`,
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
    }>;

    const normalize = (value?: string): string => (value || '').trim().toLowerCase();
    const findInternalName = (displayName: string, fallback: string): string => {
      const matchedField = fields.find((field) =>
        normalize(field.Title) === normalize(displayName) ||
        normalize(field.InternalName) === normalize(fallback)
      );

      return matchedField?.InternalName || fallback;
    };
    const findOptionalInternalName = (displayName: string, fallback: string): string => {
      const matchedField = fields.find((field) =>
        normalize(field.Title) === normalize(displayName) ||
        normalize(field.InternalName) === normalize(fallback)
      );

      return matchedField?.InternalName || '';
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

    const fieldMap: IKMDataHubEditFieldMap = {
      status: findInternalName(COLUMN_NAMES.status, COLUMN_NAMES.status),
      published: findInternalName(COLUMN_NAMES.published, COLUMN_NAMES.published),
      title: findInternalName(COLUMN_NAMES.title, COLUMN_NAMES.title),
      description: findInternalName(COLUMN_NAMES.description, COLUMN_NAMES.description),
      author: findInternalName('Author', COLUMN_NAMES.author),
      bu: findInternalName(COLUMN_NAMES.businessUnit, COLUMN_NAMES.businessUnit),
      department: findInternalName('Department / Sub Department', COLUMN_NAMES.department),
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
      reviewerComments:
        findOptionalInternalName('Reviewer Commands', 'ReviewerCommands') ||
        findOptionalInternalName('Reviewer Comments', COLUMN_NAMES.reviewerComments),
      reviewerCorrections: findOptionalInternalName('Reviewer Corrections', 'ReviewerCorrections'),
      contentRefreshDate: findInternalName('Content Refresh Date', COLUMN_NAMES.contentRefreshDate),
      views: findInternalName(COLUMN_NAMES.views, COLUMN_NAMES.views),
      likes: findInternalName(COLUMN_NAMES.likes, COLUMN_NAMES.likes),
      comments: findInternalName(COLUMN_NAMES.comments, COLUMN_NAMES.comments),
      downloads: findInternalName(COLUMN_NAMES.downloads, COLUMN_NAMES.downloads),
      follow: findInternalName(COLUMN_NAMES.follow, COLUMN_NAMES.follow),
      share: findInternalName(COLUMN_NAMES.share, COLUMN_NAMES.share),
      bookmark: findInternalName(COLUMN_NAMES.bookmark, COLUMN_NAMES.bookmark),
      projectId: findInternalName('Project ID', COLUMN_NAMES.projectId),
      versionFileName: findInternalName('Version File Name', COLUMN_NAMES.versionFileName),
      versionFileType: findInternalName('Version File Type', COLUMN_NAMES.versionFileType),
      edited: findInternalName('Edited', COLUMN_NAMES.edited),
      editedBy: findInternalName('Editor', COLUMN_NAMES.editedBy),
      modifiedBy: findInternalName('Modified By', COLUMN_NAMES.modifiedBy),
      docIcon: findInternalName('Doc Icon', COLUMN_NAMES.docIcon),
      fileLeafRef: findInternalName('File Name', COLUMN_NAMES.fileLeafRef)
    };

    kmDataHubEditFieldMapRef.current = fieldMap;
    return fieldMap;
  }, [props.context]);

  const ensureSiteUser = React.useCallback(async (upn: string): Promise<number> => {
    if (!props.context) {
      throw new Error('Context is not available.');
    }

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
      currentUserEmail ||
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
        FieldValue: buildPersonClaimsFieldValue(currentUserLogin)
      });
    }

    return formValues;
  }, [currentUserEmail, props.context, props.userEmail]);

  const appendKMReviewAuditFormValues = React.useCallback((
    formValues: Array<{ FieldName: string; FieldValue: string }>
  ): Array<{ FieldName: string; FieldValue: string }> => (
    appendFormValues(formValues, buildKMReviewAuditFormValues())
  ), [buildKMReviewAuditFormValues]);

  const updateKMDataHubTextFields = React.useCallback(async (
    itemId: number,
    metadata: Record<string, any>,
    fieldMap: IKMDataHubEditFieldMap,
    previousStatus?: string
  ): Promise<void> => {
    if (!props.context) {
      return;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const body: Record<string, any> = {
      [fieldMap.title]: metadata.title || '-',
      [fieldMap.description]: metadata.description || '-',
      [fieldMap.sensitiveTerms]: metadata.sensitiveTerms || ''
    };

    if (fieldMap.reviewerComments) {
      body[fieldMap.reviewerComments] = metadata.reviewerComments || '';
    }
    if (fieldMap.reviewerCorrections) {
      body[fieldMap.reviewerCorrections] = metadata.reviewerCorrections || '';
    }
    if (fieldMap.status && String(metadata.status || '').trim()) {
      body[fieldMap.status] = metadata.status;
    }
    appendPublishedDateUpdate(body, fieldMap, metadata.status, previousStatus);

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

    const formValues = appendKMReviewAuditFormValues(Object.keys(body)
      .map((fieldName) => ({
        FieldName: fieldName,
        FieldValue: body[fieldName]
      }))
      .filter((entry) => entry.FieldValue !== undefined && entry.FieldValue !== null));

    const effectiveLibraryName = props.listTitle || DEFAULT_LIBRARY_NAME;
    let response: SPHttpClientResponse | undefined;
    await updateWithoutVersion(
      props.context.spHttpClient,
      webUrl,
      itemId,
      async () => {
        response = await props.context!.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${itemId})/ValidateUpdateListItem`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: JSON.stringify({
              formValues,
              bNewDocumentUpdate: true
            })
          }
        );
      }
    );

    if (!response || !response.ok) {
      const errorText = response ? await response.text() : '';
      throw new Error(`Failed to update metadata text fields: ${response?.status || 'unknown'} ${errorText}`);
    }

    const responseJson = await response.json();
    const fieldResults = responseJson?.value || responseJson?.d?.ValidateUpdateListItem?.results || responseJson?.d?.results || [];
    const fieldErrors = fieldResults.filter((entry: any) => entry.HasException);
    if (fieldErrors.length > 0) {
      throw new Error(fieldErrors.map((entry: any) => `${entry.FieldName}: ${entry.ErrorMessage}`).join('; '));
    }
  }, [appendKMReviewAuditFormValues, props.context]);

  const updateKMDataHubTextFieldsRaw = React.useCallback(async (
    itemId: number,
    metadata: Record<string, any>,
    fieldMap: IKMDataHubEditFieldMap,
    previousStatus?: string
  ): Promise<void> => {
    if (!props.context) {
      return;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const body: Record<string, any> = {
      [fieldMap.title]: metadata.title || '-',
      [fieldMap.description]: metadata.description || '-',
      [fieldMap.sensitiveTerms]: metadata.sensitiveTerms || ''
    };

    if (fieldMap.reviewerComments) {
      body[fieldMap.reviewerComments] = metadata.reviewerComments || '';
    }
    if (fieldMap.reviewerCorrections) {
      body[fieldMap.reviewerCorrections] = metadata.reviewerCorrections || '';
    }
    if (fieldMap.status && String(metadata.status || '').trim()) {
      body[fieldMap.status] = metadata.status;
    }
    appendPublishedDateUpdate(body, fieldMap, metadata.status, previousStatus);

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

    const formValues = appendKMReviewAuditFormValues(Object.keys(body)
      .map((fieldName) => ({
        FieldName: fieldName,
        FieldValue: body[fieldName]
      }))
      .filter((entry) => entry.FieldValue !== undefined && entry.FieldValue !== null));

    const effectiveLibraryName = props.listTitle || DEFAULT_LIBRARY_NAME;
    console.log('formValues being sent:', JSON.stringify(formValues, null, 2));
    const response = await props.context.spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${itemId})/ValidateUpdateListItem`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: JSON.stringify({
          formValues,
          bNewDocumentUpdate: false
        })
      }
    );

    if (!response || !response.ok) {
      const errorText = response ? await response.text() : '';
      throw new Error(
        `Failed to update text fields: ${response?.status} ${errorText}`
      );
    }

    const responseJson = await response.json();
    const fieldResults = responseJson?.value ||
      responseJson?.d?.ValidateUpdateListItem?.results ||
      responseJson?.d?.results || [];
    const fieldErrors = fieldResults.filter((e: any) => e.HasException);
    if (fieldErrors.length > 0) {
      throw new Error(
        fieldErrors.map((e: any) => `${e.FieldName}: ${e.ErrorMessage}`).join('; ')
      );
    }
  }, [appendKMReviewAuditFormValues, props.context, props.listTitle]);

  const fetchMetadataSnippetRows = React.useCallback(async (itemId: number): Promise<void> => {
    if (!props.context) {
      setMetadataSnippetRows([]);
      return;
    }

    try {
      const fieldMap = await getKMDataHubEditFieldMap();
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const selectFields = [
        COLUMN_NAMES.title,
        fieldMap.title,
        COLUMN_NAMES.description,
        `${fieldMap.author}/Title`,
        fieldMap.documentType,
        fieldMap.client,
        fieldMap.geography,
        fieldMap.diseaseArea,
        fieldMap.therapyArea,
        fieldMap.sensitiveTerms,
        fieldMap.contentRefreshDate || COLUMN_NAMES.contentRefreshDate,
        fieldMap.views || COLUMN_NAMES.views,
        fieldMap.likes || COLUMN_NAMES.likes,
        fieldMap.comments || COLUMN_NAMES.comments,
        fieldMap.downloads || COLUMN_NAMES.downloads,
        fieldMap.follow || COLUMN_NAMES.follow,
        fieldMap.share || COLUMN_NAMES.share,
        fieldMap.bookmark || COLUMN_NAMES.bookmark,
        fieldMap.reviewerComments || COLUMN_NAMES.reviewerComments,
        fieldMap.projectId || COLUMN_NAMES.projectId,
        fieldMap.versionFileName || COLUMN_NAMES.versionFileName,
        fieldMap.versionFileType || COLUMN_NAMES.versionFileType,
        `${fieldMap.editedBy || COLUMN_NAMES.editedBy}/Title`,
        `${fieldMap.editedBy || COLUMN_NAMES.editedBy}/EMail`,
        `${fieldMap.editedBy || COLUMN_NAMES.editedBy}/Id`,
        fieldMap.docIcon || COLUMN_NAMES.docIcon,
        fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef
      ];
      const noteFields = [
        fieldMap.documentTypeNote,
        fieldMap.clientNote,
        fieldMap.geographyNote,
        fieldMap.diseaseAreaNote,
        fieldMap.therapyAreaNote
      ].filter(Boolean);

      const effectiveLibraryName = props.listTitle || DEFAULT_LIBRARY_NAME;
      const [itemResponse, textResponse] = await Promise.all([
        props.context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${itemId})?$select=${Array.from(new Set(selectFields.concat(noteFields))).join(',')}&$expand=${Array.from(new Set([fieldMap.author, fieldMap.editedBy || COLUMN_NAMES.editedBy])).join(',')}`,
          SPHttpClient.configurations.v1
        ),
        props.context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${itemId})/FieldValuesAsText`,
          SPHttpClient.configurations.v1
        )
      ]);

      if (!itemResponse.ok) {
        setMetadataSnippetRows([]);
        return;
      }

      const rawItem = await itemResponse.json();
      const item = rawItem?.d || rawItem || {};
      const textValuesJson = textResponse.ok ? await textResponse.json() : {};
      const textValues = textValuesJson?.d || textValuesJson || {};
      const fieldValue = (internalName: string, noteInternalName?: string): string =>
        formatTaxonomyDisplayValue(item[internalName]) ||
        formatTaxonomyDisplayValue(textValues[internalName]) ||
        (noteInternalName ? formatTaxonomyDisplayValue(item[noteInternalName]) || formatTaxonomyDisplayValue(textValues[noteInternalName]) : '');
      const directValue = (internalName?: string): string =>
        internalName
          ? formatTaxonomyDisplayValue(item[internalName]) || formatTaxonomyDisplayValue(textValues[internalName])
          : '';
      const modifiedByValue = item[fieldMap.editedBy || COLUMN_NAMES.editedBy] || item[COLUMN_NAMES.editedBy] || item.Editor;

      const rows: IMetadataSnippetRow[] = [
        { label: 'Document Type', value: fieldValue(fieldMap.documentType, fieldMap.documentTypeNote) },
        { label: 'Client', value: fieldValue(fieldMap.client, fieldMap.clientNote) },
        { label: 'Geography', value: fieldValue(fieldMap.geography, fieldMap.geographyNote) },
        { label: 'Disease Area', value: fieldValue(fieldMap.diseaseArea, fieldMap.diseaseAreaNote) },
        { label: 'Therapy Area', value: fieldValue(fieldMap.therapyArea, fieldMap.therapyAreaNote) },
        { label: 'Content Refresh Date', value: directValue(fieldMap.contentRefreshDate || COLUMN_NAMES.contentRefreshDate) },
        { label: 'Views', value: directValue(fieldMap.views || COLUMN_NAMES.views) },
        { label: 'Likes', value: directValue(fieldMap.likes || COLUMN_NAMES.likes) },
        { label: 'Comments', value: directValue(fieldMap.comments || COLUMN_NAMES.comments) },
        { label: 'Downloads', value: directValue(fieldMap.downloads || COLUMN_NAMES.downloads) },
        { label: 'Follow', value: directValue(fieldMap.follow || COLUMN_NAMES.follow) },
        { label: 'Share', value: directValue(fieldMap.share || COLUMN_NAMES.share) },
        { label: 'Bookmark', value: directValue(fieldMap.bookmark || COLUMN_NAMES.bookmark) },
        { label: 'Reviewer Comments', value: directValue(fieldMap.reviewerComments || COLUMN_NAMES.reviewerComments) },
        { label: 'Project ID', value: directValue(fieldMap.projectId || COLUMN_NAMES.projectId) },
        { label: 'Version File Name', value: directValue(fieldMap.versionFileName || COLUMN_NAMES.versionFileName) },
        { label: 'Version File Type', value: directValue(fieldMap.versionFileType || COLUMN_NAMES.versionFileType) },
        { label: 'Modified By', value: formatTaxonomyDisplayValue(modifiedByValue?.Title) },
        { label: 'Doc Icon', value: directValue(fieldMap.docIcon || COLUMN_NAMES.docIcon) },
        { label: 'File Name', value: directValue(fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef) }
      ];

      setMetadataSnippetRows(rows);
    } catch (error) {
      console.error('Error loading metadata snippet:', error);
      setMetadataSnippetRows([]);
    }
  }, [getKMDataHubEditFieldMap, props.context]);

  const handleDownload = async () => {
    if (!document || !props.context || shouldHideDownloadActions || isDownloadStarting) return;

    const webUrl = props.context.pageContext.web.absoluteUrl;
    let serverRelativeUrl = document.serverRelativeUrl;

    if (!serverRelativeUrl.startsWith('/')) {
      serverRelativeUrl = `/${serverRelativeUrl}`;
    }

    setIsDownloadStarting(true);
    try {
      await downloadSharePointFile(
        props.context,
        webUrl,
        serverRelativeUrl,
        buildTitleDownloadFileName(document.name, document.fileName, document.fileType)
      );
      void recordDownload();
    } catch (error) {
      console.error('Download error:', error);
    } finally {
      window.setTimeout(() => setIsDownloadStarting(false), 1200);
    }
  };

  const formatVersionDate = (value: string): string => {
    const date = value ? new Date(value) : null;
    if (!date || Number.isNaN(date.getTime())) {
      return 'Date unavailable';
    }

    const indiaDateParts = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Asia/Kolkata',
      day: 'numeric',
      month: 'short',
      year: 'numeric'
    }).formatToParts(date);
    const indiaTimeParts = new Intl.DateTimeFormat('en-US', {
      timeZone: 'Asia/Kolkata',
      hour: 'numeric',
      minute: '2-digit',
      hour12: true
    }).formatToParts(date);
    const getDatePart = (type: string): string =>
      indiaDateParts.find((part) => part.type === type)?.value || '';
    const getTimePart = (type: string): string =>
      indiaTimeParts.find((part) => part.type === type)?.value || '';
    const day = Number(getDatePart('day'));
    const suffix = day % 10 === 1 && day !== 11
      ? 'st'
      : day % 10 === 2 && day !== 12
        ? 'nd'
        : day % 10 === 3 && day !== 13
          ? 'rd'
          : 'th';
    const month = getDatePart('month');
    const year = getDatePart('year');
    const time = `${getTimePart('hour')}:${getTimePart('minute')} ${getTimePart('dayPeriod')}`.toLowerCase();

    return `${day}${suffix} ${month},${year} | ${time}`;
  };

  const handleVersionHistoryOpen = React.useCallback(async (): Promise<void> => {
    if (!document || !props.context) {
      return;
    }

    setIsVersionHistoryOpen(true);
    setIsVersionHistoryLoading(true);
    setVersionHistoryError('');

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const items = await getVersionHistory(props.context.spHttpClient, webUrl, document.id);
      const sortedItems = [...items].sort((a, b) => {
        if (a.isCurrentVersion !== b.isCurrentVersion) {
          return a.isCurrentVersion ? -1 : 1;
        }

        const versionA = Number.parseFloat(String(a.versionLabel || '').replace(/[^\d.]/g, ''));
        const versionB = Number.parseFloat(String(b.versionLabel || '').replace(/[^\d.]/g, ''));
        if (!Number.isNaN(versionA) && !Number.isNaN(versionB) && versionA !== versionB) {
          return versionB - versionA;
        }

        return new Date(b.created || 0).getTime() - new Date(a.created || 0).getTime();
      });
      setVersionHistoryItems(sortedItems);
    } catch (error) {
      setVersionHistoryError('Unable to load version history.');
      setVersionHistoryItems([]);
    } finally {
      setIsVersionHistoryLoading(false);
    }
  }, [document, props.context]);

  const handleVersionHistoryClose = React.useCallback((): void => {
    setIsVersionHistoryOpen(false);
  }, []);

  const handleVersionPreviewClose = React.useCallback((): void => {
    if (viewingVersionPreview?.isBlob && viewingVersionPreview.url) {
      URL.revokeObjectURL(viewingVersionPreview.url);
    }
    setViewingVersion(null);
    setViewingVersionPreview(null);
  }, [viewingVersionPreview]);

  const buildImmediateVersionDetail = React.useCallback((version: IVersionEntry): IVersionDetail | null => {
    if (!document) {
      return null;
    }

    return {
      versionLabel: version.versionLabel,
      versionId: version.id,
      isCurrentVersion: version.isCurrentVersion,
      created: version.created,
      modifiedDate: version.created,
      createdDate: version.created,
      title: document.name || document.fileName,
      status: document.status || '-',
      fileName: version.fileName,
      fileType: version.fileType,
      description: document.abstract || '-',
      author: version.createdBy,
      authorEmail: version.createdByEmail,
      modifiedBy: version.createdBy,
      modifiedByEmail: version.createdByEmail,
      createdBy: version.createdBy,
      createdByEmail: version.createdByEmail,
      published: '',
      sensitiveTerms: '-',
      contentRefreshDate: '',
      projectId: '-',
      documentType: '-',
      geography: '-',
      client: '-',
      diseaseArea: '-',
      therapyArea: '-',
      bu: '-',
      department: '-',
      url: version.versionUrl
    };
  }, [document]);

  const handleVersionView = React.useCallback(async (version: IVersionEntry): Promise<void> => {
    if (!document || !props.context) return;
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const immediateDetail = buildImmediateVersionDetail(version);
    if (viewingVersionPreview?.isBlob && viewingVersionPreview.url) {
      URL.revokeObjectURL(viewingVersionPreview.url);
    }
    setViewingVersion(null);
    setViewingVersionPreview(null);
    setIsVersionPreviewLoading(true);

    try {
      const detail = version.isCurrentVersion
        ? await fetchCurrentVersionDetail(props.context.spHttpClient, webUrl, document.id)
        : await fetchVersionDetail(props.context.spHttpClient, webUrl, document.id, version.id);

      const previewKind = getDetailVersionPreviewKind(detail?.fileType || version.fileType);
      let preview: TDetailVersionPreview = {
        url: '',
        kind: 'unsupported',
        isBlob: false,
        message: 'Preview is not available for this version.'
      };

      if (version.isCurrentVersion && previewKind === 'office') {
        preview = {
          url: version.versionUrl,
          kind: 'office',
          isBlob: false
        };
      } else if (!version.isCurrentVersion && previewKind === 'office') {
        preview = {
          url: buildFileApiUrl(webUrl, version.fileRef, `/Versions(${version.id})/$value`),
          kind: 'unsupported',
          isBlob: false,
          message: 'Preview not available for this version. Use the Download button to view this file.'
        };
      } else if (previewKind !== 'unsupported') {
        const downloadUrl = version.isCurrentVersion
          ? buildFileApiUrl(webUrl, version.fileRef, '/$value')
          : buildFileApiUrl(webUrl, version.fileRef, `/Versions(${version.id})/$value`);
        const resp = await props.context.spHttpClient.get(downloadUrl, SPHttpClient.configurations.v1);
        if (resp.ok) {
          const blob = await resp.blob();
          const mimeType = getDetailVersionMimeType(detail?.fileType || version.fileType);
          const previewBlob = blob.type === mimeType ? blob : new Blob([blob], { type: mimeType });
          preview = {
            url: URL.createObjectURL(previewBlob),
            kind: previewKind,
            isBlob: true
          };
        }
      }

      setViewingVersion(detail || immediateDetail);
      setViewingVersionPreview(preview);
    } catch (err) {
      console.error('Version detail load failed:', err);
      setViewingVersion(immediateDetail);
      setViewingVersionPreview({
        url: '',
        kind: 'unsupported',
        isBlob: false,
        message: 'Failed to load version preview.'
      });
    } finally {
      setIsVersionPreviewLoading(false);
    }
  }, [buildImmediateVersionDetail, document, props.context, viewingVersionPreview]);

  const getVersionDownloadKey = React.useCallback((version: IVersionEntry): string => (
    `${version.isCurrentVersion ? 'current' : 'version'}-${version.id || version.versionLabel || version.fileRef}`
  ), []);

  const handleVersionDownload = React.useCallback(async (version: IVersionEntry): Promise<void> => {
    if (!document || !props.context) return;
    const versionDownloadKey = getVersionDownloadKey(version);
    if (downloadingVersionKeysRef.current.has(versionDownloadKey)) {
      return;
    }

    downloadingVersionKeysRef.current.add(versionDownloadKey);
    setDownloadingVersionKeys((current) => ({ ...current, [versionDownloadKey]: true }));

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const fileExt = version.fileType?.toLowerCase() || '';
    const fileName = version.fileName || `document.${fileExt}`;
    const fileNameWithExt = fileName.includes('.')
      ? fileName.substring(0, fileName.lastIndexOf('.')) + '.' + fileExt
      : `${fileName}.${fileExt}`;

    if (version.isCurrentVersion) {
      try {
        await downloadSharePointFile(
          props.context,
          webUrl,
          version.fileRef,
          fileNameWithExt
        );
      } catch (err) {
        console.error('Download error:', err);
      } finally {
        window.setTimeout(() => {
          downloadingVersionKeysRef.current.delete(versionDownloadKey);
          setDownloadingVersionKeys((current) => {
            const next = { ...current };
            delete next[versionDownloadKey];
            return next;
          });
        }, 1200);
      }
      return;
    }

    const downloadUrl = version.isCurrentVersion
      ? buildFileApiUrl(webUrl, version.fileRef, '/$value')
      : buildFileApiUrl(webUrl, version.fileRef, `/Versions(${version.id})/$value`);

    if (!downloadUrl) return;

    try {
      const resp = await props.context.spHttpClient.get(downloadUrl, SPHttpClient.configurations.v1);
      if (!resp.ok) return;
      const blob = await resp.blob();
      const blobUrl = URL.createObjectURL(blob);
      const anchor = window.document.createElement('a');
      anchor.href = blobUrl;
      anchor.download = fileNameWithExt;
      window.document.body.appendChild(anchor);
      anchor.click();
      window.document.body.removeChild(anchor);
      URL.revokeObjectURL(blobUrl);
    } catch (err) {
      console.error('Download error:', err);
    } finally {
      downloadingVersionKeysRef.current.delete(versionDownloadKey);
      setDownloadingVersionKeys((current) => {
        const next = { ...current };
        delete next[versionDownloadKey];
        return next;
      });
    }
  }, [document, getVersionDownloadKey, props.context]);

  const handleEdit = async (): Promise<void> => {
    if (!document || !props.context || !canEditOrUpdateDocument) {
      return;
    }

    setEditDialogOpen(true);
    setEditError(null);
    const hasLoadedEditValues = editPrefetchDocumentIdRef.current === document.id && !!editInitialValues && !!taxonomyOptions;
    setEditInitialValues(hasLoadedEditValues ? editInitialValues : null);
    setEditLoading(!hasLoadedEditValues);

    if (isOfficeDocument) {
      void releaseOfficePreviewForEdit();
    }

    if (editPrefetchInFlightRef.current) {
      await editPrefetchInFlightRef.current;
      return;
    }

    try {
      const prefetchPromise = (async () => {
        const [loadedTaxonomyOptions, kmsUsers, fieldMap] = await Promise.all([
          loadTaxonomyOptions(),
          loadKmsUsers(),
          getKMDataHubEditFieldMap()
        ]);
        if (!loadedTaxonomyOptions) {
          throw new Error('Unable to load taxonomy options.');
        }

        const webUrl = props.context.pageContext.web.absoluteUrl;
        const selectFields = [
          'Id',
          COLUMN_NAMES.title,
          COLUMN_NAMES.description,
          `${fieldMap.author}/Title`,
          `${fieldMap.author}/EMail`,
          `${fieldMap.author}/Name`,
          `${fieldMap.author}/Id`,
          `${fieldMap.author}Id`,
          fieldMap.sensitiveTerms,
          fieldMap.bu,
          fieldMap.department,
          KM_DATA_HUB_BU_FIELD_INTERNAL_NAME,
          KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME,
          fieldMap.documentType,
          fieldMap.client,
          fieldMap.geography,
          fieldMap.diseaseArea,
          fieldMap.therapyArea
        ].filter(Boolean);
        const noteFields = [
          fieldMap.documentTypeNote,
          fieldMap.clientNote,
          fieldMap.geographyNote,
          fieldMap.diseaseAreaNote,
          fieldMap.therapyAreaNote
        ].filter(Boolean);

        const effectiveLibraryName = props.listTitle || DEFAULT_LIBRARY_NAME;
        const [response, textResponse] = await Promise.all([
          props.context.spHttpClient.get(
            `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${document.id})?$select=${selectFields.concat(noteFields).join(',')}&$expand=${fieldMap.author}`,
            SPHttpClient.configurations.v1
          ),
          props.context.spHttpClient.get(
            `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${document.id})/FieldValuesAsText`,
            SPHttpClient.configurations.v1
          )
        ]);

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to load document metadata: ${response.status} ${errorText}`);
        }

        const rawItem = await response.json();
        const item = rawItem?.d || rawItem || {};
        const textValuesJson = textResponse.ok ? await textResponse.json() : {};
        const textValues = textValuesJson?.d || textValuesJson || {};
        const resolveTermList = (internalName: string, rawValue: any, fallbackText: string, options: ITaxonomyTerm[]) => {
          const displayTextTerms = extractTermsFromValue(textValues[internalName], options);
          if (displayTextTerms.length > 0) {
            return uniqueTerms(displayTextTerms);
          }

          const rawTerms = extractTermsFromValue(rawValue, options);
          return rawTerms.length > 0
            ? uniqueTerms(rawTerms)
            : uniqueTerms(extractTermsFromValue(fallbackText, options));
        };
        const resolveAuthorUpns = (): string[] => {
          const authorObject = item[fieldMap.author];
          const authorEntries = extractPersonEntries(authorObject);
          const resolvedUpns: string[] = [];

          authorEntries.forEach((entry: any) => {
            const directCandidates = [
              entry?.EMail,
              entry?.Email,
              entry?.Name,
              entry?.UserPrincipalName,
              entry?.LoginName
            ].filter((value): value is string => typeof value === 'string' && value.trim().length > 0);

            directCandidates.forEach((candidate) => {
              const normalizedCandidate = candidate.trim().toLowerCase();
              const matchedUser = kmsUsers.find((user) =>
                user.upn.toLowerCase() === normalizedCandidate ||
                user.email.toLowerCase() === normalizedCandidate ||
                user.displayName.toLowerCase() === normalizedCandidate
              );

              if (matchedUser && resolvedUpns.indexOf(matchedUser.upn) === -1) {
                resolvedUpns.push(matchedUser.upn);
              }
            });

            const authorTitle = entry?.Title || '';
            if (authorTitle) {
              const matchedByDisplayName = kmsUsers.find((user) => user.displayName.toLowerCase() === authorTitle.toLowerCase());
              if (matchedByDisplayName && resolvedUpns.indexOf(matchedByDisplayName.upn) === -1) {
                resolvedUpns.push(matchedByDisplayName.upn);
              }
            }
          });

          if (resolvedUpns.length > 0) {
            return resolvedUpns;
          }

          const fallbackCandidates = String(textValues[fieldMap.author] || '')
            .split(/[;,]/)
            .map((value) => value.trim())
            .filter(Boolean);

          fallbackCandidates.forEach((candidate) => {
            const normalizedCandidate = candidate.toLowerCase();
            const matchedUser = kmsUsers.find((user) =>
              user.upn.toLowerCase() === normalizedCandidate ||
              user.email.toLowerCase() === normalizedCandidate ||
              user.displayName.toLowerCase() === normalizedCandidate
            );

            if (matchedUser && resolvedUpns.indexOf(matchedUser.upn) === -1) {
              resolvedUpns.push(matchedUser.upn);
            }
          });

          return uniqueStrings(resolvedUpns);
        };
      const getFieldTextFallback = (internalName?: string, noteInternalName?: string): string =>
        [
          internalName ? item[internalName] : '',
          internalName ? textValues[internalName] : '',
          noteInternalName ? item[noteInternalName] : '',
          noteInternalName ? textValues[noteInternalName] : ''
        ].filter(Boolean).join(';#');

      const buDepartmentFallbackText = [
        getFieldTextFallback(fieldMap.department),
        getFieldTextFallback(fieldMap.bu)
      ].filter(Boolean).join('; ');
      const buDepartmentTerms = resolveBuDepartmentTermsFromSeparateFields(item, textValues, loadedTaxonomyOptions.buDepartment);
      const documentTypeTerms = resolveTermList(fieldMap.documentType, item[fieldMap.documentType], getFieldTextFallback(fieldMap.documentType, fieldMap.documentTypeNote), loadedTaxonomyOptions.documentType);
      const clientTerms = resolveTermList(fieldMap.client, item[fieldMap.client], getFieldTextFallback(fieldMap.client, fieldMap.clientNote), loadedTaxonomyOptions.client);
      const geographyTerms = resolveTermList(fieldMap.geography, item[fieldMap.geography], getFieldTextFallback(fieldMap.geography, fieldMap.geographyNote), loadedTaxonomyOptions.geography);
      const diseaseAreaTerms = resolveTermList(fieldMap.diseaseArea, item[fieldMap.diseaseArea], getFieldTextFallback(fieldMap.diseaseArea, fieldMap.diseaseAreaNote), loadedTaxonomyOptions.diseaseArea);
      const therapyAreaTerms = resolveTermList(fieldMap.therapyArea, item[fieldMap.therapyArea], getFieldTextFallback(fieldMap.therapyArea, fieldMap.therapyAreaNote), loadedTaxonomyOptions.therapyArea);

      setEditInitialValues({
        title: item[fieldMap.title] || item.Title || textValues[fieldMap.title] || document.name || '',
        description: item[fieldMap.description] || item.Description || textValues[fieldMap.description] || document.abstract || '',
        selectedAuthorUpns: resolveAuthorUpns(),
        buDepartmentTerms,
        buDepartmentTerm: buDepartmentTerms[0] || null,
        documentTypeTerms,
        documentTypeTerm: documentTypeTerms[0] || null,
        clientTerms,
        clientTerm: clientTerms[0] || null,
        geographyTerms,
        geographyTerm: geographyTerms[0] || null,
        diseaseAreaTerms,
        diseaseAreaTerm: diseaseAreaTerms[0] || null,
        therapyAreaTerms,
        therapyAreaTerm: therapyAreaTerms[0] || null,
        sensitiveTerms: item[fieldMap.sensitiveTerms] || textValues[fieldMap.sensitiveTerms] || '',
        reviewerComments:
          (fieldMap.reviewerComments &&
            (item[fieldMap.reviewerComments] || textValues[fieldMap.reviewerComments])) ||
          '',
        reviewerCorrections:
          (fieldMap.reviewerCorrections &&
            (item[fieldMap.reviewerCorrections] || textValues[fieldMap.reviewerCorrections])) ||
          '',
        status: item[fieldMap.status] || textValues[fieldMap.status] || document.status || 'Under Review'
      });
      editPrefetchDocumentIdRef.current = document.id;
      })();

      editPrefetchInFlightRef.current = prefetchPromise;
      await prefetchPromise;
    } catch (error) {
      console.error('Error preparing edit dialog:', error);
      setEditError(error instanceof Error ? error.message : 'Unable to load document details for editing.');
    } finally {
      editPrefetchInFlightRef.current = null;
      setEditLoading(false);
    }
  };

  const handleCloseEdit = React.useCallback(() => {
    if (editSaving) {
      return;
    }

    setEditDialogOpen(false);
    setEditInitialValues(null);
    setEditError(null);
    setEditLoading(false);
    restoreOfficePreviewAfterEdit();
  }, [editSaving, restoreOfficePreviewAfterEdit]);

  React.useEffect(() => {
    if (!editDialogOpen) {
      return;
    }

    const resetEditScroll = (): void => {
      const editBody = editMetadataBodyRef.current;
      if (!editBody) {
        return;
      }

      editBody.scrollTop = 0;
      editBody.scrollLeft = 0;
    };

    resetEditScroll();
    window.requestAnimationFrame(resetEditScroll);
    const resetTimeout = window.setTimeout(resetEditScroll, 50);

    return () => {
      window.clearTimeout(resetTimeout);
    };
  }, [editDialogOpen]);

  React.useEffect(() => {
    if (!document || !props.context) {
      return;
    }

    if (editPrefetchDocumentIdRef.current !== document.id) {
      setEditInitialValues(null);
    }

    if (editPrefetchDocumentIdRef.current === document.id || editPrefetchInFlightRef.current) {
      return;
    }

    editPrefetchInFlightRef.current = (async () => {
      try {
        const [loadedTaxonomyOptions, kmsUsers, fieldMap] = await Promise.all([
          loadTaxonomyOptions(),
          loadKmsUsers(),
          getKMDataHubEditFieldMap()
        ]);
        if (!loadedTaxonomyOptions) {
          return;
        }

        const webUrl = props.context.pageContext.web.absoluteUrl;
        const selectFields = [
          'Id',
          COLUMN_NAMES.title,
          COLUMN_NAMES.description,
          `${fieldMap.author}/Title`,
          `${fieldMap.author}/EMail`,
          `${fieldMap.author}/Name`,
          `${fieldMap.author}/Id`,
          `${fieldMap.author}Id`,
          fieldMap.status,
          fieldMap.sensitiveTerms,
          fieldMap.reviewerComments,
          fieldMap.reviewerCorrections,
          fieldMap.bu,
          fieldMap.department,
          KM_DATA_HUB_BU_FIELD_INTERNAL_NAME,
          KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME,
          fieldMap.documentType,
          fieldMap.client,
          fieldMap.geography,
          fieldMap.diseaseArea,
          fieldMap.therapyArea
        ].filter(Boolean);
        const noteFields = [
          fieldMap.documentTypeNote,
          fieldMap.clientNote,
          fieldMap.geographyNote,
          fieldMap.diseaseAreaNote,
          fieldMap.therapyAreaNote
        ].filter(Boolean);

        const effectiveLibraryName = props.listTitle || DEFAULT_LIBRARY_NAME;
        const [response, textResponse] = await Promise.all([
          props.context.spHttpClient.get(
            `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${document.id})?$select=${selectFields.concat(noteFields).join(',')}&$expand=${fieldMap.author}`,
            SPHttpClient.configurations.v1
          ),
          props.context.spHttpClient.get(
            `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${document.id})/FieldValuesAsText`,
            SPHttpClient.configurations.v1
          )
        ]);

        if (!response.ok) {
          return;
        }

        const rawItem = await response.json();
        const item = rawItem?.d || rawItem || {};
        const textValuesJson = textResponse.ok ? await textResponse.json() : {};
        const textValues = textValuesJson?.d || textValuesJson || {};
        const resolveTermList = (internalName: string, rawValue: any, fallbackText: string, options: ITaxonomyTerm[]) => {
          const displayTextTerms = extractTermsFromValue(textValues[internalName], options);
          if (displayTextTerms.length > 0) {
            return uniqueTerms(displayTextTerms);
          }

          const rawTerms = extractTermsFromValue(rawValue, options);
          return rawTerms.length > 0
            ? uniqueTerms(rawTerms)
            : uniqueTerms(extractTermsFromValue(fallbackText, options));
        };
        const resolveAuthorUpns = (): string[] => {
          const authorObject = item[fieldMap.author];
          const authorEntries = extractPersonEntries(authorObject);
          const resolvedUpns: string[] = [];

          authorEntries.forEach((entry: any) => {
            const directCandidates = [
              entry?.EMail,
              entry?.Email,
              entry?.Name,
              entry?.UserPrincipalName,
              entry?.LoginName
            ].filter((value): value is string => typeof value === 'string' && value.trim().length > 0);

            directCandidates.forEach((candidate) => {
              const normalizedCandidate = candidate.trim().toLowerCase();
              const matchedUser = kmsUsers.find((user) =>
                user.upn.toLowerCase() === normalizedCandidate ||
                user.email.toLowerCase() === normalizedCandidate ||
                user.displayName.toLowerCase() === normalizedCandidate
              );

              if (matchedUser && resolvedUpns.indexOf(matchedUser.upn) === -1) {
                resolvedUpns.push(matchedUser.upn);
              }
            });

            const authorTitle = entry?.Title || '';
            if (authorTitle) {
              const matchedByDisplayName = kmsUsers.find((user) => user.displayName.toLowerCase() === authorTitle.toLowerCase());
              if (matchedByDisplayName && resolvedUpns.indexOf(matchedByDisplayName.upn) === -1) {
                resolvedUpns.push(matchedByDisplayName.upn);
              }
            }
          });

          if (resolvedUpns.length > 0) {
            return resolvedUpns;
          }

          const fallbackCandidates = String(textValues[fieldMap.author] || '')
            .split(/[;,]/)
            .map((value) => value.trim())
            .filter(Boolean);

          fallbackCandidates.forEach((candidate) => {
            const normalizedCandidate = candidate.toLowerCase();
            const matchedUser = kmsUsers.find((user) =>
              user.upn.toLowerCase() === normalizedCandidate ||
              user.email.toLowerCase() === normalizedCandidate ||
              user.displayName.toLowerCase() === normalizedCandidate
            );

            if (matchedUser && resolvedUpns.indexOf(matchedUser.upn) === -1) {
              resolvedUpns.push(matchedUser.upn);
            }
          });

          return uniqueStrings(resolvedUpns);
        };
        const getFieldTextFallback = (internalName?: string, noteInternalName?: string): string =>
          [
            internalName ? item[internalName] : '',
            internalName ? textValues[internalName] : '',
            noteInternalName ? item[noteInternalName] : '',
            noteInternalName ? textValues[noteInternalName] : ''
          ].filter(Boolean).join(';#');

        const buDepartmentFallbackText = [
          getFieldTextFallback(fieldMap.department),
          getFieldTextFallback(fieldMap.bu)
        ].filter(Boolean).join('; ');
        const buDepartmentTerms = resolveBuDepartmentTermsFromSeparateFields(item, textValues, loadedTaxonomyOptions.buDepartment);
        const documentTypeTerms = resolveTermList(fieldMap.documentType, item[fieldMap.documentType], getFieldTextFallback(fieldMap.documentType, fieldMap.documentTypeNote), loadedTaxonomyOptions.documentType);
        const clientTerms = resolveTermList(fieldMap.client, item[fieldMap.client], getFieldTextFallback(fieldMap.client, fieldMap.clientNote), loadedTaxonomyOptions.client);
        const geographyTerms = resolveTermList(fieldMap.geography, item[fieldMap.geography], getFieldTextFallback(fieldMap.geography, fieldMap.geographyNote), loadedTaxonomyOptions.geography);
        const diseaseAreaTerms = resolveTermList(fieldMap.diseaseArea, item[fieldMap.diseaseArea], getFieldTextFallback(fieldMap.diseaseArea, fieldMap.diseaseAreaNote), loadedTaxonomyOptions.diseaseArea);
        const therapyAreaTerms = resolveTermList(fieldMap.therapyArea, item[fieldMap.therapyArea], getFieldTextFallback(fieldMap.therapyArea, fieldMap.therapyAreaNote), loadedTaxonomyOptions.therapyArea);

        setEditInitialValues({
          title: item[fieldMap.title] || item.Title || textValues[fieldMap.title] || document.name || '',
          description: item[fieldMap.description] || item.Description || textValues[fieldMap.description] || document.abstract || '',
          selectedAuthorUpns: resolveAuthorUpns(),
          buDepartmentTerms,
          buDepartmentTerm: buDepartmentTerms[0] || null,
          documentTypeTerms,
          documentTypeTerm: documentTypeTerms[0] || null,
          clientTerms,
          clientTerm: clientTerms[0] || null,
          geographyTerms,
          geographyTerm: geographyTerms[0] || null,
          diseaseAreaTerms,
          diseaseAreaTerm: diseaseAreaTerms[0] || null,
          therapyAreaTerms,
          therapyAreaTerm: therapyAreaTerms[0] || null,
          sensitiveTerms: item[fieldMap.sensitiveTerms] || textValues[fieldMap.sensitiveTerms] || '',
          reviewerComments:
            (fieldMap.reviewerComments &&
              (item[fieldMap.reviewerComments] || textValues[fieldMap.reviewerComments])) ||
            '',
          reviewerCorrections:
            (fieldMap.reviewerCorrections &&
              (item[fieldMap.reviewerCorrections] || textValues[fieldMap.reviewerCorrections])) ||
            '',
          status: item[fieldMap.status] || textValues[fieldMap.status] || document.status || 'Under Review'
        });

        editPrefetchDocumentIdRef.current = document.id;
      } catch (error) {
        console.error('Error preloading edit dialog:', error);
      } finally {
        editPrefetchInFlightRef.current = null;
      }
    })();
  }, [document, getKMDataHubEditFieldMap, loadKmsUsers, loadTaxonomyOptions, props.context, props.listTitle]);

  const handleEditSubmit = React.useCallback(async (formData: Record<string, any>): Promise<void> => {
    if (!props.context || !document || !canEditOrUpdateDocument) {
      return;
    }

    setEditSaving(true);
    setEditError(null);

    const persistMetadata = async (): Promise<void> => {
      const webUrl = props.context!.pageContext.web.absoluteUrl;
      const fieldMap = await getKMDataHubEditFieldMap();
      const taxonomyFormValues = buildKMDataHubTaxonomyFormValues(formData, fieldMap);
      const effectiveLibraryName = props.listTitle || DEFAULT_LIBRARY_NAME;

      await updateWithoutVersion(
        props.context!.spHttpClient,
        webUrl,
        document.id,
        async () => {
          await updateKMDataHubTextFieldsRaw(
            document.id,
            formData,
            fieldMap,
            document.status
          );

          if (taxonomyFormValues.length > 0) {
            let response: SPHttpClientResponse | undefined;
            response = await props.context!.spHttpClient.post(
              `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${document.id})/ValidateUpdateListItem`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  Accept: 'application/json;odata=verbose',
                  'Content-Type': 'application/json;odata=verbose',
                  'odata-version': ''
                },
                body: JSON.stringify({
                  formValues: taxonomyFormValues,
                  bNewDocumentUpdate: false
                })
              }
            );

            if (!response || !response.ok) {
              const errorText = response ? await response.text() : '';
              throw new Error(
                `Failed to update taxonomy: ${response?.status} ${errorText}`
              );
            }

            const responseJson = await response.json();
            const fieldResults = responseJson?.value ||
              responseJson?.d?.ValidateUpdateListItem?.results ||
              responseJson?.d?.results || [];
            const fieldErrors = fieldResults.filter(
              (entry: any) => entry.HasException
            );
            if (fieldErrors.length > 0) {
              throw new Error(
                fieldErrors.map(
                  (entry: any) => `${entry.FieldName}: ${entry.ErrorMessage}`
                ).join('; ')
              );
            }
          }

          await props.context!.spHttpClient.post(
            `${webUrl}/_api/web/lists/getbytitle('${effectiveLibraryName}')/items(${document.id})/ValidateUpdateListItem`,
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
                      FieldName: COLUMN_NAMES.versionFileName,
                      FieldValue: document.fileName || ''
                  },
                    {
                      FieldName: COLUMN_NAMES.versionFileType,
                      FieldValue: (document.fileType || '').toLowerCase()
                    }
                  ]),
                  bNewDocumentUpdate: false
                })
            }
          );

          await updateItemUrlFieldRaw(
            props.context!.spHttpClient,
            webUrl,
            document.id,
            formData.title || document.name
          );
        }
      );
    };

    try {
      try {
        await persistMetadata();
      } catch (error) {
        if (!isSharedUseLockError(error) || !(await releaseOfficePreviewForEdit())) {
          throw error;
        }

        await persistMetadata();
      }

      const newStatus = String(formData.status || '').trim();
      const oldStatus = String(document.status || '').trim();
      if (newStatus.toLowerCase() === 'active') {
        await createDocumentMetricsRow(
          props.context,
          document.id,
          {},
          String(formData.title || document.name || document.fileName || '').trim() || document.name
        );
      }

      if (newStatus && newStatus !== oldStatus) {
        try {
          const webUrl = props.context!.pageContext.web.absoluteUrl;
          const resolvedCurrentUserId = currentUserId || Number((props.context!.pageContext as any)?.legacyPageContext?.userId || 0);
          await props.context!.spHttpClient.post(
            `${webUrl}/_api/web/lists/GetByTitle('${LIST_NAMES.auditLog}')/items`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'odata-version': ''
              },
              body: JSON.stringify({
                __metadata: { type: 'SP.Data.Audit_x0020_LogListItem' },
                Title: document.name || document.fileName || '',
                FileName: document.fileName || document.fileRef?.split('/').pop() || document.name || '',
                Action: newStatus,
                PerformedById: resolvedCurrentUserId,
                TimeStamp: new Date().toISOString()
              })
            }
          );
        } catch (auditError) {
          console.warn('Audit log entry creation failed:', auditError);
        }
      }

      const createNotificationQueueItem = async (): Promise<void> => {
        try {
          const shouldNotify =
            (newStatus === 'Active' || newStatus === 'Reject') &&
            newStatus !== oldStatus;

          if (!shouldNotify) {
            return;
          }

          const webUrl = props.context!.pageContext.web.absoluteUrl;
          let creatorEmail = document.createdByEmail || '';
          let creatorName = document.createdBy || '';
          const resolveUserEmail = (userData: any): string => {
            const candidates = [
              userData?.d?.Email,
              userData?.Email,
              userData?.d?.UserPrincipalName,
              userData?.UserPrincipalName,
              userData?.d?.LoginName,
              userData?.LoginName
            ].filter((value): value is string => typeof value === 'string' && value.trim().length > 0);
            const emailCandidate = candidates.find((value) => value.includes('@')) || '';
            return emailCandidate.includes('|')
              ? emailCandidate.split('|').pop() || ''
              : emailCandidate;
          };

          if (document.creatorUserId > 0) {
            try {
              const userResp = await props.context!.spHttpClient.get(
                `${webUrl}/_api/web/getUserById(${document.creatorUserId})?$select=Email,Title,LoginName,UserPrincipalName`,
                SPHttpClient.configurations.v1
              );
              if (userResp.ok) {
                const userData = await userResp.json();
                creatorEmail = resolveUserEmail(userData) || creatorEmail;
                creatorName = userData?.d?.Title || userData?.Title || document.createdBy || '';
              }
            } catch (userErr) {
              console.warn('Could not fetch creator email:', userErr);
            }
          }

          const notificationDocumentUrl = new URL(
            `${PAGE_URLS.assets}?assetID=${document.id}`,
            SITE_URL
          );
          notificationDocumentUrl.searchParams.set('env', 'WebViewList');
          const documentUrl = notificationDocumentUrl.toString();
          const reviewerComments = formData.reviewerComments as string || '';
          const kmComments = formData.reviewerCorrections as string || '';
          const queueEntityType = await getListEntityType(LIST_NAMES.notificationQueue);

          if (!queueEntityType) {
            console.warn('Notification queue entity type could not be resolved.');
            return;
          }

          const queueResp = await props.context!.spHttpClient.post(
            `${webUrl}/_api/web/lists/getbytitle('${LIST_NAMES.notificationQueue}')/items`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
                'odata-version': ''
              },
              body: JSON.stringify({
                __metadata: {
                  type: queueEntityType
                },
                Title: document.name || document.fileName || '',
                DocumentId: document.id,
                AuthorEmail: creatorEmail,
                AuthorName: creatorName,
                NewStatus: newStatus,
                DocumentUrl: documentUrl,
                ReviewerComments: reviewerComments,
                KMComments: kmComments
              })
            }
          );

          if (queueResp.ok) {
            console.log('Notification queue item created ✅', {
              documentId: document.id,
              newStatus
            });
          } else {
            const errText = await queueResp.text();
            console.warn('Notification queue item failed:', queueResp.status, errText);
          }
        } catch (notifErr) {
          console.warn('Notification queue error:', notifErr);
        }
      };

      await createNotificationQueueItem();

      const selectedAuthorUpns = Array.isArray(formData.selectedAuthorUpns)
        ? formData.selectedAuthorUpns.filter(Boolean)
        : formData.selectedAuthorUpn
          ? [formData.selectedAuthorUpn].filter(Boolean)
          : [];
      const kmsUsers = selectedAuthorUpns.length > 0 ? await loadKmsUsers() : [];
      const optimisticAuthor = selectedAuthorUpns.length > 0
        ? selectedAuthorUpns.map((upn) => {
          const normalizedUpn = String(upn || '').trim().toLowerCase();
          const matchedUser = kmsUsers.find((user) =>
            user.upn.toLowerCase() === normalizedUpn ||
            user.email.toLowerCase() === normalizedUpn
          );

          return matchedUser?.displayName || String(upn || '').trim();
        }).filter(Boolean).join(', ')
        : document.author;

      setDocument((prevDoc) => prevDoc ? {
        ...prevDoc,
        name: String(formData.title || prevDoc.name || '').trim() || prevDoc.name,
        abstract: String(formData.description || prevDoc.abstract || '').trim() || prevDoc.abstract,
        author: optimisticAuthor || prevDoc.author
      } : prevDoc);

      await fetchMetadataSnippetRows(document.id);
      await new Promise((resolve) => window.setTimeout(resolve, 400));
      if (knowledgeSearchApiClient.isConfigured()) {
        // TODO: Restore Easy Auth/token enforcement when backend sync auth is enabled.
        knowledgeSearchApiClient
          .triggerSync('frontend_upload_or_metadata_update')
          .catch((searchError) => console.warn('Search backend sync trigger skipped/failed after metadata edit:', searchError));
      }

      emitDocumentDataChanged({
        documentIds: [document.id],
        reason: 'metadata'
      });
      const previousPublishedStatus = isPublishedStatus(document.status);
      const nextPublishedStatus = isPublishedStatus(formData.status);
      const statusChanged = String(formData.status || '').trim().toLowerCase() !==
        String(document.status || '').trim().toLowerCase();
      if (statusChanged && previousPublishedStatus !== nextPublishedStatus) {
        void touchRecentlyPublishedCacheForDocument(
          props.context.spHttpClient,
          props.context.pageContext.web.absoluteUrl,
          document.id,
          nextPublishedStatus
        );
      }
      await fetchDocumentDetails();
      handleCloseEdit();
    } catch (error) {
      console.error('Error saving document metadata:', error);
      setEditError(error instanceof Error ? error.message : 'Unable to save document changes.');
    } finally {
      setEditSaving(false);
    }
  }, [canEditOrUpdateDocument, currentUserId, document, fetchMetadataSnippetRows, fetchDocumentDetails, getKMDataHubEditFieldMap, handleCloseEdit, loadKmsUsers, props.context, releaseOfficePreviewForEdit, updateKMDataHubTextFieldsRaw]);

  const getShareAssetUrl = (): string => {
    const webUrl = (props.context?.pageContext?.web?.absoluteUrl || window.location.origin).replace(/\/$/, '');
    const assetUrl = new URL(`${webUrl}/SitePages/Assets.aspx`);
    if (document?.id) {
      assetUrl.searchParams.set('assetID', String(document.id));
    }
    assetUrl.searchParams.set('env', 'WebViewList');
    return assetUrl.toString();
  };

  const getShareAssetsPageUrl = (): string => {
    const webUrl = (props.context?.pageContext?.web?.absoluteUrl || window.location.origin).replace(/\/$/, '');
    const assetsPageUrl = new URL(`${webUrl}/SitePages/Assets.aspx`);
    assetsPageUrl.searchParams.set('env', 'WebViewList');
    return assetsPageUrl.toString();
  };

  const getShareLoginEmail = (value?: string): string => {
    const match = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i.exec(value || '');
    return match ? match[0] : '';
  };

  const isSystemShareUser = (value: string): boolean => {
    const normalizedValue = value.toLowerCase();
    return (
      normalizedValue.indexOf('app@sharepoint') !== -1 ||
      normalizedValue.indexOf('spo-grid-all-users') !== -1 ||
      normalizedValue.indexOf('everyone except external users') !== -1 ||
      normalizedValue.indexOf('system account') !== -1 ||
      normalizedValue.indexOf('sharepoint app') !== -1
    );
  };

  const loadShareSiteUsers = async (): Promise<SharePerson[]> => {
    if (shareSiteUsersRef.current) {
      return shareSiteUsersRef.current;
    }

    const siteUrl = props.context.pageContext.site.absoluteUrl;
    let nextUrl = `${siteUrl}/_api/web/siteusers?$select=Id,Title,Email,LoginName,PrincipalType,IsHiddenInUI&$top=5000`;
    const peopleByKey: Record<string, SharePerson> = {};

    while (nextUrl) {
      const response: SPHttpClientResponse = await props.context.spHttpClient.get(
        nextUrl,
        SPHttpClient.configurations.v1
      );
      const json = await response.json();
      const rows: ShareSiteUser[] = json?.value || json?.d?.results || [];

      rows.forEach((user: ShareSiteUser) => {
        const title = (user.Title || '').trim();
        const loginName = (user.LoginName || '').trim();
        const email = (user.Email || getShareLoginEmail(loginName)).trim();
        const principalType = Number(user.PrincipalType || 0);
        const searchableText = `${title} ${email} ${loginName}`;

        if (
          user.IsHiddenInUI ||
          !title ||
          isSystemShareUser(searchableText) ||
          (principalType && (principalType & 1) !== 1)
        ) {
          return;
        }

        const key = (email || loginName || title).toLowerCase();
        if (!key || peopleByKey[key]) return;

        peopleByKey[key] = {
          name: title,
          email,
          accountName: loginName || email,
          entityType: 'User'
        };
      });

      nextUrl = json?.['@odata.nextLink'] || json?.d?.__next || '';
    }

    const sitePeople = Object.keys(peopleByKey)
      .map(key => peopleByKey[key])
      .sort((left, right) => left.name.localeCompare(right.name));

    shareSiteUsersRef.current = sitePeople;
    return sitePeople;
  };

  const searchSharePeople = async (query: string): Promise<void> => {
    const trimmedQuery = query.trim();
    const requestId = ++shareSearchRequestIdRef.current;
    if (!trimmedQuery) {
      setSharePeopleResults([]);
      setSharePeopleLoading(false);
      setShareActivePersonIndex(0);
      setShareHoveredPersonIndex(null);
      return;
    }

    setSharePeopleLoading(true);
    try {
      const siteUrl = props.context.pageContext.site.absoluteUrl;
      const normalizedQuery = trimmedQuery.toLowerCase();
      const peopleByKey: Record<string, SharePerson> = {};
      const addPerson = (person: SharePerson): void => {
        const key = (person.email || person.accountName || person.name || '').toLowerCase();
        if (!key || peopleByKey[key]) return;
        peopleByKey[key] = person;
      };

      try {
        const resp = await props.context.spHttpClient.post(
          `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'odata-version': ''
            },
            body: JSON.stringify({
              queryParams: {
                AllowEmailAddresses: false,
                AllowMultipleEntities: true,
                AllUrlZones: false,
                MaximumEntitySuggestions: 8,
                PrincipalSource: 15,
                PrincipalType: 1,
                QueryString: trimmedQuery,
                Required: false,
                SharePointGroupID: 0,
                UrlZone: 0
              }
            })
          }
        );
        const json = await resp.json();
        const rawPeople = json?.ClientPeoplePickerSearchUser || json?.d?.ClientPeoplePickerSearchUser || json?.value || [];
        const pickerPeople = typeof rawPeople === 'string' ? JSON.parse(rawPeople || '[]') : rawPeople;
        (Array.isArray(pickerPeople) ? pickerPeople : []).forEach((item: any) => {
          const entityData = item.EntityData || {};
          const name = item.DisplayText || entityData.DisplayName || entityData.Title || item.Description || item.Key || '';
          const accountName = entityData.AccountName || item.Key || name;
          const email = entityData.Email || getShareLoginEmail(accountName) || getShareLoginEmail(item.Key) || (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(item.Description || '') ? item.Description : '');
          const entityType = item.EntityType || entityData.PrincipalType || '';
          const entityText = `${entityType || ''} ${accountName || ''} ${item.Description || ''}`.toLowerCase();
          const isGroup = /\b(group|secgroup|spgroup|distributionlist|dl)\b/.test(entityText);
          const isUser = !isGroup && (!!email || entityText.indexOf('membership') !== -1 || entityText.indexOf('user') !== -1);
          const subtitle = entityData.JobTitle || entityData.Department || '';

          if (!name || !accountName || !isUser) return;
          addPerson({
            name,
            email,
            accountName,
            subtitle,
            entityType
          });
        });
      } catch {
        // Fall back to the site-user cache below.
      }

      if (Object.keys(peopleByKey).length < 8) {
        try {
          const sitePeople = await loadShareSiteUsers();
          const matchedSitePeople = sitePeople
            .filter(person => {
              const searchableText = `${person.name} ${person.email} ${person.accountName}`.toLowerCase();
              return searchableText.indexOf(normalizedQuery) !== -1;
            })
            .sort((left, right) => {
              const leftName = left.name.toLowerCase();
              const rightName = right.name.toLowerCase();
              const leftEmail = (left.email || '').toLowerCase();
              const rightEmail = (right.email || '').toLowerCase();
              const leftRank =
                leftName.indexOf(normalizedQuery) === 0 ? 0 :
                  leftEmail.indexOf(normalizedQuery) === 0 ? 1 :
                    leftName.indexOf(normalizedQuery) !== -1 ? 2 : 3;
              const rightRank =
                rightName.indexOf(normalizedQuery) === 0 ? 0 :
                  rightEmail.indexOf(normalizedQuery) === 0 ? 1 :
                    rightName.indexOf(normalizedQuery) !== -1 ? 2 : 3;
              return leftRank - rightRank || leftName.localeCompare(rightName);
            })
            .slice(0, 8 - Object.keys(peopleByKey).length);

          matchedSitePeople.forEach(addPerson);
        } catch {
          // The native picker response above is the preferred source.
        }
      }

      const people = Object.keys(peopleByKey)
        .map(key => peopleByKey[key])
        .sort((left, right) => {
          const leftName = left.name.toLowerCase();
          const rightName = right.name.toLowerCase();
          const leftEmail = (left.email || '').toLowerCase();
          const rightEmail = (right.email || '').toLowerCase();
          const leftStarts = leftName.indexOf(normalizedQuery) === 0 ? 0 : leftEmail.indexOf(normalizedQuery) === 0 ? 1 : 2;
          const rightStarts = rightName.indexOf(normalizedQuery) === 0 ? 0 : rightEmail.indexOf(normalizedQuery) === 0 ? 1 : 2;
          return leftStarts - rightStarts || leftName.localeCompare(rightName);
        })
        .slice(0, 8);

      if (requestId !== shareSearchRequestIdRef.current) {
        return;
      }
      setSharePeopleResults(people);
      setShareActivePersonIndex(0);
      setShareHoveredPersonIndex(null);
    } catch {
      if (requestId !== shareSearchRequestIdRef.current) {
        return;
      }
      setSharePeopleResults([]);
      setShareActivePersonIndex(0);
      setShareHoveredPersonIndex(null);
    }
    if (requestId === shareSearchRequestIdRef.current) {
      setSharePeopleLoading(false);
    }
  };

  const renderSharePersonaAvatar = (size: number): React.ReactElement => {
    const headSize = Math.round(size * 0.34);
    const shoulderWidth = Math.round(size * 0.62);
    const shoulderHeight = Math.round(size * 0.28);
    return (
      <span
        aria-hidden="true"
        style={{
          width: size,
          height: size,
          borderRadius: '50%',
          background: 'rgb(200, 198, 196)',
          display: 'inline-flex',
          alignItems: 'center',
          justifyContent: 'center',
          position: 'relative',
          overflow: 'hidden',
          flexShrink: 0
        }}
      >
        <span
          style={{
            position: 'absolute',
            top: Math.round(size * 0.2),
            width: headSize,
            height: headSize,
            borderRadius: '50%',
            background: '#ffffff'
          }}
        />
        <span
          style={{
            position: 'absolute',
            bottom: Math.round(size * 0.16),
            width: shoulderWidth,
            height: shoulderHeight,
            borderRadius: `${shoulderHeight}px ${shoulderHeight}px 4px 4px`,
            background: '#ffffff'
          }}
        />
      </span>
    );
  };

  const selectSharePerson = (person: SharePerson): void => {
    if (!shareSelectedPeople.find(selectedPerson =>
      selectedPerson.accountName === person.accountName ||
      (!!person.email && selectedPerson.email === person.email)
    )) {
      setShareSelectedPeople(prev => [...prev, person]);
    }
    setSharePeopleQuery('');
    setSharePeopleResults([]);
    setShareActivePersonIndex(0);
    setShareHoveredPersonIndex(null);
    setShareValidationError('');
    window.setTimeout(() => sharePeopleInputRef.current?.focus(), 0);
  };

  const handleShareCopyLink = async (): Promise<void> => {
    const assetUrl = getShareAssetUrl();
    try {
      await navigator.clipboard.writeText(assetUrl);
    } catch {
      const textarea = window.document.createElement('textarea');
      textarea.value = assetUrl;
      window.document.body.appendChild(textarea);
      textarea.select();
      window.document.execCommand('copy');
      window.document.body.removeChild(textarea);
    }
    setShareCopied(true);
    void recordShare();
    setTimeout(() => setShareCopied(false), 2000);
  };

  const getShareInviteeText = (): string => {
    const invitees = shareSelectedPeople
      .map(person => person.name || person.email || person.accountName)
      .filter(Boolean);

    if (invitees.length <= 1) {
      return invitees[0] || 'the selected people';
    }
    if (invitees.length === 2) {
      return `${invitees[0]} and ${invitees[1]}`;
    }
    return `${invitees.slice(0, -1).join(', ')}, and ${invitees[invitees.length - 1]}`;
  };

  const getSharePersonEmail = (person: SharePerson): string => (
    (person.email || getShareLoginEmail(person.accountName)).trim()
  );

  const getShareRecipientPeople = (): SharePerson[] => {
    const peopleByEmail: Record<string, SharePerson> = {};

    shareSelectedPeople.forEach(person => {
      const email = getSharePersonEmail(person).toLowerCase();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || peopleByEmail[email]) {
        return;
      }
      peopleByEmail[email] = person;
    });

    return Object.keys(peopleByEmail).map(email => peopleByEmail[email]);
  };

  const escapeShareEmailHtml = (value?: string): string => (
    (value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
  );

  const buildShareInviteEmailHtml = (
    docName: string,
    assetUrl: string,
    assetsPageUrl: string,
    senderName: string,
    senderEmail: string,
    recipientPeople: SharePerson[]
  ): string => {
    const safeDocName = escapeShareEmailHtml(docName);
    const safeAssetUrl = escapeShareEmailHtml(assetUrl);
    const safeAssetsPageUrl = escapeShareEmailHtml(assetsPageUrl);
    const safeSenderName = escapeShareEmailHtml(senderName || 'A colleague');
    const safeSenderEmail = escapeShareEmailHtml(senderEmail);
    const safeRecipientText = escapeShareEmailHtml(
      recipientPeople
        .map(person => person.name || getSharePersonEmail(person))
        .filter(Boolean)
        .join(', ')
    );
    const safeMessage = escapeShareEmailHtml(shareMessage.trim()).replace(/\r?\n/g, '<br />');

    return `
      <div style="margin:0;padding:0;background:#f6f6f6;font-family:'Segoe UI',Arial,sans-serif;color:#323130;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;background:#f6f6f6;margin:0;padding:0;">
          <tr>
            <td align="center" style="padding:32px 16px;">
              <table role="presentation" width="640" cellspacing="0" cellpadding="0" style="width:640px;max-width:100%;border-collapse:collapse;background:#ffffff;border:1px solid #e1dfdd;border-radius:8px;overflow:hidden;">
                <tr>
                  <td style="padding:28px 32px 24px;text-align:center;border-bottom:1px solid #edebe9;">
                    <div style="width:48px;height:48px;margin:0 auto 18px;border-radius:50%;background:#fbf4f4;border:1px solid #e3afb2;color:#a4262c;font-size:25px;line-height:48px;font-weight:700;">&#8599;</div>
                    <div style="font-size:24px;line-height:32px;font-weight:600;color:#323130;">${safeSenderName} invited you to view</div>
                    <div style="font-size:20px;line-height:28px;font-weight:600;color:#323130;margin-top:6px;">&quot;${safeDocName}&quot;</div>
                  </td>
                </tr>
                <tr>
                  <td style="padding:28px 32px 10px;">
                    <p style="margin:0 0 18px;font-size:15px;line-height:22px;color:#3b3a39;">Here is a page ${safeSenderName} wants you to see in Knowledge Hub.</p>
                    ${safeMessage ? `<div style="margin:0 0 22px;padding:14px 16px;background:#faf9f8;border-left:4px solid #a4262c;border-radius:4px;font-size:14px;line-height:21px;color:#323130;">${safeMessage}</div>` : ''}
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin:0 0 24px;background:#faf9f8;border:1px solid #edebe9;border-radius:6px;">
                      <tr>
                        <td style="padding:12px 16px;width:100px;font-size:12px;line-height:18px;font-weight:700;color:#605e5c;text-transform:uppercase;">Title</td>
                        <td style="padding:12px 16px;font-size:14px;line-height:20px;color:#323130;">${safeDocName}</td>
                      </tr>
                      <tr>
                        <td style="padding:12px 16px;width:100px;font-size:12px;line-height:18px;font-weight:700;color:#605e5c;text-transform:uppercase;border-top:1px solid #edebe9;">From</td>
                        <td style="padding:12px 16px;font-size:14px;line-height:20px;color:#323130;border-top:1px solid #edebe9;">${safeSenderName}${safeSenderEmail ? ` &lt;${safeSenderEmail}&gt;` : ''}</td>
                      </tr>
                      <tr>
                        <td style="padding:12px 16px;width:100px;font-size:12px;line-height:18px;font-weight:700;color:#605e5c;text-transform:uppercase;border-top:1px solid #edebe9;">To</td>
                        <td style="padding:12px 16px;font-size:14px;line-height:20px;color:#323130;border-top:1px solid #edebe9;">${safeRecipientText}</td>
                      </tr>
                    </table>
                    <table role="presentation" cellspacing="0" cellpadding="0" align="center" style="border-collapse:collapse;margin:0 auto 22px;">
                      <tr>
                        <td style="padding:0 6px 10px;">
                          <a href="${safeAssetUrl}" style="display:inline-block;min-width:168px;text-align:center;background:#a4262c;color:#ffffff;text-decoration:none;border-radius:4px;padding:12px 20px;font-size:15px;line-height:20px;font-weight:600;">Open asset</a>
                        </td>
                        <td style="padding:0 6px 10px;">
                          <a href="${safeAssetsPageUrl}" style="display:inline-block;min-width:168px;text-align:center;background:#ffffff;color:#a4262c;text-decoration:none;border:1px solid #a4262c;border-radius:4px;padding:11px 20px;font-size:15px;line-height:20px;font-weight:600;">Browse assets</a>
                        </td>
                      </tr>
                    </table>
                    <p style="margin:0 0 18px;text-align:center;font-size:12px;line-height:18px;color:#605e5c;">This invite works only for people with existing access.</p>
                  </td>
                </tr>
                <tr>
                  <td style="padding:16px 32px;background:#faf9f8;border-top:1px solid #edebe9;font-size:12px;line-height:18px;color:#605e5c;">
                    This email was sent from Knowledge Hub. If the button does not open, copy this link into your browser:<br />
                    <a href="${safeAssetUrl}" style="color:#a4262c;text-decoration:none;word-break:break-all;">${safeAssetUrl}</a>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </div>
    `;
  };

  const handleShareSend = async (): Promise<void> => {
    if (!shareSelectedPeople.length) {
      setShareValidationError('Add people to share the link.');
      return;
    }

    const recipientPeople = getShareRecipientPeople();
    const shareRecipients = recipientPeople.map(getSharePersonEmail);

    if (!shareRecipients.length) {
      setShareValidationError('Selected people do not have email addresses.');
      return;
    }

    setShareValidationError('');
    setShareSending(true);
    try {
      if (!props.context.msGraphClientFactory) {
        setShareValidationError('Mail send is not available in this SharePoint context.');
        return;
      }

      const assetUrl = getShareAssetUrl();
      const assetsPageUrl = getShareAssetsPageUrl();
      const docName = document?.name || document?.fileName || 'Document';
      const senderName = props.context.pageContext.user.displayName;
      const senderEmail = normalizeIdentityValue(
        currentUserEmail ||
        props.context.pageContext.user.email ||
        props.context.pageContext.user.loginName ||
        ''
      );
      const graphClient = await props.context.msGraphClientFactory.getClient('3');
      await graphClient.api('/me/sendMail').version('v1.0').post({
        message: {
          subject: `${senderName} shared "${docName}" with you`,
          body: {
            contentType: 'HTML',
            content: buildShareInviteEmailHtml(docName, assetUrl, assetsPageUrl, senderName, senderEmail, recipientPeople)
          },
          toRecipients: recipientPeople.map(person => ({
            emailAddress: {
              address: getSharePersonEmail(person),
              name: person.name || getSharePersonEmail(person)
            }
          })),
          replyTo: senderEmail ? [{
            emailAddress: {
              address: senderEmail,
              name: senderName
            }
          }] : undefined
        },
        saveToSentItems: true
      });
      void recordShare();
      setShareSent(true);
    } catch (error) {
      console.warn('Share invite failed:', error);
      setShareValidationError('Unable to send invite email. Please try again.');
    } finally {
      setShareSending(false);
    }
  };

  const handleShare = () => {
    if (!document || !props.context) return;
    setShowShareDialog(true);
    setShareSelectedPeople([]);
    setShareMessage('');
    setSharePeopleQuery('');
    setSharePeopleResults([]);
    setShareSent(false);
    setShareCopied(false);
    setShareValidationError('');
    setShareFocusedField(null);
    setShareTooltip(null);
    setShareActivePersonIndex(0);
    setShareHoveredPersonIndex(null);
  };

  const getShareDialogTitle = (): string => {
    const title = document?.name || document?.fileName || 'Document';
    if (title.length <= 36) return title;
    return `${title.slice(0, 13)}...${title.slice(-17)}`;
  };

  const renderShareTooltip = (
    key: ShareTooltipKey,
    text: string,
    options: {
      top?: number | string;
      right?: number | string;
      bottom?: number | string;
      left?: number | string;
      transform?: string;
      width?: number;
      arrow: 'top' | 'bottom';
      arrowLeft?: number | string;
      arrowRight?: number | string;
      iconName?: string;
    }
  ): React.ReactElement | null => {
    if (shareTooltip !== key) return null;

    const tooltipStyle: React.CSSProperties = {
      position: 'absolute',
      zIndex: 100002,
      top: options.top,
      right: options.right,
      bottom: options.bottom,
      left: options.left,
      transform: options.transform,
      width: options.width,
      boxSizing: 'border-box',
      border: '1px solid rgb(237, 235, 233)',
      borderRadius: 3,
      background: '#ffffff',
      boxShadow: '0 2px 8px rgba(0, 0, 0, 0.18)',
      color: 'rgb(50, 49, 48)',
      display: 'inline-flex',
      alignItems: 'center',
      gap: 8,
      fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
      fontSize: 12,
      fontWeight: 400,
      lineHeight: '16px',
      padding: '8px 10px',
      pointerEvents: 'none',
      whiteSpace: options.width ? 'normal' : 'nowrap'
    };

    const arrowStyle: React.CSSProperties = {
      position: 'absolute',
      width: 10,
      height: 10,
      background: '#ffffff',
      transform: 'rotate(45deg)',
      left: options.arrowLeft,
      right: options.arrowRight,
      ...(options.arrowLeft === undefined && options.arrowRight === undefined
        ? { left: '50%', marginLeft: -5 }
        : {}),
      ...(options.arrow === 'top'
        ? {
            top: -6,
            borderLeft: '1px solid rgb(237, 235, 233)',
            borderTop: '1px solid rgb(237, 235, 233)'
          }
        : {
            bottom: -6,
            borderRight: '1px solid rgb(237, 235, 233)',
            borderBottom: '1px solid rgb(237, 235, 233)'
          })
    };

    return (
      <span role="tooltip" style={tooltipStyle}>
        <span aria-hidden="true" style={arrowStyle} />
        {options.iconName && (
          <span
            aria-hidden="true"
            style={{
              width: 18,
              height: 18,
              border: '1px solid rgb(0, 120, 212)',
              borderRadius: '50%',
              color: 'rgb(0, 120, 212)',
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              flexShrink: 0
            }}
          >
            <Icon iconName={options.iconName} styles={{ root: { fontSize: 11, lineHeight: '11px' } }} />
          </span>
        )}
        <span>{text}</span>
      </span>
    );
  };

  React.useEffect(() => {
    if (!showShareDialog) return;
    window.setTimeout(() => {
      sharePeopleInputRef.current?.focus();
    }, 0);
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setShowShareDialog(false);
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [showShareDialog]);

  React.useEffect(() => {
    const listElement = sharePeopleListRef.current;
    if (!listElement || shareActivePersonIndex < 0) return;
    const activeElement = listElement.querySelector<HTMLElement>(`[data-share-person-index="${shareActivePersonIndex}"]`);
    activeElement?.scrollIntoView({ block: 'nearest' });
  }, [shareActivePersonIndex]);

  const recordShare = async (): Promise<void> => {
    if (!props.context || !document || !currentUserId) return;
    if ((document.status || '').toLowerCase() !== 'active') return;

    try {
      const newCount = await recordUserEvent(props.context, document.id, currentUserId, 'Share', document.name);
      setShareCount(newCount);
    } catch (error) {
      console.error('Error recording share:', error);
    }
  };

  const fetchCurrentUser = async () => {
    if (!props.context) return;
    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const response = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1
      );
      if (response.ok) {
        const user = await response.json();
        const userEmail = normalizeIdentityValue(
          user?.Email ||
          user?.d?.Email ||
          user?.UserPrincipalName ||
          user?.d?.UserPrincipalName ||
          user?.LoginName ||
          user?.d?.LoginName ||
          ''
        );
        setCurrentUserId(Number(user?.Id || user?.d?.Id || 0) || resolvedCurrentUserId);
        setCurrentUserName(user?.Title || user?.d?.Title || user?.LoginName || user?.d?.LoginName || 'User');
        setCurrentUserEmail(userEmail);
        return;
      }
    } catch (error) {
      console.error('Error fetching current user:', error);
    }

    if (resolvedCurrentUserId > 0) {
      setCurrentUserId(resolvedCurrentUserId);
      setCurrentUserName(((props.context.pageContext as any)?.legacyPageContext?.userDisplayName as string) || 'User');
      setCurrentUserEmail(resolvedCurrentUserEmail);
    }
  };

  const ensureListExists = async (listName: string, fields: Array<{ name: string, type: string, required?: boolean }>) => {
    if (!props.context) return false;

    if (VERIFIED_LISTS_DETAIL_PAGE.has(listName)) {
      return true;
    }

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;

      // Check if list exists
      const checkResp = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (checkResp.ok) {
        VERIFIED_LISTS_DETAIL_PAGE.add(listName);
        console.log(`List "${listName}" already exists`);
        return true;
      }

      // List doesn't exist, create it
      console.log(`Creating list "${listName}"...`);

      const createListBody = {
        __metadata: { type: 'SP.List' },
        Title: listName,
        BaseTemplate: 100, // Custom list
        Description: `List for storing ${listName.toLowerCase()}`
      };

      const createResp = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/lists`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "odata-version": ""
          },
          body: JSON.stringify(createListBody)
        }
      );

      if (!createResp.ok) {
        const errorText = await createResp.text();
        console.error(`Failed to create list "${listName}":`, createResp.status, errorText);
        return false;
      }

      const listData = await createResp.json();
      const listId = listData.d.Id;

      console.log(`List "${listName}" created with ID: ${listId}`);

      // Wait a bit for the list to be ready
      await new Promise(resolve => setTimeout(resolve, 1000));

      // Add fields to the list
      for (const field of fields) {
        try {
          let fieldBody: any = {
            __metadata: { type: 'SP.Field' },
            Title: field.name,
            Required: field.required || false
          };

          if (field.type === 'Number') {
            fieldBody.FieldTypeKind = 9; // Number
            fieldBody = {
              ...fieldBody,
              __metadata: { type: 'SP.FieldNumber' }
            };
          } else if (field.type === 'Text') {
            fieldBody.FieldTypeKind = 2; // Text
            fieldBody.MaxLength = 255;
            fieldBody = {
              ...fieldBody,
              __metadata: { type: 'SP.FieldText' }
            };
          } else if (field.type === 'Note') {
            fieldBody.FieldTypeKind = 3; // Note (Multiple lines of text)
            fieldBody = {
              ...fieldBody,
              __metadata: { type: 'SP.FieldMultiLineText' }
            };
          }

          const fieldResp = await props.context.spHttpClient.post(
            `${webUrl}/_api/web/lists(guid'${listId}')/fields`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "odata-version": ""
              },
              body: JSON.stringify(fieldBody)
            }
          );

          if (!fieldResp.ok) {
            const errorText = await fieldResp.text();
            // Field might already exist, that's okay
            if (!errorText.includes('already exists') && !errorText.includes('duplicate')) {
              console.warn(`Failed to add field "${field.name}" to "${listName}":`, errorText);
            }
          } else {
            console.log(`Field "${field.name}" added to "${listName}"`);
          }

          // Wait a bit between field additions
          await new Promise(resolve => setTimeout(resolve, 500));
        } catch (fieldError) {
          console.warn(`Error adding field "${field.name}":`, fieldError);
        }
      }

      VERIFIED_LISTS_DETAIL_PAGE.add(listName);
      return true;
    } catch (error) {
      console.error(`Error ensuring list "${listName}" exists:`, error);
      return false;
    }
  };

  const fetchLikes = async () => {
    if (!props.context || !document || !currentUserId) return;
    try {
      const [metrics, interactionState] = await Promise.all([
        getDocumentMetrics(props.context, document.id),
        getUserInteractionState(props.context, document.id, currentUserId)
      ]);
      setLikeCount(metrics.likes);
      setIsLiked(interactionState.isLiked);
    } catch (error) {
      setLikeCount(0);
      setIsLiked(false);
    }
  };

  const fetchComments = async () => {
    if (!props.context || !document) return;
    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const listName = LIST_NAMES.documentComments;

      // Fetch all comments for this document (ordered by newest first)
      const response = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=DocumentId eq ${document.id}&$orderby=Created desc&$select=Id,UserName,Comment,Created,UserId`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const items = data.value || [];
        const commentsData = items.map((item: any) => ({
          id: item.Id,
          userName: item.UserName || 'User',
          comment: item.Comment || '',
          created: new Date(item.Created),
          userId: item.UserId || 0
        }));
        setComments(commentsData);
      } else {
        setComments([]);
      }
    } catch (error) {
      console.error('Error fetching comments:', error);
      setComments([]);
    }
  };

  const getActionListName = (action: DocumentToggleAction): string => {
    return 'DocumentReviewFlags';
  };

  const getActionListFields = (
    action: DocumentToggleAction
  ): Array<{ name: string; type: string; required: boolean }> => {
    if (action === 'follow') {
      return [
        { name: 'DocumentId', type: 'Number', required: true },
        { name: 'UserId', type: 'Number', required: true }
      ];
    }

    if (action === 'bookmark') {
      return [
        { name: 'DocumentId', type: 'Number', required: true },
        { name: 'UserId', type: 'Number', required: true }
      ];
    }

    return [
      { name: 'DocumentId', type: 'Number', required: true },
      { name: 'UserId', type: 'Number', required: false },
      { name: 'UserName', type: 'Text', required: false },
      { name: 'Reason', type: 'Text', required: false },
      { name: COLUMN_NAMES.comments, type: 'Note', required: false }
    ];
  };

  const ensureListFieldsExist = React.useCallback(async (
    listName: string,
    fields: Array<{ name: string; type: string; required?: boolean }>
  ): Promise<boolean> => {
    if (!props.context) {
      return false;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const fieldsResponse = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName`,
      SPHttpClient.configurations.v1
    );

    if (!fieldsResponse.ok) {
      const errorText = await fieldsResponse.text();
      console.error(`Unable to read fields for "${listName}":`, fieldsResponse.status, errorText);
      return false;
    }

    const fieldsJson = await fieldsResponse.json();
    const existingFields = (fieldsJson?.d?.results || fieldsJson?.value || []) as Array<{ Title?: string; InternalName?: string }>;
    const normalizeFieldName = (value?: string): string => (value || '').trim().toLowerCase();
    const hasField = (fieldName: string): boolean =>
      existingFields.some((field) =>
        normalizeFieldName(field.Title) === normalizeFieldName(fieldName) ||
        normalizeFieldName(field.InternalName) === normalizeFieldName(fieldName)
      );

    for (const field of fields) {
      if (hasField(field.name)) {
        continue;
      }

      let fieldBody: any = {
        __metadata: { type: 'SP.Field' },
        Title: field.name,
        Required: field.required || false
      };

      if (field.type === 'Number') {
        fieldBody.FieldTypeKind = 9;
        fieldBody = {
          ...fieldBody,
          __metadata: { type: 'SP.FieldNumber' }
        };
      } else if (field.type === 'Text') {
        fieldBody.FieldTypeKind = 2;
        fieldBody.MaxLength = 255;
        fieldBody = {
          ...fieldBody,
          __metadata: { type: 'SP.FieldText' }
        };
      } else if (field.type === 'Note') {
        fieldBody.FieldTypeKind = 3;
        fieldBody = {
          ...fieldBody,
          __metadata: { type: 'SP.FieldMultiLineText' }
        };
      }

      const createFieldResponse = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: JSON.stringify(fieldBody)
        }
      );

      if (!createFieldResponse.ok) {
        const errorText = await createFieldResponse.text();
        console.error(`Unable to add field "${field.name}" to "${listName}":`, createFieldResponse.status, errorText);
        return false;
      }
    }

    return true;
  }, [props.context]);

  const getListFieldNames = React.useCallback(async (listName: string): Promise<Set<string>> => {
    if (!props.context) {
      return new Set<string>();
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const fieldsResponse = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName`,
      SPHttpClient.configurations.v1
    );

    if (!fieldsResponse.ok) {
      return new Set<string>();
    }

    const fieldsJson = await fieldsResponse.json();
    const existingFields = (fieldsJson?.d?.results || fieldsJson?.value || []) as Array<{ Title?: string; InternalName?: string }>;
    const normalizeFieldName = (value?: string): string => (value || '').trim().toLowerCase();

    return new Set<string>(
      existingFields.reduce<string[]>((allFieldNames, field) => {
        const nextFieldNames = [normalizeFieldName(field.Title), normalizeFieldName(field.InternalName)].filter(Boolean);
        return allFieldNames.concat(nextFieldNames);
      }, [])
    );
  }, [props.context]);

  const getCountListName = (action: DocumentCountAction): string => {
    return action === 'flag' ? getActionListName(action) : 'DocumentReviewFlags';
  };

  const setActionState = (action: DocumentToggleAction, value: boolean): void => {
    if (action === 'follow') {
      setIsFollowing(value);
      return;
    }

    if (action === 'bookmark') {
      setIsBookmarked(value);
      return;
    }

    setIsFlaggedForReview(value);
  };

  const setActionCount = (action: DocumentCountAction, value: number): void => {
    if (action === 'follow') {
      setFollowCount(value);
      return;
    }

    if (action === 'bookmark') {
      setBookmarkCount(value);
      return;
    }

    if (action === 'flag') {
      setFlagCount(value);
      return;
    }

    setShareCount(value);
  };

  const fetchDocumentActionCount = async (action: DocumentCountAction) => {
    if (!props.context || !document) return;

    try {
      if (action === 'share') {
        const metrics = await getDocumentMetrics(props.context, document.id);
        setShareCount(metrics.share);
        return;
      }

      const webUrl = props.context.pageContext.web.absoluteUrl;
      const listName = getCountListName(action);
      const response = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=DocumentId eq ${document.id}&$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        setActionCount(action, (data.value || []).length);
      } else {
        setActionCount(action, 0);
      }
    } catch (error) {
      console.error(`Error fetching ${action} count:`, error);
      setActionCount(action, 0);
    }
  };

  const fetchDocumentActionMetrics = async (action: DocumentToggleAction) => {
    if (!props.context || !document || !resolvedCurrentUserId) return;

    try {
      if (action === 'bookmark' || action === 'follow') {
        const [metrics, interactionState] = await Promise.all([
          getDocumentMetrics(props.context, document.id),
          getUserInteractionState(props.context, document.id, resolvedCurrentUserId)
        ]);

        if (action === 'bookmark') {
          setIsBookmarked(interactionState.isBookmarked);
          setBookmarkCount(metrics.bookmark);
        } else {
          setIsFollowing(interactionState.isFollowed);
          setFollowCount(metrics.follow);
        }
        return;
      }

      const webUrl = props.context.pageContext.web.absoluteUrl;
      const listName = getActionListName(action);
      const listExists = await ensureListExists(listName, getActionListFields(action));

      if (!listExists) {
        setActionState(action, false);
        setActionCount(action, 0);
        return;
      }

      const response = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=DocumentId eq ${document.id}&$select=Id,UserId`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const actionItems = data.value || [];
        setActionState(action, actionItems.some((item: any) => item.UserId === resolvedCurrentUserId));
        setActionCount(action, actionItems.length);
      } else {
        setActionState(action, false);
        setActionCount(action, 0);
      }
    } catch (error) {
      console.error(`Error fetching ${action} metrics:`, error);
      setActionState(action, false);
      setActionCount(action, 0);
    }
  };

  const fetchMetricCount = async (
    listName: string,
    setCount: React.Dispatch<React.SetStateAction<number>>
  ) => {
    if (!props.context || !document) return;

    try {
      const metrics = await getDocumentMetrics(props.context, document.id);
      if (listName === 'views') setCount(metrics.views);
      else if (listName === 'downloads') setCount(metrics.downloads);
      else if (listName === 'likes') setCount(metrics.likes);
      else if (listName === 'comments') setCount(metrics.comments);
      else setCount(0);
    } catch (error) {
      console.error(`Error fetching ${listName} count:`, error);
      setCount(0);
    }
  };

  const getListEntityType = async (listName: string): Promise<string | null> => {
    if (!props.context) return null;

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const response = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return null;
    }

    const listInfo = await response.json();
    return listInfo.ListItemEntityTypeFullName || null;
  };

  const recordUniqueView = React.useCallback(async (): Promise<void> => {
    if (!props.context || !document || !currentUserId) return;
    if (!document.id || !document.name) return;
    if ((document.status || '').toLowerCase() !== 'active') return;

    const recordKey = `${document.id}-${currentUserId}`;
    if (recordedViewKeyRef.current === recordKey || pendingViewKeyRef.current === recordKey) {
      return;
    }

    pendingViewKeyRef.current = recordKey;

    try {
      const newCount = await recordUserEvent(props.context, document.id, currentUserId, 'View', document.name);
      setViewCount(newCount);
      recordedViewKeyRef.current = recordKey;
    } catch (error) {
      console.error('Error recording unique view:', error);
    } finally {
      if (pendingViewKeyRef.current === recordKey) {
        pendingViewKeyRef.current = null;
      }
    }
  }, [currentUserId, document?.id, document?.name, document?.status, props.context]);

  React.useEffect(() => {
    if (!document?.id || !document.name || !resolvedCurrentUserId || !props.context) {
      return;
    }

    void recordUniqueView();
  }, [document?.id, document?.name, resolvedCurrentUserId, props.context, recordUniqueView]);

  const recordDownload = async (): Promise<void> => {
    if (!props.context || !document || !currentUserId) return;
    if ((document.status || '').toLowerCase() !== 'active') return;

    try {
      const newCount = await recordUserEvent(props.context, document.id, currentUserId, 'Download', document.name);
      setDownloadCount(newCount);
      props.onDownloadRecorded?.(document.id);
    } catch (error) {
      console.error('Error recording download:', error);
    }
  };

  const toggleDocumentAction = async (action: DocumentToggleAction) => {
    if (!props.context || !document || !resolvedCurrentUserId) {
      return;
    }
    if (action !== 'bookmark' && (document.status || '').toLowerCase() !== 'active') return;

    const currentValue =
      action === 'follow' ? isFollowing :
        action === 'bookmark' ? isBookmarked :
          isFlaggedForReview;

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;

      if (action === 'bookmark') {
        const interactionKey = `${document.id}_Bookmark`;
        if (socialInteractionInProgressRef.current.has(interactionKey)) {
          console.log('[DEBOUNCE] Bookmark interaction in progress, skipping');
          return;
        }

        const nextActive = !isBookmarked;
        const previousActive = isBookmarked;
        const previousCount = bookmarkCount;
        setSocialInteractionPending(interactionKey, true);
        setIsBookmarked(nextActive);
        setBookmarkCount((prev) => Math.max(0, prev + (nextActive ? 1 : -1)));
        setActionMessage(nextActive ? 'Bookmark successful' : 'Bookmark removed');
        try {
          const result = await toggleUserInteraction(props.context, document.id, resolvedCurrentUserId, 'Bookmark', document.name);
          setIsBookmarked(result.isActive);
          setBookmarkCount(result.newCount);
          emitDocumentDataChanged({
            documentIds: [document.id],
            reason: 'bookmark'
          });
        } catch (error) {
          setIsBookmarked(previousActive);
          setBookmarkCount(previousCount);
          console.error('Error toggling bookmark:', error);
          setActionMessage('Unable to update bookmark right now. Please try again.');
        } finally {
          setSocialInteractionPending(interactionKey, false);
        }
        return;
      }

      if (action === 'follow') {
        const interactionKey = `${document.id}_Follow`;
        if (socialInteractionInProgressRef.current.has(interactionKey)) {
          console.log('[DEBOUNCE] Follow interaction in progress, skipping');
          return;
        }

        const nextActive = !isFollowing;
        const previousActive = isFollowing;
        const previousCount = followCount;
        setSocialInteractionPending(interactionKey, true);
        setIsFollowing(nextActive);
        setFollowCount((prev) => Math.max(0, prev + (nextActive ? 1 : -1)));
        setActionMessage(nextActive ? 'Now following this article' : 'Follow removed');
        try {
          const result = await toggleUserInteraction(props.context, document.id, resolvedCurrentUserId, 'Follow', document.name);
          setIsFollowing(result.isActive);
          setFollowCount(result.newCount);
        } catch (error) {
          setIsFollowing(previousActive);
          setFollowCount(previousCount);
          console.error('Error toggling follow:', error);
          setActionMessage('Unable to update follow right now. Please try again.');
        } finally {
          setSocialInteractionPending(interactionKey, false);
        }
        return;
      }

      const listName = getActionListName(action);
      const listExists = await ensureListExists(listName, getActionListFields(action));

      if (!listExists) {
        setActionMessage('Unable to update this action right now. Please try again.');
        return;
      }

      if (currentValue) {
        const findResponse = await props.context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=DocumentId eq ${document.id} and UserId eq ${currentUserId}&$select=Id`,
          SPHttpClient.configurations.v1
        );

        if (findResponse.ok) {
          const findData = await findResponse.json();
          if (findData.value && findData.value.length > 0) {
            const actionItemId = findData.value[0].Id;
            await props.context.spHttpClient.post(
              `${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${actionItemId})`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  "IF-MATCH": "*",
                  "X-HTTP-Method": "DELETE"
                }
              }
            );
          }
        }

        setActionState(action, false);
        setActionCount(action, Math.max(0, flagCount - 1));
        setActionMessage('Review flag removed');
        return;
      }

      const entityType = await getListEntityType(listName);
      if (!entityType) {
        throw new Error(`Could not get entity type for ${listName}`);
      }

      const body = {
        __metadata: { type: entityType },
        Title: `${document.name || 'Document'} ${action}`,
        DocumentId: document.id,
        UserId: currentUserId
      };

      const createResponse = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/items`,
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

      if (!createResponse.ok) {
        const errorText = await createResponse.text();
        throw new Error(`Failed to save ${action}: ${createResponse.status} ${errorText}`);
      }

      setActionState(action, true);
      setActionCount(action, flagCount + 1);
      setActionMessage('Article flagged for review');
    } catch (error) {
      console.error(`Error toggling ${action}:`, error);
      setActionMessage(
        error instanceof Error
          ? error.message
          : `Unable to update ${action} right now. Please try again.`
      );
    }
  };

  const toggleLike = async () => {
    if (!props.context || !document || !currentUserId) return;
    if ((document.status || '').toLowerCase() !== 'active') return;

    const interactionKey = `${document.id}_Like`;
    if (socialInteractionInProgressRef.current.has(interactionKey)) {
      console.log('[DEBOUNCE] Like interaction in progress, skipping');
      return;
    }

    try {
      const nextActive = !isLiked;
      const previousActive = isLiked;
      const previousCount = likeCount;
      setSocialInteractionPending(interactionKey, true);
      setIsLiked(nextActive);
      setLikeCount((prev) => Math.max(0, prev + (nextActive ? 1 : -1)));
      try {
        const result = await toggleUserInteraction(props.context, document.id, currentUserId, 'Like', document.name);
        setIsLiked(result.isActive);
        setLikeCount(result.newCount);
      } catch (error) {
        setIsLiked(previousActive);
        setLikeCount(previousCount);
        console.error('Error toggling like:', error);
      } finally {
        setSocialInteractionPending(interactionKey, false);
      }
    } catch (error) {
      console.error('Error toggling like:', error);
      setSocialInteractionPending(interactionKey, false);
    }
  };

  const addComment = async () => {
    if (!props.context || !document || !currentUserId || !newComment.trim()) {
      console.log('Cannot add comment - missing requirements:', {
        hasContext: !!props.context,
        hasDocument: !!document,
        hasUserId: !!currentUserId,
        hasComment: !!newComment.trim()
      });
      return;
    }
    if ((document.status || '').toLowerCase() !== 'active') return;

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const listName = LIST_NAMES.documentComments;

      console.log('Adding comment:', {
        documentId: document.id,
        userName: currentUserName,
        userId: currentUserId,
        comment: newComment.trim()
      });

      // Ensure list exists before trying to use it
      const listExists = await ensureListExists(LIST_NAMES.documentComments, [
        { name: 'DocumentId', type: 'Number', required: true },
        { name: 'UserName', type: 'Text', required: true },
        { name: 'Comment', type: 'Note', required: true },
        { name: 'UserId', type: 'Number', required: true }
      ]);

      if (!listExists) {
        alert('Failed to create or access the DocumentComments list. Please try again.');
        return;
      }

      const listInfoResp = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      );

      if (!listInfoResp.ok) {
        const errorText = await listInfoResp.text();
        console.error('Failed to get list info:', listInfoResp.status, errorText);
        alert('Failed to add comment. Please try again.');
        return;
      }

      const listInfo = await listInfoResp.json();
      const entityType = listInfo.ListItemEntityTypeFullName;

      const body = {
        __metadata: { type: entityType },
        DocumentId: document.id,
        UserName: currentUserName,
        Comment: newComment.trim(),
        UserId: currentUserId
      };

      console.log('Posting comment with body:', body);

      const postResp = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/items`,
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

      if (!postResp.ok) {
        const errorText = await postResp.text();
        console.error('Failed to add comment:', postResp.status, errorText);
        alert('Failed to add comment. Please try again.');
        return;
      }

      console.log('Comment added successfully');
      setNewComment('');
      const newCount = await incrementMetricWithRetry(props.context, document.id, 'CommentCount', 1);
      setCommentCount(newCount);
      await fetchComments();
    } catch (error) {
      console.error('Error adding comment:', error);
      alert('An error occurred while adding the comment. Please try again.');
    }
  };

  const deleteComment = async (commentId: number) => {
    if (!props.context || !commentId) return;

    if (!confirm('Are you sure you want to delete this comment?')) {
      return;
    }

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const listName = LIST_NAMES.documentComments;

      const deleteResp = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${commentId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "IF-MATCH": "*",
            "X-HTTP-Method": "DELETE"
          }
        }
      );

      if (!deleteResp.ok) {
        const errorText = await deleteResp.text();
        console.error('Failed to delete comment:', deleteResp.status, errorText);
        alert('Failed to delete comment. Please try again.');
        return;
      }

      console.log('Comment deleted successfully');
      const newCount = await incrementMetricWithRetry(props.context, document.id, 'CommentCount', -1);
      setCommentCount(newCount);
      await fetchComments();
    } catch (error) {
      console.error('Error deleting comment:', error);
      alert('An error occurred while deleting the comment. Please try again.');
    }
  };

  const formatTimeAgo = (date: Date): string => {
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);
    const diffWeeks = Math.floor(diffDays / 7);
    const diffYears = Math.floor(diffDays / 365);

    if (diffMins < 1) return 'just now';
    if (diffMins < 60) return `${diffMins} ${diffMins === 1 ? 'minute' : 'minutes'} ago`;
    if (diffHours < 24) return `${diffHours} ${diffHours === 1 ? 'hour' : 'hours'} ago`;
    if (diffDays < 7) return `${diffDays} ${diffDays === 1 ? 'day' : 'days'} ago`;
    if (diffWeeks < 52) return `${diffWeeks} ${diffWeeks === 1 ? 'week' : 'weeks'} ago`;
    return `${diffYears} ${diffYears === 1 ? 'year' : 'years'} ago`;
  };

  const handleBack = () => {
    if (props.backTo === 'library' && props.onBackToLibrary) {
      props.onBackToLibrary();
    } else if (props.onClose) {
      props.onClose();
    }
  };

  const backButtonText =
    props.backButtonLabel ||
    (props.backTo === 'library'
      ? 'Back to Library'
      : props.backTo === 'home' || props.backTo === 'section'
        ? 'Back to Home'
        : 'Back');
  const documentStatus = (document?.status || '').toLowerCase();
  const isDocumentActive = documentStatus === 'active';
  const disableSocialActions =
    props.disableSocialActions === true
      ? true
      : !isDocumentActive;
  const hideTopDocumentActions = props.hideTopDocumentActions === true;

  const showViewerNotice = React.useCallback((message: string) => {
    setViewerNotice(message);
  }, []);

  const handleOpenUpdateUpload = React.useCallback(() => {
    if (!canEditOrUpdateDocument) {
      return;
    }
    setIsUpdateUploadOpen(true);
  }, [canEditOrUpdateDocument]);

  const handleCloseUpdateUpload = React.useCallback(() => {
    setIsUpdateUploadOpen(false);
  }, []);

  const handleUpdateUploadComplete = React.useCallback(() => {
    setActionMessage('Document updated successfully.');
    void fetchDocumentDetails();
  }, [fetchDocumentDetails]);

  // SharePoint often blocks or destabilizes the remote pdf.js worker used by react-pdf.
  // Prefer the browser's native PDF viewer for blob URLs so PDFs still open reliably.
  const isPdfCustomViewer = false;
  const canPrintDocument = isKmAdmin || isApprover || isContributor;
  const isPrintDisabled = isVideoDocument || isAudioDocument || !canPrintDocument;
  const printButtonTitle = isPrintDisabled
    ? !canPrintDocument
      ? 'Print is restricted for your role'
      : `Print is unavailable for ${document?.fileType || 'this file type'}`
    : 'Print';
  const viewerAdapterKind: ViewerAdapterKind = isPdfDocument
    ? 'pdf'
    : isOfficeDocument
      ? 'office'
      : isImageDocument
        ? 'image'
        : isTextDocument
          ? 'text'
          : isVideoDocument
            ? 'video'
            : isAudioDocument
              ? 'audio'
              : 'unsupported';
  const viewerAdapter: IViewerAdapter = {
    kind: viewerAdapterKind,
    capabilities: {
      canToggleThumbnails:
        isPdfDocument ||
        isPresentationDocument ||
        isWordProcessingDocument ||
        isImageDocument ||
        (isTextDocument && !isSimpleOfficeStyleToolbarDocument),
      canZoom: isPdfDocument || isImageDocument || isTextDocument || isAudioDocument || isOfficeDocument || isExcelLikeDocument,
      canNavigatePages:
        isPdfDocument ||
        ((isPresentationDocument || isWordProcessingDocument) && embeddedPageCount > 1),
      canWindowMode: isPreviewSupportedDocument,
      canRotate: isPdfDocument || isImageDocument || (isTextDocument && !isSimpleOfficeStyleToolbarDocument),
      canChangeLayout: isPdfDocument,
      pageUnitLabel: isPresentationDocument ? 'slide' : isExcelLikeDocument ? 'sheet' : 'page',
      pageUnitLabelPlural: isPresentationDocument ? 'slides' : isExcelLikeDocument ? 'sheets' : 'pages'
    }
  };
  const viewerCapabilities = viewerAdapter.capabilities;
  const shouldShowViewerThumbnailToggle = !isVideoDocument && !isAudioDocument;
  const supportsThumbnailSidebar = viewerCapabilities.canToggleThumbnails;
  const shouldRenderCustomViewerToolbar =
    !!previewUrl &&
    !isPreviewLoading &&
    !shouldUsePdfPreviewViewer &&
    !isImageDocument &&
    !isMediaDocument &&
    !isUnsupportedPreviewDocument;
  const resolvedPageTotal = isPdfDocument || isPdfCustomViewer ? pdfPageCount : embeddedPageCount;
  const hasResolvedPageTotal = resolvedPageTotal > 0;
  const activePageTotal = hasResolvedPageTotal ? resolvedPageTotal : 1;
  const activePageNumber = isPdfDocument || isPdfCustomViewer
    ? Math.min(currentPdfPage, activePageTotal)
    : Math.min(currentEmbeddedPage, activePageTotal);
  const supportsPrecisePageNavigation = viewerCapabilities.canNavigatePages && hasResolvedPageTotal;
  const pageIndicatorText = supportsPrecisePageNavigation
    ? `${activePageNumber} / ${activePageTotal} ${viewerCapabilities.pageUnitLabelPlural}`
    : hasResolvedPageTotal
      ? `${activePageTotal} ${activePageTotal === 1 ? viewerCapabilities.pageUnitLabel : viewerCapabilities.pageUnitLabelPlural}`
      : isAudioDocument
        ? 'Audio preview'
        : isVideoDocument
          ? 'Video preview'
          : `${document?.fileType || 'File'} preview`;
  const usesOfficeIframeZoomSync =
    isOfficeDocument &&
    (isPresentationDocument || isSpreadsheetDocument) &&
    previewUrl.indexOf('/_layouts/15/WopiFrame.aspx') !== -1;
  const activeZoomPercent = Math.round((isPdfCustomViewer ? pdfZoom : usesOfficeIframeZoomSync ? officeToolbarZoom : embeddedZoom) * 100);
  React.useEffect(() => {
    if (!isEditingZoom) {
      setZoomInputValue(String(activeZoomPercent));
    }
  }, [activeZoomPercent, isEditingZoom]);
  const toolbarSummaryText = isPresentationDocument
    ? 'PowerPoint viewer with slide thumbnails and navigation.'
    : isExcelLikeDocument
      ? 'Excel viewer with zoom, fullscreen, and print controls.'
      : isWordProcessingDocument
        ? 'Word viewer with page-aware status when the document exposes page count.'
        : isPdfDocument
          ? 'PDF viewer uses the browser-safe embedded preview inside SharePoint.'
          : isVideoDocument
            ? 'Media viewer with the shared toolbar shell and window view support.'
            : isAudioDocument
              ? 'Audio viewer with the shared toolbar shell and playback-only controls.'
              : isTextDocument
                ? 'Text viewer with zoom, fullscreen, and print controls.'
                : isUnsupportedPreviewDocument
                  ? 'Preview is not available for this file format.'
                  : 'Document preview is rendered inside the shared custom toolbar shell.';

  const getViewerFeatureUnavailableMessage = React.useCallback((feature: ViewerFeatureKey): string => {
    const fileLabel = document?.fileType || 'This file';

    switch (feature) {
      case 'thumbnails':
        return `${fileLabel} preview does not provide thumbnail navigation here.`;
      case 'zoom':
        return `${fileLabel} preview does not support custom zoom controls here.`;
      case 'pageNavigation':
        return `${fileLabel} preview does not expose reliable page navigation in this embedded viewer.`;
      case 'windowMode':
        return `${fileLabel} preview does not support window mode changes right now.`;
      case 'rotate':
        return `${fileLabel} preview does not support custom rotation controls here.`;
      default:
        return `${fileLabel} preview does not support that action in this custom toolbar.`;
    }
  }, [document?.fileType]);

  const handleSidebarToggle = (panel: Exclude<PdfSidebarPanel, null>) => {
    if (panel === 'thumbnails' && supportsThumbnailSidebar) {
      setPdfSidebarPanel((current) => current === panel ? null : panel);
      return;
    }

    if (!isPdfCustomViewer) {
      showViewerNotice(getViewerFeatureUnavailableMessage('thumbnails'));
      return;
    }

    setPdfSidebarPanel((current) => current === panel ? null : panel);
  };

  const handlePdfDocumentLoad = ({ numPages }: { numPages: number }) => {
    setPdfPageCount(numPages);
    setCurrentPdfPage(1);
    setPreviewError('');
  };

  const changePdfPage = (pageNumber: number) => {
    if (pdfPageCount < 1) return;

    const nextPage = Math.min(Math.max(pageNumber, 1), pdfPageCount);
    setCurrentPdfPage(nextPage);
  };

  const changeEmbeddedPage = (pageNumber: number, _preferredPreviewUrl?: string) => {
    if (!document || !props.context || !isOfficeDocument) {
      return;
    }

    const nextPage = Math.min(Math.max(pageNumber, 1), Math.max(embeddedPageCount, 1));
    if (isPresentationDocument) {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const siteId = String(props.context.pageContext.site.id || '');
      const webId = String(props.context.pageContext.web.id || '');
      const slidePreviewCandidates = buildSlideImagePreviewCandidates(
        webUrl,
        normalizeServerRelativeUrl(document.serverRelativeUrl),
        nextPage,
        document.fileUniqueId,
        siteId,
        webId
      );
      const selectedSlidePreviewUrl = buildOfficePreviewUrl(
        webUrl,
        document.serverRelativeUrl,
        normalizedFileType,
        nextPage
      );
      const orderedSlidePreviewCandidates = selectedSlidePreviewUrl
        ? [selectedSlidePreviewUrl, ...slidePreviewCandidates.filter((candidateUrl) => candidateUrl !== selectedSlidePreviewUrl)]
        : slidePreviewCandidates;

      if (selectedSlidePreviewUrl === previewUrl) {
        return;
      }

      setCurrentEmbeddedPage(nextPage);
      setPreviewImageUrls(orderedSlidePreviewCandidates);
      setPreviewUrl(selectedSlidePreviewUrl);
      return;
    }

    const nextPreviewUrl = buildOfficePreviewUrl(
      props.context.pageContext.web.absoluteUrl,
      document.serverRelativeUrl,
      normalizedFileType,
      nextPage
    );

    if (nextPreviewUrl === previewUrl) {
      return;
    }

    setCurrentEmbeddedPage(nextPage);
    setPreviewUrl(nextPreviewUrl);
  };

  const getThumbnailCandidateKey = (pageNumber: number): string =>
    `${document?.id || document?.serverRelativeUrl || 'document'}:${pageNumber}`;

  const getSidebarThumbnailUrls = (pageNumber: number): string[] => {
    if (isImageDocument || normalizedFileType === 'svg') {
      return previewUrl ? [previewUrl] : [];
    }

    if ((isPdfDocument || isOfficeDocument) && document && props.context) {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const siteId = String(props.context.pageContext.site.id || '');
      const webId = String(props.context.pageContext.web.id || '');
      const normalizedServerRelativeUrl = normalizeServerRelativeUrl(document.serverRelativeUrl);

      if (isPresentationDocument) {
        return [
          buildWopiFrameImagePreviewUrl(webUrl, normalizedServerRelativeUrl, normalizedFileType, pageNumber),
          ...buildSlideImagePreviewCandidates(
            webUrl,
            normalizedServerRelativeUrl,
            pageNumber,
            document.fileUniqueId,
            siteId,
            webId
          )
        ];
      }

      if (isWordProcessingDocument) {
        const safePageNumber = Math.max(pageNumber, 1);
        const rawPageIndex = Math.max(safePageNumber - 1, 0);
        const wordPreviewUrls = [
          buildIndexedPreviewImageUrl(webUrl, normalizedServerRelativeUrl, safePageNumber, document.fileUniqueId, siteId, webId),
          buildRawIndexedPreviewImageUrl(webUrl, normalizedServerRelativeUrl, rawPageIndex, document.fileUniqueId, siteId, webId),
          buildRawIndexedPreviewImageUrl(webUrl, normalizedServerRelativeUrl, safePageNumber, document.fileUniqueId, siteId, webId),
          buildIndexedPreviewImageUrl(webUrl, normalizedServerRelativeUrl, safePageNumber),
          buildRawIndexedPreviewImageUrl(webUrl, normalizedServerRelativeUrl, rawPageIndex),
          buildRawIndexedPreviewImageUrl(webUrl, normalizedServerRelativeUrl, safePageNumber),
          safePageNumber === 1 ? buildPreviewImageUrl(webUrl, normalizedServerRelativeUrl) : '',
          buildWopiFrameImagePreviewUrl(webUrl, normalizedServerRelativeUrl, normalizedFileType, safePageNumber),
          buildOfficePreviewUrl(webUrl, normalizedServerRelativeUrl, normalizedFileType, safePageNumber)
        ].filter((url) => Boolean(url));

        return Array.from(new Set(wordPreviewUrls));
      }

      const pageIndexes = isOfficeDocument
        ? [pageNumber, pageNumber + 1, pageNumber - 1]
        : [pageNumber, pageNumber + 1];
      const urls = [
        ...pageIndexes.map((pageIndex) =>
          buildRawIndexedPreviewImageUrl(
            webUrl,
            normalizedServerRelativeUrl,
            pageIndex,
            document.fileUniqueId,
            siteId,
            webId
          )
        ),
        ...pageIndexes.map((pageIndex) =>
          buildRawIndexedPreviewImageUrl(
            webUrl,
            normalizedServerRelativeUrl,
            pageIndex
          )
        )
      ].filter((url) => Boolean(url));

      return Array.from(new Set(urls));
    }

    return [];
  };

  const getSidebarThumbnailUrl = (pageNumber: number): string | null => {
    const candidates = getSidebarThumbnailUrls(pageNumber);
    const candidateIndex = thumbnailCandidateIndexes[getThumbnailCandidateKey(pageNumber)] || 0;
    return candidates[candidateIndex] || null;
  };

  const hasMoreThumbnailCandidates = (pageNumber: number): boolean => {
    const candidates = getSidebarThumbnailUrls(pageNumber);
    const candidateIndex = thumbnailCandidateIndexes[getThumbnailCandidateKey(pageNumber)] || 0;
    return candidateIndex < candidates.length - 1;
  };

  const handleThumbnailLoadError = (pageNumber: number): void => {
    const candidates = getSidebarThumbnailUrls(pageNumber);
    const candidateKey = getThumbnailCandidateKey(pageNumber);
    const candidateIndex = thumbnailCandidateIndexes[candidateKey] || 0;

    if (candidateIndex >= candidates.length - 1) {
      return;
    }

    setThumbnailCandidateIndexes((current) => ({
      ...current,
      [candidateKey]: candidateIndex + 1
    }));
  };

  const shouldPrioritizeThumbnail = (pageNumber: number): boolean =>
    pageNumber === 1 || Math.abs(pageNumber - activePageNumber) <= 2;

  const getThumbnailLoadingMode = (pageNumber: number): 'eager' | 'lazy' =>
    shouldPrioritizeThumbnail(pageNumber) ? 'eager' : 'lazy';

  const centerMediaPreviewScroll = React.useCallback((): void => {
    const container = previewContainerRef.current;
    if (!container) {
      return;
    }

    window.requestAnimationFrame(() => {
      window.requestAnimationFrame(() => {
        container.scrollLeft = Math.max(0, (container.scrollWidth - container.clientWidth) / 2);
        container.scrollTop = Math.max(0, (container.scrollHeight - container.clientHeight) / 2);
      });
    });
  }, []);

  const applyExcelIframeZoom = React.useCallback((zoomValue: number): boolean => {
    const iframe = iframeRef.current;
    if (!iframe) {             
      return false;
    }

    try {
      const iframeDocument = iframe.contentDocument;
      if (!iframeDocument) {
        return false;
      }

      const clampedZoom = Math.min(Math.max(Number(zoomValue.toFixed(2)), 0.25), 2.5);
      const zoomCssValue = String(clampedZoom);

      iframeDocument.documentElement.style.setProperty('zoom', zoomCssValue);
      iframeDocument.body?.style.setProperty('zoom', zoomCssValue);
      iframeDocument.documentElement.style.setProperty('transform-origin', 'top left');
      iframeDocument.body?.style.setProperty('transform-origin', 'top left');
      return true;
    } catch (error) {
      return false;
    }
  }, []);

  const updateZoom = (direction: 'in' | 'out' | 'reset') => {
    if (!viewerCapabilities.canZoom) {
      showViewerNotice(getViewerFeatureUnavailableMessage('zoom'));
      return;
    }

    if (isPdfCustomViewer) {
      if (direction === 'reset') {
        setPdfZoom(1);
        setIsFitToContainerMode(true);
        return;
      }

      setPdfZoom((previousZoom) => {
        const nextZoom = direction === 'in' ? previousZoom + 0.25 : previousZoom - 0.25;
        return Math.min(Math.max(Number(nextZoom.toFixed(2)), 0.25), 2.5);
      });
      setIsFitToContainerMode(false);
      return;
    }

    if (direction === 'reset') {
      setEmbeddedZoom(1);
      setOfficeToolbarZoom(1);
      if (usesOfficeIframeZoomSync && isSpreadsheetDocument) {
        applyExcelIframeZoom(1);
      }
      setIsFitToContainerMode(true);
      if (isMediaDocument) {
        centerMediaPreviewScroll();
      }
      return;
    }

    if (usesOfficeIframeZoomSync && isSpreadsheetDocument) {
      setOfficeToolbarZoom((previousZoom) => {
        const nextZoom = direction === 'in' ? previousZoom + 0.25 : previousZoom - 0.25;
        const clampedZoom = Math.min(Math.max(Number(nextZoom.toFixed(2)), 0.25), 2.5);
        applyExcelIframeZoom(clampedZoom);
        return clampedZoom;
      });
      setIsFitToContainerMode(false);
      return;
    }

    setEmbeddedZoom((previousZoom) => {
      const nextZoom = direction === 'in' ? previousZoom + 0.25 : previousZoom - 0.25;
      const clampedZoom = Math.min(Math.max(Number(nextZoom.toFixed(2)), 0.25), 2.5);
      if (isMediaDocument && direction === 'out') {
        centerMediaPreviewScroll();
      }
      setOfficeToolbarZoom(clampedZoom);
      return clampedZoom;
    });
    setIsFitToContainerMode(false);
  };

  const rotatePdf = (direction: 'left' | 'right') => {
    if (!viewerCapabilities.canRotate) {
      showViewerNotice(getViewerFeatureUnavailableMessage('rotate'));
      return;
    }

    const getNextRotation = (previousRotation: number): number => {
      const nextRotation = previousRotation + (direction === 'right' ? 90 : -90);
      return ((nextRotation % 360) + 360) % 360;
    };

    if (isPdfCustomViewer) {
      setPdfRotation(getNextRotation);
      return;
    }

    if (isPresentationDocument) {
      changeEmbeddedPage(currentEmbeddedPage);
    }

    setEmbeddedRotation(getNextRotation);
  };

  const applyZoomPreset = (preset: number) => {
    if (!viewerCapabilities.canZoom) {
      showViewerNotice(getViewerFeatureUnavailableMessage('zoom'));
      return;
    }

    const nextZoom = preset / 100;

    if (isPdfCustomViewer) {
      setPdfZoom(nextZoom);
    } else if (usesOfficeIframeZoomSync && isSpreadsheetDocument) {
      setOfficeToolbarZoom(nextZoom);
      applyExcelIframeZoom(nextZoom);
    } else {
      setEmbeddedZoom(nextZoom);
      setOfficeToolbarZoom(nextZoom);
    }

    setIsFitToContainerMode(false);
    setIsZoomMenuOpen(false);
  };

  const commitZoomInput = (): void => {
    if (!viewerCapabilities.canZoom) {
      showViewerNotice(getViewerFeatureUnavailableMessage('zoom'));
      setZoomInputValue(String(activeZoomPercent));
      setIsEditingZoom(false);
      return;
    }

    const numericValue = Number(zoomInputValue.replace('%', '').trim());
    if (!Number.isFinite(numericValue)) {
      setZoomInputValue(String(activeZoomPercent));
      setIsEditingZoom(false);
      return;
    }

    const nextZoom = Math.min(Math.max(Number((numericValue / 100).toFixed(2)), 0.25), 2.5);
    if (isPdfCustomViewer) {
      setPdfZoom(nextZoom);
    } else if (usesOfficeIframeZoomSync && isSpreadsheetDocument) {
      setOfficeToolbarZoom(nextZoom);
      applyExcelIframeZoom(nextZoom);
    } else {
      setEmbeddedZoom(nextZoom);
      setOfficeToolbarZoom(nextZoom);
    }

    if (isMediaDocument) {
      centerMediaPreviewScroll();
    }
    setIsFitToContainerMode(false);
    setIsZoomMenuOpen(false);
    setZoomInputValue(String(Math.round(nextZoom * 100)));
    setIsEditingZoom(false);
  };

  const changeViewerPage = (direction: 'prev' | 'next') => {
    if (!viewerCapabilities.canNavigatePages) {
      showViewerNotice(getViewerFeatureUnavailableMessage('pageNavigation'));
      return;
    }

    if (isPdfDocument || isPdfCustomViewer) {
      changePdfPage(currentPdfPage + (direction === 'next' ? 1 : -1));
      return;
    }

    if (isOfficeDocument) {
      const nextPage = Math.min(
        Math.max(currentEmbeddedPage + (direction === 'next' ? 1 : -1), 1),
        Math.max(embeddedPageCount, 1)
      );

      if (nextPage !== currentEmbeddedPage) {
        changeEmbeddedPage(nextPage);
      }
      return;
    }
  };

  const toggleViewerLayout = () => {
    if (!viewerCapabilities.canChangeLayout) {
      showViewerNotice(getViewerFeatureUnavailableMessage('pageNavigation'));
      return;
    }

    setPdfScrollMode((mode) => (mode === 'single' ? 'vertical' : 'single'));
  };

  const fallbackToPresentationIframePreview = () => {
    if (!document || !props.context || !isPresentationDocument) {
      setPreviewError('Failed to load slide preview');
      return;
    }

    setPreviewImageUrls([]);
    setPreviewUrl(buildOfficePreviewUrl(
      props.context.pageContext.web.absoluteUrl,
      document.serverRelativeUrl,
      normalizedFileType,
      currentEmbeddedPage
    ));
    setPreviewError('');
  };

  const handlePresentationImageError = () => {
    setPreviewImageUrls((currentUrls) => {
      const remainingUrls = currentUrls.filter((url) => url !== previewUrl);
      const nextUrl = remainingUrls[0];

      if (nextUrl) {
        setPreviewUrl(nextUrl);
        return remainingUrls;
      }

      fallbackToPresentationIframePreview();
      return [];
    });
  };

  const canSyncTouchpadZoom = viewerCapabilities.canZoom && (isPresentationDocument || isSpreadsheetDocument || isTextDocument || isAudioDocument);

  const syncToolbarZoomFromWheel = React.useCallback((deltaY: number, anchorPoint?: { clientX: number; clientY: number }): boolean => {
    if (!canSyncTouchpadZoom || deltaY === 0) {
      return false;
    }

    const now = Date.now();
    if (now - lastTouchpadZoomActionRef.current < 40) {
      return true;
    }

    lastTouchpadZoomActionRef.current = now;
    const zoomDelta = deltaY < 0 ? 0.05 : -0.05;
    const updateZoomState = (previousZoom: number): number => {
      const nextZoom = previousZoom + zoomDelta;
      return Math.min(Math.max(Number(nextZoom.toFixed(2)), 0.25), 2.5);
    };

    if (usesOfficeIframeZoomSync && isSpreadsheetDocument) {
      setOfficeToolbarZoom((previousZoom) => {
        const nextZoom = updateZoomState(previousZoom);
        applyExcelIframeZoom(nextZoom);
        return nextZoom;
      });
      setIsFitToContainerMode(false);
      return true;
    }

    setEmbeddedZoom((previousZoom) => {
      const nextZoom = updateZoomState(previousZoom);
      if (isMediaDocument && anchorPoint && nextZoom !== previousZoom) {
        const container = previewContainerRef.current;
        if (container) {
          if (nextZoom < previousZoom) {
            centerMediaPreviewScroll();
          } else {
            const bounds = container.getBoundingClientRect();
            const anchorX = anchorPoint.clientX - bounds.left;
            const anchorY = anchorPoint.clientY - bounds.top;
            const contentX = container.scrollLeft + anchorX;
            const contentY = container.scrollTop + anchorY;
            const previousLayoutScale = Math.max(previousZoom, 1);
            const nextLayoutScale = Math.max(nextZoom, 1);
            const zoomRatio = nextLayoutScale / previousLayoutScale;

            window.requestAnimationFrame(() => {
              container.scrollLeft = Math.max(0, (contentX * zoomRatio) - anchorX);
              container.scrollTop = Math.max(0, (contentY * zoomRatio) - anchorY);
            });
          }
        }
      }
      setOfficeToolbarZoom(nextZoom);
      return nextZoom;
    });

    setIsFitToContainerMode(false);
    return true;
  }, [applyExcelIframeZoom, canSyncTouchpadZoom, centerMediaPreviewScroll, isMediaDocument, isSpreadsheetDocument, usesOfficeIframeZoomSync]);

  React.useEffect(() => {
    if (!canSyncTouchpadZoom) {
      return;
    }

    const clampZoom = (zoomValue: number): number =>
      Math.min(Math.max(Number(zoomValue.toFixed(2)), 0.25), 2.5);

    const isGestureInsidePreview = (event: Event): boolean => {
      const container = previewContainerRef.current;
      if (!container) {
        return false;
      }

      const target = event.target as Node | null;
      const eventPath = typeof event.composedPath === 'function' ? event.composedPath() : [];
      if (
        eventPath.indexOf(container) >= 0 ||
        (target instanceof Node && container.contains(target))
      ) {
        return true;
      }

      const gestureEvent = event as Event & { clientX?: number; clientY?: number };
      if (typeof gestureEvent.clientX !== 'number' || typeof gestureEvent.clientY !== 'number') {
        return false;
      }

      const bounds = container.getBoundingClientRect();
      return (
        gestureEvent.clientX >= bounds.left &&
        gestureEvent.clientX <= bounds.right &&
        gestureEvent.clientY >= bounds.top &&
        gestureEvent.clientY <= bounds.bottom
      );
    };

    const applyGestureZoom = (nextZoom: number): void => {
      const clampedZoom = clampZoom(nextZoom);
      if (usesOfficeIframeZoomSync && isSpreadsheetDocument) {
        applyExcelIframeZoom(clampedZoom);
      } else {
        setEmbeddedZoom(clampedZoom);
      }
      setOfficeToolbarZoom(clampedZoom);

      setIsFitToContainerMode(false);
    };

    const handleGestureStart = (event: Event): void => {
      if (!isGestureInsidePreview(event)) {
        return;
      }

      gestureStartToolbarZoomRef.current = usesOfficeIframeZoomSync ? officeToolbarZoom : embeddedZoom;
    };

    const handleGestureChange = (event: Event): void => {
      if (!isGestureInsidePreview(event)) {
        return;
      }

      const gestureScale = Number((event as Event & { scale?: number }).scale || 1);
      applyGestureZoom(gestureStartToolbarZoomRef.current * gestureScale);
    };

    window.addEventListener('gesturestart', handleGestureStart, { capture: true, passive: true } as AddEventListenerOptions);
    window.addEventListener('gesturechange', handleGestureChange, { capture: true, passive: true } as AddEventListenerOptions);
    window.document.addEventListener('gesturestart', handleGestureStart, { capture: true, passive: true } as AddEventListenerOptions);
    window.document.addEventListener('gesturechange', handleGestureChange, { capture: true, passive: true } as AddEventListenerOptions);

    return () => {
      window.removeEventListener('gesturestart', handleGestureStart, true);
      window.removeEventListener('gesturechange', handleGestureChange, true);
      window.document.removeEventListener('gesturestart', handleGestureStart, true);
      window.document.removeEventListener('gesturechange', handleGestureChange, true);
    };
  }, [applyExcelIframeZoom, canSyncTouchpadZoom, embeddedZoom, isSpreadsheetDocument, officeToolbarZoom, usesOfficeIframeZoomSync]);

  const attachOfficeIframeWheelSync = React.useCallback((iframe: HTMLIFrameElement): void => {
    if (officeIframeWheelCleanupRef.current) {
      officeIframeWheelCleanupRef.current();
      officeIframeWheelCleanupRef.current = null;
    }

    if (!usesOfficeIframeZoomSync) {
      return;
    }

    try {
      const iframeWindow = iframe.contentWindow;
      const iframeDocument = iframe.contentDocument;
      if (!iframeWindow && !iframeDocument) {
        return;
      }

      const handleIframeWheel = (event: WheelEvent): void => {
        if (!event.ctrlKey && !event.metaKey) {
          return;
        }

        void syncToolbarZoomFromWheel(event.deltaY);
      };

      const handleIframeGestureStart = (event: Event): void => {
        gestureStartToolbarZoomRef.current = usesOfficeIframeZoomSync ? officeToolbarZoom : embeddedZoom;
      };

      const handleIframeGestureChange = (event: Event): void => {
        const gestureScale = Number((event as Event & { scale?: number }).scale || 1);
        const nextZoom = gestureStartToolbarZoomRef.current * gestureScale;
        const clampedZoom = Math.min(Math.max(Number(nextZoom.toFixed(2)), 0.25), 2.5);

        if (usesOfficeIframeZoomSync && isSpreadsheetDocument) {
          applyExcelIframeZoom(clampedZoom);
        } else {
          setEmbeddedZoom(clampedZoom);
        }
        setOfficeToolbarZoom(clampedZoom);

        setIsFitToContainerMode(false);
      };

      iframeWindow?.addEventListener('wheel', handleIframeWheel, { capture: true, passive: true });
      iframeDocument?.addEventListener('wheel', handleIframeWheel, { capture: true, passive: true });
      iframeWindow?.addEventListener('gesturestart', handleIframeGestureStart, { capture: true, passive: true } as AddEventListenerOptions);
      iframeWindow?.addEventListener('gesturechange', handleIframeGestureChange, { capture: true, passive: true } as AddEventListenerOptions);
      iframeDocument?.addEventListener('gesturestart', handleIframeGestureStart, { capture: true, passive: true } as AddEventListenerOptions);
      iframeDocument?.addEventListener('gesturechange', handleIframeGestureChange, { capture: true, passive: true } as AddEventListenerOptions);
      officeIframeWheelCleanupRef.current = () => {
        iframeWindow?.removeEventListener('wheel', handleIframeWheel, true);
        iframeDocument?.removeEventListener('wheel', handleIframeWheel, true);
        iframeWindow?.removeEventListener('gesturestart', handleIframeGestureStart, true);
        iframeWindow?.removeEventListener('gesturechange', handleIframeGestureChange, true);
        iframeDocument?.removeEventListener('gesturestart', handleIframeGestureStart, true);
        iframeDocument?.removeEventListener('gesturechange', handleIframeGestureChange, true);
      };
    } catch (error) {
      officeIframeWheelCleanupRef.current = null;
    }
  }, [applyExcelIframeZoom, embeddedZoom, isSpreadsheetDocument, officeToolbarZoom, syncToolbarZoomFromWheel, usesOfficeIframeZoomSync]);

  const handlePreviewWheel = (event: React.WheelEvent<HTMLDivElement>) => {
    const now = Date.now();

    if (event.ctrlKey) {
      if (syncToolbarZoomFromWheel(event.deltaY, { clientX: event.clientX, clientY: event.clientY })) {
        event.preventDefault();
        event.stopPropagation();
      }
      return;
    }

    const isHorizontalSwipe = Math.abs(event.deltaX) > Math.abs(event.deltaY) && Math.abs(event.deltaX) > 24;
    if (!isHorizontalSwipe || !supportsPrecisePageNavigation) {
      return;
    }

    if (now - lastViewerWheelActionRef.current < 450) {
      return;
    }

    lastViewerWheelActionRef.current = now;
    changeViewerPage(event.deltaX > 0 ? 'next' : 'prev');
  };

  React.useEffect(() => {
    const handleDocumentWheel = (event: WheelEvent) => {
      if (!event.ctrlKey) {
        return;
      }

      const container = previewContainerRef.current;
      if (container) {
        const eventPath = typeof event.composedPath === 'function' ? event.composedPath() : [];
        const eventTarget = event.target;
        const isFromPreviewTarget =
          eventPath.indexOf(container) >= 0 ||
          (eventTarget instanceof Node && container.contains(eventTarget));
        if (!isFromPreviewTarget) {
          return;
        }

        const bounds = container.getBoundingClientRect();
        const isInsidePreview =
          event.clientX >= bounds.left &&
          event.clientX <= bounds.right &&
          event.clientY >= bounds.top &&
          event.clientY <= bounds.bottom;

        if (isInsidePreview) {
          if (syncToolbarZoomFromWheel(event.deltaY, { clientX: event.clientX, clientY: event.clientY })) {
            event.preventDefault();
            event.stopPropagation();
          }
        }
      }
    };

    window.addEventListener('wheel', handleDocumentWheel, { capture: true, passive: false });
    window.document.addEventListener('wheel', handleDocumentWheel, { capture: true, passive: false });
    return () => {
      window.removeEventListener('wheel', handleDocumentWheel, true);
      window.document.removeEventListener('wheel', handleDocumentWheel, true);
    };
  }, [syncToolbarZoomFromWheel]);

  React.useEffect(() => {
    return () => {
      if (officeIframeWheelCleanupRef.current) {
        officeIframeWheelCleanupRef.current();
        officeIframeWheelCleanupRef.current = null;
      }
    };
  }, [previewUrl]);

  const openFlagReviewModal = React.useCallback((): void => {
    if (disableSocialActions || isSubmittingFlagReview) {
      return;
    }

    setFlagReviewReason(REVIEW_FLAG_REASONS[REVIEW_FLAG_REASONS.length - 1]);
    setFlagReviewDetails('');
    setFlagReviewError('');
    setIsFlagReviewModalOpen(true);
  }, [disableSocialActions, isSubmittingFlagReview]);

  const closeFlagReviewModal = React.useCallback((): void => {
    if (isSubmittingFlagReview) {
      return;
    }

    setFlagReviewError('');
    setIsFlagReviewModalOpen(false);
  }, [isSubmittingFlagReview]);

  const submitFlagReview = React.useCallback(async (): Promise<void> => {
    if (!props.context || !document || !currentUserId || !flagReviewReason.trim()) {
      return;
    }

    if (!flagReviewDetails.trim()) {
      setFlagReviewError('Please enter comments before submitting.');
      return;
    }

    try {
      setIsSubmittingFlagReview(true);
      setFlagReviewError('');

      const webUrl = props.context.pageContext.web.absoluteUrl;
      const listName = 'DocumentReviewFlags';
      const listExists = await ensureListExists(listName, [
        { name: 'DocumentId', type: 'Number', required: true },
        { name: 'UserId', type: 'Number', required: true },
        { name: 'UserName', type: 'Text', required: false },
        { name: 'Reason', type: 'Text', required: true },
        { name: COLUMN_NAMES.comments, type: 'Note', required: true }
      ]);

      if (!listExists) {
        throw new Error('Unable to prepare review feedback storage.');
      }

      const fieldsReady = await ensureListFieldsExist(listName, [
        { name: 'DocumentId', type: 'Number', required: true },
        { name: 'UserId', type: 'Number', required: true },
        { name: 'UserName', type: 'Text', required: false },
        { name: 'Reason', type: 'Text', required: true },
        { name: COLUMN_NAMES.comments, type: 'Note', required: true }
      ]);
      const availableFieldNames = await getListFieldNames(listName);
      const canStoreComments = fieldsReady || availableFieldNames.has('comments') || availableFieldNames.has('details');
      const commentsFieldName = availableFieldNames.has('comments') ? COLUMN_NAMES.comments : availableFieldNames.has('details') ? 'Details' : COLUMN_NAMES.comments;
      const canStoreUserName = availableFieldNames.has('username');

      if (!fieldsReady && !canStoreComments) {
        throw new Error('Unable to prepare review feedback fields.');
      }

      const entityType = await getListEntityType(listName);
      if (!entityType) {
        throw new Error('Unable to resolve review feedback list type.');
      }

      const findResponse = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=DocumentId eq ${document.id} and UserId eq ${currentUserId}&$select=Id`,
        SPHttpClient.configurations.v1
      );

      const payload: Record<string, any> = {
        __metadata: { type: entityType },
        Title: document.name,
        DocumentId: document.id,
        UserId: currentUserId,
        Reason: flagReviewReason.trim()
      };

      if (canStoreUserName) {
        payload.UserName = currentUserName || 'User';
      }

      payload[commentsFieldName] = flagReviewDetails.trim();

      let createdNewFlag = false;
      let existingItemId: number | null = null;

      if (findResponse.ok) {
        const findData = await findResponse.json();
        existingItemId = findData.value?.[0]?.Id || null;
      } else {
        const errorText = await findResponse.text();
        console.warn('Unable to check for existing review feedback. Proceeding with create flow.', findResponse.status, errorText);
      }

      if (existingItemId) {
        const updateResponse = await props.context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${listName}')/items(${existingItemId})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'IF-MATCH': '*',
              'X-HTTP-Method': 'MERGE',
              'odata-version': ''
            },
            body: JSON.stringify(payload)
          }
        );

        if (!updateResponse.ok) {
          throw new Error('Unable to update your review feedback.');
        }
      } else {
        const createResponse = await props.context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${listName}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: JSON.stringify(payload)
          }
        );

        if (!createResponse.ok) {
          throw new Error('Unable to submit your review feedback.');
        }

        createdNewFlag = true;
      }

      setIsFlaggedForReview(true);
      if (createdNewFlag) {
        setFlagCount((current) => current + 1);
      }
      setIsFlagReviewModalOpen(false);
      setFlagReviewReason(REVIEW_FLAG_REASONS[REVIEW_FLAG_REASONS.length - 1]);
      setFlagReviewDetails('');
      setFlagReviewError('');
      setActionMessage('Feedback submitted successfully.');
    } catch (error) {
      console.error('Error submitting review feedback:', error);
      setActionMessage('Unable to submit feedback right now.');
    } finally {
      setIsSubmittingFlagReview(false);
    }
  }, [
    currentUserId,
    currentUserName,
    document,
    ensureListExists,
    ensureListFieldsExist,
    flagReviewDetails,
    flagReviewReason,
    getListFieldNames,
    getListEntityType,
    props.context
  ]);

  const handleFlagReviewAction = React.useCallback((): void => {
    if (disableSocialActions || isSubmittingFlagReview) {
      return;
    }

    if (isFlaggedForReview) {
      void toggleDocumentAction('flag');
      return;
    }

    openFlagReviewModal();
  }, [disableSocialActions, isFlaggedForReview, isSubmittingFlagReview, openFlagReviewModal]);

  const handlePrint = async () => {
    if (!document) {
      return;
    }

    if (isPrintInProgress) {
      showViewerNotice('Print is already opening. Please wait a moment.');
      return;
    }

    if (!canPrintDocument) {
      showViewerNotice('Print is restricted for your role.');
      return;
    }

    if (isPrintDisabled) {
      showViewerNotice(`Printing is disabled for ${document.fileType} files in the embedded viewer.`);
      return;
    }

    if (isTextDocument) {
      if (!textContent.trim()) {
        showViewerNotice('Text content is still loading, so there is nothing to print yet.');
        return;
      }

      setIsTextPrintMode(true);
      window.setTimeout(() => window.print(), 50);
      return;
    }

    if (isOfficeDocument) {
      const officePrintUrl = getOfficePrintUrl();
      const officePrintPreviewPages = await getOfficePrintPreviewPages();
      if (officePrintPreviewPages.length > 0 && openOfficePrintPreviewPopup(officePrintPreviewPages)) {
        return;
      }

      if (officePrintUrl && openNativeOfficePrintWindow(officePrintUrl)) {
        return;
      }

      if (triggerEmbeddedViewerPrint()) {
        return;
      }

      if (officePrintUrl && await printOfficeViewerInHiddenIframe(officePrintUrl)) {
        return;
      }

      if (previewUrl && await printOfficeViewerInHiddenIframe(previewUrl)) {
        return;
      }
    }

    if (isPdfDocument) {
      const pdfPrintUrl = await getPdfPrintUrl();
      if (pdfPrintUrl) {
        const printed = printUrlInHiddenIframe(pdfPrintUrl);
        if (pdfPrintUrl.startsWith('blob:') && pdfPrintUrl !== previewUrl) {
          window.setTimeout(() => window.URL.revokeObjectURL(pdfPrintUrl), 30000);
        }
        if (printed || openIsolatedPrintWindow(pdfPrintUrl)) {
          return;
        }
      }
    }

    if (previewUrl && openIsolatedPrintWindow(previewUrl)) {
      return;
    }

    showViewerNotice(`${document.fileType} preview is not ready to print yet.`);
  };

const getPdfViewerSrc = (): string => {
  if (!previewUrl) return '';

  const [baseUrl, existingFragment] = previewUrl.split('#');
  const fragmentParams = new URLSearchParams(existingFragment || '');
  Object.keys(PDF_VIEWER_FRAGMENT_OPTIONS).forEach((key) => {
    fragmentParams.set(key, PDF_VIEWER_FRAGMENT_OPTIONS[key]);
  });
  fragmentParams.set('page', String(Math.max(currentPdfPage || 1, 1)));
  fragmentParams.set('zoom', isFitToContainerMode ? 'page-fit' : String(Math.max(Math.round(embeddedZoom * 100), 25)));
  fragmentParams.set('view', pdfScrollMode === 'single' ? 'Fit' : 'FitV');

  return `${baseUrl}#${fragmentParams.toString()}`;
};

  React.useEffect(() => {
    const handleRestrictedViewerShortcut = (event: KeyboardEvent): void => {
      const isModifiedShortcut = event.ctrlKey || event.metaKey;
      if (!isModifiedShortcut) {
        return;
      }

      const key = event.key.toLowerCase();
      if (key === 's') {
        event.preventDefault();
        event.stopPropagation();
        showViewerNotice('Save and download are disabled in view-only mode.');
        return;
      }

      if (key === 'z' || key === 'y' || (key === 'u' && event.shiftKey)) {
        event.preventDefault();
        event.stopPropagation();
        showViewerNotice('Restricted viewer actions are disabled in view-only mode.');
        return;
      }

      if (key === 'p' && !canPrintDocument) {
        event.preventDefault();
        event.stopPropagation();
        showViewerNotice('Print is restricted for your role.');
      }
    };

    window.addEventListener('keydown', handleRestrictedViewerShortcut, { capture: true });
    window.document.addEventListener('keydown', handleRestrictedViewerShortcut, { capture: true });

    return () => {
      window.removeEventListener('keydown', handleRestrictedViewerShortcut, true);
      window.document.removeEventListener('keydown', handleRestrictedViewerShortcut, true);
    };
  }, [canPrintDocument, showViewerNotice]);

  const findViewerControl = React.useCallback((frameDocument: Document, terms: string[]): HTMLElement | null => {
    const candidates = Array.from(
      frameDocument.querySelectorAll('button, [role="button"], [aria-label], [title], input[type="button"], input[type="submit"]')
    ) as HTMLElement[];

    const normalizedTerms = terms.map((term) => term.toLowerCase());

    for (const candidate of candidates) {
      const dataAutomationId = candidate.getAttribute('data-automationid') || '';
      const dataTooltip = candidate.getAttribute('data-tooltip') || '';
      const candidateId = candidate.getAttribute('id') || '';
      const candidateName = candidate.getAttribute('name') || '';
      const ariaKeyShortcuts = candidate.getAttribute('aria-keyshortcuts') || '';
      const className = typeof candidate.className === 'string' ? candidate.className : '';
      const candidateText = [
        candidate.getAttribute('aria-label') || '',
        candidate.getAttribute('title') || '',
        candidate.textContent || '',
        candidate.innerText || '',
        dataAutomationId,
        dataTooltip,
        candidateId,
        candidateName,
        ariaKeyShortcuts,
        className
      ].join(' ').toLowerCase().trim();

      if (normalizedTerms.some((term) => candidateText.includes(term))) {
        return candidate;
      }
    }

    return null;
  }, []);

  const findViewerControlInTree = React.useCallback((rootDocument: Document, terms: string[]): HTMLElement | null => {
    const directMatch = findViewerControl(rootDocument, terms);
    if (directMatch) {
      return directMatch;
    }

    const nestedFrames = Array.from(rootDocument.querySelectorAll('iframe')) as HTMLIFrameElement[];
    for (const nestedFrame of nestedFrames) {
      try {
        const nestedDocument = nestedFrame.contentDocument;
        if (!nestedDocument) {
          continue;
        }

        const nestedMatch = findViewerControlInTree(nestedDocument, terms);
        if (nestedMatch) {
          return nestedMatch;
        }
      } catch (error) {
        console.warn('Unable to inspect nested viewer frame:', error);
      }
    }

    return null;
  }, [findViewerControl]);

  const triggerViewerPrintShortcut = React.useCallback((rootWindow: Window, rootDocument: Document): boolean => {
    try {
      const target = (rootDocument.activeElement as HTMLElement | null) || rootDocument.body;
      target?.focus?.();

      const keyboardEvent = new KeyboardEvent('keydown', {
        bubbles: true,
        cancelable: true,
        ctrlKey: true,
        key: 'p',
        code: 'KeyP'
      });

      Object.defineProperty(keyboardEvent, 'keyCode', { get: () => 80 });
      Object.defineProperty(keyboardEvent, 'which', { get: () => 80 });

      target?.dispatchEvent(keyboardEvent);
      rootDocument.dispatchEvent(keyboardEvent);
      rootWindow.dispatchEvent(keyboardEvent);
      return true;
    } catch (error) {
      console.warn('Unable to trigger viewer print shortcut:', error);
      return false;
    }
  }, []);

  const triggerEmbeddedViewerPrint = React.useCallback((): boolean => {
    const frame = iframeRef.current;
    const frameWindow = frame?.contentWindow;
    const frameDocument = frame?.contentDocument;

    if (!frameWindow || !frameDocument) {
      return false;
    }

    const tryPrintWithinDocumentTree = (currentWindow: Window, currentDocument: Document): boolean => {
      const printControl = findViewerControlInTree(currentDocument, ['print', 'print document', 'printer']);
      if (printControl) {
        printControl.click();
        printControl.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
        return true;
      }

      if (triggerViewerPrintShortcut(currentWindow, currentDocument)) {
        return true;
      }

      const nestedFrames = Array.from(currentDocument.querySelectorAll('iframe')) as HTMLIFrameElement[];
      for (const nestedFrame of nestedFrames) {
        try {
          const nestedWindow = nestedFrame.contentWindow;
          const nestedDocument = nestedFrame.contentDocument;

          if (!nestedWindow || !nestedDocument) {
            continue;
          }

          if (tryPrintWithinDocumentTree(nestedWindow, nestedDocument)) {
            return true;
          }
        } catch (error) {
          console.warn('Unable to trigger print inside nested viewer frame:', error);
        }
      }

      return false;
    };

    try {
      return tryPrintWithinDocumentTree(frameWindow, frameDocument);
    } catch (error) {
      console.warn('Unable to access embedded viewer for printing:', error);
      return false;
    }
  }, [findViewerControlInTree, triggerViewerPrintShortcut]);

  const triggerOfficeViewerPrintInDocument = React.useCallback((rootWindow: Window, rootDocument: Document): boolean => {
    const printControl = findViewerControlInTree(rootDocument, ['print', 'print document', 'printer']);
    if (printControl) {
      printControl.click();
      printControl.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      return true;
    }

    return triggerViewerPrintShortcut(rootWindow, rootDocument);
  }, [findViewerControlInTree, triggerViewerPrintShortcut]);

  const openIsolatedPrintWindow = React.useCallback((
    url: string,
    onLoaded?: (printWindow: Window, printDocument: Document) => boolean | void,
    allowWindowPrint: boolean = true
  ): boolean => {
    const printWindow = window.open('', 'documentPrintViewer', buildPrintPopupFeatures());
    if (!printWindow) {
      return false;
    }

    const cleanup = () => {
      window.setTimeout(() => {
        try {
          printWindow.close();
        } catch (error) {
          console.warn('Unable to close isolated print window:', error);
        }
      }, 250);
    };

    printWindow.addEventListener('afterprint', cleanup, { once: true });
    printWindow.location.replace(url);
    printWindow.addEventListener('load', () => {
      let handled = false;
      try {
        printWindow.focus();
        if (onLoaded && printWindow.document) {
          handled = Boolean(onLoaded(printWindow, printWindow.document));
        }

        if (!handled && allowWindowPrint) {
          printWindow.print();
        }
      } catch (error) {
        console.warn('Unable to print from isolated window:', error);
      }
    }, { once: true });

    return true;
  }, []);

  const printUrlInHiddenIframe = React.useCallback((url: string, onLoaded?: (frame: HTMLIFrameElement) => boolean | void): boolean => {
    const frame = window.document.createElement('iframe');
    frame.setAttribute('aria-hidden', 'true');
    frame.style.position = 'fixed';
    frame.style.left = '-10000px';
    frame.style.top = '0';
    frame.style.width = '1366px';
    frame.style.height = '960px';
    frame.style.border = '0';
    frame.style.opacity = '0';
    frame.style.pointerEvents = 'none';

    let cleanedUp = false;
    const cleanup = () => {
      if (cleanedUp) {
        return;
      }

      cleanedUp = true;
      try {
        frame.remove();
      } catch (error) {
        console.warn('Unable to remove hidden print iframe:', error);
      }
    };

    const handleLoad = () => {
      const frameWindow = frame.contentWindow;
      if (!frameWindow) {
        cleanup();
        return;
      }

      let handled = false;
      if (onLoaded) {
        try {
          handled = Boolean(onLoaded(frame));
        } catch (error) {
          console.warn('Custom iframe print handler failed:', error);
        }
      }

      if (!handled) {
        try {
          frameWindow.focus();
          frameWindow.print();
        } catch (error) {
          console.warn('Unable to print from hidden iframe:', error);
          cleanup();
          return;
        }
      }

      frameWindow.addEventListener('afterprint', cleanup, { once: true });
      window.setTimeout(cleanup, 30000);
    };

    frame.addEventListener('load', handleLoad, { once: true });
    frame.src = url;
    window.document.body.appendChild(frame);
    return true;
  }, []);

  const printOfficeViewerInHiddenIframe = React.useCallback(async (url: string): Promise<boolean> => {
    setIsPrintInProgress(true);
    showViewerNotice('Preparing document print view...');

    return await new Promise<boolean>((resolve) => {
      const frame = window.document.createElement('iframe');
      frame.setAttribute('aria-hidden', 'true');
      frame.style.position = 'fixed';
      frame.style.left = '-10000px';
      frame.style.top = '0';
      frame.style.width = '1366px';
      frame.style.height = '960px';
      frame.style.border = '0';
      frame.style.opacity = '0';
      frame.style.pointerEvents = 'none';

      let settled = false;
      const timeoutIds: number[] = [];

      const cleanup = (didPrint: boolean) => {
        if (settled) {
          return;
        }

        settled = true;
        timeoutIds.forEach((timeoutId) => window.clearTimeout(timeoutId));
        setIsPrintInProgress(false);

        try {
          frame.remove();
        } catch (error) {
          console.warn('Unable to remove Office hidden print iframe:', error);
        }

        resolve(didPrint);
      };

      frame.addEventListener('load', () => {
        const frameWindow = frame.contentWindow;
        const frameDocument = frame.contentDocument || frameWindow?.document;

        if (!frameWindow || !frameDocument) {
          cleanup(false);
          return;
        }

        const attemptDelays = [250, 800, 1800, 3200, 5000, 7500];
        attemptDelays.forEach((delay, index) => {
          const timeoutId = window.setTimeout(() => {
            if (settled || !frame.isConnected) {
              return;
            }

            try {
              frameWindow.focus();

              if (triggerOfficeViewerPrintInDocument(frameWindow, frameDocument)) {
                cleanup(true);
                return;
              }

              if (index === attemptDelays.length - 1) {
                cleanup(false);
              }
            } catch (error) {
              console.warn('Unable to print Office viewer from hidden iframe:', error);
              if (index === attemptDelays.length - 1) {
                cleanup(false);
              }
            }
          }, delay);

          timeoutIds.push(timeoutId);
        });
      }, { once: true });

      const safetyTimeoutId = window.setTimeout(() => cleanup(false), 11000);
      timeoutIds.push(safetyTimeoutId);

      frame.src = url;
      window.document.body.appendChild(frame);
    });
  }, [showViewerNotice, triggerOfficeViewerPrintInDocument]);

  const getOfficePrintUrl = React.useCallback((): string | null => {
    if (!document || !props.context || !isOfficeDocument) {
      return null;
    }

    return buildOfficePrintUrl(
      props.context.pageContext.web.absoluteUrl,
      normalizeServerRelativeUrl(document.serverRelativeUrl),
      normalizedFileType
    );
  }, [document, isOfficeDocument, normalizedFileType, props.context]);

  const openNativeOfficePrintWindow = React.useCallback((url: string): boolean => {
    const printWindow = window.open('', 'documentPrintViewer', buildPrintPopupFeatures());
    if (!printWindow) {
      return false;
    }

    setIsPrintInProgress(true);
    showViewerNotice('Opening full document print view...');

    let cleanedUp = false;
    let closeWatcherId: number | null = null;
    const timeoutIds: number[] = [];

    const cleanup = (shouldCloseWindow: boolean = false) => {
      if (cleanedUp) {
        return;
      }

      cleanedUp = true;
      setIsPrintInProgress(false);

      if (closeWatcherId !== null) {
        window.clearInterval(closeWatcherId);
      }

      timeoutIds.forEach((timeoutId) => window.clearTimeout(timeoutId));

      if (shouldCloseWindow) {
        window.setTimeout(() => {
          try {
            printWindow.close();
          } catch (error) {
            console.warn('Unable to close native Office print window:', error);
          }
        }, 250);
      }
    };

    const tryTriggerPrint = (): boolean => {
      try {
        printWindow.focus();
        return Boolean(
          printWindow.document &&
          triggerOfficeViewerPrintInDocument(printWindow, printWindow.document)
        );
      } catch (error) {
        console.warn('Unable to trigger native Office print from popup:', error);
        return false;
      }
    };

    const schedulePrintAttempts = () => {
      const attemptDelays = [400, 1200, 2400, 4000, 6500, 9000];

      attemptDelays.forEach((delay, index) => {
        const timeoutId = window.setTimeout(() => {
          if (tryTriggerPrint()) {
            const afterPrintCleanup = () => cleanup(true);
            try {
              printWindow.addEventListener('afterprint', afterPrintCleanup, { once: true });
            } catch (error) {
              console.warn('Unable to attach afterprint handler to native Office print window:', error);
            }

            timeoutIds.push(window.setTimeout(() => cleanup(false), 2500));
            timeoutIds.push(window.setTimeout(() => cleanup(true), 30000));
            return;
          }

          if (index === attemptDelays.length - 1) {
            cleanup(false);
          }
        }, delay);

        timeoutIds.push(timeoutId);
      });
    };

    closeWatcherId = window.setInterval(() => {
      if (printWindow.closed) {
        cleanup(false);
      }
    }, 700);

    printWindow.addEventListener('load', schedulePrintAttempts, { once: true });
    printWindow.location.replace(url);
    timeoutIds.push(window.setTimeout(() => cleanup(false), 20000));

    return true;
  }, [showViewerNotice, triggerOfficeViewerPrintInDocument]);

  const getOfficePrintPreviewUrls = React.useCallback((): string[] => {
    if (!document || !props.context || !isOfficeDocument) {
      return [];
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const serverRelativeUrl = normalizeServerRelativeUrl(document.serverRelativeUrl);
    const siteId = String(props.context.pageContext.site.id || '');
    const webId = String(props.context.pageContext.web.id || '');
    const totalPages = Math.max(embeddedPageCount, 1);

    return Array.from({ length: totalPages }, (_, index) =>
      buildIndexedPreviewImageUrl(
        webUrl,
        serverRelativeUrl,
        index + 1,
        document.fileUniqueId,
        siteId,
        webId
      )
    );
  }, [document, embeddedPageCount, isOfficeDocument, props.context]);

  const getOfficePrintPreviewPages = React.useCallback(async (): Promise<OfficePrintPreviewPage[]> => {
    if (!props.context) {
      return [];
    }

    const previewUrls = getOfficePrintPreviewUrls();
    if (previewUrls.length === 0) {
      return [];
    }

    const officeFrameUrls = Array.from({ length: previewUrls.length }, (_, index) =>
      document
        ? buildOfficePreviewUrl(
          props.context.pageContext.web.absoluteUrl,
          normalizeServerRelativeUrl(document.serverRelativeUrl),
          normalizedFileType,
          index + 1
        )
        : ''
    );

    const previewSources = await Promise.all(
      previewUrls.map(async (url, index) => {
        const officeFrameUrl = officeFrameUrls[index];
        try {
          const response = await window.fetch(url, {
            method: 'GET',
            credentials: 'include',
            headers: {
              Accept: 'image/*,*/*;q=0.8'
            }
          });

          if (!response.ok) {
            return { kind: 'iframe' as const, src: officeFrameUrl || url };
          }

          const contentType = response.headers.get('content-type') || '';
          if (!contentType.toLowerCase().startsWith('image/')) {
            console.warn('Office print preview request did not return an image:', { url, contentType });
            return { kind: 'iframe' as const, src: officeFrameUrl || url };
          }

          const blob = await response.blob();
          return { kind: 'image' as const, src: window.URL.createObjectURL(blob) };
        } catch (error) {
          console.warn('Unable to fetch Office print preview image, falling back to direct URL:', error);
          return { kind: 'iframe' as const, src: officeFrameUrl || url };
        }
      })
    );

    const imagePages = previewSources.filter((page) => page.kind === 'image');
    if (imagePages.length > 0) {
      return imagePages;
    }

    return previewSources;
  }, [document, getOfficePrintPreviewUrls, normalizedFileType, props.context]);

  const openOfficePrintPreviewPopup = React.useCallback((pages: OfficePrintPreviewPage[]): boolean => {
    if (!document || pages.length === 0) {
      return false;
    }

    setIsPrintInProgress(true);
    showViewerNotice('Preparing full document print preview...');

    const frame = window.document.createElement('iframe');
    frame.setAttribute('aria-hidden', 'true');
    frame.style.position = 'fixed';
    frame.style.left = '-10000px';
    frame.style.top = '0';
    frame.style.width = '1366px';
    frame.style.height = '960px';
    frame.style.border = '0';
    frame.style.opacity = '0';
    frame.style.pointerEvents = 'none';

    let cleanedUp = false;
    const cleanup = () => {
      if (cleanedUp) {
        return;
      }

      cleanedUp = true;
      setIsPrintInProgress(false);
      if (printResetTimeoutRef.current) {
        window.clearTimeout(printResetTimeoutRef.current);
        printResetTimeoutRef.current = null;
      }
      pages.forEach((page) => {
        if (page.kind === 'image' && page.src.startsWith('blob:')) {
          try {
            window.URL.revokeObjectURL(page.src);
          } catch (error) {
            console.warn('Unable to revoke Office print preview blob URL:', error);
          }
        }
      });
      try {
        frame.remove();
      } catch (error) {
        console.warn('Unable to remove Office print iframe:', error);
      }
    };

    const pageMarkup = pages
      .map((page, index) => `
        <section class="print-page">
          ${page.kind === 'iframe'
          ? `<iframe src="${page.src}" title="${escapeHtml(document.name)} page ${index + 1}" loading="eager"></iframe>`
          : `<img src="${page.src}" alt="${escapeHtml(document.name)} page ${index + 1}" loading="eager" decoding="sync" />`
        }
        </section>
      `)
      .join('');

    const html = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>${escapeHtml(document.name)} - Print Preview</title>
    <style>
      html, body {
        margin: 0;
        padding: 0;
        background: #ffffff;
        font-family: Arial, sans-serif;
        min-height: 100vh;
      }
      .print-shell {
        padding: 0;
        width: 100%;
        max-width: 1100px;
        margin: 0 auto;
      }
      .print-page {
        width: 100%;
        max-width: 1100px;
        margin: 0 auto;
        display: flex;
        justify-content: center;
        page-break-after: always;
      }
      .print-page:last-child {
        page-break-after: auto;
      }
      .print-page img {
        width: 100%;
        height: auto;
        display: block;
        background: #ffffff;
        border: 0;
      }
      .print-page iframe {
        width: 100%;
        min-height: 780px;
        display: block;
        background: #ffffff;
        border: 0;
      }
      @page {
        margin: 12mm;
      }
    </style>
  </head>
  <body>
    <main class="print-shell">${pageMarkup}</main>
  </body>
</html>`;

    window.document.body.appendChild(frame);

    const frameDocument = frame.contentWindow?.document || frame.contentDocument;
    const frameWindow = frame.contentWindow;
    if (!frameDocument || !frameWindow) {
      cleanup();
      return false;
    }

    frameDocument.open();
    frameDocument.write(html);
    frameDocument.close();

    const resourcePromises = [
      ...Array.from(frameDocument.images).map((image) => new Promise<void>((resolve) => {
        if (image.complete) {
          resolve();
          return;
        }

        image.addEventListener('load', () => resolve(), { once: true });
        image.addEventListener('error', () => resolve(), { once: true });
      })),
      ...Array.from(frameDocument.querySelectorAll('iframe')).map((iframe) => new Promise<void>((resolve) => {
        iframe.addEventListener('load', () => resolve(), { once: true });
        window.setTimeout(() => resolve(), 7000);
      }))
    ];

    Promise.all(resourcePromises).finally(() => {
      window.setTimeout(() => {
        try {
          frameWindow.focus();
          frameWindow.print();
          frameWindow.addEventListener('afterprint', cleanup, { once: true });
          printResetTimeoutRef.current = window.setTimeout(cleanup, 45000);
        } catch (error) {
          console.warn('Unable to print Office preview from hidden iframe:', error);
          cleanup();
        }
      }, 200);
    });

    return true;
  }, [document, showViewerNotice]);

  const getPdfPrintUrl = React.useCallback(async (): Promise<string | null> => {
    if (!document || !props.context || !isPdfDocument) {
      return null;
    }

    if (previewUrl && previewUrl.startsWith('blob:')) {
      return previewUrl;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const serverRelativeUrl = normalizeServerRelativeUrl(document.serverRelativeUrl);
    const downloadUrl = buildFileApiUrl(webUrl, serverRelativeUrl);

    try {
      const response = await props.context.spHttpClient.get(
        downloadUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return buildAbsoluteFileUrl(webUrl, serverRelativeUrl);
      }

      const blob = await response.blob();
      return window.URL.createObjectURL(new Blob([blob], { type: 'application/pdf' }));
    } catch (error) {
      console.warn('Unable to fetch PDF print source, falling back to absolute file URL:', error);
      return buildAbsoluteFileUrl(webUrl, serverRelativeUrl);
    }
  }, [document, isPdfDocument, previewUrl, props.context]);

  if (loading) {
    return (
      <div className={styles.documentDetailPage} onClick={(e) => e.stopPropagation()}>
        <div className={styles.loading}>Loading document...</div>
      </div>
    );
  }

  if (!document) {
    return (
      <div className={styles.documentDetailPage} onClick={(e) => e.stopPropagation()}>
        <div className={styles.errorContainer}>
          <h2 className={styles.errorTitle}>Document not found</h2>
          {fetchError && <p className={styles.errorMessage}>{fetchError}</p>}
          <p className={styles.errorSubtitle}>Document ID: {props.documentId}</p>
        </div>
      </div>
    );
  }

  const normalizedEmbeddedRotation = ((embeddedRotation % 360) + 360) % 360;
  const isEmbeddedRotationActive = normalizedEmbeddedRotation !== 0;
  const isEmbeddedRotated = normalizedEmbeddedRotation === 90 || normalizedEmbeddedRotation === 270;
  const embeddedRotationBaseWidth = Math.max(previewViewportWidth - 16, 320);
  const embeddedRotationBaseHeight =
    isPresentationDocument || isVideoDocument
      ? Math.max(Math.round(embeddedRotationBaseWidth * 9 / 16), 240)
      : Math.max(previewViewportHeight - 16, 460);
  const rotationScaleModifier =
    isEmbeddedRotated && (isPresentationDocument || isSpreadsheetDocument)
      ? embeddedRotationBaseWidth / embeddedRotationBaseHeight
      : 1;
  const effectiveEmbeddedScale = embeddedZoom * rotationScaleModifier;
  const embeddedCssScale = isPdfDocument ? rotationScaleModifier : effectiveEmbeddedScale;
  const embeddedRotatedVisualWidth = (isEmbeddedRotated ? embeddedRotationBaseHeight : embeddedRotationBaseWidth) * embeddedCssScale;
  const embeddedRotatedVisualHeight = (isEmbeddedRotated ? embeddedRotationBaseWidth : embeddedRotationBaseHeight) * embeddedCssScale;
  const getEmbeddedTransform = (): string => {
    const scaleTransform = `scale(${embeddedCssScale})`;

    switch (normalizedEmbeddedRotation) {
      case 90:
        return `translateX(${embeddedRotationBaseHeight * embeddedCssScale}px) rotate(90deg) ${scaleTransform}`;
      case 180:
        return `translate(${embeddedRotationBaseWidth * embeddedCssScale}px, ${embeddedRotationBaseHeight * embeddedCssScale}px) rotate(180deg) ${scaleTransform}`;
      case 270:
        return `translateY(${embeddedRotationBaseWidth * embeddedCssScale}px) rotate(270deg) ${scaleTransform}`;
      default:
        return scaleTransform;
    }
  };
  const isMediaZoomedIn = isMediaDocument && embeddedCssScale > 1.01;
  const previewContainerStyle: React.CSSProperties = isPdfDocument
    ? {
      overflow: 'hidden',
      overscrollBehavior: 'contain',
      background: '#ffffff',
      padding: 0
    }
    : isTextDocument
    ? isEmbeddedRotationActive
      ? {
        overflowX: 'auto',
        overflowY: 'auto',
        overscrollBehavior: 'contain',
        background: '#ffffff',
        padding: 0
      }
      : {
        overflow: 'hidden',
        overscrollBehavior: 'contain',
        background: '#ffffff',
        padding: 0
      }
    : isMediaDocument
    ? {
      overflowX: isVideoDocument ? 'hidden' : isMediaZoomedIn ? 'auto' : 'hidden',
      overflowY: isVideoDocument ? 'hidden' : isMediaZoomedIn ? 'auto' : 'hidden',
      overscrollBehavior: 'contain',
      background: '#ffffff',
      padding: 0,
      justifyContent: isVideoDocument ? 'center' : undefined,
      alignItems: isVideoDocument ? 'center' : undefined
    }
    : isPresentationDocument && isEmbeddedRotationActive
    ? {
      overflowX: 'auto',
      overflowY: 'auto',
      overscrollBehavior: 'contain',
      background: '#f3f6fb',
      padding: 0
    }
    : isSpreadsheetDocument
    ? isEmbeddedRotationActive
      ? { overflowX: 'auto', overflowY: 'auto', overscrollBehavior: 'contain' }
      : { overflow: 'hidden' }
    : pdfScrollMode === 'horizontal'
      ? { overflowX: 'auto', overflowY: 'hidden' }
      : isEmbeddedRotationActive
        ? { overflowX: 'auto', overflowY: 'auto', overscrollBehavior: 'contain' }
      : !isPdfDocument && effectiveEmbeddedScale > 1.01
      ? { overflowX: 'auto', overflowY: 'auto', overscrollBehavior: 'contain' }
      : isPdfDocument
      ? { overflowX: 'auto', overflowY: 'auto', overscrollBehavior: 'contain' }
      : { overflowX: 'hidden', overflowY: 'auto' };

  const shouldCenterZoomedPreview = (embeddedCssScale < 1 || isMediaDocument) && !isEmbeddedRotationActive && !isSpreadsheetDocument;
  const embeddedPreviewStyle: React.CSSProperties = !isPdfCustomViewer
    ? {
      transform: isMediaDocument || isTextDocument ? 'none' : getEmbeddedTransform(),
      transformOrigin: isMediaDocument || shouldCenterZoomedPreview ? 'center center' : 'top left',
      position: isEmbeddedRotationActive ? 'absolute' : undefined,
      top: isEmbeddedRotationActive ? 0 : undefined,
      left: isEmbeddedRotationActive ? 0 : undefined,
      width: isVideoDocument
        ? '100%'
        : isMediaDocument
        ? '100%'
        : isTextDocument
        ? '100%'
        : isEmbeddedRotationActive
        ? `${embeddedRotationBaseWidth}px`
        : embeddedCssScale > 1
          ? `${100 / embeddedCssScale}%`
          : undefined,
      height: isVideoDocument
        ? '100%'
        : isEmbeddedRotationActive
          ? `${embeddedRotationBaseHeight}px`
          : undefined,
      minHeight: isVideoDocument
        ? '100%'
        : isEmbeddedRotationActive
          ? `${embeddedRotationBaseHeight}px`
          : undefined
    }
    : {};
  const textPreviewContentStyle: React.CSSProperties = isTextDocument
    ? {
      fontSize: `${14 * embeddedZoom}px`,
      lineHeight: 1.55,
      minWidth: '100%'
    }
    : {};
  const embeddedRotationWrapperStyle: React.CSSProperties = isEmbeddedRotationActive
    ? {
      position: 'relative',
      flex: '0 0 auto',
      width: `${embeddedRotatedVisualWidth}px`,
      height: `${embeddedRotatedVisualHeight}px`,
      minWidth: `${embeddedRotatedVisualWidth}px`,
      minHeight: `${embeddedRotatedVisualHeight}px`,
      overflow: 'visible'
    }
    : { display: 'contents' };
  const embeddedPreviewStageStyle: React.CSSProperties = !isPdfCustomViewer
    ? {
      display: 'flex',
      alignItems: shouldCenterZoomedPreview ? 'center' : 'flex-start',
      justifyContent: shouldCenterZoomedPreview ? 'center' : 'flex-start',
      width: isVideoDocument
        ? '100%'
        : isMediaDocument
        ? `${Math.max(embeddedCssScale, 1) * 100}%`
        : isTextDocument
          ? '100%'
        : isEmbeddedRotationActive
        ? `${Math.max(embeddedRotatedVisualWidth, previewViewportWidth)}px`
        : embeddedCssScale > 1
          ? `${embeddedCssScale * 100}%`
          : '100%',
      minWidth: isVideoDocument
        ? '100%'
        : isMediaDocument
          ? `${Math.max(embeddedCssScale, 1) * 100}%`
          : isTextDocument
            ? '100%'
          : '100%',
      maxWidth: isTextDocument ? '100%' : embeddedCssScale > 1 || isEmbeddedRotationActive ? 'none' : '100%',
      minHeight: isVideoDocument
        ? '100%'
        : isMediaDocument
        ? `${Math.max(embeddedCssScale, 1) * 100}%`
        : isTextDocument
        ? '100%'
        : isEmbeddedRotationActive
        ? `${Math.max(embeddedRotatedVisualHeight, previewViewportHeight)}px`
        : embeddedCssScale > 1
          ? `${embeddedCssScale * 100}%`
          : '100%',
      height: isEmbeddedRotationActive
        ? `${Math.max(embeddedRotatedVisualHeight, previewViewportHeight)}px`
        : isVideoDocument
          ? '100%'
          : undefined,
      overflow: 'visible',
      padding: isPresentationDocument && isEmbeddedRotationActive ? '0 0 16px 0' : undefined
    }
    : {};
  const detailUserName = props.userName || props.context?.pageContext.user.displayName;
  const detailUserEmail = props.userEmail || props.context?.pageContext.user.email || props.context?.pageContext.user.loginName;
  const showDetailReviewerNav = props.showReviewerNav || isKmAdmin || isApprover;
  const hideDetailDocumentsNav = props.hideDocumentsNav || isLearner;
  const handleDetailFooterBackHome = (): void => {
    if (props.onHomeOpen) {
      props.onHomeOpen();
      return;
    }

    pushPageUrl(NAV_PATHS.home);
  };
  const keepViewerToolbarItemsVisible = (): undefined => undefined;
  const togglePreviewSectionFullscreen = (): void => {
    const container = previewSectionRef.current;
    if (!container) {
      return;
    }

    if (window.document.fullscreenElement === container) {
      void window.document.exitFullscreen?.();
      return;
    }

    if (container.requestFullscreen) {
      void container.requestFullscreen();
      return;
    }

    setIsWindowMode((value) => !value);
  };
  const toggleExcelViewerFullscreen = (): void => {
    const container = excelPreviewShellRef.current;
    if (!container) {
      return;
    }

    if (window.document.fullscreenElement === container) {
      void window.document.exitFullscreen?.();
      return;
    }

    void container.requestFullscreen?.();
  };
  const excelViewerLeftToolbarItems: ICommandBarItemProps[] = [];
  const excelViewerCenterToolbarItems: ICommandBarItemProps[] = [
    makeIconOnlyToolbarItem({
      key: 'zoomOut',
      iconProps: { iconName: 'ZoomOut' },
      onClick: () => updateZoom('out')
    }, 'Zoom out'),
    {
      key: 'zoomValue',
      onRender: () => (
        <label className={styles.excelPdfToolbarZoomValue} title="Edit zoom percentage">
          <input
            className={styles.toolbarZoomInput}
            value={isEditingZoom ? zoomInputValue : String(activeZoomPercent)}
            inputMode="numeric"
            aria-label="Zoom percentage"
            onFocus={(event) => {
              setIsEditingZoom(true);
              setZoomInputValue(String(activeZoomPercent));
              event.currentTarget.select();
            }}
            onChange={(event) => setZoomInputValue(event.currentTarget.value)}
            onBlur={commitZoomInput}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                event.currentTarget.blur();
              } else if (event.key === 'Escape') {
                setZoomInputValue(String(activeZoomPercent));
                setIsEditingZoom(false);
                event.currentTarget.blur();
              }
            }}
          />
          <span aria-hidden="true">%</span>
        </label>
      )
    },
    makeIconOnlyToolbarItem({
      key: 'zoomIn',
      iconProps: { iconName: 'ZoomIn' },
      onClick: () => updateZoom('in')
    }, 'Zoom in')
  ];
  const excelViewerRightToolbarItems: ICommandBarItemProps[] = viewerCapabilities.canWindowMode
    ? [
      makeIconOnlyToolbarItem({
        key: 'fullscreen',
        iconProps: { iconName: isExcelViewerFullscreen ? 'BackToWindow' : 'FullScreen' },
        onClick: toggleExcelViewerFullscreen
      }, isExcelViewerFullscreen ? 'Exit fullscreen' : 'Fullscreen')
    ]
    : [];

  return (
    <div className={`${styles.documentDetailPage} ${isTextPrintMode ? styles.documentPrintTextMode : ''}`} onClick={(e) => e.stopPropagation()}>
      <IKShellHeader
        userName={detailUserName}
        userEmail={detailUserEmail}
        userPhotoUrl={props.userPhotoUrl}
        compact={true}
        onLogout={props.onLogout}
        onHomeOpen={props.onHomeOpen}
        onAllDocumentsOpen={props.onAllDocumentsOpen}
        onBusinessUnitsOpen={props.onBusinessUnitsOpen}
        onBookmarksOpen={props.onBookmarksOpen}
        onDocumentsOpen={props.onDocumentsOpen}
        onContactOpen={props.onContactOpen}
        onAuditLogOpen={props.onAuditLogOpen}
        onAnalyticsOpen={props.onAnalyticsOpen}
        showReviewerNav={showDetailReviewerNav}
        hideDocumentsNav={hideDetailDocumentsNav}
      />

      <div className={styles.content}>
        <div className={styles.titleSection}>
          <div className={styles.titleContent}>
            <div className={styles.titleMetaRow}>
              <div className={styles.fileTypeBadge}>
                <span>{(document.fileType || 'File').toUpperCase()}</span>
              </div>
              {!hideTopDocumentActions && (
                <div className={styles.titleActions}>
                  {!shouldHideDownloadActions && (
                    <button
                      type="button"
                      className={styles.titleDownloadButton}
                      disabled={isDownloadStarting}
                      onClick={() => { handleDownload(); }}
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                        <polyline points="7 10 12 15 17 10" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                        <line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                      {isDownloadStarting ? 'Downloading...' : 'Download'}
                    </button>
                  )}
                  {canEditOrUpdateDocument && (
                    <button
                      type="button"
                      className={styles.titleEditButton}
                      onClick={() => { void handleEdit(); }}
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
                        <path d="M12 20h9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                        <path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                      Edit
                    </button>
                  )}
                </div>
              )}
            </div>
            <div className={styles.titleTopRow}>
              <h1 className={styles.documentTitle}>{document.name}</h1>
            </div>

            <p className={styles.documentAbstract}>{document.abstract || ''}</p>

            <div className={styles.metadata}>
              <div className={styles.metadataItem}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <circle cx="12" cy="7" r="4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                <span>{document.author}</span>
              </div>
              <div className={styles.metadataItem}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <rect x="3" y="4" width="18" height="18" rx="2" ry="2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <line x1="16" y1="2" x2="16" y2="6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <line x1="8" y1="2" x2="8" y2="6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <line x1="3" y1="10" x2="21" y2="10" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                <span>{document.date}</span>
              </div>
            </div>
          </div>
        </div>

        <div className={styles.contentWrapper}>
          <div className={styles.mainContentWrapper}>
            <div className={styles.mainContent}>
              <div ref={previewSectionRef} className={`${styles.previewSection} ${isWindowMode ? styles.previewSectionWindowMode : ''}`}>
                <div className={styles.previewHeader}>
                  <div>
                    <h2 className={styles.previewTitle}>Document Preview</h2>
                    <p className={styles.previewSubtitle}>
                      {toolbarSummaryText}
                    </p>
                    {viewerNotice && <div className={styles.viewerNotice}>{viewerNotice}</div>}
                  </div>
                  <div className={styles.previewHeaderActions}>
                    {isMediaDocument && viewerCapabilities.canWindowMode && (
                      <button
                        type="button"
                        className={`${styles.viewerSecondaryButton} ${styles.viewerIconOnlyButton}`}
                        onClick={togglePreviewSectionFullscreen}
                        title={isExcelViewerFullscreen ? 'Exit fullscreen' : 'Fullscreen'}
                        aria-label={isExcelViewerFullscreen ? 'Exit fullscreen' : 'Fullscreen'}
                      >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
                          <polyline points="15 3 21 3 21 9" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                          <line x1="21" y1="3" x2="14" y2="10" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                          <polyline points="9 21 3 21 3 15" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                          <line x1="3" y1="21" x2="10" y2="14" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                        </svg>
                      </button>
                    )}
                    <button
                      type="button"
                      className={styles.viewerSecondaryButton}
                      onClick={() => { void handleVersionHistoryOpen(); }}
                    >
                      View version history
                    </button>
                    {!hideTopDocumentActions && canEditOrUpdateDocument && (
                      <button
                        type="button"
                      className={styles.viewerActionButton}
                      onClick={handleOpenUpdateUpload}
                    >
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
                          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                          <path d="M7 10l5-5 5 5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                          <path d="M12 5v12" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                        </svg>
                        Upload
                      </button>
                    )}
                    <button
                      type="button"
                      className={styles.viewerSecondaryButton}
                      onClick={() => setPreviewCollapsed((value) => !value)}
                    >
                      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">
                        <polyline
                          points={previewCollapsed ? "6 9 12 15 18 9" : "6 15 12 9 18 15"}
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      </svg>
                      {previewCollapsed ? 'Expand' : 'Collapse'}
                    </button>
                  </div>
                </div>

                {!previewCollapsed && (
                  <div
                    ref={isExcelStyleToolbarDocument ? excelPreviewShellRef : undefined}
                    className={`${styles.viewerChromeShell} ${isExcelStyleToolbarDocument ? styles.excelPdfViewerShell : ''} ${isMediaDocument ? styles.mediaViewerShell : ''}`}
                  >
                    {shouldRenderCustomViewerToolbar && (
                    isExcelStyleToolbarDocument ? (
                    <div className={styles.excelPdfToolbarShell}>
                      <div className={styles.excelPdfToolbarLayout}>
                        <CommandBar
                          className={styles.excelPdfToolbarCommandBar}
                          items={excelViewerLeftToolbarItems}
                          onReduceData={keepViewerToolbarItemsVisible}
                          ariaLabel="Excel viewer left controls"
                        />
                        <CommandBar
                          className={styles.excelPdfToolbarCommandBar}
                          items={excelViewerCenterToolbarItems}
                          onReduceData={keepViewerToolbarItemsVisible}
                          ariaLabel="Excel viewer zoom controls"
                        />
                        <CommandBar
                          className={styles.excelPdfToolbarCommandBar}
                          items={excelViewerRightToolbarItems}
                          onReduceData={keepViewerToolbarItemsVisible}
                          ariaLabel="Excel viewer right controls"
                        />
                      </div>
                    </div>
                    ) : (
                    <div
                      className={`${styles.viewerToolbar} ${isExcelStyleToolbarDocument ? styles.excelViewerToolbar : ''} ${isPdfLikeToolbarDocument ? styles.pdfLikeViewerToolbar : ''} ${isAudioDocument ? styles.audioViewerToolbar : ''}`}
                      role="toolbar"
                      aria-label="Document viewer controls"
                    >
                      {shouldShowViewerThumbnailToggle && viewerCapabilities.canToggleThumbnails && (
                        <button
                          type="button"
                          className={`${styles.viewerToolbarIconButton} ${pdfSidebarPanel === 'thumbnails' ? styles.viewerToolbarButtonActive : ''}`}
                          onClick={() => handleSidebarToggle('thumbnails')}
                          title="Toggle thumbnail panel"
                          aria-label="Toggle thumbnail panel"
                        >
                          <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <line x1="8" y1="6" x2="18" y2="6" strokeWidth="2" strokeLinecap="round" />
                            <line x1="8" y1="12" x2="18" y2="12" strokeWidth="2" strokeLinecap="round" />
                            <line x1="8" y1="18" x2="18" y2="18" strokeWidth="2" strokeLinecap="round" />
                            <circle cx="5" cy="6" r="1" fill="currentColor" stroke="none" />
                            <circle cx="5" cy="12" r="1" fill="currentColor" stroke="none" />
                            <circle cx="5" cy="18" r="1" fill="currentColor" stroke="none" />
                          </svg>
                        </button>
                      )}

                      {viewerCapabilities.canRotate && (
                        <>
                          <button
                            type="button"
                            className={styles.viewerToolbarIconButton}
                            onClick={() => rotatePdf('left')}
                            title="Rotate left"
                            aria-label="Rotate left"
                          >
                            <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                              <path d="M7.25 7.9A7.25 7.25 0 1 1 5.1 13.2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                              <polyline points="4.75 5.7 4.75 11.25 10.3 11.25" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </button>
                          <button
                            type="button"
                            className={styles.viewerToolbarIconButton}
                            onClick={() => rotatePdf('right')}
                            title="Rotate right"
                            aria-label="Rotate right"
                          >
                            <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                              <path d="M16.75 7.9A7.25 7.25 0 1 0 18.9 13.2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                              <polyline points="19.25 5.7 19.25 11.25 13.7 11.25" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </button>
                        </>
                      )}

                      {(viewerCapabilities.canToggleThumbnails || viewerCapabilities.canRotate) && viewerCapabilities.canZoom && (
                        <div className={styles.toolbarDivider} />
                      )}

                      {viewerCapabilities.canZoom && (
                        <>
                        <button
                          type="button"
                          className={`${styles.viewerToolbarIconButton} ${styles.viewerToolbarZoomButton}`}
                          onClick={() => updateZoom('out')}
                          title="Zoom out"
                          aria-label="Zoom out"
                        >
                          {usesLineIconToolbar ? (
                            <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                              <circle cx="11" cy="11" r="7" strokeWidth="2" />
                              <line x1="16.5" y1="16.5" x2="21" y2="21" strokeWidth="2" strokeLinecap="round" />
                              <line x1="8" y1="11" x2="14" y2="11" strokeWidth="2" strokeLinecap="round" />
                            </svg>
                          ) : '-'}
                        </button>
                        <div className={styles.viewerMenuWrapper} ref={zoomMenuRef}>
                        <div
                          className={`${styles.viewerToolbarZoomValue} ${isZoomMenuOpen ? styles.viewerToolbarButtonActive : ''}`}
                          title="Edit or select zoom"
                        >
                          <input
                            className={styles.toolbarZoomInput}
                            value={isEditingZoom ? zoomInputValue : String(activeZoomPercent)}
                            inputMode="numeric"
                            aria-label="Zoom percentage"
                            onFocus={(event) => {
                              setIsEditingZoom(true);
                              setZoomInputValue(String(activeZoomPercent));
                              event.currentTarget.select();
                            }}
                            onChange={(event) => setZoomInputValue(event.currentTarget.value)}
                            onBlur={commitZoomInput}
                            onKeyDown={(event) => {
                              if (event.key === 'Enter') {
                                event.currentTarget.blur();
                              } else if (event.key === 'Escape') {
                                setZoomInputValue(String(activeZoomPercent));
                                setIsEditingZoom(false);
                                event.currentTarget.blur();
                              }
                            }}
                          />
                          <span aria-hidden="true">%</span>
                          <button
                            type="button"
                            className={styles.viewerToolbarZoomMenuButton}
                            onClick={() => setIsZoomMenuOpen((value) => !value)}
                            title="Select zoom"
                            aria-label={`Select zoom. Current zoom ${activeZoomPercent}%`}
                            aria-expanded={isZoomMenuOpen}
                          >
                            <svg width="14" height="14" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                              <polyline points="6 9 12 15 18 9" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </button>
                        </div>
                        {isZoomMenuOpen && (
                          <div className={`${styles.viewerMenu} ${styles.viewerZoomMenu}`} role="menu" aria-label="Zoom levels">
                            {ZOOM_PRESETS.map((preset) => (
                              <button
                                key={preset}
                                type="button"
                                className={`${styles.viewerMenuItem} ${activeZoomPercent === preset ? styles.viewerMenuItemActive : ''}`}
                                onClick={() => applyZoomPreset(preset)}
                                role="menuitem"
                              >
                                {preset}%
                              </button>
                            ))}
                          </div>
                        )}
                        </div>
                        <button
                          type="button"
                          className={`${styles.viewerToolbarIconButton} ${styles.viewerToolbarZoomButton}`}
                          onClick={() => updateZoom('in')}
                          title="Zoom in"
                          aria-label="Zoom in"
                        >
                          {usesLineIconToolbar ? (
                            <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
                              <circle cx="11" cy="11" r="7" strokeWidth="2" />
                              <line x1="16.5" y1="16.5" x2="21" y2="21" strokeWidth="2" strokeLinecap="round" />
                              <line x1="8" y1="11" x2="14" y2="11" strokeWidth="2" strokeLinecap="round" />
                              <line x1="11" y1="8" x2="11" y2="14" strokeWidth="2" strokeLinecap="round" />
                            </svg>
                          ) : '+'}
                        </button>
                        {!isPresentationDocument && !isSimpleOfficeStyleToolbarDocument && !isAudioDocument && !isVideoDocument && (
                          <button
                            type="button"
                            className={`${styles.viewerToolbarIconButton} ${isFitToContainerMode ? styles.viewerToolbarButtonActive : ''}`}
                            onClick={() => updateZoom('reset')}
                            title="Fit to page"
                            aria-label="Fit to page"
                          >
                            <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                              <rect x="5" y="4" width="14" height="16" rx="1.5" strokeWidth="2" />
                              <path d="M9 8H7V6" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                              <path d="M15 8h2V6" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                              <path d="M9 16H7v2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                              <path d="M15 16h2v2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </button>
                        )}
                        </>
                      )}

                      {viewerCapabilities.canZoom && viewerCapabilities.canNavigatePages && (
                        <div className={styles.toolbarDivider} />
                      )}

                      {viewerCapabilities.canNavigatePages && (
                        <>
                          <button
                            type="button"
                            className={`${styles.viewerToolbarIconButton} ${styles.viewerToolbarPageButton}`}
                            onClick={() => changeViewerPage('prev')}
                            title={`Previous ${viewerCapabilities.pageUnitLabel}`}
                            aria-label={`Previous ${viewerCapabilities.pageUnitLabel}`}
                            disabled={!supportsPrecisePageNavigation}
                            aria-disabled={!supportsPrecisePageNavigation}
                          >
                            <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                              <polyline points="15 18 9 12 15 6" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </button>
                          <div className={styles.viewerToolbarPageIndicator}>
                            {supportsPrecisePageNavigation ? `${activePageNumber} of ${activePageTotal}` : pageIndicatorText}
                          </div>
                          <button
                            type="button"
                            className={`${styles.viewerToolbarIconButton} ${styles.viewerToolbarPageButton}`}
                            onClick={() => changeViewerPage('next')}
                            title={`Next ${viewerCapabilities.pageUnitLabel}`}
                            aria-label={`Next ${viewerCapabilities.pageUnitLabel}`}
                            disabled={!supportsPrecisePageNavigation}
                            aria-disabled={!supportsPrecisePageNavigation}
                          >
                            <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                              <polyline points="9 18 15 12 9 6" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </button>
                        </>
                      )}

                      {(viewerCapabilities.canNavigatePages || viewerCapabilities.canZoom || viewerCapabilities.canRotate || viewerCapabilities.canToggleThumbnails) &&
                        (viewerCapabilities.canWindowMode || viewerCapabilities.canChangeLayout || (canPrintDocument && !isPrintDisabled)) && (
                          <>
                            {isExcelStyleToolbarDocument && <div className={styles.excelToolbarSpacer} aria-hidden="true" />}
                            <div className={styles.toolbarDivider} />
                          </>
                        )}

                      {viewerCapabilities.canWindowMode && (
                        <button
                          type="button"
                          className={`${styles.viewerToolbarIconButton} ${isWindowMode ? styles.viewerToolbarButtonActive : ''}`}
                          onClick={() => setIsWindowMode((value) => !value)}
                          title={isWindowMode ? 'Exit window view' : 'Expand into window view'}
                          aria-label={isWindowMode ? 'Exit window view' : 'Expand into window view'}
                        >
                          <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <polyline points="15 3 21 3 21 9" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            <line x1="21" y1="3" x2="14" y2="10" strokeWidth="2" strokeLinecap="round" />
                            <polyline points="9 21 3 21 3 15" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            <line x1="3" y1="21" x2="10" y2="14" strokeWidth="2" strokeLinecap="round" />
                          </svg>
                        </button>
                      )}
                      {viewerCapabilities.canChangeLayout && (
                        <button
                          type="button"
                          className={`${styles.viewerToolbarIconButton} ${pdfScrollMode !== 'single' ? styles.viewerToolbarButtonActive : ''}`}
                          onClick={toggleViewerLayout}
                          title={pdfScrollMode === 'single' ? 'Switch to continuous page view' : 'Switch to single page view'}
                          aria-label={pdfScrollMode === 'single' ? 'Switch to continuous page view' : 'Switch to single page view'}
                        >
                          <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <rect x="5" y="3" width="10" height="8" rx="1.2" strokeWidth="2" />
                            <rect x="9" y="13" width="10" height="8" rx="1.2" strokeWidth="2" />
                            <path d="M17 5h2v8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            <path d="M7 19H5v-8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                        </button>
                      )}
                      {canPrintDocument && !isPrintDisabled && !isSimpleOfficeStyleToolbarDocument && !isAudioDocument && !isVideoDocument && (
                        <button
                          type="button"
                          className={styles.viewerToolbarIconButton}
                          onClick={handlePrint}
                          title={printButtonTitle}
                          aria-label={printButtonTitle}
                          disabled={isPrintInProgress}
                        >
                          <svg width="18" height="18" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                            <polyline points="6 9 6 3 18 3 18 9" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            <path d="M6 17H4a2 2 0 0 1-2-2v-4a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v4a2 2 0 0 1-2 2h-2" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            <rect x="6" y="14" width="12" height="7" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                        </button>
                      )}
                    </div>
                    )
                    )}

                    <div className={styles.previewWorkspace}>
                      {!shouldUsePdfPreviewViewer && !isImageDocument && pdfSidebarPanel && (isPdfCustomViewer || supportsThumbnailSidebar) && (
                        <div className={styles.pdfSidebarPanel}>
                          <div className={styles.pdfSidebarPanelHeader}>
                            <span className={styles.pdfSidebarPanelTitle}>
                              {pdfSidebarPanel === 'thumbnails' ? 'Pages' : pdfSidebarPanel === 'bookmarks' ? 'Bookmarks' : 'Attachments'}
                            </span>
                            <button type="button" className={styles.pdfSidebarClose} onClick={() => setPdfSidebarPanel(null)}>
                              <svg width="14" height="14" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                                <line x1="6" y1="6" x2="18" y2="18" strokeWidth="2" strokeLinecap="round" />
                                <line x1="18" y1="6" x2="6" y2="18" strokeWidth="2" strokeLinecap="round" />
                              </svg>
                            </button>
                          </div>
                          {pdfSidebarPanel === 'thumbnails' ? (
                            isPdfCustomViewer || isPdfDocument ? (
                              pdfThumbnailError ? (
                                <PdfImageFallbackThumbnailStrip
                                  pageCount={activePageTotal}
                                  activePageNumber={activePageNumber}
                                  onSelectPage={changePdfPage}
                                  getThumbnailUrl={getSidebarThumbnailUrl}
                                  getLoadingMode={getThumbnailLoadingMode}
                                  onThumbnailError={handleThumbnailLoadError}
                                  hasMoreCandidates={hasMoreThumbnailCandidates}
                                />
                              ) : !pdfThumbnailDocument ? (
                                <PdfImageFallbackThumbnailStrip
                                  pageCount={activePageTotal}
                                  activePageNumber={activePageNumber}
                                  onSelectPage={changePdfPage}
                                  getThumbnailUrl={getSidebarThumbnailUrl}
                                  getLoadingMode={getThumbnailLoadingMode}
                                  onThumbnailError={handleThumbnailLoadError}
                                  hasMoreCandidates={hasMoreThumbnailCandidates}
                                />
                              ) : (
                                <PdfCanvasThumbnailStrip
                                  pdfDocument={pdfThumbnailDocument}
                                  pageCount={activePageTotal}
                                  activePageNumber={activePageNumber}
                                  onSelectPage={changePdfPage}
                                />
                              )
                            ) : (
                              <div className={styles.pdfThumbnailStrip}>
                                {Array.from({ length: activePageTotal }, (_, index) => index + 1).map((pageNumber) => {
                                  const thumbnailUrl = getSidebarThumbnailUrl(pageNumber);

                                  return (
                                    <button
                                      key={`thumb-${pageNumber}`}
                                      type="button"
                                      className={`${styles.pdfThumbnailButton} ${isPresentationDocument ? styles.pdfThumbnailButtonPresentation : ''} ${activePageNumber === pageNumber ? styles.pdfThumbnailButtonActive : ''}`}
                                      onClick={() => {
                                        if (isPdfDocument) {
                                          changePdfPage(pageNumber);
                                        } else if (isOfficeDocument) {
                                          changeEmbeddedPage(pageNumber, isPresentationDocument ? thumbnailUrl || undefined : undefined);
                                        }
                                      }}
                                    >
                                      <span className={styles.pdfThumbnailIndex}>{pageNumber}</span>
                                      <span className={`${styles.pdfThumbnailCardBody} ${isPresentationDocument ? styles.pdfThumbnailCardBodyPresentation : ''}`}>
                                        <span
                                          className={`${styles.pdfThumbnailPreview} ${isPresentationDocument ? styles.pdfThumbnailPreviewOfficePresentation : ''
                                            }`}
                                          aria-hidden="true"
                                        >
                                          {(isPresentationDocument || isWordProcessingDocument) && thumbnailUrl && thumbnailUrl.indexOf('/_layouts/15/WopiFrame.aspx') !== -1 ? (
                                            <iframe
                                              src={thumbnailUrl}
                                              title={`${isPresentationDocument ? 'Slide' : 'Page'} ${pageNumber} preview`}
                                              className={styles.pdfThumbnailFrame}
                                              loading={getThumbnailLoadingMode(pageNumber)}
                                              scrolling="no"
                                              tabIndex={-1}
                                              aria-hidden="true"
                                            />
                                          ) : thumbnailUrl ? (
                                            <img
                                              src={thumbnailUrl}
                                              alt=""
                                              className={styles.pdfThumbnailImage}
                                              loading={getThumbnailLoadingMode(pageNumber)}
                                              decoding="async"
                                              onError={() => handleThumbnailLoadError(pageNumber)}
                                            />
                                          ) : hasMoreThumbnailCandidates(pageNumber) ? (
                                            <span className={styles.pdfThumbnailPreviewInner} />
                                          ) : (
                                            <span className={styles.pdfSidebarEmpty}>Preview unavailable</span>
                                          )}
                                        </span>
                                        {!isPresentationDocument && (
                                          <span className={styles.pdfThumbnailLabel}>
                                            {isPresentationDocument
                                              ? `Slide ${pageNumber}`
                                              : isSpreadsheetDocument
                                                ? `Sheet ${pageNumber}`
                                                : `Page ${pageNumber}`}
                                          </span>
                                        )}
                                      </span>
                                    </button>
                                  );
                                })}
                              </div>
                            )
                          ) : (
                            <div className={styles.pdfSidebarEmpty}>
                              {pdfSidebarPanel === 'bookmarks' ? 'No bookmarks available for this PDF.' : 'No attachments available for this PDF.'}
                            </div>
                          )}
                        </div>
                      )}

                      <div className={styles.previewStage}>
                        <div
                          className={styles.previewContainer}
                          ref={previewContainerRef}
                          style={previewContainerStyle}
                          onWheel={handlePreviewWheel}
                        >
                          {isPreviewLoading ? (
                            <div className={styles.previewPlaceholder} role="status" aria-live="polite">
                              <div className={styles.previewUnavailableTitle}>Preview loading</div>
                            </div>
                          ) : previewUrl ? (
                            shouldUsePdfPreviewViewer ? (
                              <div className={styles.pdfPreviewFrameShell}>
                                <React.Suspense fallback={<div className={styles.previewPlaceholder}>Loading PDF viewer...</div>}>
                                  <PdfViewer
                                    fileUrl={previewUrl}
                                    fileName={document.name}
                                    canPrint={canPrintDocument}
                                    fitToPageOnLoad={isConvertedOfficePdfPreview && isPresentationDocument}
                                  />
                                </React.Suspense>
                              </div>
                            ) : isImageDocument ? (
                              <div className={styles.pdfPreviewFrameShell}>
                                <React.Suspense fallback={<div className={styles.previewPlaceholder}>Loading image viewer...</div>}>
                                  <ImageViewer
                                    fileUrl={previewUrl}
                                    fileName={document.name}
                                  />
                                </React.Suspense>
                              </div>
                            ) : (
                              <div className={styles.previewZoomStage} style={embeddedPreviewStageStyle}>
                                <div style={embeddedRotationWrapperStyle}>
                              {normalizedFileType === 'svg' ? (
                                <object
                                  data={previewUrl}
                                  type="image/svg+xml"
                                  className={styles.previewFrame}
                                  style={embeddedPreviewStyle}
                                  title="Document Preview"
                                  onError={(e) => {
                                    console.error('SVG load error:', e);
                                    setPreviewError('Failed to load SVG preview');
                                  }}
                                  onLoad={() => setPreviewError('')}
                                >
                                  <img
                                    src={previewUrl}
                                    alt="Document Preview"
                                    className={styles.previewFrame}
                                    style={{ objectFit: 'contain', maxWidth: '100%', maxHeight: '100%', width: '100%', height: '100%' }}
                                  />
                                </object>
                              ) : isTextDocument ? (
                                <div className={styles.previewTextContainer} style={embeddedPreviewStyle}>
                                  <pre className={styles.previewTextContent} style={textPreviewContentStyle}>{textContent}</pre>
                                </div>
                              ) : isVideoDocument ? (
                                <video
                                  src={previewUrl}
                                  className={`${styles.previewMedia} ${styles.previewVideoMedia}`}
                                  style={embeddedPreviewStyle}
                                  controls
                                  controlsList="nodownload"
                                  preload="metadata"
                                  onError={() => setPreviewError('Failed to load video preview')}
                                  onLoadedData={() => setPreviewError('')}
                                />
                              ) : isAudioDocument ? (
                                <div className={styles.previewAudioShell} style={embeddedPreviewStyle}>
                                  <div style={{ textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '24px', width: '100%', maxWidth: '500px' }}>
                                    <div style={{ width: '100px', height: '100px', background: '#dce5f5', borderRadius: '24px', display: 'flex', alignItems: 'center', justifyContent: 'center', boxShadow: '0 8px 16px rgba(37, 99, 235, 0.15)' }}>
                                      <svg width="50" height="50" viewBox="0 0 24 24" fill="none" stroke="#2f6fdb" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                                        <path d="M9 18V5l12-2v13"></path>
                                        <circle cx="6" cy="18" r="3"></circle>
                                        <circle cx="18" cy="16" r="3"></circle>
                                      </svg>
                                    </div>
                                    <div style={{ fontSize: '18px', fontWeight: 600, color: '#1e293b', wordBreak: 'break-all', fontFamily: 'Inter, sans-serif' }}>
                                      {document.name}
                                    </div>
                                    <audio
                                      src={previewUrl}
                                      className={styles.previewAudio}
                                      controls
                                      controlsList="nodownload"
                                      preload="metadata"
                                      onError={() => setPreviewError('Failed to load audio preview')}
                                      onLoadedData={() => setPreviewError('')}
                                      style={{ width: '100%', outline: 'none' }}
                                    />
                                  </div>
                                </div>
                              ) : isImageDocument || (isPresentationDocument && previewUrl.indexOf('/_layouts/15/WopiFrame.aspx') === -1) ? (
                                <img
                                  src={previewUrl}
                                  alt={`${document.name} preview`}
                                  className={`${styles.previewImage} ${isPresentationDocument ? styles.previewPresentationImage : ''}`}
                                  style={embeddedPreviewStyle}
                                  onError={() => {
                                    if (isPresentationDocument) {
                                      handlePresentationImageError();
                                      return;
                                    }

                                    setPreviewError('Failed to load image preview');
                                  }}
                                  onLoad={() => setPreviewError('')}
                                />
                              ) : isOfficeDocument ? (
                                <iframe
                                  ref={iframeRef}
                                  src={previewUrl}
                                  className={`${styles.previewFrame} ${styles.previewFrameOffice} ${isPresentationDocument
                                    ? styles.previewFramePresentation
                                    : isSpreadsheetDocument
                                      ? styles.previewFrameSpreadsheet
                                      : ''
                                    }`}
                                  style={embeddedPreviewStyle}
                                  title="Document Preview"
                                  onError={() => setPreviewError('Failed to load Office preview')}
                                  onLoad={(event) => {
                                    setPreviewError('');
                                    syncOfficeViewerPageState(event.currentTarget);
                                    resetSpreadsheetViewerPosition(event.currentTarget);
                                    attachOfficeIframeWheelSync(event.currentTarget);
                                    if (isSpreadsheetDocument && usesOfficeIframeZoomSync) {
                                      window.requestAnimationFrame(() => applyExcelIframeZoom(officeToolbarZoom));
                                    }
                                  }}
                                />
                              ) : (
                                <iframe
                                  ref={iframeRef}
                                  src={previewUrl}
                                  className={styles.previewFrame}
                                  style={embeddedPreviewStyle}
                                  title="Document Preview"
                                    onError={() => setPreviewError('Failed to load preview')}
                                  onLoad={() => setPreviewError('')}
                                />
                              )}
                                </div>
                              </div>
                            )
                          ) : (
                            <div className={styles.previewPlaceholder} role="status" aria-live="polite">
                              <div className={styles.previewUnavailableIcon} aria-hidden="true">
                                <svg width="34" height="34" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                                  <path d="M14 2H7a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V7l-5-5Z" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
                                  <path d="M14 2v5h5" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
                                  <path d="M9 14h6" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
                                  <path d="M9 17h4" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" />
                                </svg>
                              </div>
                              <div className={styles.previewUnavailableTitle}>Preview not available</div>
                              <div className={styles.previewUnavailableText}>
                                This file format cannot be previewed in the browser.
                              </div>
                              {document.fileType && (
                                <div className={styles.previewUnavailableBadge}>{document.fileType}</div>
                              )}
                            </div>
                          )}
                          {previewError && (
                            <div className={styles.previewError}>
                              <div className={styles.errorIcon}>⚠️</div>
                              <div className={styles.errorText}>{previewError}</div>
                              <button
                                className={styles.retryButton}
                                onClick={() => {
                                  setPreviewError('');
                                  fetchDocumentDetails();
                                }}
                              >
                                Retry
                              </button>
                            </div>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>

            <div className={styles.sidebar}>
              <div className={styles.panel}>
                <div className={styles.panelHeader}>
                  <div className={styles.panelIcon}>
                    <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                      <rect x="2" y="1" width="10" height="13" rx="1.5" stroke="currentColor" strokeWidth="1.2" />
                      <path d="M5 5h6M5 8h6M5 11h4" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
                    </svg>
                  </div>
                  <h3 className={styles.panelTitle}>Document information</h3>
                </div>
                <div className={styles.buSection}>
                  <div className={styles.sectionLabel}>Business Unit &amp; Department</div>
                  <BUDeptDisplay
                    spHttpClient={props.context.spHttpClient}
                    siteUrl={props.context.pageContext.web.absoluteUrl}
                    itemId={props.documentId}
                  />
                </div>
                <div className={styles.metaGrid}>
                  {documentInfoFields.map((field) => {
                    const hasRenderableValue = !!field.value && field.value !== '-';
                    const displayValue = hasRenderableValue ? field.value : '\u2014';
                    const fieldValues = field.supportsMultipleValues && hasRenderableValue
                      ? splitDocumentInfoValues(field.value)
                      : [];

                    return (
                      <div className={styles.metaCell} key={field.label}>
                        <span className={styles.metaLabel}>{field.label}</span>
                        {fieldValues.length > 0 ? (
                          <span className={styles.metaValueList}>
                            {fieldValues.map((fieldValue) => (
                              <span className={styles.metaValueItem} key={fieldValue}>{fieldValue}</span>
                            ))}
                          </span>
                        ) : (
                          <span className={`${styles.metaValue} ${!hasRenderableValue ? styles.metaValueEmpty : ''}`}>
                            {fieldValues[0] || displayValue}
                          </span>
                        )}
                      </div>
                    );
                  })}
                  <div className={styles.metaCell}>
                    <span className={styles.metaLabel}>Status</span>
                    <span className={getStatusBadgeClassName(documentStatusValue)}>
                      {documentStatusValue && documentStatusValue !== '-' ? documentStatusValue : '\u2014'}
                    </span>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className={styles.socialSection}>
            {!disableSocialActions && (
              <>
                <div className={styles.socialActions}>
                  <button
                    className={`${styles.actionPill} ${styles.likeButton} ${!disableSocialActions && isLiked ? styles.actionPillActive : ''} ${!disableSocialActions && isLiked ? styles.liked : ''}`}
                    onClick={toggleLike}
                    aria-label="Like"
                    disabled={disableSocialActions || Boolean(document && pendingSocialActions[`${document.id}_Like`])}
                  >
                    <svg width="24" height="24" viewBox="0 0 24 24" fill={isLiked ? "url(#likeGradient)" : "none"} xmlns="http://www.w3.org/2000/svg">
                      <defs>
                        <linearGradient id="likeGradient" x1="0%" y1="0%" x2="100%" y2="0%">
                          <stop offset="0%" stopColor="#ff6b9d" />
                          <stop offset="30%" stopColor="#c44569" />
                          <stop offset="70%" stopColor="#6c5ce7" />
                          <stop offset="100%" stopColor="#4834d4" />
                        </linearGradient>
                      </defs>
                      <path
                        d="M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z"
                        fill={isLiked ? "url(#likeGradient)" : "none"}
                        stroke={isLiked ? "url(#likeGradient)" : "currentColor"}
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                    <span className={styles.actionLabel}>Like</span>
                    <span className={styles.count}>{disableSocialActions ? 0 : likeCount}</span>
                  </button>
                  <div className={`${styles.actionPill} ${disableSocialActions ? styles.actionPillDisabled : ''}`} aria-label="Views">
                    <svg width="22" height="22" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M1 12s4-7 11-7 11 7 11 7-4 7-11 7S1 12 1 12z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                      <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    <span className={styles.actionLabel}>Views</span>
                    <span className={styles.count}>{disableSocialActions ? 0 : viewCount}</span>
                  </div>
                  {!shouldHideDownloadActions && (
                    <div className={`${styles.actionPill} ${disableSocialActions ? styles.actionPillDisabled : ''}`} aria-label="Downloads">
                      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                        <polyline points="7 10 12 15 17 10" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                        <line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                      </svg>
                      <span className={styles.actionLabel}>Download</span>
                      <span className={styles.count}>{disableSocialActions ? 0 : downloadCount}</span>
                    </div>
                  )}
                  <button
                    className={`${styles.actionPill} ${!disableSocialActions && showComments ? styles.actionPillActive : ''}`}
                    ref={commentButtonRef}
                    onClick={handleCommentButtonClick}
                    aria-label="Comment"
                    disabled={disableSocialActions}
                  >
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path
                        d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"
                        stroke="currentColor"
                        strokeWidth="2"
                        strokeLinecap="round"
                        strokeLinejoin="round"
                      />
                    </svg>
                    <span className={styles.actionLabel}>Comment</span>
                    <span className={styles.count}>{disableSocialActions ? 0 : commentCount}</span>
                  </button>
                  <button
                    className={`${styles.actionPill} ${!disableSocialActions && isFlaggedForReview ? styles.actionPillWarning : ''}`}
                    onClick={handleFlagReviewAction}
                    aria-label="Flag for review"
                    disabled={disableSocialActions || isSubmittingFlagReview}
                  >
                    <svg width="22" height="22" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <path d="M5 21V5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                      <path d="M5 5h10l-1 4 1 4H5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    <span className={styles.actionLabel}>{!disableSocialActions && isFlaggedForReview ? 'Flagged' : 'Flag for review'}</span>
                  </button>
                  <button
                    className={`${styles.actionPill} ${isBookmarked ? styles.actionPillActive : ''}`}
                    onClick={() => toggleDocumentAction('bookmark')}
                    aria-label="Bookmark"
                    disabled={disableSocialActions || Boolean(document && pendingSocialActions[`${document.id}_Bookmark`])}
                  >
                    <svg width="22" height="22" viewBox="0 0 24 24" fill={isBookmarked ? "currentColor" : "none"} xmlns="http://www.w3.org/2000/svg">
                      <path d="M6 3h12a1 1 0 0 1 1 1v17l-7-4-7 4V4a1 1 0 0 1 1-1z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    <span className={styles.actionLabel}>{isBookmarked ? 'Bookmarked' : 'Bookmark'}</span>
                    <span className={styles.count}>{bookmarkCount}</span>
                  </button>
                  {SHOW_DOCUMENT_FOLLOW_ACTION && (
                    <button
                      className={`${styles.actionPill} ${isFollowing ? styles.actionPillActive : ''}`}
                      onClick={() => { void toggleDocumentAction('follow'); }}
                      aria-label="Follow"
                      disabled={disableSocialActions || Boolean(document && pendingSocialActions[`${document.id}_Follow`])}
                    >
                      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path
                          d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          fill={isFollowing ? 'currentColor' : 'none'}
                        />
                        <path
                          d="M13.73 21a2 2 0 0 1-3.46 0"
                          stroke="currentColor"
                          strokeWidth="2"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      </svg>
                      <span className={styles.actionLabel}>
                        {isFollowing ? 'Following' : 'Follow'}
                      </span>
                      <span className={styles.count}>
                        {disableSocialActions ? 0 : followCount}
                      </span>
                    </button>
                  )}
                  <button className={`${styles.actionPill} ${styles.shareButton}`} onClick={handleShare} aria-label="Share" disabled={disableSocialActions}>
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                      <line x1="22" y1="2" x2="11" y2="13" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                      <polygon points="22 2 15 22 11 13 2 9 22 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                    <span className={styles.actionLabel}>Share</span>
                    <span className={styles.count}>{disableSocialActions ? 0 : shareCount}</span>
                  </button>
                </div>
                {isFlagReviewModalOpen && (
                  <div className={styles.flagReviewModalOverlay} onClick={closeFlagReviewModal}>
                    <div
                      className={styles.flagReviewModal}
                      role="dialog"
                      aria-modal="true"
                      aria-labelledby="flag-review-title"
                      onClick={(event) => event.stopPropagation()}
                    >
                      <div className={styles.flagReviewModalHeader}>
                        <div>
                          <h3 id="flag-review-title" className={styles.flagReviewModalTitle}>Flag for Review</h3>
                          <p className={styles.flagReviewModalSubtitle}>
                            Provide the reason why this document needs a revision or is not appropriate.
                          </p>
                        </div>
                        <button
                          type="button"
                          className={styles.flagReviewModalClose}
                          onClick={closeFlagReviewModal}
                          aria-label="Close flag for review dialog"
                          disabled={isSubmittingFlagReview}
                        >
                          ×
                        </button>
                      </div>

                      <div className={styles.flagReviewForm}>
                        <label className={styles.flagReviewField}>
                          <span className={styles.flagReviewLabel}>Reason *</span>
                          <textarea
                            className={styles.flagReviewTextarea}
                            value={flagReviewDetails}
                            onChange={(event) => setFlagReviewDetails(event.target.value)}
                            placeholder="Type the reason"
                            rows={5}
                            maxLength={FLAG_REVIEW_MAX_LENGTH}
                            disabled={isSubmittingFlagReview}
                          />
                          <div className={styles.flagReviewCharacterCount}>
                            {flagReviewDetails.length}/{FLAG_REVIEW_MAX_LENGTH}
                          </div>
                        </label>
                        {flagReviewError && (
                          <div className={styles.flagReviewError}>{flagReviewError}</div>
                        )}
                      </div>

                      <div className={styles.flagReviewModalActions}>
                        <button
                          type="button"
                          className={styles.flagReviewSecondaryButton}
                          onClick={closeFlagReviewModal}
                          disabled={isSubmittingFlagReview}
                        >
                          Cancel
                        </button>
                        <button
                          type="button"
                          className={styles.flagReviewPrimaryButton}
                          onClick={() => void submitFlagReview()}
                          disabled={
                            isSubmittingFlagReview ||
                            !flagReviewDetails.trim()
                          }
                        >
                          {isSubmittingFlagReview ? 'Sending...' : 'Send'}
                        </button>
                      </div>
                    </div>
                  </div>
                )}
                {actionMessage && <div className={styles.actionMessage}>{actionMessage}</div>}

                {showComments && (
                  <div className={styles.commentsBox} ref={commentsPanelRef}>
                    <div className={styles.commentsList}>
                      {comments.length > 0 ? (
                        comments.map((comment) => (
                          <div key={comment.id} className={styles.commentItem}>
                            <div className={styles.commentHeader}>
                              <div className={styles.commentHeaderLeft}>
                                <span className={styles.commentUserName}>{comment.userName}</span>
                                <span className={styles.commentTime}>{formatTimeAgo(comment.created)}</span>
                              </div>
                              {comment.userId === currentUserId && (
                                <button
                                  className={styles.deleteCommentButton}
                                  onClick={() => deleteComment(comment.id)}
                                  aria-label="Delete comment"
                                >
                                  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                                    <polyline points="3 6 5 6 21 6" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                                    <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                                  </svg>
                                </button>
                              )}
                            </div>
                            <p className={styles.commentText}>{comment.comment}</p>
                          </div>
                        ))
                      ) : (
                        <div className={styles.noComments}>No comments yet</div>
                      )}
                    </div>
                    <div className={styles.addCommentSection}>
                      <div className={styles.commentInputWrapper}>
                        <input
                          type="text"
                          className={styles.commentInput}
                          placeholder="Add a comment..."
                          value={newComment}
                          maxLength={500}
                          onChange={(e) => setNewComment(e.target.value)}
                          onKeyPress={(e) => {
                            if (e.key === 'Enter' && newComment.trim() && newComment.length <= 500) {
                              addComment();
                            }
                          }}
                        />
                        <span className={`${styles.charCount} ${newComment.length >= 450 ? styles.charCountWarning : ''} ${newComment.length >= 500 ? styles.charCountError : ''}`}>
                          {newComment.length}/500
                        </span>
                      </div>
                      <button
                        className={styles.sendButton}
                        onClick={addComment}
                        disabled={!newComment.trim() || newComment.length > 500}
                        aria-label="Send comment"
                      >
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                          <line x1="22" y1="2" x2="11" y2="13" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                          <polygon points="22 2 15 22 11 13 2 9 22 2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                        </svg>
                      </button>
                    </div>
                  </div>
                )}
              </>
            )}
          </div>
        </div>
      </div>
      <IKShellFooter compact={true} onBackHome={handleDetailFooterBackHome} />
      {editDialogOpen && (
        <div className={styles.editOverlay} onClick={handleCloseEdit}>
          <div className={styles.editShell} onClick={(event) => event.stopPropagation()}>
            <IKShellHeader
              userName={detailUserName}
              userEmail={detailUserEmail}
              userPhotoUrl={props.userPhotoUrl}
              compact={true}
              onLogout={props.onLogout}
              onHomeOpen={props.onHomeOpen}
              onAllDocumentsOpen={props.onAllDocumentsOpen}
              onBusinessUnitsOpen={props.onBusinessUnitsOpen}
              onBookmarksOpen={props.onBookmarksOpen}
              onDocumentsOpen={props.onDocumentsOpen}
              onContactOpen={props.onContactOpen}
              onAuditLogOpen={props.onAuditLogOpen}
              onAnalyticsOpen={props.onAnalyticsOpen}
              showReviewerNav={showDetailReviewerNav}
              hideDocumentsNav={hideDetailDocumentsNav}
            />
            <div className={styles.editMetadataBody} ref={editMetadataBodyRef}>
              <div className={styles.editModal}>
                {editInitialValues ? (
                  <>
                    {editError && <div className={styles.editError}>{editError}</div>}
                    <MetadataForm
                      context={props.context}
                      taxonomyOptions={taxonomyOptions || EMPTY_TAXONOMY_OPTIONS}
                      initialValues={editInitialValues}
                      showReviewerFields={isKmAdmin || isApprover}
                      density="compact"
                      onClose={handleCloseEdit}
                      onSubmit={(data) => { void handleEditSubmit(data); }}
                    />
                    {editSaving && (
                      <div className={styles.editSaving} role="status" aria-live="polite">
                        <div className={styles.editSavingCard}>
                          <span className={styles.editSavingSpinner} aria-hidden="true" />
                          <div>
                            <div className={styles.editSavingTitle}>Saving changes</div>
                            <div className={styles.editSavingText}>Updating the document metadata...</div>
                          </div>
                        </div>
                      </div>
                    )}
                  </>
                ) : editLoading ? (
                  <div className={styles.editSaving} role="status" aria-live="polite">
                    <div className={styles.editSavingCard}>
                      <span className={styles.editSavingSpinner} aria-hidden="true" />
                      <div>
                        <div className={styles.editSavingTitle}>Loading metadata</div>
                        <div className={styles.editSavingText}>Loading the latest document metadata...</div>
                      </div>
                    </div>
                  </div>
                ) : null}
                {!editLoading && editError && !editInitialValues && (
                  <div className={styles.editActions}>
                    <button className={styles.closeEditButton} onClick={handleCloseEdit}>
                      Close
                    </button>
                  </div>
                )}
              </div>
            </div>
            <IKShellFooter compact={true} onBackHome={handleDetailFooterBackHome} />
          </div>
        </div>
      )}
      {isUpdateUploadOpen && (
        <FileUpload
          context={props.context}
          onClose={handleCloseUpdateUpload}
          onUploaded={handleUpdateUploadComplete}
          targetItemId={document?.id}
          targetFileRef={document?.fileRef || document?.serverRelativeUrl}
          variant="replace"
          flowDensity="compact"
          userName={detailUserName}
          userEmail={detailUserEmail}
          userPhotoUrl={props.userPhotoUrl}
          onLogout={props.onLogout}
          onAllDocumentsOpen={props.onAllDocumentsOpen}
          onBusinessUnitsOpen={props.onBusinessUnitsOpen}
          onBookmarksOpen={props.onBookmarksOpen}
          onDocumentsOpen={props.onDocumentsOpen}
          onContactOpen={props.onContactOpen}
          onAuditLogOpen={props.onAuditLogOpen}
          onAnalyticsOpen={props.onAnalyticsOpen}
          showReviewerNav={showDetailReviewerNav}
          hideDocumentsNav={hideDetailDocumentsNav}
        />
      )}
      {isVersionHistoryOpen && (
        <div className={styles.versionHistoryOverlay} onClick={handleVersionHistoryClose}>
          <div className={styles.versionHistoryModal} onClick={(event) => event.stopPropagation()}>
            <h2 className={`${styles.versionHistoryTitle} ${styles.versionHistoryListTitle}`}>Version History</h2>
            <button
              type="button"
              className={styles.versionHistoryClose}
              onClick={handleVersionHistoryClose}
              aria-label="Close version history"
            >
              ×
            </button>
            <div className={styles.versionHistoryTableWrap}>
              <table className={styles.versionHistoryTable}>
                <thead>
                  <tr>
                    <th>Version</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {isVersionHistoryLoading ? (
                    <tr>
                      <td colSpan={2} className={styles.versionHistoryMessage}>Loading version history...</td>
                    </tr>
                  ) : versionHistoryError ? (
                    <tr>
                      <td colSpan={2} className={styles.versionHistoryMessage}>{versionHistoryError}</td>
                    </tr>
                  ) : versionHistoryItems.length > 0 ? (
                    versionHistoryItems.map((version) => {
                      const versionDownloadKey = getVersionDownloadKey(version);
                      const isVersionDownloading = downloadingVersionKeys[versionDownloadKey] === true;

                      return (
                      <tr
                        key={`${version.id}-${version.versionLabel}`}
                        style={version.isCurrentVersion ? {
                          background: 'linear-gradient(90deg, #f4f0ff 0%, #fff 100%)',
                          boxShadow: 'inset 4px 0 0 #5b35d5'
                        } : undefined}
                      >
                        <td>
                          <span style={version.isCurrentVersion ? { fontWeight: 700, color: '#3b238f' } : undefined}>
                            {`Version ${version.versionLabel} | ${formatVersionDate(version.contentRefreshDate || version.created)}`}
                          </span>
                          {version.isCurrentVersion && (
                            <span style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              marginLeft: 10,
                              padding: '3px 10px',
                              borderRadius: 999,
                              background: '#5b35d5',
                              color: '#fff',
                              fontSize: 11,
                              fontWeight: 700
                            }}>
                              Current
                            </span>
                          )}
                        </td>
                        <td>
                          <div className={styles.versionHistoryActions}>
                            <button
                              type="button"
                              className={styles.versionHistoryIconButton}
                              onClick={() => handleVersionView(version)}
                              aria-label={`View version ${version.versionLabel}`}
                            >
                              <svg width="13" height="13" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                                <path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6S2 12 2 12Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                                <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="2" />
                              </svg>
                            </button>
                            {!shouldHideDownloadActions && (
                              <button
                                type="button"
                                className={styles.versionHistoryIconButton}
                                disabled={isVersionDownloading}
                                onClick={() => handleVersionDownload(version)}
                                aria-label={isVersionDownloading ? `Downloading version ${version.versionLabel}` : `Download version ${version.versionLabel}`}
                                title={isVersionDownloading ? 'Downloading...' : 'Download'}
                                style={isVersionDownloading ? { opacity: 0.7 } : undefined}
                              >
                                {isVersionDownloading ? (
                                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                                    <circle cx="12" cy="12" r="8" stroke="currentColor" strokeWidth="2" opacity="0.25" />
                                    <path d="M20 12a8 8 0 0 0-8-8" stroke="currentColor" strokeWidth="2" strokeLinecap="round">
                                      <animateTransform attributeName="transform" type="rotate" from="0 12 12" to="360 12 12" dur="0.8s" repeatCount="indefinite" />
                                    </path>
                                  </svg>
                                ) : (
                                  <svg width="13" height="13" viewBox="0 0 24 24" fill="none" aria-hidden="true">
                                    <path d="M12 4v10m0 0 4-4m-4 4-4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                                    <path d="M5 16v3h14v-3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                                  </svg>
                                )}
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                      );
                    })
                  ) : (
                    <tr>
                      <td colSpan={2} className={styles.versionHistoryMessage}>No version history found.</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
      {(isVersionPreviewLoading || viewingVersion) && (
        <div className={styles.versionHistoryOverlay} onClick={handleVersionPreviewClose}>
          <div
            className={styles.versionInformationModal}
            onClick={(event) => event.stopPropagation()}
            style={{
              position: 'relative',
              width: 'min(920px, calc(100vw - 48px))',
              height: 'min(820px, calc(100vh - 48px))',
              padding: '84px 28px 28px',
              borderRadius: 14,
              background: '#f8fafc',
              boxShadow: '0 10px 24px rgba(17, 24, 39, 0.22)',
              boxSizing: 'border-box',
              display: 'flex',
              flexDirection: 'column',
              overflow: 'hidden'
            }}
          >
            <h2
              className={styles.versionHistoryTitle}
              style={{
                position: 'absolute',
                top: 24,
                left: 28,
                right: 72,
                margin: 0,
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                whiteSpace: 'nowrap'
              }}
            >
              {viewingVersion ? `Version ${viewingVersion.versionLabel}` : 'Version Information'}
            </h2>
            {viewingVersion?.fileName && (
              <div style={{
                position: 'absolute',
                top: 56,
                left: 28,
                right: 72,
                color: '#4b5563',
                fontSize: 13,
                overflow: 'hidden',
                textOverflow: 'ellipsis',
                whiteSpace: 'nowrap'
              }}>
                {viewingVersion.fileName}
              </div>
            )}
            <button
              type="button"
              className={styles.versionHistoryClose}
              onClick={handleVersionPreviewClose}
              aria-label="Close version preview"
            >
              ×
            </button>
            <div style={{ flex: 1, minHeight: 0, display: 'grid', gridTemplateColumns: '1fr', gap: 24, overflow: 'hidden' }}>
              <div style={{ minHeight: 0, display: 'none', flexDirection: 'column', overflow: 'hidden' }}>
                <h3 style={{ margin: '0 0 12px', fontSize: 16, color: '#111827', fontWeight: 700 }}>
                  Document Preview
                </h3>
                <div style={{ flex: 1, minHeight: 0, border: '1px solid #e5e7eb', background: '#fff', overflow: 'hidden', borderRadius: 10, boxShadow: '0 1px 3px rgba(15, 23, 42, 0.08)' }}>
                  {isVersionPreviewLoading ? (
                    <div style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      height: '100%',
                      color: '#6b7280',
                      fontSize: 14
                    }}>
                      Loading preview...
                    </div>
                  ) : viewingVersionPreview?.kind === 'image' ? (
                    <div style={{ width: '100%', height: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#f9fafb' }}>
                      <img
                        src={viewingVersionPreview.url}
                        alt={`Version ${viewingVersion?.versionLabel} preview`}
                        style={{ maxWidth: '100%', maxHeight: '100%', objectFit: 'contain' }}
                      />
                    </div>
                  ) : viewingVersionPreview?.kind === 'video' ? (
                    <video
                      src={viewingVersionPreview.url}
                      controls
                      style={{ width: '100%', height: '100%', background: '#000' }}
                    />
                  ) : viewingVersionPreview?.kind === 'audio' ? (
                    <div style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      height: '100%',
                      padding: 24
                    }}>
                      <audio src={viewingVersionPreview.url} controls style={{ width: '100%' }} />
                    </div>
                  ) : viewingVersionPreview?.kind === 'unsupported' ? (
                    <div style={{
                      display: 'flex',
                      flexDirection: 'column',
                      alignItems: 'center',
                      justifyContent: 'center',
                      height: '100%',
                      gap: 12,
                      color: '#6b7280',
                      textAlign: 'center',
                      padding: 24
                    }}>
                      <svg width="56" height="56" viewBox="0 0 24 24" fill="none" stroke="#c8c8c8" strokeWidth="1">
                        <path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" />
                        <polyline points="14 2 14 8 20 8" />
                      </svg>
                      <p style={{ margin: 0, fontSize: 14, color: '#888' }}>
                        {viewingVersionPreview.message || 'Preview is not available for this version.'}
                      </p>
                    </div>
                  ) : viewingVersionPreview?.url ? (
                    <iframe
                      title={`Version ${viewingVersion?.versionLabel} preview`}
                      src={viewingVersionPreview.url}
                      style={{ width: '100%', height: '100%', border: 'none' }}
                      allowFullScreen
                    />
                  ) : (
                    <div style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      height: '100%',
                      color: '#6b7280',
                      fontSize: 14
                    }}>
                      Preview is not available for this version.
                    </div>
                  )}
                </div>
              </div>
              <div style={{ minHeight: 0, overflowY: 'auto', paddingRight: 4 }}>
                <h3 style={{ margin: '0 0 16px', fontSize: 18, color: '#111827', fontWeight: 800 }}>
                  Version Information
                </h3>
                {viewingVersion && (
                  <div style={{
                    display: 'grid',
                    gridTemplateColumns: 'minmax(120px, 0.72fr) minmax(260px, 1.56fr) minmax(150px, 0.72fr)',
                    alignItems: 'stretch',
                    gap: 12,
                    marginBottom: 16
                  }}>
                    {[
                      { label: 'Version', value: viewingVersion.versionLabel },
                      { label: 'Document Type', value: viewingVersion.documentType || '-' },
                      { label: 'Status', value: viewingVersion.status || '-', isStatus: true }
                    ].map((item) => (
                      <div
                        key={item.label}
                        style={{
                          background: '#ffffff',
                          border: '1px solid #e5e7eb',
                          borderRadius: 12,
                          minHeight: 0,
                          padding: '16px 20px',
                          boxShadow: '0 1px 3px rgba(15, 23, 42, 0.08)',
                          display: 'flex',
                          flexDirection: 'column',
                          alignItems: 'center',
                          justifyContent: 'center',
                          minWidth: 0,
                          textAlign: 'center'
                        }}
                      >
                        <div style={{ color: '#6b7280', fontSize: 12, fontWeight: 800, textTransform: 'uppercase', letterSpacing: 0 }}>
                          {item.label}
                        </div>
                        {item.isStatus ? (
                          <span
                            className={getStatusBadgeClassName(String(item.value || ''))}
                            style={{ marginTop: 10, fontSize: 13, lineHeight: '18px', padding: '4px 10px' }}
                          >
                            {item.value || '-'}
                          </span>
                        ) : (
                          <div style={{
                            color: '#111827',
                            fontSize: item.label === 'Document Type' ? 19 : 21,
                            fontWeight: 800,
                            lineHeight: item.label === 'Document Type' ? 1.25 : 1.15,
                            marginTop: 10,
                            overflowWrap: 'anywhere',
                            wordBreak: 'normal',
                            textAlign: 'center'
                          }}>
                            {item.value}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
                {isVersionPreviewLoading && !viewingVersion && (
                  <section
                    style={{
                      background: '#ffffff',
                      border: '1px solid #e5e7eb',
                      borderRadius: 10,
                      boxShadow: '0 1px 3px rgba(15, 23, 42, 0.08)',
                      marginBottom: 12,
                      overflow: 'hidden'
                    }}
                  >
                    <div style={{
                      padding: '10px 14px',
                      background: '#f4ecff',
                      color: '#3b238f',
                      fontSize: 12,
                      fontWeight: 800,
                      textTransform: 'uppercase',
                      letterSpacing: 0
                    }}>
                      Loading
                    </div>
                    <div style={{
                      alignItems: 'center',
                      color: '#4b5563',
                      display: 'flex',
                      fontSize: 15,
                      fontWeight: 700,
                      justifyContent: 'center',
                      minHeight: 120,
                      padding: '18px 14px',
                      textAlign: 'center'
                    }}>
                      Loading version details...
                    </div>
                  </section>
                )}
                {viewingVersion && [
                  {
                    title: 'Document Details',
                    rows: [
                      { label: 'Title', value: viewingVersion.title },
                      { label: 'File Name', value: viewingVersion.fileName },
                      { label: 'Description', value: viewingVersion.description }
                    ]
                  },
                  {
                    title: 'People',
                    rows: [
                      { label: 'Author', value: viewingVersion.author + (viewingVersion.authorEmail ? ` (${viewingVersion.authorEmail})` : '') },
                      { label: 'Modified By', value: viewingVersion.modifiedBy + (viewingVersion.modifiedByEmail ? ` (${viewingVersion.modifiedByEmail})` : '') },
                      { label: 'Created By', value: viewingVersion.createdBy + (viewingVersion.createdByEmail ? ` (${viewingVersion.createdByEmail})` : '') }
                    ]
                  },
                  {
                    title: 'Classification',
                    rows: [
                      { label: 'BU', value: viewingVersion.bu },
                      { label: 'Department', value: viewingVersion.department },
                      { label: 'Document Type', value: viewingVersion.documentType },
                      { label: 'Geography', value: viewingVersion.geography },
                      { label: 'Client', value: viewingVersion.client },
                      { label: 'Disease Area', value: viewingVersion.diseaseArea },
                      { label: 'Therapy Area', value: viewingVersion.therapyArea },
                      { label: 'Project ID', value: viewingVersion.projectId }
                    ]
                  },
                  {
                    title: 'Version Details',
                    rows: [
                      { label: 'Version', value: viewingVersion.versionLabel },
                      { label: 'Status', value: viewingVersion.status, isStatus: true },
                      { label: 'Modified', value: formatVersionDate(viewingVersion.modifiedDate || viewingVersion.created) },
                      { label: 'Created', value: formatVersionDate(viewingVersion.createdDate || viewingVersion.created) },
                      { label: 'Published', value: formatVersionDate(viewingVersion.published) },
                      { label: 'Content Refresh Date', value: formatVersionDate(viewingVersion.contentRefreshDate) }
                    ]
                  }
                ].map((section) => (
                  <section
                    key={section.title}
                    style={{
                      background: '#ffffff',
                      border: '1px solid #e5e7eb',
                      borderRadius: 10,
                      boxShadow: '0 1px 3px rgba(15, 23, 42, 0.08)',
                      marginBottom: 12,
                      overflow: 'hidden'
                    }}
                  >
                    <div style={{
                      padding: '10px 14px',
                      background: '#f4ecff',
                      color: '#3b238f',
                      fontSize: 12,
                      fontWeight: 800,
                      textTransform: 'uppercase',
                      letterSpacing: 0
                    }}>
                      {section.title}
                    </div>
                    <div style={{ padding: '4px 14px' }}>
                      {section.rows.map((row) => (
                        <div
                          key={row.label}
                          style={{
                            display: 'grid',
                            gridTemplateColumns: '180px minmax(0, 1fr)',
                            gap: 16,
                            padding: '12px 0',
                            borderBottom: '1px solid #f3f4f6',
                            fontSize: 13
                          }}
                        >
                          <div style={{ color: '#6b7280', fontWeight: 700 }}>{row.label}</div>
                          <div style={{ color: '#1f2937', overflowWrap: 'anywhere', lineHeight: 1.35 }}>
                            {(row as any).isStatus ? (
                              <span className={getStatusBadgeClassName(String(row.value || ''))}>
                                {row.value || '-'}
                              </span>
                            ) : (
                              row.value || '-'
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  </section>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}
      {showShareDialog && (
        <div
          onClick={() => setShowShareDialog(false)}
          style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100%',
            height: '100%',
            backgroundColor: 'rgba(0,0,0,0.48)',
            zIndex: 99999,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}
        >
          <div
            onClick={event => event.stopPropagation()}
            style={{
              width: shareSent ? 440 : 480,
              minHeight: shareSent ? 288 : 319,
              background: '#ffffff',
              borderRadius: 8,
              boxShadow: '0 24px 56px rgba(0,0,0,0.28)',
              fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
              overflow: 'hidden',
              boxSizing: 'border-box',
              transform: 'translateY(-30px)'
            }}
            role="dialog"
            aria-modal="true"
            aria-label="Share document"
          >
            {shareSent ? (
              <div
                style={{
                  minHeight: 288,
                  position: 'relative',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'flex-start',
                  textAlign: 'center',
                  padding: '52px 60px 64px',
                  boxSizing: 'border-box'
                }}
              >
                <button
                  aria-label="Close"
                  style={{
                    position: 'absolute',
                    top: 24,
                    right: 24,
                    width: 32,
                    height: 32,
                    background: 'transparent',
                    border: 'none',
                    borderRadius: 2,
                    cursor: 'pointer',
                    padding: 0,
                    color: '#605e5c',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontFamily: '"Segoe UI", sans-serif'
                  }}
                  onClick={() => {
                    setShareSent(false);
                    setShowShareDialog(false);
                    setShareSelectedPeople([]);
                    setShareMessage('');
                  }}
                  onMouseEnter={event => {
                    event.currentTarget.style.background = '#f3f2f1';
                    event.currentTarget.style.color = '#201f1e';
                  }}
                  onMouseLeave={event => {
                    event.currentTarget.style.background = 'transparent';
                    event.currentTarget.style.color = '#605e5c';
                  }}
                >
                  <Icon iconName="Cancel" styles={{ root: { fontSize: 14, lineHeight: '14px' } }} />
                </button>
                <div
                  aria-hidden="true"
                  style={{
                    width: 52,
                    height: 52,
                    borderRadius: '50%',
                    background: 'rgb(237, 249, 240)',
                    border: '1px solid rgb(186, 216, 192)',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    marginBottom: 20
                  }}
                >
                  <Icon iconName="CompletedSolid" styles={{ root: { color: 'rgb(16, 124, 16)', fontSize: 26, lineHeight: '26px' } }} />
                </div>
                <div
                  style={{
                    color: 'rgb(50, 49, 48)',
                    fontSize: 20,
                    lineHeight: '27px',
                    fontWeight: 600,
                    maxWidth: 384
                  }}
                >
                  You&apos;ve invited {getShareInviteeText()} to view
                  <br />
                  <span style={{ display: 'inline-block', marginTop: 4 }}>
                    &quot;{getShareDialogTitle()}&quot;
                  </span>
                </div>
              </div>
            ) : (
              <>
                <div style={{ padding: '24px 24px 10px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', boxSizing: 'border-box' }}>
                  <span style={{ fontSize: 20, lineHeight: '26.6667px', fontWeight: 600, color: 'rgb(50, 49, 48)', letterSpacing: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: 376 }}>
                    Share &quot;{getShareDialogTitle()}&quot;
                  </span>
                  <span style={{ position: 'relative', display: 'inline-flex' }}>
                    <button
                      aria-label="Close"
                      style={{
                        width: 32,
                        height: 32,
                        background: 'transparent',
                        border: 'none',
                        borderRadius: 2,
                        cursor: 'pointer',
                        padding: 0,
                        color: '#605e5c',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        fontFamily: '"Segoe UI", sans-serif'
                      }}
                      onClick={() => {
                        setShareTooltip(null);
                        setShowShareDialog(false);
                      }}
                      onMouseEnter={event => {
                        setShareTooltip('close');
                        event.currentTarget.style.background = '#f3f2f1';
                        event.currentTarget.style.color = '#201f1e';
                      }}
                      onMouseLeave={event => {
                        setShareTooltip(null);
                        event.currentTarget.style.background = 'transparent';
                        event.currentTarget.style.color = '#605e5c';
                      }}
                    >
                      <Icon iconName="Cancel" styles={{ root: { fontSize: 14, lineHeight: '14px' } }} />
                    </button>
                    {renderShareTooltip('close', 'Close', {
                      top: 38,
                      right: -7,
                      arrow: 'top',
                      arrowRight: 16
                    })}
                  </span>
                </div>

                <div style={{ padding: '0 24px', position: 'relative', zIndex: 2 }}>
                  <div
                    style={{
                      border: shareValidationError
                        ? '1px solid rgb(232, 17, 35)'
                        : '0.666667px solid rgb(200, 198, 196)',
                      borderRadius: 4,
                      display: 'flex',
                      alignItems: shareSelectedPeople.length ? 'flex-start' : 'center',
                      flexWrap: 'wrap',
                      padding: shareSelectedPeople.length ? '10px 8px 12px 38px' : '0 0 0 38px',
                      gap: 6,
                      minHeight: shareSelectedPeople.length ? 76 : 44,
                      height: shareSelectedPeople.length ? 'auto' : 44,
                      width: 432,
                      overflow: 'visible',
                      position: 'relative',
                      boxSizing: 'border-box',
                      cursor: 'text'
                    }}
                    onClick={() => {
                      setShareFocusedField('people');
                      sharePeopleInputRef.current?.focus();
                    }}
                  >
                    <Icon iconName="Contact" styles={{ root: { color: 'rgb(50, 49, 48)', position: 'absolute', left: 13, top: shareSelectedPeople.length ? 18 : 15, flexShrink: 0, fontSize: 12, lineHeight: '12px', width: 12, height: 12 } }} />
                    {shareSelectedPeople.map((person, index) => (
                      <span key={person.accountName || person.email || index} style={{
                        background: 'rgb(243, 242, 241)',
                        borderRadius: 16,
                        height: 28,
                        padding: '0 4px 0 0',
                        fontSize: 14,
                        color: '#201f1e',
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: 4,
                        maxWidth: shareSelectedPeople.length > 1 ? 184 : 250,
                        overflow: 'hidden'
                      }}>
                        {renderSharePersonaAvatar(28)}
                        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {person.name}
                        </span>
                        <button
                          type="button"
                          aria-label={`Remove ${person.name}`}
                          onClick={event => {
                            event.stopPropagation();
                            setShareSelectedPeople(prev => prev.filter((_, itemIndex) => itemIndex !== index));
                            window.setTimeout(() => sharePeopleInputRef.current?.focus(), 0);
                          }}
                          style={{
                            width: 20,
                            height: 20,
                            border: 'none',
                            borderRadius: 2,
                            background: 'transparent',
                            color: 'rgb(164, 38, 44)',
                            cursor: 'pointer',
                            display: 'inline-flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            padding: 0,
                            flexShrink: 0
                          }}
                        >
                          <Icon iconName="ChromeClose" styles={{ root: { fontSize: 10, lineHeight: '10px', fontWeight: 400 } }} />
                        </button>
                      </span>
                    ))}
                    {shareSelectedPeople.length > 0 && (
                      <span aria-hidden="true" style={{ flexBasis: '100%', height: 0, minWidth: '100%' }} />
                    )}
                    <input
                      ref={sharePeopleInputRef}
                      role="combobox"
                      aria-label={shareSelectedPeople.length ? 'Add more' : 'Add a name or email'}
                      aria-expanded={sharePeopleResults.length > 0 || sharePeopleLoading || sharePeopleQuery.trim().length > 0}
                      placeholder={shareSelectedPeople.length ? 'Add more' : 'Add a name or email'}
                      value={sharePeopleQuery}
                      onFocus={() => setShareFocusedField('people')}
                      onBlur={() => setShareFocusedField(null)}
                      onKeyDown={event => {
                        if (event.key === 'ArrowDown') {
                          event.preventDefault();
                          setShareActivePersonIndex(prev => Math.min(prev + 1, Math.max(sharePeopleResults.length - 1, 0)));
                        } else if (event.key === 'ArrowUp') {
                          event.preventDefault();
                          setShareActivePersonIndex(prev => Math.max(prev - 1, 0));
                        } else if (event.key === 'Enter' && sharePeopleResults[shareActivePersonIndex]) {
                          event.preventDefault();
                          selectSharePerson(sharePeopleResults[shareActivePersonIndex]);
                        } else if (event.key === 'Escape') {
                          setSharePeopleQuery('');
                          setSharePeopleResults([]);
                          setShareActivePersonIndex(0);
                          setShareHoveredPersonIndex(null);
                        }
                      }}
                      onChange={event => {
                        setSharePeopleQuery(event.target.value);
                        searchSharePeople(event.target.value);
                        if (shareValidationError) setShareValidationError('');
                      }}
                      style={{
                        border: 'none',
                        outline: 'none',
                        fontSize: 14,
                        color: 'rgb(50, 49, 48)',
                        flex: 1,
                        minWidth: shareSelectedPeople.length ? 120 : 160,
                        maxWidth: shareSelectedPeople.length ? 318 : undefined,
                        fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                        padding: 0,
                        lineHeight: 'normal',
                        height: shareSelectedPeople.length ? 20 : 18,
                        background: '#ffffff',
                        cursor: 'text'
                      }}
                    />
                    <span
                      aria-hidden="true"
                      style={{
                        position: 'absolute',
                        left: 0,
                        bottom: 0,
                        width: '100%',
                        height: 4,
                        borderBottom: shareValidationError
                          ? '1px solid rgb(232, 17, 35)'
                          : shareFocusedField === 'people'
                            ? '2px solid rgb(164, 38, 44)'
                            : '1px solid rgb(59, 58, 57)',
                        borderBottomLeftRadius: 200,
                        borderBottomRightRadius: 200,
                        boxSizing: 'border-box',
                        clipPath: 'inset(calc(100% - 2px) 0 0)',
                        pointerEvents: 'none'
                      }}
                    />
                  </div>
                  {shareValidationError && (
                    <div style={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: 6,
                      color: 'rgb(164, 38, 44)',
                      fontSize: 12,
                      lineHeight: '16px',
                      marginTop: 6,
                      fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
                    }}>
                      <span aria-hidden="true" style={{
                        width: 11,
                        height: 11,
                        borderRadius: '50%',
                        background: 'rgb(164, 38, 44)',
                        color: '#ffffff',
                        display: 'inline-flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        fontSize: 8,
                        fontWeight: 700,
                        lineHeight: '11px',
                        flexShrink: 0
                      }}>
                        !
                      </span>
                      <span>{shareValidationError}</span>
                    </div>
                  )}

                  {(sharePeopleResults.length > 0 || sharePeopleLoading || sharePeopleQuery.trim().length > 0) && (
                    <div ref={sharePeopleListRef} role="listbox" style={{
                      border: '1px solid #edebe9',
                      borderRadius: 2,
                      background: '#ffffff',
                      boxShadow: 'rgba(0, 0, 0, 0.133) 0px 6.4px 14.4px 0px, rgba(0, 0, 0, 0.11) 0px 1.2px 3.6px 0px',
                      minHeight: sharePeopleLoading && sharePeopleResults.length === 0 ? 54 : undefined,
                      height: sharePeopleResults.length > 0 ? sharePeopleResults.length * 54 : undefined,
                      maxHeight: 136,
                      overflowY: sharePeopleResults.length > 2 ? 'auto' : 'visible',
                      zIndex: 100000,
                      position: 'absolute',
                      top: '100%',
                      left: 24,
                      width: 432,
                      boxSizing: 'border-box',
                      display: 'flex',
                      flexDirection: 'column'
                    }}>
                      {sharePeopleLoading && sharePeopleResults.length === 0 && (
                        <div style={{
                          minHeight: 54,
                          display: 'flex',
                          flexDirection: 'column',
                          alignItems: 'center',
                          justifyContent: 'center',
                          gap: 4,
                          color: 'rgb(164, 38, 44)',
                          fontSize: 12,
                          lineHeight: '16px',
                          fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
                        }}>
                          <Spinner
                            size={SpinnerSize.medium}
                            styles={{ circle: { borderTopColor: 'rgb(164, 38, 44)', borderRightColor: 'rgba(164, 38, 44, 0.25)', borderBottomColor: 'rgba(164, 38, 44, 0.25)', borderLeftColor: 'rgba(164, 38, 44, 0.25)' } }}
                          />
                          <span>Searching...</span>
                        </div>
                      )}
                      {!sharePeopleLoading && sharePeopleResults.length === 0 && sharePeopleQuery.trim().length > 0 && (
                        <div style={{
                          minHeight: 32,
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          color: 'rgb(96, 94, 92)',
                          fontSize: 12,
                          lineHeight: '16px',
                          fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
                        }}>
                          No results
                        </div>
                      )}
                      {sharePeopleResults.map((person, index) => (
                        <button
                          type="button"
                          key={person.accountName || person.email || index}
                          data-share-person-index={index}
                          role="option"
                          aria-selected={shareActivePersonIndex === index}
                          onClick={() => selectSharePerson(person)}
                          style={{
                            minHeight: 54,
                            height: 54,
                            width: '100%',
                            padding: '0 10px',
                            cursor: 'pointer',
                            display: 'flex',
                            alignItems: 'center',
                            gap: 8,
                            fontSize: 14,
                            color: '#201f1e',
                            boxSizing: 'border-box',
                            background: shareActivePersonIndex === index
                              ? 'rgb(237, 235, 233)'
                              : shareHoveredPersonIndex === index
                                ? 'rgb(243, 242, 241)'
                                : '#ffffff',
                            border: shareActivePersonIndex === index ? '1px solid rgb(96, 94, 92)' : '1px solid transparent',
                            borderRadius: 0,
                            fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                            textAlign: 'left'
                          }}
                          onMouseEnter={() => setShareHoveredPersonIndex(index)}
                          onMouseLeave={() => setShareHoveredPersonIndex(null)}
                        >
                          {renderSharePersonaAvatar(40)}
                          <div style={{ minWidth: 0, overflow: 'hidden' }}>
                            <div style={{ fontWeight: 400, lineHeight: '20px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{person.name}</div>
                            {person.subtitle && (
                              <div style={{ fontSize: 12, lineHeight: '16px', color: '#605e5c', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{person.subtitle}</div>
                            )}
                          </div>
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                <div style={{ padding: '12px 24px 0 24px' }}>
                  <div
                    style={{
                      borderTop: '1px solid rgb(200, 198, 196)',
                      borderRight: '1px solid rgb(200, 198, 196)',
                      borderLeft: '1px solid rgb(200, 198, 196)',
                      borderBottom: 'unset',
                      borderRadius: 4,
                      width: 432,
                      height: 105.667,
                      boxSizing: 'border-box',
                      position: 'relative',
                      display: 'flex'
                    }}
                    onClick={() => {
                      setShareFocusedField('message');
                      shareMessageInputRef.current?.focus();
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'flex-start', width: '100%' }}>
                      <Icon iconName="Handwriting" styles={{ root: { color: 'rgb(50, 49, 48)', flexShrink: 0, fontSize: 14, lineHeight: '14px', width: 14, height: 14, padding: '13px 0 0 13px' } }} />
                      <textarea
                        ref={shareMessageInputRef}
                        placeholder="Add a message"
                        value={shareMessage}
                        onFocus={() => setShareFocusedField('message')}
                        onBlur={() => setShareFocusedField(null)}
                        onChange={event => setShareMessage(event.target.value)}
                        style={{
                          width: '100%',
                          border: 'none',
                          outline: 'none',
                          resize: 'none',
                          fontSize: 14,
                          color: 'rgb(32, 31, 30)',
                          fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif',
                          height: 89,
                          background: 'transparent',
                          lineHeight: '20px',
                          padding: '10px 8px 6px'
                        }}
                      />
                    </div>
                    <span
                      aria-hidden="true"
                      style={{
                        position: 'absolute',
                        left: 0,
                        bottom: 0,
                        width: '100%',
                        height: 4,
                        borderBottom: shareFocusedField === 'message'
                          ? '2px solid rgb(164, 38, 44)'
                          : '1px solid rgb(59, 58, 57)',
                        borderBottomLeftRadius: 200,
                        borderBottomRightRadius: 200,
                        boxSizing: 'border-box',
                        clipPath: 'inset(calc(100% - 2px) 0 0)',
                        pointerEvents: 'none'
                      }}
                    />
                  </div>
                </div>

                <div style={{
                  padding: '34px 24px 24px 24px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'flex-end',
                  gap: 12,
                  rowGap: '1.25rem'
                }}>
                  <span style={{ position: 'relative', display: 'inline-flex', alignItems: 'center' }}>
                    <button
                      onClick={handleShareCopyLink}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: 4,
                        minHeight: '2rem',
                        height: '2rem',
                        minWidth: '5.9375rem',
                        padding: '0 16px',
                        border: '1px solid rgb(174, 56, 62)',
                        borderRadius: 4,
                        background: 'rgb(251, 244, 244)',
                        cursor: 'pointer',
                        fontSize: '0.875rem',
                        fontWeight: 600,
                        color: 'rgb(147, 34, 39)',
                        fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
                      }}
                      onMouseEnter={event => {
                        setShareTooltip('copy');
                        event.currentTarget.style.background = 'rgb(227, 175, 178)';
                      }}
                      onMouseLeave={event => {
                        setShareTooltip(null);
                        event.currentTarget.style.background = 'rgb(251, 244, 244)';
                      }}
                    >
                      <Icon iconName="Link" styles={{ root: { fontSize: 16, lineHeight: '16px', width: 16, height: 16, padding: '2px 0', margin: '0 4px', position: 'relative', top: 1 } }} />
                      {shareCopied ? 'Copied!' : 'Copy link'}
                    </button>
                    {renderShareTooltip('copy', 'People in Indegene Limited with the link can view.', {
                      bottom: 40,
                      left: '50%',
                      transform: 'translateX(-50%)',
                      width: 320,
                      arrow: 'bottom',
                      iconName: 'Lock'
                    })}
                  </span>
                  <span style={{ position: 'relative', display: 'inline-flex' }}>
                    <button
                      onClick={handleShareSend}
                      disabled={shareSending}
                      style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: 4,
                        minHeight: '2rem',
                        height: '2rem',
                        minWidth: '5.9375rem',
                        padding: '0 16px',
                        border: 'none',
                        borderRadius: 4,
                        background: shareSending ? 'rgb(243, 242, 241)' : 'rgb(164, 38, 44)',
                        cursor: shareSending ? 'default' : 'pointer',
                        fontSize: '0.875rem',
                        fontWeight: 600,
                        color: shareSending ? 'rgb(164, 38, 44)' : '#ffffff',
                        fontFamily: '"Segoe UI", "Segoe UI Web (West European)", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif'
                      }}
                      onMouseEnter={event => {
                        setShareTooltip('send');
                        if (!shareSending) event.currentTarget.style.background = 'rgb(147, 34, 39)';
                      }}
                      onMouseLeave={event => {
                        setShareTooltip(null);
                        if (!shareSending) event.currentTarget.style.background = 'rgb(164, 38, 44)';
                      }}
                    >
                      {shareSending ? (
                        <Spinner
                          size={SpinnerSize.medium}
                          styles={{
                            root: { width: 20, height: 20, padding: 0 },
                            circle: {
                              width: 20,
                              height: 20,
                              borderWidth: 2.5,
                              borderTopColor: 'rgb(164, 38, 44)',
                              borderRightColor: 'rgba(164, 38, 44, 0.22)',
                              borderBottomColor: 'rgba(164, 38, 44, 0.22)',
                              borderLeftColor: 'rgba(164, 38, 44, 0.22)'
                            }
                          }}
                        />
                      ) : (
                        <>
                          <Icon iconName="Send" styles={{ root: { fontSize: 16, lineHeight: '16px', width: 16, height: 16, padding: '2px 0', margin: '0 4px', position: 'relative', top: 1 } }} />
                          Send
                        </>
                      )}
                    </button>
                    {renderShareTooltip('send', 'Send an invite in an email.', {
                      bottom: 40,
                      right: 0,
                      arrow: 'bottom',
                      arrowRight: 28
                    })}
                  </span>
                </div>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
};
