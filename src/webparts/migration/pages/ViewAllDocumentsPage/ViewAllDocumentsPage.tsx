import * as React from 'react';
import { IViewAllDocumentsPageProps } from './IViewAllDocumentsPageProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { configurePermalinkService, NAV_PATHS, pushPageUrl, updateItemUrlField } from '../../services/permalinkService';
import { MetadataForm } from '../../components/metadata/MetadataForm/MetadataForm';
import styles from './ViewAllDocumentsPage.module.scss';
import {
  buildKMDataHubItemQuery,
  fetchKMDataHubReadFieldMap,
  formatDocumentPublishedDate,
  resolveDocumentPublishedValue,
  resolveDocumentAuthor,
  splitDocumentAuthorDisplay
} from '../../utils/documentMetadata';
import { SIDEBAR_GROUP_IDS } from '../../components/access.constants';
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
import { readAllDocumentBookmarksFromSharePoint } from '../../services/documentBookmarks';
import { recordUserEvent } from '../../services/kmMetricsService';
import { getDriveItemThumbnailUrl, getExcelThumbnailUrl, getSearchServiceInstance, SharePointSearchService } from '../../services/SharePointSearchService';
import { IKShellFooter, IKShellHeader, IKShellHeaderNavKey } from '../../components/shell/IKShellChrome';
import { CACHE_KEYS, COLUMN_NAMES, KM_REVIEW_HUB_DRIVE_ID, LIBRARY_NAMES, LIST_NAMES, PAGE_SIZES, SEARCH_PROPERTIES } from '../../config/appConfig';
import { downloadSharePointFile } from '../../utils/fileDownload';

interface DocumentItem {
  id: number;
  name: string;
  fileName: string;
  abstract: string;
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
  author: string;
  fileType: string;
  date: string;
  createdTime: number;
  status: string;
  fileSize: string;
  serverRelativeUrl: string;
  fileRef: string;
  fileUniqueId?: string;
  versionLabel?: string;
}

type ThumbnailCandidate = string | { isVideo: true; posterUrl: string | null; sourceUrl?: string; useVideoElement?: boolean };

interface IKMDataHubFieldMap {
  status: string;
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
  projectId?: string;
  contentRefreshDate?: string;
  versionFileName?: string;
  versionFileType?: string;
  docIcon?: string;
  modifiedBy?: string;
}

const LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;
const DOCUMENTS_PER_PAGE = PAGE_SIZES.documentsPerPage;
const FETCH_BUFFER = 5;
const FIELD_MAP_CACHE_KEY = CACHE_KEYS.fieldMap;
const TAXONOMY_CACHE_KEY = CACHE_KEYS.taxonomy;
const USER_CACHE_KEY = CACHE_KEYS.currentUser;
let fieldMapModuleCache: any = null;
type DocumentSortOrder = 'newToOld' | 'oldToNew';
type DocumentStatusFilter = 'all' | 'allNonActive' | 'Under Review' | 'Active' | 'Archive' | 'Reject';
type DocumentViewMode = 'list' | 'grid';
type ControlDropdownKey = 'sort' | 'status' | null;
const MY_DOC_ADMIN_GROUP_IDS = [SIDEBAR_GROUP_IDS.admins];
const KM_REVIEW_ALL_STATUS_FILTERS: DocumentStatusFilter[] = ['Archive', 'Under Review', 'Reject'];
const MY_DOCUMENT_ALL_STATUS_FILTERS: DocumentStatusFilter[] = ['Active', 'Archive', 'Under Review', 'Reject'];
type PaginationItem = number | 'ellipsis-start' | 'ellipsis-end';

const getDocumentFieldValue = (item: any, fieldName?: string, fallbacks: string[] = []): any => {
  const fieldNames = [fieldName, ...fallbacks].filter((value): value is string => Boolean(value));
  for (const name of fieldNames) {
    const value = item?.[name];
    if (value !== undefined && value !== null && value !== '') {
      return value;
    }
  }
  return undefined;
};

const normalizeNumberField = (value: any): number => {
  const numericValue = Number(value);
  return Number.isFinite(numericValue) ? numericValue : 0;
};

const normalizeBooleanField = (value: any): boolean => {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value > 0;
  const normalizedValue = String(value || '').trim().toLowerCase();
  return normalizedValue === 'true' || normalizedValue === 'yes' || normalizedValue === '1';
};

const normalizeStringArrayField = (value: any): string[] => {
  if (Array.isArray(value)) {
    return value
      .map((entry) => String(entry?.Label || entry?.Title || entry || '').trim())
      .filter(Boolean);
  }
  return String(value || '')
    .split(/[;,|]/)
    .map((entry) => entry.trim())
    .filter(Boolean);
};

interface IMyDocAccess {
  id: number;
  title: string;
  email: string;
  loginName: string;
  isKmAdmin: boolean;
}

interface IViewDocumentsFetchContext {
  applyScopeFilters: (sourceItems: any[]) => any[];
  mapDocuments: (items: any[], authorLookup?: Record<number, string>) => DocumentItem[];
  orderByField?: string;
  queryParts?: { select: string; expand: string };
  statusField?: string;
}

const getSessionCache = <T,>(key: string): T | null => {
  try {
    const cached = window.sessionStorage.getItem(key);
    return cached ? JSON.parse(cached) as T : null;
  } catch {
    return null;
  }
};

const setSessionCache = (key: string, value: unknown): void => {
  try {
    window.sessionStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore storage failures; data fetching still works without cache.
  }
};

const getCachedFieldMap = <T,>(): T | null => getSessionCache<T>(FIELD_MAP_CACHE_KEY);
const setCachedFieldMap = (map: unknown): void => setSessionCache(FIELD_MAP_CACHE_KEY, map);
const getCachedTaxonomy = (): ITaxonomyFieldOptions | null => getSessionCache<ITaxonomyFieldOptions>(TAXONOMY_CACHE_KEY);
const setCachedTaxonomy = (taxonomy: ITaxonomyFieldOptions): void => setSessionCache(TAXONOMY_CACHE_KEY, taxonomy);
const getCachedUser = (): { Id: number; Title?: string; Email?: string; LoginName?: string } | null =>
  getSessionCache<{ Id: number; Title?: string; Email?: string; LoginName?: string }>(USER_CACHE_KEY);
const setCachedUser = (user: { Id: number; Title?: string; Email?: string; LoginName?: string }): void =>
  setSessionCache(USER_CACHE_KEY, user);

const KM_REVIEW_STATUS_FILTER_OPTIONS: Array<{ value: DocumentStatusFilter; label: string }> = [
  { value: 'allNonActive', label: 'All Status' },
  { value: 'Archive', label: 'Archive' },
  { value: 'Under Review', label: 'Under Review' },
  { value: 'Reject', label: 'Reject' }
];

const SORT_OPTIONS: Array<{ value: DocumentSortOrder; label: string }> = [
  { value: 'newToOld', label: 'Newest first' },
  { value: 'oldToNew', label: 'Oldest first' }
];

const MY_DOCUMENT_STATUS_FILTER_OPTIONS: Array<{ value: DocumentStatusFilter; label: string }> = [
  { value: 'all', label: 'All Status' },
  { value: 'Active', label: 'Active' },
  { value: 'Archive', label: 'Archive' },
  { value: 'Under Review', label: 'Under Review' },
  { value: 'Reject', label: 'Reject' }
];

const ACTIVE_ONLY_STATUS_FILTER_OPTIONS: Array<{ value: DocumentStatusFilter; label: string }> = [
  { value: 'Active', label: 'Active' }
];

const buildPaginationItems = (currentPage: number, totalPages: number): PaginationItem[] => {
  if (totalPages <= 7) {
    return Array.from({ length: totalPages }, (_, index) => index + 1);
  }

  const pages = new Set<number>([1, totalPages, currentPage - 1, currentPage, currentPage + 1]);

  if (currentPage <= 4) {
    [2, 3, 4, 5].forEach((page) => pages.add(page));
  }

  if (currentPage >= totalPages - 3) {
    [totalPages - 4, totalPages - 3, totalPages - 2, totalPages - 1].forEach((page) => pages.add(page));
  }

  const orderedPages = Array.from(pages)
    .filter((page) => page >= 1 && page <= totalPages)
    .sort((first, second) => first - second);

  return orderedPages.reduce<PaginationItem[]>((items, page, index) => {
    const previousPage = orderedPages[index - 1];
    if (index > 0 && page - previousPage > 1) {
      items.push(previousPage === 1 ? 'ellipsis-start' : 'ellipsis-end');
    }
    items.push(page);
    return items;
  }, []);
};

const normalizeIdentityValue = (value: unknown): string =>
  typeof value === 'string' ? value.trim().toLowerCase() : '';

const collectIdentityTokens = (currentUser: IMyDocAccess): Set<string> => {
  const normalizedTokens = [
    currentUser.title,
    currentUser.email,
    currentUser.loginName,
    currentUser.loginName.split('|').pop() || ''
  ]
    .map(normalizeIdentityValue)
    .filter(Boolean);

  return new Set(normalizedTokens);
};

const collectPersonFieldTokens = (value: any): string[] => {
  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.reduce<string[]>((tokens, entry) => {
      tokens.push(...collectPersonFieldTokens(entry));
      return tokens;
    }, []);
  }

  if (typeof value === 'string') {
    return value
      .split(/[;,]/)
      .map((entry) => entry.trim())
      .filter(Boolean);
  }

  if (typeof value === 'object') {
    return [
      value.Title,
      value.Name,
      value.Email,
      value.EMail,
      value.LoginName
    ]
      .filter((entry): entry is string => typeof entry === 'string' && entry.trim().length > 0)
      .map((entry) => entry.trim());
  }

  return [];
};

const toNumberArray = (value: unknown): number[] => {
  if (Array.isArray(value)) {
    return value.reduce<number[]>((ids, entry) => {
      ids.push(...toNumberArray(entry));
      return ids;
    }, []);
  }

  if (value && typeof value === 'object') {
    const resultValues = (value as { results?: unknown }).results;
    if (resultValues !== undefined) {
      return toNumberArray(resultValues);
    }
  }

  const numericValue = Number(value);
  return Number.isFinite(numericValue) && numericValue > 0 ? [numericValue] : [];
};

const documentMatchesMyDocAccess = (
  item: any,
  currentUser: IMyDocAccess,
  authorFieldInternalName: string
): boolean => {
  // Created By in SharePoint document libraries is the system Author field.
  if (Number(item?.AuthorId) === currentUser.id) {
    return true;
  }

  // Published By is maintained in the KM author/publisher metadata field.
  const authorIds = toNumberArray(item?.[`${authorFieldInternalName}Id`]);
  if (authorIds.includes(currentUser.id)) {
    return true;
  }

  const identityTokens = collectIdentityTokens(currentUser);
  const authorTokens = collectPersonFieldTokens(item?.[authorFieldInternalName]).map(normalizeIdentityValue);
  const creatorTokens = collectPersonFieldTokens(item?.Author).map(normalizeIdentityValue);

  return [...authorTokens, ...creatorTokens].some((token) => identityTokens.has(token));
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

const normalizeGuid = (value?: string): string => (value || '').replace(/[{}]/g, '').trim().toLowerCase();

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

const resolveBuDepartmentTermsFromSeparateFields = (
  item: Record<string, any>,
  textValues: Record<string, any>,
  allTerms: ITaxonomyTerm[]
): ITaxonomyTerm[] => {
  const resolveTerms = (internalName: string): ITaxonomyTerm[] => {
    const resolvedFromRaw = extractTermsFromValue(item[internalName], allTerms);
    if (resolvedFromRaw.length > 0) {
      return resolvedFromRaw;
    }

    return extractTermsFromValue(textValues[internalName], allTerms);
  };

  const departmentTerms = resolveTerms(KM_DATA_HUB_DEPARTMENT_FIELD_INTERNAL_NAME);
  if (departmentTerms.length > 0) {
    return departmentTerms;
  }

  return resolveTerms(KM_DATA_HUB_BU_FIELD_INTERNAL_NAME);
};

const buildKMDataHubTaxonomyFormValues = (
  metadata: Record<string, any>,
  fieldMap: IKMDataHubFieldMap
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

const getStatusBadgeClassName = (status?: string): string => {
  const normalizedStatus = (status || '').trim().toLowerCase();

  if (normalizedStatus === 'active' || normalizedStatus === 'approved') {
    return styles.statusBadgeActive;
  }

  if (normalizedStatus === 'under review' || normalizedStatus === 'review') {
    return styles.statusBadgeReview;
  }

  if (normalizedStatus === 'reject' || normalizedStatus === 'rejected') {
    return styles.statusBadgeRejected;
  }

  if (normalizedStatus === 'archive' || normalizedStatus === 'archived') {
    return styles.statusBadgeArchive;
  }

  return styles.statusBadgeDefault;
};

const getDisplayStatus = (status?: string): string => {
  const trimmedStatus = (status || '').trim();
  return trimmedStatus.toLowerCase() === 'rejected' ? 'Reject' : trimmedStatus;
};

const documentMatchesStatusFilter = (documentStatus: string, filter: DocumentStatusFilter): boolean => {
  const normalizedStatus = (documentStatus || '').trim().toLowerCase();
  const normalizedFilter = filter.toLowerCase();

  if (normalizedFilter === 'all') {
    return true;
  }

  if (normalizedFilter === 'allnonactive') {
    return documentMatchesAnyStatusFilter(documentStatus, KM_REVIEW_ALL_STATUS_FILTERS);
  }

  if (normalizedFilter === 'reject') {
    return normalizedStatus === 'reject' || normalizedStatus === 'rejected';
  }

  if (normalizedFilter === 'archive') {
    return normalizedStatus === 'archive' || normalizedStatus === 'archived';
  }

  if (normalizedFilter === 'under review') {
    return normalizedStatus === 'under review' || normalizedStatus === 'underreview' || normalizedStatus === 'review';
  }

  return normalizedStatus === normalizedFilter;
};

const documentMatchesAnyStatusFilter = (documentStatus: string, filters: DocumentStatusFilter[]): boolean =>
  filters.some((filter) => documentMatchesStatusFilter(documentStatus, filter));

const documentMatchesKmReviewStatusFilter = (documentStatus: string, filter: DocumentStatusFilter): boolean =>
  filter === 'allNonActive'
    ? documentMatchesAnyStatusFilter(documentStatus, KM_REVIEW_ALL_STATUS_FILTERS)
    : documentMatchesStatusFilter(documentStatus, filter);

const documentMatchesMyDocumentsStatusFilter = (documentStatus: string, filter: DocumentStatusFilter): boolean =>
  filter === 'all'
    ? documentMatchesAnyStatusFilter(documentStatus, MY_DOCUMENT_ALL_STATUS_FILTERS)
    : documentMatchesStatusFilter(documentStatus, filter);

const buildStatusFilterExpression = (fieldName: string, filters: DocumentStatusFilter[]): string =>
  filters
    .map((filter) => `${fieldName} eq '${filter.replace(/'/g, "''")}'`)
    .join(' or ');

const renderAuthorStack = (author: string, containerClassName: string, itemClassName: string): JSX.Element => {
  const authorEntries = splitDocumentAuthorDisplay(author);
  const entriesToRender = authorEntries.length > 0 ? authorEntries : ['Internal'];

  return (
    <div className={containerClassName}>
      {entriesToRender.map((entry, index) => (
        <span key={`${entry}-${index}`} className={itemClassName}>
          {entry}
        </span>
      ))}
    </div>
  );
};

export const ViewAllDocumentsPage: React.FunctionComponent<IViewAllDocumentsPageProps> = (props) => {
  const knowledgeSearchApiClient = React.useMemo(
    () => new KnowledgeSearchApiClient(undefined, props.context),
    [props.context]
  );
  const pageScope = props.scope || 'all';
  const isKmReviewHubScope = pageScope === 'kmReviewHub';
  const isRecentlyPublishedScope = pageScope === 'recentlyPublished';
  const isMyDocScope = pageScope === 'myDoc';
  const isBookmarksScope = pageScope === 'bookmarks';
  const pageHeading = isRecentlyPublishedScope
    ? 'Knowledge Base'
    : isBookmarksScope
      ? 'My Bookmarks'
      : isMyDocScope
      ? 'My Documents'
      : 'KM Review Hub';
  const pageTagline = isBookmarksScope
    ? 'View knowledge assets you have bookmarked for quick and easy access.'
    : isMyDocScope
      ? 'View and manage knowledge assets you have uploaded or listed as an author.'
      : undefined;
  const activeHeaderNavKey: IKShellHeaderNavKey | undefined = isBookmarksScope
    ? 'bookmarks'
    : isMyDocScope
      ? 'documents'
      : isRecentlyPublishedScope
        ? undefined
        : 'kmReviewHub';
  const hideDownloadButtonsForLearner = props.isLearner === true;
  const [documents, setDocuments] = React.useState<DocumentItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [countLoading, setCountLoading] = React.useState<boolean>(false);
  const [thumbnailAttemptByDocument, setThumbnailAttemptByDocument] = React.useState<Record<number, number>>({});
  const [thumbnailLoadedByDocument, setThumbnailLoadedByDocument] = React.useState<Record<number, boolean>>({});
  const [downloadingDocumentIds, setDownloadingDocumentIds] = React.useState<Record<number, boolean>>({});
  const [, setCurrentPage] = React.useState<number>(1);
  const [nextPageUrl, setNextPageUrl] = React.useState<string | null>(null);
  const [totalCount, setTotalCount] = React.useState<number>(0);
  const [currentRestPage, setCurrentRestPage] = React.useState<number>(1);
  const [sortOrder, setSortOrder] = React.useState<DocumentSortOrder>('newToOld');
  const [statusFilter, setStatusFilter] = React.useState<DocumentStatusFilter>(
    isKmReviewHubScope ? 'allNonActive' : isRecentlyPublishedScope || isBookmarksScope ? 'Active' : isMyDocScope ? 'all' : 'Active'
  );
  const [openControlDropdown, setOpenControlDropdown] = React.useState<ControlDropdownKey>(null);
  const [viewMode, setViewMode] = React.useState<DocumentViewMode>('list');
  const [taxonomyOptions, setTaxonomyOptions] = React.useState<ITaxonomyFieldOptions | null>(null);
  const [editingDocument, setEditingDocument] = React.useState<DocumentItem | null>(null);
  const [editInitialValues, setEditInitialValues] = React.useState<Record<string, any> | null>(null);
  const [editLoading, setEditLoading] = React.useState<boolean>(false);
  const [editSaving, setEditSaving] = React.useState<boolean>(false);
  const [editError, setEditError] = React.useState<string | null>(null);
  const kmDataHubFieldMapRef = React.useRef<IKMDataHubFieldMap | null>(null);
  const kmsUsersRef = React.useRef<IKmsUser[] | null>(null);
  const documentsFetchInitializedRef = React.useRef<boolean>(false);
  const countRequestIdRef = React.useRef<number>(0);
  const restCountCacheRef = React.useRef<Map<string, {
    restCount: number;
    activeRestCount: number;
    totalSearchCount: number;
    libraryTotal: number;
  }>>(new Map());
  const pageUrlMapRef = React.useRef<Map<number, string>>(new Map());
  const documentsFetchContextRef = React.useRef<IViewDocumentsFetchContext | null>(null);
  const bookmarkedIdsArrayRef = React.useRef<number[]>([]);
  const searchStartOffsetAdjustmentRef = React.useRef<number>(0);
  const sortDirection: 'ascending' | 'descending' = sortOrder === 'oldToNew' ? 'ascending' : 'descending';
  const restSortDirection: 'asc' | 'desc' = sortOrder === 'oldToNew' ? 'asc' : 'desc';
  const hasExplicitStatusFilter =
    statusFilter !== 'all' &&
    statusFilter !== 'allNonActive' &&
    statusFilter !== 'Active';
  const effectiveSearchSortDirection: 'ascending' | 'descending' = hasExplicitStatusFilter ? 'descending' : sortDirection;
  const effectiveRestSortDirection: 'asc' | 'desc' = hasExplicitStatusFilter ? 'desc' : restSortDirection;

  React.useEffect(() => {
    if (props.context) {
      configurePermalinkService(props.context.spHttpClient, props.context.pageContext.web.absoluteUrl);
      documentsFetchInitializedRef.current = false;
      countRequestIdRef.current += 1;
      pageUrlMapRef.current = new Map();
      searchStartOffsetAdjustmentRef.current = 0;
      setCurrentRestPage(1);
      setCurrentPage(1);
      setTotalCount(0);
      setCountLoading(false);
      setDocuments([]);
      void fetchAllDocuments();
      void loadTaxonomyOptions();
    }
  }, [pageScope, props.context]);

  React.useEffect(() => {
    if (!documentsFetchInitializedRef.current) {
      return;
    }
    pageUrlMapRef.current = new Map();
    countRequestIdRef.current += 1;
    setCurrentRestPage(1);
    setCurrentPage(1);
    setTotalCount(0);
    setCountLoading(false);
    void fetchAllDocuments();
  }, [sortOrder, statusFilter]);

  React.useEffect(() => {
    setStatusFilter(isKmReviewHubScope ? 'allNonActive' : isRecentlyPublishedScope || isBookmarksScope ? 'Active' : isMyDocScope ? 'all' : 'Active');
  }, [isBookmarksScope, isKmReviewHubScope, isMyDocScope, isRecentlyPublishedScope]);

  const ensureListExists = React.useCallback(
    async (
      listName: string,
      fields: Array<{ name: string; type: string; required?: boolean }>
    ): Promise<boolean> => {
      if (!props.context) return false;

      const webUrl = props.context.pageContext.web.absoluteUrl;

      try {
        const checkResp = await props.context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${listName}')?$select=Id`,
          SPHttpClient.configurations.v1
        );

        if (checkResp.ok) {
          return true;
        }

        const createListResp = await props.context.spHttpClient.post(
          `${webUrl}/_api/web/lists`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose'
            },
            body: JSON.stringify({
              __metadata: { type: 'SP.List' },
              BaseTemplate: 100,
              Title: listName
            })
          }
        );

        if (!createListResp.ok) {
          return false;
        }

        for (const field of fields) {
          const fieldBody = {
            __metadata: { type: 'SP.Field' },
            Title: field.name,
            FieldTypeKind: field.type === 'Number' ? 9 : field.type === 'Note' ? 3 : 2,
            Required: field.required || false
          };

          const fieldResp = await props.context.spHttpClient.post(
            `${webUrl}/_api/web/lists/getbytitle('${listName}')/fields`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose'
              },
              body: JSON.stringify(fieldBody)
            }
          );

          if (!fieldResp.ok) {
            return false;
          }
        }

        return true;
      } catch {
        return false;
      }
    },
    [props.context]
  );

  const getListEntityType = React.useCallback(async (listName: string): Promise<string | null> => {
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
  }, [props.context]);

  const loadTaxonomyOptions = React.useCallback(async (): Promise<ITaxonomyFieldOptions | null> => {
    if (!props.context) {
      return null;
    }

    const cachedTaxonomy = getCachedTaxonomy();
    if (cachedTaxonomy) {
      setTaxonomyOptions(cachedTaxonomy);
      return cachedTaxonomy;
    }

    if (taxonomyOptions) {
      return taxonomyOptions;
    }

    try {
      const loadedOptions = await fetchAllTaxonomyOptions(props.context);
      setTaxonomyOptions(loadedOptions);
      setCachedTaxonomy(loadedOptions);
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

  const resolveMyDocAccess = React.useCallback(async (): Promise<IMyDocAccess> => {
    if (!props.context) {
      throw new Error('Context is not available.');
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const cachedUser = getCachedUser();
    if (cachedUser?.Id) {
      return {
        id: Number(cachedUser.Id || 0),
        title: String(cachedUser.Title || ''),
        email: String(cachedUser.Email || ''),
        loginName: String(cachedUser.LoginName || ''),
        isKmAdmin: false
      };
    }

    const userResponse = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/currentuser?$select=Id,Title,Email,LoginName`,
      SPHttpClient.configurations.v1
    );

    if (!userResponse.ok) {
      const errorText = await userResponse.text();
      throw new Error(`Failed to load current user: ${userResponse.status} ${errorText}`);
    }

    const user = await userResponse.json();
    setCachedUser({
      Id: Number(user?.Id || 0),
      Title: String(user?.Title || ''),
      Email: String(user?.Email || ''),
      LoginName: String(user?.LoginName || '')
    });
    let isKmAdmin = false;

    try {
      const graphClient = await props.context.msGraphClientFactory.getClient('3');
      const currentUserGroupIds = new Set<string>();
      let requestPath = '/me/transitiveMemberOf?$select=id';

      while (requestPath) {
        const membershipResponse: {
          value?: Array<{ id?: string }>;
          '@odata.nextLink'?: string;
        } = await graphClient
          .api(requestPath)
          .version('v1.0')
          .get();

        (membershipResponse.value || []).forEach((entry) => {
          const groupId = (entry.id || '').trim().toLowerCase();
          if (groupId) {
            currentUserGroupIds.add(groupId);
          }
        });

        const nextLink = membershipResponse['@odata.nextLink'];
        if (!nextLink) {
          requestPath = '';
          continue;
        }

        try {
          const nextUrl = new URL(nextLink);
          requestPath = `${nextUrl.pathname}${nextUrl.search}`.replace(/^\/v1\.0/i, '');
        } catch {
          requestPath = nextLink.replace(/^https:\/\/graph\.microsoft\.com\/v1\.0/i, '');
        }
      }

      isKmAdmin = MY_DOC_ADMIN_GROUP_IDS.some((groupId) =>
        currentUserGroupIds.has(groupId.toLowerCase())
      );
    } catch (error) {
      console.warn('Unable to resolve KM admin access for My Doc. Falling back to author/creator/modifier checks only.', error);
    }

    return {
      id: Number(user?.Id || 0),
      title: String(user?.Title || ''),
      email: String(user?.Email || ''),
      loginName: String(user?.LoginName || ''),
      isKmAdmin
    };
  }, [props.context]);

  const resolveCurrentSharePointUserId = React.useCallback(async (): Promise<number> => {
    const legacyUserId = Number((props.context?.pageContext as any)?.legacyPageContext?.userId || 0);
    if (legacyUserId > 0) {
      return legacyUserId;
    }

    if (!props.context) {
      return 0;
    }

    try {
      const cachedUser = getCachedUser();
      if (cachedUser?.Id) {
        return Number(cachedUser.Id || 0);
      }

      const webUrl = props.context.pageContext.web.absoluteUrl;
      const userResponse = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/currentuser?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (!userResponse.ok) {
        return 0;
      }

      const user = await userResponse.json();
      setCachedUser({
        Id: Number(user?.Id || user?.d?.Id || 0),
        Title: String(user?.Title || user?.d?.Title || '')
      });
      return Number(user?.Id || user?.d?.Id || 0);
    } catch (error) {
      console.warn('Unable to resolve current SharePoint user for bookmarks.', error);
      return 0;
    }
  }, [props.context]);

  const fetchAuthorTitlesByIds = React.useCallback(
    async (authorIds: number[]): Promise<Record<number, string>> => {
      if (!props.context || authorIds.length === 0) {
        return {};
      }

      const authorMap: Record<number, string> = {};
      authorIds.forEach((authorId) => {
        if (typeof authorId !== 'number' || authorId <= 0 || authorMap[authorId]) {
          return;
        }
        const cachedTitle = SharePointSearchService.getCachedAuthorTitle(authorId);
        if (cachedTitle) {
          authorMap[authorId] = cachedTitle;
        }
      });

      const uniqueAuthorIds = Array.from(new Set(authorIds.filter((id) => typeof id === 'number' && id > 0 && !authorMap[id])));
      if (uniqueAuthorIds.length === 0) {
        return authorMap;
      }

      const webUrl = props.context.pageContext.web.absoluteUrl;

      await Promise.all(
        uniqueAuthorIds.map(async (authorId) => {
          try {
            const response = await props.context!.spHttpClient.get(
              `${webUrl}/_api/web/getuserbyid(${authorId})?$select=Id,Title`,
              SPHttpClient.configurations.v1
            );

            if (!response.ok) {
              return;
            }

            const user = await response.json();
            if (user?.Title) {
              SharePointSearchService.setCachedAuthorTitle(authorId, user.Title);
              authorMap[authorId] = user.Title;
            }
          } catch {
            // Ignore individual author lookup failures and keep existing fallbacks.
          }
        })
      );

      return authorMap;
    },
    [props.context]
  );

  const recordDownloadMetric = React.useCallback(async (documentId: number, documentTitle?: string): Promise<void> => {
    if (!props.context) {
      return;
    }

    const cachedUser = getCachedUser();
    let currentUserId = Number(cachedUser?.Id || 0);

    if (!currentUserId) {
      const userResponse = await props.context.spHttpClient.get(
        `${props.context.pageContext.web.absoluteUrl}/_api/web/currentuser?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (!userResponse.ok) {
        return;
      }

      const user = await userResponse.json();
      setCachedUser({
        Id: Number(user?.Id || 0),
        Title: String(user?.Title || '')
      });
      currentUserId = Number(user.Id || 0);
    }

    if (!currentUserId) {
      return;
    }

    await recordUserEvent(props.context, documentId, currentUserId, 'Download', documentTitle);
  }, [props.context]);

  const getKMDataHubFieldMap = React.useCallback(async (): Promise<IKMDataHubFieldMap> => {
    if (!props.context) {
      throw new Error('Context is not available.');
    }

    if (kmDataHubFieldMapRef.current) {
      return kmDataHubFieldMapRef.current;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const response = await props.context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/fields?$select=Title,InternalName,Hidden,Id,TextField&$top=500`,
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
      sensitiveTerms: findInternalName('Sensitive Terms', COLUMN_NAMES.sensitiveTerms)
    };

    kmDataHubFieldMapRef.current = fieldMap;
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

  const updateKMDataHubTextFields = React.useCallback(async (
    itemId: number,
    metadata: Record<string, any>,
    fieldMap: IKMDataHubFieldMap
  ): Promise<void> => {
    if (!props.context) {
      return;
    }

    const webUrl = props.context.pageContext.web.absoluteUrl;
    const body: Record<string, any> = {
      [fieldMap.status]: 'Under Review',
      [fieldMap.title]: metadata.title || '-',
      [fieldMap.description]: metadata.description || '-',
      [fieldMap.sensitiveTerms]: metadata.sensitiveTerms || ''
    };

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

    const formValues = Object.keys(body)
      .map((fieldName) => ({
        FieldName: fieldName,
        FieldValue: body[fieldName]
      }))
      .filter((entry) => entry.FieldValue !== undefined && entry.FieldValue !== null);

    const response = await props.context.spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})/ValidateUpdateListItem`,
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

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to update KM Data Hub text fields: ${response.status} ${errorText}`);
    }

    const responseJson = await response.json();
    const fieldResults = responseJson?.value || responseJson?.d?.ValidateUpdateListItem?.results || responseJson?.d?.results || [];
    const fieldErrors = fieldResults.filter((entry: any) => entry.HasException);
    if (fieldErrors.length > 0) {
      throw new Error(fieldErrors.map((entry: any) => `${entry.FieldName}: ${entry.ErrorMessage}`).join('; '));
    }
  }, [props.context]);

  const handleOpenEdit = React.useCallback(async (doc: DocumentItem): Promise<void> => {
    if (!props.context) {
      return;
    }

    setEditingDocument(doc);
    setEditLoading(true);
    setEditError(null);
    setEditInitialValues(null);

    try {
      const loadedTaxonomyOptions = await loadTaxonomyOptions();
      if (!loadedTaxonomyOptions) {
        throw new Error('Unable to load taxonomy options.');
      }
      const kmsUsers = await loadKmsUsers();

      const fieldMap = await getKMDataHubFieldMap();
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const selectFields = [
        'Id',
        COLUMN_NAMES.title,
        COLUMN_NAMES.description,
        `${fieldMap.author}/Title`,
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
      const response = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${doc.id})?$select=${selectFields.concat(noteFields).join(',')}&$expand=${fieldMap.author}`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to load document metadata: ${response.status} ${errorText}`);
      }

      const textResponse = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${doc.id})/FieldValuesAsText`,
        SPHttpClient.configurations.v1
      );

      const rawItem = await response.json();
      const item = rawItem?.d || rawItem || {};
      const textValuesJson = textResponse.ok ? await textResponse.json() : {};
      const textValues = textValuesJson?.d || textValuesJson || {};
      const resolveTermList = (rawValue: any, fallbackText: string, options: ITaxonomyTerm[]) => {
        const resolvedFromRaw = extractTermsFromValue(rawValue, options);
        if (resolvedFromRaw.length > 0) {
          return resolvedFromRaw;
        }

        return extractTermsFromValue(fallbackText, options);
      };
      const resolveAuthorUpns = (): string[] => {
        const authorObject = item[fieldMap.author];
        const authorEntries = Array.isArray(authorObject) ? authorObject : authorObject ? [authorObject] : [];
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

        return resolvedUpns;
      };
      const getFieldTextFallback = (internalName?: string, noteInternalName?: string): string =>
        (internalName ? item[internalName] || textValues[internalName] : '') ||
        (noteInternalName ? item[noteInternalName] || textValues[noteInternalName] : '') ||
        '';

      const buDepartmentTerms = resolveBuDepartmentTermsFromSeparateFields(item, textValues, loadedTaxonomyOptions.buDepartment);
      const documentTypeTerms = resolveTermList(item[fieldMap.documentType], getFieldTextFallback(fieldMap.documentType, fieldMap.documentTypeNote), loadedTaxonomyOptions.documentType);
      const clientTerms = resolveTermList(item[fieldMap.client], getFieldTextFallback(fieldMap.client, fieldMap.clientNote), loadedTaxonomyOptions.client);
      const geographyTerms = resolveTermList(item[fieldMap.geography], getFieldTextFallback(fieldMap.geography, fieldMap.geographyNote), loadedTaxonomyOptions.geography);
      const diseaseAreaTerms = resolveTermList(item[fieldMap.diseaseArea], getFieldTextFallback(fieldMap.diseaseArea, fieldMap.diseaseAreaNote), loadedTaxonomyOptions.diseaseArea);
      const therapyAreaTerms = resolveTermList(item[fieldMap.therapyArea], getFieldTextFallback(fieldMap.therapyArea, fieldMap.therapyAreaNote), loadedTaxonomyOptions.therapyArea);

      setEditInitialValues({
        title: item[fieldMap.title] || item.Title || textValues[fieldMap.title] || doc.name || '',
        description: item[fieldMap.description] || item.Description || textValues[fieldMap.description] || '',
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
        sensitiveTerms: item[fieldMap.sensitiveTerms] || textValues[fieldMap.sensitiveTerms] || ''
      });
    } catch (error) {
      console.error('Error preparing edit dialog:', error);
      setEditError(error instanceof Error ? error.message : 'Unable to load document details for editing.');
    } finally {
      setEditLoading(false);
    }
  }, [getKMDataHubFieldMap, loadKmsUsers, loadTaxonomyOptions, props.context]);

  const handleCloseEdit = React.useCallback(() => {
    if (editSaving) {
      return;
    }

    setEditingDocument(null);
    setEditInitialValues(null);
    setEditError(null);
    setEditLoading(false);
  }, [editSaving]);

  const handleEditSubmit = React.useCallback(async (formData: Record<string, any>): Promise<void> => {
    if (!props.context || !editingDocument) {
      return;
    }

    setEditSaving(true);
    setEditError(null);

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const fieldMap = await getKMDataHubFieldMap();
      // TODO: Wrap ViewAllDocumentsPage metadata updates with updateWithoutVersion once this edit flow is versioning-ready.
      await updateKMDataHubTextFields(editingDocument.id, formData, fieldMap);
      const previousStatus = String(editingDocument.status || '').trim();
      const newStatus = 'Under Review';
      if (newStatus && newStatus !== previousStatus) {
        try {
          const currentUserId = await resolveCurrentSharePointUserId();
          await props.context.spHttpClient.post(
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
                Title: formData.title || editingDocument.fileName || editingDocument.name,
                FileName: editingDocument.fileName || editingDocument.name,
                Action: newStatus,
                PerformedById: currentUserId,
                TimeStamp: new Date().toISOString()
              })
            }
          );
        } catch (auditError) {
          console.warn('Audit log entry creation failed:', auditError);
        }
      }

      const formValues = buildKMDataHubTaxonomyFormValues(formData, fieldMap);
      if (formValues.length > 0) {
        const response = await props.context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${editingDocument.id})/ValidateUpdateListItem`,
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

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to update metadata: ${response.status} ${errorText}`);
        }

        const responseJson = await response.json();
        const fieldResults = responseJson?.value || responseJson?.d?.ValidateUpdateListItem?.results || responseJson?.d?.results || [];
        const fieldErrors = fieldResults.filter((entry: any) => entry.HasException);
        if (fieldErrors.length > 0) {
          throw new Error(fieldErrors.map((entry: any) => `${entry.FieldName}: ${entry.ErrorMessage}`).join('; '));
        }
      }

      try {
        await updateItemUrlField(editingDocument.id, formData.title || editingDocument.name);
      } catch (urlError) {
        console.warn('Unable to sync permalink after metadata update:', urlError);
      }
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
        : editingDocument.author;
      if (knowledgeSearchApiClient.isConfigured()) {
        // TODO: Restore Easy Auth/token enforcement when backend sync auth is enabled.
        knowledgeSearchApiClient
          .triggerSync('frontend_upload_or_metadata_update')
          .catch((searchError) => console.warn('Search backend sync trigger skipped/failed after metadata edit:', searchError));
      }

      emitDocumentDataChanged({
        documentIds: [editingDocument.id],
        reason: 'metadata'
      });
      await fetchAllDocuments();
      handleCloseEdit();
    } catch (error) {
      console.error('Error saving document metadata:', error);
      setEditError(error instanceof Error ? error.message : 'Unable to save document changes.');
    } finally {
      setEditSaving(false);
    }
  }, [editingDocument, getKMDataHubFieldMap, handleCloseEdit, loadKmsUsers, props.context, resolveCurrentSharePointUserId, updateKMDataHubTextFields]);

  const fetchAllDocuments = async () => {
    if (!props.context) {
      setLoading(false);
      return;
    }

    try {
      setLoading(true);
      setThumbnailAttemptByDocument({});
      setThumbnailLoadedByDocument({});
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const libraryName = LIBRARY_NAMES.kmDataHub;
      const cachedFieldMap = fieldMapModuleCache || getCachedFieldMap<any>();
      const currentUserIdPromise = isBookmarksScope
        ? resolveCurrentSharePointUserId()
        : Promise.resolve(0);
      const [fieldMap, myDocAccess, resolvedBookmarkUserId] = await Promise.all([
        cachedFieldMap
          ? Promise.resolve(cachedFieldMap).then((map) => {
            fieldMapModuleCache = map;
            return map;
          })
          : fetchKMDataHubReadFieldMap(props.context.spHttpClient, webUrl, libraryName).then((map) => {
            setCachedFieldMap(map);
            fieldMapModuleCache = map;
            return map;
          }),
        isMyDocScope ? resolveMyDocAccess() : Promise.resolve(null),
        currentUserIdPromise
      ]);
      const queryParts = buildKMDataHubItemQuery(
        fieldMap,
        [
          'ID',
          fieldMap.title || COLUMN_NAMES.title,
          fieldMap.description || COLUMN_NAMES.description,
          fieldMap.docIcon || COLUMN_NAMES.docIcon,
          COLUMN_NAMES.created,
          COLUMN_NAMES.modified,
          fieldMap.published,
          COLUMN_NAMES.authorId,
          'Author/Title',
          'Author0/Title',
          'Author0Id',
          fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef,
          fieldMap.fileRef || COLUMN_NAMES.fileRef,
          'File/UniqueId',
          fieldMap.url || COLUMN_NAMES.url,
          COLUMN_NAMES.version,
          fieldMap.status
        ],
        ['Author', 'Author0', 'File']
      );
      let statusFilterExpression = isBookmarksScope || isRecentlyPublishedScope
        ? `${fieldMap.status} eq 'Active'`
        : isKmReviewHubScope
        ? buildStatusFilterExpression(fieldMap.status, KM_REVIEW_ALL_STATUS_FILTERS)
        : !isMyDocScope
        ? `${fieldMap.status} eq 'Active'`
        : !isMyDocScope && statusFilter !== 'all'
        ? `${fieldMap.status} eq '${statusFilter.replace(/'/g, "''")}'`
        : '';

      if (isMyDocScope && myDocAccess) {
        statusFilterExpression = '';
      }

      let bookmarkedIds = new Set<number>();
      let bookmarkedPaths = new Set<string>();
      let bookmarkRecords: Array<{ documentId: number; documentUrl: string }> = [];

      if (isBookmarksScope) {
        const currentUserId = resolvedBookmarkUserId;
        bookmarkRecords = await readAllDocumentBookmarksFromSharePoint(
          props.context.spHttpClient,
          webUrl,
          currentUserId
        );
        bookmarkedIds = new Set<number>(
          bookmarkRecords
            .map((bookmark) => Number(bookmark.documentId))
            .filter((documentId) => Number.isFinite(documentId) && documentId > 0)
        );
        bookmarkedPaths = new Set<string>(
          bookmarkRecords
            .map((bookmark) => (bookmark.documentUrl || '').trim().toLowerCase())
            .filter(Boolean)
        );
        bookmarkedIdsArrayRef.current = Array.from(bookmarkedIds);
        console.log('[BOOKMARKS COUNT] Total bookmarked IDs for this user:', bookmarkedIdsArrayRef.current.length);
        if (bookmarkedIds.size > 0) {
          statusFilterExpression = bookmarkedIdsArrayRef.current
            .slice(0, DOCUMENTS_PER_PAGE)
            .map((documentId) => `ID eq ${documentId}`)
            .join(' or ');
        } else {
          bookmarkedIdsArrayRef.current = [];
          setDocuments([]);
          setTotalCount(0);
          setCurrentRestPage(1);
          setCurrentPage(1);
          setLoading(false);
          return;
        }
      }

      const applyScopeFilters = (sourceItems: any[]): any[] => {
        let items = isMyDocScope && myDocAccess
          ? sourceItems.filter((item: any) => documentMatchesMyDocAccess(item, myDocAccess, fieldMap.author))
          : sourceItems;

        if (isBookmarksScope) {
          items = items.filter((item: any) => {
            const itemId = Number(item.Id || 0);
            const itemPath = String(item.FileRef || item.File?.ServerRelativeUrl || '').trim().toLowerCase();
            return bookmarkedIds.has(itemId) || (!!itemPath && bookmarkedPaths.has(itemPath));
          });
        }

        return items;
      };

      const mapDocuments = (items: any[], authorLookup: Record<number, string> = {}): DocumentItem[] =>
        items.map((item: any) => {
          const fileName = getDocumentFieldValue(item, fieldMap.fileLeafRef, [COLUMN_NAMES.fileLeafRef, COLUMN_NAMES.title]) || '';
          const fileExtension = fileName.split('.').pop()?.toUpperCase() || '';
          const titleFromColumn = String(getDocumentFieldValue(item, fieldMap.title, [COLUMN_NAMES.title]) || '').trim();
          const displayName = titleFromColumn && titleFromColumn !== '-' ? titleFromColumn : (fileName || `Document ${item.Id}`);
          const abstract = getDocumentFieldValue(item, fieldMap.description, [COLUMN_NAMES.description]) || '';
          const author = resolveDocumentAuthor(item, authorLookup[item.AuthorId] || 'Internal', fieldMap.author);
          const publishedValue = resolveDocumentPublishedValue(item, fieldMap.published);
          const publishedDate = publishedValue ? new Date(publishedValue) : null;
          const sortDateValue = isKmReviewHubScope
            ? (item.Created || publishedValue)
            : isMyDocScope
            ? (item.Created || publishedValue)
            : publishedValue;
          const sortDate = sortDateValue ? new Date(sortDateValue) : null;
          const itemStatus = String(item[fieldMap.status] || item.Status || '').trim();
          const displayDateValue = isMyDocScope
            ? (itemStatus === 'Active' ? publishedValue : item.Created)
            : isKmReviewHubScope
            ? item.Created
            : publishedValue;
          const formattedDate = formatDocumentPublishedDate(displayDateValue, 'en-GB', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
          });
          const status = itemStatus || (isRecentlyPublishedScope ? 'Active' : '');
          return {
            id: item.Id,
            name: displayName,
            fileName,
            abstract: abstract,
            reviewerComments: '',
            docIcon: String(getDocumentFieldValue(item, fieldMap.docIcon, [COLUMN_NAMES.docIcon]) || ''),
            author,
            fileType: fileExtension,
            date: formattedDate,
            createdTime: sortDate && !Number.isNaN(sortDate.getTime())
              ? sortDate.getTime()
              : publishedDate && !Number.isNaN(publishedDate.getTime())
                ? publishedDate.getTime()
                : 0,
            status,
            fileSize: '',
            serverRelativeUrl: getDocumentFieldValue(item, fieldMap.fileRef, [COLUMN_NAMES.fileRef]) || '',
            fileRef: getDocumentFieldValue(item, fieldMap.fileRef, [COLUMN_NAMES.fileRef]) || '',
            fileUniqueId: item.File?.UniqueId || '',
            versionLabel: item[COLUMN_NAMES.version]
          };
        });

      documentsFetchContextRef.current = {
        applyScopeFilters,
        mapDocuments,
        queryParts,
        statusField: fieldMap.status || COLUMN_NAMES.status
      };

      const orderByField = isBookmarksScope
        ? `${fieldMap.published || 'Created'} desc`
        : isRecentlyPublishedScope
        ? `${fieldMap.published || 'Created'} ${sortOrder === 'oldToNew' ? 'asc' : 'desc'}`
        : isKmReviewHubScope
        ? `${COLUMN_NAMES.created} ${effectiveRestSortDirection}`
        : isMyDocScope || hasExplicitStatusFilter
        ? `${COLUMN_NAMES.modified} ${effectiveRestSortDirection}`
        : sortOrder === 'oldToNew'
        ? `${fieldMap.published || 'Created'} asc`
        : `${fieldMap.published || 'Created'} desc`;
      const firstPageUrl =
        `${webUrl}/_api/web/lists/getbytitle('${libraryName}')/items` +
        `?$select=${queryParts.select}` +
        `&$expand=${queryParts.expand}` +
        `${statusFilterExpression ? `&$filter=${encodeURIComponent(statusFilterExpression)}` : ''}` +
        `&$orderby=${orderByField}` +
        `&$top=${DOCUMENTS_PER_PAGE}`;
      documentsFetchContextRef.current = {
        applyScopeFilters,
        mapDocuments,
        orderByField,
        queryParts,
        statusField: fieldMap.status || COLUMN_NAMES.status
      };

      if (isMyDocScope && myDocAccess) {
        const searchService = getSearchServiceInstance(props.context.spHttpClient, webUrl, libraryName);
        searchService.setFieldMapCache(fieldMap);
        const authorField = fieldMap.author || COLUMN_NAMES.author;
        const searchSortDirection = effectiveRestSortDirection === 'asc' ? 'ascending' : 'descending';
        const myDocumentsStatusFilter = statusFilter !== 'all' ? statusFilter : MY_DOCUMENT_ALL_STATUS_FILTERS;
        const [{ ids: authorIds }, authorTotal, author0Ids, author0Total, latestIds] = await Promise.all([
          searchService.getDocumentIdsByAuthor(
            myDocAccess.id,
            authorField,
            0,
            DOCUMENTS_PER_PAGE,
            myDocumentsStatusFilter,
            effectiveRestSortDirection
          ),
          searchService.getListItemCount(
            myDocumentsStatusFilter,
            undefined,
            myDocAccess.id,
            authorField
          ),
          searchService.getDocumentIdsByAuthor0Search(
            myDocAccess.email,
            searchSortDirection,
            0,
            DOCUMENTS_PER_PAGE,
            myDocumentsStatusFilter
          ),
          searchService.getAuthor0DocumentCount(myDocAccess.email, myDocumentsStatusFilter),
          searchService.getLatestDocumentIds(40)
        ]);
        const mergedIds = Array.from(new Set([...authorIds, ...author0Ids, ...latestIds])).slice(0, DOCUMENTS_PER_PAGE + 10);
        console.log('[MY DOCS COUNT] AuthorId REST count:', authorTotal);
        console.log('[MY DOCS COUNT] Author0 Search POST count:', author0Total);
        console.log('[MY DOCS COUNT] AuthorId page IDs:', authorIds.length);
        console.log('[MY DOCS COUNT] Author0 Search page IDs:', author0Ids.length);
        console.log('[MY DOCS COUNT] Latest REST IDs (new uploads):', latestIds.length);
        console.log('[MY DOCS COUNT] Merged unique IDs (this page):', mergedIds.length);
        console.log('[MY DOCS COUNT] Final estimated total shown to user:', Math.max(authorTotal, author0Total));
        const estimatedTotal = Math.max(authorTotal, author0Total);
        const rawItems = await searchService.getDocumentDetailsByIds(mergedIds, queryParts);
        const items = applyScopeFilters(rawItems).filter((item: any) =>
          documentMatchesMyDocAccess(item, myDocAccess, authorField) &&
          documentMatchesMyDocumentsStatusFilter(String(item[fieldMap.status || COLUMN_NAMES.status] || item.Status || ''), statusFilter)
        );
        const uniqueItems: any[] = Array.from(
          new Map(items.map((item: any) => [Number(item.ID || item.Id), item])).values()
        );
        const processedDocuments: DocumentItem[] = mapDocuments(uniqueItems, {});
        const seenDocumentIds = new Set<number>();
        const uniqueDocuments = processedDocuments.filter((document) => {
          if (seenDocumentIds.has(document.id)) {
            return false;
          }
          seenDocumentIds.add(document.id);
          return true;
        });
        uniqueDocuments.sort((a, b) => effectiveRestSortDirection === 'asc'
          ? (a.createdTime || 0) - (b.createdTime || 0)
          : (b.createdTime || 0) - (a.createdTime || 0)
        );

        pageUrlMapRef.current = new Map();
        setNextPageUrl(null);
        setCurrentRestPage(1);
        setCurrentPage(1);
        setTotalCount(estimatedTotal);
        setCountLoading(false);
        setDocuments(uniqueDocuments);
        documentsFetchInitializedRef.current = true;
        setLoading(false);
        return;
      }

      if (!isMyDocScope && !isBookmarksScope) {
        const searchService = getSearchServiceInstance(props.context.spHttpClient, webUrl, libraryName);
        searchService.setFieldMapCache(fieldMap);
        const statusFromItem = (item: any): string =>
          String(item[fieldMap.status || COLUMN_NAMES.status] || item.Status || '').trim();
        const selectedKmStatus = isKmReviewHubScope && statusFilter !== 'allNonActive'
          ? statusFilter
          : '';
        const kmReviewStatusFilter = selectedKmStatus || KM_REVIEW_ALL_STATUS_FILTERS;
        const searchStatus = isRecentlyPublishedScope
          ? 'Active'
          : isKmReviewHubScope
          ? kmReviewStatusFilter
          : 'Active';
        const sortField = isRecentlyPublishedScope
          ? SEARCH_PROPERTIES.published
          : isKmReviewHubScope
          ? SEARCH_PROPERTIES.write
          : SEARCH_PROPERTIES.published;
        const searchRowLimit = isKmReviewHubScope
          ? DOCUMENTS_PER_PAGE
          : DOCUMENTS_PER_PAGE + FETCH_BUFFER;
        const countCacheKey = `${pageScope}:${statusFilter}:${searchStatus}`;
        const cachedCountEntry = restCountCacheRef.current.get(countCacheKey);
        const cachedCount = cachedCountEntry && !isKmReviewHubScope && cachedCountEntry.libraryTotal > 0
          ? cachedCountEntry
          : undefined;
        const restCountPromise = cachedCount
          ? Promise.resolve(cachedCount.restCount)
          : isKmReviewHubScope
            ? selectedKmStatus
              ? searchService.getListItemCount(selectedKmStatus)
              : searchService.getListItemCount(KM_REVIEW_ALL_STATUS_FILTERS)
            : searchService.getListItemCount(searchStatus);
        const activeRestCountPromise = cachedCount
          ? Promise.resolve(cachedCount.activeRestCount)
          : isKmReviewHubScope
            ? Promise.resolve(0)
            : Promise.resolve(0);
        const totalSearchCountPromise = cachedCount
          ? Promise.resolve(cachedCount.totalSearchCount)
          : isKmReviewHubScope
            ? Promise.resolve(0)
            : searchService.getListItemCount();
        const libraryTotalPromise = cachedCount
          ? Promise.resolve(cachedCount.libraryTotal)
          : isKmReviewHubScope
            ? Promise.resolve(0)
            : searchService.getLibraryItemCount();
        const [
          searchResult,
          restCount,
          activeRestCount,
          totalSearchCount,
          libraryTotal,
          latestIds
        ] = await Promise.all([
          searchService.getDocumentIdsByStatus(searchStatus, 0, searchRowLimit, sortField, effectiveSearchSortDirection),
          restCountPromise,
          activeRestCountPromise,
          totalSearchCountPromise,
          libraryTotalPromise,
          searchService.getLatestDocumentIds(DOCUMENTS_PER_PAGE + 10)
        ]);
        let searchIds = searchResult.ids;
        const searchTotalRows = searchResult.totalRows;
        if (isKmReviewHubScope && searchIds.length === 0 && restCount > 0) {
          const restFallbackResult = await searchService.getDocumentIdsByStatusRest(
            kmReviewStatusFilter,
            0,
            searchRowLimit,
            fieldMap.status || COLUMN_NAMES.status,
            COLUMN_NAMES.created,
            effectiveRestSortDirection
          );
          searchIds = restFallbackResult.ids;
        }
        if (!cachedCount && !isKmReviewHubScope && libraryTotal > 0) {
          restCountCacheRef.current.set(countCacheKey, { restCount, activeRestCount, totalSearchCount, libraryTotal });
        }
        const uncrawledGap = isKmReviewHubScope
          ? 0
          : Math.max(0, libraryTotal - totalSearchCount);
        const mergedIds = Array.from(new Set([...searchIds, ...latestIds])).slice(0, DOCUMENTS_PER_PAGE + 10);
        const [
          [rawItems, indexedActiveLatestIds, indexedKmLatestIds],
          libraryStatusCounts
        ] = await Promise.all([
          Promise.all([
            searchService.getDocumentDetailsByIds(mergedIds, queryParts),
            isKmReviewHubScope && !selectedKmStatus
              ? Promise.resolve<number[] | null>([])
              : !isKmReviewHubScope
                ? searchService.getActiveSearchIdsByIds(latestIds)
                : Promise.resolve<number[] | null>([]),
            isKmReviewHubScope
              ? searchService.getSearchIdsByIds(latestIds, selectedKmStatus || KM_REVIEW_ALL_STATUS_FILTERS)
              : Promise.resolve<number[] | null>([])
          ]),
          !isMyDocScope && !isBookmarksScope
            ? searchService.getStatusCountsBySearch()
            : Promise.resolve(null)
        ]);
        const items = applyScopeFilters(rawItems).filter((item: any) => {
          const status = statusFromItem(item);
          if (isKmReviewHubScope) {
            return documentMatchesKmReviewStatusFilter(status, statusFilter);
          }
          return status === 'Active';
        });
        const latestIdSet = new Set(latestIds);
        const indexedActiveLatestIdSet = new Set(indexedActiveLatestIds || []);
        const restActiveLatestIds = rawItems
          .filter((item: any) => latestIdSet.has(Number(item.Id || item.ID || 0)) && statusFromItem(item) === 'Active')
          .map((item: any) => Number(item.Id || item.ID || 0))
          .filter((id: number) => id > 0);
        const restActiveNotIndexedCount = indexedActiveLatestIds
          ? restActiveLatestIds.filter((id: number) => !indexedActiveLatestIdSet.has(id)).length
          : uncrawledGap;
        const restInactiveButIndexedActiveCount = indexedActiveLatestIds
          ? rawItems
              .filter((item: any) => {
                const id = Number(item.Id || item.ID || 0);
                return latestIdSet.has(id) && indexedActiveLatestIdSet.has(id) && statusFromItem(item) !== 'Active';
              })
              .length
          : 0;
        const indexedKmLatestIdSet = new Set(indexedKmLatestIds || []);
        const liveKmMatchingLatestIds = rawItems
          .filter((item: any) => {
            const id = Number(item.Id || item.ID || 0);
            if (!latestIdSet.has(id)) return false;
            const status = statusFromItem(item);
            return documentMatchesKmReviewStatusFilter(status, statusFilter);
          })
          .map((item: any) => Number(item.Id || item.ID || 0))
          .filter((id: number) => id > 0);
        const liveKmMatchingLatestIdSet = new Set(liveKmMatchingLatestIds);
        const kmRestNotIndexedCount = isKmReviewHubScope && indexedKmLatestIds
          ? liveKmMatchingLatestIds.filter((id: number) => !indexedKmLatestIdSet.has(id)).length
          : 0;
        const kmIndexedButNoLongerMatchingCount = isKmReviewHubScope && indexedKmLatestIds
          ? Array.from(indexedKmLatestIdSet).filter((id: number) => !liveKmMatchingLatestIdSet.has(id)).length
          : 0;
        const sortedItems = [...items].sort((a, b) => {
          const dateA = new Date(
            isKmReviewHubScope
              ? a.Created || 0
              : a[fieldMap.published || COLUMN_NAMES.published] || a.Published || a.Created || 0
          ).getTime();
          const dateB = new Date(
            isKmReviewHubScope
              ? b.Created || 0
              : b[fieldMap.published || COLUMN_NAMES.published] || b.Published || b.Created || 0
          ).getTime();
          if (dateA || dateB) {
            return effectiveSearchSortDirection === 'ascending' ? dateA - dateB : dateB - dateA;
          }
          return Number(b.ID || b.Id) - Number(a.ID || a.Id);
        });
        const displayItems = sortedItems;
        const processedDocuments: DocumentItem[] = mapDocuments(displayItems, {});
        const seenDocumentIds = new Set<number>();
        const uniqueDocuments = processedDocuments.filter((document) => {
          if (seenDocumentIds.has(document.id)) {
            return false;
          }
          seenDocumentIds.add(document.id);
          return true;
        });
        uniqueDocuments.sort((a, b) => {
          if (isRecentlyPublishedScope || isKmReviewHubScope || (!isMyDocScope && !isBookmarksScope)) {
            const dateA = a.createdTime || new Date(a.date || 0).getTime();
            const dateB = b.createdTime || new Date(b.date || 0).getTime();
            return effectiveSearchSortDirection === 'ascending' ? dateA - dateB : dateB - dateA;
          }
          return b.id - a.id;
        });
        const finalDocuments = isKmReviewHubScope
          ? uniqueDocuments
          : uniqueDocuments.slice(0, DOCUMENTS_PER_PAGE);

        pageUrlMapRef.current = new Map();
        setNextPageUrl(null);
        setCurrentRestPage(1);
        setCurrentPage(1);
        const knownNonActiveCount = isKmReviewHubScope
          ? 0
          : Math.max(0, totalSearchCount - restCount);
        const activeCountFloor = isKmReviewHubScope
          ? 0
          : Math.max(0, libraryTotal - knownNonActiveCount - restInactiveButIndexedActiveCount);
        const searchNonActiveCount = Math.max(0, restCount - activeRestCount);
        const shouldApplyRestCorrection = !isKmReviewHubScope && (
          libraryTotal === 0 ||
          restCount < activeCountFloor ||
          totalSearchCount < libraryTotal ||
          restActiveNotIndexedCount > uncrawledGap
        );
        const adjustedRestActiveNotIndexedCount = shouldApplyRestCorrection
          ? restActiveNotIndexedCount
          : 0;
        const adjustedRestInactiveButIndexedActiveCount = !isKmReviewHubScope
          ? restInactiveButIndexedActiveCount
          : 0;
        const kmCorrectedCount = Math.max(0, searchNonActiveCount + kmRestNotIndexedCount - kmIndexedButNoLongerMatchingCount);
        const restCountBase = isKmReviewHubScope
          ? selectedKmStatus && libraryStatusCounts
            ? Number(libraryStatusCounts.statusCounts[selectedKmStatus] || 0)
            : libraryStatusCounts
            ? KM_REVIEW_ALL_STATUS_FILTERS.reduce(
              (total, filter) => total + Number(libraryStatusCounts.statusCounts[filter] || 0),
              0
            )
            : kmCorrectedCount <= DOCUMENTS_PER_PAGE
            ? Math.max(kmCorrectedCount, finalDocuments.length)
            : kmCorrectedCount
          : isRecentlyPublishedScope && libraryStatusCounts
          ? Number(libraryStatusCounts.statusCounts.Active || 0)
          : Math.max(restCount + adjustedRestActiveNotIndexedCount - adjustedRestInactiveButIndexedActiveCount, activeCountFloor);
        searchStartOffsetAdjustmentRef.current = isKmReviewHubScope
          ? 0
          : Math.max(0, restCountBase - restCount);
        console.log('[VIEW ALL COUNT] Search crawled Active:', searchTotalRows);
        console.log('[VIEW ALL COUNT] REST uncrawled/new:', adjustedRestActiveNotIndexedCount);
        console.log('[VIEW ALL COUNT] Total shown to user:', restCountBase || searchTotalRows);
        if (isKmReviewHubScope) {
          console.log('[KM REVIEW HUB COUNT] Search non-Active crawled:', searchTotalRows);
          console.log('[KM REVIEW HUB COUNT] REST latest window:', latestIds.length);
          console.log('[KM REVIEW HUB COUNT] Total shown to user:', restCountBase);
        }
        if (restCountBase > 0) {
          setTotalCount(restCountBase);
        } else {
          setTotalCount(searchTotalRows);
        }
        setCountLoading(false);
        setDocuments(finalDocuments);
        documentsFetchInitializedRef.current = true;
        setLoading(false);
        return;
      }
      if (isBookmarksScope) {
        const searchService = getSearchServiceInstance(props.context.spHttpClient, webUrl, libraryName);
        searchService.setFieldMapCache(fieldMap);
        const initialBookmarkIds = bookmarkedIdsArrayRef.current.slice(0, DOCUMENTS_PER_PAGE);
        const rawItems = await searchService.getDocumentDetailsByIds(initialBookmarkIds, queryParts);
        console.log('[BOOKMARKS COUNT] Items fetched from REST:', rawItems?.length ?? 0);
        const items = applyScopeFilters(rawItems).filter((item: any) =>
          String(item[fieldMap.status || COLUMN_NAMES.status] || item.Status || '').trim().toLowerCase() === 'active'
        );
        const processedDocuments: DocumentItem[] = mapDocuments(items, {});
        const seenDocumentIds = new Set<number>();
        const uniqueDocuments = processedDocuments.filter((document) => {
          if (seenDocumentIds.has(document.id)) {
            return false;
          }
          seenDocumentIds.add(document.id);
          return true;
        });
        uniqueDocuments.sort((a, b) => (b.createdTime || 0) - (a.createdTime || 0));
        bookmarkedIdsArrayRef.current = uniqueDocuments.map((document) => document.id);
        pageUrlMapRef.current = new Map();
        setNextPageUrl(null);
        setCurrentRestPage(1);
        setCurrentPage(1);
        setTotalCount(uniqueDocuments.length);
        setCountLoading(false);
        setDocuments(uniqueDocuments.slice(0, DOCUMENTS_PER_PAGE));
        documentsFetchInitializedRef.current = true;
        setLoading(false);
        return;
      }
      const fetchCountWithRetry = async (url: string, retries = 3): Promise<number> => {
        for (let attempt = 0; attempt <= retries; attempt++) {
          try {
            let pageUrl = url;
            let count = 0;
            while (pageUrl) {
              const response = await props.context!.spHttpClient.get(
                pageUrl,
                SPHttpClient.configurations.v1
              );
              if (response.status === 429) {
                const retryAfter = parseInt(response.headers.get('Retry-After') || '5', 10);
                await new Promise((resolve) => setTimeout(resolve, retryAfter * 1000));
                throw new Error('Retry count request after throttling');
              }
              if (!response.ok) {
                return 0;
              }
              const data = await response.json();
              count += (data.value || data?.d?.results || []).length;
              pageUrl = data['@odata.nextLink'] || data?.d?.__next || '';
            }
            return count;
          } catch {
            if (attempt === retries) {
              return 0;
            }
            await new Promise((resolve) => setTimeout(resolve, 2000));
          }
        }
        return 0;
      };
      let countPromise: Promise<number>;
      if (isMyDocScope && myDocAccess) {
        const searchService = getSearchServiceInstance(props.context.spHttpClient, webUrl, libraryName);
        const myDocumentsStatusFilter = statusFilter !== 'all' ? statusFilter : MY_DOCUMENT_ALL_STATUS_FILTERS;
        countPromise = searchService.getListItemCount(
          myDocumentsStatusFilter,
          undefined,
          myDocAccess.id,
          fieldMap.author || COLUMN_NAMES.author
        );
      } else if (isMyDocScope) {
        countPromise = Promise.resolve(0);
      } else if (isBookmarksScope) {
        countPromise = Promise.resolve(bookmarkedIds.size);
      } else if (isRecentlyPublishedScope) {
        countPromise = fetchCountWithRetry(
          `${webUrl}/_api/web/lists/getbytitle('${libraryName}')/items` +
          `?$select=ID` +
          `&$filter=${encodeURIComponent(`${fieldMap.status} eq 'Active'`)}` +
          `&$top=500`
        );
      } else {
        countPromise = fetchCountWithRetry(
          `${webUrl}/_api/web/lists/getbytitle('${libraryName}')/items` +
          `?$select=ID` +
          `&$filter=${encodeURIComponent(buildStatusFilterExpression(fieldMap.status, KM_REVIEW_ALL_STATUS_FILTERS))}` +
          `&$top=500`
        );
      }
      const firstPageResp = await props.context.spHttpClient.get(firstPageUrl, SPHttpClient.configurations.v1);

      if (!firstPageResp.ok) {
        throw new Error(`Failed to fetch first documents page: ${firstPageResp.status}`);
      }

      const firstPageData = await firstPageResp.json();
      const firstPageItems = firstPageData.value || firstPageData?.d?.results || [];
      const items = applyScopeFilters(firstPageItems);

      const processedDocuments: DocumentItem[] = mapDocuments(items, {});
      const seenDocumentIds = new Set<number>();
      const uniqueDocuments = processedDocuments.filter((document) => {
        if (seenDocumentIds.has(document.id)) {
          return false;
        }
        seenDocumentIds.add(document.id);
        return true;
      });
      uniqueDocuments.sort((a, b) => b.id - a.id);

      const initialRestCount = isBookmarksScope ? bookmarkedIds.size : uniqueDocuments.length;
      const next = firstPageData['@odata.nextLink'] || firstPageData?.d?.__next || null;
      pageUrlMapRef.current = new Map();
      if (next) {
        pageUrlMapRef.current.set(2, next);
      }
      setNextPageUrl(next);
      setCurrentRestPage(1);
      setCurrentPage(1);
      setTotalCount(initialRestCount);
      setCountLoading(true);
      setDocuments(uniqueDocuments);
      documentsFetchInitializedRef.current = true;
      setLoading(false);
      void countPromise.then((countResult) => {
        setTotalCount(countResult);
      }).finally(() => {
        setCountLoading(false);
      });
    } catch (error) {
      console.error('Error fetching documents:', error);
      setDocuments([]);
      setLoading(false);
    }
  };

  const fetchPage = async (pageNum: number, pageUrl: string): Promise<void> => {
    if (!props.context || !documentsFetchContextRef.current) return;
    setLoading(true);
    try {
      const resp = await props.context.spHttpClient.get(pageUrl, SPHttpClient.configurations.v1);
      if (!resp.ok) return;
      const data = await resp.json();
      const items = data.value || data?.d?.results || [];
      const filtered = documentsFetchContextRef.current.applyScopeFilters(items);
      const docs = documentsFetchContextRef.current.mapDocuments(filtered, {});
      const seenDocumentIds = new Set<number>();
      const uniqueDocuments = docs.filter((document) => {
        if (seenDocumentIds.has(document.id)) {
          return false;
        }
        seenDocumentIds.add(document.id);
        return true;
      });
      uniqueDocuments.sort((a, b) => b.id - a.id);
      setDocuments(uniqueDocuments);
      setCurrentRestPage(pageNum);
      setCurrentPage(pageNum);
      const next = data['@odata.nextLink'] || data?.d?.__next || null;
      setNextPageUrl(next);
      if (next) pageUrlMapRef.current.set(pageNum + 1, next);
    } catch (err) {
      console.error('fetchPage error:', err);
    } finally {
      setLoading(false);
    }
  };

  const fetchBookmarkPage = async (pageNum: number): Promise<void> => {
    if (!props.context || !documentsFetchContextRef.current) return;
    setLoading(true);
    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const start = (pageNum - 1) * DOCUMENTS_PER_PAGE;
      const chunk = bookmarkedIdsArrayRef.current.slice(start, start + DOCUMENTS_PER_PAGE);
      if (chunk.length === 0) return;

      const idFilter = chunk.map((documentId) => `ID eq ${documentId}`).join(' or ');
      let fieldMap = fieldMapModuleCache || getCachedFieldMap<any>();
      if (!fieldMap) {
        fieldMap = await fetchKMDataHubReadFieldMap(props.context.spHttpClient, webUrl, LIBRARY_NAME);
        setCachedFieldMap(fieldMap);
      }
      fieldMapModuleCache = fieldMap;
      const queryParts = documentsFetchContextRef.current.queryParts || buildKMDataHubItemQuery(
        fieldMap,
        [
          'ID',
          fieldMap.title || COLUMN_NAMES.title,
          fieldMap.description || COLUMN_NAMES.description,
          fieldMap.docIcon || COLUMN_NAMES.docIcon,
          COLUMN_NAMES.created,
          fieldMap.published,
          COLUMN_NAMES.authorId,
          'Author/Title',
          'Author0/Title',
          'Author0Id',
          fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef,
          fieldMap.fileRef || COLUMN_NAMES.fileRef,
          'File/UniqueId',
          fieldMap.url || COLUMN_NAMES.url,
          COLUMN_NAMES.version,
          fieldMap.status
        ],
        ['Author', 'Author0', 'File']
      );
      const pageUrl =
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items` +
        `?$select=${queryParts.select}` +
        `&$expand=${queryParts.expand}` +
        `&$filter=${encodeURIComponent(idFilter)}` +
        `&$top=${DOCUMENTS_PER_PAGE}`;

      const resp = await props.context.spHttpClient.get(pageUrl, SPHttpClient.configurations.v1);
      if (!resp.ok) return;
      const data = await resp.json();
      const items = data.value || data?.d?.results || [];
      const docs = documentsFetchContextRef.current.mapDocuments(items, {});
      const chunkOrder = new Map(chunk.map((documentId, index) => [documentId, index]));
      docs.sort((firstDocument, secondDocument) =>
        (chunkOrder.get(firstDocument.id) ?? Number.MAX_SAFE_INTEGER) -
        (chunkOrder.get(secondDocument.id) ?? Number.MAX_SAFE_INTEGER)
      );
      setDocuments(docs);
      setCurrentRestPage(pageNum);
      setCurrentPage(pageNum);
    } catch (err) {
      console.error('fetchBookmarkPage error:', err);
    } finally {
      setLoading(false);
    }
  };

  const fetchHybridPage = async (pageNum: number): Promise<void> => {
    if (!props.context || !documentsFetchContextRef.current) return;
      setLoading(true);
    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const searchService = getSearchServiceInstance(props.context.spHttpClient, webUrl, LIBRARY_NAMES.kmDataHub);
      const queryParts = documentsFetchContextRef.current.queryParts;
      if (!queryParts) return;
      let mergedIds: number[] = [];
      let myDocAccessForFilter: IMyDocAccess | null = null;
      let myDocAuthorField = COLUMN_NAMES.author;
      const selectedKmStatus = isKmReviewHubScope && statusFilter !== 'allNonActive'
        ? statusFilter
        : '';

      if (isMyDocScope) {
        let fieldMap = fieldMapModuleCache || getCachedFieldMap<any>();
        if (!fieldMap) {
          fieldMap = await fetchKMDataHubReadFieldMap(props.context.spHttpClient, webUrl, LIBRARY_NAME);
          setCachedFieldMap(fieldMap);
        }
        fieldMapModuleCache = fieldMap;
        searchService.setFieldMapCache(fieldMap);
        const myDocAccess = await resolveMyDocAccess();
        if (!myDocAccess) return;
        const authorField = fieldMap.author || COLUMN_NAMES.author;
        myDocAccessForFilter = myDocAccess;
        myDocAuthorField = authorField;
        const searchSortDirection = effectiveRestSortDirection === 'asc' ? 'ascending' : 'descending';
        const myDocumentsStatusFilter = statusFilter !== 'all' ? statusFilter : MY_DOCUMENT_ALL_STATUS_FILTERS;
        const [{ ids: authorIds }, author0Ids, latestIds] = await Promise.all([
          searchService.getDocumentIdsByAuthor(
            myDocAccess.id,
            authorField,
            (pageNum - 1) * DOCUMENTS_PER_PAGE,
            DOCUMENTS_PER_PAGE,
            myDocumentsStatusFilter,
            effectiveRestSortDirection
          ),
          searchService.getDocumentIdsByAuthor0Search(
            myDocAccess.email,
            searchSortDirection,
            (pageNum - 1) * DOCUMENTS_PER_PAGE,
            DOCUMENTS_PER_PAGE,
            myDocumentsStatusFilter
          ),
          pageNum === 1 ? searchService.getLatestDocumentIds(DOCUMENTS_PER_PAGE + 10) : Promise.resolve([])
        ]);
        mergedIds = Array.from(new Set([...authorIds, ...author0Ids, ...latestIds])).slice(0, DOCUMENTS_PER_PAGE + 10);
      } else {
        const searchStatus = isRecentlyPublishedScope
          ? 'Active'
          : isKmReviewHubScope
          ? (selectedKmStatus || KM_REVIEW_ALL_STATUS_FILTERS)
          : 'Active';
        const sortField = isRecentlyPublishedScope
          ? SEARCH_PROPERTIES.published
          : isKmReviewHubScope
          ? SEARCH_PROPERTIES.write
          : SEARCH_PROPERTIES.published;
        const searchRowLimit = isKmReviewHubScope
          ? DOCUMENTS_PER_PAGE
          : DOCUMENTS_PER_PAGE + FETCH_BUFFER;
        const searchStartRow = isKmReviewHubScope
          ? (pageNum - 1) * DOCUMENTS_PER_PAGE
          : Math.max(0, (pageNum - 1) * DOCUMENTS_PER_PAGE - searchStartOffsetAdjustmentRef.current);
        const [searchResult, latestIds] = await Promise.all([
          searchService.getDocumentIdsByStatus(
            searchStatus,
            searchStartRow,
            searchRowLimit,
            sortField,
            effectiveSearchSortDirection
          ),
          pageNum === 1 ? searchService.getLatestDocumentIds(DOCUMENTS_PER_PAGE + 10) : Promise.resolve([])
        ]);
        let searchIds = searchResult.ids;
        if (isKmReviewHubScope && searchIds.length === 0) {
          const restFallbackResult = await searchService.getDocumentIdsByStatusRest(
            searchStatus,
            (pageNum - 1) * DOCUMENTS_PER_PAGE,
            searchRowLimit,
            documentsFetchContextRef.current.statusField || COLUMN_NAMES.status,
            COLUMN_NAMES.created,
            effectiveRestSortDirection
          );
          searchIds = restFallbackResult.ids;
        }
        if (searchIds.length === 0 && pageNum > 1 && searchResult.totalRows > 0) {
          const fallbackStartRow = Math.max(0, searchResult.totalRows - DOCUMENTS_PER_PAGE);
          if (fallbackStartRow !== searchStartRow) {
            const fallbackResult = await searchService.getDocumentIdsByStatus(
              searchStatus,
              fallbackStartRow,
              DOCUMENTS_PER_PAGE,
              sortField,
              effectiveSearchSortDirection
            );
            searchIds = fallbackResult.ids;
          }
        }
        mergedIds = Array.from(new Set([...searchIds, ...latestIds])).slice(0, DOCUMENTS_PER_PAGE + 10);
      }

      const statusField = documentsFetchContextRef.current.statusField || COLUMN_NAMES.status;
      const rawItems = await searchService.getDocumentDetailsByIds(mergedIds, queryParts);
      const filteredItems = documentsFetchContextRef.current.applyScopeFilters(rawItems).filter((item: any) => {
        const status = String(item[statusField] || item.Status || '').trim();
        if (isMyDocScope) {
          return myDocAccessForFilter
            ? documentMatchesMyDocAccess(item, myDocAccessForFilter, myDocAuthorField) &&
              documentMatchesMyDocumentsStatusFilter(status, statusFilter)
            : false;
        }
        if (isKmReviewHubScope) {
          return documentMatchesKmReviewStatusFilter(status, statusFilter);
        }
        return status === 'Active';
      });
      const docs = documentsFetchContextRef.current.mapDocuments(filteredItems, {});
      const seenDocumentIds = new Set<number>();
      const uniqueDocuments = docs.filter((document) => {
        if (seenDocumentIds.has(document.id)) {
          return false;
        }
        seenDocumentIds.add(document.id);
        return true;
      });
      uniqueDocuments.sort((a, b) => {
        if (isRecentlyPublishedScope || isKmReviewHubScope || (!isMyDocScope && !isBookmarksScope)) {
          const dateA = a.createdTime || new Date(a.date || 0).getTime();
          const dateB = b.createdTime || new Date(b.date || 0).getTime();
          return effectiveSearchSortDirection === 'ascending' ? dateA - dateB : dateB - dateA;
        }
        if (isMyDocScope) {
          const dateA = a.createdTime || new Date(a.date || 0).getTime();
          const dateB = b.createdTime || new Date(b.date || 0).getTime();
          return effectiveRestSortDirection === 'asc' ? dateA - dateB : dateB - dateA;
        }
        return effectiveRestSortDirection === 'asc' ? a.id - b.id : b.id - a.id;
      });
      const finalDocuments = isKmReviewHubScope
        ? uniqueDocuments
        : uniqueDocuments.slice(0, DOCUMENTS_PER_PAGE);
      setDocuments(finalDocuments);
      setCurrentRestPage(pageNum);
      setCurrentPage(pageNum);
      setNextPageUrl(null);
    } catch (err) {
      console.error('fetchHybridPage error:', err);
    } finally {
      setLoading(false);
    }
  };

  const handlePageChange = async (newPage: number): Promise<void> => {
    if (newPage === currentRestPage) return;
    if (isBookmarksScope) {
      await fetchBookmarkPage(newPage);
      return;
    }

    if (!isBookmarksScope) {
      await fetchHybridPage(newPage);
      return;
    }

    let targetUrl = pageUrlMapRef.current.get(newPage);
    if (targetUrl) {
      await fetchPage(newPage, targetUrl);
      return;
    }

    if (newPage > currentRestPage) {
      let pageToFetch = currentRestPage + 1;
      targetUrl = nextPageUrl;
      while (targetUrl && pageToFetch <= newPage) {
        await fetchPage(pageToFetch, targetUrl);
        targetUrl = pageUrlMapRef.current.get(pageToFetch + 1) || null;
        pageToFetch += 1;
      }
    }
  };

  React.useEffect(() => {
    return subscribeToDocumentDataChanged(() => {
      if (props.context) {
        setCurrentPage(1);
        void fetchAllDocuments();
      }
    });
  }, [props.context]);

  const getFileTypeIcon = (fileType: string): string => {
    const type = (fileType || '').toLowerCase();
    if (type === 'pdf') return '📄';
    if (type === 'pptx' || type === 'ppt') return '📊';
    if (type === 'docx' || type === 'doc') return '📝';
    if (type === 'xlsx' || type === 'xls') return '📈';
    if (['mp4', 'mov', 'webm', 'mkv', 'avi', 'm4v', 'wmv'].indexOf(type) !== -1 || type.indexOf('video') !== -1) return '▶';
    if (['mp3', 'wav', 'm4a', 'aac', 'flac', 'ogg'].indexOf(type) !== -1 || type.indexOf('audio') !== -1) return '♪';
    return '📎';
  };

  const isExcelFileType = (fileType: string): boolean =>
    ['xls', 'xlsx', 'xlsm', 'xlsb'].indexOf((fileType || '').toLowerCase()) !== -1;

  const isAudioFileType = (fileType: string): boolean =>
    ['mp3', 'wav', 'm4a', 'aac', 'flac', 'ogg', 'wma'].indexOf((fileType || '').toLowerCase()) !== -1;

  const getThumbnailCandidates = React.useCallback((serverRelativeUrl: string, fileType: string, fileUniqueId?: string): ThumbnailCandidate[] => {
    if (!props.context || !serverRelativeUrl) return [];
    const webUrl = props.context.pageContext.web.absoluteUrl;
    const normalizedServerRelativeUrl = serverRelativeUrl.startsWith('/') ? serverRelativeUrl : `/${serverRelativeUrl}`;
    const origin = new URL(webUrl).origin;
    const absoluteFileUrl = `${origin}${normalizedServerRelativeUrl}`;
    const normalizedFileType = fileType.toLowerCase();
    if (isAudioFileType(normalizedFileType)) return [];
    const isVideoFile = ['mp4', 'mov', 'avi', 'wmv'].indexOf(normalizedFileType) !== -1;
    const isExcelFile = ['xls', 'xlsx', 'xlsm', 'xlsb'].indexOf(normalizedFileType) !== -1;
    const candidates: ThumbnailCandidate[] = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'].indexOf(normalizedFileType) !== -1 ? [absoluteFileUrl] : [];

    if (fileUniqueId) {
      const siteId = props.context.pageContext.site.id.toString();
      const webId = props.context.pageContext.web.id.toString();
      const guidPreviewBase = `${webUrl}/_layouts/15/getpreview.ashx?guidSite=${encodeURIComponent(siteId)}&guidWeb=${encodeURIComponent(webId)}&guidFile=${encodeURIComponent(fileUniqueId)}`;
      candidates.push(
        `${guidPreviewBase}&resolution=6&index=0&force=1&size=large`,
        `${guidPreviewBase}&resolution=3&index=0&force=1`
      );
    }

    if (!isExcelFile) {
      candidates.push(`${webUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(absoluteFileUrl)}&resolution=2&index=1&force=1`);
    }
    if (isVideoFile) {
      candidates.push({ isVideo: true, posterUrl: null, sourceUrl: absoluteFileUrl, useVideoElement: true });
    }
    return candidates.slice(0, isExcelFile ? 3 : 3);
  }, [props.context]);

  const handleView = (item: DocumentItem) => {
    props.onViewDocument(item.id, []);
  };

  const handleDownload = async (item: DocumentItem) => {
    if (props.context && item.serverRelativeUrl) {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      let serverRelativeUrl = item.serverRelativeUrl;
      if (!serverRelativeUrl.startsWith('/')) serverRelativeUrl = `/${serverRelativeUrl}`;
      if (downloadingDocumentIds[item.id]) return;
      setDownloadingDocumentIds((current) => ({ ...current, [item.id]: true }));
      try {
        await downloadSharePointFile(
          props.context,
          webUrl,
          serverRelativeUrl,
          buildTitleDownloadFileName(item.name, item.fileName, item.fileType)
        );
        void recordDownloadMetric(item.id, item.name || item.fileName)
          .then(() => props.onDownloadRecorded?.(item.id))
          .catch(() => undefined);
      } catch (error) {
        console.error('Download error:', error);
      } finally {
        window.setTimeout(() => {
          setDownloadingDocumentIds((current) => {
            const next = { ...current };
            delete next[item.id];
            return next;
          });
        }, 1200);
      }
    }
  };

  const visibleDocuments = React.useMemo(() => {
    return isBookmarksScope
      ? documents.filter((doc) => doc.status.trim().toLowerCase() === 'active')
      : isRecentlyPublishedScope || (!isKmReviewHubScope && !isMyDocScope)
      ? documents
      : documents.filter((doc) => isKmReviewHubScope
        ? documentMatchesKmReviewStatusFilter(doc.status, statusFilter)
        : documentMatchesMyDocumentsStatusFilter(doc.status, statusFilter)
      );
  }, [documents, isBookmarksScope, isKmReviewHubScope, isMyDocScope, isRecentlyPublishedScope, statusFilter]);

  const totalPages = Math.max(1, Math.ceil(totalCount / DOCUMENTS_PER_PAGE));
  const isFilterActive = !isRecentlyPublishedScope
    && !isBookmarksScope
    && statusFilter !== 'all'
    && statusFilter !== 'allNonActive'
    && statusFilter !== 'Active';
  const displayCount = isFilterActive ? visibleDocuments.length : totalCount;
  const displayStart = displayCount > 0 ? (currentRestPage - 1) * DOCUMENTS_PER_PAGE + 1 : 0;
  const displayEnd = isFilterActive
    ? visibleDocuments.length
    : Math.min(currentRestPage * DOCUMENTS_PER_PAGE, totalCount);
  const statusFilterOptions = isRecentlyPublishedScope || isBookmarksScope
    ? ACTIVE_ONLY_STATUS_FILTER_OPTIONS
    : isMyDocScope
    ? MY_DOCUMENT_STATUS_FILTER_OPTIONS
    : KM_REVIEW_STATUS_FILTER_OPTIONS;
  const selectedSortLabel = SORT_OPTIONS.find((option) => option.value === sortOrder)?.label || 'Newest first';
  const selectedStatusLabel = statusFilterOptions.find((option) => option.value === statusFilter)?.label || 'All Status';
  const paginationItems = React.useMemo(() => buildPaginationItems(currentRestPage, totalPages), [currentRestPage, totalPages]);
  const paginatedDocuments = React.useMemo(() => {
    return visibleDocuments;
  }, [visibleDocuments]);

  React.useEffect(() => {
    if (currentRestPage > totalPages) {
      setCurrentRestPage(totalPages);
      setCurrentPage(totalPages);
    }
  }, [currentRestPage, totalPages]);

  return (
    <div className={styles.viewAllPage}>
      <IKShellHeader
        userName={props.context?.pageContext.user.displayName}
        userEmail={props.context?.pageContext.user.email || props.context?.pageContext.user.loginName}
        userPhotoUrl={props.userPhotoUrl}
        onLogout={props.onLogout}
        onHomeOpen={props.onHomeOpen}
        onAllDocumentsOpen={props.onAllDocumentsOpen}
        onBusinessUnitsOpen={props.onBusinessUnitsOpen}
        onBookmarksOpen={props.onBookmarksOpen}
        onDocumentsOpen={props.onDocumentsOpen}
        onContactOpen={props.onContactOpen}
        onAuditLogOpen={props.onAuditLogOpen}
        onAnalyticsOpen={props.onAnalyticsOpen}
        showReviewerNav={props.showReviewerNav}
        hideDocumentsNav={props.hideDocumentsNav}
        activeNavKey={activeHeaderNavKey}
      />

      <div className={styles.mainContentArea}>
        <div className={styles.pageHeadingHeader}>
          <div className={styles.pageHeadingText}>
            <h1 className={styles.pageHeadingTitle}>{pageHeading}</h1>
            {pageTagline && (
              <p className={styles.pageHeadingTagline}>{pageTagline}</p>
            )}
          </div>
        </div>

        <div className={styles.contentHeader}>
          <div className={styles.headerControlsContainer}>
            <div className={styles.controlGroup}>
              <div className={styles.controlLabel}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M3 6h18M6 12h12M10 18h4" />
                </svg>
                Sort
              </div>
              <div className={styles.selectWrapper}>
                <button
                  type="button"
                  className={`${styles.customSelect} ${openControlDropdown === 'sort' ? styles.customSelectOpen : ''}`}
                  onClick={() => setOpenControlDropdown(openControlDropdown === 'sort' ? null : 'sort')}
                  aria-haspopup="listbox"
                  aria-expanded={openControlDropdown === 'sort'}
                >
                  <span className={styles.customSelectValue}>{selectedSortLabel}</span>
                  <svg className={styles.customSelectChevron} width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                    <path d="M6 9l6 6 6-6" />
                  </svg>
                </button>
                {openControlDropdown === 'sort' && (
                  <div className={styles.customDropdown} role="listbox">
                    {SORT_OPTIONS.map((option) => (
                      <button
                        key={option.value}
                        type="button"
                        role="option"
                        aria-selected={sortOrder === option.value}
                        className={`${styles.customDropdownOption} ${sortOrder === option.value ? styles.customDropdownOptionActive : ''}`}
                        onClick={() => {
                          setSortOrder(option.value);
                          setOpenControlDropdown(null);
                        }}
                      >
                        {option.label}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </div>

            {!isRecentlyPublishedScope && !isBookmarksScope && (
              <>
                <div className={styles.controlDivider} />

                <div className={styles.controlGroup}>
                  <div className={styles.controlLabel}>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                      <path d="M22 3H2l8 9.46V19l4 2v-8.54L22 3z" />
                    </svg>
                    Status
                  </div>
                  <div className={styles.selectWrapper}>
                    <button
                      type="button"
                      className={`${styles.customSelect} ${openControlDropdown === 'status' ? styles.customSelectOpen : ''}`}
                      onClick={() => setOpenControlDropdown(openControlDropdown === 'status' ? null : 'status')}
                      aria-haspopup="listbox"
                      aria-expanded={openControlDropdown === 'status'}
                    >
                      <span className={styles.customSelectValue}>{selectedStatusLabel}</span>
                      <svg className={styles.customSelectChevron} width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
                        <path d="M6 9l6 6 6-6" />
                      </svg>
                    </button>
                    {openControlDropdown === 'status' && (
                      <div className={styles.customDropdown} role="listbox">
                        {statusFilterOptions.map((opt) => (
                          <button
                            key={opt.value}
                            type="button"
                            role="option"
                            aria-selected={statusFilter === opt.value}
                            className={`${styles.customDropdownOption} ${statusFilter === opt.value ? styles.customDropdownOptionActive : ''}`}
                            onClick={() => {
                              setStatusFilter(opt.value);
                              setOpenControlDropdown(null);
                            }}
                          >
                            {opt.label}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
                <div className={styles.controlDivider} />
              </>
            )}

            <div className={styles.viewModeToggles}>
              <button
                className={`${styles.viewModeBtn} ${viewMode === 'list' ? styles.viewModeBtnActive : ''}`}
                onClick={() => setViewMode('list')}
                title="List View"
              >
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" />
                </svg>
              </button>
              <button
                className={`${styles.viewModeBtn} ${viewMode === 'grid' ? styles.viewModeBtnActive : ''}`}
                onClick={() => setViewMode('grid')}
                title="Grid View"
              >
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                  <rect x="3" y="3" width="7" height="7" /><rect x="14" y="3" width="7" height="7" />
                  <rect x="14" y="14" width="7" height="7" /><rect x="3" y="14" width="7" height="7" />
                </svg>
              </button>
            </div>
          </div>
        </div>

        <div className={styles.paginationHeader}>
          <span>
            {loading
              ? 'Loading matches...'
              : countLoading
              ? <><span className={styles.countSpinner} aria-hidden="true" /> Counting matches...</>
              : displayCount > 0
              ? `Showing ${displayStart}-${displayEnd} of ${displayCount} matching assets`
              : 'Showing 0 matching assets'}
          </span>
        </div>

        <div className={styles.scrollContent}>
         {loading ? (
  <div className={styles.loadingState}>
    <style>{`
      @keyframes ikShimmer {
        0% { background-position: 200% 0; }
        100% { background-position: -200% 0; }
      }
      .ik-skel {
        background: linear-gradient(90deg, #f0f0f0 25%, #e8e8e8 50%, #f0f0f0 75%);
        background-size: 200% 100%;
        animation: ikShimmer 1.5s infinite;
        border-radius: 4px;
      }
    `}</style>
    {[0,1,2,3,4,5].map((i) => (
      <div key={i} style={{display:'flex', gap:'16px', padding:'16px 0', borderBottom:'1px solid #f0f0f0', alignItems:'center'}}>
        <div style={{flex:1}}>
          <div className="ik-skel" style={{height:'18px', width:'60%', marginBottom:'8px'}} />
          <div className="ik-skel" style={{height:'14px', width:'85%', marginBottom:'6px'}} />
          <div className="ik-skel" style={{height:'14px', width:'40%'}} />
        </div>
        <div style={{display:'flex', flexDirection:'column', gap:'8px', alignItems:'flex-end', flexShrink:0}}>
          <div className="ik-skel" style={{width:'80px', height:'14px'}} />
          <div className="ik-skel" style={{width:'60px', height:'14px'}} />
        </div>
        <div style={{display:'flex', gap:'8px', flexShrink:0}}>
          <div className="ik-skel" style={{width:'64px', height:'32px', borderRadius:'4px'}} />
          <div className="ik-skel" style={{width:'80px', height:'32px', borderRadius:'4px'}} />
        </div>
      </div>
    ))}
  </div>
) : visibleDocuments.length === 0 ? (
            <div className={styles.emptyState}>No documents found matching your criteria.</div>
          ) : (
            <div className={viewMode === 'grid' ? styles.documentGrid : styles.documentList}>
              {paginatedDocuments.map((doc) => (
                <div key={doc.id} className={viewMode === 'grid' ? styles.documentGridCard : styles.documentRow}>
                  {viewMode === 'grid' ? (
                    <>
                      <div className={styles.gridThumbnailSection}>
                        {(() => {
                          const thumbnailCandidates = getThumbnailCandidates(doc.serverRelativeUrl, doc.fileType, doc.fileUniqueId);
                          const activeThumbnailIndex = thumbnailAttemptByDocument[doc.id] ?? 0;
                          const hasThumbnailFailed = activeThumbnailIndex < 0;
                          const fallbackThumbnailIndex = hasThumbnailFailed ? 0 : activeThumbnailIndex;
                          const activeThumbnailCandidate = hasThumbnailFailed ? undefined : thumbnailCandidates[fallbackThumbnailIndex];
                          const isVideoThumbnail = typeof activeThumbnailCandidate !== 'string' && !!activeThumbnailCandidate?.isVideo;
                          const shouldRenderVideoElement = isVideoThumbnail && !!activeThumbnailCandidate?.useVideoElement;
                          const activeThumbnailUrl = typeof activeThumbnailCandidate === 'string' ? activeThumbnailCandidate : '';
                          const videoThumbnailSource = typeof activeThumbnailCandidate !== 'string' ? activeThumbnailCandidate?.sourceUrl || doc.serverRelativeUrl : '';
                          const isExcelPreview = isExcelFileType(doc.fileType);
                          const isAudioPreview = isAudioFileType(doc.fileType);
                          const excelPreviewUrl = isExcelPreview && doc.fileUniqueId
                            ? getExcelThumbnailUrl(props.context.pageContext.web.absoluteUrl, KM_REVIEW_HUB_DRIVE_ID, doc.fileUniqueId)
                            : '';
                          const drivePreviewUrl = !isExcelPreview && !isAudioPreview && doc.fileUniqueId && activeThumbnailIndex >= 0
                            ? getDriveItemThumbnailUrl(props.context.pageContext.web.absoluteUrl, KM_REVIEW_HUB_DRIVE_ID, doc.fileUniqueId)
                            : '';
                          const shouldShowFallback = !isExcelPreview && !isAudioPreview && !isVideoThumbnail && (activeThumbnailIndex < 0 || !activeThumbnailUrl);
                          const usesDocumentPreviewCrop = ['PDF', 'DOC', 'DOCX', 'PPT', 'PPTX', 'XLS', 'XLSX'].indexOf(doc.fileType.toUpperCase()) !== -1;
                          const isPresentationPreview = ['PPT', 'PPTX'].indexOf(doc.fileType.toUpperCase()) !== -1;
                          const isWordPreview = ['DOC', 'DOCX'].indexOf(doc.fileType.toUpperCase()) !== -1;

                          return (
                            <>
                              {isAudioPreview ? (
                                <div className={styles.audioThumbnailContainer}>
                                  <div className={styles.audioThumbnailIcon}>
                                    <span />
                                    <span />
                                    <span />
                                    <span />
                                  </div>
                                  <span className={styles.audioThumbnailLabel}>AUDIO</span>
                                </div>
                              ) : isExcelPreview && excelPreviewUrl ? (
                                <img
                                  src={excelPreviewUrl}
                                  alt={doc.name}
                                  loading="eager"
                                  decoding="async"
                                  style={{
                                    position: 'absolute',
                                    top: '0',
                                    left: '0',
                                    width: '100%',
                                    height: '100%',
                                    objectFit: 'cover',
                                    objectPosition: 'top left',
                                    display: 'block',
                                    transform: 'scale(3.5)',
                                    transformOrigin: 'top left',
                                    imageRendering: 'auto',
                                    filter: 'contrast(1.04) saturate(1.02)'
                                  }}
                                  onError={(event) => {
                                    const sourceFallbackUrl = getExcelThumbnailUrl(
                                      props.context.pageContext.web.absoluteUrl,
                                      KM_REVIEW_HUB_DRIVE_ID,
                                      doc.fileUniqueId || '',
                                      'source'
                                    );
                                    const largeFallbackUrl = getExcelThumbnailUrl(
                                      props.context.pageContext.web.absoluteUrl,
                                      KM_REVIEW_HUB_DRIVE_ID,
                                      doc.fileUniqueId || '',
                                      'large'
                                    );
                                    if (event.currentTarget.src !== sourceFallbackUrl && event.currentTarget.src !== largeFallbackUrl) {
                                      event.currentTarget.src = sourceFallbackUrl;
                                    } else if (event.currentTarget.src !== largeFallbackUrl) {
                                      event.currentTarget.src = largeFallbackUrl;
                                    } else {
                                      event.currentTarget.style.display = 'none';
                                    }
                                  }}
                                />
                              ) : isExcelPreview ? (
                                <div className={styles.excelPlaceholder}>
                                  <span className={styles.excelPlaceholderLabel}>XLS</span>
                                </div>
                              ) : drivePreviewUrl ? (
                                <img
                                  src={drivePreviewUrl}
                                  alt=""
                                  className={`${styles.thumbnailImage} ${usesDocumentPreviewCrop ? styles.thumbnailDocumentImage : ''}`}
                                  loading="eager"
                                  decoding="async"
                                  style={isPresentationPreview ? {
                                    objectFit: 'contain',
                                    objectPosition: 'center center',
                                    background: '#ffffff'
                                  } : isWordPreview ? {
                                    imageRendering: 'auto',
                                    filter: 'contrast(1.08) saturate(1.03)',
                                    transform: 'translateZ(0)'
                                  } : undefined}
                                  onError={(event) => {
                                    const image = event.currentTarget;
                                    const sourceFallbackUrl = getDriveItemThumbnailUrl(
                                      props.context.pageContext.web.absoluteUrl,
                                      KM_REVIEW_HUB_DRIVE_ID,
                                      doc.fileUniqueId || '',
                                      'source'
                                    );
                                    const largeFallbackUrl = getDriveItemThumbnailUrl(
                                      props.context.pageContext.web.absoluteUrl,
                                      KM_REVIEW_HUB_DRIVE_ID,
                                      doc.fileUniqueId || '',
                                      'large'
                                    );
                                    if (image.src !== sourceFallbackUrl && image.src !== largeFallbackUrl) {
                                      image.src = sourceFallbackUrl;
                                    } else if (image.src !== largeFallbackUrl) {
                                      image.src = largeFallbackUrl;
                                    } else {
                                      setThumbnailAttemptByDocument(prev => ({ ...prev, [doc.id]: -1 }));
                                    }
                                    event.preventDefault();
                                  }}
                                />
                              ) : shouldRenderVideoElement ? (
                                <video
                                  className={styles.thumbnailImage}
                                  preload="metadata"
                                  muted
                                  playsInline
                                  onLoadedMetadata={(event) => {
                                    const video = event.currentTarget;
                                    video.currentTime = 1;
                                  }}
                                  onSeeked={(event) => {
                                    const video = event.currentTarget;
                                    video.style.opacity = '1';
                                  }}
                                  style={{ opacity: 0, transition: 'opacity 0.3s' }}
                                >
                                  <source src={videoThumbnailSource} type="video/mp4" />
                                </video>
                              ) : isVideoThumbnail ? (
                                <div className={styles.videoThumbnailContainer}>
                                  <div className={styles.videoThumbnailPlaceholder}>
                                    <svg viewBox="0 0 24 24" width="40" height="40">
                                      <circle cx="12" cy="12" r="12" fill="rgba(0,0,0,0.55)" />
                                      <polygon points="9,7 19,12 9,17" fill="white" />
                                    </svg>
                                  </div>
                                </div>
                              ) : activeThumbnailUrl && (
                                <img
                                  src={activeThumbnailUrl}
                                  alt=""
                                  className={`${styles.thumbnailImage} ${usesDocumentPreviewCrop ? styles.thumbnailDocumentImage : ''}`}
                                  loading="eager"
                                  decoding="sync"
                                  style={isPresentationPreview ? {
                                    objectFit: 'contain',
                                    objectPosition: 'center center',
                                    background: '#ffffff'
                                  } : undefined}
                                  onLoad={() => setThumbnailLoadedByDocument(prev => ({ ...prev, [doc.id]: true }))}
                                  onError={(event) => {
                                    const image = event.currentTarget;
                                    const candidates = getThumbnailCandidates(doc.serverRelativeUrl, doc.fileType, doc.fileUniqueId);
                                    const currentIndex = candidates.indexOf(image.src);
                                    const nextIndex = currentIndex >= 0 && currentIndex < candidates.length - 1 ? currentIndex + 1 : fallbackThumbnailIndex + 1;
                                    const nextCandidate = candidates[nextIndex];
                                    if (typeof nextCandidate === 'string') {
                                      image.src = nextCandidate;
                                      setThumbnailAttemptByDocument(prev => ({ ...prev, [doc.id]: nextIndex }));
                                    } else if (nextCandidate?.isVideo) {
                                      setThumbnailAttemptByDocument(prev => ({ ...prev, [doc.id]: nextIndex }));
                                    } else {
                                      image.onerror = null;
                                      setThumbnailAttemptByDocument(prev => ({ ...prev, [doc.id]: -1 }));
                                    }
                                    event.preventDefault();
                                  }}
                                />
                              )}
                              {shouldShowFallback && (
                                <div className={styles.thumbnailFallback}>
                                  <div className={styles.fileTypeIcon}>{getFileTypeIcon(doc.docIcon || doc.fileType)}</div>
                                  <span className={styles.fileTypeLabel}>{doc.docIcon || doc.fileType}</span>
                                </div>
                              )}
                            </>
                          );
                        })()}
                      </div>
                      <div className={styles.gridCardBody}>
                        <div className={styles.gridCardMeta}>
                          {renderAuthorStack(doc.author, styles.authorStack, styles.authorChip)}
                          <span className={styles.sourceType}>{doc.fileType}</span>
                        </div>
                        <h3 className={styles.gridDocumentTitle}>{doc.name}</h3>
                        <p className={styles.searchResultDescription}>{doc.abstract}</p>
                        <div className={styles.gridCardFooter}>
                          <span>{doc.date}</span>
                          <span className={`${styles.statusBadge} ${getStatusBadgeClassName(doc.status)}`}>{getDisplayStatus(doc.status) || 'Active'}</span>
                        </div>
                        <div className={styles.gridActionButtons}>
                          <button className={styles.viewButton} onClick={() => handleView(doc)}>View</button>
                          {!hideDownloadButtonsForLearner && (
                            <button
                              className={styles.downloadButton}
                              disabled={!!downloadingDocumentIds[doc.id]}
                              onClick={() => handleDownload(doc)}
                            >
                              {downloadingDocumentIds[doc.id] ? 'Downloading...' : 'Download'}
                            </button>
                          )}
                        </div>
                      </div>
                    </>
                  ) : (
                    <div className={styles.listContent}>
                      <div className={styles.searchResultContent}>
                        <h3 className={styles.searchResultTitle}>{doc.name}</h3>
                        <p className={styles.searchResultDescription}>{doc.abstract}</p>
                        <div className={styles.searchResultStatusLine}>
                          <span className={styles.searchResultStatusLabel}>Status :</span>
                          <span className={`${styles.statusBadge} ${getStatusBadgeClassName(doc.status)}`}>{getDisplayStatus(doc.status) || 'Active'}</span>
                        </div>
                      </div>
                      <div className={styles.searchResultDivider} aria-hidden="true" />
                      <div className={styles.searchResultMeta}>
                        {renderAuthorStack(doc.author, styles.authorStack, styles.searchResultMetaStrong)}
                        <span className={styles.sourceDate}>{doc.date}</span>
                      </div>
                      <div className={styles.searchResultActions}>
                        <button className={styles.searchResultViewButton} onClick={() => handleView(doc)}>View</button>
                        {!hideDownloadButtonsForLearner && (
                          <button
                            className={styles.searchResultDownloadButton}
                            disabled={!!downloadingDocumentIds[doc.id]}
                            onClick={() => handleDownload(doc)}
                          >
                            {downloadingDocumentIds[doc.id] ? 'Downloading...' : 'Download'}
                          </button>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}

          {!loading && visibleDocuments.length > 0 && (
            <div className={styles.paginationFooter}>
              <button
                disabled={currentRestPage === 1}
                onClick={() => void handlePageChange(currentRestPage - 1)}
                className={styles.pageArrowBtn}
              >
                Previous
              </button>
              <div className={styles.pageNumbers}>
                {paginationItems.map((item) => (
                  typeof item === 'number' ? (
                    <button
                      key={item}
                      onClick={() => void handlePageChange(item)}
                      className={`${styles.pageNumberBtn} ${currentRestPage === item ? styles.pageNumberBtnActive : ''}`}
                    >
                      {item}
                    </button>
                  ) : (
                    <span key={item} className={styles.pageEllipsis}>...</span>
                  )
                ))}
              </div>
              <button
                disabled={currentRestPage === totalPages}
                onClick={() => void handlePageChange(currentRestPage + 1)}
                className={styles.pageArrowBtn}
              >
                Next
              </button>
            </div>
          )}
        </div>
      </div>

      <IKShellFooter className={styles.fixedShellFooter} onBackHome={props.onHomeOpen || (() => { pushPageUrl(NAV_PATHS.home); })} />

      {editingDocument && (
        <div className={styles.editOverlay}>
          <div className={styles.editModal}>
            <MetadataForm
              context={props.context}
              taxonomyOptions={taxonomyOptions!}
              initialValues={editInitialValues!}
              onClose={handleCloseEdit}
              onSubmit={(data) => { void handleEditSubmit(data); }}
            />
          </div>
        </div>
      )}
    </div>
  );
};
