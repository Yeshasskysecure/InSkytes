import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  buildKMDataHubItemQuery,
  fetchKMDataHubReadFieldMap,
  formatDocumentPublishedDate,
  resolveDocumentPublishedValue,
  resolveDocumentAuthor
} from '../../utils/documentMetadata';
import { IKShellFooter, IKShellHeader } from '../../components/shell/IKShellChrome';
import { fetchAllTaxonomyOptions, ITaxonomyFieldOptions, ITaxonomyTerm, TAXONOMY_FIELD_CONFIGS } from '../../services/TaxonomyService';
import { getDriveItemThumbnailUrl, getExcelThumbnailUrl, getSearchServiceInstance } from '../../services/SharePointSearchService';
import { fetchBUDepartmentTerms, Term } from '../../utils/termStore';
import { CACHE_KEYS, COLUMN_NAMES, KM_REVIEW_HUB_DRIVE_ID, LIBRARY_NAMES, PAGE_SIZES, SEARCH_PROPERTIES } from '../../config/appConfig';
import { NAV_PATHS, pushPageUrl } from '../../services/permalinkService';
import { downloadSharePointFile } from '../../utils/fileDownload';
import styles from './CategoryDocumentsPage.module.scss';

const LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;
const DOCUMENTS_PER_PAGE = PAGE_SIZES.documentsPerPage;
const FIELD_MAP_CACHE_KEY = CACHE_KEYS.fieldMap;
const TAXONOMY_CACHE_KEY = CACHE_KEYS.taxonomy;
const FIELDS_CACHE_KEY = CACHE_KEYS.fields;
const BU_TERMS_CACHE_KEY = CACHE_KEYS.buTerms;
let fieldMapModuleCache: any = null;

interface ICategoryDocument {
  id: number;
  title: string;
  description: string;
  documentType: string;
  author: string;
  date: string;
  createdTime: number;
  status: string;
  statusKey: TStatusKey;
  businessUnit: string;
  department: string;
  client: string;
  geography: string;
  therapyArea: string;
  diseaseArea: string;
  fileUrl: string;
  fileName: string;
  fileType: string;
  fileRef: string;
  fileUniqueId?: string;
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

type ThumbnailCandidate = string | { isVideo: true; posterUrl: string | null; sourceUrl?: string; useVideoElement?: boolean };

type TFilterKey =
  'businessUnit' |
  'department' |
  'client' |
  'geography' |
  'therapyArea' |
  'diseaseArea' |
  'reviewerComments' |
  'projectId' |
  'contentRefreshDate' |
  'versionFileType';
type TDocumentViewMode = 'list' | 'grid';
type PaginationItem = number | 'ellipsis-start' | 'ellipsis-end';
interface ISharedFilterOption {
  title: string;
  count: number;
  level?: number;
  parentTitle?: string;
}
interface ISharedFilterGroup {
  key: string;
  title: string;
  children: ISharedFilterOption[];
}
type TStatusKey = 'active' | 'underReview' | 'archive' | 'reject' | 'unknown';
type TFilterRefinerCounts = Partial<Record<TFilterKey, Record<string, number>>>;

interface ICategoryFetchContext {
  mapCategoryDocuments: (items: any[]) => ICategoryDocument[];
  mapSearchRowsToItems: (rows: any[]) => any[];
  queryParts: { select: string; expand: string };
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
const getCachedFields = (): Array<{ Title?: string; InternalName?: string; Id?: string; TextField?: string }> | null =>
  getSessionCache<Array<{ Title?: string; InternalName?: string; Id?: string; TextField?: string }>>(FIELDS_CACHE_KEY);
const setCachedFields = (fields: Array<{ Title?: string; InternalName?: string; Id?: string; TextField?: string }>): void =>
  setSessionCache(FIELDS_CACHE_KEY, fields);
const getCachedBuTerms = (): Term[] | null => getSessionCache<Term[]>(BU_TERMS_CACHE_KEY);
const setCachedBuTerms = (terms: Term[]): void => setSessionCache(BU_TERMS_CACHE_KEY, terms);

const FILTER_LABELS: Array<{ key: TFilterKey | null; label: string }> = [
  { key: null, label: 'All' },
  { key: 'businessUnit', label: 'Business Unit' },
  { key: 'department', label: 'Department' },
  { key: 'client', label: 'Client' },
  { key: 'geography', label: 'Geography' },
  { key: 'therapyArea', label: 'Therapy Area' },
  { key: 'diseaseArea', label: 'Disease Area' }
];

const CATEGORY_FILTER_SEARCH_PROPERTIES: Partial<Record<TFilterKey, string>> = {
  businessUnit: SEARCH_PROPERTIES.businessUnit,
  department: SEARCH_PROPERTIES.department,
  client: SEARCH_PROPERTIES.client,
  geography: SEARCH_PROPERTIES.geography,
  therapyArea: SEARCH_PROPERTIES.therapyArea,
  diseaseArea: SEARCH_PROPERTIES.diseaseArea
};

const CATEGORY_SEARCH_BASIC_SELECT_PROPERTIES = [
  SEARCH_PROPERTIES.title,
  SEARCH_PROPERTIES.path,
  SEARCH_PROPERTIES.listItemId
].join(',');

const escapeSearchFilterValue = (value: string): string =>
  value.replace(/\\/g, '\\\\').replace(/"/g, '\\"').trim();

const normalizeRefinerKey = (value: string): string => normalizeText(value).toLowerCase();

const getCategoryFilterSearchConditions = (
  filters: Partial<Record<TFilterKey, string | null>>,
  excludedKey?: TFilterKey
): string[] =>
  (Object.keys(CATEGORY_FILTER_SEARCH_PROPERTIES) as TFilterKey[])
    .map((key) => {
      if (key === excludedKey) return '';
      const property = CATEGORY_FILTER_SEARCH_PROPERTIES[key];
      const value = filters[key];
      if (!property || !value) return '';
      return `${property}:"${escapeSearchFilterValue(value)}"`;
    })
    .filter((condition): condition is string => Boolean(condition));

const getCategoryDocumentFilterValue = (document: ICategoryDocument, key: TFilterKey): string => {
  switch (key) {
    case 'reviewerComments':
      return document.reviewerComments ? 'Has reviewer comments' : 'No reviewer comments';
    case 'contentRefreshDate':
      return document.contentRefreshDate || '';
    case 'projectId':
      return document.projectId || '';
    case 'versionFileType':
      return document.versionFileType || '';
    default:
      return normalizeText(document[key]);
  }
};

const doesCategoryDocumentMatchFilters = (
  document: ICategoryDocument,
  filters: Partial<Record<TFilterKey, string | null>>
): boolean =>
  FILTER_LABELS.every((filter) => {
    if (!filter.key) return true;
    const selectedValue = filters[filter.key];
    return !selectedValue || doesFilterValueMatch(getCategoryDocumentFilterValue(document, filter.key), selectedValue);
  });

const buildPaginationItems = (currentPage: number, totalPages: number): PaginationItem[] => {
  if (totalPages <= 7) {
    return Array.from({ length: totalPages }, (_, index) => index + 1);
  }

  if (currentPage <= 4) {
    return [1, 2, 3, 4, 5, 'ellipsis-end', totalPages];
  }

  if (currentPage >= totalPages - 3) {
    return [1, 'ellipsis-start', totalPages - 4, totalPages - 3, totalPages - 2, totalPages - 1, totalPages];
  }

  return [1, 'ellipsis-start', currentPage - 1, currentPage, currentPage + 1, 'ellipsis-end', totalPages];
};

interface ICategoryDocumentsPageProps {
  context: WebPartContext;
  categoryTitle: string;
  taxonomyOptions?: ITaxonomyFieldOptions | null;
  filterGroups?: ISharedFilterGroup[];
  onHomeOpen?: () => void;
  onAllDocumentsOpen?: () => void;
  onBusinessUnitsOpen?: () => void;
  onBookmarksOpen?: () => void;
  onDocumentsOpen?: () => void;
  onContactOpen?: () => void;
  onAuditLogOpen?: () => void;
  onAnalyticsOpen?: () => void;
  onViewDocument?: (documentId: number) => void;
  onLogout?: () => void;
  userName?: string;
  userEmail?: string;
  userPhotoUrl?: string;
  showReviewerNav?: boolean;
  hideDocumentsNav?: boolean;
  isLearner?: boolean;
}

const stringifyMetadataValue = (value: unknown): string => {
  if (value === undefined || value === null) {
    return '';
  }

  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value);
  }

  if (Array.isArray(value)) {
    return value.map(stringifyMetadataValue).filter(Boolean).join('; ');
  }

  if (typeof value === 'object') {
    const objectValue = value as Record<string, any>;

    if (Array.isArray(objectValue.results)) {
      return stringifyMetadataValue(objectValue.results);
    }

    if (Array.isArray(objectValue.value)) {
      return stringifyMetadataValue(objectValue.value);
    }

    return stringifyMetadataValue(
      objectValue.Label ||
      objectValue.TermLabel ||
      objectValue.LookupValue ||
      objectValue.Title ||
      objectValue.Name ||
      objectValue.Value ||
      ''
    );
  }

  return '';
};

const normalizeText = (value: unknown): string =>
  stringifyMetadataValue(value)
    .replace(/\|[0-9a-f-]{8,}/gi, '')
    .replace(/-?\d+;#/g, '')
    .replace(/;#/g, ';')
    .replace(/\s*;\s*/g, '; ')
    .trim();

const getSearchRowCells = (row: any): Record<string, string> =>
  (row?.Cells || []).reduce((accumulator: Record<string, string>, cell: { Key: string; Value: string }) => ({
    ...accumulator,
    [cell.Key]: cell.Value
  }), {});

const getFileNameFromSearchPath = (path: string): string => {
  const lastSegment = (path || '').split('/').pop() || '';
  try {
    return decodeURIComponent(lastSegment);
  } catch {
    return lastSegment;
  }
};

const normalizeCategoryMatchText = (value: unknown): string =>
  normalizeText(value)
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\bstudies\b/g, 'study')
    .replace(/\b([a-z]+)ies\b/g, '$1y')
    .replace(/\b([a-z]+)s\b/g, '$1')
    .replace(/\s+/g, ' ')
    .trim();

const getDocumentTypeTokens = (value: unknown): string[] =>
  normalizeText(value)
    .split(/[;,\n]+/)
    .map((part) => normalizeCategoryMatchText(part))
    .filter(Boolean);

const CATEGORY_DOCUMENT_TYPE_MATCHES: Record<string, string[]> = {
  'case study': ['case study'],
  capability: ['capability'],
  proposal: ['proposal'],
  'lesson learned': ['lesson learned', 'lessons learned'],
  'knowledge sharing session': ['ks2', 'knowledge sharing session', 'knowledge sharing sessions'],
};

const CATEGORY_SEARCH_LABELS: Record<string, string> = {
  'case study': 'Case Study',
  capability: 'Capability',
  proposal: 'Proposal',
  'lesson learned': 'Lesson Learned',
  'knowledge sharing session': 'KS2',
};

const matchesDocumentTypeCategory = (value: unknown, categoryTitle: string): boolean => {
  const normalizedCategory = normalizeCategoryMatchText(categoryTitle);
  const allowedDocumentTypes = CATEGORY_DOCUMENT_TYPE_MATCHES[normalizedCategory] || [normalizedCategory];

  if (!normalizedCategory) {
    return false;
  }

  return getDocumentTypeTokens(value).some((documentType) => allowedDocumentTypes.indexOf(documentType) !== -1);
};

const getFieldValue = (item: any, internalName?: string, fallbackNames: string[] = []): unknown => {
  const candidates = [internalName, ...fallbackNames].filter(Boolean) as string[];
  for (const candidate of candidates) {
    if (item && item[candidate] !== undefined && item[candidate] !== null) {
      return item[candidate];
    }
  }
  return undefined;
};

const normalizeNumberField = (value: unknown): number | undefined => {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const numericValue = Number(normalizeText(value));
  return Number.isFinite(numericValue) ? numericValue : undefined;
};

const normalizeBooleanField = (value: unknown): boolean | undefined => {
  if (typeof value === 'boolean') {
    return value;
  }

  const normalizedValue = normalizeText(value).toLowerCase();
  if (['true', 'yes', '1'].indexOf(normalizedValue) !== -1) {
    return true;
  }
  if (['false', 'no', '0'].indexOf(normalizedValue) !== -1) {
    return false;
  }
  return undefined;
};

const normalizeStringArrayField = (value: unknown): string[] =>
  normalizeText(value)
    .split(/[;,\n]+/)
    .map((part) => part.trim())
    .filter(Boolean);

const getNormalizedMetadataValue = (item: any, internalName?: string, fallbackNames: string[] = []): string =>
  normalizeText(getFieldValue(item, internalName, fallbackNames));

const getNormalizedMetadataValues = (item: any, internalNames: Array<string | undefined>): string =>
  normalizeText(
    internalNames
      .filter(Boolean)
      .map((internalName) => getFieldValue(item, internalName))
      .filter((value) => value !== undefined && value !== null)
      .map(stringifyMetadataValue)
      .filter(Boolean)
      .join('; ')
  );

const normalizeFilterValue = (value: string): string =>
  normalizeText(value)
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();

const doesFilterValueMatch = (documentValue: string, selectedValue: string): boolean => {
  const normalizedDocumentValue = normalizeFilterValue(documentValue);
  const normalizedSelectedValue = normalizeFilterValue(selectedValue);

  if (!normalizedDocumentValue || !normalizedSelectedValue) {
    return false;
  }

  return normalizedDocumentValue === normalizedSelectedValue ||
    normalizedDocumentValue.indexOf(normalizedSelectedValue) !== -1 ||
    normalizedSelectedValue.indexOf(normalizedDocumentValue) !== -1;
};

const getUniqueTermOptions = (terms: ITaxonomyTerm[]): Array<{ title: string; count: number }> => {
  const uniqueOptions = new Map<string, { title: string; count: number }>();

  terms.forEach((term) => {
    const title = normalizeText(term.label);
    const key = title.toLowerCase();

    if (!title || uniqueOptions.has(key)) {
      return;
    }

    uniqueOptions.set(key, { title, count: 0 });
  });

  return Array.from(uniqueOptions.values()).sort((left, right) => left.title.localeCompare(right.title));
};

const getTermHierarchyLevel = (term: ITaxonomyTerm): number => {
  if (typeof term.level === 'number' && term.level >= 0) {
    return term.level;
  }

  const pathSegments = (term.path || term.label || '')
    .replace(/;/g, ' > ')
    .replace(/\|/g, ' > ')
    .split(/\s*>\s*/)
    .map((segment) => segment.trim())
    .filter(Boolean);

  if (pathSegments.length > 1) {
    return pathSegments.length - 1;
  }

  return term.parentId ? 1 : 0;
};

const getUniqueNamedOptions = (names: string[], level = 0): ISharedFilterOption[] => {
  const uniqueOptions = new Map<string, ISharedFilterOption>();

  names.forEach((name) => {
    const title = normalizeText(name);
    const key = title.toLowerCase();

    if (!title || uniqueOptions.has(key)) {
      return;
    }

    uniqueOptions.set(key, { title, count: 0, level });
  });

  return Array.from(uniqueOptions.values()).sort((left, right) => left.title.localeCompare(right.title));
};

const getBusinessUnitOptionsFromTree = (terms: Term[]): ISharedFilterOption[] =>
  getUniqueNamedOptions(terms.map((term) => term.name), 0);

const getDepartmentOptionsFromTree = (
  terms: Term[],
  selectedBusinessUnit?: string | null
): ISharedFilterOption[] => {
  const normalizedSelectedBusinessUnit = normalizeFilterValue(selectedBusinessUnit || '');
  const scopedTerms = normalizedSelectedBusinessUnit
    ? terms.filter((term) => normalizeFilterValue(term.name) === normalizedSelectedBusinessUnit)
    : terms;
  const uniqueOptions = new Map<string, ISharedFilterOption>();

  scopedTerms.forEach((businessUnit) => {
    (businessUnit.children || []).forEach((department) => {
      const departmentTitle = normalizeText(department.name || '');
      const departmentKey = departmentTitle.toLowerCase();

      if (departmentTitle && !uniqueOptions.has(departmentKey)) {
        uniqueOptions.set(departmentKey, {
          title: departmentTitle,
          count: 0,
          level: 0
        });
      }

      (department.children || []).forEach((subDepartment) => {
        const subDepartmentTitle = normalizeText(subDepartment.name || '');
        const subDepartmentKey = subDepartmentTitle.toLowerCase();

        if (subDepartmentTitle && !uniqueOptions.has(subDepartmentKey)) {
          uniqueOptions.set(subDepartmentKey, {
            title: subDepartmentTitle,
            count: 0,
            level: 1,
            parentTitle: departmentTitle
          });
        }
      });
    });
  });

  return Array.from(uniqueOptions.values());
};

const getStatusClassName = (status?: string): string => {
  const normalized = normalizeStatusKey(status);
  if (normalized === 'active') return styles.statusBadgeActive;
  if (normalized === 'underReview') return styles.statusBadgeReview;
  if (normalized === 'reject') return styles.statusBadgeRejected;
  if (normalized === 'archive') return styles.statusBadgeArchive;
  return styles.statusBadgeDefault;
};

const normalizeStatusKey = (status?: string): TStatusKey => {
  const normalizedStatus = normalizeText(status)
    .toLowerCase()
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!normalizedStatus) {
    return 'unknown';
  }

  const tokens = normalizedStatus
    .split(/[;,/|]+/)
    .map((token) => token.trim())
    .filter(Boolean);
  const candidates = tokens.length > 0 ? tokens : [normalizedStatus];
  const hasCandidate = (values: string[]): boolean =>
    candidates.some((candidate) => values.indexOf(candidate) !== -1);

  if (hasCandidate(['reject', 'rejected']) || /\breject(ed)?\b/.test(normalizedStatus)) return 'reject';
  if (hasCandidate(['archive', 'archived']) || /\barchive(d)?\b/.test(normalizedStatus)) return 'archive';
  if (hasCandidate(['under review', 'underreview', 'review']) || /\bunder\s*review\b/.test(normalizedStatus)) return 'underReview';
  if (hasCandidate(['active', 'approved']) || /^(active|approved)$/.test(normalizedStatus)) return 'active';

  return 'unknown';
};

const getDisplayStatus = (status?: string): string => {
  const trimmedStatus = String(status || '').trim();
  if (!trimmedStatus) {
    return 'Unknown';
  }
  return trimmedStatus.toLowerCase() === 'rejected' ? 'Reject' : trimmedStatus;
};

const buildFileUrl = (webUrl: string, fileRef: string): string => {
  if (!fileRef) {
    return '';
  }

  if (/^https?:\/\//i.test(fileRef)) {
    return fileRef;
  }

  const origin = new URL(webUrl).origin;
  return `${origin}${fileRef}`;
};

const getServerRelativeFileUrl = (webUrl: string, fileRef: string): string => {
  const trimmedFileRef = String(fileRef || '').trim();
  if (!trimmedFileRef) {
    return '';
  }

  if (/^https?:\/\//i.test(trimmedFileRef)) {
    try {
      return new URL(trimmedFileRef).pathname;
    } catch {
      return trimmedFileRef;
    }
  }

  return trimmedFileRef.startsWith('/') ? trimmedFileRef : `/${trimmedFileRef}`;
};

export const CategoryDocumentsPage: React.FC<ICategoryDocumentsPageProps> = (props) => {
  const [documents, setDocuments] = React.useState<ICategoryDocument[]>([]);
  const [isLoading, setIsLoading] = React.useState(true);
  const [countLoading, setCountLoading] = React.useState(false);
  const [fallbackTaxonomyOptions, setFallbackTaxonomyOptions] = React.useState<ITaxonomyFieldOptions | null>(null);
  const [buDepartmentTree, setBuDepartmentTree] = React.useState<Term[]>([]);
  const [activeFilterKey, setActiveFilterKey] = React.useState<TFilterKey | null>(null);
  const [selectedFilters, setSelectedFilters] = React.useState<Partial<Record<TFilterKey, string | null>>>({});
  const [filterRefinerCounts, setFilterRefinerCounts] = React.useState<TFilterRefinerCounts>({});
  const [filterRefinerLoadingKey, setFilterRefinerLoadingKey] = React.useState<TFilterKey | null>(null);
  const [, setCurrentPage] = React.useState(1);
  const [totalCount, setTotalCount] = React.useState<number>(0);
  const [currentRestPage, setCurrentRestPage] = React.useState<number>(1);
  const [viewMode, setViewMode] = React.useState<TDocumentViewMode>('list');
  const [thumbnailAttemptByDocument, setThumbnailAttemptByDocument] = React.useState<Record<number, number>>({});
  const [thumbnailLoadedByDocument, setThumbnailLoadedByDocument] = React.useState<Record<number, boolean>>({});
  const [downloadingDocumentIds, setDownloadingDocumentIds] = React.useState<Record<number, boolean>>({});
  const categoryFetchContextRef = React.useRef<ICategoryFetchContext | null>(null);
  const categorySearchStartOffsetAdjustmentRef = React.useRef<number>(0);
  const categoryDocumentsRequestRef = React.useRef(0);
  const categoryFilterRefinerRequestRef = React.useRef(0);
  const taxonomyOptions = props.taxonomyOptions || fallbackTaxonomyOptions;

  React.useEffect(() => {
    if (props.taxonomyOptions) {
      setFallbackTaxonomyOptions(null);
      return;
    }

    let isDisposed = false;
    const cachedTaxonomy = getCachedTaxonomy();
    if (cachedTaxonomy) {
      setFallbackTaxonomyOptions(cachedTaxonomy);
      return () => {
        isDisposed = true;
      };
    }

    fetchAllTaxonomyOptions(props.context)
      .then((options) => {
        if (!isDisposed) {
          setFallbackTaxonomyOptions(options);
          setCachedTaxonomy(options);
        }
      })
      .catch((error) => {
        console.error('Unable to load category filter taxonomy:', error);
        if (!isDisposed) {
          setFallbackTaxonomyOptions(null);
        }
      });

    return () => {
      isDisposed = true;
    };
  }, [props.context, props.taxonomyOptions]);

  React.useEffect(() => {
    let isDisposed = false;

    const loadBuDepartmentTerms = async (): Promise<void> => {
      try {
        const cachedTerms = getCachedBuTerms();
        if (cachedTerms) {
          if (!isDisposed) {
            setBuDepartmentTree(cachedTerms);
          }
          return;
        }

        const terms = await fetchBUDepartmentTerms(props.context.spHttpClient, props.context.pageContext.web.absoluteUrl);
        setCachedBuTerms(terms);
        if (!isDisposed) {
          setBuDepartmentTree(terms);
        }
      } catch (error) {
        console.error('Unable to load category BU/department tree:', error);
        if (!isDisposed) {
          setBuDepartmentTree([]);
        }
      }
    };

    void loadBuDepartmentTerms();

    return () => {
      isDisposed = true;
    };
  }, [props.context]);

  React.useEffect(() => {
    let isDisposed = false;
    const requestId = categoryDocumentsRequestRef.current + 1;
    categoryDocumentsRequestRef.current = requestId;
    const isCurrentRequest = (): boolean => !isDisposed && categoryDocumentsRequestRef.current === requestId;
    setCurrentRestPage(1);
    setCurrentPage(1);
    setTotalCount(0);
    setCountLoading(false);
    categorySearchStartOffsetAdjustmentRef.current = 0;
    setDocuments([]);

    const fetchDocuments = async (): Promise<void> => {
      setIsLoading(true);
      const webUrl = props.context.pageContext.web.absoluteUrl;

      try {
        const normalizedCategory = normalizeCategoryMatchText(props.categoryTitle);
        const searchLabel = CATEGORY_SEARCH_LABELS[normalizedCategory];
        if (!searchLabel) {
          if (isCurrentRequest()) {
            setDocuments([]);
            setTotalCount(0);
            setIsLoading(false);
          }
          return;
        }
        const SEARCH_PAGE_SIZE = PAGE_SIZES.searchPageSize;
        const startRow = 0;
        const searchService = getSearchServiceInstance(
          props.context.spHttpClient,
          webUrl,
          LIBRARY_NAMES.kmDataHub
        );
        const rootFolderUrl = await searchService.getLibraryRootFolderUrl();
        const categorySearchFilterConditions = getCategoryFilterSearchConditions(selectedFilters);
        const hasActiveCategorySearchFilters = categorySearchFilterConditions.length > 0;
        const queryPartsForSearch = [
          rootFolderUrl ? `Path:"${rootFolderUrl}/*"` : '',
          `${SEARCH_PROPERTIES.documentType}:"${escapeSearchFilterValue(searchLabel)}"`,
          `${SEARCH_PROPERTIES.status}:Active`,
          ...categorySearchFilterConditions
        ].filter(Boolean);
        const query = queryPartsForSearch.join(' AND ');
        const buildCategorySearchUrl = (selectProperties: string, sortProperty?: string, queryText: string = query): string =>
          `${webUrl}/_api/search/query` +
          `?querytext='${encodeURIComponent(queryText)}'` +
          `&selectproperties='${selectProperties}'` +
          `&rowlimit=${SEARCH_PAGE_SIZE}` +
          `&startrow=${startRow}` +
          `${sortProperty ? `&sortlist='${sortProperty}:descending'` : ''}` +
          `&trimduplicates=false`;
        const searchUrl = buildCategorySearchUrl(CATEGORY_SEARCH_BASIC_SELECT_PROPERTIES);
        const filterCountUrl =
          `${webUrl}/_api/search/query` +
          `?querytext='${encodeURIComponent(query)}'` +
          `&rowlimit=1` +
          `&selectproperties='${SEARCH_PROPERTIES.listItemId},${SEARCH_PROPERTIES.status},${SEARCH_PROPERTIES.documentType}'` +
          `&trimduplicates=false` +
          `&clienttype='ContentSearchRegular'`;
        const cachedFieldMap = fieldMapModuleCache || getCachedFieldMap<any>();
        const cachedFields = getCachedFields();
        const [readFieldMap, fields, searchResp, filterCountResp, categoryRestCount, latestIds] = await Promise.all([
          cachedFieldMap
            ? Promise.resolve(cachedFieldMap).then((map) => {
              fieldMapModuleCache = map;
              return map;
            })
            : fetchKMDataHubReadFieldMap(props.context.spHttpClient, webUrl, LIBRARY_NAME).then((map) => {
              setCachedFieldMap(map);
              fieldMapModuleCache = map;
              return map;
            }),
          cachedFields
            ? Promise.resolve(cachedFields)
            : props.context.spHttpClient.get(
              `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/fields?$select=Title,InternalName,Id,TextField`,
              SPHttpClient.configurations.v1,
              { headers: { Accept: 'application/json;odata.metadata=minimal' } }
            ).then(async (fieldsResponse) => {
              const fieldsJson = fieldsResponse.ok ? await fieldsResponse.json() : {};
              const loadedFields = (fieldsJson?.value || fieldsJson?.d?.results || []) as Array<{ Title?: string; InternalName?: string; Id?: string; TextField?: string }>;
              setCachedFields(loadedFields);
              return loadedFields;
          }),
          props.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1),
          hasActiveCategorySearchFilters
            ? props.context.spHttpClient.get(filterCountUrl, SPHttpClient.configurations.v1)
            : Promise.resolve(null),
          searchService.getListItemCount('Active', searchLabel),
          searchService.getLatestDocumentIds(40)
        ]);
        if (isCurrentRequest()) {
          setTotalCount(categoryRestCount || 0);
          setCountLoading(false);
        }

        const normalizeGuidValue = (value?: string): string => (value || '').replace(/[{}]/g, '').trim().toLowerCase();
        const findField = (titles: string[], fallbacks: string[]): string | undefined => {
          const normalizedTitles = titles.map((title) => title.toLowerCase());
          const matched = fields.find((field) =>
            normalizedTitles.includes(String(field.Title || '').trim().toLowerCase()) ||
            fallbacks.includes(String(field.InternalName || '').trim())
          );
          return matched?.InternalName || fallbacks.find(Boolean);
        };
        const findTaxonomyNoteField = (taxonomyInternalName?: string): string | undefined => {
          if (!taxonomyInternalName) {
            return undefined;
          }

          const taxonomyField = fields.find((field) => String(field.InternalName || '').trim().toLowerCase() === taxonomyInternalName.toLowerCase());
          const textFieldId = normalizeGuidValue(taxonomyField?.TextField);
          if (!textFieldId) {
            return undefined;
          }

          return fields.find((field) => normalizeGuidValue(field.Id) === textFieldId)?.InternalName;
        };

        const documentTypeField = findField(['Document Type', 'DocumentType'], [TAXONOMY_FIELD_CONFIGS.documentType.fieldInternalName, 'DocumentType', 'Document_x0020_Type']);
        const documentTypeNoteField = findTaxonomyNoteField(documentTypeField);
        const descriptionField = findField([COLUMN_NAMES.description, 'Abstract'], [COLUMN_NAMES.description, 'Abstract']);
        const businessUnitField = findField(['BU', 'Business Unit', 'BusinessUnit'], ['BU', 'BusinessUnit']);
        const departmentField = findField(['Department / Sub Department', 'Department', 'Sub Department'], ['Department_x0020__x002f__x0020_Sub_x0020_Department', 'Department', 'SubDepartment']);
        const clientField = findField(['Client'], ['Client']);
        const geographyField = findField(['Geography', 'Region'], ['Geography', 'Region']);
        const therapyAreaField = findField(['Therapy Area', 'TherapyArea'], ['TherapyArea', 'Therapy_x0020_Area']);
        const diseaseAreaField = findField(['Disease Area', 'DiseaseArea'], ['DiseaseArea', 'Disease_x0020_Area']);
        const urlField = readFieldMap.url || findField(['URL'], ['URL']);
        const titleField = readFieldMap.title || COLUMN_NAMES.title;
        const docIconField = readFieldMap.docIcon || COLUMN_NAMES.docIcon;
        const fileLeafRefField = readFieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef;
        const fileRefField = readFieldMap.fileRef || COLUMN_NAMES.fileRef;

        const queryParts = buildKMDataHubItemQuery(
          readFieldMap,
          [
            'ID',
            titleField,
            COLUMN_NAMES.created,
            fileLeafRefField,
            fileRefField,
            'File/UniqueId',
            descriptionField || COLUMN_NAMES.description,
            documentTypeField || 'DocumentType',
            businessUnitField || 'BU',
            departmentField || 'Department',
            clientField || 'Client',
            geographyField || 'Geography',
            therapyAreaField || 'TherapyArea',
            diseaseAreaField || 'DiseaseArea',
            docIconField,
            readFieldMap.published,
            readFieldMap.status,
            urlField || 'URL'
          ],
          ['File']
        );

        const setFallbackField = (item: Record<string, any>, internalName: string | undefined, value: unknown): void => {
          if (internalName && value !== undefined && value !== null && value !== '') {
            item[internalName] = value;
          }
        };

        const mapSearchRowsToItems = (searchRows: any[]): any[] =>
          searchRows
            .map((row) => {
              const cells = getSearchRowCells(row);
              const id = Number(cells[SEARCH_PROPERTIES.listItemId]);
              if (!Number.isFinite(id) || id <= 0) {
                return null;
              }

              const path = cells[SEARCH_PROPERTIES.path] || '';
              const fileName = cells[COLUMN_NAMES.fileLeafRef] || getFileNameFromSearchPath(path);
              const fileExtension = cells.FileExtension || fileName.split('.').pop() || '';
              const title = cells[SEARCH_PROPERTIES.title] || fileName || `Document ${id}`;
              const documentTypeValue = cells[SEARCH_PROPERTIES.documentType] || searchLabel;
              const statusValue = cells[SEARCH_PROPERTIES.status] || 'Active';
              const publishedValue = cells[SEARCH_PROPERTIES.published] || cells[SEARCH_PROPERTIES.write] || '';
              const businessUnitValue = cells[SEARCH_PROPERTIES.businessUnit] || selectedFilters.businessUnit || '';
              const departmentValue = cells[SEARCH_PROPERTIES.department] || selectedFilters.department || '';
              const clientValue = cells[SEARCH_PROPERTIES.client] || selectedFilters.client || '';
              const geographyValue = cells[SEARCH_PROPERTIES.geography] || selectedFilters.geography || '';
              const therapyAreaValue = cells[SEARCH_PROPERTIES.therapyArea] || selectedFilters.therapyArea || '';
              const diseaseAreaValue = cells[SEARCH_PROPERTIES.diseaseArea] || selectedFilters.diseaseArea || '';
              const item: Record<string, any> = {
                Id: id,
                ID: id,
                Title: title,
                FileLeafRef: fileName,
                FileRef: path,
                URL: path,
                Status: statusValue,
                Document_x0020_Type: documentTypeValue,
                Description: cells[COLUMN_NAMES.description] || '',
                Published: publishedValue,
                FileExtension: fileExtension,
                Author0: { Title: 'Internal' }
              };

              setFallbackField(item, titleField, title);
              setFallbackField(item, fileLeafRefField, fileName);
              setFallbackField(item, fileRefField, path);
              setFallbackField(item, urlField || COLUMN_NAMES.url, path);
              setFallbackField(item, readFieldMap.status || COLUMN_NAMES.status, statusValue);
              setFallbackField(item, documentTypeField || COLUMN_NAMES.documentType, documentTypeValue);
              setFallbackField(item, documentTypeNoteField, documentTypeValue);
              setFallbackField(item, descriptionField || COLUMN_NAMES.description, cells[COLUMN_NAMES.description] || '');
              setFallbackField(item, readFieldMap.published || COLUMN_NAMES.published, publishedValue);
              setFallbackField(item, businessUnitField || COLUMN_NAMES.businessUnit, businessUnitValue);
              setFallbackField(item, departmentField || COLUMN_NAMES.department, departmentValue);
              setFallbackField(item, clientField || COLUMN_NAMES.client, clientValue);
              setFallbackField(item, geographyField || COLUMN_NAMES.geography, geographyValue);
              setFallbackField(item, therapyAreaField || COLUMN_NAMES.therapyArea, therapyAreaValue);
              setFallbackField(item, diseaseAreaField || COLUMN_NAMES.diseaseArea, diseaseAreaValue);
              return item;
            })
            .filter(Boolean);

        const mapCategoryDocuments = (items: any[]): ICategoryDocument[] =>
          items
            .filter((item) => matchesDocumentTypeCategory(
              getNormalizedMetadataValues(item, [documentTypeField, documentTypeNoteField, 'DocumentType', 'Document_x0020_Type']),
              props.categoryTitle
            ))
            .map<ICategoryDocument>((item) => {
              const fileName = normalizeText(getFieldValue(item, fileLeafRefField, [COLUMN_NAMES.fileLeafRef]));
              const title = normalizeText(getFieldValue(item, titleField, [COLUMN_NAMES.title])) || fileName || `Document ${item.Id}`;
              const fileRef = normalizeText(getFieldValue(item, fileRefField, [COLUMN_NAMES.fileRef])) || normalizeText(getFieldValue(item, urlField, [COLUMN_NAMES.url]));
              const publishedValue = resolveDocumentPublishedValue(item, readFieldMap.published);
              const publishedDate = publishedValue ? new Date(publishedValue) : null;
              const documentType = getNormalizedMetadataValues(item, [documentTypeField, documentTypeNoteField, 'DocumentType', 'Document_x0020_Type']) || props.categoryTitle;
              const status = getNormalizedMetadataValue(item, readFieldMap.status, [COLUMN_NAMES.status]);
              return {
                id: Number(item.Id || item.ID),
                title,
                description: normalizeText(getFieldValue(item, descriptionField, [COLUMN_NAMES.description, 'Abstract'])) || 'No abstract available',
                documentType,
                author: resolveDocumentAuthor(item, 'Internal', readFieldMap.author),
                date: formatDocumentPublishedDate(publishedValue, 'en-GB', {
                  day: '2-digit',
                  month: '2-digit',
                  year: 'numeric'
                }) || 'No date available',
                createdTime: publishedDate && !Number.isNaN(publishedDate.getTime()) ? publishedDate.getTime() : 0,
                status,
                statusKey: normalizeStatusKey(status),
                businessUnit: getNormalizedMetadataValue(item, businessUnitField, ['BU', 'BusinessUnit']),
                department: getNormalizedMetadataValue(item, departmentField, ['Department_x0020__x002f__x0020_Sub_x0020_Department', 'Department', 'SubDepartment']),
                client: getNormalizedMetadataValue(item, clientField, ['Client']),
                geography: getNormalizedMetadataValue(item, geographyField, ['Geography', 'Region']),
                therapyArea: getNormalizedMetadataValue(item, therapyAreaField, ['TherapyArea', 'Therapy_x0020_Area']),
                diseaseArea: getNormalizedMetadataValue(item, diseaseAreaField, ['DiseaseArea', 'Disease_x0020_Area']),
                fileUrl: buildFileUrl(webUrl, fileRef),
                fileName,
                fileType: (item.FileExtension || fileName.split('.').pop() || '').toUpperCase(),
                fileRef,
                fileUniqueId: item.File?.UniqueId || '',
                reviewerComments: '',
                docIcon: normalizeText(getFieldValue(item, docIconField, [COLUMN_NAMES.docIcon]))
              };
            });

        categoryFetchContextRef.current = {
          mapCategoryDocuments,
          mapSearchRowsToItems,
          queryParts
        };

        let effectiveSearchResp = searchResp;
        if (!effectiveSearchResp.ok) {
          effectiveSearchResp = await props.context.spHttpClient.get(
            buildCategorySearchUrl(CATEGORY_SEARCH_BASIC_SELECT_PROPERTIES, SEARCH_PROPERTIES.write),
            SPHttpClient.configurations.v1
          );
        }
        if (!effectiveSearchResp.ok) {
          if (isCurrentRequest()) {
            setDocuments([]);
          }
          return;
        }
        const searchData = await effectiveSearchResp.json();
        const searchTotalRows = searchData?.PrimaryQueryResult?.RelevantResults?.TotalRows ?? 0;
        let filterSearchTotalRows = searchTotalRows;
        if (hasActiveCategorySearchFilters && filterCountResp) {
          if (filterCountResp.ok) {
            const filterCountData = await filterCountResp.json();
            filterSearchTotalRows =
              filterCountData?.PrimaryQueryResult?.RelevantResults?.TotalRows ??
              filterCountData?.d?.query?.PrimaryQueryResult?.RelevantResults?.TotalRows ??
              searchTotalRows;
          }
        }
        const rows = searchData?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];
        const listItemIds: number[] = rows
          .map((row: any) => {
            const cells = getSearchRowCells(row);
            return Number(cells.ListItemID);
          })
          .filter((id: number) => Number.isFinite(id) && id > 0);
        const fallbackItemsById = new Map<number, any>(
          mapSearchRowsToItems(rows).map((item: any) => [Number(item.Id || item.ID), item])
        );
        const isActiveCategoryItem = (item: any): boolean => {
          const status = String(item[readFieldMap.status] || item.Status || '').trim();
          const docType = getNormalizedMetadataValues(item, [documentTypeField, documentTypeNoteField, 'DocumentType', 'Document_x0020_Type']);
          return status === 'Active' && matchesDocumentTypeCategory(docType, normalizedCategory);
        };
        const detailIds = Array.from(new Set([...listItemIds, ...latestIds])).slice(0, 35);
        const detailItems = detailIds.length > 0
          ? await searchService.getDocumentDetailsByIds(detailIds, queryParts)
          : [];
        const detailItemsById = new Map<number, any>(
          detailItems.map((item: any) => [Number(item.Id || item.ID), item])
        );
        const indexedItems = listItemIds
          .map((id) => detailItemsById.get(id) || fallbackItemsById.get(id))
          .filter(Boolean);
        const searchIdSet = new Set(listItemIds);
        const latestDetails = latestIds
          .map((id) => detailItemsById.get(id))
          .filter(Boolean);
        const latestRestCategoryDocuments: ICategoryDocument[] = hasActiveCategorySearchFilters
          ? []
          : mapCategoryDocuments(latestDetails.filter((item: any) => {
            const id = Number(item.Id || item.ID);
            return !searchIdSet.has(id) && isActiveCategoryItem(item);
          }));
        const latestCategoryCorrection = latestRestCategoryDocuments.length;
        const indexedCategoryDocuments = mapCategoryDocuments(indexedItems.filter(isActiveCategoryItem));
        const categoryDocuments = [...latestRestCategoryDocuments, ...indexedCategoryDocuments];
        const seenDocumentIds = new Set<number>();
        const uniqueDocuments = categoryDocuments.filter((document) => {
          if (seenDocumentIds.has(document.id)) {
            return false;
          }
          seenDocumentIds.add(document.id);
          return true;
        });
        uniqueDocuments.sort((a, b) => (b.createdTime || 0) - (a.createdTime || 0) || b.id - a.id);

        if (isCurrentRequest()) {
          const baseCategoryCount = hasActiveCategorySearchFilters ? filterSearchTotalRows : (categoryRestCount || searchTotalRows);
          const stableLatestCategoryCorrection = hasActiveCategorySearchFilters ? 0 : latestCategoryCorrection;
          const correctedCategoryCount = Math.max(0, baseCategoryCount + stableLatestCategoryCorrection);
          console.log('[CATEGORY COUNT] Search crawled Active:', searchTotalRows);
          console.log('[CATEGORY COUNT] REST uncrawled new docs:', latestCategoryCorrection);
          console.log('[CATEGORY COUNT] Total shown to user:', correctedCategoryCount);
          setTotalCount(correctedCategoryCount);
          setCountLoading(false);
          setCurrentRestPage(1);
          setCurrentPage(1);
          categorySearchStartOffsetAdjustmentRef.current = Math.max(0, stableLatestCategoryCorrection);
          setDocuments(uniqueDocuments);
          setIsLoading(false);
        }
      } catch (error) {
        console.error('Unable to load category documents:', error);
        if (isCurrentRequest()) {
          setDocuments([]);
        }
      } finally {
        if (isCurrentRequest()) {
          setIsLoading(false);
        }
      }
    };

    void fetchDocuments();

    return () => {
      isDisposed = true;
    };
  }, [props.categoryTitle, props.context, selectedFilters]);

  const fetchCategoryPage = async (pageNum: number): Promise<void> => {
    if (!props.context || !categoryFetchContextRef.current) return;
    setIsLoading(true);
    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const SEARCH_PAGE_SIZE = DOCUMENTS_PER_PAGE;
      const startRow = Math.max(
        0,
        (pageNum - 1) * DOCUMENTS_PER_PAGE - categorySearchStartOffsetAdjustmentRef.current
      );
      const normalizedCategory = normalizeCategoryMatchText(props.categoryTitle);
      const searchLabel = CATEGORY_SEARCH_LABELS[normalizedCategory];
      if (!searchLabel) return;

      const searchService = getSearchServiceInstance(
        props.context.spHttpClient,
        webUrl,
        LIBRARY_NAMES.kmDataHub
      );
      const rootFolderUrl = await searchService.getLibraryRootFolderUrl();
      const categorySearchFilterConditions = getCategoryFilterSearchConditions(selectedFilters);
      const queryPartsForSearch = [
        rootFolderUrl ? `Path:"${rootFolderUrl}/*"` : '',
        `${SEARCH_PROPERTIES.documentType}:"${escapeSearchFilterValue(searchLabel)}"`,
        `${SEARCH_PROPERTIES.status}:Active`,
        ...categorySearchFilterConditions
      ].filter(Boolean);
      const query = queryPartsForSearch.join(' AND ');
      const buildCategorySearchUrl = (
        selectProperties: string,
        rowLimit: number,
        rowStart: number,
        sortProperty?: string,
        queryText: string = query
      ): string =>
        `${webUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(queryText)}'` +
        `&selectproperties='${selectProperties}'` +
        `&rowlimit=${rowLimit}` +
        `&startrow=${rowStart}` +
        `${sortProperty ? `&sortlist='${sortProperty}:descending'` : ''}` +
        `&trimduplicates=false`;
      const searchUrl = buildCategorySearchUrl(
        CATEGORY_SEARCH_BASIC_SELECT_PROPERTIES,
        SEARCH_PAGE_SIZE,
        startRow
      );

      const extractSearchRows = (searchData: any): any[] =>
        searchData?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];
      const extractSearchIds = (searchRows: any[]): number[] =>
        searchRows
          .map((row: any) => Number(getSearchRowCells(row)[SEARCH_PROPERTIES.listItemId]))
          .filter((id: number) => Number.isFinite(id) && id > 0);

      let searchResp = await props.context.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
      if (!searchResp.ok) {
        searchResp = await props.context.spHttpClient.get(
          buildCategorySearchUrl(
            CATEGORY_SEARCH_BASIC_SELECT_PROPERTIES,
            SEARCH_PAGE_SIZE,
            startRow,
            SEARCH_PROPERTIES.write
          ),
          SPHttpClient.configurations.v1
        );
      }
      if (!searchResp.ok) return;
      const searchData = await searchResp.json();
      const searchTotalRows = searchData?.PrimaryQueryResult?.RelevantResults?.TotalRows ?? 0;
      let searchRows = extractSearchRows(searchData);
      let listItemIds: number[] = extractSearchIds(searchRows);
      if (listItemIds.length === 0 && pageNum > 1 && searchTotalRows > 0) {
        const fallbackStartRow = Math.max(0, searchTotalRows - DOCUMENTS_PER_PAGE);
        if (fallbackStartRow !== startRow) {
          const fallbackSearchUrl = buildCategorySearchUrl(
            CATEGORY_SEARCH_BASIC_SELECT_PROPERTIES,
            DOCUMENTS_PER_PAGE,
            fallbackStartRow
          );
          let fallbackResp = await props.context.spHttpClient.get(fallbackSearchUrl, SPHttpClient.configurations.v1);
          if (!fallbackResp.ok) {
            fallbackResp = await props.context.spHttpClient.get(
              buildCategorySearchUrl(
                CATEGORY_SEARCH_BASIC_SELECT_PROPERTIES,
                DOCUMENTS_PER_PAGE,
                fallbackStartRow,
                SEARCH_PROPERTIES.write
              ),
              SPHttpClient.configurations.v1
            );
          }
          if (fallbackResp.ok) {
            searchRows = extractSearchRows(await fallbackResp.json());
            listItemIds = extractSearchIds(searchRows);
          }
        }
      }

      if (listItemIds.length === 0) {
        setDocuments([]);
        setCurrentRestPage(pageNum);
        setCurrentPage(pageNum);
        return;
      }

      const fallbackItemsById = new Map<number, any>(
        categoryFetchContextRef.current.mapSearchRowsToItems(searchRows).map((item: any) => [Number(item.Id || item.ID), item])
      );
      const slicedListItemIds = listItemIds.slice(0, DOCUMENTS_PER_PAGE + 5);
      const detailItems = await searchService.getDocumentDetailsByIds(slicedListItemIds, categoryFetchContextRef.current.queryParts);
      const detailItemsById = new Map<number, any>(
        detailItems.map((item: any) => [Number(item.Id || item.ID), item])
      );
      const items = listItemIds
        .map((id) => detailItemsById.get(id) || fallbackItemsById.get(id))
        .filter(Boolean);
      const categoryDocuments = categoryFetchContextRef.current.mapCategoryDocuments(items);
      const seenDocumentIds = new Set<number>();
      const uniqueDocuments = categoryDocuments.filter((document) => {
        if (seenDocumentIds.has(document.id)) {
          return false;
        }
        seenDocumentIds.add(document.id);
        return true;
      });
      uniqueDocuments.sort((a, b) => (b.createdTime || 0) - (a.createdTime || 0) || b.id - a.id);
      setDocuments(uniqueDocuments);
      setCurrentRestPage(pageNum);
      setCurrentPage(pageNum);
    } catch (err) {
      console.error('fetchCategoryPage error:', err);
    } finally {
      setIsLoading(false);
    }
  };

  const handlePageChange = async (newPage: number): Promise<void> => {
    if (newPage === currentRestPage) return;
    await fetchCategoryPage(newPage);
  };

  const getDocumentFilterValue = React.useCallback((document: ICategoryDocument, key: TFilterKey): string => {
    return getCategoryDocumentFilterValue(document, key);
  }, []);

  const activeDocuments = React.useMemo(
    () => documents.filter((document) => document.statusKey === 'active'),
    [documents]
  );

  const filteredDocuments = React.useMemo(() => {
    const filtered = activeDocuments.filter((document) => {
      return FILTER_LABELS.every((filter) => {
        if (!filter.key) return true;
        const selectedValue = selectedFilters[filter.key];
        return !selectedValue || doesFilterValueMatch(getDocumentFilterValue(document, filter.key), selectedValue);
      });
    });

    return filtered.sort((first, second) =>
      second.createdTime - first.createdTime
    );
  }, [activeDocuments, getDocumentFilterValue, selectedFilters]);

  const totalPages = Math.max(1, Math.ceil(totalCount / DOCUMENTS_PER_PAGE));
  const categoryDisplayCount = totalCount;
  const categoryDisplayStart = categoryDisplayCount > 0 ? (currentRestPage - 1) * DOCUMENTS_PER_PAGE + 1 : 0;
  const categoryDisplayEnd = Math.min(currentRestPage * DOCUMENTS_PER_PAGE, totalCount);
  const paginationItems = React.useMemo(() => buildPaginationItems(currentRestPage, totalPages), [currentRestPage, totalPages]);
  const paginatedDocuments = React.useMemo(() => {
    return filteredDocuments;
  }, [filteredDocuments]);

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
    const origin = new URL(webUrl).origin;
    const trimmedFileUrl = serverRelativeUrl.trim();
    const absoluteFileUrl = /^https?:\/\//i.test(trimmedFileUrl)
      ? trimmedFileUrl
      : `${origin}${trimmedFileUrl.startsWith('/') ? trimmedFileUrl : `/${trimmedFileUrl}`}`;
    const normalizedFileType = (fileType || '').toLowerCase();
    if (isAudioFileType(normalizedFileType)) return [];
    const isVideoFile = ['mp4', 'mov', 'avi', 'wmv'].indexOf(normalizedFileType) !== -1;
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

    if (isVideoFile) {
      candidates.push({ isVideo: true, posterUrl: null, sourceUrl: absoluteFileUrl, useVideoElement: true });
    }
    return candidates.slice(0, 3);
  }, [props.context]);

  React.useEffect(() => {
    setCurrentPage(1);
    setCurrentRestPage(1);
  }, [selectedFilters]);

  React.useEffect(() => {
    if (currentRestPage > totalPages) {
      setCurrentRestPage(totalPages);
      setCurrentPage(totalPages);
    }
  }, [currentRestPage, totalPages]);

  React.useEffect(() => {
    setThumbnailAttemptByDocument({});
    setThumbnailLoadedByDocument({});
  }, [props.categoryTitle, currentRestPage, viewMode]);

  React.useEffect(() => {
    if (!activeFilterKey) return;

    const refinerProperty = CATEGORY_FILTER_SEARCH_PROPERTIES[activeFilterKey];
    if (!refinerProperty) return;

    let isDisposed = false;
    const requestId = categoryFilterRefinerRequestRef.current + 1;
    categoryFilterRefinerRequestRef.current = requestId;
    setFilterRefinerLoadingKey(activeFilterKey);

    const fetchFilterRefinerCounts = async (): Promise<void> => {
      try {
        const webUrl = props.context.pageContext.web.absoluteUrl;
        const normalizedCategory = normalizeCategoryMatchText(props.categoryTitle);
        const searchLabel = CATEGORY_SEARCH_LABELS[normalizedCategory];
        if (!searchLabel) return;

        const searchService = getSearchServiceInstance(
          props.context.spHttpClient,
          webUrl,
          LIBRARY_NAMES.kmDataHub
        );
        const rootFolderUrl = await searchService.getLibraryRootFolderUrl();
        const queryParts = [
          rootFolderUrl ? `Path:"${rootFolderUrl}/*"` : '',
          `${SEARCH_PROPERTIES.documentType}:"${escapeSearchFilterValue(searchLabel)}"`,
          `${SEARCH_PROPERTIES.status}:Active`,
          ...getCategoryFilterSearchConditions(selectedFilters, activeFilterKey)
        ].filter(Boolean);
        const query = queryParts.join(' AND ');
        const refinerUrl =
          `${webUrl}/_api/search/query` +
          `?querytext='${encodeURIComponent(query)}'` +
          `&rowlimit=1` +
          `&selectproperties='${SEARCH_PROPERTIES.listItemId}'` +
          `&refiners='${refinerProperty}'` +
          `&trimduplicates=false` +
          `&clienttype='ContentSearchRegular'`;

        const response = await props.context.spHttpClient.get(refinerUrl, SPHttpClient.configurations.v1);
        if (!response.ok) {
          return;
        }

        const data = await response.json();
        const refiners =
          data?.PrimaryQueryResult?.RefinementResults?.Refiners ??
          data?.d?.query?.PrimaryQueryResult?.RefinementResults?.Refiners?.results ??
          data?.d?.query?.PrimaryQueryResult?.RefinementResults?.Refiners ??
          [];
        const refiner = (refiners?.results || refiners || []).find((item: any) =>
          String(item.Name || '').toLowerCase() === refinerProperty.toLowerCase()
        );
        const entries = refiner?.Entries?.results || refiner?.Entries || [];
        const counts = entries.reduce((accumulator: Record<string, number>, entry: any) => {
          const title = normalizeText(entry.RefinementName || entry.RefinementValue || '');
          const count = Number(String(entry.RefinementCount || '0').replace(/,/g, ''));
          if (title) {
            accumulator[normalizeRefinerKey(title)] = count;
          }
          return accumulator;
        }, {});

        if (!isDisposed && categoryFilterRefinerRequestRef.current === requestId) {
          setFilterRefinerCounts((current) => ({
            ...current,
            [activeFilterKey]: counts
          }));
        }
      } catch {
        return;
      } finally {
        if (!isDisposed && categoryFilterRefinerRequestRef.current === requestId) {
          setFilterRefinerLoadingKey(null);
        }
      }
    };

    void fetchFilterRefinerCounts();

    return () => {
      isDisposed = true;
    };
  }, [activeFilterKey, props.categoryTitle, props.context, selectedFilters]);

  const filterOptions = React.useMemo(() => {
    const sharedOptions = (props.filterGroups || []).reduce<Partial<Record<TFilterKey, ISharedFilterOption[]>>>(
      (options, group) => {
        const normalizedKey = group.key === 'region' ? 'geography' : group.key;
        if (
          normalizedKey !== 'businessUnit' &&
          normalizedKey !== 'department' &&
          normalizedKey !== 'client' &&
          normalizedKey !== 'geography' &&
          normalizedKey !== 'therapyArea' &&
          normalizedKey !== 'diseaseArea'
        ) {
          return options;
        }

        if (normalizedKey === 'businessUnit' || normalizedKey === 'department') {
          return options;
        }

        options[normalizedKey] = (group.children || [])
          .map((option) => ({
            title: normalizeText(option.title),
            count: 0
          }))
          .filter((option) => option.title)
          .sort((first, second) => first.title.localeCompare(second.title));

        return options;
      },
      {}
    );

    const getTermStoreOptions = (key: TFilterKey): ISharedFilterOption[] | null => {
      if (key === 'businessUnit' && buDepartmentTree.length > 0) {
        return getBusinessUnitOptionsFromTree(buDepartmentTree);
      }

      if (key === 'department' && buDepartmentTree.length > 0) {
        return getDepartmentOptionsFromTree(buDepartmentTree, selectedFilters.businessUnit);
      }

      if (!taxonomyOptions) {
        return null;
      }

      if (key === 'businessUnit') {
        const buTerms = taxonomyOptions.buDepartment.filter((term) => getTermHierarchyLevel(term) === 0);
        return buTerms.length > 0 ? getUniqueTermOptions(buTerms) : null;
      }

      if (key === 'department') {
        const departmentTerms = taxonomyOptions.buDepartment.filter((term) => {
          if (getTermHierarchyLevel(term) <= 0) {
            return false;
          }

          if (!selectedFilters.businessUnit) {
            return true;
          }

          const path = normalizeText(term.path || term.label);
          const firstPathSegment = path.split(/\s*>\s*/)[0] || '';
          return doesFilterValueMatch(firstPathSegment, selectedFilters.businessUnit);
        });
        return departmentTerms.length > 0 ? getUniqueTermOptions(departmentTerms) : null;
      }

      if (key === 'client') return getUniqueTermOptions(taxonomyOptions.client);
      if (key === 'geography') return getUniqueTermOptions(taxonomyOptions.geography);
      if (key === 'therapyArea') return getUniqueTermOptions(taxonomyOptions.therapyArea);
      if (key === 'diseaseArea') return getUniqueTermOptions(taxonomyOptions.diseaseArea);
      return null;
    };

    return FILTER_LABELS.reduce<Partial<Record<TFilterKey, ISharedFilterOption[]>>>((options, filter) => {
      if (!filter.key) return options;

      const refinerCounts = filterRefinerCounts[filter.key];
      const getRefinerCount = (title: string): number =>
        refinerCounts ? (refinerCounts[normalizeRefinerKey(title)] || 0) : 0;
      const counts = new Map<string, number>();
      activeDocuments.forEach((document) => {
        const value = getDocumentFilterValue(document, filter.key!);
        if (!value) return;
        counts.set(value, (counts.get(value) || 0) + 1);
      });

      if (sharedOptions[filter.key]) {
        options[filter.key] = sharedOptions[filter.key]?.map((option) => ({
          ...option,
          count: getRefinerCount(option.title)
        }));
        return options;
      }

      const termStoreOptions = getTermStoreOptions(filter.key);

      options[filter.key] = termStoreOptions
        ? termStoreOptions.map((option) => ({
          ...option,
          count: getRefinerCount(option.title)
        }))
        : Array.from(counts.entries())
          .map(([title, count]) => ({
            title,
            count: refinerCounts ? getRefinerCount(title) : count
          }))
          .sort((first, second) => first.title.localeCompare(second.title));
      return options;
    }, {});
  }, [activeDocuments, buDepartmentTree, filterRefinerCounts, getDocumentFilterValue, props.filterGroups, selectedFilters.businessUnit, taxonomyOptions]);

  const hasSelectedFilters = Object.values(selectedFilters).some(Boolean);

  const handleChipClick = (key: TFilterKey | null): void => {
    if (!key) {
      setSelectedFilters({});
      setActiveFilterKey(null);
      return;
    }

    setActiveFilterKey((current) => current === key ? null : key);
  };

  const handleFilterOptionClick = (key: TFilterKey, option: string): void => {
    setSelectedFilters((current) => {
      const nextValue = current[key] === option ? null : option;
      const nextFilters = {
        ...current,
        [key]: nextValue
      };

      if (key === 'businessUnit' && current.businessUnit !== nextValue) {
        nextFilters.department = null;
      }

      return nextFilters;
    });
    setActiveFilterKey(null);
  };

  const renderFilterDropdown = (key: TFilterKey | null): JSX.Element | null => {
    if (!key || activeFilterKey !== key) return null;
    const currentOptions = filterOptions[key] || [];
    const label = FILTER_LABELS.find((filter) => filter.key === key)?.label || '';

    return (
      <div className={styles.filterDropdown}>
        <div className={styles.filterDropdownHeader}>
          <span>{label}</span>
          <div className={styles.filterHeaderActions}>
            {selectedFilters[key] && (
              <button type="button" className={styles.filterClear} onClick={() => handleFilterOptionClick(key, selectedFilters[key]!)}>
                Clear
              </button>
            )}
            <button type="button" className={styles.filterClose} onClick={() => setActiveFilterKey(null)} aria-label="Close filter">×</button>
          </div>
        </div>
        {filterRefinerLoadingKey === key && (
          <div className={styles.filterEmpty}>Loading counts...</div>
        )}
        {currentOptions.length > 0 ? currentOptions.map((option) => {
          const isSelected = selectedFilters[key] === option.title;
          const isHierarchyChild = key === 'department' && option.level === 1;
          return (
            <button
              key={option.title}
              type="button"
              className={`${styles.filterOption} ${isSelected ? styles.filterOptionActive : ''}`}
              style={isHierarchyChild ? { paddingLeft: 26 } : undefined}
              onClick={() => handleFilterOptionClick(key, option.title)}
            >
              <span className={styles.filterOptionTitle}>
                {isHierarchyChild ? `› ${option.title}` : option.title}
              </span>
              <span className={styles.filterOptionCount}>{isSelected ? totalCount : option.count}</span>
            </button>
          );
        }) : (
          <div className={styles.filterEmpty}>No options found</div>
        )}
      </div>
    );
  };

  const handleDownload = async (document: ICategoryDocument): Promise<void> => {
    if (!document.fileRef && !document.fileUrl) {
      return;
    }
    if (downloadingDocumentIds[document.id]) {
      return;
    }

    const fileName = document.fileName || document.title || 'Document';
    setDownloadingDocumentIds((current) => ({ ...current, [document.id]: true }));

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      await downloadSharePointFile(
        props.context,
        webUrl,
        document.fileRef || document.fileUrl,
        fileName
      );
    } catch (error) {
      console.error('Category document download error:', error);
    }

    window.setTimeout(() => {
      setDownloadingDocumentIds((current) => {
        const next = { ...current };
        delete next[document.id];
        return next;
      });
    }, 1200);
  };

  const renderDocumentCard = (document: ICategoryDocument): JSX.Element => {
    if (viewMode === 'grid') {
      const thumbnailCandidates = getThumbnailCandidates(document.fileRef, document.fileType, document.fileUniqueId);
      const activeThumbnailIndex = thumbnailAttemptByDocument[document.id] ?? 0;
      const hasThumbnailFailed = activeThumbnailIndex < 0;
      const fallbackThumbnailIndex = hasThumbnailFailed ? 0 : activeThumbnailIndex;
      const activeThumbnailCandidate = hasThumbnailFailed ? undefined : thumbnailCandidates[fallbackThumbnailIndex];
      const isVideoThumbnail = typeof activeThumbnailCandidate !== 'string' && !!activeThumbnailCandidate?.isVideo;
      const shouldRenderVideoElement = isVideoThumbnail && !!activeThumbnailCandidate?.useVideoElement;
      const activeThumbnailUrl = typeof activeThumbnailCandidate === 'string' ? activeThumbnailCandidate : '';
      const videoThumbnailSource = typeof activeThumbnailCandidate !== 'string' ? activeThumbnailCandidate?.sourceUrl || document.fileRef : '';
      const isExcelPreview = isExcelFileType(document.fileType);
      const isAudioPreview = isAudioFileType(document.fileType);
      const excelPreviewUrl = isExcelPreview && document.fileUniqueId
        ? getExcelThumbnailUrl(props.context.pageContext.web.absoluteUrl, KM_REVIEW_HUB_DRIVE_ID, document.fileUniqueId)
        : '';
      const drivePreviewUrl = !isExcelPreview && !isAudioPreview && document.fileUniqueId && activeThumbnailIndex >= 0
        ? getDriveItemThumbnailUrl(props.context.pageContext.web.absoluteUrl, KM_REVIEW_HUB_DRIVE_ID, document.fileUniqueId)
        : '';
      const shouldShowFallback = !isExcelPreview && !isAudioPreview && !isVideoThumbnail && (activeThumbnailIndex < 0 || !activeThumbnailUrl);
      const usesDocumentPreviewCrop = ['PDF', 'DOC', 'DOCX', 'PPT', 'PPTX', 'XLS', 'XLSX'].indexOf(document.fileType.toUpperCase()) !== -1;
      const isPresentationPreview = ['PPT', 'PPTX'].indexOf(document.fileType.toUpperCase()) !== -1;
      const isWordPreview = ['DOC', 'DOCX'].indexOf(document.fileType.toUpperCase()) !== -1;

      return (
        <article key={document.id} className={styles.documentGridCard}>
          <div className={styles.gridThumbnailSection}>
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
                alt={document.title}
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
                    document.fileUniqueId || '',
                    'source'
                  );
                  const largeFallbackUrl = getExcelThumbnailUrl(
                    props.context.pageContext.web.absoluteUrl,
                    KM_REVIEW_HUB_DRIVE_ID,
                    document.fileUniqueId || '',
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
                    document.fileUniqueId || '',
                    'source'
                  );
                  const largeFallbackUrl = getDriveItemThumbnailUrl(
                    props.context.pageContext.web.absoluteUrl,
                    KM_REVIEW_HUB_DRIVE_ID,
                    document.fileUniqueId || '',
                    'large'
                  );
                  if (image.src !== sourceFallbackUrl && image.src !== largeFallbackUrl) {
                    image.src = sourceFallbackUrl;
                  } else if (image.src !== largeFallbackUrl) {
                    image.src = largeFallbackUrl;
                  } else {
                    setThumbnailAttemptByDocument(prev => ({ ...prev, [document.id]: -1 }));
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
                onLoad={() => setThumbnailLoadedByDocument(prev => ({ ...prev, [document.id]: true }))}
                onError={(event) => {
                  const image = event.currentTarget;
                  const candidates = getThumbnailCandidates(document.fileRef, document.fileType, document.fileUniqueId);
                  const currentIndex = candidates.indexOf(image.src);
                  const nextIndex = currentIndex >= 0 && currentIndex < candidates.length - 1 ? currentIndex + 1 : fallbackThumbnailIndex + 1;
                  const nextCandidate = candidates[nextIndex];
                  if (typeof nextCandidate === 'string') {
                    image.src = nextCandidate;
                    setThumbnailAttemptByDocument(prev => ({ ...prev, [document.id]: nextIndex }));
                  } else if (nextCandidate?.isVideo) {
                    setThumbnailAttemptByDocument(prev => ({ ...prev, [document.id]: nextIndex }));
                  } else {
                    image.onerror = null;
                    setThumbnailAttemptByDocument(prev => ({ ...prev, [document.id]: -1 }));
                  }
                  event.preventDefault();
                }}
              />
            )}
            {shouldShowFallback && (
              <div className={styles.thumbnailFallback}>
                <div className={styles.fileTypeIcon}>{getFileTypeIcon(document.docIcon || document.fileType)}</div>
                <span className={styles.fileTypeLabel}>{document.docIcon || document.fileType || 'FILE'}</span>
              </div>
            )}
          </div>
          <div className={styles.gridCardBody}>
            <div className={styles.gridCardMeta}>
              <span className={styles.authorChip}>{document.author || 'Internal'}</span>
              <span className={styles.sourceType}>{document.docIcon || document.fileType || document.documentType}</span>
            </div>
            <h2 className={styles.gridDocumentTitle}>{document.title}</h2>
            <p className={styles.searchResultDescription}>{document.description}</p>
            <div className={styles.gridCardFooter}>
              <span>{document.date}</span>
              <span className={`${styles.statusBadge} ${getStatusClassName(document.status)}`}>{getDisplayStatus(document.status)}</span>
            </div>
            <div className={styles.gridActionButtons}>
              <button type="button" className={styles.searchResultViewButton} onClick={() => props.onViewDocument?.(document.id)}>View</button>
              {!props.isLearner && (
                <button
                  type="button"
                  className={styles.searchResultDownloadButton}
                  disabled={(!document.fileUrl && !document.fileRef) || !!downloadingDocumentIds[document.id]}
                  onClick={() => { handleDownload(document); }}
                >
                  {downloadingDocumentIds[document.id] ? 'Downloading...' : 'Download'}
                </button>
              )}
            </div>
          </div>
        </article>
      );
    }

    return (
      <article key={document.id} className={styles.documentRow}>
        <div className={styles.listContent}>
          <div className={styles.searchResultContent}>
            <h2 className={styles.searchResultTitle}>{document.title}</h2>
            <p className={styles.searchResultDescription}>{document.description}</p>
            <div className={styles.searchResultStatusLine}>
              <span className={styles.searchResultStatusLabel}>Status :</span>
              <span className={`${styles.statusBadge} ${getStatusClassName(document.status)}`}>{getDisplayStatus(document.status)}</span>
            </div>
          </div>
          <div className={styles.searchResultDivider} aria-hidden="true" />
          <div className={styles.searchResultMeta}>
            <strong>{document.author || 'Internal'}</strong>
            <span className={styles.sourceDate}>{document.date}</span>
          </div>
          <div className={styles.searchResultActions}>
            <button type="button" className={styles.searchResultViewButton} onClick={() => props.onViewDocument?.(document.id)}>View</button>
            {!props.isLearner && (
              <button
                type="button"
                className={styles.searchResultDownloadButton}
                disabled={(!document.fileUrl && !document.fileRef) || !!downloadingDocumentIds[document.id]}
                onClick={() => { handleDownload(document); }}
              >
                {downloadingDocumentIds[document.id] ? 'Downloading...' : 'Download'}
              </button>
            )}
          </div>
        </div>
      </article>
    );
  };

  return (
    <section className={styles.categoryPage}>
      <IKShellHeader
        userName={props.userName}
        userEmail={props.userEmail}
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
      />

      <main className={styles.pageBody}>
        <div className={styles.contentHeader}>
          <h1 className={styles.pageTitle}>Category - {props.categoryTitle}</h1>
        </div>

        <div className={`${styles.filterChipRow} ${activeFilterKey ? styles.filterChipRowOpen : ''}`}>
          {FILTER_LABELS.map((filter) => {
            const isActive = filter.key === null
              ? !hasSelectedFilters && !activeFilterKey
              : activeFilterKey === filter.key || Boolean(selectedFilters[filter.key]);
            return (
              <div key={filter.key || 'all'} className={styles.filterChipWrap}>
                <button type="button" className={`${styles.filterChip} ${isActive ? styles.filterChipActive : ''}`} onClick={() => handleChipClick(filter.key)}>
                  <span className={styles.filterChipLabel}>{filter.label}</span>
                  {filter.key && <span className={styles.filterChipChevron}>▼</span>}
                </button>
                {renderFilterDropdown(filter.key)}
              </div>
            );
          })}
          <div className={styles.viewModeToggles}>
            <button
              type="button"
              className={`${styles.viewModeBtn} ${viewMode === 'list' ? styles.viewModeBtnActive : ''}`}
              onClick={() => setViewMode('list')}
              title="List View"
            >
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" />
              </svg>
            </button>
            <button
              type="button"
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

        {hasSelectedFilters && (
          <div className={styles.selectedFilters}>
            {FILTER_LABELS.filter((filter) => filter.key && selectedFilters[filter.key]).map((filter) => (
              <button key={filter.key} type="button" className={styles.selectedFilterPill} onClick={() => filter.key && handleFilterOptionClick(filter.key, selectedFilters[filter.key]!)}>
                <span>{filter.label}</span>
                <strong>{filter.key ? selectedFilters[filter.key] : ''}</strong>
                <span aria-hidden="true">×</span>
              </button>
            ))}
          </div>
        )}

        <div className={styles.paginationHeader}>
          {isLoading
            ? 'Loading matches...'
            : countLoading
            ? <><span className={styles.countSpinner} aria-hidden="true" /> Counting matches...</>
            : `Showing ${categoryDisplayStart}-${categoryDisplayEnd} of ${categoryDisplayCount} matching assets`}
        </div>

        <div className={styles.scrollContent}>
          <div className={styles.contentWrap}>
            {isLoading ? (
              <div className={styles.stateMessage}>Loading documents...</div>
            ) : filteredDocuments.length === 0 ? (
              <div className={styles.stateMessage}>No {props.categoryTitle.toLowerCase()} documents found.</div>
            ) : (
              <div className={viewMode === 'grid' ? styles.documentGrid : styles.documentList}>
                {paginatedDocuments.map(renderDocumentCard)}
              </div>
            )}
          </div>

          {!isLoading && filteredDocuments.length > 0 && (
            <div className={styles.paginationFooter}>
              <button
                type="button"
                disabled={currentRestPage === 1}
                onClick={() => void handlePageChange(currentRestPage - 1)}
                className={styles.pageArrowBtn}
              >
                Previous
              </button>
              <div className={styles.pageNumbers}>
                {paginationItems.map((item) =>
                  typeof item === 'number' ? (
                    <button
                      type="button"
                      key={item}
                      onClick={() => void handlePageChange(item)}
                      className={`${styles.pageNumberBtn} ${currentRestPage === item ? styles.pageNumberBtnActive : ''}`}
                    >
                      {item}
                    </button>
                  ) : (
                    <span key={item} className={styles.pageEllipsis}>...</span>
                  )
                )}
              </div>
              <button
                type="button"
                disabled={currentRestPage === totalPages}
                onClick={() => void handlePageChange(currentRestPage + 1)}
                className={styles.pageArrowBtn}
              >
                Next
              </button>
            </div>
          )}
        </div>
      </main>

      <IKShellFooter className={styles.fixedShellFooter} onBackHome={props.onHomeOpen || (() => { pushPageUrl(NAV_PATHS.home); })} />
    </section>
  );
};

export default CategoryDocumentsPage;
