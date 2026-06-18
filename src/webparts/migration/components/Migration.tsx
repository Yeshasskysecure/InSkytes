import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { IMigrationProps } from './IMigrationProps';
import GenericSearchDropdown, {
  buildKnowledgeSearchFilters,
  buildKnowledgeSearchFiltersKey,
  type FilterState,
  ISearchResultsChangeMeta
} from './search/FilterDropdown/GenericSearchDropdown';
import { KnowledgeHubSection } from './home/KnowledgeHubSection/KnowledgeHubSection';
import { MainContentRouter } from './layout/MainContentRouter/MainContentRouter';
import { AboutCategorySection } from '../pages/AboutPage/AboutPage';
import { BusinessUnitDetailPage } from '../pages/BusinessUnitDetailPage/BusinessUnitDetailPage';
import RecentlyPublishedSection from './home/RecentlyPublishedSection/RecentlyPublishedSection';
import { SIDEBAR_GROUP_IDS } from './access.constants';
import { MyDocuments } from '../pages/MyDocuments/MyDocuments';
import { MyBookmarks } from '../pages/MyBookmarks/MyBookmarks';
import { RenameBusinessUnitDialog } from './businessUnit/RenameBusinessUnitDialog/RenameBusinessUnitDialog';
import { FileUpload } from './upload/FileUpload/FileUpload';
import { DocumentDetailPage } from '../pages/DocumentDetailPage/DocumentDetailPage';
import {
  detectCurrentPage,
  parseUrlParams,
  pushPageUrl,
  openAppPageInNewTab,
  LIST_PATHS,
  NAV_PATHS,
  getItemIdFromHash
} from '../services/permalinkService';
import { ViewAllDocumentsPage } from '../pages/ViewAllDocumentsPage/ViewAllDocumentsPage';
import { CategoryDocumentsPage } from '../pages/CategoryDocumentsPage/CategoryDocumentsPage';
import { IKShellFooter, IKShellHeader } from './shell/IKShellChrome';
import { fetchAllTaxonomyOptions, ITaxonomyFieldOptions } from '../services/TaxonomyService';
import { CACHE_KEYS, COLUMN_NAMES, GROUP_IDS, IFRAME_URLS, KM_REVIEW_HUB_DRIVE_ID, LIBRARY_NAMES, PAGE_SIZES, SITE_RELATIVE_URL } from '../config/appConfig';
import { ConfigService } from '../services/ConfigService';
import { IKnowledgeSearchPageInfo } from '../services/KnowledgeSearchApiClient';
import { openOutlookCompose } from '../utils/contactActions';
import { downloadSharePointFile } from '../utils/fileDownload';
import { getDriveItemThumbnailUrl, getExcelThumbnailUrl } from '../services/SharePointSearchService';
import styles from './Migration.module.scss';

type TSidebarView = 'home' | 'profile' | 'audit-log' | 'analytics' | 'kmartifacts' | 'km-library' | 'my-documents' | 'my-bookmarks' | 'published' | 'all-documents';
type TSearchSuggestionItem = {
  value: string;
  source: 'title' | 'query';
  count?: number;
};
type TSearchFilterKey =
  | 'fileType'
  | 'businessUnit'
  | 'department'
  | 'documentType'
  | 'client'
  | 'region'
  | 'therapyArea'
  | 'diseaseArea';
type TSearchFilterOption = {
  title: string;
  value?: string;
  count: number;
  level?: number;
  parentTitle?: string;
};
type TSearchFilterGroup = {
  key: TSearchFilterKey;
  title: string;
  children: TSearchFilterOption[];
};

const buildSearchPageBackendFilters = (
  filters: Partial<Record<TSearchFilterKey, string | null>>
): FilterState => ({
  fileType: filters.fileType || null,
  businessUnit: filters.businessUnit || null,
  department: filters.department || null,
  documentType: filters.documentType || null,
  client: filters.client || null,
  region: filters.region || null,
  therapyArea: filters.therapyArea || null,
  diseaseArea: filters.diseaseArea || null,
});

type TSearchResultItem = {
  id: number;
  title: string;
  fileName?: string;
  contributor?: string;
  updated?: string;
  description?: string;
  fileUrl?: string;
  fileType?: string;
  serverRelativeUrl?: string;
  fileUniqueId?: string;
  businessUnit?: string;
  department?: string;
  documentType?: string;
  client?: string;
  region?: string;
  therapyArea?: string;
  diseaseArea?: string;
  status?: string;
};
type TSearchFilterChip = {
  label: string;
  key: TSearchFilterKey | null;
};
type TSearchChipRowVariant = 'search' | 'results';
type TSearchResultsViewMode = 'list' | 'grid';
type TSearchSortOrder = 'relevance' | 'newest' | 'oldest';
type TSearchPaginationToken = number | 'ellipsis';

const buildSearchPaginationTokens = (currentPage: number, pageCount: number): TSearchPaginationToken[] => {
  const safeCurrentPage = Math.min(Math.max(currentPage, 1), Math.max(pageCount, 1));
  const pages = new Set<number>([1, pageCount]);

  for (let page = safeCurrentPage - 2; page <= safeCurrentPage + 2; page += 1) {
    if (page >= 1 && page <= pageCount) {
      pages.add(page);
    }
  }

  const orderedPages = Array.from(pages).sort((left, right) => left - right);
  const tokens: TSearchPaginationToken[] = [];

  orderedPages.forEach((page) => {
    const previous = tokens[tokens.length - 1];
    if (typeof previous === 'number' && page - previous > 1) {
      tokens.push('ellipsis');
    }
    tokens.push(page);
  });

  return tokens;
};

const cleanSearchFilterTextValue = (value?: string | null): string => {
  const normalizedValue = String(value || '').trim();
  if (!normalizedValue) {
    return '';
  }

  return normalizedValue
    .split(/;#|;/)
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const parts = entry.split('|').map((part) => part.trim()).filter(Boolean);
      const label = parts[0] || entry;
      return /^\d+$/.test(label) ? '' : label;
    })
    .filter(Boolean)
    .filter((entry, index, entries) => entries.indexOf(entry) === index)
    .join(', ');
};

const normalizeSearchFilterCountKey = (value?: string | null): string => {
  const normalizedValue = cleanSearchFilterTextValue(value)
    .trim()
    .toLowerCase()
    .replace(/＆/g, '&')
    .replace(/[_-]+/g, ' ')
    .replace(/\s+/g, ' ');

  if (['na', 'n a', 'n/a', 'notapplicable', 'not applicable'].indexOf(normalizedValue.replace(/\s+/g, '')) !== -1) {
    return 'not applicable';
  }

  return normalizedValue;
};

const isSearchPageFacetDebugEnabled = (): boolean => {
  try {
    if (typeof window === 'undefined') {
      return false;
    }

    return window.localStorage?.getItem('IKNOWLEDGE_DEBUG_LOGS') === 'true' ||
      new URLSearchParams(window.location.search).get('ikSearchDebug') === 'true';
  } catch {
    return false;
  }
};

const SEARCH_FRONTEND_BUNDLE_MARKER = 'search-backend-facets-only';

const getSearchStatusBadgeClassName = (status?: string): string => {
  const normalizedStatus = (status || '').trim().toLowerCase();

  if (normalizedStatus === 'active') {
    return styles.statusBadgeActive;
  }

  if (normalizedStatus === 'under review' || normalizedStatus === 'underreview') {
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

const getSearchDisplayStatus = (status?: string): string => {
  const trimmedStatus = (status || '').trim();
  return trimmedStatus.toLowerCase() === 'rejected' ? 'Reject' : trimmedStatus;
};

const KM_ARTIFACTS_LIBRARY = LIBRARY_NAMES.kmDataHub;
const SUB_DEPARTMENT_INTERNAL_NAME = 'SubDepartment';
const SUB_DEPARTMENT_DISPLAY_NAME = 'Sub-Department';
const STATUS_INTERNAL_NAME = COLUMN_NAMES.status;
const PUBLISHED_INTERNAL_NAME = COLUMN_NAMES.published;
const PUBLISHED_DISPLAY_NAME = 'Published';
const ADMIN_GROUP_IDS = [SIDEBAR_GROUP_IDS.admins, SIDEBAR_GROUP_IDS.approvers];
const CONTRIBUTOR_GROUP_IDS = [SIDEBAR_GROUP_IDS.contributors];
const LEARNER_GROUP_IDS = [SIDEBAR_GROUP_IDS.learners];

const normalizeFieldKey = (value?: string): string => (value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
const normalizeGroupId = (value?: string): string => (value || '').trim().toLowerCase();
const OPEN_ROUTED_DOCUMENT_EVENT = 'ikn:open-routed-document';
const UPLOAD_SUCCESS_RESTORE_STORAGE_KEY = CACHE_KEYS.uploadSuccessRestore;
const getRoutedDocumentIdFromLocation = (): number | null => {
  const hashItemId = getItemIdFromHash(window.location.hash);
  if (hashItemId !== null) {
    return hashItemId;
  }

  const searchParams = new URLSearchParams(window.location.search);
  const itemIdText = searchParams.get('docId') || searchParams.get('id');
  if (itemIdText) {
    const itemId = parseInt(itemIdText, 10);
    if (Number.isFinite(itemId)) {
      return itemId;
    }
  }

  return null;
};

const getCategoryTitleFromUrlValue = (value: string): string | null => {
  const normalized = String(value || '').trim().toLowerCase().replace(/\s+/g, '-');
  const categoryMap: Record<string, string> = {
    // Current categories
    'case-studies': 'Case Studies',
    'case-study': 'Case Studies',
    proposals: 'Proposals',
    proposal: 'Proposals',
    capabilities: 'Capabilities',
    capability: 'Capabilities',
    'lessons-learned': 'Lessons Learned',
    'lesson-learned': 'Lessons Learned',
    'knowledge-sharing-sessions': 'Knowledge Sharing Sessions',
    'knowledge-sharing-session': 'Knowledge Sharing Sessions',
    ks2: 'Knowledge Sharing Sessions',
  };

  return categoryMap[normalized] || null;
};

const getCategoryUrlValue = (categoryTitle: string): string => {
  const map: Record<string, string> = {
    'Case Studies': 'case-studies',
    'Capabilities': 'capabilities',
    'Proposals': 'proposals',
    'Lessons Learned': 'lessons-learned',
    'Knowledge Sharing Sessions': 'knowledge-sharing-sessions',
  };
  return map[categoryTitle] || categoryTitle.toLowerCase().replace(/\s+/g, '-');
};

const Migration: React.FC<IMigrationProps> = (props) => {
  const { context, projectId: initialProjectId } = props;
  const configServiceRef = React.useRef<ConfigService | null>(null);
  if (!configServiceRef.current && context) {
    configServiceRef.current = new ConfigService(context);
  }
  const { docId: initialDocId, query: initialQuery, projectId: urlProjectId, category: initialCategory } = parseUrlParams();
  const [activePage, setActivePage] = React.useState<string>('about');
  const [selectedBU, setSelectedBU] = React.useState<string | null>(() => detectCurrentPage() === 'businessUnits' ? 'Business Unit Name' : null);
  const [activeSidebarView, setActiveSidebarView] = React.useState<TSidebarView>(() => {
    const page = detectCurrentPage();
    const map: Record<string, TSidebarView> = {
      home: 'home',
      businessUnits: 'home',
      categories: 'home',
      search: 'home',
      asset: initialDocId ? 'all-documents' : 'published',
      bookmark: 'my-bookmarks',
      myDocuments: 'my-documents',
      analytics: 'analytics',
      auditLog: 'audit-log',
      kmReviewHub: 'all-documents',
      kmHarvestHub: 'kmartifacts',
      kmLibrary: 'km-library',
      legacy: 'home'
    };
    return map[page] || 'home';
  });
  const [isReviewerFrameLoaded, setIsReviewerFrameLoaded] = React.useState(false);
  const [analyticsFrameLoadedByPage, setAnalyticsFrameLoadedByPage] = React.useState<Record<string, boolean>>({});
  const [isRenameBusinessUnitOpen, setIsRenameBusinessUnitOpen] = React.useState(false);
  const [routedDocumentId, setRoutedDocumentId] = React.useState<number | null>(
    initialDocId ?? getRoutedDocumentIdFromLocation()
  );
  const [projectId, setProjectId] = React.useState<string | null>(urlProjectId || initialProjectId || null);
  const [searchPageQuery, setSearchPageQuery] = React.useState<string>(initialQuery);
  const [searchPageDraftQuery, setSearchPageDraftQuery] = React.useState<string>(initialQuery);
  const [searchPageSuggestions, setSearchPageSuggestions] = React.useState<TSearchSuggestionItem[]>([]);
  const [searchPageResults, setSearchPageResults] = React.useState<TSearchResultItem[]>([]);
  const [searchPageTotalCount, setSearchPageTotalCount] = React.useState<number>(0);
  const [searchPageInfo, setSearchPageInfo] = React.useState<IKnowledgeSearchPageInfo | undefined>(undefined);
  const [isSearchPageResultsLoading, setIsSearchPageResultsLoading] = React.useState<boolean>(Boolean(initialQuery.trim()));
  const [hasSearchPageResultsResolved, setHasSearchPageResultsResolved] = React.useState<boolean>(!initialQuery.trim());
  const [searchResultsPage, setSearchResultsPage] = React.useState<number>(1);
  const [searchResultsViewMode, setSearchResultsViewMode] = React.useState<TSearchResultsViewMode>('list');
  const [searchSortOrder, setSearchSortOrder] = React.useState<TSearchSortOrder>('relevance');
  const [searchThumbnailAttemptByDocument, setSearchThumbnailAttemptByDocument] = React.useState<Record<number, number>>({});
  const [searchThumbnailLoadedByDocument, setSearchThumbnailLoadedByDocument] = React.useState<Record<number, boolean>>({});
  const [searchProviderRefreshKey, setSearchProviderRefreshKey] = React.useState<number>(0);
  const [activeSearchFilterKey, setActiveSearchFilterKey] = React.useState<TSearchFilterKey | null>(null);
  const [searchPageFilterGroups, setSearchPageFilterGroups] = React.useState<TSearchFilterGroup[]>([]);
  const [searchTaxonomyOptions, setSearchTaxonomyOptions] = React.useState<ITaxonomyFieldOptions | null>(null);
  const [searchPageFilters, setSearchPageFilters] = React.useState<Partial<Record<TSearchFilterKey, string | null>>>({});
  const [categoryDocumentsTitle, setCategoryDocumentsTitle] = React.useState<string | null>(() =>
    detectCurrentPage() === 'categories' ? getCategoryTitleFromUrlValue(initialCategory) : null
  );
  const [showHomeUploader, setShowHomeUploader] = React.useState<boolean>(() => Boolean(urlProjectId || initialProjectId));
  const [isSearchPageOpen, setIsSearchPageOpen] = React.useState<boolean>(() => {
    const page = detectCurrentPage();
    return page === 'search';
  });
  const [isLearner, setIsLearner] = React.useState(false);
  const [canAccessKmReviewHub, setCanAccessKmReviewHub] = React.useState(false);
  const [isKmAdmin, setIsKmAdmin] = React.useState(false);
  const [isReviewerAccessResolved, setIsReviewerAccessResolved] = React.useState(false);
  const [currentUserPhotoUrl, setCurrentUserPhotoUrl] = React.useState<string>('');
  const rootRef = React.useRef<HTMLDivElement>(null);
  const searchResultsScrollAreaRef = React.useRef<HTMLDivElement>(null);
  const hasSearchPageResultsLoadingStartedRef = React.useRef<boolean>(Boolean(initialQuery.trim()));
  const appReadyAnnouncedRef = React.useRef<boolean>(false);

  React.useEffect(() => {
    if (!isSearchPageOpen || typeof window === 'undefined') {
      return;
    }

    const marker = {
      marker: SEARCH_FRONTEND_BUNDLE_MARKER,
      page: 'search',
      usesBackendFacetGroupsOnly: true,
      query: searchPageQuery
    };
    (window as any).__IKNOWLEDGE_SEARCH_FRONTEND_MARKER = marker;

    if (isSearchPageFacetDebugEnabled()) {
      console.info('[Search] Frontend bundle marker', marker);
    }
  }, [isSearchPageOpen, searchPageQuery]);

  React.useEffect(() => {
    if (!document.getElementById('ik-google-fonts')) {
      const link = document.createElement('link');
      link.id = 'ik-google-fonts';
      link.rel = 'stylesheet';
      link.href = 'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap';
      document.head.appendChild(link);
    }

    const handlePopState = (): void => {
      const page = detectCurrentPage();
      const { docId, query, projectId: nextProjectId, category } = parseUrlParams();

      const map: Record<string, TSidebarView> = {
        home: 'home',
        search: 'home',
        asset: docId ? 'all-documents' : 'published',
        bookmark: 'my-bookmarks',
        businessUnits: 'home',
        categories: 'home',
        myDocuments: 'my-documents',
        analytics: 'analytics',
        auditLog: 'audit-log',
        kmReviewHub: 'all-documents',
        kmHarvestHub: 'kmartifacts',
        kmLibrary: 'km-library',
        legacy: 'home'
      };
      setActiveSidebarView(map[page] || 'home');
      setRoutedDocumentId(docId);
      setIsSearchPageOpen(page === 'search');
      setSearchPageQuery(query);
      setSearchPageDraftQuery(query);
      setProjectId(nextProjectId);
      setShowHomeUploader(Boolean(nextProjectId && docId === null && !window.history.state?.iknShowHome));
      setSearchPageResults([]);
      setSelectedBU(page === 'businessUnits' ? 'Business Unit Name' : null);
      setCategoryDocumentsTitle(page === 'categories' ? getCategoryTitleFromUrlValue(category) : null);
    };
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, []);

  React.useLayoutEffect(() => {
    const onRouteChange = (): void => {
      const parsedParams = parseUrlParams();
      const page = detectCurrentPage();
      const nextDocumentId = parsedParams.docId ?? getRoutedDocumentIdFromLocation();
      setRoutedDocumentId(nextDocumentId);
      setProjectId(parsedParams.projectId);
      setShowHomeUploader(Boolean(parsedParams.projectId && nextDocumentId === null && !window.history.state?.iknShowHome));

      if (nextDocumentId !== null) {
        setActiveSidebarView('home');
        setSelectedBU(null);
        setCategoryDocumentsTitle(null);
      } else if (page === 'businessUnits') {
        setActiveSidebarView('home');
        setSelectedBU('Business Unit Name');
        setCategoryDocumentsTitle(null);
      } else if (page === 'categories') {
        setActiveSidebarView('home');
        setSelectedBU(null);
        setCategoryDocumentsTitle(getCategoryTitleFromUrlValue(parsedParams.category));
      }
    };

    onRouteChange();
    window.addEventListener('hashchange', onRouteChange);

    const handleOpenRoutedDocument = (event: Event): void => {
      const customEvent = event as CustomEvent<{ documentId?: number }>;
      const nextDocumentId = Number(customEvent.detail?.documentId);

      if (!Number.isFinite(nextDocumentId) || nextDocumentId <= 0) {
        return;
      }

      setRoutedDocumentId(nextDocumentId);
      setActiveSidebarView('home');
      setSelectedBU(null);
      setCategoryDocumentsTitle(null);
      setIsSearchPageOpen(false);
    };

    window.addEventListener(OPEN_ROUTED_DOCUMENT_EVENT, handleOpenRoutedDocument as EventListener);

    return () => {
      window.removeEventListener('hashchange', onRouteChange);
      window.removeEventListener(OPEN_ROUTED_DOCUMENT_EVENT, handleOpenRoutedDocument as EventListener);
    };
  }, []);

  React.useEffect(() => {
    if (!document.querySelector('link[href*="fonts.googleapis.com/css2?family=Inter"]')) {
      const link = document.createElement('link');
      link.href = 'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap';
      link.rel = 'stylesheet';
      document.head.appendChild(link);
    }
  }, []);

  React.useEffect(() => {
    if (isSearchPageOpen) {
      return;
    }

    let isDisposed = false;

    const loadTaxonomyOptions = async (): Promise<void> => {
      try {
        const TAXONOMY_KEY = CACHE_KEYS.taxonomy;
        const cachedTaxonomy = (() => {
          try {
            const cached = sessionStorage.getItem(TAXONOMY_KEY);
            return cached ? JSON.parse(cached) : null;
          } catch {
            return null;
          }
        })();
        if (cachedTaxonomy) {
          if (!isDisposed) {
            setSearchTaxonomyOptions(cachedTaxonomy);
          }
          return;
        }

        const options = await fetchAllTaxonomyOptions(context);
        if (!isDisposed) {
          setSearchTaxonomyOptions(options);
        }
        try {
          sessionStorage.setItem(TAXONOMY_KEY, JSON.stringify(options));
        } catch {
          // Ignore session storage failures.
        }
      } catch (error) {
        console.warn('Unable to load taxonomy options:', error);
        if (!isDisposed) {
          setSearchTaxonomyOptions(null);
        }
      }
    };

    void loadTaxonomyOptions();

    return () => {
      isDisposed = true;
    };
  }, [context, isSearchPageOpen]);

  React.useEffect(() => {
    if (!context) return;
    let objectUrl = '';

    void (async () => {
      try {
        const graphClient = await context.msGraphClientFactory.getClient('3');
        const response = await graphClient
          .api('/me/photo/$value')
          .version('v1.0')
          .responseType(ResponseType.BLOB)
          .get();
        if (response) {
          objectUrl = URL.createObjectURL(response);
          setCurrentUserPhotoUrl(objectUrl);
        }
      } catch (error) {
        void error;
        setCurrentUserPhotoUrl('');
      }
    })();

    return () => {
      if (objectUrl) {
        URL.revokeObjectURL(objectUrl);
      }
    };
  }, [context]);

  React.useEffect(() => {
    let isDisposed = false;

    const ensureSubDepartmentColumn = async (): Promise<void> => {
      const webUrl = context.pageContext.web.absoluteUrl;

      const fieldsResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/fields?$select=InternalName,Title`,
        SPHttpClient.configurations.v1
      );

      if (!fieldsResponse.ok) {
        throw new Error(`Failed to read ${KM_ARTIFACTS_LIBRARY} fields.`);
      }

      const fieldsData = await fieldsResponse.json();
      const fields = fieldsData.value || [];
      const existingField = fields.find((field: { InternalName?: string; Title?: string }) => {
        const candidates = [field.InternalName, field.Title].map(normalizeFieldKey);
        return candidates.indexOf(normalizeFieldKey(SUB_DEPARTMENT_INTERNAL_NAME)) !== -1 ||
          candidates.indexOf(normalizeFieldKey(SUB_DEPARTMENT_DISPLAY_NAME)) !== -1;
      });

      const resolvedFieldName = existingField?.InternalName || SUB_DEPARTMENT_INTERNAL_NAME;

      if (!existingField) {
        const createFieldResponse = await context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/fields/createfieldasxml`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: JSON.stringify({
              parameters: {
                SchemaXml: `<Field Type="Text" DisplayName="${SUB_DEPARTMENT_DISPLAY_NAME}" Name="${SUB_DEPARTMENT_INTERNAL_NAME}" StaticName="${SUB_DEPARTMENT_INTERNAL_NAME}" Group="Custom Columns" />`,
                Options: 0
              }
            })
          }
        );

        if (!createFieldResponse.ok) {
          const errorText = await createFieldResponse.text();
          throw new Error(`Failed to create ${SUB_DEPARTMENT_DISPLAY_NAME}: ${createFieldResponse.status} ${errorText}`);
        }
      }

      const viewFieldsResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/DefaultView/ViewFields`,
        SPHttpClient.configurations.v1
      );

      if (!viewFieldsResponse.ok) {
        throw new Error(`Failed to read ${KM_ARTIFACTS_LIBRARY} default view fields.`);
      }

      const viewFieldsData = await viewFieldsResponse.json();
      const viewFields: string[] = viewFieldsData.Items || viewFieldsData.value || [];

      if (viewFields.indexOf(resolvedFieldName) === -1) {
        const addFieldToViewResponse = await context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/DefaultView/ViewFields/AddViewField('${resolvedFieldName}')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
            }
          }
        );

        if (!addFieldToViewResponse.ok) {
          const errorText = await addFieldToViewResponse.text();
          throw new Error(`Failed to add ${SUB_DEPARTMENT_DISPLAY_NAME} to the default view: ${addFieldToViewResponse.status} ${errorText}`);
        }
      }

      const refreshedViewFieldsResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/DefaultView/ViewFields`,
        SPHttpClient.configurations.v1
      );

      if (!refreshedViewFieldsResponse.ok) {
        throw new Error(`Failed to refresh ${KM_ARTIFACTS_LIBRARY} default view fields.`);
      }

      const refreshedViewFieldsData = await refreshedViewFieldsResponse.json();
      const refreshedViewFields: string[] = refreshedViewFieldsData.Items || refreshedViewFieldsData.value || [];
      const departmentIndex = refreshedViewFields.indexOf('Department');
      const subDepartmentIndex = refreshedViewFields.indexOf(resolvedFieldName);

      if (departmentIndex !== -1 && subDepartmentIndex !== -1 && subDepartmentIndex !== departmentIndex + 1) {
        const moveFieldResponse = await context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/DefaultView/ViewFields/moveViewFieldTo`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: JSON.stringify({
              field: resolvedFieldName,
              index: departmentIndex + 1
            })
          }
        );

        if (!moveFieldResponse.ok) {
          const errorText = await moveFieldResponse.text();
          throw new Error(`Failed to position ${SUB_DEPARTMENT_DISPLAY_NAME} after Department: ${moveFieldResponse.status} ${errorText}`);
        }
      }
    };

    const ensurePublishedColumn = async (): Promise<void> => {
      const webUrl = context.pageContext.web.absoluteUrl;

      const fieldsResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/fields?$select=InternalName,Title`,
        SPHttpClient.configurations.v1
      );

      if (!fieldsResponse.ok) {
        throw new Error(`Failed to read ${KM_ARTIFACTS_LIBRARY} fields.`);
      }

      const fieldsData = await fieldsResponse.json();
      const fields = fieldsData.value || [];
      const existingField = fields.find((field: { InternalName?: string; Title?: string }) => {
        const candidates = [field.InternalName, field.Title].map(normalizeFieldKey);
        return candidates.indexOf(normalizeFieldKey(PUBLISHED_INTERNAL_NAME)) !== -1 ||
          candidates.indexOf(normalizeFieldKey(PUBLISHED_DISPLAY_NAME)) !== -1;
      });

      const resolvedFieldName = existingField?.InternalName || PUBLISHED_INTERNAL_NAME;

      if (!existingField) {
        const createFieldResponse = await context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/fields/createfieldasxml`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: JSON.stringify({
              parameters: {
                SchemaXml: `<Field Type="DateTime" DisplayName="${PUBLISHED_DISPLAY_NAME}" Name="${PUBLISHED_INTERNAL_NAME}" StaticName="${PUBLISHED_INTERNAL_NAME}" Format="DateTime" FriendlyDisplayFormat="Disabled" Group="Custom Columns" />`,
                Options: 0
              }
            })
          }
        );

        if (!createFieldResponse.ok) {
          const errorText = await createFieldResponse.text();
          throw new Error(`Failed to create ${PUBLISHED_DISPLAY_NAME}: ${createFieldResponse.status} ${errorText}`);
        }
      }

      const viewFieldsResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/DefaultView/ViewFields`,
        SPHttpClient.configurations.v1
      );

      if (!viewFieldsResponse.ok) {
        throw new Error(`Failed to read ${KM_ARTIFACTS_LIBRARY} default view fields.`);
      }

      const viewFieldsData = await viewFieldsResponse.json();
      const viewFields: string[] = viewFieldsData.Items || viewFieldsData.value || [];

      if (viewFields.indexOf(resolvedFieldName) === -1) {
        const addFieldToViewResponse = await context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/DefaultView/ViewFields/AddViewField('${resolvedFieldName}')`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
            }
          }
        );

        if (!addFieldToViewResponse.ok) {
          const errorText = await addFieldToViewResponse.text();
          throw new Error(`Failed to add ${PUBLISHED_DISPLAY_NAME} to the default view: ${addFieldToViewResponse.status} ${errorText}`);
        }
      }

      const refreshedViewFieldsResponse = await context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/DefaultView/ViewFields`,
        SPHttpClient.configurations.v1
      );

      if (!refreshedViewFieldsResponse.ok) {
        throw new Error(`Failed to refresh ${KM_ARTIFACTS_LIBRARY} default view fields.`);
      }

      const refreshedViewFieldsData = await refreshedViewFieldsResponse.json();
      const refreshedViewFields: string[] = refreshedViewFieldsData.Items || refreshedViewFieldsData.value || [];
      const statusIndex = refreshedViewFields.indexOf(STATUS_INTERNAL_NAME);
      const publishedIndex = refreshedViewFields.indexOf(resolvedFieldName);

      if (statusIndex !== -1 && publishedIndex !== -1 && publishedIndex !== statusIndex + 1) {
        const moveFieldResponse = await context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${KM_ARTIFACTS_LIBRARY}')/DefaultView/ViewFields/moveViewFieldTo`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: JSON.stringify({
              field: resolvedFieldName,
              index: statusIndex + 1
            })
          }
        );

        if (!moveFieldResponse.ok) {
          const errorText = await moveFieldResponse.text();
          throw new Error(`Failed to position ${PUBLISHED_DISPLAY_NAME} after Status: ${moveFieldResponse.status} ${errorText}`);
        }
      }
    };

    Promise.all([
      ensureSubDepartmentColumn(),
      ensurePublishedColumn()
    ]).catch((error) => {
      if (!isDisposed) {
        console.error('Unable to ensure KM Data Hub default columns:', error);
      }
    });

    return () => {
      isDisposed = true;
    };
  }, [context]);

  const webUrl = context.pageContext.web.absoluteUrl;

  React.useEffect(() => {
    if (!context.msGraphClientFactory) {
      setCanAccessKmReviewHub(false);
      setIsKmAdmin(false);
      setIsReviewerAccessResolved(true);
      return;
    }

    let isDisposed = false;

    const fetchKmReviewHubAccess = async (): Promise<void> => {
      try {
        const graphClient = await context.msGraphClientFactory.getClient('3');
        const currentUserGroupIds = new Set<string>();
        let requestPath = '/me/transitiveMemberOf?$select=id';

        while (requestPath) {
          const membershipResponse = await graphClient.api(requestPath).version('v1.0').get();

          (membershipResponse.value || []).forEach((entry: { id?: string }) => {
            const groupId = normalizeGroupId(entry.id);
            if (groupId) {
              currentUserGroupIds.add(groupId);
            }
          });

          const nextLink = membershipResponse['@odata.nextLink'] as string | undefined;
          requestPath = nextLink
            ? nextLink
              .replace('https://graph.microsoft.com/v1.0', '')
              .replace(/^\/v1\.0/i, '')
            : '';
        }

        if (!isDisposed) {
          const hasAdminOrApproverAccess = ADMIN_GROUP_IDS.some((groupId) =>
            currentUserGroupIds.has(normalizeGroupId(groupId))
          );
          const hasKmAdminAccess = currentUserGroupIds.has(normalizeGroupId(GROUP_IDS.admins));
          const hasContributorAccess = CONTRIBUTOR_GROUP_IDS.some((groupId) =>
            currentUserGroupIds.has(normalizeGroupId(groupId))
          );
          const hasLearnerAccess = LEARNER_GROUP_IDS.some((groupId) =>
            currentUserGroupIds.has(normalizeGroupId(groupId))
          );
          const shouldUseLearnerView = hasLearnerAccess && !hasAdminOrApproverAccess && !hasContributorAccess;

          setCanAccessKmReviewHub(hasAdminOrApproverAccess);
          setIsKmAdmin(hasKmAdminAccess);
          setIsLearner(shouldUseLearnerView);
          setIsReviewerAccessResolved(true);
        }
      } catch (error) {
        console.warn('Unable to resolve KM Review Hub access:', error);
        if (!isDisposed) {
          setCanAccessKmReviewHub(false);
          setIsKmAdmin(false);
          setIsLearner(false);
          setIsReviewerAccessResolved(true);
        }
      }
    };

    void fetchKmReviewHubAccess();

    return () => {
      isDisposed = true;
    };
  }, [context]);

  const handleSidebarChange = React.useCallback((view: TSidebarView) => {
    const pathMap: Record<TSidebarView | 'km-review-hub', string> = {
      home: NAV_PATHS.home,
      profile: NAV_PATHS.home,
      'my-documents': NAV_PATHS.myDocuments,
      'my-bookmarks': NAV_PATHS.bookmark,
      'all-documents': NAV_PATHS.kmReviewHub,
      published: NAV_PATHS.asset,
      'audit-log': NAV_PATHS.auditLog,
      analytics: NAV_PATHS.analytics,
      'km-review-hub': NAV_PATHS.kmReviewHub,
      kmartifacts: NAV_PATHS.kmHarvestHub,
      'km-library': NAV_PATHS.kmLibrary
    };

    pushPageUrl(pathMap[view] || NAV_PATHS.home);
    setActiveSidebarView(view);
    setActivePage('about');
    setSearchPageQuery('');
    setSearchPageDraftQuery('');
    setSearchPageSuggestions([]);
    setSearchPageResults([]);
    setSearchResultsPage(1);
    setSelectedBU(null);
    setCategoryDocumentsTitle(null);
    setRoutedDocumentId(null);
    setIsSearchPageOpen(false);
  }, []);

  const currentUserName = React.useMemo(() => {
    const pageUser = context.pageContext.user;
    return pageUser?.displayName || pageUser?.loginName || 'User';
  }, [context.pageContext.user]);

  const currentUserEmail = React.useMemo(() => {
    const pageUser = context.pageContext.user;
    return pageUser?.email || pageUser?.loginName || '-';
  }, [context.pageContext.user]);

  const handleLogout = React.useCallback((): void => {
    window.location.href = `${webUrl}/_layouts/15/SignOut.aspx`;
  }, [webUrl]);

  const handleBackToMain = React.useCallback((): void => {
    setSelectedBU(null);
    setCategoryDocumentsTitle(null);
    setActivePage('about');
    setSearchPageQuery('');
  }, []);

  const pushHomeUrlQuietly = React.useCallback((): void => {
    const url = new URL(NAV_PATHS.home, window.location.origin);
    window.history.pushState({ sameTabNav: true }, document.title, url.toString());
  }, []);

  const navigateHomeSmooth = React.useCallback((): void => {
    setShowHomeUploader(false);
    setRoutedDocumentId(null);
    setIsSearchPageOpen(false);
    setSearchPageQuery('');
    setSearchPageDraftQuery('');
    setSearchPageSuggestions([]);
    setSearchPageResults([]);
    setSelectedBU(null);
    setCategoryDocumentsTitle(null);
    setActiveSidebarView('home');
    setActivePage('about');
    setSearchResultsPage(1);
    pushHomeUrlQuietly();
    window.scrollTo({ top: 0, behavior: 'auto' });
  }, [pushHomeUrlQuietly]);

  const handleCloseRoutedDocument = React.useCallback((): void => {
    navigateHomeSmooth();
  }, [navigateHomeSmooth]);

  const handleHeaderSearchSubmit = React.useCallback((query: string): void => {
    const trimmed = query.trim();
    const isSameSubmittedSearch =
      trimmed.toLowerCase() === searchPageQuery.trim().toLowerCase() &&
      isSearchPageOpen;

    if (!isSameSubmittedSearch) {
      pushPageUrl(NAV_PATHS.search, { query: trimmed });
    }

    setRoutedDocumentId(null);
    setSelectedBU(null);
    setCategoryDocumentsTitle(null);
    setActiveSidebarView('home');
    setIsSearchPageOpen(true);
    setSearchPageQuery(trimmed);
    setSearchPageDraftQuery(trimmed);
    setSearchPageSuggestions([]);
    if (!isSameSubmittedSearch) {
      setSearchPageResults([]);
      setSearchPageTotalCount(0);
      setSearchPageInfo(undefined);
      setHasSearchPageResultsResolved(!trimmed);
      hasSearchPageResultsLoadingStartedRef.current = Boolean(trimmed);
      setIsSearchPageResultsLoading(Boolean(trimmed));
      setSearchProviderRefreshKey((value) => value + 1);
    } else {
      setIsSearchPageResultsLoading(false);
    }
    setSearchResultsPage(1);
  }, [isSearchPageOpen, searchPageQuery]);

  const handleOpenSearchPage = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.search);
  }, []);

  const handleOpenUploader = React.useCallback((): void => {
    setShowHomeUploader(true);
  }, []);

  const handleHomeCategorySelect = React.useCallback((category: string): void => {
    const categoryTitle = String(category || '').trim();
    if (!categoryTitle) {
      return;
    }

    openAppPageInNewTab(NAV_PATHS.categories, { category: getCategoryUrlValue(categoryTitle) });
  }, []);

  const handleHomeBusinessUnitsOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.businessUnits);
  }, []);

  const handleHomeAllDocumentsOpen = React.useCallback((): void => {
    setSelectedBU(null);
    setCategoryDocumentsTitle(null);
    setIsSearchPageOpen(false);
    handleSidebarChange('all-documents');
  }, [handleSidebarChange]);

  const handleRecentlyPublishedOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.asset);
  }, []);

  const handleKmReviewHubOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.kmReviewHub);
  }, []);

  const handleHomeBookmarksOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.bookmark);
  }, []);

  const handleHomeDocumentsOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.myDocuments);
  }, []);

  const handleLearnerDocumentsOpen = isLearner ? undefined : handleHomeDocumentsOpen;

  const handleHomeContactOpen = React.useCallback((): void => {
    openOutlookCompose();
  }, []);

  const handleAuditLogOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.auditLog);
  }, []);

  const handleAnalyticsOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.analytics);
  }, []);

  const handleKmHarvestHubOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.kmHarvestHub);
  }, []);

  const handleKmLibraryOpen = React.useCallback((): void => {
    openAppPageInNewTab(NAV_PATHS.kmLibrary);
  }, []);

  const handleReviewerAuditLogOpen = canAccessKmReviewHub ? handleAuditLogOpen : undefined;
  const handleReviewerAnalyticsOpen = canAccessKmReviewHub ? handleAnalyticsOpen : undefined;
  const handleReviewerKmReviewHubOpen = canAccessKmReviewHub ? handleKmReviewHubOpen : undefined;
  const handleReviewerKmHarvestHubOpen = canAccessKmReviewHub ? handleKmHarvestHubOpen : undefined;
  const handleReviewerKmLibraryOpen = canAccessKmReviewHub ? handleKmLibraryOpen : undefined;

  React.useEffect(() => {
    const handleOpenKmHarvestHub = (): void => {
      if (canAccessKmReviewHub) {
        handleKmHarvestHubOpen();
      }
    };

    window.addEventListener('ikn:open-km-harvest-hub', handleOpenKmHarvestHub);
    return () => window.removeEventListener('ikn:open-km-harvest-hub', handleOpenKmHarvestHub);
  }, [canAccessKmReviewHub, handleKmHarvestHubOpen]);

  const handleSearchPageInputChange = React.useCallback((query: string): void => {
    setRoutedDocumentId(null);
    setSelectedBU(null);
    setCategoryDocumentsTitle(null);
    setActiveSidebarView('home');
    setIsSearchPageOpen(true);
    setSearchPageDraftQuery(query);
  }, []);

  const handleCloseSearchPage = React.useCallback((): void => {
    navigateHomeSmooth();
  }, [navigateHomeSmooth]);

  const handleShellHomeOpen = React.useCallback((): void => {
    navigateHomeSmooth();
  }, [navigateHomeSmooth]);

  React.useEffect(() => {
    if (!isReviewerAccessResolved || canAccessKmReviewHub) {
      return;
    }

    if (
      activeSidebarView !== 'all-documents' &&
      activeSidebarView !== 'audit-log' &&
      activeSidebarView !== 'analytics' &&
      activeSidebarView !== 'kmartifacts' &&
      activeSidebarView !== 'km-library'
    ) {
      return;
    }

    handleShellHomeOpen();
  }, [activeSidebarView, canAccessKmReviewHub, handleShellHomeOpen, isReviewerAccessResolved]);

  const handleSearchResultViewDocument = React.useCallback((documentId: number): void => {
    openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
  }, []);

  const handleSearchPageResultsChange = React.useCallback((
    items: TSearchResultItem[],
    isLoading: boolean,
    totalCount?: number,
    page?: IKnowledgeSearchPageInfo,
    meta?: ISearchResultsChangeMeta
  ): void => {
    if (meta) {
      const expectedQuery = searchPageQuery.trim() || '*';
      const expectedSkip = Math.max(0, (Math.max(1, searchResultsPage) - 1) * PAGE_SIZES.searchResultsPerPage);
      const expectedFiltersKey = buildKnowledgeSearchFiltersKey(
        buildKnowledgeSearchFilters(buildSearchPageBackendFilters(searchPageFilters))
      );

      if (
        meta.query !== expectedQuery ||
        meta.sort !== searchSortOrder ||
        meta.skip !== expectedSkip ||
        meta.top !== PAGE_SIZES.searchResultsPerPage ||
        meta.filtersKey !== expectedFiltersKey
      ) {
        return;
      }
    }

    if (isLoading && items.length === 0) {
      setIsSearchPageResultsLoading(true);
      setHasSearchPageResultsResolved(false);
      hasSearchPageResultsLoadingStartedRef.current = true;
      setSearchPageInfo(page);
      if (typeof totalCount === 'number') {
        setSearchPageTotalCount(totalCount);
      }
      return;
    }

    if (
      !isLoading &&
      items.length === 0 &&
      !hasSearchPageResultsResolved &&
      !hasSearchPageResultsLoadingStartedRef.current &&
      searchPageResults.length === 0 &&
      searchPageQuery.trim()
    ) {
      return;
    }

    setSearchPageResults(items);
    setSearchPageInfo(page);
    if (typeof totalCount === 'number') {
      setSearchPageTotalCount(totalCount);
    } else if (!isLoading) {
      setSearchPageTotalCount(items.length);
    }
    setHasSearchPageResultsResolved(!isLoading);
    setIsSearchPageResultsLoading(isLoading);
  }, [hasSearchPageResultsResolved, searchPageFilters, searchPageQuery, searchPageResults.length, searchResultsPage, searchSortOrder]);

  const isSearchSpreadsheetFileType = React.useCallback((fileType?: string): boolean => {
    return ['xls', 'xlsx', 'xlsm', 'xlsb'].indexOf((fileType || '').toLowerCase()) !== -1;
  }, []);

  const getSearchFileTypeIcon = React.useCallback((fileType?: string): string => {
    const normalizedType = (fileType || '').toLowerCase();
    if (normalizedType.indexOf('pdf') !== -1) return '📄';
    if (normalizedType.indexOf('ppt') !== -1) return '📊';
    if (normalizedType.indexOf('xls') !== -1) return '📈';
    if (normalizedType.indexOf('doc') !== -1 || normalizedType.indexOf('txt') !== -1) return '📝';
    return '📎';
  }, []);

  const getSearchThumbnailCandidates = React.useCallback((item: TSearchResultItem): string[] => {
    const serverRelativeUrl = item.serverRelativeUrl || '';
    if (!webUrl) return [];

    const normalizedServerRelativeUrl = serverRelativeUrl
      ? (serverRelativeUrl.startsWith('/') ? serverRelativeUrl : `/${serverRelativeUrl}`)
      : '';
    const origin = new URL(webUrl).origin;
    const absoluteFileUrl = item.fileUrl && item.fileUrl.startsWith('http')
      ? item.fileUrl
      : normalizedServerRelativeUrl
        ? `${origin}${normalizedServerRelativeUrl}`
        : '';
    const normalizedFileType = (item.fileType || item.fileName?.split('.').pop() || '').toLowerCase();
    const candidates = ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'].indexOf(normalizedFileType) !== -1
      ? (absoluteFileUrl ? [absoluteFileUrl] : [])
      : [];

    if (item.fileUniqueId) {
      candidates.push(
        isSearchSpreadsheetFileType(normalizedFileType)
          ? getExcelThumbnailUrl(webUrl, KM_REVIEW_HUB_DRIVE_ID, item.fileUniqueId)
          : getDriveItemThumbnailUrl(webUrl, KM_REVIEW_HUB_DRIVE_ID, item.fileUniqueId)
      );
    }

    return candidates;
  }, [isSearchSpreadsheetFileType, webUrl]);

  const handleSearchResultDownload = React.useCallback(async (item: TSearchResultItem): Promise<void> => {
    const fileUrl = item.serverRelativeUrl || item.fileUrl || '';
    if (!fileUrl) {
      return;
    }

    try {
      await downloadSharePointFile(
        context,
        webUrl,
        fileUrl,
        item.fileName || item.title || 'download'
      );
    } catch (error) {
      console.error('Search result download failed:', error);
    }
  }, [context, webUrl]);

  React.useEffect(() => {
    setSearchResultsPage(1);
  }, [searchResultsViewMode]);

  const handleSearchResultsPageChange = React.useCallback((nextPage: number): void => {
    const safeNextPage = Math.max(1, nextPage);
    if (safeNextPage === searchResultsPage) {
      return;
    }

    setIsSearchPageResultsLoading(true);
    setHasSearchPageResultsResolved(false);
    setSearchResultsPage(safeNextPage);
    window.setTimeout(() => {
      searchResultsScrollAreaRef.current?.scrollTo({ top: 0, behavior: 'smooth' });
    }, 0);
  }, [searchResultsPage]);

  const handleSearchPageClearAll = React.useCallback((): void => {
    pushPageUrl(NAV_PATHS.search);
    setSearchPageQuery('');
    setSearchPageDraftQuery('');
    setSearchPageSuggestions([]);
    setSearchPageResults([]);
    setSearchPageTotalCount(0);
    setSearchPageInfo(undefined);
    setSearchPageFilters({});
    setActiveSearchFilterKey(null);
    setSearchResultsPage(1);
    setSearchSortOrder('relevance');
    setIsSearchPageResultsLoading(false);
    setHasSearchPageResultsResolved(true);
  }, []);

  const isSearchMode = !routedDocumentId && !selectedBU && activeSidebarView === 'home' && isSearchPageOpen;
  const isReviewerPage = activeSidebarView === 'audit-log' || activeSidebarView === 'analytics' || activeSidebarView === 'kmartifacts' || activeSidebarView === 'km-library';
  const reviewerFrameSrc = activeSidebarView === 'kmartifacts'
    ? `${webUrl}${LIST_PATHS.kmHarvestHub.replace(SITE_RELATIVE_URL, '')}`
    : activeSidebarView === 'km-library'
      ? IFRAME_URLS.kmLibrary
      : `${webUrl}/Lists/Audit%20Log/AllItems.aspx`;
  const reviewerFrameTitle = activeSidebarView === 'kmartifacts'
    ? 'KM Harvest Hub'
    : activeSidebarView === 'km-library'
      ? 'KM Library'
      : 'Audit Log';
  const analyticsPageGroups = IFRAME_URLS.powerBI.pages.reduce((groups: Array<{ sectionName: string; pages: typeof IFRAME_URLS.powerBI.pages }>, page) => {
    const sectionName = page.sectionName || page.name;
    const existingGroup = groups.find((group) => group.sectionName === sectionName);
    if (existingGroup) {
      existingGroup.pages.push(page);
    } else {
      groups.push({ sectionName, pages: [page] });
    }
    return groups;
  }, []);
  const isAnyAnalyticsFrameLoaded = IFRAME_URLS.powerBI.pages.some((page) => analyticsFrameLoadedByPage[page.pageName]);
  const isHomeFullPage = !routedDocumentId && !selectedBU && !isReviewerPage && (activeSidebarView === 'home' || activeSidebarView === 'all-documents' || activeSidebarView === 'published' || activeSidebarView === 'my-documents' || activeSidebarView === 'my-bookmarks') && !isSearchPageOpen && activePage === 'about';
  const showRoutedDetailFullScreen = Boolean(routedDocumentId && context);
  const submittedSearchQuery = searchPageQuery.trim();
  const searchResultsPageSize = PAGE_SIZES.searchResultsPerPage;
  const backendAvailableCount = typeof searchPageInfo?.totalAvailable === 'number'
    ? searchPageInfo.totalAvailable
    : searchPageTotalCount;
  const searchResultsTotalCount = Math.max(backendAvailableCount || 0, searchPageResults.length);
  const searchResultsPageCount = Math.max(1, Math.ceil(searchResultsTotalCount / searchResultsPageSize));
  const safeSearchResultsPage = Math.min(searchResultsPage, searchResultsPageCount);
  const searchResultsPaginationTokens = buildSearchPaginationTokens(safeSearchResultsPage, searchResultsPageCount);
  const searchResultStartIndex = (safeSearchResultsPage - 1) * searchResultsPageSize;
  const searchResultItems = searchPageResults;
  const backendPageSkip = typeof searchPageInfo?.skip === 'number' ? searchPageInfo.skip : searchResultStartIndex;
  const backendPageReturned = typeof searchPageInfo?.returned === 'number' ? searchPageInfo.returned : searchResultItems.length;
  const searchResultStartLabel = searchResultsTotalCount > 0 ? backendPageSkip + 1 : 0;
  const searchResultEndLabel = Math.min(backendPageSkip + backendPageReturned, searchResultsTotalCount);
  const isSearchResultWindowCapped = Boolean(searchPageInfo?.capped);
  const searchResultsCountLabel = isSearchPageResultsLoading
    ? 'Loading matching assets...'
    : isSearchResultWindowCapped
      ? `Showing ${searchResultStartLabel}\u2013${searchResultEndLabel} of ${searchResultsTotalCount} best matching assets`
      : `Showing ${searchResultStartLabel}\u2013${searchResultEndLabel} of ${searchResultsTotalCount} matching assets`;

  React.useEffect(() => {
    if (appReadyAnnouncedRef.current || !isReviewerAccessResolved) {
      return;
    }

    appReadyAnnouncedRef.current = true;
    let innerFrameId = 0;
    const outerFrameId = window.requestAnimationFrame(() => {
      innerFrameId = window.requestAnimationFrame(() => {
        const appReady = (window as any).__appReady;
        if (typeof appReady === 'function') {
          appReady();
        }
      });
    });

    return () => {
      window.cancelAnimationFrame(outerFrameId);
      if (innerFrameId) {
        window.cancelAnimationFrame(innerFrameId);
      }
    };
  }, [isReviewerAccessResolved]);

  React.useEffect(() => {
    if (activeSidebarView === 'analytics') {
      setAnalyticsFrameLoadedByPage({});
      return;
    }

    if (activeSidebarView === 'audit-log' || activeSidebarView === 'kmartifacts' || activeSidebarView === 'km-library') {
      setIsReviewerFrameLoaded(false);
    }
  }, [activeSidebarView, reviewerFrameSrc]);

  React.useEffect(() => {
    if (searchResultsViewMode !== 'grid' || searchResultItems.length === 0) return;

    const preloaders: HTMLImageElement[] = [];
    searchResultItems.forEach((item) => {
      const thumbnailCandidates = getSearchThumbnailCandidates(item);
      const activeThumbnailIndex = searchThumbnailAttemptByDocument[item.id] || 0;
      const activeThumbnailUrl = thumbnailCandidates[activeThumbnailIndex] || '';
      if (!activeThumbnailUrl || searchThumbnailLoadedByDocument[item.id]) return;

      const preloader = new Image();
      preloader.decoding = 'async';
      preloader.loading = 'eager';
      preloader.onload = () => setSearchThumbnailLoadedByDocument((previous) => ({ ...previous, [item.id]: true }));
      preloader.onerror = () => setSearchThumbnailAttemptByDocument((previous) => ({ ...previous, [item.id]: activeThumbnailIndex + 1 }));
      preloader.src = activeThumbnailUrl;
      preloaders.push(preloader);
    });

    return () => {
      preloaders.forEach((preloader) => {
        preloader.onload = null;
        preloader.onerror = null;
      });
    };
  }, [
    getSearchThumbnailCandidates,
    searchResultItems,
    searchResultsViewMode,
    searchThumbnailAttemptByDocument,
    searchThumbnailLoadedByDocument
  ]);

  const searchFilterLabels: TSearchFilterChip[] = [
    { label: 'All', key: null },
    { label: 'Document Type', key: 'documentType' },
    { label: 'Business Unit', key: 'businessUnit' },
    { label: 'Department', key: 'department' },
    { label: 'Client', key: 'client' },
    { label: 'Geography', key: 'region' },
    { label: 'Therapy Area', key: 'therapyArea' },
    { label: 'Disease Area', key: 'diseaseArea' }
  ];
  const handleSearchPageFilterGroupsChange = React.useCallback((groups: TSearchFilterGroup[]): void => {
    const nextGroups = Array.isArray(groups) ? groups : [];

    if (isSearchPageFacetDebugEnabled()) {
      const departmentGroup = nextGroups.find((group) => group.key === 'department');
      console.info('[Search] Search page received filter groups from provider', {
        groupKeys: nextGroups.map((group) => group.key),
        departmentOptionCount: departmentGroup?.children.length || 0,
        topDepartmentOptions: (departmentGroup?.children || []).slice(0, 20).map((option) => ({
          title: option.title,
          value: option.value || option.title,
          count: option.count,
          level: option.level || 0
        }))
      });
    }

    setSearchPageFilterGroups(nextGroups);
  }, []);
  const displaySearchFilterGroups = React.useMemo<TSearchFilterGroup[]>(() => {
    const shouldShowResultCounts = Boolean(searchPageQuery.trim());
    const getDepartmentParentValue = (option: TSearchFilterOption): string => {
      if (!option.value || Number(option.level || 0) <= 0) {
        return '';
      }

      const valueParts = String(option.value).split(':').map((part) => part.trim()).filter(Boolean);
      if (valueParts.length > 1) {
        return valueParts.slice(0, -1).join(':');
      }

      return String(option.parentTitle || '').trim();
    };
    const orderDepartmentOptions = (group: TSearchFilterGroup): TSearchFilterGroup => {
      if (group.key !== 'department') {
        return group;
      }

      const compareByTitleThenCount = (left: TSearchFilterOption, right: TSearchFilterOption): number =>
        left.title.localeCompare(right.title) || right.count - left.count;
      const selectedValue = searchPageFilters.department || '';
      const normalizedSelectedValue = normalizeSearchFilterCountKey(selectedValue);
      const topLevelOptions = group.children.filter((option) => Number(option.level || 0) === 0);
      const childOptions = group.children.filter((option) => Number(option.level || 0) > 0);
      const childrenByParent = new Map<string, TSearchFilterOption[]>();

      childOptions.forEach((option) => {
        const parentValue = getDepartmentParentValue(option);
        if (!parentValue) {
          return;
        }

        const normalizedParent = normalizeSearchFilterCountKey(parentValue);
        const siblings = childrenByParent.get(normalizedParent) || [];
        siblings.push(option);
        childrenByParent.set(normalizedParent, siblings);
      });

      const ordered: TSearchFilterOption[] = [];
      topLevelOptions.sort(compareByTitleThenCount).forEach((option) => {
        ordered.push(option);
        const optionValue = option.value || option.title;
        const isSelectedParent = Boolean(normalizedSelectedValue) &&
          normalizeSearchFilterCountKey(optionValue) === normalizedSelectedValue;
        const children = childrenByParent.get(normalizeSearchFilterCountKey(optionValue)) || [];
        if (isSelectedParent || children.length > 0) {
          children.sort(compareByTitleThenCount).forEach((child) => {
            ordered.push(child);
          });
        }
      });

      return {
        ...group,
        children: ordered
      };
    };
    const keepNonZeroOrSelectedOptions = (group: TSearchFilterGroup): TSearchFilterGroup => {
      if (!shouldShowResultCounts) {
        return orderDepartmentOptions(group);
      }

      const selectedValue = searchPageFilters[group.key];
      const normalizedSelectedValue = normalizeSearchFilterCountKey(selectedValue || '');

      return orderDepartmentOptions({
        ...group,
        children: group.children.filter((option) => {
          if (option.count > 0) {
            return true;
          }

          const optionValue = option.value || option.title;
          return Boolean(normalizedSelectedValue) &&
            normalizeSearchFilterCountKey(optionValue) === normalizedSelectedValue;
        })
      });
    };

    if (searchPageFilterGroups.length > 0) {
      const groupsByKey = new Map<TSearchFilterKey, TSearchFilterGroup>(
        searchPageFilterGroups.map((group) => [group.key, group])
      );
      const orderedGroups: TSearchFilterGroup[] = [];

      searchFilterLabels.forEach((filter) => {
        if (!filter.key) {
          return;
        }

        const group = groupsByKey.get(filter.key);
        if (group) {
          orderedGroups.push(group);
        }
      });

      searchPageFilterGroups.forEach((group) => {
        if (!orderedGroups.some((orderedGroup) => orderedGroup.key === group.key)) {
          orderedGroups.push(group);
        }
      });

      return orderedGroups.map(keepNonZeroOrSelectedOptions);
    }

    return [];
  }, [
    searchFilterLabels,
    searchPageFilterGroups,
    searchPageQuery,
    searchPageFilters
  ]);
  const activeSearchFilterGroup = activeSearchFilterKey
    ? displaySearchFilterGroups.find((group) => group.key === activeSearchFilterKey)
    : null;

  React.useEffect(() => {
    if (!isSearchPageOpen || !isSearchPageFacetDebugEnabled()) {
      return;
    }

    const departmentGroup = displaySearchFilterGroups.find((group) => group.key === 'department');
    if (!departmentGroup) {
      console.warn('[Search] Visible Department chip group is missing', {
        query: searchPageQuery,
        availableGroups: displaySearchFilterGroups.map((group) => group.key)
      });
      return;
    }

    const topDepartmentOptions = departmentGroup.children.slice(0, 20).map((option) => ({
      title: option.title,
      value: option.value || option.title,
      count: option.count,
      level: option.level || 0
    }));
    console.info('[Search] Visible Department chip options', {
      query: searchPageQuery,
      optionCount: departmentGroup.children.length,
      topDepartmentOptions
    });
  }, [displaySearchFilterGroups, isSearchPageOpen, searchPageQuery]);

  const handleSearchFilterChipClick = React.useCallback((filterKey: TSearchFilterKey | null): void => {
    if (filterKey === null) {
      setActiveSearchFilterKey(null);
      setSearchPageFilters({});
      setSearchResultsPage(1);
      return;
    }

    setActiveSearchFilterKey((currentFilterKey) => currentFilterKey === filterKey ? null : filterKey);
  }, []);
  const getSearchFilterOptionValue = React.useCallback((option: TSearchFilterOption): string =>
    option.value || option.title,
  []);
  const getSearchFilterSelectedLabel = React.useCallback((filterKey: TSearchFilterKey, selectedValue: string): string => {
    const group = displaySearchFilterGroups.find((item) => item.key === filterKey);
    const normalizedSelectedValue = normalizeSearchFilterCountKey(selectedValue);
    const match = group?.children.find((option) =>
      normalizeSearchFilterCountKey(getSearchFilterOptionValue(option)) === normalizedSelectedValue ||
      normalizeSearchFilterCountKey(option.title) === normalizedSelectedValue
    );

    return match?.title || selectedValue.split(/\s*[:>]\s*/g).filter(Boolean).pop() || selectedValue;
  }, [displaySearchFilterGroups, getSearchFilterOptionValue]);
  const handleSearchFilterOptionClick = React.useCallback((filterKey: TSearchFilterKey, optionValue: string): void => {
    setSearchPageFilters((current) => {
      const nextValue = current[filterKey] === optionValue ? null : optionValue;
      const nextFilters = {
        ...current,
        [filterKey]: nextValue
      };

      if (filterKey === 'businessUnit' && current.businessUnit !== nextValue) {
        nextFilters.department = null;
      }

      return nextFilters;
    });
    setSearchResultsPage(1);
    setActiveSearchFilterKey(null);
  }, []);
  const hasSearchPageFilters = Object.values(searchPageFilters).some(Boolean);
  const renderSearchFilterChipContent = (filter: TSearchFilterChip): JSX.Element => {
    return (
      <>
        <span className={styles.searchPageChipLabel}>{filter.label}</span>
        {filter.key && (
          <span className={styles.searchPageChipChevron} aria-hidden="true">▾</span>
        )}
      </>
    );
  };
  const renderSearchFilterDropdown = (filterKey: TSearchFilterKey | null): JSX.Element | null => {
    if (
      !filterKey ||
      activeSearchFilterKey !== filterKey ||
      !activeSearchFilterGroup ||
      activeSearchFilterGroup.key !== filterKey
    ) {
      return null;
    }

    return (
      <div className={styles.searchPageFilterDropdown}>
        <div className={styles.searchPageFilterDropdownHeader}>
          <span>{activeSearchFilterGroup.title}</span>
          <div className={styles.searchPageFilterHeaderActions}>
            {searchPageFilters[activeSearchFilterKey] && (
              <button
                type="button"
                className={styles.searchPageFilterClear}
                onClick={() => handleSearchFilterOptionClick(activeSearchFilterKey, searchPageFilters[activeSearchFilterKey]!)}
              >
                Clear
              </button>
            )}
            <button
              type="button"
              className={styles.searchPageFilterClose}
              onClick={() => setActiveSearchFilterKey(null)}
              aria-label="Close filter options"
            >
              ×
            </button>
          </div>
        </div>
        {activeSearchFilterGroup.children.length === 0 ? (
          <div className={styles.searchPageFilterEmpty}>No filters available for this search</div>
        ) : activeSearchFilterGroup.children.map((option) => {
          const optionValue = getSearchFilterOptionValue(option);
          const isSelected = searchPageFilters[activeSearchFilterKey] === optionValue;
          const isHierarchyChild = activeSearchFilterGroup.key === 'department' && Number(option.level || 0) > 0;
          return (
            <button
              key={`${activeSearchFilterKey}-${optionValue}`}
              type="button"
              className={`${styles.searchPageFilterOption} ${isSelected ? styles.searchPageFilterOptionActive : ''}`}
              style={isHierarchyChild ? { paddingLeft: 26 } : undefined}
              onClick={() => handleSearchFilterOptionClick(activeSearchFilterKey, optionValue)}
            >
              <span className={styles.searchPageFilterOptionTitle}>
                {isHierarchyChild ? `› ${option.title}` : option.title}
              </span>
              <span className={styles.searchPageFilterOptionCount}>{option.count}</span>
            </button>
          );
        })}
      </div>
    );
  };
  const renderSearchFilterChips = (variant: TSearchChipRowVariant): JSX.Element => (
    <>
      <div
        className={`${styles.searchFilterChipRow} ${variant === 'results' ? styles.searchFilterChipRowResults : styles.searchFilterChipRowLanding} ${activeSearchFilterKey && activeSearchFilterGroup ? styles.searchFilterChipRowOpen : ''}`}
        aria-label="Search filters"
      >
        {searchFilterLabels.map((filter) => {
          const isChipActive = filter.key === null
            ? !hasSearchPageFilters && activeSearchFilterKey === null
            : activeSearchFilterKey === filter.key || Boolean(searchPageFilters[filter.key]);

          return (
            <div key={filter.key || 'all'} className={styles.searchPageChipWrap}>
              <button
                type="button"
                className={`${styles.searchPageChip} ${isChipActive ? styles.searchPageChipActive : ''}`}
                onClick={() => handleSearchFilterChipClick(filter.key)}
              >
                {renderSearchFilterChipContent(filter)}
              </button>
              {renderSearchFilterDropdown(filter.key)}
            </div>
          );
        })}
      </div>
      {hasSearchPageFilters && (
        <div className={styles.searchSelectedFilters}>
          {searchFilterLabels
            .filter((filter) => filter.key && searchPageFilters[filter.key])
            .map((filter) => (
              <button
                key={filter.key || filter.label}
                type="button"
                className={styles.searchSelectedFilterPill}
                onClick={() => filter.key && handleSearchFilterOptionClick(filter.key, searchPageFilters[filter.key]!)}
              >
                <span>{filter.label}</span>
                <strong>{filter.key ? getSearchFilterSelectedLabel(filter.key, searchPageFilters[filter.key]!) : ''}</strong>
                <span aria-hidden="true">×</span>
              </button>
            ))}
        </div>
      )}
    </>
  );

  return (
    <div
      ref={rootRef}
      className={`${styles.projectRoot} ${styles.appFadeIn} ${styles.isExpandedFull} ${!isReviewerPage && !showRoutedDetailFullScreen ? styles.visualScale80 : ''} ${isSearchMode ? styles.searchFullPageRoot : isHomeFullPage ? styles.homeFullPageRoot : ''}`}
    >
      {showRoutedDetailFullScreen ? (
        <div className={styles.routedDetailStandalone}>
          <DocumentDetailPage
            context={context!}
            documentId={routedDocumentId!}
            onClose={handleCloseRoutedDocument}
            backTo="home"
            backButtonLabel={sessionStorage.getItem(UPLOAD_SUCCESS_RESTORE_STORAGE_KEY) ? 'Back' : undefined}
            userName={currentUserName}
            userEmail={currentUserEmail}
            userPhotoUrl={currentUserPhotoUrl}
            onLogout={handleLogout}
            onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
            onHomeOpen={handleShellHomeOpen}
            onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
            onBookmarksOpen={handleHomeBookmarksOpen}
            onDocumentsOpen={handleLearnerDocumentsOpen}
            onContactOpen={handleHomeContactOpen}
            onAuditLogOpen={handleReviewerAuditLogOpen}
            onAnalyticsOpen={handleReviewerAnalyticsOpen}
            showReviewerNav={canAccessKmReviewHub}
            hideDocumentsNav={isLearner}
          />
        </div>
      ) : (
        <>
      <main className={`${styles.mainContainer} ${isSearchMode || isHomeFullPage ? styles.searchModeMainContainer : ''}`}>
        {!showHomeUploader && !isSearchMode && !isSearchPageOpen && !categoryDocumentsTitle && !selectedBU &&
          activeSidebarView !== 'my-documents' && activeSidebarView !== 'my-bookmarks' &&
          activeSidebarView !== 'all-documents' && activeSidebarView !== 'published' &&
          activeSidebarView !== 'audit-log' && activeSidebarView !== 'analytics' &&
          activeSidebarView !== 'kmartifacts' && activeSidebarView !== 'km-library' && (
            <div className={styles.homeFixedHeader}>
              <IKShellHeader
                userName={currentUserName}
                userEmail={currentUserEmail}
                userPhotoUrl={currentUserPhotoUrl}
                onLogout={handleLogout}
                onHomeOpen={handleShellHomeOpen}
                onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                onBookmarksOpen={handleHomeBookmarksOpen}
                onDocumentsOpen={handleLearnerDocumentsOpen}
                onContactOpen={handleHomeContactOpen}
                onAuditLogOpen={handleReviewerAuditLogOpen}
                onAnalyticsOpen={handleReviewerAnalyticsOpen}
                onKmLibraryOpen={handleReviewerKmLibraryOpen}
                showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
              />
            </div>
        )}
        <div className={styles.scrollArea}>
          {activeSidebarView === 'my-documents' ? (
            <section className={styles.embeddedPageShell}>
              <MyDocuments
                context={context}
                onClose={handleShellHomeOpen}
                onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                onHomeOpen={handleShellHomeOpen}
                onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                onBookmarksOpen={handleHomeBookmarksOpen}
                onDocumentsOpen={handleLearnerDocumentsOpen}
                onContactOpen={handleHomeContactOpen}
                onAuditLogOpen={handleReviewerAuditLogOpen}
                onAnalyticsOpen={handleReviewerAnalyticsOpen}
                showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
                onLogout={handleLogout}
                userPhotoUrl={currentUserPhotoUrl}
                onViewDocument={(documentId) => {
                  openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
                }}
                isLearner={isLearner}
              />
            </section>
          ) : activeSidebarView === 'my-bookmarks' ? (
            <section className={styles.embeddedPageShell}>
              <MyBookmarks
                context={context}
                onClose={handleShellHomeOpen}
                onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                onHomeOpen={handleShellHomeOpen}
                onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                onBookmarksOpen={handleHomeBookmarksOpen}
                onDocumentsOpen={handleLearnerDocumentsOpen}
                onContactOpen={handleHomeContactOpen}
                onAuditLogOpen={handleReviewerAuditLogOpen}
                onAnalyticsOpen={handleReviewerAnalyticsOpen}
                showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
                onLogout={handleLogout}
                userPhotoUrl={currentUserPhotoUrl}
                onViewDocument={(documentId) => {
                  openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
                }}
                isLearner={isLearner}
              />
            </section>
          ) : activeSidebarView === 'all-documents' ? (
            <section className={styles.embeddedPageShell}>
              <ViewAllDocumentsPage
                context={context}
                scope="kmReviewHub"
                onClose={handleShellHomeOpen}
                onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                onHomeOpen={handleShellHomeOpen}
                onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                onBookmarksOpen={handleHomeBookmarksOpen}
                onDocumentsOpen={handleLearnerDocumentsOpen}
                onContactOpen={handleHomeContactOpen}
                onAuditLogOpen={handleReviewerAuditLogOpen}
                onAnalyticsOpen={handleReviewerAnalyticsOpen}
                showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
                userPhotoUrl={currentUserPhotoUrl}
                onViewDocument={(documentId) => {
                  openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
                }}
                isLearner={isLearner}
              />
            </section>
          ) : activeSidebarView === 'published' ? (
            <section className={styles.embeddedPageShell}>
              <ViewAllDocumentsPage
                context={context}
                scope="recentlyPublished"
                onClose={handleShellHomeOpen}
                onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                onHomeOpen={handleShellHomeOpen}
                onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                onBookmarksOpen={handleHomeBookmarksOpen}
                onDocumentsOpen={handleLearnerDocumentsOpen}
                onContactOpen={handleHomeContactOpen}
                onAuditLogOpen={handleReviewerAuditLogOpen}
                onAnalyticsOpen={handleReviewerAnalyticsOpen}
                showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
                userPhotoUrl={currentUserPhotoUrl}
                onViewDocument={(documentId) => {
                  openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
                }}
                isLearner={isLearner}
              />
            </section>
          ) : categoryDocumentsTitle ? (
            <>
              <CategoryDocumentsPage
                context={context}
                categoryTitle={categoryDocumentsTitle}
                taxonomyOptions={searchTaxonomyOptions}
                onHomeOpen={handleShellHomeOpen}
                onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                onBookmarksOpen={handleHomeBookmarksOpen}
                onDocumentsOpen={handleLearnerDocumentsOpen}
                onContactOpen={handleHomeContactOpen}
                onAuditLogOpen={handleReviewerAuditLogOpen}
                onAnalyticsOpen={handleReviewerAnalyticsOpen}
                onViewDocument={handleSearchResultViewDocument}
                userName={currentUserName}
                userEmail={currentUserEmail}
                userPhotoUrl={currentUserPhotoUrl}
                onLogout={handleLogout}
                showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
                isLearner={isLearner}
              />
            </>
          ) : selectedBU ? (
            <BusinessUnitDetailPage
              selectedBU={selectedBU}
              onBack={handleBackToMain}
              context={context}
              onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
              onHomeOpen={handleShellHomeOpen}
              onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
              onBookmarksOpen={handleHomeBookmarksOpen}
              onDocumentsOpen={handleLearnerDocumentsOpen}
              onContactOpen={handleHomeContactOpen}
              onAuditLogOpen={handleReviewerAuditLogOpen}
              onAnalyticsOpen={handleReviewerAnalyticsOpen}
              showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
              onLogout={handleLogout}
              userPhotoUrl={currentUserPhotoUrl}
              isLearner={isLearner}
            />
          ) : isSearchPageOpen ? (
            <section className={styles.searchPageShell}>
              <IKShellHeader
                userName={currentUserName}
                userEmail={currentUserEmail}
                userPhotoUrl={currentUserPhotoUrl}
                onLogout={handleLogout}
                onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                onHomeOpen={handleShellHomeOpen}
                onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                onBookmarksOpen={handleHomeBookmarksOpen}
                onDocumentsOpen={handleLearnerDocumentsOpen}
                onContactOpen={handleHomeContactOpen}
                onAuditLogOpen={handleReviewerAuditLogOpen}
                onAnalyticsOpen={handleReviewerAnalyticsOpen}
                onKmHarvestHubOpen={handleReviewerKmHarvestHubOpen}
                onKmLibraryOpen={handleReviewerKmLibraryOpen}
                showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
              />
              <div className={submittedSearchQuery
                ? styles.searchResultsPageBody
                : `${styles.searchPageBody} ${searchPageDraftQuery.trim().length >= 1 && searchPageSuggestions.length > 0 ? styles.searchPageBodyWithSuggestions : ''}`}>
                {submittedSearchQuery ? (
                  <>
                    <div className={`${styles.searchPageSearchArea} ${styles.searchResultsSearchArea}`}>
                      <div className={styles.searchPageSearchContainer}>
                        <span className={styles.searchPageSearchIcon} aria-hidden="true">
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
                            <circle cx="11" cy="11" r="6" stroke="currentColor" strokeWidth="2" />
                            <path d="m16 16 4 4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                          </svg>
                        </span>
                        <input
                          type="text"
                          className={styles.searchPageSearchInput}
                          placeholder="Search by keyword, title, client, therapy area, or document content"
                          value={searchPageDraftQuery}
                          onChange={(event) => handleSearchPageInputChange(event.target.value)}
                          onKeyDown={(event) => {
                            if (event.key === 'Enter') {
                              event.preventDefault();
                              handleHeaderSearchSubmit(searchPageDraftQuery);
                            }
                          }}
                        />
                      </div>
                      {(submittedSearchQuery || hasSearchPageFilters) && (
                        <button
                          type="button"
                          className={styles.searchPageClearAllButton}
                          onClick={handleSearchPageClearAll}
                        >
                          Clear all
                        </button>
                      )}
                    </div>
                    {renderSearchFilterChips('results')}
                    <div className={styles.searchResultsViewToolbar}>
                      <div className={styles.searchResultsMatchCount}>
                        {searchResultsCountLabel}
                      </div>
                      <label className={styles.searchResultsSortControl}>
                        <span>Sort By</span>
                        <select
                          value={searchSortOrder}
                          onChange={(event) => {
                            setSearchSortOrder(event.target.value as TSearchSortOrder);
                            setSearchResultsPage(1);
                            setSearchPageResults([]);
                            setSearchPageInfo(undefined);
                            setIsSearchPageResultsLoading(Boolean(searchPageQuery.trim()));
                            setHasSearchPageResultsResolved(!searchPageQuery.trim());
                          }}
                        >
                          <option value="relevance">Relevance</option>
                          <option value="newest">Newest First</option>
                          <option value="oldest">Oldest First</option>
                        </select>
                      </label>
                      <div className={styles.searchResultsViewModeToggles}>
                        <button
                          type="button"
                          className={`${styles.searchResultsViewModeButton} ${searchResultsViewMode === 'list' ? styles.searchResultsViewModeButtonActive : ''}`}
                          onClick={() => setSearchResultsViewMode('list')}
                          title="List view"
                        >
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" />
                          </svg>
                        </button>
                        <button
                          type="button"
                          className={`${styles.searchResultsViewModeButton} ${searchResultsViewMode === 'grid' ? styles.searchResultsViewModeButtonActive : ''}`}
                          onClick={() => setSearchResultsViewMode('grid')}
                          title="Grid view"
                        >
                          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                            <rect x="3" y="3" width="7" height="7" />
                            <rect x="14" y="3" width="7" height="7" />
                            <rect x="14" y="14" width="7" height="7" />
                            <rect x="3" y="14" width="7" height="7" />
                          </svg>
                        </button>
                      </div>
                    </div>
                    <div className={styles.searchResultsScrollArea} ref={searchResultsScrollAreaRef}>
                      <div className={searchResultsViewMode === 'grid' ? styles.searchResultsGrid : styles.searchResultsList}>
                        {isSearchPageResultsLoading ? (
                          <div className={styles.searchResultsStatus}>Loading matching documents...</div>
                        ) : searchResultItems.length > 0 ? (
                          searchResultItems.map((item) => {
                            const downloadUrl = item.fileUrl || item.serverRelativeUrl || '';
                            const statusLabel = getSearchDisplayStatus(item.status) || 'Active';
                            const fileType = (item.fileType || item.fileName?.split('.').pop() || 'FILE').toUpperCase();
                            if (searchResultsViewMode === 'grid') {
                              const thumbnailCandidates = getSearchThumbnailCandidates(item);
                              const activeThumbnailIndex = searchThumbnailAttemptByDocument[item.id] || 0;
                              const activeThumbnailUrl = thumbnailCandidates[activeThumbnailIndex] || '';
                              const isThumbnailLoaded = !!searchThumbnailLoadedByDocument[item.id];
                              const shouldShowFallback = !activeThumbnailUrl || !isThumbnailLoaded;
                              const usesDocumentPreviewCrop = ['PDF', 'DOC', 'DOCX', 'PPT', 'PPTX', 'XLS', 'XLSX'].indexOf(fileType) !== -1;

                              return (
                                <article key={item.id} className={styles.searchResultGridCard}>
                                  <div className={styles.searchResultGridPreview}>
                                    {activeThumbnailUrl && (
                                      <img
                                        src={activeThumbnailUrl}
                                        alt=""
                                        className={`${styles.searchResultThumbnailImage} ${usesDocumentPreviewCrop ? styles.searchResultThumbnailDocumentImage : ''}`}
                                        loading="eager"
                                        decoding="async"
                                        onLoad={() => setSearchThumbnailLoadedByDocument((previous) => ({ ...previous, [item.id]: true }))}
                                        onError={() => setSearchThumbnailAttemptByDocument((previous) => ({ ...previous, [item.id]: activeThumbnailIndex + 1 }))}
                                      />
                                    )}
                                    {shouldShowFallback && (
                                      <div className={styles.searchResultThumbnailFallback}>
                                        <div className={styles.searchResultFileTypeIcon}>{getSearchFileTypeIcon(fileType)}</div>
                                        <span className={styles.searchResultGridFileType}>{fileType}</span>
                                      </div>
                                    )}
                                  </div>
                                  <div className={styles.searchResultGridBody}>
                                    <div className={styles.searchResultGridMeta}>
                                      <span className={styles.searchResultAuthorChip}>{item.contributor || 'Internal'}</span>
                                      <span className={styles.searchResultFileTypeChip}>{fileType}</span>
                                    </div>
                                    <h2 className={styles.searchResultGridTitle}>{item.title || item.fileName || 'Untitled document'}</h2>
                                    <p className={styles.searchResultGridDescription}>{item.description || ''}</p>
                                    <div className={styles.searchResultGridFooter}>
                                      <span>{item.updated || 'No date available'}</span>
                                      <span className={`${styles.statusBadge} ${getSearchStatusBadgeClassName(statusLabel)}`}>{statusLabel}</span>
                                    </div>
                                    <div className={styles.searchResultGridActions}>
                                      <button
                                        type="button"
                                        className={styles.searchResultViewButton}
                                        onClick={() => handleSearchResultViewDocument(item.id)}
                                      >
                                        View
                                      </button>
                                      {!isLearner && (
                                        <button
                                          type="button"
                                          className={styles.searchResultDownloadButton}
                                          aria-label={`Download ${item.title || item.fileName || 'document'}`}
                                          disabled={!downloadUrl}
                                          onClick={() => {
                                            if (downloadUrl) {
                                              void handleSearchResultDownload(item);
                                            }
                                          }}
                                        >
                                          Download
                                        </button>
                                      )}
                                    </div>
                                  </div>
                                </article>
                              );
                            }

                            return (
                              <article key={item.id} className={styles.searchResultCard}>
                                <div className={styles.searchResultListContent}>
                                  <div className={styles.searchResultContent}>
                                    <h2 className={styles.searchResultTitle}>{item.title || item.fileName || 'Untitled document'}</h2>
                                    <p className={styles.searchResultDescription}>{item.description || ''}</p>
                                    <div className={styles.searchResultStatusLine}>
                                      <span className={styles.searchResultStatusLabel}>Status :</span>
                                      <span className={`${styles.statusBadge} ${getSearchStatusBadgeClassName(statusLabel)}`}>{statusLabel}</span>
                                    </div>
                                  </div>
                                  <div className={styles.searchResultDivider} aria-hidden="true" />
                                  <div className={styles.searchResultMeta}>
                                    <strong>{item.contributor || 'Internal'}</strong>
                                    <span className={styles.searchResultDate}>{item.updated || 'No date available'}</span>
                                  </div>
                                  <div className={styles.searchResultActions}>
                                    <button
                                      type="button"
                                      className={styles.searchResultViewButton}
                                      onClick={() => handleSearchResultViewDocument(item.id)}
                                    >
                                      View
                                    </button>
                                    {!isLearner && (
                                      <button
                                        type="button"
                                        className={styles.searchResultDownloadButton}
                                        aria-label={`Download ${item.title || item.fileName || 'document'}`}
                                        disabled={!downloadUrl}
                                        onClick={() => {
                                          if (downloadUrl) {
                                            void handleSearchResultDownload(item);
                                          }
                                        }}
                                      >
                                        Download
                                      </button>
                                    )}
                                  </div>
                                </div>
                              </article>
                            );
                          })
                        ) : (
                          <div className={styles.searchResultsStatus}>No documents found.</div>
                        )}
                      </div>
                      {searchResultsTotalCount > searchResultsPageSize && (
                        <nav className={styles.searchResultsPagination} aria-label="Search results pagination">
                          <button
                            type="button"
                            className={styles.searchResultsPageButton}
                            disabled={safeSearchResultsPage <= 1}
                            onClick={() => handleSearchResultsPageChange(Math.max(1, safeSearchResultsPage - 1))}
                          >
                            Previous
                          </button>
                          {searchResultsPaginationTokens.map((token, tokenIndex) => {
                            if (token === 'ellipsis') {
                              return <span key={`ellipsis-${tokenIndex}`} className={styles.searchResultsPageEllipsis}>...</span>;
                            }

                            return (
                              <button
                                key={token}
                                type="button"
                                className={`${styles.searchResultsPageNumber} ${token === safeSearchResultsPage ? styles.searchResultsPageNumberActive : ''}`}
                                onClick={() => handleSearchResultsPageChange(token)}
                              >
                                {token}
                              </button>
                            );
                          })}
                          <button
                            type="button"
                            className={styles.searchResultsPageButton}
                            disabled={safeSearchResultsPage >= searchResultsPageCount}
                            onClick={() => handleSearchResultsPageChange(Math.min(searchResultsPageCount, safeSearchResultsPage + 1))}
                          >
                            Next
                          </button>
                        </nav>
                      )}
                    </div>
                  </>
                ) : (
                  <>
                    <h1 className={styles.searchPageTitle}>Discover Knowledge Faster</h1>
                    <p className={styles.searchPageHelperText}>
                      Search by title, keywords, client, therapy area, business unit, or document content. Use filters to narrow your results and quickly find the most relevant assets.
                    </p>
                    {renderSearchFilterChips('search')}

                    <div className={styles.searchPageSearchArea}>
                      <div className={styles.searchPageSearchContainer}>
                        <span className={styles.searchPageSearchIcon} aria-hidden="true">
                          <svg width="14" height="14" viewBox="0 0 24 24" fill="none">
                            <circle cx="11" cy="11" r="6" stroke="currentColor" strokeWidth="2" />
                            <path d="m16 16 4 4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
                          </svg>
                        </span>
                        <input
                          type="text"
                          className={styles.searchPageSearchInput}
                          placeholder="Search by keyword, title, client, therapy area, or document content"
                          value={searchPageDraftQuery}
                          onChange={(event) => handleSearchPageInputChange(event.target.value)}
                          onKeyDown={(event) => {
                            if (event.key === 'Enter') {
                              event.preventDefault();
                              handleHeaderSearchSubmit(searchPageDraftQuery);
                            }
                          }}
                          autoFocus
                        />
                      </div>
                      {searchPageDraftQuery.trim().length >= 1 && searchPageSuggestions.length > 0 && (
                        <div className={styles.searchPageSuggestionDropdown}>
                          {searchPageSuggestions.slice(0, 5).map((suggestion) => (
                            <button
                              key={`${suggestion.source}-${suggestion.value}`}
                              type="button"
                              className={styles.searchPageSuggestionItem}
                              onMouseDown={(event) => {
                                event.preventDefault();
                                handleHeaderSearchSubmit(suggestion.value);
                              }}
                            >
                              <span className={styles.searchPageSuggestionText}>{suggestion.value}</span>
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                    <div className={styles.searchPageExamples} aria-label="Example searches">
                      <span>Try searching for:</span>
                      <button type="button" onClick={() => handleHeaderSearchSubmit('Oncology')}>Oncology</button>
                      <button type="button" onClick={() => handleHeaderSearchSubmit('SOP')}>SOP</button>
                      <button type="button" onClick={() => handleHeaderSearchSubmit('Generative AI')}>Generative AI</button>
                      <button type="button" onClick={() => handleHeaderSearchSubmit('Nephrology')}>Nephrology</button>
                    </div>
                  </>
                )}
              </div>
              <IKShellFooter onBackHome={handleCloseSearchPage} />
              {/* Hide KnowledgeHubSection in search mode as per Figma */}
              {/* <KnowledgeHubSection context={context} showPublishedSection={false} /> */}
              <div className={styles.hiddenSearchProvider} aria-hidden="true">
                <GenericSearchDropdown
                  searchText={searchPageQuery}
                  draftSearchText={searchPageDraftQuery}
                  isFilterPanelOpen={true}
                  activeMetadataFilter={null}
                  externalFilters={searchPageFilters}
                  spHttpClient={context.spHttpClient}
                  siteUrl={context.pageContext.web.absoluteUrl}
                  context={context}
                  onSuggestionSelect={handleHeaderSearchSubmit}
                  onSearchSubmit={handleHeaderSearchSubmit}
                  onViewDocument={handleSearchResultViewDocument}
                  onSuggestionsChange={setSearchPageSuggestions}
                  onResultsChange={handleSearchPageResultsChange}
                  onFilterGroupsChange={handleSearchPageFilterGroupsChange}
                  renderSuggestionsSection={false}
                  fastResultsOnly={true}
                  isLearner={isLearner}
                  refreshKey={searchProviderRefreshKey}
                  resultsPage={safeSearchResultsPage}
                  resultsPageSize={searchResultsPageSize}
                  searchSortOrder={searchSortOrder}
                />
              </div>
            </section>
          ) : (
            <>
              {(activeSidebarView === 'audit-log' || activeSidebarView === 'analytics' || activeSidebarView === 'kmartifacts' || activeSidebarView === 'km-library') ? (
                <section className={styles.reviewerShell}>
                  <IKShellHeader
                    userName={currentUserName}
                    userEmail={currentUserEmail}
                    userPhotoUrl={currentUserPhotoUrl}
                    onLogout={handleLogout}
                    onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                    onHomeOpen={handleShellHomeOpen}
                    onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                    onBookmarksOpen={handleHomeBookmarksOpen}
                    onDocumentsOpen={handleLearnerDocumentsOpen}
                    onContactOpen={handleHomeContactOpen}
                    onAuditLogOpen={handleReviewerAuditLogOpen}
                    onAnalyticsOpen={handleReviewerAnalyticsOpen}
                    onKmHarvestHubOpen={handleReviewerKmHarvestHubOpen}
                    onKmLibraryOpen={handleReviewerKmLibraryOpen}
                    showReviewerNav={canAccessKmReviewHub}
                    hideDocumentsNav={isLearner}
                    activeNavKey={activeSidebarView === 'analytics'
                      ? 'analytics'
                      : activeSidebarView === 'kmartifacts'
                        ? 'kmHarvestHub'
                        : activeSidebarView === 'km-library'
                          ? 'kmLibrary'
                          : 'auditLog'}
                  />
                  <section className={styles.reviewerFrameShell}>
                    <div className={styles.embeddedPageHeader}>
                      <h1 className={styles.embeddedPageTitle}>
                        {activeSidebarView === 'analytics'
                          ? 'Analytics'
                          : activeSidebarView === 'kmartifacts'
                            ? 'KM Harvest Hub'
                            : activeSidebarView === 'km-library'
                              ? 'KM Library'
                              : 'Audit Log'}
                      </h1>
                    </div>
                    {activeSidebarView === 'analytics' ? (
                      <div className={styles.analyticsFrameHost}>
                        {!isAnyAnalyticsFrameLoaded && (
                          <div className={styles.reviewerFrameLoadingMask} role="status" aria-live="polite">
                            <span className={styles.reviewerFrameSpinner} aria-hidden="true" />
                            <span>Loading Analytics...</span>
                          </div>
                        )}
                        <div className={styles.analyticsContainer}>
                          {analyticsPageGroups.map((group) => (
                            <div key={group.sectionName} className={styles.analyticsSection}>
                              <div className={styles.analyticsFrameStack}>
                                {group.pages.map((page) => {
                                  const embedUrl =
                                    `${IFRAME_URLS.powerBI.baseEmbedUrl}` +
                                    `?reportId=${page.reportId || IFRAME_URLS.powerBI.reportId}` +
                                    `&groupId=${page.groupId || IFRAME_URLS.powerBI.groupId}` +
                                    `&pageName=${page.pageName}` +
                                    `&autoAuth=${IFRAME_URLS.powerBI.autoAuth}` +
                                    `&ctid=${IFRAME_URLS.powerBI.ctid}` +
                                    `&filterPaneEnabled=false` +
                                    `&navContentPaneEnabled=false` +
                                    `&pageView=fitToWidth`;
                                  return (
                                    <div key={page.pageName} className={styles.analyticsFramePanel}>
                                      <iframe
                                        className={`${styles.analyticsFrame} ${analyticsFrameLoadedByPage[page.pageName] ? styles.iframeLoaded : styles.iframeLoading}`}
                                        src={embedUrl}
                                        title={page.name}
                                        allowFullScreen
                                        loading="lazy"
                                        onLoad={() => {
                                          setAnalyticsFrameLoadedByPage((previous) => ({
                                            ...previous,
                                            [page.pageName]: true
                                          }));
                                        }}
                                      />
                                    </div>
                                  );
                                })}
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    ) : (
                      <div className={styles.reviewerFrameHost}>
                        {!isReviewerFrameLoaded && (
                          <div className={styles.reviewerFrameLoadingMask} role="status" aria-live="polite">
                            <span className={styles.reviewerFrameSpinner} aria-hidden="true" />
                            <span>Loading {reviewerFrameTitle}...</span>
                          </div>
                        )}
                        <iframe
                          className={`${styles.reviewerFrame} ${isReviewerFrameLoaded ? styles.iframeLoaded : styles.iframeLoading}`}
                          src={reviewerFrameSrc}
                          title={reviewerFrameTitle}
                          allowFullScreen
                          onLoad={() => setIsReviewerFrameLoaded(true)}
                        />
                      </div>
                    )}
                  </section>
                  <IKShellFooter onBackHome={handleShellHomeOpen} />
                </section>
              ) : (
                <>
                  {!showHomeUploader && (
                    <>
                      <section className={styles.homePageContent}>
                        <MainContentRouter
                          activePage={activePage}
                          context={context}
                          onSearchOpen={handleOpenSearchPage}
                          onUploadOpen={handleOpenUploader}
                          onCategorySelect={handleHomeCategorySelect}
                          hideUploadButton={isLearner}
                          hideCategorySection={true}
                          isAdmin={isKmAdmin}
                          configService={configServiceRef.current}
                        />
                        <RecentlyPublishedSection
                          context={context}
                          isLearner={isLearner}
                          onViewAllOpen={handleRecentlyPublishedOpen}
                          onViewDocument={(documentId) => {
                            openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
                          }}
                        />
                        <AboutCategorySection onCategorySelect={handleHomeCategorySelect} />
                      </section>
                      <IKShellFooter isHomeFooter />
                      <KnowledgeHubSection context={context} showPublishedSection={false} isLearner={isLearner} />
                    </>
                  )}
                </>
              )}
              {showHomeUploader && (
                projectId ? (
                  <FileUpload
                    onClose={handleShellHomeOpen}
                    context={context}
                    projectId={projectId}
                    variant="kmArtifact"
                    sidebarOffset={0}
                    userName={currentUserName}
                    userEmail={currentUserEmail}
                    userPhotoUrl={currentUserPhotoUrl}
                    onLogout={handleLogout}
                    onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                    onHomeOpen={handleShellHomeOpen}
                    onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                    onBookmarksOpen={handleHomeBookmarksOpen}
                    onDocumentsOpen={handleLearnerDocumentsOpen}
                    onContactOpen={handleHomeContactOpen}
                    onAuditLogOpen={handleReviewerAuditLogOpen}
                    onAnalyticsOpen={handleReviewerAnalyticsOpen}
                    showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
                  />
                ) : (
                  <FileUpload
                    onClose={handleShellHomeOpen}
                    context={context}
                    sidebarOffset={0}
                    userName={currentUserName}
                    userEmail={currentUserEmail}
                    userPhotoUrl={currentUserPhotoUrl}
                    onLogout={handleLogout}
                    onAllDocumentsOpen={handleReviewerKmReviewHubOpen}
                    onHomeOpen={handleShellHomeOpen}
                    onBusinessUnitsOpen={handleHomeBusinessUnitsOpen}
                    onBookmarksOpen={handleHomeBookmarksOpen}
                    onDocumentsOpen={handleLearnerDocumentsOpen}
                    onContactOpen={handleHomeContactOpen}
                    onAuditLogOpen={handleReviewerAuditLogOpen}
                    onAnalyticsOpen={handleReviewerAnalyticsOpen}
                    showReviewerNav={canAccessKmReviewHub}
                hideDocumentsNav={isLearner}
                  />
                )
              )}
            </>
          )}
        </div>
      </main>
        </>
      )}

      <RenameBusinessUnitDialog
        context={context}
        isOpen={isRenameBusinessUnitOpen}
        onClose={() => setIsRenameBusinessUnitOpen(false)}
      />
    </div>
  );
};

export default Migration;

