import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { fetchKMDataHubReadFieldMap } from '../utils/documentMetadata';
import { updateWithoutVersion } from './kmVersionControl';
import {
  LIBRARY_NAMES,
  LIST_PATHS as APP_LIST_PATHS,
  PAGE_URLS as APP_PAGE_URLS,
  SITE_URL
} from '../config/appConfig';

export const BASE_PAGE_URL = `${SITE_URL}/SitePages/Home.aspx`;
export const SITE_BASE = SITE_URL;

export const PAGE_URLS = {
  home: `${SITE_BASE}/SitePages/Home.aspx`,
  businessUnits: `${SITE_BASE}/SitePages/Business-Units.aspx`,
  categories: `${SITE_BASE}/SitePages/Categories.aspx`,
  search: `${SITE_BASE}/SitePages/Search.aspx`,
  asset: `${SITE_BASE}/SitePages/Assets.aspx`,
  bookmark: `${SITE_BASE}/SitePages/Bookmark.aspx`,
  myDocuments: `${SITE_BASE}/SitePages/MyDocuments.aspx`,
  analytics: `${SITE_BASE}/SitePages/Analytics.aspx`,
  auditLog: `${SITE_BASE}/SitePages/Audit-Log.aspx`,
  kmReviewHub: `${SITE_BASE}/SitePages/KM-Review-Hub.aspx`,
  kmHarvestHub: `${SITE_BASE}/SitePages/KM-Harvest-Hub.aspx`,
  kmLibrary: `${SITE_BASE}/SitePages/KM-Library.aspx`,
};

export const PAGE_PATHS = {
  home: '/sites/iKnowledgeNext/SitePages/Home.aspx',
  businessUnits: '/SitePages/Business-Units.aspx',
  categories: '/SitePages/Categories.aspx',
  search: '/SitePages/Search.aspx',
  asset: '/SitePages/Assets.aspx',
  bookmark: '/SitePages/Bookmark.aspx',
  myDocuments: '/SitePages/MyDocuments.aspx',
  analytics: '/SitePages/Analytics.aspx',
  auditLog: '/SitePages/Audit-Log.aspx',
  kmReviewHub: '/SitePages/KM-Review-Hub.aspx',
  kmHarvestHub: '/SitePages/KM-Harvest-Hub.aspx',
  kmLibrary: '/SitePages/KM-Library.aspx',
};

export const NAV_PATHS = {
  home: APP_PAGE_URLS.home,
  businessUnits: APP_PAGE_URLS.businessUnits,
  categories: APP_PAGE_URLS.categories,
  search: APP_PAGE_URLS.search,
  asset: APP_PAGE_URLS.assets,
  bookmark: APP_PAGE_URLS.bookmark,
  myDocuments: APP_PAGE_URLS.myDocuments,
  analytics: APP_PAGE_URLS.analytics,
  auditLog: APP_PAGE_URLS.auditLog,
  kmReviewHub: APP_PAGE_URLS.kmReviewHub,
  kmHarvestHub: APP_PAGE_URLS.kmHarvestHub,
  kmLibrary: APP_PAGE_URLS.kmLibrary,
};

export const LIST_PATHS = {
  kmHarvestHub: APP_LIST_PATHS.kmHarvestHub,
};

const LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;
const MAX_SLUG_LENGTH = 120;
const BASE_PAGE = new URL(BASE_PAGE_URL);
const TOPIC_HOME_PATH = BASE_PAGE.pathname.toLowerCase();
const WEB_ABSOLUTE_URL = `${BASE_PAGE.origin}${BASE_PAGE.pathname.substring(0, BASE_PAGE.pathname.toLowerCase().lastIndexOf('/sitepages/'))}`;
export const LOCAL_APP_NAVIGATION_MODE_KEY = 'IKNOWLEDGE_APP_NAVIGATION_MODE';

let serviceContext: { spHttpClient: SPHttpClient; webUrl: string } | null = null;

export const configurePermalinkService = (spHttpClient: SPHttpClient, webUrl: string): void => {
  serviceContext = { spHttpClient, webUrl };
};

export const slugify = (title: string): string => {
  const normalizedSlug = (title || '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, '-')
    .replace(/[^a-z0-9-]/g, '')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '');

  return normalizedSlug.substring(0, MAX_SLUG_LENGTH).replace(/^-+|-+$/g, '');
};

export const buildPrettyUrl = (id: number, _title: string): string => {
  return buildAssetUrl(id);
};

export const detectCurrentPage = (): 'home' | 'businessUnits' | 'categories' | 'search' | 'asset' | 'bookmark' | 'myDocuments' | 'analytics' | 'auditLog' | 'kmReviewHub' | 'kmHarvestHub' | 'kmLibrary' | 'legacy' => {
  const appPage = new URLSearchParams(window.location.search).get('ikpage') || '';
  if (appPage === 'home') return 'home';
  if (appPage === 'businessUnits') return 'businessUnits';
  if (appPage === 'categories') return 'categories';
  if (appPage === 'search') return 'search';
  if (appPage === 'asset') return 'asset';
  if (appPage === 'bookmark') return 'bookmark';
  if (appPage === 'myDocuments') return 'myDocuments';
  if (appPage === 'analytics') return 'analytics';
  if (appPage === 'auditLog') return 'auditLog';
  if (appPage === 'kmReviewHub') return 'kmReviewHub';
  if (appPage === 'kmHarvestHub') return 'kmHarvestHub';
  if (appPage === 'kmLibrary') return 'kmLibrary';

  const path = window.location.pathname.toLowerCase();
  if (path.includes('/home.aspx')) return 'home';
  if (path.includes('/business-units.aspx')) return 'businessUnits';
  if (path.includes('/categories.aspx')) return 'categories';
  if (path.includes('/search.aspx')) return 'search';
  if (path.includes('/assets.aspx')) return 'asset';
  if (path.includes('/bookmark.aspx')) return 'bookmark';
  if (path.includes('/mydocuments.aspx')) return 'myDocuments';
  if (path.includes('/analytics.aspx')) return 'analytics';
  if (path.includes('/audit-log.aspx') || path.includes('/lists/audit%20log/allitems.aspx')) return 'auditLog';
  if (path.includes('/km-review-hub.aspx')) return 'kmReviewHub';
  if (path.includes('/km-harvest-hub.aspx') || path.includes('/lists/km%20harvest%20hub/allitems.aspx')) return 'kmHarvestHub';
  if (path.includes('/km-library.aspx')) return 'kmLibrary';
  return 'legacy';
};

export const parseUrlParams = (): {
  docId: number | null;
  query: string;
  taxonomy: string;
  projectId: string | null;
  category: string;
} => {
  const params = new URLSearchParams(window.location.search);
  const idText =
    params.get('assetID') ||
    params.get('docId') ||
    params.get('id');
  const docId = idText ? parseInt(idText, 10) : null;
  return {
    docId: Number.isFinite(docId) ? docId : null,
    query: params.get('query') || '',
    taxonomy: params.get('taxonomy') || '',
    projectId: params.get('projectId'),
    category: params.get('category') || ''
  };
};

export const pushPageUrl = (
  pagePath: string,
  params?: Record<string, string>
): void => {
  const currentUrl = new URL(window.location.href);
  const projectId = currentUrl.searchParams.get('projectId');

  const url = new URL(window.location.href);
  url.hash = '';

  if (!shouldKeepAppNavigationInCurrentTab()) {
    url.pathname = pagePath;
    url.search = '';
  } else {
    // Workbench/local testing can force app-page navigation to stay on the
    // Workbench URL. Packaged SharePoint keeps the normal/new-tab behavior.
    ['ikpage', 'query', 'taxonomy', 'assetID', 'docId', 'id', 'category'].forEach((key) => {
      url.searchParams.delete(key);
    });
    const appPage = getAppPageKeyFromPath(pagePath);
    if (appPage) {
      url.searchParams.set('ikpage', appPage);
    }
  }

  // Preserve projectId if it exists
  if (projectId) {
    url.searchParams.set('projectId', projectId);
  }

  if (params) {
    Object.entries(params).forEach(([key, value]) => {
      if (value) url.searchParams.set(key, value);
    });
  }

  window.history.pushState({ sameTabNav: true }, document.title, url.toString());
  window.dispatchEvent(new PopStateEvent('popstate'));
};

export const buildAppPageUrl = (
  pagePath: string,
  params?: Record<string, string>
): string => {
  const currentUrl = new URL(window.location.href);
  const projectId = currentUrl.searchParams.get('projectId');
  const url = pagePath.toLowerCase().startsWith('http')
    ? new URL(pagePath)
    : new URL(pagePath, currentUrl.origin);

  url.search = '';
  url.hash = '';

  if (projectId) {
    url.searchParams.set('projectId', projectId);
  }

  if (params) {
    Object.entries(params).forEach(([key, value]) => {
      if (value) {
        url.searchParams.set(key, value);
      }
    });
  }

  return url.toString();
};

const getAppPageKeyFromPath = (pagePath: string): ReturnType<typeof detectCurrentPage> | '' => {
  const normalized = String(pagePath || '').toLowerCase();
  if (normalized.includes('home.aspx')) return 'home';
  if (normalized.includes('business-units.aspx')) return 'businessUnits';
  if (normalized.includes('categories.aspx')) return 'categories';
  if (normalized.includes('search.aspx')) return 'search';
  if (normalized.includes('assets.aspx')) return 'asset';
  if (normalized.includes('bookmark.aspx')) return 'bookmark';
  if (normalized.includes('mydocuments.aspx')) return 'myDocuments';
  if (normalized.includes('analytics.aspx')) return 'analytics';
  if (normalized.includes('audit-log.aspx')) return 'auditLog';
  if (normalized.includes('km-review-hub.aspx')) return 'kmReviewHub';
  if (normalized.includes('km-harvest-hub.aspx')) return 'kmHarvestHub';
  if (normalized.includes('km-library.aspx')) return 'kmLibrary';
  return '';
};

export const shouldKeepAppNavigationInCurrentTab = (): boolean => {
  const href = window.location.href.toLowerCase();
  const params = new URLSearchParams(window.location.search);
  const queryMode = (params.get('iknav') || params.get('ikNavigationMode') || '').toLowerCase();
  const storedMode = (() => {
    try {
      return (window.localStorage.getItem(LOCAL_APP_NAVIGATION_MODE_KEY) || '').toLowerCase();
    } catch {
      return '';
    }
  })();

  if (queryMode === 'newtab' || storedMode === 'newtab') {
    return false;
  }

  if (queryMode === 'inline' || storedMode === 'inline') {
    return true;
  }

  // Default local safety: Workbench/local testing must not leave the Workbench
  // page, while the packaged SharePoint app keeps its configured new-tab flow.
  return href.indexOf('workbench.aspx') !== -1 || href.indexOf('localhost') !== -1;
};

export const openAppPageInNewTab = (
  pagePath: string,
  params?: Record<string, string>
): void => {
  if (shouldKeepAppNavigationInCurrentTab()) {
    pushPageUrl(pagePath, params);
    return;
  }

  const newTabUrl = new URL(buildAppPageUrl(pagePath, params));
  newTabUrl.searchParams.set('env', 'WebViewList');
  window.open(newTabUrl.toString(), '_blank', 'noopener,noreferrer');
};

export const openUrlInNewTab = (url: string): void => {
  window.open(url, '_blank', 'noopener,noreferrer');
};

export const buildHomeUrl = (): string => SITE_BASE + PAGE_PATHS.home;

export const buildSearchUrl = (query?: string, taxonomy?: string): string => {
  const url = new URL(SITE_BASE + PAGE_PATHS.search);
  if (query) url.searchParams.set('query', query);
  if (taxonomy) url.searchParams.set('taxonomy', taxonomy);
  return url.toString();
};

export const buildAssetUrl = (docId?: number): string => {
  const url = new URL(SITE_BASE + PAGE_PATHS.asset);
  if (docId) url.searchParams.set('assetID', docId.toString());
  return url.toString();
};

export const buildBookmarkUrl = (docId?: number): string => {
  return buildAssetUrl(docId);
};

export const buildMyDocumentUrl = (docId?: number): string => {
  return buildAssetUrl(docId);
};

export const navigateToPage = (url: string): void => {
  window.location.href = url;
};

export const updatePageParams = (params: Record<string, string>): void => {
  const url = new URL(window.location.href);
  Object.entries(params).forEach(([key, value]) => {
    if (value) {
      url.searchParams.set(key, value);
    } else {
      url.searchParams.delete(key);
    }
  });
  window.history.pushState({ sameTabNav: true }, document.title, url.toString());
};

const buildStoredDocumentUrl = (id: number, _title: string): string => {
  return buildAssetUrl(id);
};

const getServiceContext = (): { spHttpClient: SPHttpClient; webUrl: string } => {
  if (serviceContext) {
    return serviceContext;
  }

  throw new Error('Permalink service is not configured.');
};

export const updateItemUrlField = async (itemId: number, title: string): Promise<void> => {
  const { spHttpClient, webUrl } = getServiceContext();
  const fieldMap = await fetchKMDataHubReadFieldMap(spHttpClient, webUrl, LIBRARY_NAME);

  // 1. Fetch current status AND URL to check if update is actually needed
  const itemResponse = await spHttpClient.get(
    `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})?$select=${fieldMap.status},${fieldMap.url}`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata.metadata=minimal'
      }
    }
  );

  if (!itemResponse.ok) {
    const errorText = await itemResponse.text();
    throw new Error(`Failed to read KM Data Hub item status: ${itemResponse.status} ${errorText}`);
  }

  const itemJson = await itemResponse.json();
  const itemStatus = (itemJson?.[fieldMap.status] || itemJson?.d?.[fieldMap.status] || '').toString().trim().toLowerCase();
  const currentUrlObj = itemJson?.[fieldMap.url] || itemJson?.d?.[fieldMap.url];
  const currentUrl = currentUrlObj?.Url || '';

  const shouldShowUrl = itemStatus === 'active';
  const resolvedTitle = (title || `Document ${itemId}`).trim() || `Document ${itemId}`;
  const storedUrl = buildStoredDocumentUrl(itemId, resolvedTitle);

  // 🚩 OPTIMIZATION: If the URL is already what we want, STOP here to avoid 409 Conflicts
  if (shouldShowUrl && currentUrl.toLowerCase() === storedUrl.toLowerCase()) {
    return;
  }
  if (!shouldShowUrl && !currentUrl) {
    return; // Already null/empty
  }

  let response: SPHttpClientResponse | undefined;
  await updateWithoutVersion(
    spHttpClient,
    webUrl,
    itemId,
    async () => {
      response = await spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})/ValidateUpdateListItem`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: JSON.stringify({
            formValues: [
              {
                FieldName: fieldMap.url,
                FieldValue: shouldShowUrl ? `${storedUrl}, ${resolvedTitle}` : ''
              }
            ],
            bNewDocumentUpdate: true
          })
        }
      );
    }
  );

  if (!response || !response.ok) {
    const errorText = response ? await response.text() : '';
    throw new Error(`Failed to update KM Data Hub URL field: ${response?.status || 'unknown'} ${errorText}`);
  }

  const responseJson = await response.json();
  const fieldResults = responseJson?.value || responseJson?.d?.ValidateUpdateListItem?.results || responseJson?.d?.results || [];
  const fieldErrors = fieldResults.filter((entry: { HasException?: boolean }) => entry.HasException);

  if (fieldErrors.length > 0) {
    throw new Error(`Failed to update KM Data Hub URL field: ${fieldErrors.map((entry: { FieldName?: string; ErrorMessage?: string }) => `${entry.FieldName}: ${entry.ErrorMessage}`).join('; ')}`);
  }
};

export const updateItemUrlFieldRaw = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number,
  title: string
): Promise<void> => {
  const fieldMap = await fetchKMDataHubReadFieldMap(spHttpClient, webUrl, LIBRARY_NAME);
  const storedUrl = buildAssetUrl(itemId);
  const resolvedTitle = (title || `Document ${itemId}`).trim();

  const itemResponse = await spHttpClient.get(
    `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})?$select=${fieldMap.status},${fieldMap.url}`,
    SPHttpClient.configurations.v1
  );

  if (!itemResponse.ok) {
    const errorText = await itemResponse.text();
    throw new Error(`Failed to read KM Data Hub item status: ${itemResponse.status} ${errorText}`);
  }

  const itemData = await itemResponse.json();
  const currentStatus = itemData?.d?.[fieldMap.status] ||
    itemData?.[fieldMap.status] || '';
  const currentUrl = itemData?.d?.[fieldMap.url]?.Url ||
    itemData?.[fieldMap.url]?.Url || '';
  const shouldShowUrl = !['draft', 'inactive', 'rejected']
    .includes(String(currentStatus).toLowerCase());

  if (shouldShowUrl && currentUrl.toLowerCase() === storedUrl.toLowerCase()) {
    return;
  }
  if (!shouldShowUrl && !currentUrl) {
    return;
  }

  const response = await spHttpClient.post(
    `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items(${itemId})/ValidateUpdateListItem`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': ''
      },
      body: JSON.stringify({
        formValues: [{
          FieldName: fieldMap.url,
          FieldValue: shouldShowUrl
            ? `${storedUrl}, ${resolvedTitle}`
            : ''
        }],
        bNewDocumentUpdate: false
      })
    }
  );
  if (response && !response.ok) {
    const text = await response.text();
    console.warn('updateItemUrlFieldRaw failed:', response.status, text);
  }
};

export const syncPrettyUrlForItem = async (
  item: {
    Id?: number;
    Title?: string;
    FileLeafRef?: string;
    URL?: { Url?: string };
    Status?: string;
  },
  getTitle?: (item: { Id?: number; Title?: string; FileLeafRef?: string }) => string
): Promise<void> => {
  const itemId = item?.Id;
  if (typeof itemId !== 'number') {
    return;
  }

  const itemStatus = (item?.Status || '').toString().trim().toLowerCase();
  if (itemStatus !== 'active') {
    return;
  }

  const currentUrl = item?.URL?.Url;
  if (currentUrl && currentUrl.trim().length > 0) {
    const expectedTitle = (getTitle ? getTitle(item) : item?.Title || item?.FileLeafRef || `Document ${itemId}`).trim();
    const expectedUrl = buildStoredDocumentUrl(itemId, expectedTitle);
    if (currentUrl.trim().toLowerCase() === expectedUrl.trim().toLowerCase()) {
      return;
    }
  }

  const resolvedTitle = (getTitle ? getTitle(item) : item?.Title || item?.FileLeafRef || `Document ${itemId}`).trim();
  await updateItemUrlField(itemId, resolvedTitle || `Document ${itemId}`);
};

export const syncPrettyUrlsForItems = async (
  items: Array<{
    Id?: number;
    Title?: string;
    FileLeafRef?: string;
    URL?: { Url?: string };
    Status?: string;
  }>,
  getTitle?: (item: { Id?: number; Title?: string; FileLeafRef?: string }) => string
): Promise<void> => {
  await Promise.all(
    (items || []).map(async (item) => {
      try {
        await syncPrettyUrlForItem(item, getTitle);
      } catch (error) {
        console.warn(`Unable to sync pretty URL for KM Data Hub item ${item?.Id}:`, error);
      }
    })
  );
};

export const getItemIdFromHash = (hash: string): number | null => {
  const normalizedHash = (hash || '').trim();
  if (!normalizedHash.startsWith('#/')) {
    return null;
  }

  const routeValue = normalizedHash.substring(2);
  if (!routeValue) {
    return null;
  }

  const separatorIndex = routeValue.indexOf('-');
  const itemIdText = separatorIndex === -1 ? routeValue : routeValue.substring(0, separatorIndex);
  const itemId = parseInt(itemIdText, 10);

  return Number.isFinite(itemId) ? itemId : null;
};

export const isTopicHomePage = (locationLike: Pick<Location, 'pathname'> = window.location): boolean =>
  (locationLike.pathname || '').toLowerCase() === TOPIC_HOME_PATH;

export const getTopicHomeWebAbsoluteUrl = (): string => WEB_ABSOLUTE_URL;
