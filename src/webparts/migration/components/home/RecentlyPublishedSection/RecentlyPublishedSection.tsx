import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './RecentlyPublishedSection.module.scss';
import { IRecentlyPublishedSectionProps } from './IRecentlyPublishedSectionProps';
import { DocumentDetailPage } from '../../../pages/DocumentDetailPage/DocumentDetailPage';
import { ViewAllDocumentsPage } from '../../../pages/ViewAllDocumentsPage/ViewAllDocumentsPage';
import {
  NAV_PATHS,
  openAppPageInNewTab,
  configurePermalinkService
} from '../../../services/permalinkService';
import { getDocumentMetricsBatch, recordUserEvent } from '../../../services/kmMetricsService';
import { SharePointSearchService } from '../../../services/SharePointSearchService';
import {
  buildKMDataHubItemQuery,
  fetchKMDataHubReadFieldMap,
  formatDocumentPublishedDate,
  resolveDocumentAuthor,
  splitDocumentAuthorDisplay,
  resolveDocumentPublishedValue
} from '../../../utils/documentMetadata';
import { subscribeToDocumentDataChanged } from '../../../services/documentChangeEvents';
import { COLUMN_NAMES, LIBRARY_NAMES } from '../../../config/appConfig';
import { downloadSharePointFile } from '../../../utils/fileDownload';
import {
  fetchRecentlyPublishedCacheIds,
  saveRecentlyPublishedCacheIds
} from '../../../services/recentlyPublishedCache';

interface IRecentDocument {
  id: number;
  title: string;
  name: string;
  fileName: string;
  abstract: string;
  fileType: string;
  author: string;
  date: string;
  serverRelativeUrl: string;
  views: number;
  comments: number;
  likes: number;
  downloads: number;
  contentRefreshDate?: string;
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

const PREFERRED_RECENT_DOC_PATTERNS = [
  /oncology med info support/i,
  /support_case/i,
  /lessons learnt/i
];

let fieldMapModuleCache: any = null;
const RECENTLY_PUBLISHED_FETCH_LIMIT = 8;
const RECENTLY_PUBLISHED_MAX_VISIBLE_TILES = 8;
const RECENTLY_PUBLISHED_CACHE_TTL_MS = 30 * 1000;
const RECENTLY_PUBLISHED_TILE_COMFORTABLE_WIDTH = 300;
const RECENTLY_PUBLISHED_TILE_GAP = 24;
const RECENTLY_PUBLISHED_MAX_COLUMNS = 5;
const RECENTLY_PUBLISHED_MIN_COLUMNS = 2;
let recentlyPublishedDocumentsCache: { fetchedAt: number; items: IRecentDocument[] } | null = null;

// Session-level caches to track which lists have been verified and store user titles
const VERIFIED_LISTS: Set<string> = new Set();
const AUTHOR_CACHE: Record<number, string> = {};

const getRecentlyPublishedVisibleCount = (cardsPerRow: number): number => {
  switch (cardsPerRow) {
    case 5:
      return 5;
    case 4:
      return 8;
    case 3:
    case 2:
      return 6;
    default:
      return 6;
  }
};

const renderStatIcon = (type: 'views' | 'comments' | 'likes' | 'downloads'): JSX.Element => {
  if (type === 'views') {
    return (
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    );
  }

  if (type === 'comments') {
    return (
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M21 11.5a8.5 8.5 0 0 1-8.5 8.5c-1.2 0-2.4-.25-3.47-.72L3 21l1.9-4.5A8.47 8.47 0 0 1 4 12a8.5 8.5 0 1 1 17 0Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    );
  }

  if (type === 'downloads') {
    return (
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M12 3v11" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        <path d="M8 10l4 4 4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        <path d="M4 17v1a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-1" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    );
  }

  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M12 21s-6.7-4.35-9.33-8.08C.7 10.13 1.3 6.2 4.6 4.42c2.3-1.24 4.9-.55 6.4 1.1 1.5-1.65 4.1-2.34 6.4-1.1 3.3 1.78 3.9 5.71 1.93 8.5C18.7 16.65 12 21 12 21Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
};

const renderAuthorLines = (author: string, containerClassName: string, itemClassName: string): JSX.Element => {
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

const renderFileTypeIcon = (fileType: string): JSX.Element => {
  const normalizedType = (fileType || '').toLowerCase();

  if (['mp4', 'mov', 'webm', 'mkv', 'avi', 'm4v', 'wmv'].indexOf(normalizedType) !== -1 || normalizedType.indexOf('video') !== -1) {
    return (
      <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
        <rect x="1.5" y="2.5" width="13" height="11" rx="2" fill="#5D2CC9" opacity="0.16" />
        <path d="M6.25 5.2L11 8L6.25 10.8V5.2Z" fill="#5D2CC9" />
      </svg>
    );
  }

  return (
    <svg width="14" height="14" viewBox="0 0 16 16" fill="none" aria-hidden="true">
      <rect x="1" y="1" width="4" height="14" rx="1.2" fill="#7AE582" />
      <rect x="6" y="4" width="4" height="11" rx="1.2" fill="#58A6FF" />
      <rect x="11" y="2" width="4" height="13" rx="1.2" fill="#FF6B9D" />
      <rect x="1" y="11" width="14" height="4" rx="1.2" fill="#9B7BFF" opacity="0.28" />
    </svg>
  );
};

const getFileTypeBadge = (fileType: string): string => {
  if (!fileType) return 'DOC';
  const type = fileType.toUpperCase();
  return type;
};

const getItemFieldValue = (item: any, internalName?: string, fallbackNames: string[] = []): unknown => {
  const candidates = [internalName, ...fallbackNames].filter(Boolean) as string[];
  for (const candidate of candidates) {
    if (item && item[candidate] !== undefined && item[candidate] !== null) {
      return item[candidate];
    }
  }
  return undefined;
};

const normalizeTextField = (value: unknown): string => {
  if (value === undefined || value === null) {
    return '';
  }

  if (typeof value === 'string') {
    return value.trim();
  }

  if (Array.isArray(value)) {
    return value
      .map((entry) => normalizeTextField(entry))
      .filter(Boolean)
      .join('; ');
  }

  if (typeof value === 'object') {
    const objectValue = value as any;
    return String(objectValue.Label || objectValue.Title || objectValue.Name || objectValue.Value || '').trim();
  }

  return String(value).trim();
};

const normalizeDescriptionText = (value: unknown): string => {
  const text = normalizeTextField(value);
  return !text || text === '-' ? '' : text;
};

const normalizeNumberField = (value: unknown): number | undefined => {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const numericValue = Number(normalizeTextField(value));
  return Number.isFinite(numericValue) ? numericValue : undefined;
};

const normalizeBooleanField = (value: unknown): boolean | undefined => {
  if (typeof value === 'boolean') {
    return value;
  }

  const normalizedValue = normalizeTextField(value).toLowerCase();
  if (['true', 'yes', '1'].indexOf(normalizedValue) !== -1) {
    return true;
  }
  if (['false', 'no', '0'].indexOf(normalizedValue) !== -1) {
    return false;
  }
  return undefined;
};

const normalizeStringArrayField = (value: unknown): string[] =>
  normalizeTextField(value)
    .split(/[;,\n]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);

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

const escapeXmlValue = (value: string): string =>
  String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');

const RecentlyPublishedSection: React.FC<IRecentlyPublishedSectionProps> = ({
  context,
  sectionStyle,
  isLearner = false,
  onViewAllOpen,
  onViewDocument
}) => {
  const [documents, setDocuments] = React.useState<IRecentDocument[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [selectedDocumentId, setSelectedDocumentId] = React.useState<number | null>(null);
  const [showViewAll, setShowViewAll] = React.useState<boolean>(false);
  const [detailBackTarget, setDetailBackTarget] = React.useState<'section' | 'library'>('section');
  const [metricsRefreshKey, setMetricsRefreshKey] = React.useState<number>(0);
  const [downloadingDocumentIds, setDownloadingDocumentIds] = React.useState<Record<number, boolean>>({});
  const [cardsPerRow, setCardsPerRow] = React.useState<number>(RECENTLY_PUBLISHED_MAX_COLUMNS);
  const tilesContainerRef = React.useRef<HTMLDivElement | null>(null);
  const normalLayoutBaselineRef = React.useRef<{ width: number; devicePixelRatio: number } | null>(null);
  const previousUrlRef = React.useRef<string | null>(null);

  React.useEffect(() => {
    if (!context) {
      return;
    }

    configurePermalinkService(context.spHttpClient, context.pageContext.web.absoluteUrl);
  }, [context]);

  const ensureListExists = React.useCallback(
    async (
      listName: string,
      fields: Array<{ name: string; type: string; required?: boolean }>
    ): Promise<boolean> => {
      if (!context) return false;
      
      if (VERIFIED_LISTS.has(listName)) {
        return true;
      }

      const webUrl = context.pageContext.web.absoluteUrl;

      try {
        const checkResp = await context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${listName}')?$select=Id`,
          SPHttpClient.configurations.v1
        );

        if (checkResp.ok) {
          VERIFIED_LISTS.add(listName);
          return true;
        }

        const createListResp = await context.spHttpClient.post(
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
            FieldTypeKind:
              field.type === 'Number'
                ? 9
                : field.type === 'Note'
                ? 3
                : 2,
            Required: field.required || false
          };

          const fieldResp = await context.spHttpClient.post(
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

        VERIFIED_LISTS.add(listName);
        return true;
      } catch {
        return false;
      }
    },
    [context]
  );

  const fetchRecentMetrics = React.useCallback(
    async (documentIds: number[]) => {
      if (!context || documentIds.length === 0) {
        return {};
      }

      try {
        return await getDocumentMetricsBatch(context, documentIds);
      } catch {
        return {};
      }
    },
    [context]
  );

  const getListEntityType = React.useCallback(async (listName: string): Promise<string | null> => {
    if (!context) {
      return null;
    }

    const webUrl = context.pageContext.web.absoluteUrl;
    const response = await context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${listName}')?$select=ListItemEntityTypeFullName`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      return null;
    }

    const listInfo = await response.json();
    return listInfo.ListItemEntityTypeFullName || null;
  }, [context]);

  const recordDownloadMetric = React.useCallback(async (documentId: number, documentTitle?: string): Promise<void> => {
    if (!context) {
      return;
    }

    const userResponse = await context.spHttpClient.get(
      `${context.pageContext.web.absoluteUrl}/_api/web/currentuser?$select=Id`,
      SPHttpClient.configurations.v1
    );

    if (!userResponse.ok) {
      return;
    }

    const user = await userResponse.json();
    const currentUserId = user.Id;
    if (!currentUserId) {
      return;
    }

    await recordUserEvent(context, documentId, currentUserId, 'Download', documentTitle);
  }, [context]);

  const fetchAuthorTitlesByIds = React.useCallback(
    async (authorIds: number[]): Promise<Record<number, string>> => {
      if (!context || authorIds.length === 0) {
        return {};
      }

      const authorMap: Record<number, string> = { ...AUTHOR_CACHE };
      authorIds.forEach((authorId) => {
        if (typeof authorId !== 'number' || authorId <= 0 || authorMap[authorId]) {
          return;
        }
        const cachedTitle = SharePointSearchService.getCachedAuthorTitle(authorId);
        if (cachedTitle) {
          AUTHOR_CACHE[authorId] = cachedTitle;
          authorMap[authorId] = cachedTitle;
        }
      });

      const uniqueAuthorIds = Array.from(new Set(authorIds.filter((id) => typeof id === 'number' && id > 0 && !authorMap[id])));
      if (uniqueAuthorIds.length === 0) {
        return authorMap;
      }

      const webUrl = context.pageContext.web.absoluteUrl;

      await Promise.all(
        uniqueAuthorIds.map(async (authorId) => {
          try {
            const response = await context.spHttpClient.get(
              `${webUrl}/_api/web/getuserbyid(${authorId})?$select=Id,Title`,
              SPHttpClient.configurations.v1
            );

            if (!response.ok) {
              return;
            }

            const user = await response.json();
            if (user?.Title) {
              AUTHOR_CACHE[authorId] = user.Title;
              SharePointSearchService.setCachedAuthorTitle(authorId, user.Title);
              authorMap[authorId] = user.Title;
            }
          } catch {
            // Ignore
          }
        })
      );

      return authorMap;
    },
    [context]
  );

  React.useEffect(() => {
    return subscribeToDocumentDataChanged(() => {
      setMetricsRefreshKey((value) => value + 1);
    });
  }, []);

  React.useEffect(() => {
    const container = tilesContainerRef.current;
    if (!container) {
      return;
    }

    const updateCardsPerRow = (): void => {
      const availableWidth = container.clientWidth;
      const currentDevicePixelRatio = window.devicePixelRatio || 1;
      const currentBaseline = normalLayoutBaselineRef.current;

      if (!currentBaseline || availableWidth > currentBaseline.width) {
        normalLayoutBaselineRef.current = {
          width: Math.max(availableWidth, currentBaseline?.width || 0),
          devicePixelRatio: currentBaseline?.devicePixelRatio || currentDevicePixelRatio
        };
      }

      const baseline = normalLayoutBaselineRef.current;
      const isAtOrBelowNormalZoom = currentDevicePixelRatio <= baseline.devicePixelRatio + 0.01;

      if (isAtOrBelowNormalZoom) {
        setCardsPerRow((current) => current === RECENTLY_PUBLISHED_MAX_COLUMNS ? current : RECENTLY_PUBLISHED_MAX_COLUMNS);
        return;
      }

      const computedGap = parseFloat(window.getComputedStyle(container).columnGap || '');
      const tileGap = Number.isFinite(computedGap) ? computedGap : RECENTLY_PUBLISHED_TILE_GAP;
      const nextCardsPerRow = Math.max(
        RECENTLY_PUBLISHED_MIN_COLUMNS,
        Math.min(
          RECENTLY_PUBLISHED_MAX_COLUMNS,
          Math.floor(
            (availableWidth + tileGap) /
            (RECENTLY_PUBLISHED_TILE_COMFORTABLE_WIDTH + tileGap)
          )
        )
      );

      setCardsPerRow((current) => current === nextCardsPerRow ? current : nextCardsPerRow);
    };

    updateCardsPerRow();

    if (typeof ResizeObserver !== 'undefined') {
      const resizeObserver = new ResizeObserver(updateCardsPerRow);
      resizeObserver.observe(container);
      return () => resizeObserver.disconnect();
    }

    window.addEventListener('resize', updateCardsPerRow);
    return () => window.removeEventListener('resize', updateCardsPerRow);
  }, []);

  React.useEffect(() => {
    const fetchLatestDocuments = async (): Promise<void> => {
      if (!context) {
        setLoading(false);
        return;
      }

      try {
        if (
          recentlyPublishedDocumentsCache &&
          Date.now() - recentlyPublishedDocumentsCache.fetchedAt < RECENTLY_PUBLISHED_CACHE_TTL_MS
        ) {
          setDocuments(recentlyPublishedDocumentsCache.items);
          setLoading(false);
        }

        const webUrl = context.pageContext.web.absoluteUrl;
        const libraryName = LIBRARY_NAMES.kmDataHub;
        const fieldMap = fieldMapModuleCache
          ? fieldMapModuleCache
          : await fetchKMDataHubReadFieldMap(context.spHttpClient, webUrl, libraryName);
        fieldMapModuleCache = fieldMap;
        const statusField = fieldMap.status || COLUMN_NAMES.status;
        const publishedField = fieldMap.published || COLUMN_NAMES.published;
        const escapedLibraryName = libraryName.replace(/'/g, "''");
        const queryParts = buildKMDataHubItemQuery(
          fieldMap,
          [
            'ID',
            fieldMap.title || COLUMN_NAMES.title,
            fieldMap.description || COLUMN_NAMES.description,
            COLUMN_NAMES.created,
            'AuthorId',
            fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef,
            fieldMap.fileRef || COLUMN_NAMES.fileRef,
            'File_x0020_Type',
            'File/UniqueId',
            fieldMap.url || COLUMN_NAMES.url,
            fieldMap.docIcon || COLUMN_NAMES.docIcon,
            fieldMap.status || COLUMN_NAMES.status,
            fieldMap.published || COLUMN_NAMES.created
          ],
          ['File']
        );

        const fetchItemsByIds = async (ids: number[]): Promise<any[]> => {
          const normalizedIds = Array.from(new Set(ids.filter((id) => Number.isFinite(id) && id > 0)))
            .slice(0, RECENTLY_PUBLISHED_FETCH_LIMIT);
          if (normalizedIds.length === 0) {
            return [];
          }

          const idFilter = normalizedIds.map((id: number) => `ID eq ${id}`).join(' or ');
          const detailsResp = await context.spHttpClient.get(
            `${webUrl}/_api/web/lists/getbytitle('${escapedLibraryName}')/items` +
            `?$select=${queryParts.select}` +
            `&$expand=${queryParts.expand}` +
            `&$filter=${encodeURIComponent(idFilter)}` +
            `&$top=${normalizedIds.length}`,
            SPHttpClient.configurations.v1,
            { headers: { Accept: 'application/json;odata=verbose' } }
          );

          if (!detailsResp.ok) {
            let detailsErrorText = '';
            try {
              detailsErrorText = await detailsResp.text();
            } catch {
              detailsErrorText = '';
            }
            console.warn('Recently Published cached details query failed:', detailsResp.status, detailsResp.statusText, detailsErrorText);
            return [];
          }

          const detailsData = await detailsResp.json();
          const detailItems = detailsData.value || detailsData?.d?.results || [];
          const orderById = new Map<number, number>(
            normalizedIds.map((id: number, index: number) => [id, index] as [number, number])
          );
          return detailItems.sort((a: any, b: any) => {
            const leftIndex = orderById.get(Number(a.ID || a.Id));
            const rightIndex = orderById.get(Number(b.ID || b.Id));
            return (leftIndex === undefined ? Number.MAX_SAFE_INTEGER : leftIndex) -
              (rightIndex === undefined ? Number.MAX_SAFE_INTEGER : rightIndex);
          });
        };

        let rawItems: any[] = [];
        const cachedIds = await fetchRecentlyPublishedCacheIds(context.spHttpClient, webUrl);
        rawItems = await fetchItemsByIds(cachedIds);

        if (rawItems.length === 0) {
        const viewXml =
          `<View Scope="RecursiveAll">` +
          `<Query>` +
          `<Where>` +
          `<Eq><FieldRef Name="${escapeXmlValue(statusField)}" /><Value Type="Text">Active</Value></Eq>` +
          `</Where>` +
          `<OrderBy><FieldRef Name="${escapeXmlValue(publishedField)}" Ascending="FALSE" /></OrderBy>` +
          `</Query>` +
          `<RowLimit Paged="FALSE">${RECENTLY_PUBLISHED_FETCH_LIMIT}</RowLimit>` +
          `</View>`;
        const latestItemsResp = await context.spHttpClient.post(
          `${webUrl}/_api/web/lists/getbytitle('${escapedLibraryName}')/RenderListDataAsStream`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose',
              'Content-Type': 'application/json;odata=verbose'
            },
            body: JSON.stringify({
              parameters: {
                ViewXml: viewXml,
                RenderOptions: 2,
                DatesInUtc: true
              }
            })
          }
        );

        if (latestItemsResp.ok) {
          const latestItemsData = await latestItemsResp.json();
          const rows =
            latestItemsData.Row ||
            latestItemsData?.ListData?.Row ||
            latestItemsData?.d?.RenderListDataAsStream?.Row ||
            [];
          const latestIds = rows
            .map((row: any) => Number(row.ID || row.Id))
            .filter((id: number) => Number.isFinite(id) && id > 0)
            .slice(0, RECENTLY_PUBLISHED_FETCH_LIMIT);

          if (latestIds.length > 0) {
            rawItems = await fetchItemsByIds(latestIds);
            void saveRecentlyPublishedCacheIds(context.spHttpClient, webUrl, latestIds);
          }
        } else {
          let errorText = '';
          try {
            errorText = await latestItemsResp.text();
          } catch {
            errorText = '';
          }
          console.warn('Recently Published CAML query failed:', latestItemsResp.status, latestItemsResp.statusText, errorText);
        }
        }

        const items = rawItems
          .filter((item: any) => normalizeTextField(getItemFieldValue(item, statusField, [COLUMN_NAMES.status])) === 'Active')
          .sort((a: any, b: any) => {
            const dateA = new Date(resolveDocumentPublishedValue(a, fieldMap.published) || a.Created || 0).getTime();
            const dateB = new Date(resolveDocumentPublishedValue(b, fieldMap.published) || b.Created || 0).getTime();
            return dateB - dateA;
          })
          .slice(0, RECENTLY_PUBLISHED_MAX_VISIBLE_TILES);

        const authorIds = items
          .filter((item: any) => !item?.Author?.Title && typeof item?.AuthorId === 'number')
          .map((item: any) => item.AuthorId);
        const itemDocumentIds = items
          .map((item: any) => Number(item.Id))
          .filter((id: number) => Number.isFinite(id) && id > 0);
        const [authorLookup, metricsByDocumentId] = await Promise.all([
          fetchAuthorTitlesByIds(authorIds),
          fetchRecentMetrics(itemDocumentIds)
        ]);

        const formattedItemsBase: IRecentDocument[] = items.map((item: any) => {
          const resolvedTitle = normalizeTextField(getItemFieldValue(item, fieldMap.title, [COLUMN_NAMES.title]));
          const fileName = normalizeTextField(getItemFieldValue(item, fieldMap.fileLeafRef, [COLUMN_NAMES.fileLeafRef])) || item.File?.Name || item.FileRef?.split('/').pop() || '';
          let fileExtension = item.File_x0020_Type || (fileName.includes('.') ? fileName.split('.').pop() : '');
          
          if (!fileExtension && resolvedTitle.includes('.')) {
            fileExtension = resolvedTitle.split('.').pop();
          }

          fileExtension = (fileExtension || '').toUpperCase();

          let displayName = fileName;
          if (!displayName || displayName === '') {
            displayName = resolvedTitle || `Document ${item.Id}`;
          }

          let abstract = normalizeDescriptionText(
            getItemFieldValue(item, fieldMap.description, [COLUMN_NAMES.description])
          );
          let serverRelativeUrl = normalizeTextField(getItemFieldValue(item, fieldMap.fileRef, [COLUMN_NAMES.fileRef]));

          if (serverRelativeUrl && !serverRelativeUrl.startsWith('/')) {
            serverRelativeUrl = `/${serverRelativeUrl}`;
          }

          if (!serverRelativeUrl && fileName) {
            const siteRelativePath = context.pageContext.web.serverRelativeUrl || '';
            serverRelativeUrl = `${siteRelativePath}/${LIBRARY_NAMES.kmDataHub}/${fileName}`;
          }

          if (abstract === '-' || abstract === '') {
            abstract = '';
          }

          return {
            id: item.Id,
            title: resolvedTitle || displayName,
            name: displayName,
            fileName,
            abstract,
            fileType: fileExtension,
            author: resolveDocumentAuthor(item, authorLookup[item.AuthorId] || 'Internal', fieldMap.author),
            date: formatDocumentPublishedDate(resolveDocumentPublishedValue(item, fieldMap.published), 'en-US', {
              month: 'short',
              day: 'numeric',
              year: 'numeric'
            }),
            serverRelativeUrl,
            views: 0,
            comments: 0,
            likes: 0,
            downloads: 0,
            reviewerComments: '',
            docIcon: normalizeTextField(getItemFieldValue(item, fieldMap.docIcon, [COLUMN_NAMES.docIcon]))
          };
        });

        const formattedItems = formattedItemsBase.map((item) => ({
          ...item,
          views: metricsByDocumentId[item.id]?.views ?? 0,
          comments: metricsByDocumentId[item.id]?.comments ?? 0,
          likes: metricsByDocumentId[item.id]?.likes ?? 0,
          downloads: metricsByDocumentId[item.id]?.downloads ?? 0
        }));

        const getItemSearchText = (item: any): string =>
          [
            item.Title,
            item.TitleName,
            item.FileLeafRef,
            item.FileRef,
            item.Abstract,
            item.Author0,
            item.Author?.Title
          ]
            .filter(Boolean)
            .join(' ')
            .toLowerCase();

        const preferredIds = PREFERRED_RECENT_DOC_PATTERNS
          .map((pattern) => items.find((item: any) => pattern.test(getItemSearchText(item)))?.Id)
          .filter((id): id is number => typeof id === 'number');

        const preferredDocuments = preferredIds
          .map((id) => formattedItems.find((doc) => doc.id === id))
          .filter((doc): doc is IRecentDocument => Boolean(doc));

        const remainingDocuments = formattedItems.filter(
          (doc) => !preferredDocuments.some((preferredDoc) => preferredDoc.id === doc.id)
        );

        const latestDocuments = [...preferredDocuments, ...remainingDocuments];
        recentlyPublishedDocumentsCache = {
          fetchedAt: Date.now(),
          items: latestDocuments
        };
        setDocuments(latestDocuments);
      } catch (error) {
        console.error('fetchLatestDocuments error:', error);
        setDocuments([]);
      } finally {
        setLoading(false);
      }
    };

    fetchLatestDocuments().catch(() => setLoading(false));
  }, [context, fetchAuthorTitlesByIds, fetchRecentMetrics, metricsRefreshKey]);

  const visibleTileCount = getRecentlyPublishedVisibleCount(cardsPerRow);
  const displayTiles = React.useMemo(
    () => documents.slice(0, visibleTileCount),
    [documents, visibleTileCount]
  );

  const updateBrowserUrlForDocument = React.useCallback((itemId: number, _title: string): void => {
    openAppPageInNewTab(NAV_PATHS.asset, { assetID: itemId.toString() });
  }, []);

  const restorePreviousUrl = React.useCallback((): void => {
    if (!previousUrlRef.current) {
      return;
    }

    window.history.replaceState({}, document.title, previousUrlRef.current);
    previousUrlRef.current = null;
  }, []);

  React.useEffect(() => {
    if (!selectedDocumentId || detailBackTarget !== 'section') {
      return;
    }

    const isDocumentStillPublished = documents.some((document) => document.id === selectedDocumentId);
    if (!isDocumentStillPublished) {
      setSelectedDocumentId(null);
      restorePreviousUrl();
    }
  }, [detailBackTarget, documents, restorePreviousUrl, selectedDocumentId]);

  const handleDownloadRecorded = React.useCallback((): void => {
    setMetricsRefreshKey((value) => value + 1);
  }, []);

  const handleView = (item: IRecentDocument): void => {
    if (!item.id) {
      return;
    }

    if (onViewDocument) {
      onViewDocument(item.id);
      return;
    }

    updateBrowserUrlForDocument(item.id, item.title || item.name);
  };

  const handleDownload = async (item: IRecentDocument): Promise<void> => {
    if (!context || !item.serverRelativeUrl) {
      return;
    }
    if (downloadingDocumentIds[item.id]) {
      return;
    }

    const webUrl = context.pageContext.web.absoluteUrl;
    let serverRelativeUrl = item.serverRelativeUrl;

    if (!serverRelativeUrl.startsWith('/')) {
      serverRelativeUrl = `/${serverRelativeUrl}`;
    }

    setDownloadingDocumentIds((current) => ({ ...current, [item.id]: true }));
    try {
      await downloadSharePointFile(
        context,
        webUrl,
        serverRelativeUrl,
        buildTitleDownloadFileName(item.title, item.fileName, item.fileType)
      );
      void recordDownloadMetric(item.id, item.title || item.name)
        .then(handleDownloadRecorded)
        .catch(() => undefined);
    } catch (error) {
      console.error('Recently published download error:', error);
    } finally {
      window.setTimeout(() => {
        setDownloadingDocumentIds((current) => {
          const next = { ...current };
          delete next[item.id];
          return next;
        });
      }, 1200);
    }
  };

  const handleViewAll = (): void => {
    if (onViewAllOpen) {
      onViewAllOpen();
      return;
    }

    openAppPageInNewTab(NAV_PATHS.asset);
  };

  const handleCloseViewAll = (): void => {
    setShowViewAll(false);
    restorePreviousUrl();
  };

  const handleViewDocumentFromList = (documentId: number): void => {
    openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
  };

  const handleBackToLibrary = (): void => {
    setSelectedDocumentId(null);
    setShowViewAll(true);
    restorePreviousUrl();
  };

  return (
    <div className={styles.questionSection} style={sectionStyle}>
      <div className={styles.sectionHeader}>
        <h2 className={styles.sectionTitle}>Recently Published</h2>
        <button
          type="button"
          className={styles.viewAllButton}
          onClick={handleViewAll}
          aria-label="View all published documents"
        >
          View All
        </button>
      </div>
      <div className={styles.tilesCarousel}>
        <div
          ref={tilesContainerRef}
          className={styles.tilesContainer}
          style={{ gridTemplateColumns: `repeat(${cardsPerRow}, minmax(0, 1fr))` }}
        >
            {loading ? (
              <div className={styles.loading}>Loading...</div>
            ) : (
              displayTiles.map((doc, index) => (
                <div
                  key={doc ? doc.id : `empty-${index}`}
                  className={styles.tileCard}
                >
                  {doc ? (
                    <>
                      <div className={styles.tileBody}>
                        <div className={styles.tileTopMeta}>
                          <span className={styles.tileTopIcon}>{renderFileTypeIcon(doc.docIcon || doc.fileType)}</span>
                          <span className={styles.tileTopType}>{doc.docIcon || getFileTypeBadge(doc.fileType)}</span>
                        </div>
                        <div className={styles.tileContentSection}>
                          <div className={styles.tileTitleWrap}>
                            <h3 className={styles.tileTitle}>{doc.title || doc.name}</h3>
                            <div className={styles.tileTitleTooltip} role="tooltip">
                              {doc.title || doc.name}
                            </div>
                          </div>
                          <p className={styles.tileAbstract}>{doc.abstract || ''}</p>
                          <div className={styles.tileStatusRow}>
                            <span className={styles.statusLabel}>Status :</span>
                            <span className={`${styles.statusBadge} ${styles.statusBadgeActive}`}>Active</span>
                          </div>
                          <div className={styles.tilePublishMeta}>
                            <div className={styles.tilePublishText}>
                              <span>Published by:</span>
                              {renderAuthorLines(doc.author || 'Internal', styles.authorLineList, styles.authorLineItem)}
                            </div>
                            <p className={styles.tilePublishText}>Published date: {doc.date || '-'}</p>
                          </div>
                        </div>
                      </div>

                      <div className={styles.tileStats}>
                        <span className={styles.tileStat}>
                          {renderStatIcon('views')}
                          <span>{doc.views ?? 0}</span>
                        </span>
                        <span className={styles.tileStat}>
                          {renderStatIcon('comments')}
                          <span>{doc.comments ?? 0}</span>
                        </span>
                        <span className={styles.tileStat}>
                          {renderStatIcon('likes')}
                          <span>{doc.likes ?? 0}</span>
                        </span>
                        <span className={styles.tileStat}>
                          {renderStatIcon('downloads')}
                          <span>{doc.downloads ?? 0}</span>
                        </span>
                      </div>
                      <div className={styles.tileActions}>
                        <button className={styles.viewButton} onClick={() => handleView(doc)}>View</button>
                        {!isLearner && (
                          <button
                            className={styles.downloadButton}
                            disabled={!!downloadingDocumentIds[doc.id]}
                            onClick={() => handleDownload(doc)}
                          >
                            {downloadingDocumentIds[doc.id] ? 'Downloading...' : 'Download'}
                          </button>
                        )}
                      </div>
                    </>
                  ) : (
                    <>
                      <div className={styles.tileBody}>
                        <div className={styles.tileTopMeta}>
                          <span className={styles.tileTopIcon}>📎</span>
                          <span className={styles.tileTopType}>---</span>
                        </div>
                        <div className={styles.tileContentSection}>
                          <h3 className={styles.tileTitle}>No document</h3>
                          <p className={styles.tileAbstract} />
                          <div className={styles.tilePublishMeta}>
                            <p className={styles.tilePublishText}>Published by: -</p>
                            <p className={styles.tilePublishText}>Published date: -</p>
                          </div>
                        </div>
                      </div>
                      <div className={styles.tileStats}>
                        <span className={styles.tileStat}>{renderStatIcon('views')}<span>0</span></span>
                        <span className={styles.tileStat}>{renderStatIcon('comments')}<span>0</span></span>
                        <span className={styles.tileStat}>{renderStatIcon('likes')}<span>0</span></span>
                        <span className={styles.tileStat}>{renderStatIcon('downloads')}<span>0</span></span>
                      </div>
                      <div className={styles.tileActions}>
                        <button className={styles.viewButton} disabled>
                          View
                        </button>
                      </div>
                    </>
                  )}
                </div>
              ))
            )}
          </div>
      </div>
      {showViewAll && context && (
        <div className={styles.viewAllModal}>
          <ViewAllDocumentsPage
            context={context}
            onClose={handleCloseViewAll}
            onViewDocument={handleViewDocumentFromList}
            onDownloadRecorded={handleDownloadRecorded}
            scope="recentlyPublished"
            isLearner={isLearner}
          />
        </div>
      )}
      {selectedDocumentId && context && (
        <div className={styles.detailModal}>
          <DocumentDetailPage
            context={context}
            documentId={selectedDocumentId}
            onMetricsUpdated={handleDownloadRecorded}
            onClose={() => {
              setSelectedDocumentId(null);
              restorePreviousUrl();
            }}
            backTo={detailBackTarget}
            onBackToLibrary={handleBackToLibrary}
            onDownloadRecorded={handleDownloadRecorded}
            hideTopDocumentActions={isLearner}
          />
        </div>
      )}
    </div>
  );
};

export default RecentlyPublishedSection;

