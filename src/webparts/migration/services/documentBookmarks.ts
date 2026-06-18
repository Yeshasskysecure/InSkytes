import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { LIST_NAMES } from '../config/appConfig';

export interface IDocumentBookmarkRecord {
  itemId?: number;
  documentId: number;
  documentUrl: string;
  userId: number;
  timestamp: string;
  title?: string;
  description?: string;
  contributor?: string;
  businessUnit?: string;
  department?: string;
  fileType?: string;
}

export interface ISharePointBookmarkSummary {
  isBookmarked: boolean;
  totalCount: number;
}

export interface ISharePointBookmarkToggleResult extends ISharePointBookmarkSummary {
  success: boolean;
  message: string;
}

const DOCUMENT_METRICS_LIST = LIST_NAMES.documentMetrics;
const USER_INTERACTIONS_LIST = LIST_NAMES.userInteractions;

const wait = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

const escapeODataString = (value: string): string => (value || '').replace(/'/g, "''");

const getODataItems = (data: any): any[] => {
  const items = data?.value || data?.d?.results || [];
  return Array.isArray(items) ? items : [];
};

const getODataNextLink = (data: any): string | null =>
  data?.['@odata.nextLink'] || data?.d?.__next || null;

const getODataProperty = <T,>(data: any, propertyName: string, fallback: T): T =>
  (data?.[propertyName] ?? data?.d?.[propertyName] ?? fallback) as T;

const interactionKey = (documentId: number, userId: number, type: string): string =>
  `${documentId}_${userId}_${type}`;

const isThrottleResponse = (response: SPHttpClientResponse): boolean =>
  response.status === 429 || response.status === 503;

const getRetryDelayMs = (response: SPHttpClientResponse, attempt: number): number => {
  const retryAfter = response.headers.get('Retry-After');
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds) && seconds > 0) {
      return seconds * 1000;
    }
  }

  return Math.min(1000 * Math.pow(2, attempt), 15000);
};

const requestWithThrottleRetry = async (
  request: () => Promise<SPHttpClientResponse>,
  maxRetries: number = 4
): Promise<SPHttpClientResponse> => {
  let lastResponse: SPHttpClientResponse | null = null;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const response = await request();
    lastResponse = response;

    if (!isThrottleResponse(response)) {
      return response;
    }

    if (attempt < maxRetries) {
      await wait(getRetryDelayMs(response, attempt));
    }
  }

  return lastResponse!;
};

const getListEntityType = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listName: string
): Promise<string | null> => {
  const response = await requestWithThrottleRetry(() =>
    spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${escapeODataString(listName)}')?$select=ListItemEntityTypeFullName`,
      SPHttpClient.configurations.v1
    )
  );

  if (!response.ok) {
    return null;
  }

  const data = await response.json();
  return getODataProperty<string | null>(data, 'ListItemEntityTypeFullName', null);
};

const getBookmarkInteraction = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  userId: number,
  documentId: number
): Promise<{ Id: number } | null> => {
  const key = interactionKey(documentId, userId, 'Bookmark');
  const response = await requestWithThrottleRetry(() =>
    spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items` +
      `?$select=Id&$filter=InteractionKey eq '${escapeODataString(key)}'&$top=1`,
      SPHttpClient.configurations.v1
    )
  );

  if (!response.ok) {
    return null;
  }

  const data = await response.json();
  return getODataItems(data)[0] || null;
};

const getDocumentMetricsBookmarkCount = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  documentId: number
): Promise<number> => {
  const response = await requestWithThrottleRetry(() =>
    spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items` +
      `?$select=BookmarkCount&$filter=DocumentId eq ${documentId}&$top=1`,
      SPHttpClient.configurations.v1
    )
  );

  if (!response.ok) {
    return 0;
  }

  const data = await response.json();
  const item = getODataItems(data)[0];
  return Math.max(0, Math.floor(Number(item?.BookmarkCount || 0)));
};

const findDocumentMetricsItem = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  documentId: number
): Promise<any | null> => {
  const response = await requestWithThrottleRetry(() =>
    spHttpClient.get(
      `${siteUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items` +
      `?$select=Id,BookmarkCount&$filter=DocumentId eq ${documentId}&$top=1`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose'
        }
      }
    )
  );

  if (!response.ok) {
    return null;
  }

  const data = await response.json();
  return getODataItems(data)[0] || null;
};

const incrementBookmarkCount = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  documentId: number,
  delta: number
): Promise<number> => {
  let nextCount = 0;

  for (let attempt = 0; attempt < 4; attempt++) {
    const item = await findDocumentMetricsItem(spHttpClient, siteUrl, documentId);
    if (!item?.Id) {
      return 0;
    }

    const currentCount = Math.max(0, Math.floor(Number(item.BookmarkCount || 0)));
    nextCount = Math.max(0, currentCount + delta);
    const etag = item.__metadata?.etag;
    if (!etag) {
      await wait(100 * (attempt + 1));
      continue;
    }

    const response = await requestWithThrottleRetry(() =>
      spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items(${item.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': etag,
            'X-HTTP-Method': 'MERGE',
            'odata-version': ''
          },
          body: JSON.stringify({
            BookmarkCount: nextCount,
            LastUpdated: new Date().toISOString()
          })
        }
      )
    );

    if (response.ok) {
      return nextCount;
    }

    if (response.status === 412) {
      await wait(100 * (attempt + 1));
      continue;
    }

    break;
  }

  return nextCount;
};

const createBookmarkInteraction = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  userId: number,
  documentId: number,
  title?: string
): Promise<boolean> => {
  const entityType = await getListEntityType(spHttpClient, siteUrl, USER_INTERACTIONS_LIST);
  if (!entityType) {
    return false;
  }

  const key = interactionKey(documentId, userId, 'Bookmark');
  const response = await requestWithThrottleRetry(() =>
    spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: JSON.stringify({
          __metadata: { type: entityType },
          Title: title || key,
          DocumentId: documentId,
          UserId: userId,
          InteractionType: 'Bookmark',
          InteractionKey: key
        })
      }
    )
  );

  return response.ok;
};

const deleteBookmarkInteraction = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  itemId: number
): Promise<boolean> => {
  const response = await requestWithThrottleRetry(() =>
    spHttpClient.post(
      `${siteUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE',
          'odata-version': ''
        }
      }
    )
  );

  return response.ok;
};

export const readAllDocumentBookmarksFromSharePoint = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  userId: number
): Promise<IDocumentBookmarkRecord[]> => {
  if (!spHttpClient || !siteUrl || userId <= 0) {
    return [];
  }

  const bookmarks: IDocumentBookmarkRecord[] = [];
  let nextUrl: string | null =
    `${siteUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items` +
    `?$select=Id,Created,DocumentId,UserId&$filter=UserId eq ${userId} and InteractionType eq 'Bookmark'&$top=500`;

  while (nextUrl) {
    const response = await requestWithThrottleRetry(() =>
      spHttpClient.get(nextUrl!, SPHttpClient.configurations.v1)
    );

    if (!response.ok) {
      break;
    }

    const data = await response.json();
    getODataItems(data).forEach((item) => {
      const documentId = Number(item.DocumentId || 0);
      const bookmarkUserId = Number(item.UserId || 0);
      if (documentId > 0 && bookmarkUserId > 0) {
        bookmarks.push({
          itemId: Number(item.Id || 0),
          documentId,
          documentUrl: '',
          userId: bookmarkUserId,
          timestamp: item.Created || new Date().toISOString()
        });
      }
    });

    nextUrl = getODataNextLink(data);
  }

  return bookmarks;
};

export const isDocumentBookmarkedInSharePoint = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  userId: number,
  documentId: number
): Promise<boolean> => {
  return !!(await getBookmarkInteraction(spHttpClient, siteUrl, userId, documentId));
};

export const getDocumentBookmarkSummaryInSharePoint = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  userId: number,
  documentId: number
): Promise<ISharePointBookmarkSummary> => {
  if (!spHttpClient || !siteUrl || documentId <= 0) {
    return { isBookmarked: false, totalCount: 0 };
  }

  const [interaction, totalCount] = await Promise.all([
    userId > 0 ? getBookmarkInteraction(spHttpClient, siteUrl, userId, documentId) : Promise.resolve(null),
    getDocumentMetricsBookmarkCount(spHttpClient, siteUrl, documentId)
  ]);

  return {
    isBookmarked: !!interaction,
    totalCount
  };
};

export const toggleDocumentBookmarkInSharePoint = async (
  spHttpClient: SPHttpClient,
  siteUrl: string,
  userId: number,
  documentId: number,
  _documentUrl: string,
  snapshot?: Partial<Pick<IDocumentBookmarkRecord, 'title' | 'description' | 'contributor' | 'businessUnit' | 'department' | 'fileType'>>
): Promise<ISharePointBookmarkToggleResult> => {
  if (!spHttpClient || !siteUrl || userId <= 0 || documentId <= 0) {
    return {
      success: false,
      isBookmarked: false,
      totalCount: 0,
      message: 'Unable to bookmark. Please refresh and try again.'
    };
  }

  try {
    const existing = await getBookmarkInteraction(spHttpClient, siteUrl, userId, documentId);

    if (existing?.Id) {
      const removed = await deleteBookmarkInteraction(spHttpClient, siteUrl, existing.Id);
      const totalCount = removed
        ? await incrementBookmarkCount(spHttpClient, siteUrl, documentId, -1)
        : await getDocumentMetricsBookmarkCount(spHttpClient, siteUrl, documentId);

      return {
        success: removed,
        isBookmarked: !removed,
        totalCount,
        message: removed ? 'Bookmark removed' : 'Unable to remove bookmark. Please try again.'
      };
    }

    const created = await createBookmarkInteraction(spHttpClient, siteUrl, userId, documentId, snapshot?.title);
    const totalCount = created
      ? await incrementBookmarkCount(spHttpClient, siteUrl, documentId, 1)
      : await getDocumentMetricsBookmarkCount(spHttpClient, siteUrl, documentId);

    return {
      success: created,
      isBookmarked: created,
      totalCount,
      message: created ? 'Bookmark successful' : 'Unable to bookmark. Please try again.'
    };
  } catch (error) {
    console.warn('Unable to update SharePoint personal bookmark:', error);
    const summary = await getDocumentBookmarkSummaryInSharePoint(spHttpClient, siteUrl, userId, documentId);
    return {
      success: false,
      isBookmarked: summary.isBookmarked,
      totalCount: summary.totalCount,
      message: 'Unable to update bookmark. Please try again.'
    };
  }
};
