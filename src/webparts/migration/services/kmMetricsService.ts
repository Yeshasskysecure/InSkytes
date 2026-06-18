import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { COLUMN_NAMES, LIBRARY_NAMES, LIST_NAMES } from '../config/appConfig';

const KM_DATA_HUB_LIBRARY = LIBRARY_NAMES.kmDataHub;
const DOCUMENT_METRICS_LIST = LIST_NAMES.documentMetrics;
const USER_INTERACTIONS_LIST = LIST_NAMES.userInteractions;
const MAX_OR_FILTER_IDS = 30;

type TKMMetricName =
  | 'views'
  | 'likes'
  | 'comments'
  | 'downloads'
  | 'follow'
  | 'share'
  | 'bookmark';

type TUserInteractionType = 'Like' | 'Bookmark' | 'Follow' | 'View';
type TUserEventType = 'Download' | 'Share' | 'View';
type TMetricField =
  | 'ViewCount'
  | 'LikeCount'
  | 'DownloadCount'
  | 'CommentCount'
  | 'ShareCount'
  | 'BookmarkCount'
  | 'FollowCount';

interface IKMMetricFieldMap {
  views: string;
  likes: string;
  comments: string;
  downloads: string;
  follow: string;
  share: string;
  bookmark: string;
}

export interface IKMDocumentMetrics {
  views: number;
  likes: number;
  comments: number;
  downloads: number;
  follow: number;
  share: number;
  bookmark: number;
}

export interface IUserInteractionState {
  isLiked: boolean;
  isBookmarked: boolean;
  isFollowed: boolean;
  hasViewed: boolean;
}

interface IDocumentMetricsListItem {
  Id: number;
  DocumentId: number;
  ViewCount?: number;
  LikeCount?: number;
  DownloadCount?: number;
  CommentCount?: number;
  ShareCount?: number;
  BookmarkCount?: number;
  FollowCount?: number;
  LastUpdated?: string;
  __metadata?: {
    etag?: string;
  };
}

interface IUserInteractionListItem {
  Id: number;
  DocumentId: number;
  UserId: number;
  InteractionType: string;
  InteractionKey: string;
}

const METRIC_FIELD_TITLES: Record<TKMMetricName, string> = {
  views: COLUMN_NAMES.views,
  likes: COLUMN_NAMES.likes,
  comments: COLUMN_NAMES.comments,
  downloads: COLUMN_NAMES.downloads,
  follow: COLUMN_NAMES.follow,
  share: COLUMN_NAMES.share,
  bookmark: COLUMN_NAMES.bookmark
};

const METRIC_COUNT_FIELDS: Record<TKMMetricName, TMetricField> = {
  views: 'ViewCount',
  likes: 'LikeCount',
  comments: 'CommentCount',
  downloads: 'DownloadCount',
  follow: 'FollowCount',
  share: 'ShareCount',
  bookmark: 'BookmarkCount'
};

const INTERACTION_METRIC_FIELDS: Record<TUserInteractionType | TUserEventType, TMetricField> = {
  Like: 'LikeCount',
  Bookmark: 'BookmarkCount',
  Follow: 'FollowCount',
  View: 'ViewCount',
  Download: 'DownloadCount',
  Share: 'ShareCount'
};

const fieldMapCache: Record<string, Promise<IKMMetricFieldMap> | undefined> = {};
const listEntityTypeCache: Record<string, Promise<string | null> | undefined> = {};

const wait = (ms: number): Promise<void> => new Promise((resolve) => setTimeout(resolve, ms));

const normalizeValue = (value?: string): string => (value || '').trim().toLowerCase();

const sanitizeMetricValue = (value: number | null | undefined): number =>
  Number.isFinite(value) && value !== null && value !== undefined ? Math.max(0, Math.floor(value)) : 0;

const getODataItems = (data: any): any[] => {
  const items = data?.value || data?.d?.results || [];
  return Array.isArray(items) ? items : [];
};

const getODataNextLink = (data: any): string | null =>
  data?.['@odata.nextLink'] || data?.d?.__next || null;

const getODataProperty = <T,>(data: any, propertyName: string, fallback: T): T =>
  (data?.[propertyName] ?? data?.d?.[propertyName] ?? fallback) as T;

const escapeODataString = (value: string): string => (value || '').replace(/'/g, "''");

const chunkArray = <T,>(items: T[], size: number): T[][] => {
  const chunks: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }
  return chunks;
};

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

const getFreshDigest = async (context: WebPartContext): Promise<string> => {
  const webUrl = context.pageContext.web.absoluteUrl;
  const response = await requestWithThrottleRetry(() =>
    context.spHttpClient.post(
      `${webUrl}/_api/contextinfo`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: ''
      }
    )
  );

  if (!response.ok) {
    console.warn('[DIGEST] Failed to get request digest:', response.status);
    return '';
  }

  const data = await response.json();
  return data?.d?.GetContextWebInformation?.FormDigestValue ||
    data?.FormDigestValue ||
    data?.GetContextWebInformation?.FormDigestValue ||
    '';
};

const metricsFromItem = (item?: Partial<IDocumentMetricsListItem> | null): IKMDocumentMetrics => ({
  views: sanitizeMetricValue(item?.ViewCount),
  likes: sanitizeMetricValue(item?.LikeCount),
  comments: sanitizeMetricValue(item?.CommentCount),
  downloads: sanitizeMetricValue(item?.DownloadCount),
  follow: sanitizeMetricValue(item?.FollowCount),
  share: sanitizeMetricValue(item?.ShareCount),
  bookmark: sanitizeMetricValue(item?.BookmarkCount)
});

const getMetricsTotal = (item?: Partial<IDocumentMetricsListItem> | null): number => {
  const metrics = metricsFromItem(item);
  return metrics.views + metrics.likes + metrics.comments + metrics.downloads + metrics.follow + metrics.share + metrics.bookmark;
};

const getMetricTimestamp = (item?: Partial<IDocumentMetricsListItem> | null): number => {
  const timestamp = item?.LastUpdated ? new Date(item.LastUpdated).getTime() : 0;
  return Number.isFinite(timestamp) ? timestamp : 0;
};

const pickPreferredMetricsItem = (
  current: IDocumentMetricsListItem | undefined,
  candidate: IDocumentMetricsListItem
): IDocumentMetricsListItem => {
  if (!current) {
    return candidate;
  }

  const currentTime = getMetricTimestamp(current);
  const candidateTime = getMetricTimestamp(candidate);
  if (candidateTime !== currentTime) {
    return candidateTime > currentTime ? candidate : current;
  }

  return getMetricsTotal(candidate) > getMetricsTotal(current) ? candidate : current;
};

export const buildZeroKMDocumentMetrics = (): IKMDocumentMetrics => ({
  views: 0,
  likes: 0,
  comments: 0,
  downloads: 0,
  follow: 0,
  share: 0,
  bookmark: 0
});

const getListEntityType = async (
  context: WebPartContext,
  listName: string
): Promise<string | null> => {
  const cacheKey = `${context.pageContext.web.absoluteUrl}|${listName}`;
  if (!listEntityTypeCache[cacheKey]) {
    listEntityTypeCache[cacheKey] = (async () => {
      const webUrl = context.pageContext.web.absoluteUrl;
      const response = await requestWithThrottleRetry(() =>
        context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${escapeODataString(listName)}')?$select=ListItemEntityTypeFullName`,
          SPHttpClient.configurations.v1
        )
      );

      if (!response.ok) {
        return null;
      }

      const data = await response.json();
      return getODataProperty<string | null>(data, 'ListItemEntityTypeFullName', null);
    })();
  }

  return listEntityTypeCache[cacheKey]!;
};

const getKMMetricFieldMap = async (context: WebPartContext): Promise<IKMMetricFieldMap> => {
  const cacheKey = context.pageContext.web.absoluteUrl;
  if (!fieldMapCache[cacheKey]) {
    fieldMapCache[cacheKey] = (async () => {
      const webUrl = context.pageContext.web.absoluteUrl;
      const response = await requestWithThrottleRetry(() =>
        context.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${escapeODataString(KM_DATA_HUB_LIBRARY)}')/fields?$select=Title,InternalName&$top=5000`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata.metadata=minimal'
            }
          }
        )
      );

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to read KM Data Hub metric fields: ${response.status} ${errorText}`);
      }

      const json = await response.json();
      const fields = getODataItems(json) as Array<{ Title?: string; InternalName?: string }>;

      const findInternalName = (displayName: string): string => {
        const match = fields.find((field) =>
          normalizeValue(field.Title) === normalizeValue(displayName) ||
          normalizeValue(field.InternalName) === normalizeValue(displayName)
        );

        return match?.InternalName || displayName;
      };

      return {
        views: findInternalName(METRIC_FIELD_TITLES.views),
        likes: findInternalName(METRIC_FIELD_TITLES.likes),
        comments: findInternalName(METRIC_FIELD_TITLES.comments),
        downloads: findInternalName(METRIC_FIELD_TITLES.downloads),
        follow: findInternalName(METRIC_FIELD_TITLES.follow),
        share: findInternalName(METRIC_FIELD_TITLES.share),
        bookmark: findInternalName(METRIC_FIELD_TITLES.bookmark)
      };
    })();
  }

  return fieldMapCache[cacheKey]!;
};

const findDocumentMetricsItem = async (
  context: WebPartContext,
  documentId: number,
  includeEtag: boolean = false
): Promise<IDocumentMetricsListItem | null> => {
  if (!documentId || documentId <= 0) {
    return null;
  }

  const webUrl = context.pageContext.web.absoluteUrl;
  const response = await requestWithThrottleRetry(() =>
    context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items` +
      `?$select=Id,DocumentId,ViewCount,LikeCount,DownloadCount,CommentCount,ShareCount,BookmarkCount,FollowCount,LastUpdated` +
      `&$filter=DocumentId eq ${documentId}&$top=1`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: includeEtag ? 'application/json;odata=verbose' : 'application/json;odata=nometadata'
        }
      }
    )
  );

  if (!response.ok) {
    return null;
  }

  const data = await response.json();
  return (getODataItems(data)[0] as IDocumentMetricsListItem | undefined) || null;
};

export const createDocumentMetricsRow = async (
  context: WebPartContext,
  documentId: number,
  initialMetrics: Partial<IKMDocumentMetrics> = {},
  documentTitle?: string
): Promise<boolean> => {
  if (!documentId || documentId <= 0) {
    return false;
  }

  const existing = await findDocumentMetricsItem(context, documentId);
  if (existing?.Id) {
    return true;
  }

  const entityType = await getListEntityType(context, DOCUMENT_METRICS_LIST);
  if (!entityType) {
    return false;
  }

  const webUrl = context.pageContext.web.absoluteUrl;
  const response = await requestWithThrottleRetry(() =>
    context.spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: JSON.stringify({
          __metadata: { type: entityType },
          Title: documentTitle || String(documentId),
          DocumentId: documentId,
          ViewCount: sanitizeMetricValue(initialMetrics.views),
          LikeCount: sanitizeMetricValue(initialMetrics.likes),
          DownloadCount: sanitizeMetricValue(initialMetrics.downloads),
          CommentCount: sanitizeMetricValue(initialMetrics.comments),
          ShareCount: sanitizeMetricValue(initialMetrics.share),
          BookmarkCount: sanitizeMetricValue(initialMetrics.bookmark),
          FollowCount: sanitizeMetricValue(initialMetrics.follow),
          LastUpdated: new Date().toISOString()
        })
      }
    )
  );

  return response.ok;
};

export const getDocumentMetrics = async (
  context: WebPartContext,
  documentId: number
): Promise<IKMDocumentMetrics> => {
  const item = await findDocumentMetricsItem(context, documentId);
  return metricsFromItem(item);
};

export const getDocumentMetricsBatch = async (
  context: WebPartContext,
  documentIds: number[]
): Promise<Record<number, IKMDocumentMetrics>> => {
  const uniqueIds = Array.from(new Set(documentIds.filter((id) => Number.isFinite(id) && id > 0)));
  const metricsByDocumentId = uniqueIds.reduce((accumulator, documentId) => {
    accumulator[documentId] = buildZeroKMDocumentMetrics();
    return accumulator;
  }, {} as Record<number, IKMDocumentMetrics>);

  if (uniqueIds.length === 0) {
    return metricsByDocumentId;
  }

  const webUrl = context.pageContext.web.absoluteUrl;

  for (const chunk of chunkArray(uniqueIds, MAX_OR_FILTER_IDS)) {
    const filter = chunk.map((id) => `DocumentId eq ${id}`).join(' or ');
    const response = await requestWithThrottleRetry(() =>
      context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items` +
        `?$select=Id,DocumentId,ViewCount,LikeCount,DownloadCount,CommentCount,ShareCount,BookmarkCount,FollowCount,LastUpdated` +
        `&$filter=${encodeURIComponent(filter)}&$orderby=LastUpdated desc,Id desc&$top=5000`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      )
    );

    if (!response.ok) {
      continue;
    }

    const data = await response.json();
    const preferredItemsByDocumentId: Record<number, IDocumentMetricsListItem> = {};
    (getODataItems(data) as IDocumentMetricsListItem[]).forEach((item) => {
      const documentId = Number(item.DocumentId || 0);
      if (documentId > 0) {
        preferredItemsByDocumentId[documentId] = pickPreferredMetricsItem(preferredItemsByDocumentId[documentId], item);
      }
    });

    Object.keys(preferredItemsByDocumentId).forEach((documentIdKey) => {
      const documentId = Number(documentIdKey);
      metricsByDocumentId[documentId] = metricsFromItem(preferredItemsByDocumentId[documentId]);
    });
  }

  return metricsByDocumentId;
};

const buildInteractionKey = (documentId: number, userId: number, type: string): string =>
  `${documentId}_${userId}_${type}`;

const buildEventInteractionKey = (documentId: number, userId: number, type: string): string =>
  `${documentId}_${userId}_${type}_${Date.now()}_${Math.floor(Math.random() * 100000)}`;

const getInteractionByKey = async (
  context: WebPartContext,
  interactionKey: string
): Promise<IUserInteractionListItem | null> => {
  const webUrl = context.pageContext.web.absoluteUrl;
  console.log('[TOGGLE] InteractionKey:', interactionKey);
  const response = await requestWithThrottleRetry(() =>
    context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items` +
      `?$select=Id,DocumentId,UserId,InteractionType,InteractionKey,Title` +
      `&$filter=InteractionKey eq '${escapeODataString(interactionKey)}'&$top=1`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    )
  );

  if (!response.ok) {
    return null;
  }

  const data = await response.json();
  const existing = (getODataItems(data)[0] as IUserInteractionListItem | undefined) || null;
  console.log('[TOGGLE] Existing row found:', existing?.Id || 'none');
  return existing;
};

const createUserInteraction = async (
  context: WebPartContext,
  documentId: number,
  userId: number,
  interactionType: string,
  interactionKey: string,
  documentTitle?: string
): Promise<boolean> => {
  const entityType = await getListEntityType(context, USER_INTERACTIONS_LIST);
  if (!entityType) {
    return false;
  }

  const webUrl = context.pageContext.web.absoluteUrl;
  const digest = await getFreshDigest(context);
  const response = await requestWithThrottleRetry(() =>
    context.spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          ...(digest ? { 'X-RequestDigest': digest } : {}),
          'odata-version': ''
        },
        body: JSON.stringify({
          __metadata: { type: entityType },
          Title: documentTitle || String(documentId),
          DocumentId: documentId,
          UserId: userId,
          InteractionType: interactionType,
          InteractionKey: interactionKey
        })
      }
    )
  );

  return response.ok;
};

const deleteUserInteraction = async (
  context: WebPartContext,
  itemId: number
): Promise<boolean> => {
  const webUrl = context.pageContext.web.absoluteUrl;
  const digest = await getFreshDigest(context);
  const response = await requestWithThrottleRetry(() =>
    context.spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE',
          ...(digest ? { 'X-RequestDigest': digest } : {}),
          'odata-version': ''
        }
      }
    )
  );

  console.log('[TOGGLE] DELETE status:', response.status);
  return response.ok;
};

export const getUserInteractionState = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<IUserInteractionState> => {
  if (!documentId || !userId) {
    return { isLiked: false, isBookmarked: false, isFollowed: false, hasViewed: false };
  }

  const webUrl = context.pageContext.web.absoluteUrl;
  const response = await requestWithThrottleRetry(() =>
    context.spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items` +
      `?$select=InteractionType&$filter=DocumentId eq ${documentId} and UserId eq ${userId}&$top=100`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    )
  );

  if (!response.ok) {
    return { isLiked: false, isBookmarked: false, isFollowed: false, hasViewed: false };
  }

  const data = await response.json();
  const types = new Set(getODataItems(data).map((item) => String(item.InteractionType || '')));

  return {
    isLiked: types.has('Like'),
    isBookmarked: types.has('Bookmark'),
    isFollowed: types.has('Follow'),
    hasViewed: types.has('View')
  };
};

export const getUserInteractionsBatch = async (
  context: WebPartContext,
  documentIds: number[],
  userId: number
): Promise<Record<number, IUserInteractionState>> => {
  const uniqueIds = Array.from(new Set(documentIds.filter((id) => Number.isFinite(id) && id > 0)));
  const stateByDocumentId = uniqueIds.reduce((accumulator, documentId) => {
    accumulator[documentId] = { isLiked: false, isBookmarked: false, isFollowed: false, hasViewed: false };
    return accumulator;
  }, {} as Record<number, IUserInteractionState>);

  if (uniqueIds.length === 0 || userId <= 0) {
    return stateByDocumentId;
  }

  const webUrl = context.pageContext.web.absoluteUrl;

  for (const chunk of chunkArray(uniqueIds, MAX_OR_FILTER_IDS)) {
    const filter = `UserId eq ${userId} and (${chunk.map((id) => `DocumentId eq ${id}`).join(' or ')})`;
    const response = await requestWithThrottleRetry(() =>
      context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items` +
        `?$select=DocumentId,InteractionType&$filter=${encodeURIComponent(filter)}&$top=5000`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      )
    );

    if (!response.ok) {
      continue;
    }

    const data = await response.json();
    getODataItems(data).forEach((item) => {
      const documentId = Number(item.DocumentId || 0);
      const interactionType = String(item.InteractionType || '');
      const state = stateByDocumentId[documentId];
      if (!state) {
        return;
      }

      if (interactionType === 'Like') state.isLiked = true;
      if (interactionType === 'Bookmark') state.isBookmarked = true;
      if (interactionType === 'Follow') state.isFollowed = true;
      if (interactionType === 'View') state.hasViewed = true;
    });
  }

  return stateByDocumentId;
};

export const incrementMetricWithRetry = async (
  context: WebPartContext,
  documentId: number,
  field: TMetricField,
  delta: number,
  maxRetries: number = 3
): Promise<number> => {
  let lastKnownValue = 0;
  console.log('[INCREMENT] Starting:', field, delta, documentId);

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      console.log('[INCREMENT] GET DocumentMetrics row...');
      const normalizedDocumentId = parseInt(String(documentId), 10);
      const item = await findDocumentMetricsItem(context, normalizedDocumentId, false);
      if (!item?.Id) {
        console.error('[INCREMENT] DocumentMetrics row not found for:', documentId);
        await createDocumentMetricsRow(context, normalizedDocumentId);
        if (attempt < maxRetries) {
          continue;
        }
        return lastKnownValue;
      }

      const currentValue = sanitizeMetricValue((item as any)[field]);
      const nextValue = Math.max(0, currentValue + delta);
      lastKnownValue = nextValue;
      console.log('[INCREMENT] Current value:', currentValue);
      console.log('[INCREMENT] List item Id:', item.Id);

      const webUrl = context.pageContext.web.absoluteUrl;
      const [entityType, digest] = await Promise.all([
        getListEntityType(context, DOCUMENT_METRICS_LIST),
        getFreshDigest(context)
      ]);
      const patchUrl = `${webUrl}/_api/web/lists/getbytitle('${DOCUMENT_METRICS_LIST}')/items(${item.Id})`;
      console.log(`[INCREMENT] ${field}: ${currentValue} -> ${nextValue}`);
      console.log('[INCREMENT] PATCH URL:', patchUrl);

      const response = await requestWithThrottleRetry(() =>
        context.spHttpClient.post(
          patchUrl,
          SPHttpClient.configurations.v1,
          {
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE',
            ...(digest ? { 'X-RequestDigest': digest } : {}),
            'odata-version': ''
            },
            body: JSON.stringify({
              ...(entityType ? { __metadata: { type: entityType } } : {}),
              [field]: nextValue,
              PendingSync: true,
              LastUpdated: new Date().toISOString()
            })
          }
        )
      );

      console.log('[INCREMENT] PATCH result:', response.status);
      if (response.ok || response.status === 204) {
        console.log('[INCREMENT] Success:', field, nextValue);
        return nextValue;
      }

      if (response.status === 412 && attempt < maxRetries) {
        console.log('[INCREMENT] Conflict, retrying...');
        await wait(100 * (attempt + 1));
        continue;
      }

      const errorText = await response.text();
      console.error('[INCREMENT] PATCH failed:', response.status, errorText);
      if (!isThrottleResponse(response)) {
        break;
      }
    } catch (error) {
      console.error('[INCREMENT] Failed:', error);
      if (attempt >= maxRetries) {
        break;
      }
      await wait(200 * (attempt + 1));
    }
  }

  return lastKnownValue;
};

export const toggleUserInteraction = async (
  context: WebPartContext,
  documentId: number,
  userId: number,
  type: TUserInteractionType,
  documentTitle?: string
): Promise<{ action: 'added' | 'removed' | 'unchanged'; newCount: number; isActive: boolean }> => {
  if (!documentId || !userId) {
    return { action: 'unchanged', newCount: 0, isActive: false };
  }

  const interactionKey = buildInteractionKey(documentId, userId, type);
  console.log('[TOGGLE] documentId:', documentId);
  console.log('[TOGGLE] userId:', userId);
  console.log('[TOGGLE] type:', type);
  const existing = await getInteractionByKey(context, interactionKey);
  const field = INTERACTION_METRIC_FIELDS[type];

  if (existing?.Id) {
    const deleted = await deleteUserInteraction(context, existing.Id);
    const newCount = deleted
      ? await incrementMetricWithRetry(context, documentId, field, -1)
      : (await getDocumentMetrics(context, documentId))[metricFieldToMetricName(field)];

    return { action: deleted ? 'removed' : 'unchanged', newCount, isActive: false };
  }

  const created = await createUserInteraction(context, documentId, userId, type, interactionKey, documentTitle);
  if (!created) {
    const metrics = await getDocumentMetrics(context, documentId);
    return { action: 'unchanged', newCount: metrics[metricFieldToMetricName(field)], isActive: false };
  }

  const newCount = await incrementMetricWithRetry(context, documentId, field, 1);
  return { action: 'added', newCount, isActive: true };
};

const metricFieldToMetricName = (field: TMetricField): keyof IKMDocumentMetrics => {
  switch (field) {
    case 'ViewCount': return 'views';
    case 'LikeCount': return 'likes';
    case 'DownloadCount': return 'downloads';
    case 'CommentCount': return 'comments';
    case 'ShareCount': return 'share';
    case 'BookmarkCount': return 'bookmark';
    case 'FollowCount': return 'follow';
  }
};

export const recordUserEvent = async (
  context: WebPartContext,
  documentId: number,
  userId: number,
  type: TUserEventType,
  documentTitle?: string
): Promise<number> => {
  if (!documentId || !userId) {
    return 0;
  }

  const field = INTERACTION_METRIC_FIELDS[type];

  if (type === 'View') {
    const interactionKey = buildInteractionKey(documentId, userId, type);
    const existing = await getInteractionByKey(context, interactionKey);
    if (existing?.Id) {
      const metrics = await getDocumentMetrics(context, documentId);
      return metrics.views;
    }

    const created = await createUserInteraction(context, documentId, userId, type, interactionKey, documentTitle);
    return created ? incrementMetricWithRetry(context, documentId, field, 1) : (await getDocumentMetrics(context, documentId)).views;
  }

  const interactionKey = buildEventInteractionKey(documentId, userId, type);
  await createUserInteraction(context, documentId, userId, type, interactionKey, documentTitle);
  return incrementMetricWithRetry(context, documentId, field, 1);
};

const validateMetricUpdate = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number,
  fieldName: string,
  fieldValue: number
): Promise<void> => {
  const response = await requestWithThrottleRetry(() =>
    spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${KM_DATA_HUB_LIBRARY}')/items(${itemId})/validateUpdateListItem`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose',
          'X-Requested-With': 'XMLHttpRequest',
          'odata-version': ''
        },
        body: JSON.stringify({
          formValues: [{
            FieldName: fieldName,
            FieldValue: String(fieldValue)
          }],
          bNewDocumentUpdate: false,
          checkInComment: ''
        })
      }
    )
  );

  if (!response.ok) {
    const err = await response.text();
    console.warn(`validateMetricUpdate failed - ${fieldName} on item ${itemId}: ${response.status}`, err);
  }
};

export const syncKMDataHubMetricValues = async (
  context: WebPartContext,
  itemId: number,
  metrics: IKMDocumentMetrics
): Promise<void> => {
  const fieldMap = await getKMMetricFieldMap(context);
  const webUrl = context.pageContext.web.absoluteUrl;
  await Promise.all([
    validateMetricUpdate(context.spHttpClient, webUrl, itemId, fieldMap.views || COLUMN_NAMES.views, sanitizeMetricValue(metrics.views)),
    validateMetricUpdate(context.spHttpClient, webUrl, itemId, fieldMap.likes || COLUMN_NAMES.likes, sanitizeMetricValue(metrics.likes)),
    validateMetricUpdate(context.spHttpClient, webUrl, itemId, fieldMap.comments || COLUMN_NAMES.comments, sanitizeMetricValue(metrics.comments)),
    validateMetricUpdate(context.spHttpClient, webUrl, itemId, fieldMap.downloads || COLUMN_NAMES.downloads, sanitizeMetricValue(metrics.downloads)),
    validateMetricUpdate(context.spHttpClient, webUrl, itemId, fieldMap.follow || COLUMN_NAMES.follow, sanitizeMetricValue(metrics.follow)),
    validateMetricUpdate(context.spHttpClient, webUrl, itemId, fieldMap.share || COLUMN_NAMES.share, sanitizeMetricValue(metrics.share)),
    validateMetricUpdate(context.spHttpClient, webUrl, itemId, fieldMap.bookmark || COLUMN_NAMES.bookmark, sanitizeMetricValue(metrics.bookmark))
  ]);
};

export const syncDocumentMetricsToKMHub = async (
  context: WebPartContext,
  documentIds: number[]
): Promise<void> => {
  const metricsById = await getDocumentMetricsBatch(context, documentIds);

  for (const chunk of chunkArray(Object.keys(metricsById), 50)) {
    await Promise.all(
      chunk.map(async (itemId) => {
        const documentId = Number(itemId);
        try {
          await syncKMDataHubMetricValues(context, documentId, metricsById[documentId]);
        } catch {
          // Keep bulk sync resilient per item.
        }
      })
    );
  }
};

export const fetchKMDocumentMetricsByIds = async (
  context: WebPartContext,
  documentIds: number[]
): Promise<Record<number, IKMDocumentMetrics>> => {
  return getDocumentMetricsBatch(context, documentIds);
};

export const syncKMDataHubMetricValuesForItems = async (
  context: WebPartContext,
  itemIds: number[]
): Promise<void> => {
  return syncDocumentMetricsToKMHub(context, itemIds);
};

export const getCurrentUserId = async (
  context: WebPartContext
): Promise<number> => {
  try {
    const webUrl = context.pageContext.web.absoluteUrl;
    const resp = await context.spHttpClient.get(
      `${webUrl}/_api/web/currentuser?$select=Id`,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return 0;
    const data = await resp.json();
    return data?.Id || data?.d?.Id || 0;
  } catch {
    return 0;
  }
};

export const recordDocumentView = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<number> => recordUserEvent(context, documentId, userId, 'View');

export const toggleDocumentLike = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<{ liked: boolean; count: number }> => {
  const result = await toggleUserInteraction(context, documentId, userId, 'Like');
  return { liked: result.isActive, count: result.newCount };
};

export const toggleDocumentFollow = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<{ followed: boolean; count: number }> => {
  const result = await toggleUserInteraction(context, documentId, userId, 'Follow');
  return { followed: result.isActive, count: result.newCount };
};

export const toggleDocumentBookmark = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<{ bookmarked: boolean; count: number }> => {
  const result = await toggleUserInteraction(context, documentId, userId, 'Bookmark');
  return { bookmarked: result.isActive, count: result.newCount };
};

export const recordDocumentDownload = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<number> => recordUserEvent(context, documentId, userId, 'Download');

export const recordDocumentShare = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<number> => recordUserEvent(context, documentId, userId, 'Share');

export const checkUserLiked = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<boolean> => (await getUserInteractionState(context, documentId, userId)).isLiked;

export const checkUserFollowed = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<boolean> => (await getUserInteractionState(context, documentId, userId)).isFollowed;

export const checkUserBookmarked = async (
  context: WebPartContext,
  documentId: number,
  userId: number
): Promise<boolean> => (await getUserInteractionState(context, documentId, userId)).isBookmarked;

export const getUserBookmarkDocumentIds = async (
  context: WebPartContext,
  userId: number
): Promise<number[]> => {
  if (!userId) {
    return [];
  }

  const webUrl = context.pageContext.web.absoluteUrl;
  const bookmarkIds = new Set<number>();
  let nextUrl: string | null =
    `${webUrl}/_api/web/lists/getbytitle('${USER_INTERACTIONS_LIST}')/items` +
    `?$select=DocumentId&$filter=UserId eq ${userId} and InteractionType eq 'Bookmark'&$top=5000`;

  while (nextUrl) {
    const response = await requestWithThrottleRetry(() =>
      context.spHttpClient.get(nextUrl!, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      })
    );

    if (!response.ok) {
      break;
    }

    const data = await response.json();
    getODataItems(data).forEach((item) => {
      const documentId = Number(item.DocumentId || 0);
      if (documentId > 0) {
        bookmarkIds.add(documentId);
      }
    });
    nextUrl = getODataNextLink(data);
  }

  return Array.from(bookmarkIds);
};
