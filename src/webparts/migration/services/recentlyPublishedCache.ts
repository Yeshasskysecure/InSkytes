import { SPHttpClient } from '@microsoft/sp-http';
import { LIST_NAMES } from '../config/appConfig';

const CACHE_TITLE = 'recentlyPublishedCache';
const CACHE_LIMIT = 8;

interface IRecentlyPublishedCachePayload {
  ids: number[];
  savedAt: string;
}

const normalizeIds = (ids: number[]): number[] =>
  Array.from(new Set(ids.map((id) => Number(id)).filter((id) => Number.isFinite(id) && id > 0)))
    .slice(0, CACHE_LIMIT);

const parseCachePayload = (value: string): IRecentlyPublishedCachePayload | null => {
  try {
    const parsed = JSON.parse(value || '{}');
    return {
      ids: normalizeIds(Array.isArray(parsed.ids) ? parsed.ids : []),
      savedAt: String(parsed.savedAt || '')
    };
  } catch {
    return null;
  }
};

export const fetchRecentlyPublishedCacheIds = async (
  spHttpClient: SPHttpClient,
  webUrl: string
): Promise<number[]> => {
  try {
    const response = await spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIST_NAMES.homePageConfig}')/items` +
      `?$select=Id,Title,HeroMessage&$filter=Title eq '${CACHE_TITLE}'&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      return [];
    }

    const data = await response.json();
    const item = (data.value || [])[0];
    const payload = parseCachePayload(item?.HeroMessage || '');
    return payload?.ids || [];
  } catch {
    return [];
  }
};

export const saveRecentlyPublishedCacheIds = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  ids: number[]
): Promise<void> => {
  const normalizedIds = normalizeIds(ids);
  if (normalizedIds.length === 0) {
    return;
  }

  const payload = JSON.stringify({
    ids: normalizedIds,
    savedAt: new Date().toISOString()
  });

  try {
    const lookupResponse = await spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${LIST_NAMES.homePageConfig}')/items` +
      `?$select=Id&$filter=Title eq '${CACHE_TITLE}'&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!lookupResponse.ok) {
      return;
    }

    const lookupData = await lookupResponse.json();
    const existingItem = (lookupData.value || [])[0];

    if (existingItem?.Id) {
      await spHttpClient.post(
        `${webUrl}/_api/web/lists/getbytitle('${LIST_NAMES.homePageConfig}')/items(${existingItem.Id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify({ HeroMessage: payload })
        }
      );
      return;
    }

    await spHttpClient.post(
      `${webUrl}/_api/web/lists/getbytitle('${LIST_NAMES.homePageConfig}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({
          Title: CACHE_TITLE,
          HeroMessage: payload,
          QuoteAuthor: '',
          ImageUrl: ''
        })
      }
    );
  } catch {
    // Cache write is best-effort. The live query can still render the section.
  }
};

export const touchRecentlyPublishedCacheForDocument = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  documentId: number,
  isActive: boolean
): Promise<void> => {
  const currentIds = await fetchRecentlyPublishedCacheIds(spHttpClient, webUrl);
  const nextIds = isActive
    ? normalizeIds([documentId, ...currentIds])
    : normalizeIds(currentIds.filter((id) => id !== documentId));

  await saveRecentlyPublishedCacheIds(spHttpClient, webUrl, nextIds);
};
