import { requireValue, SearchEnv } from './env';

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';

export interface GraphPage<T = any> {
  value?: T[];
  '@odata.nextLink'?: string;
  '@odata.deltaLink'?: string;
}

const parseRetryAfterMs = (headers: Headers): number | undefined => {
  const retryAfter = headers.get('retry-after');
  if (!retryAfter) return undefined;
  const seconds = Number(retryAfter);
  if (Number.isFinite(seconds) && seconds > 0) return seconds * 1000;
  const retryDate = Date.parse(retryAfter);
  if (Number.isFinite(retryDate)) return Math.max(0, retryDate - Date.now());
  return undefined;
};

const fetchWithTimeout = async (url: string, init: RequestInit, timeoutMs: number): Promise<Response> => {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...init, signal: controller.signal });
  } finally {
    clearTimeout(timeout);
  }
};

export const getGraphToken = async (env: SearchEnv): Promise<string> => {
  const tenantId = requireValue(env, 'GRAPH_TENANT_ID');
  const params = new URLSearchParams();
  params.set('client_id', requireValue(env, 'GRAPH_CLIENT_ID'));
  params.set('client_secret', requireValue(env, 'GRAPH_CLIENT_SECRET'));
  params.set('grant_type', 'client_credentials');
  params.set('scope', env.GRAPH_SCOPE || 'https://graph.microsoft.com/.default');

  const response = await fetchWithTimeout(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params
  }, Number(env.GRAPH_TOKEN_TIMEOUT_MS || 30000));

  if (!response.ok) {
    const body = await response.text();
    if (body.includes('AADSTS7000215')) {
      throw new Error([
        'Graph token request failed: Microsoft Entra rejected GRAPH_CLIENT_SECRET.',
        'Use the client secret VALUE, not the Secret ID. If the value is hidden, create a new client secret.'
      ].join('\n'));
    }
    throw new Error(`Graph token request failed ${response.status}: ${body.slice(0, 700)}`);
  }

  const json = await response.json();
  return json.access_token;
};

export const graphFetch = async <T = any>(
  token: string,
  url: string,
  responseType: 'json' | 'arrayBuffer' = 'json'
): Promise<T> => {
  const response = await fetchWithTimeout(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: responseType === 'json' ? 'application/json' : '*/*'
    }
  }, 120000);

  if (!response.ok) {
    const body = await response.text();
    const error = new Error(`Graph GET failed ${response.status}: ${body.slice(0, 700)}`) as Error & {
      retryAfterMs?: number;
      statusCode?: number;
    };
    error.retryAfterMs = parseRetryAfterMs(response.headers);
    error.statusCode = response.status;
    throw error;
  }

  if (responseType === 'arrayBuffer') {
    return response.arrayBuffer() as unknown as Promise<T>;
  }

  return response.json() as Promise<T>;
};

export const fetchLibraryItemsPage = async (
  env: SearchEnv,
  token: string,
  top = 10,
  nextLink?: string
): Promise<GraphPage> => {
  const siteId = encodeURIComponent(requireValue(env, 'SHAREPOINT_SITE_ID'));
  const listId = encodeURIComponent(requireValue(env, 'SHAREPOINT_LIBRARY_LIST_ID'));
  const url = nextLink || `${graphBaseUrl}/sites/${siteId}/lists/${listId}/items?$expand=fields,driveItem&$top=${top}`;
  return graphFetch(token, url);
};

export const fetchDriveDeltaPage = async (
  env: SearchEnv,
  token: string,
  nextLinkOrDeltaLink?: string
): Promise<GraphPage> => {
  const driveId = encodeURIComponent(requireValue(env, 'SHAREPOINT_LIBRARY_DRIVE_ID'));
  const url = nextLinkOrDeltaLink || `${graphBaseUrl}/drives/${driveId}/root/delta`;
  return graphFetch(token, url);
};

export const fetchListItemByDriveItemId = async (
  env: SearchEnv,
  token: string,
  driveItemId: string
): Promise<any | null> => {
  const driveId = encodeURIComponent(requireValue(env, 'SHAREPOINT_LIBRARY_DRIVE_ID'));
  const itemId = encodeURIComponent(driveItemId);
  const driveItem = await graphFetch<any>(token, `${graphBaseUrl}/drives/${driveId}/items/${itemId}?$expand=listItem($expand=fields)`);
  if (!driveItem.listItem) return null;
  return {
    ...driveItem.listItem,
    driveItem
  };
};

export const fetchPeopleItems = async (
  env: SearchEnv,
  token: string,
  top = 10
): Promise<any[]> => {
  const siteId = encodeURIComponent(requireValue(env, 'SHAREPOINT_SITE_ID'));
  const listId = encodeURIComponent(requireValue(env, 'SHAREPOINT_WHOSWHO_LIST_ID'));
  const url = `${graphBaseUrl}/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=${top}`;
  const json = await graphFetch<GraphPage>(token, url);
  return json.value || [];
};

export const downloadDriveItem = async (
  env: SearchEnv,
  token: string,
  driveItemId: string
): Promise<Buffer> => {
  const driveId = encodeURIComponent(requireValue(env, 'SHAREPOINT_LIBRARY_DRIVE_ID'));
  const itemId = encodeURIComponent(driveItemId);
  const url = `${graphBaseUrl}/drives/${driveId}/items/${itemId}/content`;
  const arrayBuffer = await graphFetch<ArrayBuffer>(token, url, 'arrayBuffer');
  return Buffer.from(arrayBuffer);
};
