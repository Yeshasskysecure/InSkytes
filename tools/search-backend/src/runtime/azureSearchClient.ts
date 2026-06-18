import { requireValue, SearchEnv } from './env';

export interface AzureSearchDocument {
  [key: string]: unknown;
}

export interface AzureSearchResponse {
  value?: AzureSearchDocument[];
  [key: string]: unknown;
}

interface AzureSearchIndexActionResult {
  key?: string;
  status?: boolean;
  errorMessage?: string | null;
  statusCode?: number;
}

interface AzureSearchIndexResponse {
  value?: AzureSearchIndexActionResult[];
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

export class AzureSearchIndexingError extends Error {
  public readonly indexName: string;
  public readonly failures: AzureSearchIndexActionResult[];

  constructor(indexName: string, failures: AzureSearchIndexActionResult[]) {
    const preview = failures
      .slice(0, 5)
      .map((failure) => `${failure.key || '(missing key)'}: ${failure.statusCode || 'unknown'} ${failure.errorMessage || 'Unknown indexing error'}`)
      .join('; ');
    super(`Azure Search indexing partially failed for ${indexName}: ${preview}`);
    this.name = 'AzureSearchIndexingError';
    this.indexName = indexName;
    this.failures = failures;
  }
}

export const azureSearchFetch = async <T = AzureSearchResponse>(
  env: SearchEnv,
  method: string,
  pathPart: string,
  body?: unknown
): Promise<T> => {
  const endpoint = requireValue(env, 'AZURE_SEARCH_ENDPOINT').replace(/\/$/, '');
  const apiVersion = requireValue(env, 'AZURE_SEARCH_API_VERSION');
  const adminKey = requireValue(env, 'AZURE_SEARCH_ADMIN_KEY');
  const separator = pathPart.includes('?') ? '&' : '?';
  const response = await fetch(`${endpoint}${pathPart}${separator}api-version=${apiVersion}`, {
    method,
    headers: {
      'Content-Type': 'application/json',
      'api-key': adminKey
    },
    body: body ? JSON.stringify(body) : undefined
  });

  if (!response.ok) {
    const text = await response.text();
    const error = new Error(`Azure Search ${method} ${pathPart} failed ${response.status}: ${text.slice(0, 1000)}`) as Error & {
      retryAfterMs?: number;
      statusCode?: number;
    };
    error.retryAfterMs = parseRetryAfterMs(response.headers);
    error.statusCode = response.status;
    throw error;
  }

  if (response.status === 204) return {} as T;
  const text = await response.text();
  return text ? JSON.parse(text) as T : {} as T;
};

const assertIndexingSucceeded = (indexName: string, response: AzureSearchIndexResponse): void => {
  const failures = (response.value || []).filter((item) => item.status === false);
  if (failures.length > 0) {
    throw new AzureSearchIndexingError(indexName, failures);
  }
};

export const uploadDocuments = async (
  env: SearchEnv,
  indexName: string,
  records: AzureSearchDocument[]
): Promise<void> => {
  if (records.length === 0) return;
  const configuredBatchSize = Number(env.AZURE_SEARCH_UPLOAD_BATCH_SIZE || 1000);
  const batchSize = Math.min(Math.max(Number.isFinite(configuredBatchSize) ? configuredBatchSize : 1000, 1), 1000);
  for (let index = 0; index < records.length; index += batchSize) {
    const batch = records.slice(index, index + batchSize);
    const response = await azureSearchFetch<AzureSearchIndexResponse>(env, 'POST', `/indexes/${encodeURIComponent(indexName)}/docs/index`, {
      value: batch.map((record) => ({ '@search.action': 'mergeOrUpload', ...record }))
    });
    assertIndexingSucceeded(indexName, response);
  }
};

export const deleteDocuments = async (
  env: SearchEnv,
  indexName: string,
  keyName: string,
  keys: string[]
): Promise<void> => {
  if (keys.length === 0) return;
  const configuredBatchSize = Number(env.AZURE_SEARCH_UPLOAD_BATCH_SIZE || 1000);
  const batchSize = Math.min(Math.max(Number.isFinite(configuredBatchSize) ? configuredBatchSize : 1000, 1), 1000);
  for (let index = 0; index < keys.length; index += batchSize) {
    const batch = keys.slice(index, index + batchSize);
    const response = await azureSearchFetch<AzureSearchIndexResponse>(env, 'POST', `/indexes/${encodeURIComponent(indexName)}/docs/index`, {
      value: batch.map((key) => ({ '@search.action': 'delete', [keyName]: key }))
    });
    assertIndexingSucceeded(indexName, response);
  }
};

export const searchIndex = async (
  env: SearchEnv,
  indexName: string,
  body: unknown
): Promise<AzureSearchResponse> =>
  azureSearchFetch(env, 'POST', `/indexes/${encodeURIComponent(indexName)}/docs/search`, body);
