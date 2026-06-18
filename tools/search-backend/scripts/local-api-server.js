const http = require('http');
const https = require('https');
const os = require('os');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { loadEnv } = require('../dist/runtime/env');
const {
  chat,
  searchDocuments,
  searchDocumentsPaged,
  searchPeople,
  semanticSearch
} = require('../dist/api/knowledgeSearchApi');
const {
  analyzeMediaTranscript,
  extractUploadMetadata,
  transcribeUploadedMedia
} = require('../dist/api/uploadAiApi');

const envPath = process.argv[2] || process.env.SEARCH_ENV_PATH || 'config/search.env';
const port = Number(process.argv[3] || process.env.PORT || 7072);
const { env, absolutePath } = loadEnv(envPath);
if (!process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED && env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED) {
  process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED = env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED;
}
if (!process.env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED && env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED) {
  process.env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED = env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED;
}
const normalizedFacetsEnabled = /^(1|true|yes)$/i.test(String(process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED || ''));
const canonicalDepartmentFacetsEnabled = /^(1|true|yes)$/i.test(String(process.env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED || ''));
const searchAnalyticsEnabled = /^(1|true|yes)$/i.test(String(env.SEARCH_ANALYTICS_ENABLED || process.env.SEARCH_ANALYTICS_ENABLED || 'false'));
const searchAnalyticsSink = String(env.SEARCH_ANALYTICS_SINK || process.env.SEARCH_ANALYTICS_SINK || 'eventhub').toLowerCase();
const searchAnalyticsEnvironment = String(env.SEARCH_ANALYTICS_ENVIRONMENT || process.env.SEARCH_ANALYTICS_ENVIRONMENT || env.NODE_ENV || process.env.NODE_ENV || 'prod');
const searchAnalyticsBackendVersion = String(env.SEARCH_ANALYTICS_BACKEND_VERSION || process.env.SEARCH_ANALYTICS_BACKEND_VERSION || env.WEBSITE_SITE_NAME || process.env.WEBSITE_SITE_NAME || 'iknowledge-search-api');
const searchAnalyticsStoreRawQuery = /^(1|true|yes)$/i.test(String(env.SEARCH_ANALYTICS_STORE_RAW_QUERY || process.env.SEARCH_ANALYTICS_STORE_RAW_QUERY || 'false'));
const searchAnalyticsStoreSanitizedQuery = !/^(0|false|no)$/i.test(String(env.SEARCH_ANALYTICS_STORE_SANITIZED_QUERY || process.env.SEARCH_ANALYTICS_STORE_SANITIZED_QUERY || 'true'));
const searchAnalyticsMaxQueryTextChars = Math.max(Number(env.SEARCH_ANALYTICS_MAX_QUERY_TEXT_CHARS || process.env.SEARCH_ANALYTICS_MAX_QUERY_TEXT_CHARS || 120), 0);
const searchAnalyticsMaxResultsTracked = Math.max(Number(env.SEARCH_ANALYTICS_MAX_RESULTS_TRACKED || process.env.SEARCH_ANALYTICS_MAX_RESULTS_TRACKED || 30), 0);
const searchAnalyticsEventHubNamespace = String(env.SEARCH_ANALYTICS_EVENTHUB_NAMESPACE || process.env.SEARCH_ANALYTICS_EVENTHUB_NAMESPACE || '');
const searchAnalyticsEventHubName = String(env.SEARCH_ANALYTICS_EVENTHUB_NAME || process.env.SEARCH_ANALYTICS_EVENTHUB_NAME || '');
const searchAnalyticsEventHubConnectionString = String(env.SEARCH_ANALYTICS_EVENTHUB_CONNECTION_STRING || process.env.SEARCH_ANALYTICS_EVENTHUB_CONNECTION_STRING || '');
const searchAnalyticsEndpointPath = String(env.SEARCH_ANALYTICS_ENDPOINT_PATH || process.env.SEARCH_ANALYTICS_ENDPOINT_PATH || '/api/analytics/events');
const searchAnalyticsSendTimeoutMs = Math.max(Number(env.SEARCH_ANALYTICS_SEND_TIMEOUT_MS || process.env.SEARCH_ANALYTICS_SEND_TIMEOUT_MS || 5000), 1000);
const localHttpsEnabled = String(env.SEARCH_LOCAL_HTTPS || process.env.SEARCH_LOCAL_HTTPS || 'auto').toLowerCase() !== 'false';
const localHttpsKeyPath = env.SEARCH_LOCAL_HTTPS_KEY_PATH ||
  process.env.SEARCH_LOCAL_HTTPS_KEY_PATH ||
  path.join(os.homedir(), '.rushstack', 'rushstack-serve.key');
const localHttpsCertPath = env.SEARCH_LOCAL_HTTPS_CERT_PATH ||
  process.env.SEARCH_LOCAL_HTTPS_CERT_PATH ||
  path.join(os.homedir(), '.rushstack', 'rushstack-serve.pem');
const allowedOrigins = String(env.SEARCH_API_ALLOWED_ORIGINS || process.env.SEARCH_API_ALLOWED_ORIGINS || '')
  .split(',')
  .map((item) => item.trim())
  .filter(Boolean);
const requireEasyAuth = String(env.SEARCH_API_REQUIRE_EASY_AUTH || process.env.SEARCH_API_REQUIRE_EASY_AUTH || 'false').toLowerCase() === 'true';
const telemetryHashSalt = String(env.SEARCH_TELEMETRY_HASH_SALT || process.env.SEARCH_TELEMETRY_HASH_SALT || env.GRAPH_CLIENT_ID || 'iknowledge-search');
const chatRateLimitEnabled = String(env.SEARCH_CHAT_RATE_LIMIT_ENABLED || process.env.SEARCH_CHAT_RATE_LIMIT_ENABLED || 'false').toLowerCase() === 'true';
const debugRequests = String(env.SEARCH_API_DEBUG_REQUESTS || process.env.SEARCH_API_DEBUG_REQUESTS || 'false').toLowerCase() === 'true';
const chatRateLimitWindowMs = Math.max(Number(env.SEARCH_CHAT_RATE_LIMIT_WINDOW_MS || process.env.SEARCH_CHAT_RATE_LIMIT_WINDOW_MS || 60_000), 10_000);
const chatRateLimitMax = Math.max(Number(env.SEARCH_CHAT_RATE_LIMIT_MAX || process.env.SEARCH_CHAT_RATE_LIMIT_MAX || 20), 1);
const maxChatQuestionChars = Math.max(Number(env.SEARCH_CHAT_MAX_QUESTION_CHARS || process.env.SEARCH_CHAT_MAX_QUESTION_CHARS || 8000), 100);
const maxChatHistoryMessages = Math.max(Number(env.SEARCH_CHAT_MAX_HISTORY_MESSAGES || process.env.SEARCH_CHAT_MAX_HISTORY_MESSAGES || 8), 0);
const maxJsonBodyBytes = Math.max(Number(env.SEARCH_API_MAX_JSON_BODY_BYTES || process.env.SEARCH_API_MAX_JSON_BODY_BYTES || 8 * 1024 * 1024), 1024 * 1024);
const maxUploadMediaBytes = Math.max(Number(env.SEARCH_UPLOAD_MAX_MEDIA_BYTES || process.env.SEARCH_UPLOAD_MAX_MEDIA_BYTES || 25 * 1024 * 1024), 1024 * 1024);
const chatMemoryEnabled = String(env.SEARCH_CHAT_MEMORY_ENABLED || process.env.SEARCH_CHAT_MEMORY_ENABLED || 'false').toLowerCase() === 'true';
const chatMemoryMaxTurns = Math.max(Number(env.SEARCH_CHAT_MEMORY_MAX_TURNS || process.env.SEARCH_CHAT_MEMORY_MAX_TURNS || 12), 1);
const chatMemoryTtlDays = Math.max(Number(env.SEARCH_CHAT_MEMORY_TTL_DAYS || process.env.SEARCH_CHAT_MEMORY_TTL_DAYS || 14), 1);
const runtimeSetting = (key, fallback = '') => process.env[key] || env[key] || fallback;
const syncWorkerEnabled = String(runtimeSetting('SYNC_WORKER_ENABLED', 'false')).toLowerCase() === 'true';
const syncWorkerIntervalSeconds = Math.max(Number(runtimeSetting('SYNC_WORKER_INTERVAL_SECONDS', '180')), 30);
const syncTimerEnabled = String(runtimeSetting('SYNC_TIMER_ENABLED', 'false')).toLowerCase() === 'true';
const syncWebhookPath = runtimeSetting('SYNC_WEBHOOK_PATH', '/api/sync/notifications');
const syncManualPath = runtimeSetting('SYNC_MANUAL_PATH', '/api/sync/run');
const syncStatusPath = runtimeSetting('SYNC_STATUS_PATH', '/api/sync/status');
const syncAdminToken = runtimeSetting('SYNC_ADMIN_TOKEN', '');
const syncAdminTokenRequired = String(runtimeSetting('SYNC_ADMIN_TOKEN_REQUIRED', 'true')).toLowerCase() !== 'false';
const graphWebhookClientState = runtimeSetting('GRAPH_WEBHOOK_CLIENT_STATE', '');
const syncRunTimeoutMs = Math.max(Number(runtimeSetting('SYNC_RUN_TIMEOUT_MS', String(12 * 60 * 1000))), 60 * 1000);
const syncStaleRunMs = Math.max(Number(runtimeSetting('SYNC_STALE_RUN_MS', String(syncRunTimeoutMs + 60 * 1000))), syncRunTimeoutMs);
const syncTimerMaxChanges = String(runtimeSetting('SYNC_TIMER_MAX_CHANGES', '100'));
const syncManualMaxChanges = String(
  process.env.SYNC_MANUAL_MAX_CHANGES ||
  env.SYNC_MANUAL_MAX_CHANGES ||
  process.env.SYNC_DELTA_MAX_CHANGES ||
  env.SYNC_DELTA_MAX_CHANGES ||
  '250'
);
const syncInventoryPath = process.env.SEARCH_INVENTORY_PATH ||
  process.env.SYNC_INVENTORY_PATH ||
  env.SEARCH_INVENTORY_PATH ||
  env.SYNC_INVENTORY_PATH ||
  (fs.existsSync('/home')
    ? `/home/data/search-inventory-${String(env.AZURE_SEARCH_DOCUMENTS_INDEX || 'appservice').replace(/[^A-Za-z0-9_.-]/g, '_')}.json`
    : path.join('tools', 'search-backend', '.inventory', 'search-inventory-hosted.json'));
const chatMemoryPath = env.SEARCH_CHAT_MEMORY_PATH ||
  process.env.SEARCH_CHAT_MEMORY_PATH ||
  (fs.existsSync('/home')
    ? `/home/data/search-chat-sessions-${String(env.AZURE_SEARCH_DOCUMENTS_INDEX || 'appservice').replace(/[^A-Za-z0-9_.-]/g, '_')}.json`
    : path.join('tools', 'search-backend', '.inventory', 'search-chat-sessions-hosted.json'));
const syncState = {
  enabled: syncWorkerEnabled,
  running: false,
  queued: false,
  runCount: 0,
  activePid: undefined,
  activeMaxChanges: '',
  lastReason: '',
  lastStartedAt: '',
  lastCompletedAt: '',
  lastExitCode: undefined,
  lastExitSignal: '',
  lastError: '',
  lastOutputTail: '',
  nextTimerAt: ''
};
const isLocalRequest = (origin) => !origin || /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/i.test(origin);
const chatRateBuckets = new Map();
let chatSessionStore = { version: 1, sessions: {} };

const hashValue = (value) =>
  crypto
    .createHash('sha256')
    .update(`${telemetryHashSalt}:${String(value || 'anonymous')}`)
    .digest('hex')
    .slice(0, 32);

const newCorrelationId = () =>
  typeof crypto.randomUUID === 'function'
    ? crypto.randomUUID()
    : crypto.randomBytes(16).toString('hex');

const parseClientPrincipal = (request) => {
  const principalHeader = request.headers['x-ms-client-principal'];
  if (!principalHeader) {
    return undefined;
  }

  try {
    const json = Buffer.from(String(principalHeader), 'base64').toString('utf8');
    const principal = JSON.parse(json);
    const claims = Array.isArray(principal.claims) ? principal.claims : [];
    const claimValue = (...types) => {
      const found = claims.find((claim) => types.includes(claim.typ));
      return found ? found.val : '';
    };

    return {
      authProvider: principal.auth_typ || request.headers['x-ms-client-principal-idp'] || '',
      userId: principal.user_id || request.headers['x-ms-client-principal-id'] || claimValue('oid', 'http://schemas.microsoft.com/identity/claims/objectidentifier'),
      userDetails: principal.userDetails || request.headers['x-ms-client-principal-name'] || claimValue('preferred_username', 'upn', 'email')
    };
  } catch (error) {
    return {
      authProvider: request.headers['x-ms-client-principal-idp'] || 'unknown',
      userId: request.headers['x-ms-client-principal-id'] || '',
      userDetails: request.headers['x-ms-client-principal-name'] || '',
      parseError: error.message || String(error)
    };
  }
};

const buildRequestContext = (request) => {
  const principal = parseClientPrincipal(request);
  const fallbackIdentity = request.headers['x-forwarded-for'] || request.socket.remoteAddress || 'anonymous';
  const identity = principal?.userId || principal?.userDetails || fallbackIdentity;
  const correlationId = String(
    request.headers['x-correlation-id'] ||
    request.headers['x-ms-client-request-id'] ||
    request.headers['x-request-id'] ||
    newCorrelationId()
  );
  const userHash = hashValue(identity);

  return {
    correlationId,
    principal,
    userHash,
    rateLimitKey: userHash,
    clientIpHash: hashValue(fallbackIdentity)
  };
};

const logEvent = (context, event) => {
  const payload = {
    level: event.level || 'info',
    correlationId: context?.correlationId,
    userHash: context?.userHash,
    ...event
  };

  const writer = payload.level === 'error' ? console.error : console.log;
  writer(JSON.stringify(payload));
};
let activeSyncChild;
let activeSyncStartedAtMs = 0;

const summarizeSearchRequest = (body) => {
  const filterGroups = body && typeof body.filters === 'object' && body.filters
    ? Object.entries(body.filters)
        .filter(([, value]) => Array.isArray(value) && value.length > 0)
        .map(([key, value]) => ({ key, count: value.length }))
    : [];

  return {
    queryLength: String(body?.query || '').length,
    queryPreview: String(body?.query || '').slice(0, 160),
    top: body?.top,
    skip: body?.skip,
    includeFacets: body?.includeFacets,
    sort: body?.sort || 'relevance',
    filterGroups,
    filters: sanitizeFilterTelemetry(body?.filters)
  };
};

const summarizeFacetValues = (facets, key, limit = 12) => {
  const values = facets && Array.isArray(facets[key]) ? facets[key] : [];
  return values.slice(0, limit).map((item) => ({
    value: truncateText(item?.value || '', 120),
    count: Number(item?.count || 0)
  }));
};

const summarizeSearchResponse = (payload) => {
  if (!isPlainObject(payload)) {
    return {};
  }

  return {
    totalCount: payload.totalCount,
    resultCount: Array.isArray(payload.results) ? payload.results.length : undefined,
    topDepartmentFacets: summarizeFacetValues(payload.facets, 'department'),
    topBusinessUnitFacets: summarizeFacetValues(payload.facets, 'bu'),
    topDocumentTypeFacets: summarizeFacetValues(payload.facets, 'documentType')
  };
};

const analyticsHash = (value) =>
  crypto
    .createHash('sha256')
    .update(`${telemetryHashSalt}:analytics:${String(value || '')}`)
    .digest('hex')
    .slice(0, 32);

const isPlainObject = (value) => Boolean(value && typeof value === 'object' && !Array.isArray(value));

const truncateText = (value, maxLength) => {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  if (!maxLength || text.length <= maxLength) {
    return text;
  }
  return `${text.slice(0, maxLength - 1)}...`;
};

const countQueryTokensApprox = (query) =>
  String(query || '')
    .split(/[\s,;|/\\()[\]{}"'`]+/g)
    .map((token) => token.trim())
    .filter(Boolean).length;

const isLikelyPastedQuery = (query) => String(query || '').length > 500 || /\n/.test(String(query || ''));

const sanitizeQueryTelemetry = (query) => {
  if (!searchAnalyticsStoreSanitizedQuery || searchAnalyticsMaxQueryTextChars <= 0) {
    return '';
  }
  if (isLikelyPastedQuery(query) && !searchAnalyticsStoreRawQuery) {
    return '';
  }
  return truncateText(query, searchAnalyticsMaxQueryTextChars);
};

const sanitizeFilterTelemetry = (filters) => {
  if (!isPlainObject(filters)) {
    return {};
  }

  const allowedKeys = new Set([
    'documentType',
    'fileExtension',
    'bu',
    'businessUnit',
    'department',
    'client',
    'region',
    'geography',
    'therapyArea',
    'diseaseArea'
  ]);
  const sanitized = {};
  for (const [key, value] of Object.entries(filters)) {
    if (!allowedKeys.has(key)) {
      continue;
    }
    const values = Array.isArray(value) ? value : [value];
    sanitized[key] = values
      .map((item) => truncateText(item, 160))
      .filter(Boolean)
      .slice(0, 20);
  }
  return sanitized;
};

const normalizeSearchAnalyticsMeta = (body) => {
  const searchRequestId = newCorrelationId();
  const suppliedSessionId = String(body?.searchSessionId || body?.analytics?.searchSessionId || '').trim();
  const generatedSessionId = newCorrelationId();
  return {
    searchRequestId,
    searchSessionIdRaw: suppliedSessionId || generatedSessionId,
    searchSessionId: suppliedSessionId ? analyticsHash(suppliedSessionId) : analyticsHash(generatedSessionId)
  };
};

const buildResultRanks = (payload) => {
  const results = Array.isArray(payload)
    ? payload
    : Array.isArray(payload?.results)
      ? payload.results
      : [];

  return results.slice(0, searchAnalyticsMaxResultsTracked).map((item, index) => ({
    rank: index + 1 + Math.max(Number(payload?.page?.skip || 0), 0),
    id: String(item?.documentId || item?.id || ''),
    listItemId: String(item?.listItemId || ''),
    title: truncateText(item?.title || item?.fileName || '', 220),
    documentType: truncateText(item?.documentType || '', 120),
    businessUnit: truncateText(item?.businessUnit || item?.bu || '', 120),
    department: truncateText(item?.department || '', 180),
    score: typeof item?.score === 'number' ? item.score : undefined
  }));
};

const buildSearchCompletedEvent = (context, body, payload, meta, durationMs) => {
  const query = String(body?.query || '');
  const page = isPlainObject(payload?.page) ? payload.page : {};
  const results = Array.isArray(payload)
    ? payload
    : Array.isArray(payload?.results)
      ? payload.results
      : [];

  return {
    timestamp: new Date().toISOString(),
    eventType: 'search_completed',
    eventId: newCorrelationId(),
    searchRequestId: meta.searchRequestId,
    searchSessionId: meta.searchSessionId,
    correlationId: context?.correlationId || '',
    environment: searchAnalyticsEnvironment,
    backendVersion: searchAnalyticsBackendVersion,
    userHash: context?.userHash || '',
    queryHash: analyticsHash(query.toLowerCase()),
    querySanitized: sanitizeQueryTelemetry(query),
    queryLength: query.length,
    queryTokenCount: countQueryTokensApprox(query),
    isPastedContent: isLikelyPastedQuery(query),
    filters: sanitizeFilterTelemetry(body?.filters),
    sort: String(body?.sort || 'relevance'),
    skip: Math.max(Number(body?.skip || page.skip || 0), 0),
    top: Math.max(Number(body?.top || page.top || results.length || 0), 0),
    totalCount: Number(payload?.totalCount ?? page.totalAvailable ?? results.length ?? 0),
    returnedCount: results.length,
    resultIds: buildResultRanks(payload).map((item) => item.id).filter(Boolean),
    resultRanks: buildResultRanks(payload),
    searchPath: String(payload?.debug?.searchPath || payload?.searchPath || ''),
    fallbackApplied: Boolean(payload?.fallbackApplied || payload?.debug?.fallbackApplied || page.capped),
    normalizedFacetsEnabled,
    latencyMs: durationMs,
    statusCode: 200
  };
};

const isWildcardAnalyticsQuery = (body) =>
  String(body?.query || '').trim() === '*';

const shouldEmitSearchAnalyticsForRequest = (body) => {
  if (!isWildcardAnalyticsQuery(body)) {
    return true;
  }

  return Boolean(body?.analytics?.recordWildcardSearch || body?.recordWildcardSearch);
};

const buildSearchFailedEvent = (context, body, meta, durationMs, error) => {
  const query = String(body?.query || '');
  const statusCode = Number(error?.statusCode || error?.status || 500);
  return {
    timestamp: new Date().toISOString(),
    eventType: 'search_failed',
    eventId: newCorrelationId(),
    searchRequestId: meta.searchRequestId,
    searchSessionId: meta.searchSessionId,
    correlationId: context?.correlationId || '',
    environment: searchAnalyticsEnvironment,
    backendVersion: searchAnalyticsBackendVersion,
    userHash: context?.userHash || '',
    queryHash: analyticsHash(query.toLowerCase()),
    querySanitized: sanitizeQueryTelemetry(query),
    queryLength: query.length,
    queryTokenCount: countQueryTokensApprox(query),
    isPastedContent: isLikelyPastedQuery(query),
    filters: sanitizeFilterTelemetry(body?.filters),
    sort: String(body?.sort || 'relevance'),
    skip: Math.max(Number(body?.skip || 0), 0),
    top: Math.max(Number(body?.top || 0), 0),
    normalizedFacetsEnabled,
    latencyMs: durationMs,
    statusCode,
    errorCode: truncateText(error?.code || error?.name || 'search_error', 120),
    errorMessage: truncateText(error?.message || String(error || 'Search failed'), 500)
  };
};

const buildFrontendAnalyticsEvent = (context, body) => {
  const event = isPlainObject(body) ? body : {};
  const action = String(event.action || '').toLowerCase();
  if (!['view', 'download', 'open_reference'].includes(action)) {
    return undefined;
  }

  return {
    timestamp: new Date().toISOString(),
    eventType: 'search_result_action',
    eventId: newCorrelationId(),
    searchRequestId: truncateText(event.searchRequestId || '', 80),
    searchSessionId: event.searchSessionId ? analyticsHash(event.searchSessionId) : '',
    correlationId: context?.correlationId || '',
    environment: searchAnalyticsEnvironment,
    frontendVersion: truncateText(event.frontendVersion || '', 80),
    backendVersion: searchAnalyticsBackendVersion,
    userHash: context?.userHash || '',
    action,
    documentId: truncateText(event.documentId || event.id || '', 120),
    listItemId: truncateText(event.listItemId || '', 80),
    documentTitle: truncateText(event.documentTitle || event.title || '', 220),
    rank: Number.isFinite(Number(event.rank)) ? Number(event.rank) : undefined,
    timeSinceSearchMs: Number.isFinite(Number(event.timeSinceSearchMs)) ? Number(event.timeSinceSearchMs) : undefined,
    sourceSurface: truncateText(event.sourceSurface || 'search', 80),
    statusCode: 202
  };
};

const eventHubHttpRequest = (url, options, body) => new Promise((resolve, reject) => {
  const parsed = new URL(url);
  const transport = parsed.protocol === 'http:' ? http : https;
  const requestOptions = {
    method: options.method || 'POST',
    hostname: parsed.hostname,
    port: parsed.port || undefined,
    path: `${parsed.pathname}${parsed.search}`,
    headers: options.headers || {},
    timeout: searchAnalyticsSendTimeoutMs
  };
  const req = transport.request(requestOptions, (res) => {
    const chunks = [];
    res.on('data', (chunk) => chunks.push(chunk));
    res.on('end', () => {
      const text = Buffer.concat(chunks).toString('utf8');
      if (res.statusCode >= 200 && res.statusCode < 300) {
        resolve({ statusCode: res.statusCode, text });
      } else {
        reject(new Error(`Event Hub request failed ${res.statusCode}: ${text.slice(0, 300)}`));
      }
    });
  });
  req.on('timeout', () => {
    req.destroy(new Error(`Event Hub request timed out after ${searchAnalyticsSendTimeoutMs}ms`));
  });
  req.on('error', reject);
  if (body) {
    req.write(body);
  }
  req.end();
});

let managedIdentityTokenCache = { accessToken: '', expiresAt: 0 };

const getManagedIdentityToken = async () => {
  const now = Date.now();
  if (managedIdentityTokenCache.accessToken && managedIdentityTokenCache.expiresAt - 60_000 > now) {
    return managedIdentityTokenCache.accessToken;
  }

  const endpoint = process.env.IDENTITY_ENDPOINT || process.env.MSI_ENDPOINT;
  const secret = process.env.IDENTITY_HEADER || process.env.MSI_SECRET;
  if (!endpoint || !secret) {
    throw new Error('Managed identity endpoint is not available.');
  }

  const separator = endpoint.includes('?') ? '&' : '?';
  const tokenUrl = `${endpoint}${separator}resource=${encodeURIComponent('https://eventhubs.azure.net/')}&api-version=2019-08-01`;
  const headerName = process.env.IDENTITY_HEADER ? 'X-IDENTITY-HEADER' : 'Secret';
  const response = await eventHubHttpRequest(tokenUrl, {
    method: 'GET',
    headers: { [headerName]: secret }
  });
  const parsed = JSON.parse(response.text);
  const expiresAt = parsed.expires_on
    ? Number(parsed.expires_on) * 1000
    : Date.now() + Math.max(Number(parsed.expires_in || 300), 60) * 1000;
  managedIdentityTokenCache = {
    accessToken: parsed.access_token,
    expiresAt
  };
  return managedIdentityTokenCache.accessToken;
};

const parseConnectionString = (connectionString) => String(connectionString || '')
  .split(';')
  .map((part) => part.trim())
  .filter(Boolean)
  .reduce((acc, part) => {
    const index = part.indexOf('=');
    if (index > 0) {
      acc[part.slice(0, index)] = part.slice(index + 1);
    }
    return acc;
  }, {});

const buildEventHubSasToken = (eventHubUrl) => {
  const parts = parseConnectionString(searchAnalyticsEventHubConnectionString);
  if (!parts.SharedAccessKeyName || !parts.SharedAccessKey) {
    throw new Error('Event Hub connection string is missing SharedAccessKeyName or SharedAccessKey.');
  }
  const encodedUri = encodeURIComponent(eventHubUrl.toLowerCase());
  const expires = Math.floor(Date.now() / 1000) + 10 * 60;
  const stringToSign = `${encodedUri}\n${expires}`;
  const signature = encodeURIComponent(
    crypto
      .createHmac('sha256', parts.SharedAccessKey)
      .update(stringToSign)
      .digest('base64')
  );
  return `SharedAccessSignature sr=${encodedUri}&sig=${signature}&se=${expires}&skn=${encodeURIComponent(parts.SharedAccessKeyName)}`;
};

const publishAnalyticsEvent = async (context, event) => {
  if (!searchAnalyticsEnabled || searchAnalyticsSink !== 'eventhub' || !event) {
    return { sent: false, reason: searchAnalyticsEnabled ? 'not_configured' : 'disabled' };
  }
  if (!searchAnalyticsEventHubNamespace || !searchAnalyticsEventHubName) {
    return { sent: false, reason: 'eventhub_missing' };
  }

  const eventHubResourceUrl = `https://${searchAnalyticsEventHubNamespace}.servicebus.windows.net/${encodeURIComponent(searchAnalyticsEventHubName)}`;
  const eventHubUrl = `${eventHubResourceUrl}/messages`;
  const body = JSON.stringify(event);
  const authorization = searchAnalyticsEventHubConnectionString
    ? buildEventHubSasToken(eventHubResourceUrl)
    : `Bearer ${await getManagedIdentityToken()}`;

  await eventHubHttpRequest(eventHubUrl, {
    method: 'POST',
    headers: {
      Authorization: authorization,
      'Content-Type': 'application/json',
      'Content-Length': Buffer.byteLength(body)
    }
  }, body);

  logEvent(context, {
    level: 'info',
    event: 'search-analytics-sent',
    analyticsEventType: event.eventType,
    searchRequestId: event.searchRequestId || '',
    action: event.action || ''
  });
  return { sent: true };
};

const emitAnalyticsEvent = (context, event) => {
  if (!searchAnalyticsEnabled || !event) {
    return;
  }

  publishAnalyticsEvent(context, event).catch((error) => {
    logEvent(context, {
      level: 'error',
      event: 'search-analytics-failed',
      analyticsEventType: event.eventType,
      searchRequestId: event.searchRequestId || '',
      error: error.message || String(error)
    });
  });
};

const checkRateLimit = (key) => {
  const now = Date.now();
  const resetAt = now + chatRateLimitWindowMs;
  const bucket = chatRateBuckets.get(key);

  if (!bucket || bucket.resetAt <= now) {
    chatRateBuckets.set(key, { count: 1, resetAt });
    return { allowed: true, remaining: chatRateLimitMax - 1, resetAt };
  }

  if (bucket.count >= chatRateLimitMax) {
    return { allowed: false, remaining: 0, resetAt: bucket.resetAt };
  }

  bucket.count += 1;
  return { allowed: true, remaining: Math.max(0, chatRateLimitMax - bucket.count), resetAt: bucket.resetAt };
};

setInterval(() => {
  const now = Date.now();
  for (const [key, bucket] of chatRateBuckets.entries()) {
    if (bucket.resetAt <= now) {
      chatRateBuckets.delete(key);
    }
  }
}, Math.min(chatRateLimitWindowMs, 60_000)).unref?.();

const loadChatSessions = () => {
  if (!chatMemoryEnabled) return;
  try {
    if (!fs.existsSync(chatMemoryPath)) return;
    const parsed = JSON.parse(fs.readFileSync(chatMemoryPath, 'utf8'));
    if (parsed && parsed.sessions && typeof parsed.sessions === 'object') {
      chatSessionStore = { version: 1, sessions: parsed.sessions };
    }
  } catch (error) {
    console.warn(JSON.stringify({
      level: 'error',
      event: 'chat-memory-load-failed',
      path: chatMemoryPath,
      error: error.message || String(error)
    }));
  }
};

const saveChatSessions = () => {
  if (!chatMemoryEnabled) return;
  try {
    fs.mkdirSync(path.dirname(chatMemoryPath), { recursive: true });
    fs.writeFileSync(chatMemoryPath, JSON.stringify(chatSessionStore), 'utf8');
  } catch (error) {
    console.warn(JSON.stringify({
      level: 'error',
      event: 'chat-memory-save-failed',
      path: chatMemoryPath,
      error: error.message || String(error)
    }));
  }
};

const cleanupChatSessions = () => {
  if (!chatMemoryEnabled) return;
  const expiresBefore = Date.now() - chatMemoryTtlDays * 24 * 60 * 60 * 1000;
  for (const [id, session] of Object.entries(chatSessionStore.sessions)) {
    const updatedAt = Date.parse(session.updatedAt || '');
    if (!Number.isFinite(updatedAt) || updatedAt < expiresBefore) {
      delete chatSessionStore.sessions[id];
    }
  }
};

const normalizeConversationId = (value) => {
  const text = String(value || '').trim();
  return /^[A-Za-z0-9_-]{16,80}$/.test(text) ? text : '';
};

const normalizeChatMessages = (messages) =>
  (Array.isArray(messages) ? messages : [])
    .map((message) => ({
      role: message && message.role === 'assistant' ? 'assistant' : 'user',
      content: String((message && message.content) || '').trim().slice(0, maxChatQuestionChars),
      at: message && message.at ? String(message.at) : new Date().toISOString()
    }))
    .filter((message) => message.content)
    .slice(-chatMemoryMaxTurns * 2);

const resolveChatSession = (context, body) => {
  if (!chatMemoryEnabled) {
    return {
      conversationId: normalizeConversationId(body.conversationId) || newCorrelationId(),
      history: Array.isArray(body.history) ? body.history : [],
      persisted: false
    };
  }

  cleanupChatSessions();
  const requestedId = normalizeConversationId(body.conversationId);
  let conversationId = requestedId;
  let session = conversationId ? chatSessionStore.sessions[conversationId] : undefined;

  if (session && session.userHash !== context.userHash) {
    conversationId = '';
    session = undefined;
  }

  if (!conversationId) {
    conversationId = newCorrelationId().replace(/-/g, '');
  }

  if (!session) {
    session = {
      userHash: context.userHash,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      messages: normalizeChatMessages(body.history)
    };
    chatSessionStore.sessions[conversationId] = session;
  }

  return {
    conversationId,
    session,
    history: normalizeChatMessages(session.messages),
    persisted: true
  };
};

const rememberChatTurn = (conversation, question, answer) => {
  if (!chatMemoryEnabled || !conversation.session) return;
  const now = new Date().toISOString();
  conversation.session.messages = normalizeChatMessages([
    ...(conversation.session.messages || []),
    { role: 'user', content: question, at: now },
    { role: 'assistant', content: answer, at: now }
  ]);
  conversation.session.updatedAt = now;
  chatSessionStore.sessions[conversation.conversationId] = conversation.session;
  saveChatSessions();
};

loadChatSessions();

const resolveCorsOrigin = (origin) => {
  if (!origin) return allowedOrigins[0] || '*';
  if (allowedOrigins.includes(origin)) return origin;
  if (allowedOrigins.length === 0 && isLocalRequest(origin)) return origin;
  return '';
};

const readJsonBody = (request, maxBytes = maxJsonBodyBytes) =>
  new Promise((resolve, reject) => {
    let body = '';
    let bytes = 0;
    request.on('data', (chunk) => {
      bytes += chunk.length;
      body += chunk;
      if (bytes > maxBytes) {
        const error = new Error('Request body is too large.');
        error.statusCode = 413;
        reject(error);
        request.destroy();
      }
    });
    request.on('end', () => {
      if (!body.trim()) {
        resolve({});
        return;
      }

      try {
        resolve(JSON.parse(body));
      } catch (error) {
        reject(new Error(`Invalid JSON body: ${error.message}`));
      }
    });
    request.on('error', reject);
  });

const readRequestBuffer = (request, maxBytes) =>
  new Promise((resolve, reject) => {
    const chunks = [];
    let bytes = 0;
    request.on('data', (chunk) => {
      bytes += chunk.length;
      if (bytes > maxBytes) {
        const error = new Error('Uploaded media file is too large.');
        error.statusCode = 413;
        reject(error);
        request.destroy();
        return;
      }
      chunks.push(chunk);
    });
    request.on('end', () => resolve(Buffer.concat(chunks)));
    request.on('error', reject);
  });

const parseContentDisposition = (value) => {
  const parsed = {};
  String(value || '').split(';').forEach((part) => {
    const [rawKey, ...rest] = part.trim().split('=');
    const key = rawKey.trim().toLowerCase();
    if (!key || rest.length === 0) return;
    parsed[key] = rest.join('=').trim().replace(/^"|"$/g, '');
  });
  return parsed;
};

const parseMultipartFile = (body, contentType) => {
  const boundaryMatch = /boundary=(?:"([^"]+)"|([^;]+))/i.exec(String(contentType || ''));
  const boundary = boundaryMatch && (boundaryMatch[1] || boundaryMatch[2]);
  if (!boundary) {
    const error = new Error('multipart/form-data boundary is required.');
    error.statusCode = 400;
    throw error;
  }

  const boundaryBuffer = Buffer.from(`--${boundary}`);
  const headerBreak = Buffer.from('\r\n\r\n');
  let position = body.indexOf(boundaryBuffer);
  while (position !== -1) {
    let partStart = position + boundaryBuffer.length;
    if (body[partStart] === 45 && body[partStart + 1] === 45) break;
    if (body[partStart] === 13 && body[partStart + 1] === 10) partStart += 2;

    const headersEnd = body.indexOf(headerBreak, partStart);
    if (headersEnd === -1) break;
    const headersText = body.subarray(partStart, headersEnd).toString('utf8');
    const headers = {};
    headersText.split(/\r?\n/).forEach((line) => {
      const index = line.indexOf(':');
      if (index > -1) {
        headers[line.slice(0, index).trim().toLowerCase()] = line.slice(index + 1).trim();
      }
    });

    const contentStart = headersEnd + headerBreak.length;
    const nextBoundary = body.indexOf(Buffer.from(`\r\n--${boundary}`), contentStart);
    if (nextBoundary === -1) break;

    const disposition = parseContentDisposition(headers['content-disposition']);
    if (disposition.name === 'file') {
      const filename = disposition.filename || 'media';
      const buffer = body.subarray(contentStart, nextBoundary);
      if (buffer.length === 0) {
        const error = new Error('Uploaded file is empty.');
        error.statusCode = 400;
        throw error;
      }
      return {
        buffer,
        filename,
        contentType: headers['content-type'] || 'application/octet-stream'
      };
    }

    position = body.indexOf(boundaryBuffer, nextBoundary + 2);
  }

  const error = new Error('multipart/form-data must include a file field named "file".');
  error.statusCode = 400;
  throw error;
};

const readMultipartFile = async (request) => {
  const contentType = String(request.headers['content-type'] || '');
  if (!/^multipart\/form-data\b/i.test(contentType)) {
    const error = new Error('Content-Type must be multipart/form-data.');
    error.statusCode = 415;
    throw error;
  }
  const body = await readRequestBuffer(request, maxUploadMediaBytes + 1024 * 1024);
  const file = parseMultipartFile(body, contentType);
  if (file.buffer.length > maxUploadMediaBytes) {
    const error = new Error(`Uploaded media file exceeds ${maxUploadMediaBytes} bytes.`);
    error.statusCode = 413;
    throw error;
  }
  return file;
};

const writeJson = (response, statusCode, payload, origin, context, extraHeaders = {}) => {
  const corsOrigin = resolveCorsOrigin(origin);
  const headers = {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Credentials': 'true',
    'Access-Control-Allow-Headers': 'content-type, authorization, x-sync-admin-token, x-correlation-id',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Private-Network': 'true',
    'Vary': 'Origin',
    ...extraHeaders
  };
  if (context?.correlationId) {
    headers['x-correlation-id'] = context.correlationId;
  }
  if (corsOrigin) {
    headers['Access-Control-Allow-Origin'] = corsOrigin;
  }
  response.writeHead(statusCode, headers);
  response.end(JSON.stringify(payload));
};

const writeText = (response, statusCode, payload) => {
  response.writeHead(statusCode, {
    'Content-Type': 'text/plain; charset=utf-8'
  });
  response.end(payload);
};

const isAuthorizedSyncRequest = (request) => {
  if (!syncAdminTokenRequired) {
    return true;
  }
  if (requireEasyAuth && !request.headers['x-ms-client-principal']) {
    return false;
  }
  if (!syncAdminToken) {
    return requireEasyAuth && Boolean(request.headers['x-ms-client-principal']);
  }
  const authorization = String(request.headers.authorization || '');
  const tokenHeader = String(request.headers['x-sync-admin-token'] || '');
  return tokenHeader === syncAdminToken || authorization === `Bearer ${syncAdminToken}`;
};

const appendOutputTail = (text) => {
  if (!text) return;
  syncState.lastOutputTail = `${syncState.lastOutputTail || ''}${text}`.slice(-8000);
};

const safeMaxChangesArg = (value, fallback) => {
  const raw = String(value || fallback || '').trim();
  if (raw.toLowerCase() === 'all') return 'all';
  const parsed = Math.floor(Number(raw));
  return Number.isFinite(parsed) && parsed > 0 ? String(parsed) : String(fallback || '100');
};

const stopActiveSyncChild = (reason) => {
  if (!activeSyncChild || activeSyncChild.killed) {
    return false;
  }

  appendOutputTail(`\n${new Date().toISOString()} sync child termination requested: ${reason}\n`);
  syncState.queued = false;
  syncState.lastError = reason;
  try {
    const childToStop = activeSyncChild;
    childToStop.kill('SIGTERM');
    setTimeout(() => {
      if (activeSyncChild === childToStop && !childToStop.killed) {
        try {
          childToStop.kill('SIGKILL');
        } catch {
          childToStop.kill();
        }
      }
    }, 10_000).unref?.();
    return true;
  } catch (error) {
    syncState.lastError = `Failed to terminate sync child: ${error.message || String(error)}`;
    return false;
  }
};

const runDeltaCatchup = (reason, maxChangesArg) => new Promise((resolve) => {
  const safeMaxChanges = safeMaxChangesArg(maxChangesArg, syncTimerMaxChanges);
  syncState.running = true;
  syncState.runCount += 1;
  syncState.activeMaxChanges = safeMaxChanges;
  syncState.lastReason = reason;
  syncState.lastStartedAt = new Date().toISOString();
  syncState.lastCompletedAt = '';
  syncState.lastExitCode = undefined;
  syncState.lastExitSignal = '';
  syncState.lastError = '';
  syncState.lastOutputTail = '';
  activeSyncStartedAtMs = Date.now();

  const child = spawn(process.execPath, [
    'tools/search-backend/scripts/delta-sync.js',
    envPath,
    'delta',
    safeMaxChanges
  ], {
    cwd: process.cwd(),
    env: {
      ...process.env,
      SEARCH_ENV_PATH: envPath,
      SEARCH_INVENTORY_PATH: syncInventoryPath
    },
    windowsHide: true
  });
  activeSyncChild = child;
  syncState.activePid = child.pid;

  appendOutputTail(`${new Date().toISOString()} sync run started: reason=${reason}; maxChanges=${safeMaxChanges}; pid=${child.pid || 'unknown'}\n`);
  const timeout = setTimeout(() => {
    stopActiveSyncChild(`sync run exceeded timeout ${syncRunTimeoutMs}ms`);
  }, syncRunTimeoutMs);
  timeout.unref?.();

  child.stdout.on('data', (chunk) => appendOutputTail(chunk.toString()));
  child.stderr.on('data', (chunk) => appendOutputTail(chunk.toString()));

  child.on('error', (error) => {
    syncState.lastError = error.message || String(error);
  });

  child.on('close', (code, signal) => {
    clearTimeout(timeout);
    if (activeSyncChild === child) {
      activeSyncChild = undefined;
    }
    activeSyncStartedAtMs = 0;
    syncState.running = false;
    syncState.activePid = undefined;
    syncState.lastExitCode = code;
    syncState.lastExitSignal = signal || '';
    syncState.lastCompletedAt = new Date().toISOString();
    if (code !== 0 && !syncState.lastError) {
      syncState.lastError = `delta-sync exited with code ${code}${signal ? ` signal ${signal}` : ''}`;
    }
    console.log(JSON.stringify({
      level: code === 0 ? 'info' : 'error',
      event: 'sync-delta-completed',
      reason,
      code,
      signal,
      maxChanges: safeMaxChanges,
      inventoryPath: syncInventoryPath
    }));
    resolve(code);
  });
});

const triggerSync = (reason, options = {}) => {
  if (!syncWorkerEnabled) {
    return { accepted: false, reason: 'sync_worker_disabled', state: syncState };
  }
  if (syncState.running) {
    const elapsedMs = activeSyncStartedAtMs ? Date.now() - activeSyncStartedAtMs : 0;
    if (elapsedMs > syncStaleRunMs) {
      const stopped = stopActiveSyncChild(`stale sync run exceeded ${syncStaleRunMs}ms`);
      return {
        accepted: false,
        queued: false,
        reason: stopped ? 'sync_run_stale_terminating' : 'sync_run_stale_no_child',
        state: syncState
      };
    }
    syncState.queued = true;
    return { accepted: true, queued: true, reason: 'sync_already_running', state: syncState };
  }

  const maxChanges = safeMaxChangesArg(options.maxChanges, reason === 'timer_fallback' ? syncTimerMaxChanges : syncManualMaxChanges);
  setTimeout(async () => {
    do {
      syncState.queued = false;
      await runDeltaCatchup(reason, maxChanges);
      reason = 'queued_after_previous_run';
    } while (syncState.queued);
  }, 0);

  return { accepted: true, queued: false, maxChanges, state: syncState };
};

const validateGraphNotifications = (body) => {
  const notifications = Array.isArray(body.value) ? body.value : [];
  if (!graphWebhookClientState) {
    return {
      ok: false,
      notifications,
      error: 'GRAPH_WEBHOOK_CLIENT_STATE is required before accepting Graph notifications.'
    };
  }
  if (graphWebhookClientState) {
    const invalid = notifications.find((notification) => notification.clientState !== graphWebhookClientState);
    if (invalid) {
      return { ok: false, notifications, error: 'Invalid Graph notification clientState.' };
    }
  }
  return { ok: true, notifications };
};

const validateChatBody = (body) => {
  const question = String(body.question || '');
  if (question.length > maxChatQuestionChars) {
    return {
      ok: false,
      statusCode: 413,
      payload: { error: `Question is too large. Maximum allowed is ${maxChatQuestionChars} characters.` }
    };
  }

  if (Array.isArray(body.history) && body.history.length > maxChatHistoryMessages) {
    body.history = body.history.slice(-maxChatHistoryMessages);
  }

  return { ok: true };
};

const routeRequest = async (request, context) => {
  const url = new URL(request.url || '/', `http://${request.headers.host || 'localhost'}`);
  if (
    url.pathname === syncWebhookPath &&
    url.searchParams.has('validationToken') &&
    (request.method === 'GET' || request.method === 'POST')
  ) {
    return {
      statusCode: 200,
      textPayload: url.searchParams.get('validationToken') || ''
    };
  }

  const isMediaUploadRequest = request.method === 'POST' && url.pathname === '/api/upload/transcribe-media';
  const body = request.method === 'POST' && !isMediaUploadRequest ? await readJsonBody(request) : {};

  if (request.method === 'GET' && url.pathname === '/healthz') {
    return {
      statusCode: 200,
      payload: {
        ok: true,
        service: 'iknowledge-search-api',
        documentsIndex: env.AZURE_SEARCH_DOCUMENTS_INDEX,
        chunksIndex: env.AZURE_SEARCH_CHUNKS_INDEX,
        peopleIndex: env.AZURE_SEARCH_PEOPLE_INDEX,
        syncWorker: {
          enabled: syncWorkerEnabled,
          timerEnabled: syncTimerEnabled,
          running: syncState.running,
          queued: syncState.queued,
          lastCompletedAt: syncState.lastCompletedAt,
          lastExitCode: syncState.lastExitCode,
          lastExitSignal: syncState.lastExitSignal
        },
        chatMemory: {
          enabled: chatMemoryEnabled,
          sessions: chatMemoryEnabled ? Object.keys(chatSessionStore.sessions).length : 0
        },
        chatRateLimit: {
          enabled: chatRateLimitEnabled,
          max: chatRateLimitEnabled ? chatRateLimitMax : 0,
          windowMs: chatRateLimitEnabled ? chatRateLimitWindowMs : 0
        },
        normalizedFacets: {
          enabled: normalizedFacetsEnabled
        },
        searchAnalytics: {
          enabled: searchAnalyticsEnabled,
          sink: searchAnalyticsSink,
          endpointPath: searchAnalyticsEndpointPath,
          eventHubConfigured: Boolean(searchAnalyticsEventHubNamespace && searchAnalyticsEventHubName),
          managedIdentityAvailable: Boolean(process.env.IDENTITY_ENDPOINT || process.env.MSI_ENDPOINT),
          connectionStringConfigured: Boolean(searchAnalyticsEventHubConnectionString)
        }
      }
    };
  }

  if (request.method === 'GET' && url.pathname === '/api/me') {
    if (requireEasyAuth && !request.headers['x-ms-client-principal']) {
      return { statusCode: 401, payload: { error: 'Authentication required.' } };
    }
    return {
      statusCode: 200,
      payload: {
        authenticated: Boolean(context.principal),
        userHash: context.userHash,
        authProvider: context.principal?.authProvider || '',
        correlationId: context.correlationId
      }
    };
  }

  if (request.method === 'GET' && url.pathname === syncStatusPath) {
    if (!isAuthorizedSyncRequest(request)) {
      return { statusCode: 401, payload: { error: 'Authentication required.' } };
    }
    return {
      statusCode: 200,
      payload: {
        ...syncState,
        inventoryPath: syncInventoryPath,
        timerEnabled: syncTimerEnabled,
        intervalSeconds: syncWorkerIntervalSeconds,
        adminTokenRequired: syncAdminTokenRequired,
        webhookPath: syncWebhookPath
      }
    };
  }

  if (request.method !== 'POST') {
    return { statusCode: 405, payload: { error: 'Method not allowed' } };
  }

  if (url.pathname === syncWebhookPath) {
    const validation = validateGraphNotifications(body);
    if (!validation.ok) {
      return { statusCode: 403, payload: { error: validation.error } };
    }
    const trigger = triggerSync(`graph_webhook:${validation.notifications.length || 0}`);
    return {
      statusCode: 202,
      payload: {
        ok: true,
        notifications: validation.notifications.length,
        trigger
      }
    };
  }

  if (url.pathname === syncManualPath) {
    if (!isAuthorizedSyncRequest(request)) {
      return { statusCode: 401, payload: { error: 'Authentication required.' } };
    }
    return {
      statusCode: 202,
      payload: triggerSync(body.reason || 'manual', { maxChanges: body.maxChanges })
    };
  }

  if (requireEasyAuth && !request.headers['x-ms-client-principal']) {
    return { statusCode: 401, payload: { error: 'Authentication required.' } };
  }

  if (url.pathname === searchAnalyticsEndpointPath) {
    const events = Array.isArray(body?.events) ? body.events : [body];
    let accepted = 0;
    for (const eventBody of events.slice(0, 20)) {
      const event = buildFrontendAnalyticsEvent(context, eventBody);
      if (event) {
        accepted += 1;
        emitAnalyticsEvent(context, event);
      }
    }
    return {
      statusCode: 202,
      payload: {
        accepted,
        enabled: searchAnalyticsEnabled,
        sink: searchAnalyticsSink
      }
    };
  }

  if (url.pathname === '/api/search/documents') {
    if (debugRequests) {
      logEvent(context, {
        level: 'info',
        event: 'search-documents-request-debug',
        ...summarizeSearchRequest(body)
      });
    }

    const analyticsMeta = normalizeSearchAnalyticsMeta(body);
    const searchStartedAt = Date.now();
    const wantsPagedResponse = Boolean(body && (
      body.includeTotalCount ||
      body.includeFacets ||
      typeof body.skip !== 'undefined' ||
      body.responseShape === 'paged'
    ));
    try {
      const payload = wantsPagedResponse
        ? await searchDocumentsPaged(env, body, { ...context, log: (event) => logEvent(context, event) })
        : await searchDocuments(env, body, { ...context, log: (event) => logEvent(context, event) });
      if (debugRequests) {
        logEvent(context, {
          level: 'info',
          event: 'search-documents-response-debug',
          ...summarizeSearchResponse(payload)
        });
      }
      if (isPlainObject(payload)) {
        payload.analytics = {
          searchRequestId: analyticsMeta.searchRequestId,
          searchSessionId: analyticsMeta.searchSessionIdRaw
        };
      }
      if (shouldEmitSearchAnalyticsForRequest(body)) {
        emitAnalyticsEvent(
          context,
          buildSearchCompletedEvent(context, body, payload, analyticsMeta, Date.now() - searchStartedAt)
        );
      }
      return { statusCode: 200, payload };
    } catch (error) {
      if (shouldEmitSearchAnalyticsForRequest(body)) {
        emitAnalyticsEvent(
          context,
          buildSearchFailedEvent(context, body, analyticsMeta, Date.now() - searchStartedAt, error)
        );
      }
      throw error;
    }
  }

  if (url.pathname === '/api/search/semantic') {
    return { statusCode: 200, payload: await semanticSearch(env, body, { ...context, log: (event) => logEvent(context, event) }) };
  }

  if (url.pathname === '/api/search/people') {
    return { statusCode: 200, payload: await searchPeople(env, body, { ...context, log: (event) => logEvent(context, event) }) };
  }

  if (url.pathname === '/api/upload/extract-metadata') {
    return { statusCode: 200, payload: await extractUploadMetadata(env, body, { ...context, log: (event) => logEvent(context, event) }) };
  }

  if (url.pathname === '/api/upload/analyze-transcript') {
    return { statusCode: 200, payload: await analyzeMediaTranscript(env, body, { ...context, log: (event) => logEvent(context, event) }) };
  }

  if (url.pathname === '/api/upload/transcribe-media') {
    const upload = await readMultipartFile(request);
    return { statusCode: 200, payload: await transcribeUploadedMedia(env, upload, { ...context, log: (event) => logEvent(context, event) }) };
  }

  if (url.pathname === '/api/chat') {
    const validation = validateChatBody(body);
    if (!validation.ok) {
      return validation;
    }

    const rateLimit = chatRateLimitEnabled
      ? checkRateLimit(context.rateLimitKey)
      : { allowed: true, remaining: chatRateLimitMax, resetAt: Date.now() + chatRateLimitWindowMs };
    if (chatRateLimitEnabled && !rateLimit.allowed) {
      logEvent(context, {
        level: 'info',
        event: 'chat-rate-limited',
        path: url.pathname,
        resetAt: new Date(rateLimit.resetAt).toISOString()
      });
      return {
        statusCode: 429,
        payload: { error: 'Too many chat requests. Please wait and try again.' },
        headers: {
          'Retry-After': String(Math.ceil((rateLimit.resetAt - Date.now()) / 1000)),
          'X-RateLimit-Limit': String(chatRateLimitMax),
          'X-RateLimit-Remaining': '0',
          'X-RateLimit-Reset': String(Math.ceil(rateLimit.resetAt / 1000))
        }
      };
    }

    const conversation = resolveChatSession(context, body);
    const chatBody = {
      ...body,
      history: conversation.history.map((message) => ({
        role: message.role,
        content: message.content
      }))
    };
    const payload = await chat(env, chatBody, { ...context, log: (event) => logEvent(context, event) });
    rememberChatTurn(conversation, String(body.question || ''), payload.answer || '');
    return {
      statusCode: 200,
      payload: {
        ...payload,
        conversationId: conversation.conversationId,
        conversation: {
          id: conversation.conversationId,
          persisted: conversation.persisted
        }
      },
      headers: {
        'X-RateLimit-Enabled': String(chatRateLimitEnabled),
        'X-RateLimit-Limit': chatRateLimitEnabled ? String(chatRateLimitMax) : '',
        'X-RateLimit-Remaining': chatRateLimitEnabled ? String(rateLimit.remaining) : '',
        'X-RateLimit-Reset': chatRateLimitEnabled ? String(Math.ceil(rateLimit.resetAt / 1000)) : ''
      }
    };
  }

  return { statusCode: 404, payload: { error: 'Not found' } };
};

const requestListener = async (request, response) => {
  const startedAt = Date.now();
  const origin = request.headers.origin;
  const context = buildRequestContext(request);

  if (request.method === 'OPTIONS') {
    if (origin && !resolveCorsOrigin(origin)) {
      writeJson(response, 403, { error: 'Origin is not allowed.' }, origin, context);
      return;
    }
    writeJson(response, 204, {}, origin, context);
    return;
  }

  try {
    if (origin && !resolveCorsOrigin(origin)) {
      writeJson(response, 403, { error: 'Origin is not allowed.' }, origin, context);
      return;
    }
    const result = await routeRequest(request, context);
    if (Object.prototype.hasOwnProperty.call(result, 'textPayload')) {
      writeText(response, result.statusCode, result.textPayload);
    } else {
      writeJson(response, result.statusCode, result.payload, origin, context, result.headers || {});
    }
    logEvent(context, {
      level: 'info',
      event: 'http-request',
      method: request.method,
      path: request.url,
      status: result.statusCode,
      durationMs: Date.now() - startedAt,
      authenticated: Boolean(context.principal)
    });
  } catch (error) {
    const statusCode = Number(error.statusCode || 500);
    writeJson(response, statusCode, { error: error.message || 'Unhandled local API error' }, origin, context);
    logEvent(context, {
      level: 'error',
      event: 'http-request-failed',
      method: request.method,
      path: request.url,
      status: statusCode,
      durationMs: Date.now() - startedAt,
      error: error.message || String(error)
    });
  }
};

const localHttpsOptions = localHttpsEnabled && fs.existsSync(localHttpsKeyPath) && fs.existsSync(localHttpsCertPath)
  ? {
      key: fs.readFileSync(localHttpsKeyPath),
      cert: fs.readFileSync(localHttpsCertPath)
    }
  : undefined;

const server = localHttpsOptions
  ? https.createServer(localHttpsOptions, requestListener)
  : http.createServer(requestListener);

server.listen(port, () => {
  console.log(`Using env: ${absolutePath}`);
  console.log(`Allowed origins: ${allowedOrigins.length > 0 ? allowedOrigins.join(', ') : 'local-only/default'}`);
  console.log(`Easy Auth header required: ${requireEasyAuth}`);
  console.log(`Debug request logging: ${debugRequests ? 'enabled' : 'disabled'}`);
  console.log(`Normalized facets: ${normalizedFacetsEnabled ? 'enabled' : 'disabled'}`);
  console.log(`Canonical department facets: ${canonicalDepartmentFacetsEnabled ? 'enabled' : 'disabled'}`);
  console.log(`Chat rate limit: ${chatRateLimitEnabled ? `${chatRateLimitMax} requests per ${chatRateLimitWindowMs}ms` : 'disabled'}`);
  console.log(`Chat memory enabled: ${chatMemoryEnabled}`);
  console.log(`Chat memory path: ${chatMemoryPath}`);
  console.log(`Sync worker enabled: ${syncWorkerEnabled}`);
  console.log(`Sync timer enabled: ${syncTimerEnabled}`);
  console.log(`Sync run timeout: ${syncRunTimeoutMs}ms`);
  console.log(`Sync stale guard: ${syncStaleRunMs}ms`);
  console.log(`Sync timer max changes: ${syncTimerMaxChanges}`);
  console.log(`Sync manual max changes: ${syncManualMaxChanges}`);
  console.log(`Sync admin token required: ${syncAdminTokenRequired}`);
  console.log(`Sync inventory path: ${syncInventoryPath}`);
  console.log(`Local HTTPS enabled: ${Boolean(localHttpsOptions)}`);
  console.log(`iKnowledge search API listening at ${localHttpsOptions ? 'https' : 'http'}://localhost:${port}`);

  if (syncWorkerEnabled && syncTimerEnabled) {
    const scheduleNext = () => {
      syncState.nextTimerAt = new Date(Date.now() + syncWorkerIntervalSeconds * 1000).toISOString();
    };
    scheduleNext();
    setInterval(() => {
      scheduleNext();
      triggerSync('timer_fallback');
    }, syncWorkerIntervalSeconds * 1000);
  }
});
