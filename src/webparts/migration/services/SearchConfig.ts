// src/webparts/migration/services/SearchConfig.ts
//
// Search and AI secrets must not be compiled into SPFx. SPFx runs in the
// browser, so every bundled value is visible to users. This module only reads
// optional non-secret runtime configuration for legacy code paths. Backend
// search, chat, upload AI, and sync calls use the production API below.

export interface IKnowledgeSearchRuntimeConfig {
  azureOpenAiEndpoint?: string;
  azureOpenAiApiKey?: string;
  azureOpenAiDeployment?: string;
  azureOpenAiApiVersion?: string;
  azureOpenAiEmbeddingModel?: string;
  azureOpenAiEmbeddingApiVersion?: string;
  azureOpenAiWhisperEndpoint?: string;
  azureOpenAiWhisperApiKey?: string;
  azureOpenAiWhisperDeployment?: string;
  azureSearchEndpoint?: string;
  azureSearchIndex?: string;
  azureSearchApiVersion?: string;
  azureSearchKey?: string;
  azureSearchSemanticConfig?: string;
  azureSearchSuggesterName?: string;
  knowledgeSearchApiBaseUrl?: string;
  knowledgeSearchApiAadResource?: string;
  knowledgeBaseUrl?: string;
}

declare global {
  interface Window {
    __IKNOWLEDGE_SEARCH_CONFIG__?: IKnowledgeSearchRuntimeConfig;
  }
}

// Default production SharePoint pages to the hosted backend.
// SharePoint Workbench can be pointed at local or production via the switch
// below, URL query params, or runtime config. Local backend testing remains
// available on localhost pages, or with an explicit query override:
//   ?ikSearchApi=https://localhost:7072
// Stored localhost overrides are ignored on normal SharePoint-hosted pages so an
// old test setting cannot silently send upload/search/chat calls to a stopped
// local API while we are validating the deployed backend.
// When Entra/Easy Auth is enabled, also provide knowledgeSearchApiAadResource
// through the same runtime config path instead of hard-coding secrets here.
const PRODUCTION_SEARCH_API_BASE_URL = 'https://iknowledge-search-api-prod-ecajb0e6faddgfgb.centralus-01.azurewebsites.net';
const LOCAL_SEARCH_API_BASE_URL = 'https://localhost:7072';

// One-line backend switch for SharePoint Workbench testing.
// - 'local': Workbench uses https://localhost:7072
// - 'production': Workbench uses the deployed Azure Web App backend
// Normal non-Workbench SharePoint pages always default to production unless an
// explicit query/runtime config override is supplied.
let WORKBENCH_SEARCH_BACKEND_MODE: 'local' | 'production' = 'local';

const shouldUseLocalBackendByDefault = (): boolean => {
  if (typeof window === 'undefined') {
    return false;
  }

  return window.location.hostname === 'localhost' ||
    (isSharePointWorkbenchPage() && WORKBENCH_SEARCH_BACKEND_MODE === 'local');
};

const isSharePointHostedPage = (): boolean => {
  if (typeof window === 'undefined') {
    return false;
  }

  return /\.sharepoint\.com$/i.test(window.location.hostname);
};

const isSharePointWorkbenchPage = (): boolean => {
  if (typeof window === 'undefined') {
    return false;
  }

  return /\/_layouts\/(?:15\/)?workbench\.aspx$/i.test(window.location.pathname);
};

const isLoopbackApiBaseUrl = (value: string | undefined): boolean =>
  /^https?:\/\/(?:localhost|127(?:\.\d{1,3}){3})(?::\d+)?(?:\/|$)/i.test(String(value || '').trim());

const shouldAcceptStoredApiBaseUrl = (value: string | undefined): boolean => {
  if (!value) {
    return false;
  }

  return !(isSharePointHostedPage() && !isSharePointWorkbenchPage() && isLoopbackApiBaseUrl(value));
};

const sanitizeStoredRuntimeConfig = (config: IKnowledgeSearchRuntimeConfig): IKnowledgeSearchRuntimeConfig => {
  if (!config.knowledgeSearchApiBaseUrl || shouldAcceptStoredApiBaseUrl(config.knowledgeSearchApiBaseUrl)) {
    return config;
  }

  const { knowledgeSearchApiBaseUrl, ...safeConfig } = config;
  return safeConfig;
};

const parseLocalStorageRuntimeConfig = (): IKnowledgeSearchRuntimeConfig => {
  if (typeof window === 'undefined') {
    return {};
  }

  const config: IKnowledgeSearchRuntimeConfig = {};
  const applyStorageConfig = (storage: Storage | undefined): void => {
    if (!storage) {
      return;
    }

    const rawConfig = storage.getItem('IKNOWLEDGE_SEARCH_CONFIG');
    if (rawConfig) {
      const parsed = sanitizeStoredRuntimeConfig(JSON.parse(rawConfig) as IKnowledgeSearchRuntimeConfig);
      if (parsed && typeof parsed === 'object') {
        Object.assign(config, parsed);
      }
    }

    const localApiBaseUrl = storage.getItem('IKNOWLEDGE_SEARCH_API_BASE_URL');
    if (shouldAcceptStoredApiBaseUrl(localApiBaseUrl)) {
      config.knowledgeSearchApiBaseUrl = localApiBaseUrl;
    }

    const localApiAadResource = storage.getItem('IKNOWLEDGE_SEARCH_API_AAD_RESOURCE');
    if (localApiAadResource) {
      config.knowledgeSearchApiAadResource = localApiAadResource;
    }
  };

  try {
    applyStorageConfig(window.localStorage);
    applyStorageConfig(window.sessionStorage);

    const params = new URLSearchParams(window.location.search);
    const backendModeFromQuery = (params.get('ikBackend') || params.get('knowledgeSearchBackend') || '').toLowerCase();
    if (backendModeFromQuery === 'local') {
      config.knowledgeSearchApiBaseUrl = LOCAL_SEARCH_API_BASE_URL;
    } else if (backendModeFromQuery === 'prod' || backendModeFromQuery === 'production') {
      config.knowledgeSearchApiBaseUrl = PRODUCTION_SEARCH_API_BASE_URL;
    }

    const apiFromQuery = params.get('ikSearchApi') || params.get('knowledgeSearchApiBaseUrl');
    if (apiFromQuery) {
      config.knowledgeSearchApiBaseUrl = apiFromQuery;
    }

    const aadResourceFromQuery = params.get('ikSearchApiResource') || params.get('knowledgeSearchApiAadResource');
    if (aadResourceFromQuery) {
      config.knowledgeSearchApiAadResource = aadResourceFromQuery;
    }

    if (!config.knowledgeSearchApiBaseUrl) {
      config.knowledgeSearchApiBaseUrl = shouldUseLocalBackendByDefault()
        ? LOCAL_SEARCH_API_BASE_URL
        : PRODUCTION_SEARCH_API_BASE_URL;
    }
  } catch {
    // Runtime configuration is optional. Ignore malformed local test values.
  }

  return config;
};

const runtimeConfig: IKnowledgeSearchRuntimeConfig =
  typeof window !== 'undefined' && window.__IKNOWLEDGE_SEARCH_CONFIG__
    ? { ...parseLocalStorageRuntimeConfig(), ...sanitizeStoredRuntimeConfig(window.__IKNOWLEDGE_SEARCH_CONFIG__) }
    : parseLocalStorageRuntimeConfig();

const readRuntimeValue = (key: keyof IKnowledgeSearchRuntimeConfig): string =>
  String(runtimeConfig[key] || '');

export const AZURE_OPENAI_ENDPOINT = readRuntimeValue('azureOpenAiEndpoint');
export const AZURE_OPENAI_API_KEY = readRuntimeValue('azureOpenAiApiKey');
export const AZURE_OPENAI_DEPLOYMENT = readRuntimeValue('azureOpenAiDeployment');
export const AZURE_OPENAI_API_VERSION = readRuntimeValue('azureOpenAiApiVersion');

export const AZURE_OPENAI_EMBEDDING_MODEL = readRuntimeValue('azureOpenAiEmbeddingModel');
export const AZURE_OPENAI_EMBEDDING_API_VERSION = readRuntimeValue('azureOpenAiEmbeddingApiVersion');

export const AZURE_OPENAI_WHISPER_ENDPOINT = readRuntimeValue('azureOpenAiWhisperEndpoint');
export const AZURE_OPENAI_WHISPER_API_KEY = readRuntimeValue('azureOpenAiWhisperApiKey');
export const AZURE_OPENAI_WHISPER_DEPLOYMENT = readRuntimeValue('azureOpenAiWhisperDeployment');

export const AZURE_SEARCH_ENDPOINT = readRuntimeValue('azureSearchEndpoint');
export const AZURE_SEARCH_INDEX = readRuntimeValue('azureSearchIndex');
export const AZURE_SEARCH_API_VERSION = readRuntimeValue('azureSearchApiVersion');
export const AZURE_SEARCH_KEY = readRuntimeValue('azureSearchKey');
export const AZURE_SEARCH_SEMANTIC_CONFIG = readRuntimeValue('azureSearchSemanticConfig');
export const AZURE_SEARCH_SUGGESTER_NAME = readRuntimeValue('azureSearchSuggesterName');

export const KNOWLEDGE_SEARCH_API_BASE_URL = readRuntimeValue('knowledgeSearchApiBaseUrl') || PRODUCTION_SEARCH_API_BASE_URL;
export const KNOWLEDGE_SEARCH_API_AAD_RESOURCE = readRuntimeValue('knowledgeSearchApiAadResource');
export const KNOWLEDGE_BASE_URL = readRuntimeValue('knowledgeBaseUrl');

let lastLoggedSearchApiBaseUrl = '';

export const getKnowledgeSearchApiBaseUrl = (): string => {
  const latestRuntimeConfig =
    typeof window !== 'undefined' && window.__IKNOWLEDGE_SEARCH_CONFIG__
      ? { ...parseLocalStorageRuntimeConfig(), ...sanitizeStoredRuntimeConfig(window.__IKNOWLEDGE_SEARCH_CONFIG__) }
      : parseLocalStorageRuntimeConfig();

  const baseUrl = String(latestRuntimeConfig.knowledgeSearchApiBaseUrl || KNOWLEDGE_SEARCH_API_BASE_URL || '').replace(/\/$/, '');
  if (typeof window !== 'undefined' && isSharePointWorkbenchPage() && baseUrl !== lastLoggedSearchApiBaseUrl) {
    lastLoggedSearchApiBaseUrl = baseUrl;
    console.info(`[iKnowledgeNext] Workbench search backend: ${baseUrl}`);
  }

  return baseUrl;
};

export const getKnowledgeSearchApiAadResource = (): string => {
  const latestRuntimeConfig =
    typeof window !== 'undefined' && window.__IKNOWLEDGE_SEARCH_CONFIG__
      ? { ...parseLocalStorageRuntimeConfig(), ...window.__IKNOWLEDGE_SEARCH_CONFIG__ }
      : parseLocalStorageRuntimeConfig();

  return String(latestRuntimeConfig.knowledgeSearchApiAadResource || KNOWLEDGE_SEARCH_API_AAD_RESOURCE || '').replace(/\/$/, '');
};
