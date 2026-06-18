import { AadHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  getKnowledgeSearchApiAadResource,
  getKnowledgeSearchApiBaseUrl
} from './SearchConfig';

export interface IKnowledgeSearchRequest {
  query: string;
  top?: number;
  skip?: number;
  includeTotalCount?: boolean;
  includeFacets?: boolean;
  responseShape?: 'array' | 'paged';
  sort?: 'relevance' | 'newest' | 'oldest';
  filters?: Record<string, string | string[] | undefined>;
  searchSessionId?: string;
  analytics?: {
    recordWildcardSearch?: boolean;
  };
}

export interface IKnowledgeSearchFacetValue {
  value: string;
  count: number;
}

export type IKnowledgeSearchFacets = Record<string, IKnowledgeSearchFacetValue[] | undefined>;

export interface IKnowledgeSearchPageInfo {
  top: number;
  skip: number;
  returned: number;
  totalAvailable: number;
  resultWindow: number;
  capped: boolean;
  hasMore: boolean;
  nextSkip?: number;
}

export interface IKnowledgeSearchResult {
  id: string;
  documentId?: string;
  listItemId?: string;
  driveItemId?: string;
  fileUniqueId?: string;
  uniqueId?: string;
  fileRef?: string;
  serverRelativeUrl?: string;
  title: string;
  fileName?: string;
  fileExtension?: string;
  webUrl?: string;
  status?: string;
  publishedDate?: string;
  chunkText?: string;
  score?: number;
  rerankerScore?: number;
  contentPreview?: string;
  description?: string;
  documentType?: string;
  businessUnit?: string;
  bu?: string;
  department?: string;
  diseaseArea?: string;
  therapyArea?: string;
  client?: string;
  geography?: string;
  region?: string;
  authors?: string[] | string;
  createdDateTime?: string;
  created?: string;
  lastModifiedDateTime?: string;
  modified?: string;
  contentRefreshDate?: string;
  views?: number;
  likes?: number;
  comments?: number;
  downloads?: number;
}

export interface IKnowledgeSearchDocumentsResponse {
  results: IKnowledgeSearchResult[];
  totalCount: number;
  facets?: IKnowledgeSearchFacets;
  page?: IKnowledgeSearchPageInfo;
  analytics?: {
    searchRequestId?: string;
    searchSessionId?: string;
  };
}

export interface IKnowledgeSearchActionEvent {
  action: 'view' | 'download' | 'open_reference';
  searchRequestId?: string;
  searchSessionId?: string;
  documentId?: string;
  listItemId?: string;
  documentTitle?: string;
  rank?: number;
  timeSinceSearchMs?: number;
  sourceSurface?: string;
  frontendVersion?: string;
}

export interface IKnowledgePeopleResult {
  id: string;
  personName: string;
  email?: string;
  contacts?: string;
  allEmails?: string;
  allText?: string;
  description?: string;
  sectionTitles?: string;
  serviceLine?: string;
  role?: string;
  team?: string;
  listItemUrl?: string;
  score?: number;
}

export interface IKnowledgeChatCitation {
  number: number;
  kind: 'document' | 'people';
  id?: string;
  documentId?: string;
  listItemId?: string;
  title: string;
  url?: string;
  score?: number;
}

export interface IKnowledgeChatResponse {
  answer: string;
  citations: IKnowledgeChatCitation[];
  conversationId?: string;
  conversation?: {
    id: string;
    persisted: boolean;
  };
  usage?: {
    promptTokens?: number;
    completionTokens?: number;
    totalTokens?: number;
    deployment?: string;
  };
}

export interface IKnowledgeSyncResponse {
  ok?: boolean;
  accepted?: boolean;
  queued?: boolean;
  reason?: string;
  trigger?: {
    accepted?: boolean;
    queued?: boolean;
    reason?: string;
  };
}

export interface IUploadTranscriptionResponse {
  transcript: Array<{ time: number; text: string }>;
  transcriptText: string;
}

export interface IUploadMediaAnalysisResponse {
  abstract: string;
  businessUnit: string;
  department: string;
  documentType: string;
  client: string;
  geography: string;
  therapyArea: string;
  diseaseArea: string;
  title: string;
  sensitiveTerms?: string;
}

const DEFAULT_TIMEOUT_MS = 45000;
const UPLOAD_AI_TIMEOUT_MS = 180000;

const readJsonResponse = async <T>(response: Response, path: string, requestUrl: string): Promise<T> => {
  const text = await response.text();
  const contentType = response.headers.get('content-type') || '';

  if (!response.ok) {
    throw new Error(`Knowledge search API ${path} failed ${response.status} at ${requestUrl}: ${text.slice(0, 500)}`);
  }

  if (!/application\/json/i.test(contentType) && /^\s*</.test(text)) {
    throw new Error(`Knowledge search API ${path} returned HTML instead of JSON from ${requestUrl}. Check the backend URL and CORS.`);
  }

  try {
    return JSON.parse(text) as T;
  } catch (error) {
    throw new Error(`Knowledge search API ${path} returned invalid JSON from ${requestUrl}: ${text.slice(0, 300)}`);
  }
};

export class KnowledgeSearchApiClient {
  private readonly configuredBaseUrl: string;
  private readonly configuredAadResource: string;
  private readonly context?: WebPartContext;

  constructor(
    baseUrl: string = '',
    context?: WebPartContext,
    aadResource: string = ''
  ) {
    this.configuredBaseUrl = String(baseUrl || '').replace(/\/$/, '');
    this.configuredAadResource = String(aadResource || '').replace(/\/$/, '');
    this.context = context;
  }

  public isConfigured(): boolean {
    return Boolean(this.resolveBaseUrl());
  }

  public async searchDocuments(request: IKnowledgeSearchRequest): Promise<IKnowledgeSearchDocumentsResponse> {
    const requestWithSession: IKnowledgeSearchRequest = {
      ...request,
      searchSessionId: request.searchSessionId || this.getSearchSessionId()
    };
    const response = await this.post<IKnowledgeSearchResult[] | IKnowledgeSearchDocumentsResponse>('/api/search/documents', requestWithSession);
    if (Array.isArray(response)) {
      return {
        results: response,
        totalCount: response.length,
        facets: {},
        page: {
          top: request.top || response.length,
          skip: request.skip || 0,
          returned: response.length,
          totalAvailable: response.length,
          resultWindow: response.length,
          capped: false,
          hasMore: false
        }
      };
    }

    return {
      results: Array.isArray(response.results) ? response.results : [],
      totalCount: Number(response.totalCount || response.results?.length || 0),
      facets: response.facets || {},
      page: response.page,
      analytics: response.analytics
    };
  }

  public async trackSearchAction(event: IKnowledgeSearchActionEvent): Promise<void> {
    const analyticsEvent: IKnowledgeSearchActionEvent = {
      ...event,
      searchSessionId: event.searchSessionId || this.getSearchSessionId(),
      sourceSurface: event.sourceSurface || 'search'
    };

    try {
      await this.post<{ accepted?: number; enabled?: boolean }>('/api/analytics/events', analyticsEvent, 10000);
    } catch (error) {
      // Analytics must never interrupt the user's search or document workflow.
      console.warn('Search analytics event failed quietly:', error);
    }
  }

  public semanticSearch(request: IKnowledgeSearchRequest): Promise<IKnowledgeSearchResult[]> {
    return this.post<IKnowledgeSearchResult[]>('/api/search/semantic', request);
  }

  public searchPeople(request: IKnowledgeSearchRequest): Promise<IKnowledgePeopleResult[]> {
    return this.post<IKnowledgePeopleResult[]>('/api/search/people', request);
  }

  public chat(
    question: string,
    history: Array<{ role: string; content: string }> = [],
    conversationId?: string,
    conversationSummary?: string
  ): Promise<IKnowledgeChatResponse> {
    return this.post<IKnowledgeChatResponse>('/api/chat', { question, history, conversationId, conversationSummary });
  }

  public triggerSync(reason: string = 'spfx-manual'): Promise<IKnowledgeSyncResponse> {
    return this.post<IKnowledgeSyncResponse>('/api/sync/run', { reason });
  }

  public extractUploadMetadata(
    documentText: string,
    docTypeTermsSection?: string,
    clientTermsSection?: string
  ): Promise<Record<string, any>> {
    return this.post<Record<string, any>>('/api/upload/extract-metadata', {
      documentText,
      docTypeTermsSection,
      clientTermsSection
    }, UPLOAD_AI_TIMEOUT_MS);
  }

  public analyzeMediaTranscript(
    fileName: string,
    transcriptText: string,
    docTypeTermsSection?: string,
    clientTermsSection?: string
  ): Promise<IUploadMediaAnalysisResponse> {
    return this.post<IUploadMediaAnalysisResponse>('/api/upload/analyze-transcript', {
      fileName,
      transcriptText,
      docTypeTermsSection,
      clientTermsSection
    }, UPLOAD_AI_TIMEOUT_MS);
  }

  public transcribeMedia(file: File): Promise<IUploadTranscriptionResponse> {
    const formData = new FormData();
    formData.append('file', file, file.name);
    return this.postForm<IUploadTranscriptionResponse>('/api/upload/transcribe-media', formData, UPLOAD_AI_TIMEOUT_MS);
  }

  private async post<T>(path: string, body: unknown, timeoutMs: number = DEFAULT_TIMEOUT_MS): Promise<T> {
    const baseUrl = this.resolveBaseUrl();
    if (!baseUrl) {
      throw new Error('Knowledge search API base URL is not configured.');
    }

    const requestUrl = `${baseUrl}${path}`;
    const requestBody = JSON.stringify(body);
    const correlationId = this.createCorrelationId();
    const headers = {
      'Content-Type': 'application/json',
      'x-correlation-id': correlationId
    };
    const response = this.context?.aadHttpClientFactory
      && this.resolveAadResource()
      ? await this.postWithAadHttpClient(requestUrl, requestBody, headers, timeoutMs)
      : await this.postWithFetch(requestUrl, requestBody, headers, timeoutMs);

    return readJsonResponse<T>(response, path, requestUrl);
  }

  private async postForm<T>(path: string, formData: FormData, timeoutMs: number = DEFAULT_TIMEOUT_MS): Promise<T> {
    const baseUrl = this.resolveBaseUrl();
    if (!baseUrl) {
      throw new Error('Knowledge search API base URL is not configured.');
    }

    const requestUrl = `${baseUrl}${path}`;
    const headers = {
      'x-correlation-id': this.createCorrelationId()
    };
    const response = this.context?.aadHttpClientFactory
      && this.resolveAadResource()
      ? await this.postWithAadHttpClient(requestUrl, formData, headers, timeoutMs)
      : await this.postWithFetch(requestUrl, formData, headers, timeoutMs);

    return readJsonResponse<T>(response, path, requestUrl);
  }

  private async postWithAadHttpClient(
    requestUrl: string,
    requestBody: string | FormData,
    headers: Record<string, string>,
    timeoutMs: number = DEFAULT_TIMEOUT_MS
  ): Promise<Response> {
    const aadResource = this.resolveAadResource();
    if (!aadResource || !this.context?.aadHttpClientFactory) {
      return this.postWithFetch(requestUrl, requestBody, headers, timeoutMs);
    }

    const aadClient = await this.context.aadHttpClientFactory.getClient(aadResource);
    return aadClient.post(requestUrl, AadHttpClient.configurations.v1, {
      headers,
      body: requestBody as any
    } as any) as unknown as Promise<Response>;
  }

  private async postWithFetch(
    requestUrl: string,
    requestBody: string | FormData,
    headers: Record<string, string>,
    timeoutMs: number = DEFAULT_TIMEOUT_MS
  ): Promise<Response> {
    const controller = typeof AbortController !== 'undefined' ? new AbortController() : undefined;
    const timeoutId = controller
      ? window.setTimeout(() => controller.abort(), timeoutMs)
      : undefined;

    return fetch(requestUrl, {
      method: 'POST',
      // Easy Auth is not enabled yet. Do not send SharePoint/browser cookies to
      // the hosted backend, otherwise browser CORS requires
      // Access-Control-Allow-Credentials=true and blocks valid API responses.
      // When Entra auth is configured, set knowledgeSearchApiAadResource and the
      // AadHttpClient path above will be used instead.
      credentials: 'omit',
      signal: controller?.signal,
      headers,
      body: requestBody
    }).finally(() => {
      if (timeoutId) {
        window.clearTimeout(timeoutId);
      }
    });
  }

  private resolveBaseUrl(): string {
    return this.configuredBaseUrl || getKnowledgeSearchApiBaseUrl();
  }

  private resolveAadResource(): string {
    return this.configuredAadResource || getKnowledgeSearchApiAadResource();
  }

  private createCorrelationId(): string {
    const cryptoApi = typeof window !== 'undefined' ? window.crypto : undefined;
    if (cryptoApi && typeof cryptoApi.getRandomValues === 'function') {
      const bytes = new Uint8Array(16);
      cryptoApi.getRandomValues(bytes);
      return Array.from(bytes).map((value) => value.toString(16).padStart(2, '0')).join('');
    }
    return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
  }

  private getSearchSessionId(): string {
    const key = 'iknowledge_search_session_id';
    const cryptoApi = typeof window !== 'undefined' ? window.crypto : undefined;
    const createId = (): string => {
      if (cryptoApi && typeof cryptoApi.getRandomValues === 'function') {
        const bytes = new Uint8Array(16);
        cryptoApi.getRandomValues(bytes);
        return Array.from(bytes).map((value) => value.toString(16).padStart(2, '0')).join('');
      }
      return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2)}`;
    };

    try {
      const existing = window.sessionStorage.getItem(key);
      if (existing) {
        return existing;
      }
      const next = createId();
      window.sessionStorage.setItem(key, next);
      return next;
    } catch {
      return createId();
    }
  }
}
