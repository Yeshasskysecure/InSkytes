import { SearchEnv, toNumber } from '../runtime/env';
import { searchIndex } from '../runtime/azureSearchClient';
import { ChatCompletionResult, embedTextsInBatches, getChatCompletionResult } from '../runtime/openAiClient';
import { DocumentSearchFilters, buildDocumentFilter } from '../search/filters';
import { domainAcronymCoreQuery, normalizeDomainSearchPhrase } from '../search/domainNormalizations';
import { classifyChatIntent, isCredentialDisclosureQuery, shouldSuppressDocumentSearchQuery } from '../chat/chatPolicy';
import {
  GroundingSource,
  STRUCTURED_CHAT_RESPONSE_FORMAT,
  buildGroundedChatPrompt,
  buildNoEvidenceChatPrompt,
  buildStructuredGroundedChatPrompt,
  buildStructuredNoEvidenceChatPrompt
} from '../chat/chatPrompt';

export interface SearchExecutionContext {
  correlationId?: string;
  userHash?: string;
  log?: (event: Record<string, unknown>) => void;
}

export interface SearchRequest {
  query: string;
  top?: number;
  skip?: number;
  includeTotalCount?: boolean;
  includeFacets?: boolean;
  sort?: 'relevance' | 'newest' | 'oldest';
  filters?: Record<string, string | string[] | undefined>;
  analytics?: {
    recordWildcardSearch?: boolean;
  };
}

export interface ChatRequest {
  question: string;
  history?: Array<{ role: string; content: string }>;
  conversationSummary?: string;
}

export interface SearchHit {
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
  score?: number;
  rerankerScore?: number;
  contentWindowHits?: number;
  description?: string;
  contentPreview?: string;
  chunkText?: string;
  documentType?: string;
  documentTypeFilterValues?: string[];
  bu?: string;
  buFilterValues?: string[];
  businessUnit?: string;
  department?: string;
  departmentFilterValues?: string[];
  departmentFacetValues?: string[];
  diseaseArea?: string;
  diseaseAreaFilterValues?: string[];
  therapyArea?: string;
  therapyAreaFilterValues?: string[];
  client?: string;
  clientFilterValues?: string[];
  region?: string;
  regionFilterValues?: string[];
  geography?: string;
  authors?: string[];
  createdDateTime?: string;
  created?: string;
  lastModifiedDateTime?: string;
  modified?: string;
  publishedDate?: string;
  contentRefreshDate?: string;
}

export interface SearchFacetValue {
  value: string;
  count: number;
}

export interface SearchPageResponse {
  results: SearchHit[];
  totalCount: number;
  facets: Record<string, SearchFacetValue[]>;
  page: {
    top: number;
    skip: number;
    returned: number;
    totalAvailable: number;
    resultWindow: number;
    capped: boolean;
    hasMore: boolean;
    nextSkip?: number;
  };
}

const DOCUMENT_SELECT_FIELDS = [
  'id',
  'listItemId',
  'driveItemId',
  'fileUniqueId',
  'uniqueId',
  'fileRef',
  'serverRelativeUrl',
  'title',
  'fileName',
  'fileExtension',
  'webUrl',
  'status',
  'description',
  'contentPreview',
  'documentType',
  'bu',
  'department',
  'diseaseArea',
  'therapyArea',
  'client',
  'region',
  'authors',
  'createdDateTime',
  'created',
  'lastModifiedDateTime',
  'modified',
  'publishedDate',
  'contentRefreshDate'
].join(',');

const LEGACY_DOCUMENT_SELECT_FIELDS = [
  'id',
  'listItemId',
  'title',
  'fileName',
  'fileExtension',
  'webUrl',
  'status',
  'description',
  'contentPreview',
  'documentType',
  'bu',
  'department',
  'diseaseArea',
  'therapyArea',
  'client',
  'region',
  'authors'
].join(',');

const BASE_NORMALIZED_FILTER_SELECT_FIELDS = [
  'buFilterValues',
  'departmentFilterValues',
  'diseaseAreaFilterValues',
  'therapyAreaFilterValues',
  'clientFilterValues',
  'regionFilterValues',
  'documentTypeFilterValues'
];

const normalizedFacetsEnabled = (): boolean =>
  /^(1|true|yes)$/i.test(String((process as any).env?.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED || ''));

const canonicalDepartmentFacetsEnabled = (): boolean =>
  /^(1|true|yes)$/i.test(String((process as any).env?.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED || ''));

const normalizedFilterSelectFields = (): string =>
  [
    ...BASE_NORMALIZED_FILTER_SELECT_FIELDS,
    ...(canonicalDepartmentFacetsEnabled() ? ['departmentFacetValues'] : [])
  ].join(',');

const withNormalizedFilterSelectFields = (selectFields: string): string =>
  normalizedFacetsEnabled() ? `${selectFields},${normalizedFilterSelectFields()}` : selectFields;

export interface ChatCitation {
  number: number;
  kind: 'document' | 'people';
  id?: string;
  documentId?: string;
  listItemId?: string;
  title: string;
  url?: string;
  score?: number;
}

export interface ChatResponse {
  answer: string;
  citations: ChatCitation[];
  usage?: {
    promptTokens?: number;
    completionTokens?: number;
    totalTokens?: number;
    deployment?: string;
  };
  conversationId?: string;
  conversation?: {
    id: string;
    persisted: boolean;
  };
}

const normalizeTop = (env: SearchEnv, top?: number): number => {
  const requestedTop = Math.max(Number(top || env.AZURE_SEARCH_DEFAULT_TOP || 10), 1);
  const maxTop = Math.min(Math.max(toNumber(env.AZURE_SEARCH_MAX_TOP, 250), 1), 1000);
  return Math.min(requestedTop, maxTop);
};

const normalizePageTop = (env: SearchEnv, top?: number): number => {
  const requestedTop = Math.max(Number(top || env.AZURE_SEARCH_PAGE_TOP || 30), 1);
  const maxTop = Math.min(Math.max(toNumber(env.AZURE_SEARCH_MAX_PAGE_TOP, 30), 1), 1000);
  return Math.min(requestedTop, maxTop);
};

const normalizeSkip = (skip?: number): number =>
  Math.max(Number.isFinite(Number(skip)) ? Number(skip) : 0, 0);

const STOP_WORDS = new Set([
  'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'can', 'do', 'does', 'for', 'from',
  'give', 'how', 'in', 'is', 'it', 'me', 'of', 'on', 'or', 'show', 'that', 'the', 'this',
  'to', 'what', 'when', 'where', 'which', 'who', 'with', 'you',
  'all', 'any', 'some', 'please', 'find', 'list', 'search', 'documents', 'document',
  'docs', 'files', 'file', 'make', 'delete', 'remove',
  'under', 'his', 'her', 'their', 'name', 'names', 'about', 'tell', 'related'
]);

const meaningfulTokens = (value: string): string[] =>
  value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .split(/\s+/)
    .map((token) => token.trim())
    .filter((token) => token.length > 2 && !STOP_WORDS.has(token));

const lexicalOverlapCount = (query: string, text: string): number => {
  const haystack = ` ${text.toLowerCase().replace(/[^a-z0-9]+/g, ' ')} `;
  return meaningfulTokens(query).filter((token) => haystack.includes(` ${token} `)).length;
};

const editDistanceWithinOne = (left: string, right: string): boolean => {
  if (left === right) return true;
  if (Math.abs(left.length - right.length) > 1) return false;

  let edits = 0;
  let leftIndex = 0;
  let rightIndex = 0;

  while (leftIndex < left.length && rightIndex < right.length) {
    if (left[leftIndex] === right[rightIndex]) {
      leftIndex += 1;
      rightIndex += 1;
      continue;
    }

    edits += 1;
    if (edits > 1) return false;

    if (left.length > right.length) {
      leftIndex += 1;
    } else if (right.length > left.length) {
      rightIndex += 1;
    } else {
      leftIndex += 1;
      rightIndex += 1;
    }
  }

  if (leftIndex < left.length || rightIndex < right.length) edits += 1;
  return edits <= 1;
};

const fuzzyLexicalOverlapCount = (query: string, text: string): number => {
  const words = Array.from(new Set(
    String(text || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .split(/\s+/)
      .filter((word) => word.length >= 4)
  ));

  if (words.length === 0) return 0;

  return meaningfulTokens(query)
    .filter((token) => token.length >= 4)
    .filter((token) => words.some((word) => word[0] === token[0] && editDistanceWithinOne(token, word)))
    .length;
};

const GENERIC_AUTHOR_TOKENS = new Set(['iknowledge', 'admin', 'system', 'sharepoint']);
const GENERIC_DISTINCTIVE_TOKENS = new Set([...GENERIC_AUTHOR_TOKENS, 'indegene']);

const authorLexicalOverlapCount = (query: string, text: string): { overlap: number; tokenTotal: number } => {
  const haystack = ` ${text.toLowerCase().replace(/[^a-z0-9]+/g, ' ')} `;
  const tokens = meaningfulTokens(query).filter((token) => !GENERIC_AUTHOR_TOKENS.has(token));
  return {
    overlap: tokens.filter((token) => haystack.includes(` ${token} `)).length,
    tokenTotal: tokens.length
  };
};

const passesDistinctiveTokenGuard = (query: string, text: string): boolean => {
  const distinctiveTokens = meaningfulTokens(query)
    .filter((token) => !GENERIC_DISTINCTIVE_TOKENS.has(token))
    .filter((token) => token.length >= 8 || /\d/.test(token) || /(.)\1{2,}/.test(token));

  if (distinctiveTokens.length === 0) return true;

  const haystack = ` ${text.toLowerCase().replace(/[^a-z0-9]+/g, ' ')} `;
  return distinctiveTokens.some((token) => haystack.includes(` ${token} `));
};

const compactMeaningfulQuery = (query: string): string =>
  meaningfulTokens(query).join(' ');

const compactPrecisionQuery = (query: string): string => {
  const tokens = meaningfulTokens(query);
  if (tokens.length === 0) return query.trim();

  const priorityTokens = tokens.filter((token) =>
    token.length >= 6 ||
    /\d/.test(token) ||
    /(ai|ml|rwe|heor|crm|hcp|uat|pmo|aws|gcp)/i.test(token)
  );
  const selected = (priorityTokens.length >= 2 ? priorityTokens : tokens).slice(0, 12);
  return Array.from(new Set(selected)).join(' ');
};

const likelyPastedSearchText = (value: string): boolean => {
  const text = String(value || '').trim();
  if (!text) return false;

  return (
    likelyDenseStructuredContentSnippet(text) ||
    /[\r\n]/.test(text) ||
    text.length >= 160 ||
    (tokenCount(text) >= 18 && /[.,;:"'“”‘’–—-]/.test(text))
  );
};

const likelyDenseStructuredContentSnippet = (value: string): boolean => {
  const text = String(value || '').trim();
  if (text.length < 40) return false;
  if (rawContentIdentifierTokens(text).length > 0) return true;

  const punctuationSignals = (text.match(/[()[\];]/g) || []).length;
  const hyphenatedSignals = (text.match(/\b[a-z0-9]+(?:-[a-z0-9]+){2,}\b/gi) || []).length;
  const distinctiveTokenSignals = meaningfulTokens(text)
    .filter((token) => token.length >= 8 || /\d/.test(token))
    .length;

  return punctuationSignals + hyphenatedSignals >= 2 && distinctiveTokenSignals >= 2;
};

const buildPagedSearchText = (query: string): string => {
  if (isWildcardQuery(query) || isExplicitBooleanQuery(query)) return query;
  if (likelyPastedSearchText(query)) return compactPrecisionQuery(query);
  return normalizeDomainSearchPhrase(query);
};

const shouldRequireAllSearchTerms = (query: string): boolean =>
  !isWildcardQuery(query) &&
  !isExplicitBooleanQuery(query) &&
  !likelyPastedSearchText(query) &&
  meaningfulTokens(query).length >= 2;

const normalizeStrictSearchToken = (token: string): string => {
  const value = String(token || '').trim().toLowerCase();
  if (value.length <= 3) return value;

  if (/ies$/.test(value) && value.length > 4) {
    return `${value.slice(0, -3)}y`;
  }

  if (/(ches|shes|xes|zes|ses)$/.test(value) && value.length > 5) {
    return value.slice(0, -2);
  }

  if (/s$/.test(value) && !/(ss|us|is)$/.test(value) && value.length > 4) {
    return value.slice(0, -1);
  }

  return value;
};

const buildStrictTermSearchText = (query: string): string => {
  const tokens = meaningfulTokens(normalizeDomainSearchPhrase(query)).map(normalizeStrictSearchToken).filter(Boolean);
  return Array.from(new Set(tokens)).join(' ') || query.trim();
};


const cleanPersonNameCandidate = (value: string): string => {
  const stopAt = value
    .replace(/[?!.]/g, ' ')
    .split(/\b(?:and|which|what|where|when|why|how|docs?|documents?|files?|under|name|dept|department|role|do|does)\b/i)[0] || '';

  return stopAt
    .replace(/[^a-z0-9\s.'-]/gi, ' ')
    .replace(/\b(?:mr|mrs|ms|dr)\.?\b/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter((token) => token.length > 1)
    .slice(0, 4)
    .join(' ');
};

const extractPersonNameQuery = (query: string): string => {
  const value = String(query || '').trim();
  const patterns = [
    /\bwho\s+is\s+(.+?)$/i,
    /\bdocuments?\s+(?:by|from|under|for)\s+(.+?)$/i,
    /\bdocs?\s+(?:by|from|under|for)\s+(.+?)$/i,
    /\b(?:author|authors|uploaded by|created by|owner|owners)\s+(.+?)$/i
  ];

  for (const pattern of patterns) {
    const match = value.match(pattern);
    if (!match) continue;
    const candidate = cleanPersonNameCandidate(match[1]);
    if (meaningfulTokens(candidate).length >= 2) return candidate;
  }

  return '';
};

const isWildcardQuery = (query: string): boolean => !query.trim() || query.trim() === '*';

const tokenCount = (value: string): number =>
  value.split(/\s+/).map((token) => token.trim()).filter(Boolean).length;

const likelyPastedDocumentText = (value: string): boolean => {
  const normalized = String(value || '').toLowerCase();
  return (
    tokenCount(value) >= 12 ||
    value.length >= 120 ||
    /\b(contents|definitions|purpose and scope|termination|confidentiality|intellectual property|governing law)\b/.test(normalized)
  );
};

const likelyCompactBodyPhraseQuery = (value: string): boolean => {
  const text = String(value || '').trim();
  if (!text || isWildcardQuery(text) || isExplicitBooleanQuery(text)) return false;
  if (looksLikeSpecificDocumentQuery(text)) return false;

  const tokens = meaningfulTokens(text);
  if (tokens.length < 4 || tokens.length > 5) return false;

  const distinctiveTokens = tokens.filter((token) => token.length >= 9);
  const genericTitleTokens = tokens.filter((token) => TITLE_FILENAME_GENERIC_PROMOTION_TOKENS.has(token));
  return distinctiveTokens.length >= 2 && genericTitleTokens.length === 0;
};

const isExplicitBooleanQuery = (query: string): boolean => {
  const value = String(query || '').trim();
  if (/(^|\s)[+-][^\s]+/.test(value)) return true;
  const operators = value.match(/\b(?:AND|OR|NOT)\b/g) || [];
  if (operators.length === 0) return false;
  const tokenTotal = meaningfulTokens(value).length;
  return tokenTotal <= 10 || /["()]/.test(value);
};

const documentEvidenceText = (item: SearchHit): string => [
  item.listItemId,
  item.title,
  item.fileName,
  item.description,
  item.contentPreview,
  item.chunkText,
  item.documentType,
  item.bu,
  item.department,
  item.client,
  item.region,
  item.therapyArea,
  item.diseaseArea,
  ...(item.authors || [])
].filter(Boolean).join(' ');

const comparableText = (value: unknown): string =>
  String(value || '')
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\.[a-z0-9]{2,6}$/i, '')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');

const identifierTokens = (query: string): string[] =>
  Array.from(new Set((query.match(/\b\d{4,}\b/g) || []).map((token) => token.trim()).filter(Boolean)));

const looksLikeSpecificDocumentQuery = (query: string): boolean =>
  identifierTokens(query).length > 0 ||
  /\b[^\s"'<>]+?\.(?:pdf|docx?|pptx?|xlsx?|mp4|mp3|msg|png|jpe?g)\b/i.test(query) ||
  /[_/\\]/.test(query);

const shouldAvoidRelaxedNoResultFallback = (query: string): boolean =>
  identifierTokens(query).length > 0 ||
  meaningfulTokens(query).some((token) => token.length >= 5 && /(.)\1{2,}/.test(token));

const isLowInformationDocumentSearch = (query: string): boolean => {
  if (isWildcardQuery(query)) return false;
  if (identifierTokens(query).length > 0) return false;
  if (/\b(doc(?:u?e?m?e?n?t?s?)?|documents?|docs?|files?|source|sources|policy|sop|certificate|training|manual|deck|proposal|asset|pptx?|pdf|docx?|xlsx?|mp4|mp3|video|recording|client|bu|business unit|department|therapy|disease|geography|region)\b/i.test(query)) {
    return false;
  }

  const tokens = query
    .replace(/[^a-z0-9\s]/gi, ' ')
    .split(/\s+/)
    .map((token) => token.trim())
    .filter(Boolean);

  const uniqueAlphaNumericChars = new Set(tokens.join('').split('')).size;

  return tokens.length >= 2
    && tokens.length <= 3
    && tokens.every((token) => token.length <= 3)
    && uniqueAlphaNumericChars <= 3
    && !/^(hi|hello|hey|thanks|thank you|ok|okay)\b/i.test(query);
};

const exactDocumentMatchScore = (query: string, item: SearchHit): number => {
  if (isWildcardQuery(query)) return 0;

  const queryText = comparableText(query);
  const titleText = comparableText(item.title);
  const fileText = comparableText(item.fileName);
  const evidenceText = comparableText(documentEvidenceText(item));
  const queryTokens = meaningfulTokens(query);
  const evidenceForTokenMatch = ` ${evidenceText} `;
  const ids = identifierTokens(query);
  let score = 0;

  ids.forEach((id) => {
    if (String(item.listItemId || '') === id) score += 120;
    if (fileText.includes(id)) score += 90;
    if (titleText.includes(id)) score += 50;
  });

  if (queryText && titleText && queryText === titleText) score += 100;
  if (queryText && fileText && queryText === fileText) score += 100;
  if (queryText && titleText && queryText.includes(titleText) && meaningfulTokens(String(item.title || '')).length >= 3) {
    score += 120;
  }
  if (queryText && fileText && fileText.includes(queryText)) score += 70;
  if (queryText && titleText && titleText.includes(queryText) && queryTokens.length >= 3) score += 60;

  const titleOverlap = lexicalOverlapCount(query, [item.title, item.fileName].filter(Boolean).join(' '));
  const specificTitleTokens = queryTokens.filter((token) => !TITLE_FILENAME_GENERIC_PROMOTION_TOKENS.has(token));
  const titleEvidence = [item.title, item.fileName].filter(Boolean).join(' ');
  const specificTitleOverlap = specificTitleTokens
    .filter((token) => lexicalOverlapCount(token, titleEvidence) > 0)
    .length;
  const specificEvidenceOverlap = specificTitleTokens
    .filter((token) => evidenceForTokenMatch.includes(` ${comparableText(token)} `))
    .length;
  const leadingSpecificToken = specificTitleTokens[0];
  if (leadingSpecificToken) {
    const leadingToken = comparableText(leadingSpecificToken);
    if ((titleText && ` ${titleText} `.includes(` ${leadingToken} `)) || (fileText && ` ${fileText} `.includes(` ${leadingToken} `))) {
      score += 140;
    } else if (evidenceForTokenMatch.includes(` ${leadingToken} `)) {
      score += 70;
    }
  }
  const { overlap: authorOverlap, tokenTotal: authorQueryTokenTotal } = authorLexicalOverlapCount(query, (item.authors || []).join(' '));
  const hasAuthorIntent = /\b(author|authors|created by|uploaded by|owner|owners)\b/i.test(query);
  if (queryTokens.length >= 2 && titleOverlap >= 2) score += 85;
  if (queryTokens.length >= 3 && titleOverlap >= Math.ceil(queryTokens.length * 0.75)) score += 90;
  if (specificTitleOverlap > 0) score += specificTitleOverlap * 70;
  if (specificEvidenceOverlap > 0) score += specificEvidenceOverlap * 35;
  if (specificTitleTokens.length >= 2 && specificEvidenceOverlap >= Math.min(3, specificTitleTokens.length)) score += 90;
  if (specificTitleTokens.length >= 2 && specificTitleOverlap >= specificTitleTokens.length) score += 120;
  if (hasAuthorIntent && authorOverlap >= 1) score += 85;
  if (authorQueryTokenTotal >= 2 && authorOverlap >= Math.min(2, authorQueryTokenTotal)) score += 85;
  if (authorQueryTokenTotal >= 2 && authorOverlap >= authorQueryTokenTotal) score += 120;
  if (looksLikeSpecificDocumentQuery(query) && lexicalOverlapCount(query, evidenceText) >= 1) score += 20;

  return score;
};

const shouldPromoteExactDocument = (query: string, item: SearchHit, minScore = 80): boolean =>
  exactDocumentMatchScore(query, item) >= minScore;


const defaultMinimumLexicalMatches = (query: string): number => {
  const tokenTotal = meaningfulTokens(query).length;
  if (tokenTotal <= 0) return 0;
  if (tokenTotal <= 2) return tokenTotal;
  if (tokenTotal <= 4) return 2;
  if (tokenTotal <= 8) return 3;
  return Math.max(3, Math.ceil(tokenTotal * 0.25));
};

const passesDocumentRelevanceGuard = (env: SearchEnv, query: string, item: SearchHit): boolean => {
  if (isWildcardQuery(query)) return true;
  if (isCredentialDisclosureQuery(query)) return false;

  const tokens = meaningfulTokens(query);
  if (tokens.length === 0) return false;

  const evidenceText = documentEvidenceText(item);
  if (!passesDistinctiveTokenGuard(query, evidenceText)) return false;

  const overlap = Math.max(lexicalOverlapCount(query, evidenceText), fuzzyLexicalOverlapCount(query, evidenceText));
  const rerankerScore = item.rerankerScore ?? 0;
  const semanticMinRerankerScore = toNumber(env.AZURE_SEARCH_SEMANTIC_MIN_RERANKER_SCORE, 2);
  const semanticStrongRerankerScore = toNumber(env.AZURE_SEARCH_SEMANTIC_STRONG_RERANKER_SCORE, 3);
  const defaultGenericMinLexicalMatches = defaultMinimumLexicalMatches(query);
  const genericMinLexicalMatches = Math.max(
    toNumber(env.AZURE_SEARCH_GENERIC_MIN_LEXICAL_MATCHES, defaultGenericMinLexicalMatches),
    defaultGenericMinLexicalMatches
  );
  const semanticMinLexicalMatches = toNumber(env.AZURE_SEARCH_SEMANTIC_MIN_LEXICAL_MATCHES, 1);

  if (typeof item.rerankerScore === 'number') {
    return (
      (rerankerScore >= semanticMinRerankerScore && overlap >= semanticMinLexicalMatches) ||
      rerankerScore >= semanticStrongRerankerScore
    );
  }

  return overlap >= genericMinLexicalMatches;
};

const peopleEvidenceText = (item: any): string => [
  item.personName,
  item.email,
  item.contacts,
  item.allEmails,
  item.allText,
  item.description,
  item.sectionTitles,
  item.serviceLine,
  item.role,
  item.team,
  item.bu,
  item.region
].filter(Boolean).join(' ');

const focusedSnippet = (query: string, text: unknown, maxLength: number): string => {
  const value = String(text || '').replace(/\s+/g, ' ').trim();
  if (!value || value.length <= maxLength) return value;
  const lower = value.toLowerCase();
  const positions = meaningfulTokens(query)
    .map((token) => lower.indexOf(token.toLowerCase()))
    .filter((position) => position >= 0);
  const anchor = positions.length > 0 ? Math.min(...positions) : 0;
  const start = Math.max(0, anchor - Math.floor(maxLength * 0.35));
  return value.slice(start, start + maxLength).trim();
};

const passesPeopleRelevanceGuard = (query: string, item: any): boolean => {
  if (isWildcardQuery(query)) return true;
  if (isCredentialDisclosureQuery(query)) return false;

  const overlap = lexicalOverlapCount(query, peopleEvidenceText(item));
  const tokenTotal = meaningfulTokens(query).length;

  if (tokenTotal <= 1) return overlap >= 1;
  return overlap >= Math.min(2, tokenTotal);
};

const stripBracketCitations = (value: string): string =>
  value
    .replace(/\s*\[(?:\d+|[1-9]\d*\s*[-–]\s*[1-9]\d*)(?:\s*,\s*(?:\d+|[1-9]\d*\s*[-–]\s*[1-9]\d*))*\]/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

const cleanChatAnswer = (value: string): string =>
  String(value || '')
    .replace(/who['’]s\s+who/gi, 'people lookup')
    .replace(/^\s*[-*]?\s*Link:\s*https?:\/\/\S+\s*$/gim, '')
    .replace(/^\s*https?:\/\/\S+\s*$/gim, '')
    .replace(/\b(?:Link|URL)\s*:\s*https?:\/\/\S+/gi, '')
    .replace(/https?:\/\/\S+/gi, '')
    .replace(/^#{1,6}\s*/gm, '')
    .replace(/^\s*[-*•]\s*$/gm, '')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

const compactEvidenceSnippet = (value: string, maxLength = 220): string => {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  if (text.length <= maxLength) return text;

  const window = text.slice(0, maxLength);
  const sentenceEnd = Math.max(window.lastIndexOf('. '), window.lastIndexOf('? '), window.lastIndexOf('! '));
  if (sentenceEnd >= 80) return window.slice(0, sentenceEnd + 1).trim();

  const wordBoundary = window.replace(/\s+\S*$/, '').trim();
  return wordBoundary ? `${wordBoundary}.` : '';
};

const toStringArray = (value: string | string[] | undefined): string[] | undefined => {
  if (Array.isArray(value)) {
    const values = value.map((entry) => String(entry || '').trim()).filter(Boolean);
    return values.length > 0 ? values : undefined;
  }

  const text = String(value || '').trim();
  return text ? [text] : undefined;
};

const normalizeFileExtensions = (values: string[] | undefined): string[] | undefined => {
  const normalized = (values || [])
    .map((value) => String(value || '').trim().replace(/^\./, '').toUpperCase())
    .filter(Boolean);
  return normalized.length > 0 ? Array.from(new Set(normalized)) : undefined;
};

const normalizeFilters = (filters: SearchRequest['filters'] = {}): DocumentSearchFilters => ({
  status: 'Active',
  bu: toStringArray(filters.bu || filters.businessUnit),
  department: toStringArray(filters.department),
  diseaseArea: toStringArray(filters.diseaseArea),
  therapyArea: toStringArray(filters.therapyArea),
  client: toStringArray(filters.client),
  region: toStringArray(filters.region || filters.geography),
  documentType: toStringArray(filters.documentType),
  fileExtension: normalizeFileExtensions(toStringArray(filters.fileExtension || filters.fileType)),
  authors: toStringArray(filters.authors || filters.author)
});

const hasUserSelectedFilters = (filters: SearchRequest['filters'] = {}): boolean =>
  Object.entries(filters || {}).some(([key, value]) => {
    if (key === 'status') return false;
    return Array.isArray(value)
      ? value.some((item) => String(item || '').trim())
      : Boolean(String(value || '').trim());
  });

const SEARCH_FACET_FIELDS = [
  'status',
  'documentType',
  'bu',
  'department',
  'client',
  'region',
  'therapyArea',
  'diseaseArea',
  'fileExtension',
  'authors'
];

const NORMALIZED_SEARCH_FACET_FIELDS = [
  'status',
  'documentTypeFilterValues',
  'buFilterValues',
  'departmentFilterValues',
  'clientFilterValues',
  'regionFilterValues',
  'therapyAreaFilterValues',
  'diseaseAreaFilterValues',
  'fileExtension',
  'authors'
];

const SEARCH_FACET_ALIASES: Record<string, string[]> = {
  documentTypeFilterValues: ['documentType'],
  buFilterValues: ['bu', 'businessUnit'],
  departmentFilterValues: ['department'],
  departmentFacetValues: ['department'],
  clientFilterValues: ['client'],
  regionFilterValues: ['region', 'geography'],
  therapyAreaFilterValues: ['therapyArea'],
  diseaseAreaFilterValues: ['diseaseArea']
};

const activeFacetFields = (): string[] =>
  normalizedFacetsEnabled()
    ? NORMALIZED_SEARCH_FACET_FIELDS.map((field) =>
        field === 'departmentFilterValues' && canonicalDepartmentFacetsEnabled()
          ? 'departmentFacetValues'
          : field)
    : SEARCH_FACET_FIELDS;

const buildFacetExpressions = (env: SearchEnv): string[] => {
  const count = Math.min(Math.max(toNumber(env.AZURE_SEARCH_FACET_COUNT, 1000), 1), 1000);
  return activeFacetFields().map((field) => `${field},count:${count},sort:count`);
};

const getDocumentSortOrderBy = (sort: SearchRequest['sort']): string | undefined => {
  if (sort === 'newest') {
    return 'publishedDate desc, search.score() desc';
  }

  if (sort === 'oldest') {
    return 'publishedDate asc, search.score() desc';
  }

  return undefined;
};

const appendDocumentFilterClause = (baseFilter: string, clause: string): string =>
  baseFilter ? `${baseFilter} and ${clause}` : clause;

const comparableFacetPathPart = (value: string): string =>
  String(value || '')
    .normalize('NFKC')
    .replace(/＆/g, '&')
    .replace(/\s*&\s*/g, ' & ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[\s_-]+/g, ' ')
    .trim();

const splitFacetPath = (value: string): string[] =>
  String(value || '')
    .normalize('NFKC')
    .split(/\s*(?::|>|›|＞)\s*/g)
    .map((part) => part.trim().replace(/\s+/g, ' '))
    .filter(Boolean);

const collectBusinessUnitFacetKeys = (rawFacets: any): Set<string> => {
  const keys = new Set<string>();
  ['buFilterValues', 'bu', 'businessUnit'].forEach((field) => {
    const values = Array.isArray(rawFacets?.[field]) ? rawFacets[field] : [];
    values.forEach((item: any) => {
      const value = String(item?.value ?? '').trim();
      if (value) {
        keys.add(comparableFacetPathPart(value));
      }
    });
  });
  return keys;
};

const canonicalDepartmentFacetValue = (value: string, businessUnitFacetKeys: Set<string>): string => {
  const parts = splitFacetPath(value);
  if (parts.length <= 1) {
    return value;
  }

  const firstPartKey = comparableFacetPathPart(parts[0]);
  if (!businessUnitFacetKeys.has(firstPartKey)) {
    return parts.join(':');
  }

  return parts.slice(1).join(':');
};

const mapFacets = (rawFacets: any): Record<string, SearchFacetValue[]> => {
  const output: Record<string, SearchFacetValue[]> = {};
  if (!rawFacets || typeof rawFacets !== 'object') return output;

  const businessUnitFacetKeys = collectBusinessUnitFacetKeys(rawFacets);

  activeFacetFields().forEach((field) => {
    const values = Array.isArray(rawFacets[field]) ? rawFacets[field] : [];
    const counts = new Map<string, number>();

    values.forEach((item: any) => {
      const count = Number(item?.count || 0);
      if (count <= 0) return;

      const value = String(item?.value ?? '').trim();
      if (!value) return;

      const normalizedValue = value.toLowerCase().replace(/[\s_-]+/g, ' ') === 'not applicable' || value.toLowerCase() === 'n/a'
        ? 'Not Applicable'
        : value;
      const facetValue = field === 'departmentFacetValues' || field === 'departmentFilterValues' || field === 'department'
        ? canonicalDepartmentFacetValue(normalizedValue, businessUnitFacetKeys)
        : normalizedValue;
      const existingCount = counts.get(facetValue) || 0;
      counts.set(
        facetValue,
        field === 'departmentFacetValues' || field === 'departmentFilterValues' || field === 'department'
          ? Math.max(existingCount, count)
          : existingCount + count
      );
    });

    const mapped = Array.from(counts.entries())
      .map(([value, count]) => ({ value, count }))
      .sort((left, right) => right.count - left.count || left.value.localeCompare(right.value));

    const outputKeys = SEARCH_FACET_ALIASES[field] || [field];
    outputKeys.forEach((key) => {
      output[key] = mapped;
    });
  });

  return output;
};

const mapSearchHits = (items: any[] = []): SearchHit[] =>
  items.map((item) => ({
    id: item.id || '',
    documentId: item.documentId || item.id || '',
    listItemId: item.listItemId || '',
    driveItemId: item.driveItemId || '',
    fileUniqueId: item.fileUniqueId || item.uniqueId || '',
    uniqueId: item.uniqueId || item.fileUniqueId || '',
    fileRef: item.fileRef || item.serverRelativeUrl || '',
    serverRelativeUrl: item.serverRelativeUrl || item.fileRef || '',
    title: item.title || item.fileName || item.personName,
    fileName: item.fileName || '',
    fileExtension: item.fileExtension || '',
    webUrl: item.webUrl || item.listItemUrl || '',
    status: item.status || '',
    score: item['@search.score'],
    rerankerScore: item['@search.rerankerScore'],
    contentWindowHits: Number(item.contentWindowHits || item.__contentWindowHits || 0) || undefined,
    description: item.description || '',
    contentPreview: item.contentPreview || '',
    chunkText: item.chunkText || '',
    documentType: item.documentType || '',
    documentTypeFilterValues: Array.isArray(item.documentTypeFilterValues) ? item.documentTypeFilterValues : undefined,
    bu: item.bu || item.businessUnit || '',
    businessUnit: item.businessUnit || item.bu || '',
    buFilterValues: Array.isArray(item.buFilterValues) ? item.buFilterValues : undefined,
    department: item.department || '',
    departmentFilterValues: Array.isArray(item.departmentFilterValues) ? item.departmentFilterValues : undefined,
    departmentFacetValues: Array.isArray(item.departmentFacetValues) ? item.departmentFacetValues : undefined,
    diseaseArea: item.diseaseArea || '',
    diseaseAreaFilterValues: Array.isArray(item.diseaseAreaFilterValues) ? item.diseaseAreaFilterValues : undefined,
    therapyArea: item.therapyArea || '',
    therapyAreaFilterValues: Array.isArray(item.therapyAreaFilterValues) ? item.therapyAreaFilterValues : undefined,
    client: item.client || '',
    clientFilterValues: Array.isArray(item.clientFilterValues) ? item.clientFilterValues : undefined,
    region: item.region || item.geography || '',
    geography: item.geography || item.region || '',
    regionFilterValues: Array.isArray(item.regionFilterValues) ? item.regionFilterValues : undefined,
    authors: Array.isArray(item.authors) ? item.authors : toStringArray(item.authors) || [],
    createdDateTime: item.createdDateTime || item.created || '',
    created: item.created || item.createdDateTime || '',
    lastModifiedDateTime: item.lastModifiedDateTime || item.modified || '',
    modified: item.modified || item.lastModifiedDateTime || '',
    publishedDate: item.publishedDate || '',
    contentRefreshDate: item.contentRefreshDate || ''
  }));

const readODataCount = (result: Record<string, any>, fallback = 0): number => {
  if (Object.prototype.hasOwnProperty.call(result, '@odata.count')) {
    const count = Number(result['@odata.count']);
    return Number.isFinite(count) ? count : fallback;
  }

  return fallback;
};

const emitTelemetry = (context: SearchExecutionContext | undefined, event: Record<string, unknown>): void => {
  if (context?.log) {
    context.log(event);
    return;
  }

  console.log(JSON.stringify({
    level: 'info',
    ...event
  }));
};

const timedSearchIndex = async (
  env: SearchEnv,
  indexName: string,
  body: Record<string, unknown>,
  operation: string,
  context?: SearchExecutionContext
): Promise<any> => {
  const startedAt = Date.now();
  try {
    const result = await searchIndex(env, indexName, body);
    emitTelemetry(context, {
      level: 'info',
      event: 'azure-search-query',
      operation,
      indexName,
      durationMs: Date.now() - startedAt,
      resultCount: Array.isArray(result.value) ? result.value.length : undefined
    });
    return result;
  } catch (error: any) {
    emitTelemetry(context, {
      level: 'error',
      event: 'azure-search-query-failed',
      operation,
      indexName,
      durationMs: Date.now() - startedAt,
      error: error?.message || String(error)
    });
    throw error;
  }
};

const isUnknownSelectFieldError = (error: unknown): boolean => {
  const message = String((error as any)?.message || error || '').toLowerCase();
  return (
    message.includes('select') &&
    (
      message.includes('could not find a property named') ||
      message.includes('property') && message.includes('does not exist') ||
      message.includes('unknown field')
    )
  );
};

const timedDocumentSearchIndex = async (
  env: SearchEnv,
  body: Record<string, unknown>,
  operation: string,
  context?: SearchExecutionContext
): Promise<any> => {
  try {
    return await timedSearchIndex(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, body, operation, context);
  } catch (error) {
    if (body.select !== DOCUMENT_SELECT_FIELDS || !isUnknownSelectFieldError(error)) {
      throw error;
    }

    emitTelemetry(context, {
      level: 'warn',
      event: 'documents-select-legacy-fallback',
      operation,
      reason: 'index-schema-missing-new-fields'
    });

    return timedSearchIndex(
      env,
      env.AZURE_SEARCH_DOCUMENTS_INDEX,
      { ...body, select: LEGACY_DOCUMENT_SELECT_FIELDS },
      `${operation}-legacy-select`,
      context
    );
  }
};

const isSemanticSearchEnabled = (env: SearchEnv): boolean =>
  String(env.AZURE_SEARCH_SEMANTIC_ENABLED || 'true').toLowerCase() !== 'false';

const isSemanticBillingError = (error: unknown): boolean => {
  const message = String((error as any)?.message || error || '').toLowerCase();
  return (
    message.includes('semantic') &&
    (
      message.includes('402') ||
      message.includes('payment required') ||
      message.includes('usage exceeded') ||
      message.includes('enable semantic billing')
    )
  );
};

const escapeODataString = (value: string): string => value.replace(/'/g, "''");

const extractFileNameMentions = (query: string): string[] =>
  Array.from(new Set(
    (query.match(/[^\s"'<>]+?\.(?:pdf|docx?|pptx?|xlsx?|mp4|mp3|msg|png|jpe?g)/gi) || [])
      .map((value) => value.trim())
      .filter(Boolean)
  ));

const exactDocumentFilter = (query: string, filters?: SearchRequest['filters']): string | undefined => {
  if (isWildcardQuery(query)) return undefined;

  const clauses: string[] = [];
  identifierTokens(query).forEach((id) => {
    clauses.push(`listItemId eq '${escapeODataString(id)}'`);
  });

  const trimmedQuery = query.trim();
  if (trimmedQuery.length > 0 && trimmedQuery.length <= 240 && !/[?*:]/.test(trimmedQuery)) {
    clauses.push(`title eq '${escapeODataString(trimmedQuery)}'`);
    clauses.push(`fileName eq '${escapeODataString(trimmedQuery)}'`);
  }

  extractFileNameMentions(query).forEach((fileName) => {
    clauses.push(`fileName eq '${escapeODataString(fileName)}'`);
  });

  if (clauses.length === 0) return undefined;

  return `${buildDocumentFilter(normalizeFilters(filters))} and (${Array.from(new Set(clauses)).join(' or ')})`;
};

const mergePromotedDocumentHits = (
  query: string,
  primaryHits: SearchHit[],
  exactHits: SearchHit[],
  top: number,
  minScore: number
): SearchHit[] => {
  const promoted = exactHits
    .map((item) => ({ item, exactScore: exactDocumentMatchScore(query, item) }))
    .filter(({ item, exactScore }) => exactScore >= minScore || shouldPromoteExactDocument(query, item, minScore))
    .sort((left, right) => {
      if (right.exactScore !== left.exactScore) return right.exactScore - left.exactScore;
      return (right.item.score || 0) - (left.item.score || 0);
    })
    .map(({ item, exactScore }) => ({
      ...item,
      score: Math.max(item.score || 0, exactScore)
    }));

  const seen = new Set<string>();
  const output: SearchHit[] = [];

  [...promoted, ...primaryHits].forEach((item) => {
    const key = item.documentId || item.id || item.webUrl || item.fileName || item.title;
    if (!key || seen.has(key)) return;
    seen.add(key);
    output.push(item);
  });

  return output.slice(0, top);
};

const exactDocumentCandidates = async (
  env: SearchEnv,
  request: SearchRequest,
  top: number,
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  if (isWildcardQuery(query) || isCredentialDisclosureQuery(query)) return [];

  const select = withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS);
  const candidateTop = Math.max(top * 4, 20);
  const filter = buildDocumentFilter(normalizeFilters(request.filters));
  const exactFilter = exactDocumentFilter(query, request.filters);
  const personNameQuery = extractPersonNameQuery(query);
  const authorQuery = personNameQuery || compactMeaningfulQuery(query);
  const meaningfulQuery = compactMeaningfulQuery(query);
  const searches: Array<Promise<any>> = [
    timedDocumentSearchIndex(env, {
      search: query,
      queryType: isExplicitBooleanQuery(query) ? 'full' : 'simple',
      ...(isExplicitBooleanQuery(query) ? { searchMode: 'all' } : {}),
      filter,
      select,
      top: candidateTop
    }, 'documents-exact-candidates', context)
  ];

  if (meaningfulQuery && meaningfulQuery !== query.toLowerCase().trim()) {
    searches.push(timedDocumentSearchIndex(env, {
      search: meaningfulQuery,
      queryType: 'simple',
      filter,
      select,
      top: candidateTop
    }, 'documents-exact-meaningful', context));
  }

  if (personNameQuery && personNameQuery !== meaningfulQuery) {
    searches.push(timedDocumentSearchIndex(env, {
      search: personNameQuery,
      queryType: 'simple',
      filter,
      select,
      top: candidateTop
    }, 'documents-exact-person-name', context));
  }

  if (authorQuery) {
    searches.push(timedDocumentSearchIndex(env, {
      search: authorQuery,
      queryType: 'simple',
      searchMode: 'all',
      searchFields: 'authors',
      filter,
      select,
      top: candidateTop
    }, 'documents-exact-authors', context));

    searches.push(timedDocumentSearchIndex(env, {
      search: authorQuery,
      queryType: 'simple',
      searchMode: 'any',
      searchFields: 'authors',
      filter,
      select,
      top: candidateTop
    }, 'documents-exact-authors-any', context));
  }

  if (exactFilter) {
    searches.push(timedDocumentSearchIndex(env, {
      search: '*',
      filter: exactFilter,
      select,
      top: candidateTop
    }, 'documents-exact-filter', context));
  }

  identifierTokens(query).forEach((identifier) => {
    searches.push(timedDocumentSearchIndex(env, {
      search: identifier,
      queryType: 'simple',
      filter,
      select,
      top: candidateTop
    }, 'documents-exact-identifier', context));
  });

  extractFileNameMentions(query).forEach((fileName) => {
    searches.push(timedDocumentSearchIndex(env, {
      search: fileName,
      queryType: 'simple',
      filter,
      select,
      top: candidateTop
    }, 'documents-exact-filename', context));
  });

  const results = await Promise.all(searches);
  const byKey = new Map<string, SearchHit>();
  results.flatMap((result) => mapSearchHits(result.value as any[])).forEach((item) => {
    const key = item.id || item.webUrl || item.fileName || item.title;
    if (key && !byKey.has(key)) byKey.set(key, item);
  });

  return Array.from(byKey.values());
};

const escapeLuceneTerm = (value: string): string =>
  value.replace(/([+\-&|!(){}\[\]^"~*?:\\/])/g, '\\$1');

const fuzzyDocumentQuery = (query: string): string => {
  const tokens = meaningfulTokens(query)
    .filter((token) => token.length >= 4 && /^[a-z0-9]+$/i.test(token))
    .slice(0, 6);

  return tokens.map((token) => `${escapeLuceneTerm(token)}~1`).join(' ');
};

interface StructuredAuthorTopicCandidate {
  authorQuery: string;
  topicQuery: string;
}

const STRUCTURED_AUTHOR_TOPIC_IGNORED_TOKENS = new Set([
  'doc', 'docs', 'document', 'documents', 'file', 'files', 'asset', 'assets',
  'material', 'materials', 'related', 'about', 'from', 'under', 'for', 'by',
  'author', 'authors', 'created', 'uploaded'
]);

const structuredSearchTokens = (query: string): string[] =>
  meaningfulTokens(normalizeDomainSearchPhrase(query))
    .filter((token) => !STRUCTURED_AUTHOR_TOPIC_IGNORED_TOKENS.has(token));

const authorSpanCandidateQueries = (query: string): StructuredAuthorTopicCandidate[] => {
  if (isWildcardQuery(query) || isExplicitBooleanQuery(query) || likelyPastedSearchText(query)) {
    return [];
  }

  const tokens = structuredSearchTokens(query);
  if (tokens.length < 3 || tokens.length > 8) {
    return [];
  }

  const candidates: StructuredAuthorTopicCandidate[] = [];
  const addCandidate = (startIndex: number, length: number): void => {
    if (startIndex < 0 || startIndex + length > tokens.length || length < 2) return;
    const authorTokens = tokens.slice(startIndex, startIndex + length);
    const topicTokens = [
      ...tokens.slice(0, startIndex),
      ...tokens.slice(startIndex + length)
    ];
    if (topicTokens.length === 0) return;

    candidates.push({
      authorQuery: authorTokens.join(' '),
      topicQuery: topicTokens.join(' ')
    });
  };

  const explicitPersonQuery = extractPersonNameQuery(query);
  if (explicitPersonQuery) {
    const authorTokens = meaningfulTokens(explicitPersonQuery);
    const authorTokenSet = new Set(authorTokens);
    const topicTokens = tokens.filter((token) => !authorTokenSet.has(token));
    if (authorTokens.length >= 2 && topicTokens.length > 0) {
      candidates.push({
        authorQuery: authorTokens.join(' '),
        topicQuery: topicTokens.join(' ')
      });
    }
  }

  addCandidate(tokens.length - 2, 2);
  addCandidate(tokens.length - 3, 3);
  addCandidate(0, 2);

  if (tokens.length <= 5) {
    for (let index = 1; index < tokens.length - 2; index += 1) {
      addCandidate(index, 2);
    }
  }

  const seen = new Set<string>();
  return candidates.filter((candidate) => {
    const key = `${candidate.authorQuery}|${candidate.topicQuery}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

const authorMatchScore = (authorQuery: string, item: SearchHit): number => {
  const authorsText = Array.isArray(item.authors) ? item.authors.join(' ') : String(item.authors || '');
  if (!authorsText.trim()) return 0;

  const exact = authorLexicalOverlapCount(authorQuery, authorsText);
  const fuzzy = fuzzyLexicalOverlapCount(authorQuery, authorsText);
  const required = Math.min(2, Math.max(1, exact.tokenTotal));
  return Math.max(exact.overlap, fuzzy) >= required
    ? Math.max(exact.overlap, fuzzy) * 100
    : 0;
};

const topicMatchScore = (topicQuery: string, item: SearchHit): number => {
  const evidenceText = documentEvidenceText(item);
  const exact = lexicalOverlapCount(topicQuery, evidenceText);
  const fuzzy = fuzzyLexicalOverlapCount(topicQuery, evidenceText);
  const tokenTotal = meaningfulTokens(topicQuery).length;
  const required = Math.min(2, Math.max(1, Math.ceil(tokenTotal * 0.5)));
  const overlap = Math.max(exact, fuzzy);

  return overlap >= required
    ? (overlap * 120) + exactDocumentMatchScore(topicQuery, item) + (item.score || 0)
    : 0;
};

const uniqueDocumentHits = (items: SearchHit[]): SearchHit[] => {
  const seen = new Set<string>();
  const output: SearchHit[] = [];

  items.forEach((item) => {
    const key = item.documentId || item.id || item.webUrl || item.fileName || item.title;
    if (!key || seen.has(key)) return;
    seen.add(key);
    output.push(item);
  });

  return output;
};

const TITLE_FILENAME_GENERIC_PROMOTION_TOKENS = new Set([
  'asset', 'assets', 'material', 'materials', 'document', 'documents', 'doc', 'docs',
  'file', 'files', 'deck', 'decks', 'presentation', 'presentations',
  'case', 'study', 'studies', 'proposal', 'proposals', 'rfp', 'rfi',
  'capability', 'capabilities', 'sop', 'summary', 'overview',
  'knowledge', 'sharing', 'session', 'recording', 'recordings', 'available'
]);

const compactSpecificTitleFileNameQuery = (query: string): string =>
  meaningfulTokens(query)
    .filter((token) => !TITLE_FILENAME_GENERIC_PROMOTION_TOKENS.has(token))
    .join(' ');

const hasSpecificTitleFileNameOverlap = (query: string, item: SearchHit): boolean => {
  const evidence = ` ${comparableText([item.title, item.fileName].filter(Boolean).join(' '))} `;
  const queryTokens = Array.from(new Set(meaningfulTokens(query).map((token) => comparableText(token)).filter(Boolean)));
  const overlaps = queryTokens
    .filter((token) => evidence.includes(` ${token} `));
  const specificTokens = queryTokens.filter((token) => !TITLE_FILENAME_GENERIC_PROMOTION_TOKENS.has(token));
  const specificOverlap = overlaps.filter((token) => !TITLE_FILENAME_GENERIC_PROMOTION_TOKENS.has(token)).length;

  if (specificTokens.length > 0) {
    const requiredSpecificOverlap = specificTokens.length === 1 ? 1 : Math.min(2, specificTokens.length);
    return specificOverlap >= requiredSpecificOverlap;
  }

  return overlaps.length >= 3;
};

const titleFileNameDocumentCandidates = async (
  env: SearchEnv,
  request: SearchRequest,
  top: number,
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  if (isWildcardQuery(query) || isExplicitBooleanQuery(query) || isCredentialDisclosureQuery(query)) {
    return [];
  }

  // Chunk rows can lag document rows for normalized taxonomy fields. Pull the
  // content candidate window with only the base active-document filter, then
  // apply selected filters after hydrating from the document index. This keeps
  // facet counts and click-through totals based on the same hydrated records.
  const filter = buildDocumentFilter(normalizeFilters());
  const candidateTop = Math.min(Math.max(top * 8, 60), 120);
  const variants = [
    ...Array.from(new Set([
      query,
      normalizeDomainSearchPhrase(query),
      compactMeaningfulQuery(query)
    ].map((value) => value.trim()).filter(Boolean)))
      .map((searchText) => ({ searchText, searchMode: 'any' as const })),
    ...Array.from(new Set([
      compactSpecificTitleFileNameQuery(query)
    ].map((value) => value.trim()).filter((value) => meaningfulTokens(value).length >= 2)))
      .map((searchText) => ({ searchText, searchMode: 'all' as const }))
  ].filter((entry, index, all) =>
    all.findIndex((candidate) => candidate.searchText === entry.searchText && candidate.searchMode === entry.searchMode) === index
  );

  const results = await Promise.all(variants.map(({ searchText, searchMode }, index) =>
    timedDocumentSearchIndex(env, {
      search: searchText,
      queryType: 'simple',
      searchMode,
      searchFields: 'title,fileName',
      filter,
      select: withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS),
      top: candidateTop
    }, index === 0 ? 'documents-title-filename-candidates' : 'documents-title-filename-candidates-alt', context)
  ));

  return uniqueDocumentHits(results.flatMap((result) => mapSearchHits(result.value as any[])))
    .filter((item) => hasSpecificTitleFileNameOverlap(query, item));
};

const mergeDocumentMetadata = (metadata: SearchHit, source: SearchHit): SearchHit => ({
  ...metadata,
  chunkText: source.chunkText || metadata.chunkText,
  score: Math.max(metadata.score || 0, source.rerankerScore || source.score || 0),
  rerankerScore: source.rerankerScore ?? metadata.rerankerScore,
  contentWindowHits: source.contentWindowHits ?? metadata.contentWindowHits
});

const hydrateContentHitsFromDocumentIndex = async (
  env: SearchEnv,
  hits: SearchHit[],
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  const clauses: string[] = uniqueDocumentHits(hits)
    .flatMap((item) => [
      item.documentId ? `id eq '${escapeODataString(item.documentId)}'` : '',
      item.listItemId ? `listItemId eq '${escapeODataString(item.listItemId)}'` : ''
    ])
    .filter((clause): clause is string => Boolean(clause));

  if (clauses.length === 0) return hits;

  const hydratedByKey = new Map<string, SearchHit>();
  const batchSize = 40;

  for (let index = 0; index < clauses.length; index += batchSize) {
    const batchClauses = clauses.slice(index, index + batchSize);
    const result = await timedDocumentSearchIndex(env, {
      search: '*',
      filter: `${buildDocumentFilter(normalizeFilters())} and (${Array.from(new Set(batchClauses)).join(' or ')})`,
      select: withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS),
      top: batchClauses.length
    }, 'documents-hydrate-content-hits', context);

    mapSearchHits(result.value as any[]).forEach((item) => {
      [item.id, item.documentId, item.listItemId].filter(Boolean).forEach((key) => {
        hydratedByKey.set(String(key), item);
      });
    });
  }

  return hits.map((item) => {
    const metadata =
      (item.documentId && hydratedByKey.get(item.documentId)) ||
      (item.id && hydratedByKey.get(item.id)) ||
      (item.listItemId && hydratedByKey.get(item.listItemId));

    return metadata ? mergeDocumentMetadata(metadata, item) : item;
  });
};

const contentDocumentCandidates = async (
  env: SearchEnv,
  request: SearchRequest,
  top: number,
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  if (isWildcardQuery(query)) return [];

  const contentTop = Math.min(top, toNumber(env.AZURE_SEARCH_DOCUMENT_CONTENT_TOP, top));
  const hits = await semanticSearch(env, { ...request, query, top: contentTop }, context);
  const hydratedHits = await hydrateContentHitsFromDocumentIndex(env, hits, context);
  const minOverlap = Math.min(2, Math.max(1, meaningfulTokens(query).length));

  return hydratedHits.filter((item) => {
    const evidenceText = documentEvidenceText(item);
    const overlap = Math.max(lexicalOverlapCount(query, evidenceText), fuzzyLexicalOverlapCount(query, evidenceText));
    return overlap >= minOverlap || shouldPromoteExactDocument(query, item, toNumber(env.AZURE_SEARCH_EXACT_DOCUMENT_MIN_SCORE, 80));
  });
};

const getFacetValuesFromHit = (item: SearchHit, field: string): string[] => {
  const value = (item as any)[field];
  if (Array.isArray(value)) {
    return value.map((entry) => String(entry || '').trim()).filter(Boolean);
  }

  const text = String(value || '').trim();
  return text ? [text] : [];
};

const normalizeFacetMatchValue = (value: unknown): string =>
  String(value || '')
    .normalize('NFKC')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/&/g, 'and')
    .replace(/＆/g, 'and')
    .replace(/\s*>\s*/g, ':')
    .replace(/\s*›\s*/g, ':')
    .replace(/\s*＞\s*/g, ':')
    .replace(/[\s_-]+/g, ' ')
    .trim()
    .toLowerCase();

const expandFacetPathMatchValues = (value: string): string[] => {
  const parts = value.split(':').map((part) => part.trim()).filter(Boolean);
  if (parts.length <= 1) {
    return [value].filter(Boolean);
  }

  return [
    value,
    parts.slice(1).join(':')
  ].filter(Boolean);
};

const splitFacetMatchValues = (value: unknown): string[] => {
  if (Array.isArray(value)) {
    return value.flatMap((entry) => splitFacetMatchValues(entry));
  }

  const rawValue = String(value || '');
  const values = [
    ...expandFacetPathMatchValues(normalizeFacetMatchValue(rawValue)),
    ...rawValue
      .split(/[;,|]/g)
      .flatMap((entry) => expandFacetPathMatchValues(normalizeFacetMatchValue(entry)))
  ];

  return values
    .filter((entry, index, allValues) => entry && allValues.indexOf(entry) === index);
};

const facetValueMatchesFilter = (sourceValue: unknown, expectedValues?: string | string[]): boolean => {
  const expected = (Array.isArray(expectedValues) ? expectedValues : expectedValues ? [expectedValues] : [])
    .flatMap((value) => splitFacetMatchValues(value))
    .filter(Boolean);
  if (expected.length === 0) return true;

  const source = splitFacetMatchValues(sourceValue);
  if (source.length === 0) return false;

  return expected.some((expectedValue) => source.some((sourceValue) => sourceValue === expectedValue));
};

const hitMatchesRequestedFilters = (item: SearchHit, rawFilters?: SearchRequest['filters']): boolean => {
  const filters = normalizeFilters(rawFilters);
  const normalized = normalizedFacetsEnabled();

  return (
    facetValueMatchesFilter(item.status, filters.status) &&
    facetValueMatchesFilter(normalized ? item.documentTypeFilterValues || item.documentType : item.documentType, filters.documentType) &&
    facetValueMatchesFilter(normalized ? item.buFilterValues || item.bu || (item as any).businessUnit : item.bu || (item as any).businessUnit, filters.bu) &&
    facetValueMatchesFilter(normalized ? item.departmentFilterValues || item.department : item.department, filters.department) &&
    facetValueMatchesFilter(normalized ? item.clientFilterValues || item.client : item.client, filters.client) &&
    facetValueMatchesFilter(normalized ? item.regionFilterValues || item.region || (item as any).geography : item.region || (item as any).geography, filters.region) &&
    facetValueMatchesFilter(normalized ? item.therapyAreaFilterValues || item.therapyArea : item.therapyArea, filters.therapyArea) &&
    facetValueMatchesFilter(normalized ? item.diseaseAreaFilterValues || item.diseaseArea : item.diseaseArea, filters.diseaseArea) &&
    facetValueMatchesFilter(item.fileExtension || (item as any).fileType, filters.fileExtension) &&
    facetValueMatchesFilter(item.authors || (item as any).author, filters.authors)
  );
};

const buildFacetsFromHits = (items: SearchHit[]): Record<string, SearchFacetValue[]> => {
  const output: Record<string, SearchFacetValue[]> = {};

  activeFacetFields().forEach((field) => {
    const counts = new Map<string, number>();

    items.forEach((item) => {
      getFacetValuesFromHit(item, field).forEach((rawValue) => {
        const value = rawValue.toLowerCase().replace(/[\s_-]+/g, ' ') === 'not applicable' || rawValue.toLowerCase() === 'n/a'
          ? 'Not Applicable'
          : rawValue;
        counts.set(value, (counts.get(value) || 0) + 1);
      });
    });

    const mapped = Array.from(counts.entries())
      .map(([value, count]) => ({ value, count }))
      .filter((facet) => facet.count > 0)
      .sort((left, right) => right.count - left.count || left.value.localeCompare(right.value));

    const outputKeys = SEARCH_FACET_ALIASES[field] || [field];
    outputKeys.forEach((key) => {
      output[key] = mapped;
    });
  });

  return output;
};

const rankContentDocumentHit = (query: string, item: SearchHit): number => {
  const chunkText = item.chunkText || '';
  const evidenceText = documentEvidenceText(item);
  const queryTokens = meaningfulTokens(query);
  const titleTokens = meaningfulTokens(item.title || item.fileName || '');
  const tokenTotal = Math.max(queryTokens.length, 1);
  const chunkOverlap = lexicalOverlapCount(query, chunkText);
  const evidenceOverlap = lexicalOverlapCount(query, evidenceText);
  const fuzzyOverlap = fuzzyLexicalOverlapCount(query, evidenceText);
  const overlapRatio = Math.max(chunkOverlap, evidenceOverlap, fuzzyOverlap) / tokenTotal;
  const rerankerScore = item.rerankerScore || 0;
  const searchScore = item.score || 0;
  const exactScore = exactDocumentMatchScore(query, item);
  const leadingTitleIntentBonus = queryTokens.length >= 3
    && titleTokens.length >= 2
    && queryTokens[0] === titleTokens[0]
    && queryTokens[1] === titleTokens[1]
    ? 700
    : 0;
  const normalizedEvidenceText = normalizedIdentifierText(evidenceText);
  const identifierBonus = rawContentIdentifierTokens(query)
    .filter((identifier) => normalizedEvidenceText.includes(normalizedIdentifierText(identifier)))
    .length * 200;

  return exactScore + leadingTitleIntentBonus + identifierBonus + ((item.contentWindowHits || 0) * 180) + (overlapRatio * 1000) + (chunkOverlap * 20) + (rerankerScore * 50) + searchScore;
};

const publishedDateTimestamp = (item: SearchHit): number | null => {
  const dateValue = item.publishedDate || '';
  const timestamp = Date.parse(dateValue);
  return Number.isFinite(timestamp) ? timestamp : null;
};

const comparePublishedDate = (left: SearchHit, right: SearchHit, sort: SearchRequest['sort']): number => {
  const leftTimestamp = publishedDateTimestamp(left);
  const rightTimestamp = publishedDateTimestamp(right);
  const leftMissing = leftTimestamp === null;
  const rightMissing = rightTimestamp === null;

  if (leftMissing && rightMissing) return 0;
  if (leftMissing) return 1;
  if (rightMissing) return -1;

  if (sort === 'newest') {
    return rightTimestamp - leftTimestamp;
  }

  return leftTimestamp - rightTimestamp;
};

const specificQueryEvidenceScore = (query: string, item: SearchHit): number => {
  if (isWildcardQuery(query)) return 0;

  const tokens = Array.from(new Set(
    meaningfulTokens(query)
      .filter((token) => !TITLE_FILENAME_GENERIC_PROMOTION_TOKENS.has(token))
      .map((token) => comparableText(token))
      .filter(Boolean)
  ));
  if (tokens.length === 0) return 0;

  const titleFileText = ` ${comparableText([item.title, item.fileName].filter(Boolean).join(' '))} `;
  const evidenceText = ` ${comparableText(documentEvidenceText(item))} `;
  let score = 0;
  let matched = 0;

  tokens.forEach((token, index) => {
    const needle = ` ${token} `;
    const inTitleFile = titleFileText.includes(needle);
    const inEvidence = evidenceText.includes(needle);
    if (!inEvidence) return;

    matched += 1;
    score += inTitleFile ? 4 : 2;
    if (index === 0) {
      score += inTitleFile ? 6 : 3;
    }
  });

  if (matched >= Math.min(2, tokens.length)) score += 4;
  if (matched === tokens.length && tokens.length >= 2) score += 6;

  return score;
};

const contentSourceConfidenceTier = (query: string, item: SearchHit): number => {
  const chunkText = item.chunkText || '';
  const evidenceText = documentEvidenceText(item);
  const queryTokenCount = Math.max(meaningfulTokens(query).length, 1);
  const chunkOverlap = lexicalOverlapCount(query, chunkText);
  const evidenceOverlap = lexicalOverlapCount(query, evidenceText);
  const hasWindowEvidence = hasContentWindowEvidence(query, chunkText);
  const hasIdentifier = hasIdentifierEvidence(query, chunkText) || hasIdentifierEvidence(query, evidenceText);
  const exactScore = exactDocumentMatchScore(query, item);
  const highOverlap = Math.max(6, Math.ceil(queryTokenCount * 0.08));
  const mediumOverlap = Math.max(4, Math.ceil(queryTokenCount * 0.05));
  const windowHits = item.contentWindowHits || 0;

  if (exactScore >= 120) return 0;
  if (exactScore >= 80 && hasSpecificTitleFileNameOverlap(query, item)) return 1;
  if (hasIdentifier || (hasWindowEvidence && (chunkOverlap >= highOverlap || windowHits >= 3))) return 0;
  if (hasWindowEvidence || chunkOverlap >= highOverlap || windowHits >= 2) return 1;
  if (chunkOverlap >= mediumOverlap || evidenceOverlap >= mediumOverlap) return 2;
  return 3;
};

const sortDocumentHits = (query: string, items: SearchHit[], sort: SearchRequest['sort']): SearchHit[] => {
  if (sort === 'newest' || sort === 'oldest') {
    const sourceConfidenceFirst = likelyPastedDocumentText(query);
    return items.slice().sort((left, right) => {
      if (sourceConfidenceFirst) {
        const tierDifference = contentSourceConfidenceTier(query, left) - contentSourceConfidenceTier(query, right);
        if (tierDifference !== 0) return tierDifference;
      }

      const dateDifference = comparePublishedDate(left, right, sort);
      if (dateDifference !== 0) return dateDifference;

      if (!sourceConfidenceFirst) {
        const tierDifference = contentSourceConfidenceTier(query, left) - contentSourceConfidenceTier(query, right);
        if (tierDifference !== 0) return tierDifference;
      }

      return rankContentDocumentHit(query, right) - rankContentDocumentHit(query, left);
    });
  }

  return items.slice().sort((left, right) => {
    const specificDifference = specificQueryEvidenceScore(query, right) - specificQueryEvidenceScore(query, left);
    if (specificDifference !== 0) return specificDifference;

    const tierDifference = contentSourceConfidenceTier(query, left) - contentSourceConfidenceTier(query, right);
    if (tierDifference !== 0) return tierDifference;

    return rankContentDocumentHit(query, right) - rankContentDocumentHit(query, left);
  });
};

const hasSequentialTokenEvidence = (query: string, text: string, sequenceLength = 4): boolean => {
  const queryTokens = meaningfulTokens(query);
  const textTokens = meaningfulTokens(text);
  if (queryTokens.length < sequenceLength || textTokens.length < sequenceLength) return false;

  const textWindow = ` ${textTokens.join(' ')} `;
  for (let index = 0; index <= queryTokens.length - sequenceLength; index += 1) {
    const phrase = queryTokens.slice(index, index + sequenceLength).join(' ');
    if (textWindow.includes(` ${phrase} `)) return true;
  }

  return false;
};

const normalizedIdentifierText = (value: string): string =>
  String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, '');

const rawContentIdentifierTokens = (query: string): string[] => {
  const value = String(query || '');
  const matches = [
    ...(value.match(/\b\d{2,7}-\d{2}-\d\b/g) || []),
    ...(value.match(/\b(?:H\d{3}|GHS\d{2})\b/gi) || [])
  ];

  return Array.from(new Set(matches.map((token) => token.trim()).filter((token) => token.length >= 4)));
};

const hasIdentifierEvidence = (query: string, text: string): boolean => {
  const identifiers = rawContentIdentifierTokens(query);
  if (identifiers.length === 0) return false;

  const haystack = normalizedIdentifierText(text);
  return identifiers.some((identifier) => haystack.includes(normalizedIdentifierText(identifier)));
};

const buildContentPhraseWindows = (query: string): string[] => {
  const tokens = meaningfulTokens(query);
  if (tokens.length === 0) return [];
  if (tokens.length <= 6) return [tokens.join(' ')];

  const windowLength = tokens.length >= 24 ? 6 : 5;
  const starts = new Set<number>([
    0,
    Math.floor(tokens.length * 0.2),
    Math.floor(tokens.length * 0.45),
    Math.floor(tokens.length * 0.7),
    Math.max(0, tokens.length - windowLength)
  ]);

  tokens.forEach((token, index) => {
    if (token.length >= 10 || /\d/.test(token)) {
      starts.add(Math.max(0, Math.min(index - 2, tokens.length - windowLength)));
    }
  });

  return Array.from(starts)
    .sort((left, right) => left - right)
    .map((start) => tokens.slice(start, start + windowLength).join(' '))
    .filter((window) => meaningfulTokens(window).length >= 3)
    .filter((window, index, all) => all.indexOf(window) === index)
    .slice(0, 8);
};

const buildContentCandidateQueries = (query: string): string[] => {
  if (!likelyPastedDocumentText(query)) {
    const identifiers = rawContentIdentifierTokens(query);
    return identifiers.length > 0 ? Array.from(new Set([...identifiers, query])) : [query];
  }

  const phraseWindows = buildContentPhraseWindows(query);
  const identifiers = rawContentIdentifierTokens(query);
  const compactQuery = compactPrecisionQuery(query);

  return Array.from(new Set([...phraseWindows, ...identifiers, compactQuery]))
    .map((value) => value.trim())
    .filter(Boolean)
    .slice(0, 10);
};

const hasContentWindowEvidence = (query: string, text: string): boolean => {
  if (hasIdentifierEvidence(query, text)) return true;

  return buildContentPhraseWindows(query).some((window) => {
    const windowTokens = meaningfulTokens(window);
    const requiredLength = Math.min(4, windowTokens.length);
    return requiredLength >= 3 && hasSequentialTokenEvidence(window, text, requiredLength);
  });
};

const chunkContentDocumentCandidates = async (
  env: SearchEnv,
  request: SearchRequest,
  candidateTop: number,
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  if (isWildcardQuery(query) || isExplicitBooleanQuery(query) || isCredentialDisclosureQuery(query)) return [];

  // Chunk rows can lag document rows for normalized taxonomy fields. Pull the
  // content candidate window with only the base active-document filter, then
  // apply selected filters after hydrating from the document index. This keeps
  // facet counts and click-through totals based on the same hydrated records.
  const filter = buildDocumentFilter(normalizeFilters());
  const candidateQueries = buildContentCandidateQueries(query);
  const topPerQuery = Math.min(
    candidateTop,
    Math.max(40, Math.ceil(candidateTop / Math.max(1, Math.min(candidateQueries.length, 4))))
  );
  const rawHits: any[] = [];

  for (const candidateQuery of candidateQueries) {
    const lexicalResult = await timedSearchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
      search: candidateQuery,
      queryType: 'simple',
      searchMode: 'all',
      searchFields: 'chunkText',
      filter,
      select: withNormalizedFilterSelectFields('id,documentId,listItemId,title,fileName,webUrl,status,description,chunkText,documentType,bu,department,diseaseArea,therapyArea,client,region,authors'),
      top: topPerQuery
    }, 'documents-paged-content-lexical', context);
    rawHits.push(...(lexicalResult.value as any[]).map((item) => ({
      ...item,
      __contentWindowKey: candidateQuery
    })));
  }

  const lexicalHits = mapSearchHits(dedupeContentHitsByDocument(rawHits));
  const hydratedHits = await hydrateContentHitsFromDocumentIndex(
    env,
    uniqueDocumentHits(lexicalHits),
    context
  );
  const tokenCount = meaningfulTokens(query).length;
  const isLongContentQuery = tokenCount >= 8 || likelyPastedDocumentText(query);
  const hasRawIdentifier = rawContentIdentifierTokens(query).length > 0;
  const requiresChunkEvidence = isLongContentQuery || hasRawIdentifier;
  const minimumOverlap = isLongContentQuery
    ? Math.max(4, Math.ceil(tokenCount * 0.18))
    : Math.min(3, Math.max(1, Math.ceil(tokenCount * 0.2)));
  const minimumChunkOverlap = isLongContentQuery
    ? Math.max(3, Math.ceil(tokenCount * 0.12))
    : 1;
  const highConfidenceChunkOverlap = isLongContentQuery
    ? Math.max(minimumChunkOverlap, Math.ceil(tokenCount * 0.25))
    : minimumChunkOverlap;

  return uniqueDocumentHits(hydratedHits)
    .filter((item) => hitMatchesRequestedFilters(item, request.filters))
    .filter((item) => {
      if (shouldPromoteExactDocument(query, item, toNumber(env.AZURE_SEARCH_EXACT_DOCUMENT_MIN_SCORE, 80))) {
        return true;
      }

      const chunkOverlap = lexicalOverlapCount(query, item.chunkText || '');
      const evidenceText = documentEvidenceText(item);
      const overlap = Math.max(
        chunkOverlap,
        lexicalOverlapCount(query, evidenceText),
        fuzzyLexicalOverlapCount(query, evidenceText)
      );
      if (requiresChunkEvidence && !hasContentWindowEvidence(query, item.chunkText || '')) return false;
      if (isLongContentQuery && chunkOverlap < minimumChunkOverlap) return false;
      if (
        isLongContentQuery &&
        chunkOverlap < highConfidenceChunkOverlap &&
        !hasContentWindowEvidence(query, item.chunkText || '')
      ) {
        return false;
      }
      return overlap >= minimumOverlap;
    })
    .sort((left, right) => rankContentDocumentHit(query, right) - rankContentDocumentHit(query, left));
};

const searchContentDocumentsPaged = async (
  env: SearchEnv,
  request: SearchRequest,
  requestedTop: number,
  skip: number,
  context?: SearchExecutionContext
): Promise<SearchPageResponse | undefined> => {
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  const compactBodyPhrase = likelyCompactBodyPhraseQuery(query);
  if (!compactBodyPhrase && !likelyPastedSearchText(query) && !likelyPastedDocumentText(query) && rawContentIdentifierTokens(query).length === 0) {
    return undefined;
  }

  const maxCandidates = Math.min(Math.max(toNumber(env.AZURE_SEARCH_CONTENT_PAGE_CANDIDATE_TOP, 500), requestedTop), 1000);
  const candidateTop = Math.min(Math.max(skip + requestedTop, requestedTop * 4, 80), maxCandidates);
  let candidates = await chunkContentDocumentCandidates(env, request, candidateTop, context);
  if (likelyPastedSearchText(query) || likelyPastedDocumentText(query)) {
    const exactCandidates = (await exactDocumentCandidates(env, request, candidateTop, context))
      .filter((item) => hitMatchesRequestedFilters(item, request.filters))
      .filter((item) =>
        shouldPromoteExactDocument(query, item, toNumber(env.AZURE_SEARCH_EXACT_DOCUMENT_MIN_SCORE, 80)) ||
        passesDocumentRelevanceGuard(env, query, item)
      );

    candidates = uniqueDocumentHits([...exactCandidates, ...candidates]);
  }

  if (candidates.length === 0) {
    emitTelemetry(context, {
      level: 'warn',
      event: 'documents-paged-content-empty',
      reason: 'no-content-candidates'
    });
    if (compactBodyPhrase) {
      return undefined;
    }

    return {
      results: [],
      facets: {},
      totalCount: 0,
      page: {
        top: requestedTop,
        skip,
        returned: 0,
        totalAvailable: 0,
        resultWindow: 0,
        capped: false,
        hasMore: false
      }
    };
  }

  const sortedCandidates = sortDocumentHits(query, candidates, request.sort);
  const results = sortedCandidates.slice(skip, skip + requestedTop);
  const totalAvailable = sortedCandidates.length;
  const hasMore = skip + results.length < totalAvailable;

  emitTelemetry(context, {
    level: 'info',
    event: 'documents-paged-content-results',
    candidateCount: sortedCandidates.length,
    returned: results.length,
    skip,
    top: requestedTop
  });

  return {
    results,
    totalCount: totalAvailable,
    facets: request.includeFacets ? buildFacetsFromHits(candidates) : {},
    page: {
      top: requestedTop,
      skip,
      returned: results.length,
      totalAvailable,
      resultWindow: totalAvailable,
      capped: sortedCandidates.length >= maxCandidates,
      hasMore,
      nextSkip: hasMore ? skip + results.length : undefined
    }
  };
};

const fuzzyDocumentCandidates = async (
  env: SearchEnv,
  request: SearchRequest,
  top: number,
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  if (isWildcardQuery(query) || isCredentialDisclosureQuery(query) || isExplicitBooleanQuery(query)) return [];

  const fuzzyQuery = fuzzyDocumentQuery(query);
  if (!fuzzyQuery) return [];

  const result = await timedDocumentSearchIndex(env, {
    search: fuzzyQuery,
    queryType: 'full',
    searchMode: 'any',
    filter: buildDocumentFilter(normalizeFilters(request.filters)),
    select: withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS),
    top: Math.max(top * 3, 20)
  }, 'documents-fuzzy-candidates', context);

  return mapSearchHits(result.value as any[])
    .filter((item) => passesDocumentRelevanceGuard(env, query, item));
};

const sortStructuredAuthorTopicHits = (items: SearchHit[], sort: SearchRequest['sort']): SearchHit[] =>
  items.slice().sort((left, right) => {
    if (sort === 'newest' || sort === 'oldest') {
      const dateComparison = comparePublishedDate(left, right, sort);
      if (dateComparison !== 0) return dateComparison;
    }

    return (right.score || 0) - (left.score || 0);
  });

const structuredAuthorTopicPagedSearch = async (
  env: SearchEnv,
  request: SearchRequest,
  requestedTop: number,
  skip: number,
  query: string,
  baseFilter: string,
  context?: SearchExecutionContext
): Promise<SearchPageResponse | undefined> => {
  const structuredCandidates = authorSpanCandidateQueries(query);
  if (structuredCandidates.length === 0) {
    return undefined;
  }

  const candidateTop = Math.min(Math.max(toNumber(env.AZURE_SEARCH_STRUCTURED_AUTHOR_TOPIC_TOP, 1000), requestedTop), 1000);
  const candidateResults = await Promise.all(structuredCandidates.map((candidate, index) =>
    timedDocumentSearchIndex(env, {
      search: candidate.authorQuery,
      queryType: 'simple',
      searchMode: 'any',
      searchFields: 'authors',
      filter: baseFilter,
      select: withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS),
      top: candidateTop,
      skip: 0,
      count: false
    }, index === 0 ? 'documents-structured-author-topic' : 'documents-structured-author-topic-alt', context)
      .then((result) => ({ candidate, result }))
  ));

  let reachedCandidateLimit = false;
  const byKey = new Map<string, SearchHit>();

  candidateResults.forEach(({ candidate, result }) => {
    const rawHits = mapSearchHits(result.value as any[]);
    if (rawHits.length >= candidateTop) {
      reachedCandidateLimit = true;
    }

    rawHits.forEach((item) => {
      const authorScore = authorMatchScore(candidate.authorQuery, item);
      if (authorScore <= 0) return;

      const topicScore = topicMatchScore(candidate.topicQuery, item);
      if (topicScore <= 0) return;

      const score = authorScore + topicScore;
      const key = item.documentId || item.id || item.webUrl || item.fileName || item.title;
      const existing = key ? byKey.get(key) : undefined;
      if (!key || !existing || score > (existing.score || 0)) {
        byKey.set(key || `${item.title}-${byKey.size}`, {
          ...item,
          score
        });
      }
    });
  });

  const ranked = sortStructuredAuthorTopicHits(Array.from(byKey.values()), request.sort);
  if (ranked.length === 0) {
    return undefined;
  }

  const pageResults = ranked.slice(skip, skip + requestedTop);

  const hasMore = skip + pageResults.length < ranked.length;
  emitTelemetry(context, {
    event: 'documents-structured-author-topic-applied',
    queryTokenCount: meaningfulTokens(query).length,
    candidatePatternCount: structuredCandidates.length,
    resultWindow: ranked.length,
    returned: pageResults.length
  });

  return {
    results: pageResults,
    totalCount: ranked.length,
    facets: request.includeFacets ? buildFacetsFromHits(ranked) : {},
    page: {
      top: requestedTop,
      skip,
      returned: pageResults.length,
      totalAvailable: ranked.length,
      resultWindow: ranked.length,
      capped: reachedCandidateLimit,
      hasMore,
      nextSkip: hasMore ? skip + pageResults.length : undefined
    }
  };
};

const relaxedPagedFallback = async (
  env: SearchEnv,
  request: SearchRequest,
  requestedTop: number,
  skip: number,
  query: string,
  baseFilter: string,
  context?: SearchExecutionContext
): Promise<SearchPageResponse | undefined> => {
  if (isWildcardQuery(query) || isExplicitBooleanQuery(query) || likelyPastedSearchText(query)) {
    return undefined;
  }

  const relaxedSearchText = buildPagedSearchText(query);
  // Keep fallback paging stable across page clicks. If this window grows with
  // skip, later pages can report a different total and make the frontend page
  // count appear to shrink or hang on fallback-heavy searches.
  const candidateTop = 500;
  const orderby = getDocumentSortOrderBy(request.sort);
  const filterFallbackCandidatesLocally = hasUserSelectedFilters(request.filters);
  const fallbackCandidateRequest = filterFallbackCandidatesLocally
    ? { ...request, filters: {} }
    : { ...request, query };
  const fallbackCandidateFilter = filterFallbackCandidatesLocally
    ? buildDocumentFilter(normalizeFilters())
    : baseFilter;
  const relaxedBody: Record<string, unknown> = {
    search: relaxedSearchText,
    queryType: 'simple',
    filter: fallbackCandidateFilter,
    select: withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS),
    top: candidateTop,
    skip: 0,
    count: true
  };

  if (orderby) {
    relaxedBody.orderby = orderby;
  }

  const [relaxedResult, exactHits, contentHits, fuzzyHits] = await Promise.all([
    timedDocumentSearchIndex(env, relaxedBody, 'documents-paged-relaxed-fallback', context),
    exactDocumentCandidates(env, { ...fallbackCandidateRequest, query }, candidateTop, context),
    contentDocumentCandidates(env, { ...fallbackCandidateRequest, query }, Math.min(candidateTop, 80), context),
    fuzzyDocumentCandidates(env, { ...fallbackCandidateRequest, query }, candidateTop, context)
  ]);

  const candidates = uniqueDocumentHits([
    ...mapSearchHits(relaxedResult.value as any[]),
    ...exactHits,
    ...contentHits,
    ...fuzzyHits
  ]);
  const filteredCandidates = candidates.filter((item) =>
    hitMatchesRequestedFilters(item, request.filters) &&
    (passesDocumentRelevanceGuard(env, query, item) ||
      shouldPromoteExactDocument(query, item, 60))
  );
  const ranked = sortDocumentHits(query, filteredCandidates, request.sort);
  const pageResults = ranked.slice(skip, skip + requestedTop);
  const resultWindow = ranked.length;
  const hasMore = skip + pageResults.length < resultWindow;

  emitTelemetry(context, {
    event: 'documents-paged-relaxed-fallback-applied',
    queryTokenCount: meaningfulTokens(query).length,
    candidateCount: candidates.length,
    filteredCount: filteredCandidates.length,
    returned: pageResults.length
  });

  if (resultWindow === 0) {
    return undefined;
  }

  return {
    results: pageResults,
    totalCount: resultWindow,
    facets: request.includeFacets ? buildFacetsFromHits(ranked) : {},
    page: {
      top: requestedTop,
      skip,
      returned: pageResults.length,
      totalAvailable: resultWindow,
      resultWindow,
      capped: candidates.length >= candidateTop,
      hasMore,
      nextSkip: hasMore ? skip + pageResults.length : undefined
    }
  };
};

const domainAcronymPagedFallback = async (
  env: SearchEnv,
  request: SearchRequest,
  requestedTop: number,
  skip: number,
  query: string,
  baseFilter: string,
  context?: SearchExecutionContext
): Promise<SearchPageResponse | undefined> => {
  const coreQuery = domainAcronymCoreQuery(query);
  if (!coreQuery) return undefined;

  const candidateTop = 500;
  const result = await timedDocumentSearchIndex(env, {
    search: coreQuery,
    queryType: 'simple',
    searchMode: 'any',
    filter: baseFilter,
    select: withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS),
    top: candidateTop,
    skip: 0,
    count: true
  }, 'documents-paged-domain-acronym-fallback', context);

  const candidates = mapSearchHits(result.value as any[])
    .filter((item) =>
      hitMatchesRequestedFilters(item, request.filters) &&
      passesDocumentRelevanceGuard(env, coreQuery, item)
    );
  const ranked = sortDocumentHits(query, uniqueDocumentHits(candidates), request.sort);
  if (ranked.length === 0) return undefined;

  const pageResults = ranked.slice(skip, skip + requestedTop);
  const hasMore = skip + pageResults.length < ranked.length;

  emitTelemetry(context, {
    event: 'documents-paged-domain-acronym-fallback-applied',
    query,
    coreQuery,
    candidateCount: candidates.length,
    returned: pageResults.length
  });

  return {
    results: pageResults,
    totalCount: ranked.length,
    facets: request.includeFacets ? buildFacetsFromHits(ranked) : {},
    page: {
      top: requestedTop,
      skip,
      returned: pageResults.length,
      totalAvailable: ranked.length,
      resultWindow: ranked.length,
      capped: candidates.length >= candidateTop,
      hasMore,
      nextSkip: hasMore ? skip + pageResults.length : undefined
    }
  };
};

export const searchDocumentsPaged = async (
  env: SearchEnv,
  request: SearchRequest,
  context?: SearchExecutionContext
): Promise<SearchPageResponse> => {
  const requestedTop = normalizePageTop(env, request.top);
  const skip = normalizeSkip(request.skip);
  const query = request.query && request.query.trim() ? request.query.trim() : '*';

  const isBooleanQuery = isExplicitBooleanQuery(query);
  const isPrecisionQuery = !isWildcardQuery(query) && !isBooleanQuery && likelyPastedSearchText(query);
  if (isLowInformationDocumentSearch(query)) {
    return {
      results: [],
      totalCount: 0,
      facets: {},
      page: {
        top: requestedTop,
        skip,
        returned: 0,
        totalAvailable: 0,
        resultWindow: 0,
        capped: false,
        hasMore: false
      }
    };
  }

  const contentPage = !isBooleanQuery
    ? await searchContentDocumentsPaged(env, request, requestedTop, skip, context)
    : undefined;
  if (contentPage) {
    return contentPage;
  }

  if (shouldSuppressDocumentSearchQuery(query)) {
    return {
      results: [],
      totalCount: 0,
      facets: {},
      page: {
        top: requestedTop,
        skip,
        returned: 0,
        totalAvailable: 0,
        resultWindow: 0,
        capped: false,
        hasMore: false
      }
    };
  }

  const baseFilter = buildDocumentFilter(normalizeFilters(request.filters));
  const strictTermMode = shouldRequireAllSearchTerms(query);
  const searchText = strictTermMode ? buildStrictTermSearchText(query) : buildPagedSearchText(query);
  const body: Record<string, unknown> = {
    search: searchText,
    queryType: isBooleanQuery ? 'full' : 'simple',
    ...(isBooleanQuery || isPrecisionQuery || strictTermMode ? { searchMode: 'all' } : {}),
    filter: baseFilter,
    select: withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS),
    top: requestedTop,
    skip,
    count: request.includeTotalCount !== false
  };

  if (request.includeFacets) {
    body.facets = buildFacetExpressions(env);
  }

  const orderby = getDocumentSortOrderBy(request.sort);
  if (orderby) {
    body.orderby = orderby;
  }

  if (orderby && (request.sort === 'newest' || request.sort === 'oldest')) {
    const datedResult = await timedDocumentSearchIndex(env, {
      ...body,
      filter: appendDocumentFilterClause(baseFilter, 'publishedDate ne null'),
      facets: undefined
    }, 'documents-paged-dated', context);
    const datedResults = mapSearchHits(datedResult.value as any[]);
    const datedCount = readODataCount(datedResult);
    const results = datedResults.slice();
    let undatedCount = 0;

    if (results.length < requestedTop) {
      const undatedResult = await timedDocumentSearchIndex(env, {
        ...body,
        filter: appendDocumentFilterClause(baseFilter, 'publishedDate eq null'),
        facets: undefined,
        orderby: 'search.score() desc',
        top: requestedTop - results.length,
        skip: Math.max(0, skip - datedCount)
      }, 'documents-paged-undated', context);
      undatedCount = readODataCount(undatedResult);
      results.push(...mapSearchHits(undatedResult.value as any[]));
    } else if (request.includeTotalCount !== false) {
      const undatedCountResult = await timedDocumentSearchIndex(env, {
        ...body,
        filter: appendDocumentFilterClause(baseFilter, 'publishedDate eq null'),
        facets: undefined,
        select: 'id',
        orderby: undefined,
        top: 1,
        skip: 0
      }, 'documents-paged-undated-count', context);
      undatedCount = readODataCount(undatedCountResult);
    }

    let facets: Record<string, SearchFacetValue[]> = {};
    if (request.includeFacets) {
      const facetResult = await timedDocumentSearchIndex(env, {
        search: searchText,
        queryType: isBooleanQuery ? 'full' : 'simple',
        ...(isBooleanQuery || isPrecisionQuery || strictTermMode ? { searchMode: 'all' } : {}),
        filter: baseFilter,
        select: 'id',
        top: 1,
        skip: 0,
        count: false,
        facets: buildFacetExpressions(env)
      }, 'documents-paged-date-facets', context);
      facets = mapFacets(facetResult['@search.facets']);
    }

    const totalAvailable = datedCount + undatedCount;
    if (strictTermMode && totalAvailable < requestedTop) {
      const domainFallback = await domainAcronymPagedFallback(env, request, requestedTop, skip, query, baseFilter, context);
      if (domainFallback && domainFallback.totalCount > totalAvailable) {
        return domainFallback;
      }
    }

    if (strictTermMode && totalAvailable === 0) {
      const structuredFallback = await structuredAuthorTopicPagedSearch(env, request, requestedTop, skip, query, baseFilter, context);
      if (structuredFallback) {
        return structuredFallback;
      }

      if (shouldAvoidRelaxedNoResultFallback(query)) {
        return {
          results: [],
          totalCount: 0,
          facets: {},
          page: {
            top: requestedTop,
            skip,
            returned: 0,
            totalAvailable: 0,
            resultWindow: 0,
            capped: false,
            hasMore: false
          }
        };
      }

      const fallback = await relaxedPagedFallback(env, request, requestedTop, skip, query, baseFilter, context);
      if (fallback) {
        return fallback;
      }
    }

    const hasMore = skip + results.length < totalAvailable;

    return {
      results,
      totalCount: totalAvailable,
      facets,
      page: {
        top: requestedTop,
        skip,
        returned: results.length,
        totalAvailable,
        resultWindow: totalAvailable,
        capped: false,
        hasMore,
        nextSkip: hasMore ? skip + results.length : undefined
      }
    };
  }

  const result = await timedDocumentSearchIndex(env, body, 'documents-paged', context);
  const results = mapSearchHits(result.value as any[]);
  const totalCount = readODataCount(result, results.length + skip);

  if (strictTermMode && totalCount < requestedTop) {
    const domainFallback = await domainAcronymPagedFallback(env, request, requestedTop, skip, query, baseFilter, context);
    if (domainFallback && domainFallback.totalCount > totalCount) {
      return domainFallback;
    }
  }

  if (strictTermMode && totalCount === 0) {
    const structuredFallback = await structuredAuthorTopicPagedSearch(env, request, requestedTop, skip, query, baseFilter, context);
    if (structuredFallback) {
      return structuredFallback;
    }

    if (shouldAvoidRelaxedNoResultFallback(query)) {
      return {
        results: [],
        totalCount: 0,
        facets: {},
        page: {
          top: requestedTop,
          skip,
          returned: 0,
          totalAvailable: 0,
          resultWindow: 0,
          capped: false,
          hasMore: false
        }
      };
    }

    const fallback = await relaxedPagedFallback(env, request, requestedTop, skip, query, baseFilter, context);
    if (fallback) {
      return fallback;
    }
  }

  const promotedTitleFileNameCandidates = skip === 0 && !hasUserSelectedFilters(request.filters) && !isWildcardQuery(query) && !isBooleanQuery && !isPrecisionQuery
    ? await titleFileNameDocumentCandidates(env, request, requestedTop, context)
    : [];
  const pageResults = promotedTitleFileNameCandidates.length > 0
    ? mergePromotedDocumentHits(
      query,
      results,
      promotedTitleFileNameCandidates,
      requestedTop,
      toNumber(env.AZURE_SEARCH_EXACT_DOCUMENT_MIN_SCORE, 80)
    )
    : results;
  const totalAvailable = Math.max(totalCount, skip + pageResults.length);
  const hasMore = skip + pageResults.length < totalAvailable;

  return {
    results: pageResults,
    totalCount: totalAvailable,
    facets: request.includeFacets ? mapFacets(result['@search.facets']) : {},
    page: {
      top: requestedTop,
      skip,
      returned: pageResults.length,
      totalAvailable,
      resultWindow: totalAvailable,
      capped: false,
      hasMore,
      nextSkip: hasMore ? skip + pageResults.length : undefined
    }
  };
};

export const searchDocuments = async (
  env: SearchEnv,
  request: SearchRequest,
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  if (shouldSuppressDocumentSearchQuery(request.query || '') || isLowInformationDocumentSearch(request.query || '')) {
    return [];
  }

  const top = normalizeTop(env, request.top);
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  const isBooleanQuery = isExplicitBooleanQuery(query);

  if (!isWildcardQuery(query) && !isBooleanQuery && likelyPastedDocumentText(query)) {
    const mainSemanticTop = Math.min(top, toNumber(env.AZURE_SEARCH_MAIN_SEMANTIC_TOP, 10));
    return semanticSearch(env, { ...request, query, top: mainSemanticTop }, context);
  }

  const result = await timedDocumentSearchIndex(env, {
    search: query,
    queryType: isBooleanQuery ? 'full' : 'simple',
    ...(isBooleanQuery ? { searchMode: 'all' } : {}),
    filter: buildDocumentFilter(normalizeFilters(request.filters)),
    select: withNormalizedFilterSelectFields(DOCUMENT_SELECT_FIELDS),
    top: isWildcardQuery(query) ? top : Math.max(top * 3, 20)
  }, 'documents', context);

  const hits = mapSearchHits(result.value as any[])
    .filter((item) => passesDocumentRelevanceGuard(env, query, item))
    .slice(0, top);

  const [exactHits, contentHits, fuzzyHits] = await Promise.all([
    exactDocumentCandidates(env, request, top, context),
    contentDocumentCandidates(env, request, top, context),
    hits.length < top ? fuzzyDocumentCandidates(env, request, top, context) : Promise.resolve([])
  ]);

  const primaryHits = uniqueDocumentHits([...hits, ...contentHits, ...fuzzyHits]);
  return mergePromotedDocumentHits(
    query,
    primaryHits,
    exactHits,
    top,
    toNumber(env.AZURE_SEARCH_EXACT_DOCUMENT_MIN_SCORE, 80)
  );
};

const dedupeByDocument = (items: any[] = []): any[] => {
  const best = new Map<string, any>();
  items.forEach((item) => {
    const key = item.documentId || item.id;
    const score = item['@search.rerankerScore'] ?? item['@search.score'] ?? 0;
    const existing = best.get(key);
    const existingScore = existing ? (existing['@search.rerankerScore'] ?? existing['@search.score'] ?? 0) : -Infinity;
    if (!existing || score > existingScore) best.set(key, item);
  });
  return Array.from(best.values());
};

const dedupeContentHitsByDocument = (items: any[] = []): any[] => {
  const grouped = new Map<string, { best: any; score: number; windows: Set<string> }>();

  items.forEach((item) => {
    const key = item.documentId || item.id;
    if (!key) return;

    const score = item['@search.rerankerScore'] ?? item['@search.score'] ?? 0;
    const existing = grouped.get(key);
    const windowKey = String(item.__contentWindowKey || '').trim();

    if (!existing) {
      grouped.set(key, {
        best: item,
        score,
        windows: new Set(windowKey ? [windowKey] : [])
      });
      return;
    }

    if (windowKey) {
      existing.windows.add(windowKey);
    }

    if (score > existing.score) {
      existing.best = item;
      existing.score = score;
    }
  });

  return Array.from(grouped.values()).map((entry) => ({
    ...entry.best,
    contentWindowHits: entry.windows.size
  }));
};

const lexicalContentSearch = async (
  env: SearchEnv,
  request: SearchRequest,
  top: number,
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  const result = await timedSearchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
    search: query,
    queryType: 'full',
    searchMode: 'all',
    filter: buildDocumentFilter(normalizeFilters(request.filters)),
    select: withNormalizedFilterSelectFields('id,documentId,listItemId,title,fileName,webUrl,status,description,chunkText,documentType,bu,department,diseaseArea,therapyArea,client,region,authors'),
    top: Math.max(top * 5, 50)
  }, 'lexical-content', context);

  return mapSearchHits(dedupeByDocument(result.value as any[]))
    .slice(0, top);
};

export const semanticSearch = async (
  env: SearchEnv,
  request: SearchRequest,
  context?: SearchExecutionContext
): Promise<SearchHit[]> => {
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  if (shouldSuppressDocumentSearchQuery(query) || isLowInformationDocumentSearch(query)) {
    return [];
  }

  const top = normalizeTop(env, request.top);
  if (isExplicitBooleanQuery(query)) {
    return lexicalContentSearch(env, request, top, context);
  }

  if (!isSemanticSearchEnabled(env)) {
    emitTelemetry(context, {
      level: 'warn',
      event: 'semantic-search-disabled',
      fallback: 'lexical-content'
    });
    return lexicalContentSearch(env, request, top, context);
  }

  const vector = query === '*' ? undefined : (await embedTextsInBatches(env, [query]))[0];
  const body: Record<string, unknown> = {
    search: query,
    vectorFilterMode: 'preFilter',
    queryType: 'semantic',
    semanticConfiguration: env.AZURE_SEARCH_SEMANTIC_CONFIG,
    captions: 'extractive',
    filter: buildDocumentFilter(normalizeFilters(request.filters)),
    select: withNormalizedFilterSelectFields('id,documentId,listItemId,title,fileName,webUrl,status,description,chunkText,documentType,bu,department,diseaseArea,therapyArea,client,region,authors'),
    top: Math.max(top * 3, 20)
  };

  if (vector) {
    body.vectorQueries = [
      {
        kind: 'vector',
        vector,
        fields: env.AZURE_SEARCH_VECTOR_FIELD || 'chunkVector',
        k: toNumber(env.AZURE_SEARCH_HYBRID_K, 50),
        exhaustive: false
      }
    ];
  }

  let result: any;
  try {
    result = await timedSearchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, body, 'semantic-content', context);
  } catch (error) {
    if (!isSemanticBillingError(error)) {
      throw error;
    }

    emitTelemetry(context, {
      level: 'warn',
      event: 'semantic-search-billing-fallback',
      fallback: 'lexical-content',
      error: (error as any)?.message || String(error)
    });
    return lexicalContentSearch(env, request, top, context);
  }

  const hits = mapSearchHits(dedupeByDocument(result.value as any[]))
    .filter((item) => passesDocumentRelevanceGuard(env, query, item));
  const exactHits = await exactDocumentCandidates(env, request, top, context);
  const mergedHits = mergePromotedDocumentHits(
    query,
    hits,
    exactHits,
    Math.max(top, hits.length),
    toNumber(env.AZURE_SEARCH_SEMANTIC_EXACT_DOCUMENT_MIN_SCORE, 80)
  ).filter((item) => passesDocumentRelevanceGuard(env, query, item));

  if (likelyPastedDocumentText(query)) {
    const mainSemanticMinRerankerScore = toNumber(env.AZURE_SEARCH_MAIN_SEMANTIC_MIN_RERANKER_SCORE, 2);
    return mergedHits
      .filter((item) =>
        shouldPromoteExactDocument(query, item, toNumber(env.AZURE_SEARCH_SEMANTIC_EXACT_DOCUMENT_MIN_SCORE, 80)) ||
        (typeof item.rerankerScore === 'number' && item.rerankerScore >= mainSemanticMinRerankerScore)
      )
      .slice(0, top);
  }

  return mergedHits.slice(0, top);
};

export const searchPeople = async (
  env: SearchEnv,
  request: SearchRequest,
  context?: SearchExecutionContext
): Promise<any[]> => {
  if (isCredentialDisclosureQuery(request.query || '')) {
    return [];
  }

  const top = normalizeTop(env, request.top);
  const query = request.query && request.query.trim() ? request.query.trim() : '*';
  const result = await timedSearchIndex(env, env.AZURE_SEARCH_PEOPLE_INDEX, {
    search: query,
    filter: 'active eq true',
    select: 'id,personName,email,contacts,allEmails,allText,description,sectionTitles,serviceLine,role,team,bu,region,listItemUrl',
    top: isWildcardQuery(query) ? top : Math.max(top * 3, 20)
  }, 'people', context);

  return (result.value || [])
    .filter((item: any) => passesPeopleRelevanceGuard(query, item))
    .slice(0, top);
};

const documentMetadata = (item: SearchHit): Record<string, string | string[] | undefined> => ({
  status: item.status,
  documentType: item.documentType,
  description: item.description,
  bu: item.bu,
  department: item.department,
  diseaseArea: item.diseaseArea,
  therapyArea: item.therapyArea,
  client: item.client,
  region: item.region,
  authors: item.authors,
  publishedDate: item.publishedDate,
  createdDateTime: item.createdDateTime || item.created,
  lastModifiedDateTime: item.lastModifiedDateTime || item.modified
});

const retrieveDocumentContentSources = async (
  env: SearchEnv,
  query: string,
  top: number,
  filters?: SearchRequest['filters'],
  context?: SearchExecutionContext
): Promise<GroundingSource[]> => {
  const minRerankerScore = toNumber(env.AZURE_SEARCH_CHAT_MIN_RERANKER_SCORE, 1);
  const hits = await semanticSearch(env, { query, top, filters }, context);
  return hits
    .filter((item) => {
      const evidenceText = documentEvidenceText(item);
      const lexicalMatches = lexicalOverlapCount(query, evidenceText);
      const rerankerScore = item.rerankerScore ?? 0;
      return (rerankerScore >= minRerankerScore && lexicalMatches >= 1) || lexicalMatches >= 2;
    })
    .map((item) => ({
      kind: 'document',
      id: item.id,
      documentId: item.documentId,
      listItemId: item.listItemId,
      title: item.title || item.fileName || 'Untitled document',
      url: item.webUrl,
      text: String(item.chunkText || item.description || item.contentPreview || '').slice(0, 850),
      score: item.rerankerScore ?? item.score ?? 0,
      metadata: documentMetadata(item)
    }));
};

const retrieveDocumentMetadataSources = async (
  env: SearchEnv,
  query: string,
  top: number,
  filters?: SearchRequest['filters'],
  context?: SearchExecutionContext
): Promise<GroundingSource[]> => {
  const hits = await searchDocuments(env, { query, top, filters }, context);
  return hits.map((item) => ({
    kind: 'document',
    id: item.id,
    documentId: item.documentId,
    listItemId: item.listItemId,
    title: item.title || item.fileName || 'Untitled document',
    url: item.webUrl,
    text: [
      item.contentPreview ? `Preview: ${item.contentPreview}` : '',
      item.description ? `Description: ${item.description}` : '',
      item.fileName ? `File name: ${item.fileName}` : '',
      item.documentType ? `Document type: ${item.documentType}` : '',
      item.bu ? `BU: ${item.bu}` : '',
      item.department ? `Department: ${item.department}` : '',
      item.client ? `Client: ${item.client}` : '',
      item.region ? `Region: ${item.region}` : '',
      item.therapyArea ? `Therapy area: ${item.therapyArea}` : '',
      item.diseaseArea ? `Disease area: ${item.diseaseArea}` : '',
      item.authors && item.authors.length > 0 ? `Authors: ${item.authors.join(', ')}` : ''
    ].filter(Boolean).join('\n').slice(0, 850),
    score: item.rerankerScore ?? item.score ?? 0,
    metadata: documentMetadata(item)
  }));
};

const shouldTryMetadataScopeFilters = (query: string): boolean => {
  const cleaned = String(query || '').trim();
  return Boolean(cleaned) && !isWildcardQuery(cleaned) && meaningfulTokens(cleaned).length <= 6 && cleaned.length <= 80;
};

const metadataScopeFilterVariants = (query: string): SearchRequest['filters'][] => {
  const cleaned = String(query || '').trim();
  if (!shouldTryMetadataScopeFilters(cleaned)) {
    return [];
  }

  return [
    { bu: cleaned },
    { department: cleaned },
    { client: cleaned },
    { documentType: cleaned },
    { therapyArea: cleaned },
    { diseaseArea: cleaned },
    { region: cleaned }
  ];
};

const retrieveMetadataScopeSources = async (
  env: SearchEnv,
  query: string,
  top: number,
  context?: SearchExecutionContext
): Promise<GroundingSource[]> => {
  const filterVariants = metadataScopeFilterVariants(query);
  if (filterVariants.length === 0) {
    return [];
  }

  const batches = await Promise.all(filterVariants.map((filters) =>
    retrieveDocumentMetadataSources(env, '*', top, filters, context)
  ));

  return dedupeSources(batches.flat())
    .map((source) => ({
      ...source,
      // Exact metadata-scope matches should outrank broad keyword matches in chat lists.
      score: source.score + 1000
    }))
    .slice(0, top);
};

const retrievePeopleSources = async (
  env: SearchEnv,
  query: string,
  top: number,
  context?: SearchExecutionContext
): Promise<GroundingSource[]> => {
  const items = await searchPeople(env, { query, top }, context);
  return items.map((item) => ({
    kind: 'people',
    id: item.id,
    title: item.personName,
    url: item.listItemUrl,
    text: [
      item.serviceLine ? `Service line: ${item.serviceLine}` : '',
      item.team ? `Team: ${focusedSnippet(query, item.team, 700)}` : '',
      item.role ? `Role/text: ${focusedSnippet(query, item.role, 500)}` : '',
      item.allText ? `Details: ${focusedSnippet(query, item.allText, 900)}` : '',
      item.email || item.allEmails ? `Emails: ${focusedSnippet(query, item.email || item.allEmails, 400)}` : ''
    ].filter(Boolean).join('\n').slice(0, 1000),
    score: item['@search.score'] ?? 0,
    metadata: {
      serviceLine: item.serviceLine,
      team: item.team ? focusedSnippet(query, item.team, 300) : item.team,
      role: item.role ? focusedSnippet(query, item.role, 300) : item.role,
      email: item.email || item.allEmails,
      region: item.region,
      bu: item.bu
    }
  }));
};

// Client requested people-directory answers be disabled for KM Assistant for now.
// The retrieval helper above is retained as archived code so the feature can be restored later.
const CHAT_PEOPLE_DIRECTORY_RETRIEVAL_ENABLED = false;
const CHAT_MAX_QUESTION_CHARS = 2500;

const normalizeSourceTitleKey = (value: string): string =>
  String(value || '')
    .toLowerCase()
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\.[a-z0-9]{2,6}$/i, '')
    .replace(/\s+\d{5,}\b/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const sourceKey = (source: GroundingSource): string => {
  const titleKey = normalizeSourceTitleKey(source.title);
  if (source.kind === 'document' && titleKey.length >= 4) {
    return `${source.kind}|title:${titleKey}`;
  }

  const sourceId = String(source.documentId || source.id || '').trim().toLowerCase();
  if (sourceId) {
    return `${source.kind}|id:${sourceId}`;
  }

  return [source.kind, source.url || '', titleKey].join('|').toLowerCase();
};

const dedupeSources = (sources: GroundingSource[]): GroundingSource[] => {
  const byKey = new Map<string, GroundingSource>();

  sources.forEach((source) => {
    const key = sourceKey(source);
    const existing = byKey.get(key);
    if (!existing || source.score > existing.score) {
      byKey.set(key, source);
    }
  });

  return Array.from(byKey.values());
};

const sourcePriorityForIntent = (intentKind: string, source: GroundingSource): number => {
  if (intentKind === 'people') return source.kind === 'people' ? 0 : 1;
  if (intentKind === 'document' || intentKind === 'metadata_listing' || intentKind === 'mixed') {
    return source.kind === 'document' ? 0 : 1;
  }

  return 0;
};

const noEvidenceAnswer = (kind: string, question: string): string => {
  if (kind === 'people') {
    return 'I cannot answer employee-directory or personal-profile questions in KM Assistant right now. I can help search Active Knowledge Hub documents by author name, topic, client, BU, department, therapy area, disease area, or document type.';
  }

  if (kind === 'metadata_listing') {
    return 'I do not see Active Knowledge Hub documents that confidently match that request. The fastest way to narrow it is usually an exact title, author name, client, BU, department, or document type.';
  }

  if (kind === 'mixed') {
    return 'I do not have enough matching Knowledge Hub evidence to answer both parts confidently. A separate document question with a specific topic, client, BU, department, therapy area, disease area, author, or document type should work better.';
  }

  return 'I do not have enough evidence in the Active Knowledge Hub sources to answer that reliably. A more specific title, keyword, client, BU, department, therapy area, or disease area should help me find the right source.';
};

const sourceMetadataText = (source: GroundingSource): string => {
  const metadata = source.metadata || {};
  return [
    metadata.documentType ? `Type: ${metadata.documentType}` : '',
    metadata.bu ? `BU: ${metadata.bu}` : '',
    metadata.client ? `Client: ${metadata.client}` : '',
    metadata.region ? `Geography: ${metadata.region}` : '',
    metadata.therapyArea ? `Therapy area: ${metadata.therapyArea}` : '',
    metadata.diseaseArea ? `Disease area: ${metadata.diseaseArea}` : ''
  ].filter(Boolean).join('; ');
};

const CHAT_GENERIC_RETRIEVAL_TOKENS = new Set([
  'asset', 'assets', 'material', 'materials', 'document', 'documents', 'doc', 'docs',
  'file', 'files', 'source', 'sources', 'reference', 'references', 'link', 'links',
  'share', 'show', 'find', 'fetch', 'pull', 'list', 'give', 'tell', 'about', 'related',
  'relevant', 'matching', 'available', 'current', 'latest', 'newest', 'recent', 'more',
  'summary', 'summarize', 'explain', 'overview',
  'proposal', 'proposals', 'rfp', 'rfps', 'rfi', 'rfis', 'sop', 'sops',
  'case', 'study', 'studies', 'document', 'documents'
]);

const sourceEvidenceForGuard = (source: GroundingSource): string => {
  const metadata = source.metadata || {};
  return [
    source.title,
    source.text,
    ...Object.values(metadata).flatMap((value) => Array.isArray(value) ? value : [value])
  ].filter(Boolean).join(' ');
};

const chatQueryEvidenceTokens = (query: string): string[] =>
  meaningfulTokens(normalizeDomainSearchPhrase(query))
    .map((token) => token.toLowerCase())
    .filter((token) => !CHAT_GENERIC_RETRIEVAL_TOKENS.has(token));

const sourceTokenOverlap = (tokens: string[], source: GroundingSource): number => {
  const evidence = ` ${sourceEvidenceForGuard(source).toLowerCase().replace(/[^a-z0-9]+/g, ' ')} `;
  return tokens.filter((token) => evidence.includes(` ${token} `)).length;
};

const sourceTitleTokenOverlap = (tokens: string[], source: GroundingSource): number => {
  const titleEvidence = ` ${String(source.title || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ')} `;
  return tokens.filter((token) => titleEvidence.includes(` ${token} `)).length;
};

const filterWeakChatSources = (
  query: string,
  sources: GroundingSource[],
  intentKind: string
): GroundingSource[] => {
  if (sources.length === 0 || intentKind === 'capability' || intentKind === 'greeting') {
    return sources;
  }

  const tokens = chatQueryEvidenceTokens(query);
  if (tokens.length === 0) {
    return sources;
  }

  const minimumOverlap = tokens.length === 1 ? 1 : Math.min(2, tokens.length);
  return sources.filter((source) => {
    if (sourceTokenOverlap(tokens, source) >= minimumOverlap) {
      return true;
    }

    const titleOverlap = sourceTitleTokenOverlap(tokens, source);
    return titleOverlap >= Math.min(3, tokens.length)
      || (tokens.length >= 3 && titleOverlap >= Math.ceil(tokens.length * 0.7));
  });
};

const fallbackGroundedAnswer = (question: string, kind: string, sources: GroundingSource[]): string => {
  const documentSources = sources.filter((source) => source.kind === 'document');
  const peopleSources = sources.filter((source) => source.kind === 'people');
  const selectedSources = (kind === 'people' && peopleSources.length > 0 ? peopleSources : documentSources.length > 0 ? documentSources : sources)
    .slice(0, 5);

  if (selectedSources.length === 0) {
    return noEvidenceAnswer(kind, question);
  }

  if (/\b(summarize|summary|brief|briefly|explain|describe|overview|walk me through)\b/i.test(question)
    && /\b(first|top|strongest|best|one|result|source|reference|document|doc|asset)\b/i.test(question)) {
    const source = selectedSources[0];
    const metadata = sourceMetadataText(source);
    const evidence = String(source.text || '').replace(/\s+/g, ' ').trim();
    const snippet = compactEvidenceSnippet(evidence, 420);
    return [
      `Here is a brief of "${source.title}":`,
      '',
      metadata ? `- ${metadata}` : '',
      snippet
        ? `- ${snippet}`
        : '- The available Knowledge Hub preview is limited, so I can only provide a high-level brief from the title and metadata.',
      '',
      'Open the first reference below to review the full asset.'
    ].filter(Boolean).join('\n');
  }

  const lead = 'I found these matching Active Knowledge Hub documents:';

  const lines = selectedSources.map((source, index) => {
    const metadata = sourceMetadataText(source);
    const evidence = String(source.text || '').replace(/\s+/g, ' ').trim();
    const snippet = compactEvidenceSnippet(evidence, 220);
    return [
      `${index + 1}. ${source.title}`,
      metadata ? `   - ${metadata}` : '',
      snippet ? `   - Why it may match: ${snippet}` : ''
    ].filter(Boolean).join('\n');
  });

  const close = selectedSources.length >= 5
    ? 'I kept this to the strongest five matches here. Use the search page or add a tighter filter if you want to browse more results.'
    : 'Open the references below to review the source documents.';

  return [lead, '', ...lines, '', close].join('\n');
};

const fallbackGroundedCitationSources = (
  question: string,
  kind: string,
  sources: GroundingSource[]
): GroundingSource[] => {
  const documentSources = sources.filter((source) => source.kind === 'document');
  const peopleSources = sources.filter((source) => source.kind === 'people');
  const selectedSources = (kind === 'people' && peopleSources.length > 0 ? peopleSources : documentSources.length > 0 ? documentSources : sources)
    .slice(0, 5);

  if (/\b(summarize|summary|brief|briefly|explain|describe|overview|walk me through)\b/i.test(question)
    && /\b(first|top|strongest|best|one|result|source|reference|document|doc|asset)\b/i.test(question)) {
    return selectedSources.slice(0, 1);
  }

  return selectedSources;
};

const sourcesToCitations = (sources: GroundingSource[]): ChatCitation[] =>
  sources.map((source, index) => ({
    number: index + 1,
    kind: source.kind,
    id: source.id,
    documentId: source.documentId,
    listItemId: source.listItemId,
    title: source.title,
    url: source.url,
    score: source.score
  }));

const sourcesToCitationsBySourceNumber = (
  sources: GroundingSource[],
  sourceNumbers: number[]
): ChatCitation[] =>
  sourceNumbers
    .map((sourceNumber) => {
      const source = sources[sourceNumber - 1];
      if (!source) return null;

      return {
        number: sourceNumber,
        kind: source.kind,
        id: source.id,
        documentId: source.documentId,
        listItemId: source.listItemId,
        title: source.title,
        url: source.url,
        score: source.score
      } as ChatCitation;
    })
    .filter((citation): citation is ChatCitation => Boolean(citation));

const limitRepeatedCitationMarkers = (value: string): string => {
  const seen = new Set<string>();
  return String(value || '')
    .replace(/\s*\[([1-9]\d*)\]/g, (match, citationNumber) => {
      if (seen.has(citationNumber)) {
        return '';
      }
      seen.add(citationNumber);
      return match;
    })
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
};

const citationSourcesUsedByAnswer = (answer: string, sources: GroundingSource[]): GroundingSource[] => {
  const numbers = Array.from(String(answer || '').matchAll(/\[([1-9]\d*)\]/g))
    .map((match) => Number(match[1]))
    .filter((value) => Number.isFinite(value) && value >= 1);
  const highestCitation = numbers.length > 0 ? Math.max(...numbers) : 0;

  return highestCitation > 0 ? sources.slice(0, Math.min(highestCitation, sources.length)) : sources;
};

type StructuredChatResponseType = 'answer' | 'no_evidence' | 'refusal' | 'clarification';
type StructuredChatScope = 'knowledge_hub' | 'off_topic' | 'unsafe' | 'unsupported_action' | 'employee_directory_disabled';
type StructuredChatConfidence = 'high' | 'medium' | 'low';

interface StructuredChatAnswer {
  responseType: StructuredChatResponseType;
  scope: StructuredChatScope;
  answerMarkdown: string;
  usedSourceNumbers: number[];
  confidence: StructuredChatConfidence;
  followUpSuggestions: string[];
}

const STRUCTURED_RESPONSE_TYPES = new Set<StructuredChatResponseType>(['answer', 'no_evidence', 'refusal', 'clarification']);
const STRUCTURED_SCOPES = new Set<StructuredChatScope>(['knowledge_hub', 'off_topic', 'unsafe', 'unsupported_action', 'employee_directory_disabled']);
const STRUCTURED_CONFIDENCE_VALUES = new Set<StructuredChatConfidence>(['high', 'medium', 'low']);

const extractJsonObjectText = (value: string): string => {
  const text = String(value || '').trim();
  if (!text) return '';

  const fencedMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fencedMatch) {
    return fencedMatch[1].trim();
  }

  const firstBrace = text.indexOf('{');
  const lastBrace = text.lastIndexOf('}');
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return text.slice(firstBrace, lastBrace + 1).trim();
  }

  return text;
};

const sourceNumbersMentionedInAnswer = (answer: string, sourceCount: number): number[] =>
  Array.from(new Set(
    Array.from(String(answer || '').matchAll(/\[([1-9]\d*)\]/g))
      .map((match) => Number(match[1]))
      .filter((value) => Number.isInteger(value) && value >= 1 && value <= sourceCount)
  ));

const removeInvalidCitationMarkers = (answer: string, sourceCount: number): string =>
  String(answer || '')
    .replace(/\s*\[([1-9]\d*)\]/g, (match, sourceNumber) => {
      const parsed = Number(sourceNumber);
      return Number.isInteger(parsed) && parsed >= 1 && parsed <= sourceCount ? match : '';
    })
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

const normalizeStructuredSourceNumbers = (
  usedSourceNumbers: unknown,
  answer: string,
  sourceCount: number
): number[] => {
  const declared = Array.isArray(usedSourceNumbers)
    ? usedSourceNumbers
        .map((value) => Number(value))
        .filter((value) => Number.isInteger(value) && value >= 1 && value <= sourceCount)
    : [];
  const mentioned = sourceNumbersMentionedInAnswer(answer, sourceCount);
  return Array.from(new Set([...mentioned, ...declared]));
};

const parseStructuredChatAnswer = (
  content: string,
  sourceCount: number
): StructuredChatAnswer | null => {
  let parsed: any;
  try {
    parsed = JSON.parse(extractJsonObjectText(content));
  } catch {
    return null;
  }

  const responseType = STRUCTURED_RESPONSE_TYPES.has(parsed?.responseType)
    ? parsed.responseType as StructuredChatResponseType
    : 'answer';
  const scope = STRUCTURED_SCOPES.has(parsed?.scope)
    ? parsed.scope as StructuredChatScope
    : 'knowledge_hub';
  const confidence = STRUCTURED_CONFIDENCE_VALUES.has(parsed?.confidence)
    ? parsed.confidence as StructuredChatConfidence
    : 'medium';
  const answerMarkdown = limitRepeatedCitationMarkers(
    removeInvalidCitationMarkers(cleanChatAnswer(String(parsed?.answerMarkdown || '')), sourceCount)
  );

  if (!answerMarkdown) return null;

  return {
    responseType,
    scope,
    answerMarkdown,
    usedSourceNumbers: normalizeStructuredSourceNumbers(parsed?.usedSourceNumbers, answerMarkdown, sourceCount),
    confidence,
    followUpSuggestions: Array.isArray(parsed?.followUpSuggestions)
      ? parsed.followUpSuggestions.map((value: unknown) => String(value || '').trim()).filter(Boolean).slice(0, 3)
      : []
  };
};

const isStructuredResponseFormatUnsupported = (error: unknown): boolean => {
  const message = String((error as any)?.message || error || '').toLowerCase();
  return /response_format|json_schema|schema|invalid parameter|unsupported parameter|unknown parameter/.test(message)
    && !/content_filter|responsibleaipolicyviolation/.test(message);
};

const getStructuredChatCompletionResult = async (
  env: SearchEnv,
  messages: Array<{ role: 'system' | 'user' | 'assistant'; content: string }>,
  options: { maxTokens: number },
  context?: SearchExecutionContext,
  intentKind?: string
): Promise<ChatCompletionResult> => {
  try {
    return await getChatCompletionResult(env, messages, {
      maxTokens: options.maxTokens,
      responseFormat: STRUCTURED_CHAT_RESPONSE_FORMAT
    });
  } catch (error) {
    if (!isStructuredResponseFormatUnsupported(error)) {
      throw error;
    }

    emitTelemetry(context, {
      level: 'warn',
      event: 'chat-structured-response-format-fallback',
      intent: intentKind,
      error: (error as any)?.message || String(error)
    });

    return getChatCompletionResult(env, messages, { maxTokens: options.maxTokens });
  }
};

const asksToSummarizeNamedDocument = (question: string): boolean =>
  /\b(summarize|summary|brief|briefly|explain|describe|overview|this\s+(?:doc|document|deck|file|asset))\b/i.test(question);

const looksLikeNamedDocumentLookup = (question: string): boolean => {
  const normalized = String(question || '').replace(/[?!.]+$/g, ' ').trim();
  const tokens = meaningfulTokens(normalized)
    .filter((token) => !CHAT_GENERIC_RETRIEVAL_TOKENS.has(token.toLowerCase()));

  if (tokens.length < 3 || tokens.length > 12) {
    return false;
  }

  return !/\b(who|where|when|why|how many|count|total|compare|difference|latest|newest|oldest)\b/i.test(normalized);
};

const closeTitleMatchSourceNumbers = (question: string, sources: GroundingSource[]): number[] => {
  const queryTokens = meaningfulTokens(question)
    .filter((token) => !CHAT_GENERIC_RETRIEVAL_TOKENS.has(token))
    .filter((token) => token !== 'indegene' || meaningfulTokens(question).length <= 4);

  if (queryTokens.length < 2) return [];

  return sources
    .map((source, index) => {
      const titleText = source.title || '';
      const overlap = Math.max(
        lexicalOverlapCount(queryTokens.join(' '), titleText),
        fuzzyLexicalOverlapCount(queryTokens.join(' '), titleText)
      );
      return {
        sourceNumber: index + 1,
        overlap,
        titleTokenCount: meaningfulTokens(titleText).length
      };
    })
    .filter((candidate) =>
      candidate.overlap >= Math.min(3, queryTokens.length) ||
      (candidate.titleTokenCount >= 3 && candidate.overlap >= Math.ceil(candidate.titleTokenCount * 0.7))
    )
    .sort((left, right) => right.overlap - left.overlap)
    .map((candidate) => candidate.sourceNumber)
    .slice(0, 2);
};

const buildLimitedNamedDocumentAnswer = (
  question: string,
  sources: GroundingSource[]
): { answer: string; sourceNumbers: number[] } | null => {
  const asksForSummary = asksToSummarizeNamedDocument(question);
  if (!asksForSummary && !looksLikeNamedDocumentLookup(question)) return null;

  const sourceNumbers = closeTitleMatchSourceNumbers(question, sources);
  const firstSourceNumber = sourceNumbers[0];
  const source = firstSourceNumber ? sources[firstSourceNumber - 1] : undefined;
  if (!source) return null;

  const metadata = sourceMetadataText(source);
  const asksForMetadata = /\b(type|bu|business unit|client|geography|therapy area|disease area|metadata|tagged|tags|field|fields|details)\b/i.test(question);
  const snippet = compactEvidenceSnippet(source.text, 360);
  const naturalSnippet = snippet
    .replace(/^[A-Z][A-Za-z ]+:\s*/, '')
    .replace(/\s*\[[^\]]*$/g, '')
    .replace(/[<>\[\]]/g, ' ')
    .replace(/\bDindegene\b/gi, 'Indegene')
    .replace(/\s+/g, ' ')
    .trim();
  const shouldShowSnippet = Boolean(naturalSnippet)
    && !/\b(?:all rights reserved|confidential and proprietary|better outcomes delivered|indegene overview)\b/i.test(naturalSnippet)
    && naturalSnippet.length >= 40;
  const titleBasedBrief = source.title
    ? `Based on the title, this appears to be a Knowledge Hub asset about ${source.title.replace(/\s+/g, ' ').trim()}.`
    : 'Based on the available result, this appears to be a Knowledge Hub asset matching the document you asked for.';
  const lines = asksForSummary
    ? [
        `I found "${source.title}" in Knowledge Hub [${firstSourceNumber}].`,
        '',
        shouldShowSnippet
          ? `It appears to be about ${naturalSnippet}`
          : `${titleBasedBrief} The available preview is limited, so I can only give a high-level brief from search evidence unless you open the asset or paste a section.`,
        '',
        'Open the reference below for the full asset. If you paste a section or slide text, I can summarize it more precisely.'
      ]
    : [
        `I found "${source.title}" in Knowledge Hub [${firstSourceNumber}].`,
        asksForMetadata && metadata ? `Its indexed metadata lists ${metadata}.` : '',
        shouldShowSnippet ? `The visible preview suggests it is related to ${naturalSnippet}` : '',
        'Open the reference below to review the full asset.'
      ];

  return {
    answer: cleanChatAnswer(lines.filter(Boolean).join('\n')),
    sourceNumbers: [firstSourceNumber]
  };
};

const alternateRetrievalQuery = (query: string): string => {
  const personNameQuery = extractPersonNameQuery(query);
  if (!personNameQuery) return '';

  const normalizedPerson = personNameQuery.toLowerCase();
  const normalizedQuery = query.toLowerCase().trim();
  return normalizedPerson && normalizedPerson !== normalizedQuery ? personNameQuery : '';
};

const relaxedDocumentTypeRetrievalQuery = (query: string): string => {
  const relaxed = String(query || '')
    .replace(/\b(?:rfps?|rfis?|sops?|standard operating procedures?|proposals?|case stud(?:y|ies)|case|studies?|documents?|docs?|assets?|materials?)\b/gi, ' ')
    .replace(/\b(?:and|or|related|matching|available)\b/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (!relaxed || relaxed.toLowerCase() === String(query || '').trim().toLowerCase()) {
    return '';
  }

  return meaningfulTokens(relaxed).length >= 2 ? relaxed : '';
};

const normalizeChatHistoryContent = (content: string): string => {
  const cleaned = String(content || '')
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  const sourceSplit = cleaned.split(/sources used:/i);
  if (sourceSplit.length < 2) {
    return cleaned.slice(0, 900);
  }

  const answerText = sourceSplit[0].trim().slice(0, 650);
  const sourceText = sourceSplit.slice(1).join('Sources used:').trim().slice(0, 650);
  return [answerText, sourceText ? `Sources used:\n${sourceText}` : '']
    .filter(Boolean)
    .join('\n\n')
    .slice(0, 1300);
};

const normalizeChatHistory = (history?: ChatRequest['history']): NonNullable<ChatRequest['history']> =>
  (history || [])
    .map((message) => ({
      role: message.role === 'assistant' ? 'assistant' : 'user',
      content: normalizeChatHistoryContent(String(message.content || ''))
    }))
    .filter((message) => message.content)
    .slice(-5);

const normalizeConversationSummary = (value?: string): string =>
  String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
    .slice(0, 1400);

const asksForPreviousSources = (question: string): boolean => {
  const normalized = question.toLowerCase().replace(/\s+/g, ' ').trim();
  const asksAboutSourceArtifact = /\b(source|sources|reference|references|citation|citations|document|documents|file|files|link|links)\b/.test(normalized);
  const pointsToPreviousAnswer = /\b(use|used|using|cite|cited|based|answer|previous|above|that|those|one|only|just|share|shared|gave|provide|provided)\b/.test(normalized);

  if (asksAboutSourceArtifact && pointsToPreviousAnswer) return true;

  return /\b(just|only)\s+(one|1)\s+(source|reference|citation|document|doc|link)\b/.test(normalized)
    || /\b(what|why)\s+(did|do)\s+(you|u)\s+(share|shared|cite|cited|use|used|give|gave|provide|provided)\b/.test(normalized)
    || /\bwhere\s+(did|does)\s+(that|this|the answer)\s+(come|came)\s+from\b/.test(normalized);
};

const parseHistorySources = (content: string): Array<{ title: string; url?: string }> => {
  const sourceBlock = content.split(/sources used:/i)[1] || '';
  if (!sourceBlock.trim()) return [];

  const entries: string[] = [];
  let current = '';

  sourceBlock
    .replace(/\r\n/g, '\n')
    .split('\n')
    .map((line) => line.trim())
    .filter(Boolean)
    .forEach((line) => {
      const startsEntry = /^[-*]\s+/.test(line) || /^\d+[\).]\s+/.test(line);
      const cleaned = line.replace(/^[-*]\s+/, '').replace(/^\d+[\).]\s+/, '').trim();
      if (startsEntry) {
        if (current) entries.push(current);
        current = cleaned;
      } else if (current) {
        current = `${current} ${cleaned}`.trim();
      }
    });

  if (current) entries.push(current);

  const seen = new Set<string>();
  return entries
    .map((entry) => {
      const parenUrl = entry.match(/\((https?:\/\/[^)\s]+)\)\s*$/i);
      const dashUrl = entry.match(/^(.*?)(?:\s+-\s+(https?:\/\/\S+))$/i);
      const url = parenUrl?.[1]?.trim() || dashUrl?.[2]?.trim();
      const title = (parenUrl
        ? entry.slice(0, parenUrl.index).trim()
        : dashUrl
          ? dashUrl[1].trim()
          : entry
      )
        .replace(/\s*\(?https?:?[^)\s]*$/i, '')
        .replace(/^[-*]\s+/, '')
        .replace(/\s+/g, ' ')
        .trim();
      return { title: title || url || 'Source', url };
    })
    .filter((source) => {
      const key = `${source.title.toLowerCase()}|${source.url || ''}`;
      if ((!source.title && !source.url) || seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .slice(0, 8);
};

const requestedHistorySourceOrdinal = (question: string): number => {
  const normalized = question.toLowerCase();
  if (/\b(second|2nd|two)\b/.test(normalized)) return 2;
  if (/\b(third|3rd|three)\b/.test(normalized)) return 3;
  if (/\b(fourth|4th|four)\b/.test(normalized)) return 4;
  if (/\b(fifth|5th|five)\b/.test(normalized)) return 5;
  return 1;
};

const asksForPreviousSourceContent = (question: string): boolean =>
  /\b(summarize|summary|explain|describe|overview|more|details|tell me more|brief|briefly|walk me through)\b/i.test(question)
  && /\b(first|second|third|fourth|fifth|1st|2nd|3rd|4th|5th|one|that|this|source|reference|document|doc|asset|found)\b/i.test(question);

const previousHistorySourceFocus = (
  question: string,
  history: NonNullable<ChatRequest['history']>
): { title: string; url?: string } | null => {
  if (!asksForPreviousSourceContent(question)) return null;

  const assistantWithSources = [...history].reverse().find((message) =>
    message.role === 'assistant' && /sources used:/i.test(message.content)
  );
  if (!assistantWithSources) return null;

  const sources = parseHistorySources(assistantWithSources.content);
  return sources[requestedHistorySourceOrdinal(question) - 1] || sources[0] || null;
};

const sourceTitleMatchesFocus = (focusTitle: string, source: GroundingSource): boolean => {
  const focusTokens = meaningfulTokens(focusTitle).filter((token) => !CHAT_GENERIC_RETRIEVAL_TOKENS.has(token));
  if (focusTokens.length === 0) return false;

  const title = source.title || '';
  const overlap = Math.max(
    lexicalOverlapCount(focusTokens.join(' '), title),
    fuzzyLexicalOverlapCount(focusTokens.join(' '), title)
  );
  return overlap >= Math.min(3, focusTokens.length) || overlap >= Math.ceil(focusTokens.length * 0.7);
};

const buildPreviousSourceUnavailableResponse = (focus: { title: string; url?: string }): ChatResponse => ({
  answer: `I know you mean the previous reference "${focus.title}", but I could not retrieve enough matching Knowledge Hub text for that source in this turn to summarize it safely. Open the reference below, or paste the section you want summarized and I can help with it. [1]`,
  citations: [{
    number: 1,
    kind: 'document',
    title: focus.title,
    url: focus.url
  }]
});

const buildPreviousSourceResponse = (
  question: string,
  history: NonNullable<ChatRequest['history']>
): ChatResponse | null => {
  if (!asksForPreviousSources(question)) return null;

  const assistantWithSources = [...history].reverse().find((message) =>
    message.role === 'assistant' && /sources used:/i.test(message.content)
  );
  if (!assistantWithSources) {
    return {
      answer: 'I do not have a previous cited Knowledge Hub answer in this chat to reference yet. Ask me to find or explain a specific topic, client, BU, department, therapy area, disease area, document type, or document title first, and I can cite the matching sources.',
      citations: []
    };
  }

  const sources = parseHistorySources(assistantWithSources.content);
  if (sources.length === 0) {
    return {
      answer: 'I do not have a previous cited Knowledge Hub source in this chat to show yet. Ask me to find or explain a specific Knowledge Hub topic or document first, and I can cite the matching sources.',
      citations: []
    };
  }
  const visibleSources = sources.slice(0, 3);

  return {
    answer: [
      visibleSources.length === 1
        ? 'I used one Knowledge Hub reference for the previous answer because it was the source retrieved for that response:'
        : 'I used these strongest Knowledge Hub references from the previous answer:',
      ...visibleSources.map((source, index) => `${index + 1}. ${source.title}`)
    ].join('\n'),
    citations: visibleSources.map((source, index) => ({
      number: index + 1,
      kind: 'document',
      title: source.title,
      url: source.url
    }))
  };
};

const isTotalActiveDocumentCountQuestion = (question: string): boolean => {
  const normalized = question.toLowerCase().replace(/\s+/g, ' ').trim();
  return /\b(total|count|number|no\.?|how many)\b.*\b(documents?|docuemnts?|docuemtns?|docs?|files?)\b/.test(normalized)
    || /\b(how many)\b.*\b(active)?\s*(knowledge hub)?\s*(documents?|files?)\b/.test(normalized);
};

export const chat = async (
  env: SearchEnv,
  request: ChatRequest,
  context?: SearchExecutionContext
): Promise<ChatResponse> => {
  const rawQuestion = String(request.question || '').trim();
  if (rawQuestion.length > CHAT_MAX_QUESTION_CHARS) {
    emitTelemetry(context, {
      level: 'info',
      event: 'chat-input-too-long',
      intent: 'unsupported_action',
      sourceCount: 0,
      inputLength: rawQuestion.length,
      maxLength: CHAT_MAX_QUESTION_CHARS
    });
    return {
      answer: `That message is too long for KM Assistant right now. Please keep it under ${CHAT_MAX_QUESTION_CHARS.toLocaleString()} characters. Try pasting only the key paragraph, document title, or specific question you want answered.`,
      citations: []
    };
  }

  const question = rawQuestion;
  const history = normalizeChatHistory(request.history);
  const conversationSummary = normalizeConversationSummary(request.conversationSummary);
  const intentHistory = conversationSummary
    ? [
      {
        role: 'assistant',
        content: `Current chat memory. Use only to resolve follow-ups, not as factual evidence:\n${conversationSummary}`
      },
      ...history
    ]
    : history;
  const previousSourceResponse = buildPreviousSourceResponse(question, history);
  const previousSourceFocus = previousHistorySourceFocus(question, history);

  if (isTotalActiveDocumentCountQuestion(question)) {
    const page = await searchDocumentsPaged(env, {
      query: '*',
      top: 1,
      skip: 0,
      includeTotalCount: true,
      includeFacets: false
    }, context);

    return {
      answer: `I currently see ${page.totalCount} Active Knowledge Hub document${page.totalCount === 1 ? '' : 's'} in the search index.`,
      citations: []
    };
  }

  const intent = classifyChatIntent(question, intentHistory);
  const retrievalQuery = intent.retrievalQuery || question;
  const scopedPromptQuestion = retrievalQuery !== question
    ? [
      `Original user question: ${question}`,
      '',
      `Answer only this in-scope Knowledge Hub portion: ${retrievalQuery}`,
      '',
      'Do not answer unrelated general-world, sports, politics, celebrity, or people-directory parts of the original question.'
    ].join('\n')
    : question;
  const promptQuestion = intent.conversationContext
    ? [
      `Current user question: ${scopedPromptQuestion}`,
      '',
      'Recent conversation context:',
      intent.conversationContext
    ].join('\n')
    : scopedPromptQuestion;

  if (intent.directAnswer) {
    emitTelemetry(context, {
      level: 'info',
      event: 'chat-direct-answer',
      intent: intent.kind,
      sourceCount: 0,
      noEvidence: false
    });
    return {
      answer: intent.directAnswer,
      citations: []
    };
  }

  if (previousSourceResponse) {
    emitTelemetry(context, {
      level: 'info',
      event: 'chat-previous-source-answer',
      intent: 'document',
      sourceCount: previousSourceResponse.citations.length,
      noEvidence: false
    });
    return previousSourceResponse;
  }

  const retrievals: Array<Promise<GroundingSource[]>> = [];
  const configuredDocumentTop = toNumber(env.AZURE_SEARCH_CHAT_DOCUMENT_TOP, 4);
  const documentTop = intent.kind === 'metadata_listing'
    ? Math.max(configuredDocumentTop, 10)
    : configuredDocumentTop;
  const peopleTop = toNumber(env.AZURE_SEARCH_CHAT_PEOPLE_TOP, 3);
  const secondaryQuery = alternateRetrievalQuery(retrievalQuery);
  const relaxedDocumentTypeQuery = relaxedDocumentTypeRetrievalQuery(retrievalQuery);

  if (intent.shouldRetrieveDocuments) {
    retrievals.push(retrieveDocumentContentSources(env, retrievalQuery, documentTop, {}, context));
    if (!intent.shouldUseDocumentMetadata) {
      retrievals.push(retrieveDocumentMetadataSources(env, retrievalQuery, documentTop, {}, context));
    }
    if (secondaryQuery) {
      retrievals.push(retrieveDocumentMetadataSources(env, secondaryQuery, documentTop, {}, context));
    }
  }

  if (intent.shouldUseDocumentMetadata) {
    retrievals.push(retrieveMetadataScopeSources(env, retrievalQuery, documentTop, context));
    retrievals.push(retrieveDocumentMetadataSources(env, retrievalQuery, documentTop, {}, context));
    if (secondaryQuery) {
      retrievals.push(retrieveDocumentMetadataSources(env, secondaryQuery, documentTop, {}, context));
    }
  }

  if (intent.shouldRetrieveDocuments && relaxedDocumentTypeQuery) {
    retrievals.push(retrieveDocumentContentSources(env, relaxedDocumentTypeQuery, documentTop, {}, context));
    retrievals.push(retrieveDocumentMetadataSources(env, relaxedDocumentTypeQuery, documentTop, {}, context));
  }

  if (CHAT_PEOPLE_DIRECTORY_RETRIEVAL_ENABLED && intent.shouldRetrievePeople) {
    retrievals.push(retrievePeopleSources(env, retrievalQuery, peopleTop, context));
    if (secondaryQuery) {
      retrievals.push(retrievePeopleSources(env, secondaryQuery, peopleTop, context));
    }
  }

  const retrievalStartedAt = Date.now();
  const settledRetrievalResults = await Promise.allSettled(retrievals);
  const failedRetrievals = settledRetrievalResults.filter((result) => result.status === 'rejected');
  if (failedRetrievals.length > 0) {
    emitTelemetry(context, {
      level: 'warn',
      event: 'chat-retrieval-partial-failed',
      intent: intent.kind,
      failedRetrievalCount: failedRetrievals.length,
      totalRetrievalCount: settledRetrievalResults.length,
      error: String((failedRetrievals[0] as PromiseRejectedResult).reason?.message || (failedRetrievals[0] as PromiseRejectedResult).reason || '')
    });
  }
  const retrievalResults = settledRetrievalResults
    .filter((result): result is PromiseFulfilledResult<GroundingSource[]> => result.status === 'fulfilled')
    .map((result) => result.value);
  const candidateSourceLimit = intent.kind === 'metadata_listing'
    ? Math.max(toNumber(env.AZURE_SEARCH_CHAT_TOP, 5), 8)
    : toNumber(env.AZURE_SEARCH_CHAT_TOP, 5);
  const unrestrictedCandidateSources = dedupeSources(retrievalResults.flat())
    .filter((source) => CHAT_PEOPLE_DIRECTORY_RETRIEVAL_ENABLED || source.kind === 'document')
    .filter((source) => source.text.trim())
    .sort((left, right) => {
      const priorityDifference = sourcePriorityForIntent(intent.kind, left) - sourcePriorityForIntent(intent.kind, right);
      return priorityDifference || right.score - left.score;
    })
    .slice(0, candidateSourceLimit);
  const candidateSources = previousSourceFocus
    ? unrestrictedCandidateSources.filter((source) => sourceTitleMatchesFocus(previousSourceFocus.title, source))
    : unrestrictedCandidateSources;

  if (previousSourceFocus && candidateSources.length === 0) {
    emitTelemetry(context, {
      level: 'info',
      event: 'chat-previous-source-lock-no-match',
      intent: intent.kind,
      sourceCount: 0,
      focusedSourceTitle: previousSourceFocus.title
    });
    return buildPreviousSourceUnavailableResponse(previousSourceFocus);
  }

  const rankedSources = filterWeakChatSources(previousSourceFocus?.title || retrievalQuery, candidateSources, intent.kind);
  const answerSources = rankedSources.slice(0, 5);

  emitTelemetry(context, {
    level: 'info',
    event: 'chat-retrieval-completed',
    intent: intent.kind,
    sourceCount: answerSources.length,
    candidateSourceCount: candidateSources.length,
    retrievalDurationMs: Date.now() - retrievalStartedAt
  });

  if (rankedSources.length === 0) {
    emitTelemetry(context, {
      level: 'info',
      event: 'chat-no-evidence',
      intent: intent.kind,
      sourceCount: 0,
      noEvidence: true
    });
    if (!['greeting', 'capability'].includes(intent.kind)) {
      return {
        answer: noEvidenceAnswer(intent.kind, question),
        citations: []
      };
    }
    try {
      const structuredCompletion = await getStructuredChatCompletionResult(
        env,
        buildStructuredNoEvidenceChatPrompt(promptQuestion, intent.kind),
        { maxTokens: 420 },
        context,
        intent.kind
      );
      const structuredAnswer = parseStructuredChatAnswer(structuredCompletion.content, 0);
      if (structuredAnswer) {
        return {
          answer: cleanChatAnswer(stripBracketCitations(structuredAnswer.answerMarkdown)),
          citations: [],
          usage: {
            promptTokens: structuredCompletion.usage?.prompt_tokens,
            completionTokens: structuredCompletion.usage?.completion_tokens,
            totalTokens: structuredCompletion.usage?.total_tokens,
            deployment: structuredCompletion.deployment
          }
        };
      }

      const completion = await getChatCompletionResult(env, buildNoEvidenceChatPrompt(promptQuestion, intent.kind), { maxTokens: 420 });
      return {
        answer: cleanChatAnswer(stripBracketCitations(completion.content)),
        citations: [],
        usage: {
          promptTokens: completion.usage?.prompt_tokens,
          completionTokens: completion.usage?.completion_tokens,
          totalTokens: completion.usage?.total_tokens,
          deployment: completion.deployment
        }
      };
    } catch (error: any) {
      emitTelemetry(context, {
        level: 'error',
        event: 'chat-no-evidence-completion-failed',
        intent: intent.kind,
        sourceCount: 0,
        error: error?.message || String(error)
      });
    }

    return {
      answer: noEvidenceAnswer(intent.kind, question),
      citations: []
    };
  }

  const completionStartedAt = Date.now();
  let completion;
  try {
    completion = await getStructuredChatCompletionResult(
      env,
      buildStructuredGroundedChatPrompt(promptQuestion, intent.kind, answerSources),
      { maxTokens: 650 },
      context,
      intent.kind
    );
    emitTelemetry(context, {
      level: 'info',
      event: 'chat-completion-usage',
      intent: intent.kind,
      sourceCount: answerSources.length,
      deployment: completion.deployment,
      completionDurationMs: Date.now() - completionStartedAt,
      promptTokens: completion.usage?.prompt_tokens,
      completionTokens: completion.usage?.completion_tokens,
      totalTokens: completion.usage?.total_tokens
    });
  } catch (error: any) {
    emitTelemetry(context, {
      level: 'error',
      event: 'chat-completion-failed',
      intent: intent.kind,
      sourceCount: answerSources.length,
      completionDurationMs: Date.now() - completionStartedAt,
      error: error?.message || String(error)
    });
    return {
      answer: fallbackGroundedAnswer(question, intent.kind, answerSources),
      citations: sourcesToCitations(fallbackGroundedCitationSources(question, intent.kind, answerSources))
    };
  }

  const structuredAnswer = parseStructuredChatAnswer(completion.content, answerSources.length);
  if (structuredAnswer) {
    let usedSourceNumbers = structuredAnswer.usedSourceNumbers;
    if (structuredAnswer.usedSourceNumbers.length === 0) {
      const limitedNamedDocumentAnswer = buildLimitedNamedDocumentAnswer(question, answerSources);
      if (limitedNamedDocumentAnswer) {
        emitTelemetry(context, {
          level: 'info',
          event: 'chat-limited-named-document-answer',
          intent: intent.kind,
          sourceCount: limitedNamedDocumentAnswer.sourceNumbers.length
        });
        return {
          answer: limitedNamedDocumentAnswer.answer,
          citations: sourcesToCitationsBySourceNumber(answerSources, limitedNamedDocumentAnswer.sourceNumbers),
          usage: {
            promptTokens: completion.usage?.prompt_tokens,
            completionTokens: completion.usage?.completion_tokens,
            totalTokens: completion.usage?.total_tokens,
            deployment: completion.deployment
          }
        };
      }

      if (structuredAnswer.responseType !== 'answer') {
        return {
          answer: noEvidenceAnswer(intent.kind, question),
          citations: [],
          usage: {
            promptTokens: completion.usage?.prompt_tokens,
            completionTokens: completion.usage?.completion_tokens,
            totalTokens: completion.usage?.total_tokens,
            deployment: completion.deployment
          }
        };
      }

      return {
        answer: fallbackGroundedAnswer(question, intent.kind, answerSources),
        citations: sourcesToCitations(fallbackGroundedCitationSources(question, intent.kind, answerSources)),
        usage: {
          promptTokens: completion.usage?.prompt_tokens,
          completionTokens: completion.usage?.completion_tokens,
          totalTokens: completion.usage?.total_tokens,
          deployment: completion.deployment
        }
      };
    }

    return {
      answer: structuredAnswer.answerMarkdown,
      citations: sourcesToCitationsBySourceNumber(answerSources, usedSourceNumbers),
      usage: {
        promptTokens: completion.usage?.prompt_tokens,
        completionTokens: completion.usage?.completion_tokens,
        totalTokens: completion.usage?.total_tokens,
        deployment: completion.deployment
      }
    };
  }

  emitTelemetry(context, {
    level: 'warn',
    event: 'chat-structured-answer-parse-failed',
    intent: intent.kind,
    sourceCount: answerSources.length
  });

  try {
    completion = await getChatCompletionResult(env, buildGroundedChatPrompt(promptQuestion, intent.kind, answerSources), { maxTokens: 650 });
  } catch (error: any) {
    emitTelemetry(context, {
      level: 'error',
      event: 'chat-freeform-completion-failed',
      intent: intent.kind,
      sourceCount: answerSources.length,
      error: error?.message || String(error)
    });
    return {
      answer: fallbackGroundedAnswer(question, intent.kind, answerSources),
      citations: sourcesToCitations(fallbackGroundedCitationSources(question, intent.kind, answerSources))
    };
  }

  const answer = limitRepeatedCitationMarkers(cleanChatAnswer(completion.content));
  const citationSources = citationSourcesUsedByAnswer(answer, answerSources);

  return {
    answer,
    citations: sourcesToCitations(citationSources),
    usage: {
      promptTokens: completion.usage?.prompt_tokens,
      completionTokens: completion.usage?.completion_tokens,
      totalTokens: completion.usage?.total_tokens,
      deployment: completion.deployment
    }
  };
};
