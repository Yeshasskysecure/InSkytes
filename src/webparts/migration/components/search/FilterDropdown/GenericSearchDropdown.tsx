import * as React from "react";
import { createPortal } from "react-dom";
import styles from "./FilterDropdown.module.scss";
import { SPHttpClient } from "@microsoft/sp-http";
import { DocumentDetailPage } from "../../../pages/DocumentDetailPage/DocumentDetailPage";
import {
  formatDocumentPublishedDate,
  resolveDocumentPublishedValue,
  splitDocumentAuthorDisplay,
} from "../../../utils/documentMetadata";
import { subscribeToDocumentDataChanged } from "../../../services/documentChangeEvents";
import { CACHE_KEYS, COLUMN_NAMES, KM_REVIEW_HUB_DRIVE_ID, LIBRARY_NAMES, LIST_NAMES, PAGE_SIZES } from "../../../config/appConfig";
import {
  IKnowledgeSearchDocumentsResponse,
  IKnowledgeSearchFacetValue,
  IKnowledgeSearchFacets,
  IKnowledgeSearchPageInfo,
  IKnowledgeSearchRequest,
  IKnowledgeSearchResult,
  KnowledgeSearchApiClient
} from "../../../services/KnowledgeSearchApiClient";
import { getDriveItemThumbnailUrl, getExcelThumbnailUrl } from "../../../services/SharePointSearchService";
import { downloadSharePointFile } from "../../../utils/fileDownload";

const isSearchDebugLoggingEnabled = (): boolean => {
  try {
    return typeof window !== "undefined" &&
      window.localStorage?.getItem("IKNOWLEDGE_DEBUG_LOGS") === "true";
  } catch {
    return false;
  }
};

const searchDebugLog = (...args: unknown[]): void => {
  if (isSearchDebugLoggingEnabled()) {
    console.log(...args);
  }
};

export interface ISearchResultsChangeMeta {
  requestKey: string;
  query: string;
  sort: "relevance" | "newest" | "oldest";
  skip: number;
  top: number;
  filtersKey: string;
}

export interface IFilterDropdownProps {
  searchText: string;
  draftSearchText?: string;
  isFilterPanelOpen?: boolean;
  activeMetadataFilter?: SearchMetadataFilterKey | null;
  externalFilters?: Partial<Record<keyof FilterState, string | null>>;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  context: any;
  onSuggestionSelect?: (query: string) => void;
  onSearchSubmit?: (query: string) => void;
  onViewDocument?: (documentId: number) => void;
  onSuggestionsChange?: (suggestions: SuggestionItem[]) => void;
  onResultsChange?: (items: ResultItem[], isLoading: boolean, totalCount?: number, page?: IKnowledgeSearchPageInfo, meta?: ISearchResultsChangeMeta) => void;
  onFilterGroupsChange?: (groups: any[]) => void;
  renderSuggestionsSection?: boolean;
  fastResultsOnly?: boolean;
  isLearner?: boolean;
  refreshKey?: number;
  resultsPage?: number;
  resultsPageSize?: number;
  searchSortOrder?: "relevance" | "newest" | "oldest";
}

type FilterGroupTitle =
  | "File Format"
  | "BU"
  | "Department / Sub Department"
  | "Document Type"
  | "Client Category"
  | "Region Category"
  | "Therapy Area"
  | "Disease Area";

const renderAuthorLines = (author: string, containerClassName: string, itemClassName: string): JSX.Element => {
  const authorEntries = splitDocumentAuthorDisplay(author);
  const entriesToRender = authorEntries.length > 0 ? authorEntries : ["Internal"];

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

export type FilterState = {
  fileType: string | null;
  businessUnit: string | null;
  department: string | null;
  documentType: string | null;
  client: string | null;
  region: string | null;
  therapyArea: string | null;
  diseaseArea: string | null;
};

type SearchMetadataFilterKey =
  | "documentType"
  | "businessUnit"
  | "department"
  | "client"
  | "region"
  | "therapyArea"
  | "diseaseArea";

type ResultItem = {
  id: number;
  title: string;
  fileName: string;
  contributor: string;
  updated: string;
  updatedDateValue: string;
  description: string;
  fileUrl?: string;
  fileType?: string;
  serverRelativeUrl?: string;
  fileUniqueId?: string;
  businessUnit?: string;
  department?: string;
  documentType?: string;
  client?: string;
  region?: string;
  therapyArea?: string;
  diseaseArea?: string;
  status?: string;
  viewCount?: number;
  likeCount?: number;
  commentCount?: number;
  shareCount?: number;
  rank?: number;
  matchedByNativeSearch?: boolean;
  matchedByInstantList?: boolean;
  views?: number;
  likes?: number;
  comments?: number;
  downloads?: number;
  share?: boolean;
  docIcon?: string;
};

type SuggestionItem = {
  value: string;
  source: "title" | "query";
  count?: number;
  meta?: string;
};

type DocumentSortOrder = "newToOld" | "oldToNew";
type DocumentViewMode = "list" | "grid";

type PaginationItem = number | 'ellipsis-start' | 'ellipsis-end';

type SearchHistoryItem = {
  id?: number;
  value: string;
  count: number;
};

type SearchAnalyticsContext = {
  searchRequestId?: string;
  searchSessionId?: string;
  startedAt: number;
  ranksByListItemId: Record<number, number>;
};

type FilterOption = {
  title: string;
  value?: string;
  count: number;
  level?: number;
  parentTitle?: string;
};

type FilterGroup = {
  key: keyof FilterState;
  title: FilterGroupTitle;
  children: FilterOption[];
};

type SearchFieldMap = {
  author: string;
  isPersonField?: boolean;
  published: string;
  titleName: string;
  businessUnit: string;
  department: string;
  documentType: string;
  client: string;
  region: string;
  therapyArea: string;
  diseaseArea: string;
  status: string;
  title?: string;
  description?: string;
  views?: string;
  likes?: string;
  comments?: string;
  downloads?: string;
  share?: string;
  docIcon?: string;
  fileLeafRef?: string;
  fileRef?: string;
  allFields?: Array<{ Title?: string; InternalName: string; FieldTypeKind?: number }>;
};

const LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;
const SEARCH_QUERY_LIST = LIST_NAMES.searchQueryHistoryLog;
const LOCAL_HISTORY_KEY = CACHE_KEYS.sharedSearchHistory;
const DOCUMENT_RESULTS_PER_PAGE = PAGE_SIZES.searchResultsPerPage;
const SEARCH_METADATA_CHUNK_SIZE = 30;
const MAX_NATIVE_RANK_SCORE = 1200;

const EMPTY_FILTERS: FilterState = {
  fileType: null,
  businessUnit: null,
  department: null,
  documentType: null,
  client: null,
  region: null,
  therapyArea: null,
  diseaseArea: null,
};

const DOCUMENT_FILE_TYPES = new Set([
  "DOC",
  "DOCX",
  "PDF",
  "PPT",
  "PPTX",
  "XLS",
  "XLSX",
  "TXT",
  "MHT",
  "MHTML",
  "SVG",
]);

const VIDEO_FILE_TYPES = new Set([
  "MP4",
  "MOV",
  "AVI",
  "WMV",
  "M4V",
  "WEBM",
  "MKV",
]);

const AUDIO_FILE_TYPES = new Set([
  "MP3",
  "WAV",
  "M4A",
  "AAC",
  "OGG",
  "FLAC",
  "WMA",
]);

const QUICK_FILE_TYPE_VALUES = {
  documents: "documents",
  video: "video",
  audio: "audio",
} as const;

const FILTER_DEFINITIONS: Array<{
  key: keyof FilterState;
  title: FilterGroupTitle;
  getValue: (item: ResultItem) => string | undefined;
}> = [
  { key: "fileType", title: "File Format", getValue: (item) => item.fileType },
  { key: "businessUnit", title: "BU", getValue: (item) => item.businessUnit },
  { key: "department", title: "Department / Sub Department", getValue: (item) => item.department },
  { key: "documentType", title: "Document Type", getValue: (item) => item.documentType },
  { key: "client", title: "Client Category", getValue: (item) => item.client },
  { key: "region", title: "Region Category", getValue: (item) => item.region },
  { key: "therapyArea", title: "Therapy Area", getValue: (item) => item.therapyArea },
  { key: "diseaseArea", title: "Disease Area", getValue: (item) => item.diseaseArea },
];

const getFilterValues = (filters: FilterState): Array<string | null> =>
  FILTER_DEFINITIONS.map((definition) => filters[definition.key]);

const DEFAULT_SEARCH_FIELD_MAP: SearchFieldMap = {
  author: "Author0",
  isPersonField: true,
  published: "Published",
  titleName: "",
  businessUnit: "",
  department: "",
  documentType: "",
  client: "",
  region: "",
  therapyArea: "",
  diseaseArea: "",
  status: "",
  title: COLUMN_NAMES.title,
  description: COLUMN_NAMES.description,
  views: COLUMN_NAMES.views,
  likes: COLUMN_NAMES.likes,
  comments: COLUMN_NAMES.comments,
  downloads: COLUMN_NAMES.downloads,
  share: COLUMN_NAMES.share,
  docIcon: COLUMN_NAMES.docIcon,
  fileLeafRef: COLUMN_NAMES.fileLeafRef,
  fileRef: COLUMN_NAMES.fileRef,
};

const normalizeFieldName = (value?: string): string => (value || "").trim().toLowerCase();

const cleanTaxonomyTextValue = (value: string): string => {
  const normalizedValue = String(value || "").trim();
  if (!normalizedValue) {
    return "";
  }

  return normalizedValue
    .split(/;#|;/)
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) => {
      const parts = entry.split("|").map((part) => part.trim()).filter(Boolean);
      const label = parts[0] || entry;
      return /^\d+$/.test(label) ? "" : label;
    })
    .filter(Boolean)
    .filter((entry, index, entries) => entries.indexOf(entry) === index)
    .join(", ");
};

const extractStringValue = (val: any): string => {
  if (val === null || val === undefined) return "";
  if (typeof val === "string") return cleanTaxonomyTextValue(val) || val;
  if (typeof val === "number" || typeof val === "boolean") return String(val);

  if (Array.isArray(val)) {
    return val.map(extractStringValue).filter(Boolean).join(", ");
  }

  if (typeof val === "object") {
    // SharePoint Taxonomy / Choice arrays
    if (val.results && Array.isArray(val.results)) {
      return val.results.map(extractStringValue).filter(Boolean).join(", ");
    }
    // Managed Metadata Label
    if (val.Label) return String(val.Label);
    if (val.TermGuid) return String(val.Label || "");
    // Lookup / User Object
    if (val.Title) return String(val.Title);
    if (val.Name) return String(val.Name);

    return ""; // Avoid returning "[object Object]"
  }

  return String(val);
};

const normalizeText = (value?: string | null): string =>
  (value || "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();

const normalizeSearchKey = (value?: string | null): string =>
  normalizeText(value)
    .replace(/[_-]+/g, " ")
    .replace(/[^\w\s.]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const stripFileExtension = (value?: string | null): string => {
  const normalizedValue = String(value || "").trim();
  return normalizedValue.replace(/\.[^.]+$/, "");
};

const getExactMatchPriority = (item: ResultItem, query: string): number => {
  const normalizedQuery = normalizeSearchKey(query);
  const normalizedQueryWithoutExtension = normalizeSearchKey(stripFileExtension(query));

  if (!normalizedQuery) {
    return 5;
  }

  const title = normalizeSearchKey(item.title);
  const fileName = normalizeSearchKey(item.fileName);
  const titleWithoutExtension = normalizeSearchKey(stripFileExtension(item.title));
  const fileNameWithoutExtension = normalizeSearchKey(stripFileExtension(item.fileName));

  if (title === normalizedQuery || titleWithoutExtension === normalizedQuery) {
    return 0;
  }

  if (fileName === normalizedQuery || fileNameWithoutExtension === normalizedQuery) {
    return 1;
  }

  if (
    normalizedQueryWithoutExtension &&
    (title === normalizedQueryWithoutExtension ||
      titleWithoutExtension === normalizedQueryWithoutExtension)
  ) {
    return 2;
  }

  if (
    normalizedQueryWithoutExtension &&
    (fileName === normalizedQueryWithoutExtension ||
      fileNameWithoutExtension === normalizedQueryWithoutExtension)
  ) {
    return 3;
  }

  return 5;
};

const tokenizeSearchText = (value: string): string[] =>
  Array.from(
    new Set(
      normalizeText(value)
        .split(/[^a-z0-9]+/i)
        .filter((token) => token.length > 0)
    )
  );

const SEARCH_STOP_WORDS = new Set([
  "a",
  "an",
  "and",
  "are",
  "as",
  "at",
  "be",
  "by",
  "for",
  "from",
  "how",
  "in",
  "is",
  "it",
  "of",
  "on",
  "or",
  "that",
  "the",
  "this",
  "to",
  "use",
  "was",
  "what",
  "when",
  "where",
  "who",
  "with",
]);

const normalizeSearchWhitespace = (value: string): string =>
  value.replace(/\s+/g, " ").trim();

const isEmailSearch = (value: string): boolean =>
  /^[^\s@]+@[^\s@]+\.[^\s@]+$/i.test(normalizeSearchWhitespace(value));

const isSentenceLikeSearch = (value: string): boolean => {
  const normalizedValue = normalizeSearchWhitespace(value);
  if (
    !normalizedValue ||
    normalizedValue === "." ||
    normalizedValue === "*" ||
    hasExplicitSearchSyntax(normalizedValue) ||
    isEmailSearch(normalizedValue)
  ) {
    return false;
  }

  const tokenCount = normalizedValue.split(/\s+/).filter(Boolean).length;
  const meaningfulTermCount = buildMeaningfulSearchTerms(normalizedValue).length;
  return meaningfulTermCount >= 4 && (tokenCount >= 7 || normalizedValue.length >= 40);
};

const isSimpleConjunctionSearch = (value: string): boolean => {
  const normalizedValue = normalizeSearchWhitespace(value);
  if (
    !normalizedValue ||
    !hasExplicitSearchSyntax(normalizedValue) ||
    hasQuotedSearchSyntax(normalizedValue)
  ) {
    return false;
  }

  let hasAndOperator = false;
  let termCount = 0;

  for (const token of parseSearchSyntaxTokens(normalizedValue)) {
    if (token.type === "operator") {
      if (token.value === "OR" || token.value === "NOT") {
        return false;
      }

      if (token.value === "AND") {
        hasAndOperator = true;
      }

      continue;
    }

    termCount += 1;
  }

  return hasAndOperator && termCount >= 2 && termCount <= 3;
};

const normalizeSearchToken = (value: string): string =>
  normalizeText(value).replace(/[^a-z0-9]/gi, "");

const getPrimarySearchTokenVariant = (value: string): string => {
  const normalizedToken = normalizeSearchToken(value);
  if (!normalizedToken || !/^[a-z]+$/i.test(normalizedToken) || normalizedToken.length <= 3) {
    return normalizedToken;
  }

  if (normalizedToken.endsWith("ies") && normalizedToken.length > 4) {
    return `${normalizedToken.slice(0, -3)}y`;
  }

  if (/(ics|sis|ss|us)$/.test(normalizedToken)) {
    return normalizedToken;
  }

  if (/(ches|shes|xes|zes)$/.test(normalizedToken)) {
    return normalizedToken.slice(0, -2);
  }

  if (normalizedToken.endsWith("s")) {
    return normalizedToken.slice(0, -1);
  }

  return normalizedToken;
};

const buildSearchTokenVariants = (value: string): string[] => {
  const normalizedToken = normalizeSearchToken(value);
  if (!normalizedToken) {
    return [];
  }

  const variants = new Set<string>([normalizedToken]);
  const primaryVariant = getPrimarySearchTokenVariant(normalizedToken);

  if (primaryVariant) {
    variants.add(primaryVariant);
  }

  if (
    primaryVariant &&
    /^[a-z]+$/i.test(primaryVariant) &&
    primaryVariant.length > 3 &&
    primaryVariant.endsWith("y") &&
    !/[aeiou]y$/i.test(primaryVariant)
  ) {
    variants.add(`${primaryVariant.slice(0, -1)}ies`);
  } else if (
    primaryVariant &&
    /^[a-z]+$/i.test(primaryVariant) &&
    primaryVariant.length > 3 &&
    !/(ics|sis|ss|us|s)$/.test(primaryVariant)
  ) {
    variants.add(`${primaryVariant}s`);
  }

  return Array.from(variants);
};

const buildPrimarySearchPhraseVariant = (value: string): string => {
  const normalizedValue = normalizeSearchKey(value);
  if (!normalizedValue) {
    return "";
  }

  return normalizedValue
    .split(/\s+/)
    .map((token) => getPrimarySearchTokenVariant(token) || token)
    .join(" ")
    .trim();
};

type SearchSyntaxToken =
  | { type: "operator"; value: "AND" | "OR" | "NOT" }
  | { type: "term"; value: string; quoted: boolean };

const parseSearchSyntaxTokens = (value: string): SearchSyntaxToken[] => {
  const tokens: SearchSyntaxToken[] = [];
  const tokenPattern = /"([^"]+)"|(\S+)/g;
  let match: RegExpExecArray | null;

  while ((match = tokenPattern.exec(value)) !== null) {
    const quotedValue = normalizeSearchWhitespace(match[1] || "");
    const rawValue = normalizeSearchWhitespace(match[2] || "");

    if (quotedValue) {
      tokens.push({ type: "term", value: quotedValue, quoted: true });
      continue;
    }

    if (!rawValue) {
      continue;
    }

    const normalizedOperator = rawValue.toUpperCase();
    if (normalizedOperator === "AND" || normalizedOperator === "OR" || normalizedOperator === "NOT") {
      tokens.push({ type: "operator", value: normalizedOperator });
      continue;
    }

    tokens.push({ type: "term", value: rawValue, quoted: false });
  }

  return tokens;
};

const hasQuotedSearchSyntax = (value: string): boolean => /"[^"]+"/.test(value);

const hasExplicitUppercaseSearchOperators = (value: string): boolean =>
  /\b(?:OR|NOT|AND)\b/.test(value);

const hasExplicitLowercaseSearchOperators = (value: string): boolean =>
  /\b(?:or|not|and)\b/.test(value);

const shouldTreatLowercaseBooleanOperatorsAsExplicit = (value: string): boolean => {
  if (!hasExplicitLowercaseSearchOperators(value)) {
    return false;
  }

  const normalizedValue = normalizeSearchWhitespace(value);
  const tokenCount = normalizedValue.split(/\s+/).filter(Boolean).length;
  return tokenCount > 0 && tokenCount <= 6;
};

const hasExplicitSearchOperators = (value: string): boolean =>
  hasExplicitUppercaseSearchOperators(value) || shouldTreatLowercaseBooleanOperatorsAsExplicit(value);

const hasExplicitSearchSyntax = (value: string): boolean =>
  hasQuotedSearchSyntax(value) || hasExplicitSearchOperators(value);

const buildMeaningfulSearchTerms = (value: string): string[] =>
  Array.from(
    new Set(
      normalizeText(value)
        .split(/[^a-z0-9]+/i)
        .filter((token) => token.length > 2 && !SEARCH_STOP_WORDS.has(token))
    )
  );

const buildConversationalSearchCore = (value: string): string => {
  const normalizedValue = normalizeSearchWhitespace(value).replace(/[?!.,]+$/g, "");
  if (!normalizedValue) {
    return normalizedValue;
  }

  let strippedValue = normalizedValue;

  strippedValue = strippedValue.replace(
    /^(?:please\s+)?(?:do\s+you\s+have|can\s+you\s+(?:find|show|search)(?:\s+me)?|could\s+you\s+(?:find|show|search)(?:\s+me)?|find\s+me|search\s+for|show\s+me|give\s+me|i\s+need|i'?m\s+looking\s+for|looking\s+for|are\s+there|is\s+there)\s+/i,
    ""
  );

  strippedValue = strippedValue.replace(
    /^(?:any\s+|some\s+|all\s+)?(?:documents?|docs?|files?|presentations?|slides?|pptx?|ppts?|powerpoints?|pdfs?|content|materials?)\s*(?:that\s+are\s+|which\s+are\s+)?(?:related\s+to|regarding|about|on|for)\s+/i,
    ""
  );

  strippedValue = strippedValue.replace(
    /^(?:any\s+|some\s+|all\s+)?(?:documents?|docs?|files?|presentations?|slides?|pptx?|ppts?|powerpoints?|pdfs?|content|materials?)\b\s*/i,
    ""
  );

  strippedValue = strippedValue.replace(/^(?:related\s+to|regarding|about|on|for)\s+/i, "");
  strippedValue = strippedValue.replace(
    /^(?:something|anything)\s+(?:related\s+to|regarding|about|on|for)\s+/i,
    ""
  );
  strippedValue = strippedValue.replace(/\b(?:documents?|docs?|files?|presentations?|slides?|pptx?|ppts?|powerpoints?|pdfs?|content|materials?)$/i, "");
  strippedValue = normalizeSearchWhitespace(strippedValue);

  if (!strippedValue) {
    return normalizedValue;
  }

  const strippedTerms = buildMeaningfulSearchTerms(strippedValue);
  if (strippedTerms.length === 0) {
    return normalizedValue;
  }

  return strippedValue;
};

const buildListSearchLiterals = (value: string, fallbackValues: string[] = []): string[] => {
  const literals: string[] = [];
  const addLiteral = (literal?: string): void => {
    const normalizedLiteral = normalizeSearchWhitespace(literal || "");
    if (!normalizedLiteral || literals.indexOf(normalizedLiteral) >= 0) {
      return;
    }

    literals.push(normalizedLiteral);
  };

  const addLiteralWithVariants = (literal?: string): void => {
    const normalizedLiteral = normalizeSearchWhitespace(literal || "");
    if (!normalizedLiteral) {
      return;
    }

    addLiteral(normalizedLiteral);

    const normalizedKey = normalizeSearchKey(normalizedLiteral);
    if (normalizedKey && normalizedKey !== normalizedLiteral) {
      addLiteral(normalizedKey);
    }

    const primaryPhraseVariant = buildPrimarySearchPhraseVariant(normalizedLiteral);
    if (
      primaryPhraseVariant &&
      primaryPhraseVariant !== normalizedLiteral &&
      primaryPhraseVariant !== normalizedKey
    ) {
      addLiteral(primaryPhraseVariant);
    }

    if (normalizedKey.includes(" ")) {
      addLiteral(normalizedKey.replace(/\s+/g, "_"));
      addLiteral(normalizedKey.replace(/\s+/g, "-"));
    }
  };

  const addLiteralsFromValue = (input: string): void => {
    const normalizedValue = normalizeSearchWhitespace(input);
    if (!normalizedValue) {
      return;
    }

    if (!hasExplicitSearchSyntax(normalizedValue)) {
      addLiteralWithVariants(normalizedValue);
      buildMeaningfulSearchTerms(normalizedValue)
        .slice(0, 4)
        .forEach((keyword) => buildSearchTokenVariants(keyword).forEach((variant) => addLiteral(variant)));
      return;
    }

    const positiveLiterals: string[] = [];
    let skipNextTerm = false;

    parseSearchSyntaxTokens(normalizedValue).forEach((token) => {
      if (token.type === "operator") {
        skipNextTerm = token.value === "NOT";
        return;
      }

      if (!skipNextTerm) {
        positiveLiterals.push(token.value);
      }
      skipNextTerm = false;
    });

    positiveLiterals.forEach((literal) => addLiteralWithVariants(literal));
    buildMeaningfulSearchTerms(normalizedValue)
      .slice(0, 4)
      .forEach((keyword) => buildSearchTokenVariants(keyword).forEach((variant) => addLiteral(variant)));
  };

  addLiteralsFromValue(value);
  fallbackValues.forEach((fallbackValue) => addLiteralsFromValue(fallbackValue));
  return literals.slice(0, 8);
};

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

const escapeODataValue = (value: string): string => value.replace(/'/g, "''");

const chunkNumberArray = (values: number[], chunkSize: number): number[][] => {
  const chunks: number[][] = [];
  for (let index = 0; index < values.length; index += chunkSize) {
    chunks.push(values.slice(index, index + chunkSize));
  }
  return chunks;
};

const formatRelativeDate = (value: string): string => {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }

  return date.toLocaleDateString(undefined, {
    year: "numeric",
    month: "short",
    day: "numeric",
  });
};

const getFileTypeFromName = (fileName: string | undefined): string => {
  if (!fileName) return "FILE";
  const parts = fileName.split(".");
  if (parts.length < 2) return "FILE";
  return (parts.pop() || "FILE").toUpperCase();
};

const buildPaginationItems = (currentPage: number, totalPages: number): PaginationItem[] => {
  if (totalPages <= 7) {
    return Array.from({ length: totalPages }, (_, index) => index + 1);
  }

  if (currentPage <= 4) {
    return [1, 2, 3, 4, 5, 'ellipsis-end', totalPages];
  }

  if (currentPage >= totalPages - 3) {
    return [1, 'ellipsis-start', totalPages - 4, totalPages - 3, totalPages - 2, totalPages - 1, totalPages];
  }

  return [1, 'ellipsis-start', currentPage - 1, currentPage, currentPage + 1, 'ellipsis-end', totalPages];
};

const getDisplayStatus = (status?: string): string => {
  const trimmedStatus = String(status || "").trim();
  if (!trimmedStatus) {
    return "Unknown";
  }
  return trimmedStatus.toLowerCase() === "rejected" ? "Reject" : trimmedStatus;
};

const getStatusClassName = (status?: string): string => {
  const normalizedStatus = normalizeText(status)
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  if (/\breject(ed)?\b/.test(normalizedStatus)) return styles.statusBadgeRejected;
  if (/\barchive(d)?\b/.test(normalizedStatus)) return styles.statusBadgeArchive;
  if (/\bunder\s*review\b/.test(normalizedStatus) || normalizedStatus === "review") return styles.statusBadgeReview;
  if (/^(active|approved)$/.test(normalizedStatus)) return styles.statusBadgeActive;

  return styles.statusBadgeDefault;
};

const getFileTypeIcon = (type: string | undefined): string => {
  const normalized = (type || "").toLowerCase();
  if (normalized.includes("pdf")) return "📄";
  if (normalized.includes("ppt")) return "📊";
  if (normalized.includes("xls")) return "📈";
  if (normalized.includes("doc") || normalized.includes("txt")) return "📝";
  if (/\b(mp4|mov|webm|mkv|avi|m4v|wmv)\b/.test(normalized) || normalized.includes("video")) return "▶";
  if (/\b(mp3|wav|m4a|aac|flac|ogg)\b/.test(normalized) || normalized.includes("audio")) return "♪";
  return "📎";
};

const decodeUrlSafely = (value: string): string => {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
};

const isExcelFileType = (fileType: string | undefined): boolean =>
  ["xls", "xlsx", "xlsm", "xlsb"].indexOf(String(fileType || "").trim().toLowerCase()) !== -1;

const isAudioFileType = (fileType: string | undefined): boolean =>
  ["mp3", "wav", "m4a", "aac", "flac", "ogg", "wma"].indexOf(String(fileType || "").trim().toLowerCase()) !== -1;

const isVideoFileType = (fileType: string | undefined): boolean => {
  const normalized = String(fileType || "").trim().toLowerCase();
  return ["mp4", "mov", "webm", "mkv", "avi", "m4v", "wmv"].indexOf(normalized) !== -1 || normalized.indexOf("video") !== -1;
};

const getAbsoluteSearchFileUrl = (
  serverRelativeUrl: string | undefined,
  webUrl: string,
  fileUrl?: string
): string => {
  if (!serverRelativeUrl || !webUrl) {
    return fileUrl || "";
  }

  const decodedServerRelativeUrl = decodeUrlSafely(serverRelativeUrl);
  const normalizedServerRelativeUrl = decodedServerRelativeUrl.startsWith("/") ? decodedServerRelativeUrl : `/${decodedServerRelativeUrl}`;
  const origin = new URL(webUrl).origin;

  if (decodedServerRelativeUrl.startsWith("http")) {
    return decodedServerRelativeUrl;
  }

  if (fileUrl && fileUrl.startsWith("http")) {
    return decodeUrlSafely(fileUrl);
  }

  return `${origin}${normalizedServerRelativeUrl}`;
};

const getSearchThumbnailCandidates = (
  serverRelativeUrl: string | undefined,
  webUrl: string,
  fileType?: string,
  fileUrl?: string,
  fileUniqueId?: string
): string[] => {
  if (!serverRelativeUrl || !webUrl) {
    return [];
  }

  const absoluteFileUrl = getAbsoluteSearchFileUrl(serverRelativeUrl, webUrl, fileUrl);
  const normalizedFileType = String(fileType || "").trim().toLowerCase();

  if (isAudioFileType(normalizedFileType)) {
    return [];
  }

  const isExcelFile = isExcelFileType(normalizedFileType);
  const candidates =
    ["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg"].indexOf(normalizedFileType) !== -1
      ? [absoluteFileUrl]
      : [];

  if (fileUniqueId) {
    const driveThumbnailUrl = isExcelFile
      ? getExcelThumbnailUrl(webUrl, KM_REVIEW_HUB_DRIVE_ID, fileUniqueId)
      : getDriveItemThumbnailUrl(webUrl, KM_REVIEW_HUB_DRIVE_ID, fileUniqueId);
    candidates.push(driveThumbnailUrl);
  }

  return candidates.slice(0, 3);
};

const buildResultPreview = (item: ResultItem, _query: string, _cachedContent?: string): string => {
  return item.description || "";
};

const countTermMatches = (haystack: string, term: string): number => {
  if (!haystack || !term) {
    return 0;
  }

  const matches = haystack.match(new RegExp(`\\b${escapeRegExp(term)}\\b`, "gi"));
  return matches ? matches.length : 0;
};

const countSearchTermMatches = (haystack: string, term: string): number =>
  buildSearchTokenVariants(term).reduce(
    (highestMatchCount, variant) => Math.max(highestMatchCount, countTermMatches(haystack, variant)),
    0
  );

const countLiteralMatches = (haystack: string, literal: string, quoted = false): number => {
  const normalizedLiteral = normalizeSearchKey(literal);
  if (!haystack || !normalizedLiteral) {
    return 0;
  }

  if (quoted || /[\s./_-]/.test(normalizedLiteral)) {
    let count = 0;
    let searchIndex = 0;
    while (searchIndex < haystack.length) {
      const matchIndex = haystack.indexOf(normalizedLiteral, searchIndex);
      if (matchIndex < 0) {
        break;
      }
      count += 1;
      searchIndex = matchIndex + normalizedLiteral.length;
    }
    return count;
  }

  return countSearchTermMatches(haystack, normalizedLiteral);
};

const buildPositiveSearchLiterals = (query: string): Array<{ value: string; quoted: boolean }> => {
  const normalizedQuery = normalizeSearchWhitespace(query);
  if (!normalizedQuery) {
    return [];
  }

  if (!hasExplicitSearchSyntax(normalizedQuery)) {
    const meaningfulTerms = buildMeaningfulSearchTerms(normalizedQuery);
    if (meaningfulTerms.length > 0) {
      return meaningfulTerms.map((value) => ({ value, quoted: false }));
    }

    return [{ value: normalizedQuery, quoted: normalizedQuery.indexOf(" ") !== -1 }];
  }

  const literals: Array<{ value: string; quoted: boolean }> = [];
  let skipNextTerm = false;
  parseSearchSyntaxTokens(normalizedQuery).forEach((token) => {
    if (token.type === "operator") {
      skipNextTerm = token.value === "NOT";
      return;
    }

    if (!skipNextTerm) {
      literals.push({ value: token.value, quoted: token.quoted });
    }
    skipNextTerm = false;
  });

  return literals;
};

const countKeywordMatches = (item: ResultItem, query: string, cachedContent?: string): number => {
  const literals = buildPositiveSearchLiterals(query);
  if (literals.length === 0) {
    return 0;
  }

  const searchableText = buildResultSearchableText(item, cachedContent);
  const literalMatches = literals.reduce(
    (total, literal) => total + countLiteralMatches(searchableText, literal.value, literal.quoted),
    0
  );
  const normalizedQuery = normalizeSearchKey(query);
  const fullSentenceBoost = isSentenceLikeSearch(query) && normalizedQuery && searchableText.includes(normalizedQuery)
    ? Math.max(25, literals.length * 6)
    : 0;

  return literalMatches + fullSentenceBoost;
};

const isAndSeparatedSearch = (value: string): boolean => {
  const normalizedValue = normalizeSearchWhitespace(value);
  if (!normalizedValue || hasQuotedSearchSyntax(normalizedValue) || /\b(?:OR|NOT|or|not)\b/.test(normalizedValue)) {
    return false;
  }

  return /\band\b/i.test(normalizedValue) && buildMeaningfulSearchTerms(normalizedValue).length >= 2;
};

const getMinimumRequiredTermCoverage = (query: string, termCount: number): number => {
  if (termCount <= 1) {
    return termCount;
  }

  if (isAndSeparatedSearch(query)) {
    return termCount;
  }

  if (isSentenceLikeSearch(query)) {
    return Math.max(3, Math.ceil(termCount * 0.55));
  }

  if (termCount <= 4) {
    return Math.max(2, termCount - 1);
  }

  return Math.max(3, Math.ceil(termCount * 0.5));
};

const buildSearchRelevanceProfile = (item: ResultItem, query: string, cachedContent?: string): {
  distinctTermMatches: number;
  requiredTermMatches: number;
  phraseMatches: number;
  titleTermMatches: number;
  metadataTermMatches: number;
  contentTermMatches: number;
  minimumRequiredMatches: number;
  coverageScore: number;
} => {
  const normalizedQuery = normalizeSearchWhitespace(query);
  const terms = buildMeaningfulSearchTerms(normalizedQuery);
  if (terms.length === 0) {
    return {
      distinctTermMatches: 0,
      requiredTermMatches: 0,
      phraseMatches: 0,
      titleTermMatches: 0,
      metadataTermMatches: 0,
      contentTermMatches: 0,
      minimumRequiredMatches: 0,
      coverageScore: 0,
    };
  }

  const titleText = normalizeSearchKey([item.title, item.fileName].filter(Boolean).join(" "));
  const metadataText = normalizeSearchKey([
    item.businessUnit,
    item.department,
    item.documentType,
    item.client,
    item.region,
    item.therapyArea,
    item.diseaseArea,
    item.contributor,
  ].filter(Boolean).join(" "));
  const descriptionText = normalizeSearchKey(item.description);
  const contentText = normalizeSearchKey(cachedContent);
  const allText = `${titleText} ${metadataText} ${descriptionText} ${contentText}`;

  let titleTermMatches = 0;
  let metadataTermMatches = 0;
  let contentTermMatches = 0;
  let distinctTermMatches = 0;

  terms.forEach((term) => {
    const titleMatched = matchesSearchLiteral(titleText, term);
    const metadataMatched = matchesSearchLiteral(metadataText, term);
    const descriptionMatched = matchesSearchLiteral(descriptionText, term);
    const contentMatched = matchesSearchLiteral(contentText, term);

    if (titleMatched) titleTermMatches += 1;
    if (metadataMatched) metadataTermMatches += 1;
    if (contentMatched) contentTermMatches += 1;
    if (titleMatched || metadataMatched || descriptionMatched || contentMatched) {
      distinctTermMatches += 1;
    }
  });

  const phraseTerms = terms.map((term) => getPrimarySearchTokenVariant(term) || term);
  const phraseMatches = phraseTerms
    .slice(0, Math.max(0, phraseTerms.length - 1))
    .reduce((total, term, index) => {
      const nextTerm = phraseTerms[index + 1];
      const phrase = nextTerm ? `${term} ${nextTerm}` : "";
      return phrase && allText.includes(phrase) ? total + 1 : total;
    }, 0);
  const minimumRequiredMatches = getMinimumRequiredTermCoverage(normalizedQuery, terms.length);
  const requiredTermMatches = Math.min(distinctTermMatches, minimumRequiredMatches);
  const fullQueryMatch = normalizeSearchKey(normalizedQuery) && allText.includes(normalizeSearchKey(normalizedQuery));
  const coverageScore =
    requiredTermMatches * 10000 +
    distinctTermMatches * 2500 +
    phraseMatches * 5000 +
    titleTermMatches * 900 +
    metadataTermMatches * 1100 +
    contentTermMatches * 700 +
    (fullQueryMatch ? 250000 : 0);

  return {
    distinctTermMatches,
    requiredTermMatches,
    phraseMatches,
    titleTermMatches,
    metadataTermMatches,
    contentTermMatches,
    minimumRequiredMatches,
    coverageScore,
  };
};

const scoreResult = (item: ResultItem, query: string, cachedContent?: string): number => {
  const normalizedQuery = normalizeSearchKey(query);
  if (!normalizedQuery) {
    return 0;
  }

  const meaningfulTokens = buildMeaningfulSearchTerms(normalizedQuery);
  const queryTokens = meaningfulTokens.length > 0 ? meaningfulTokens : tokenizeSearchText(normalizedQuery);
  const queryPhrases = Array.from(
    new Set(
      queryTokens
        .slice(0, Math.max(0, queryTokens.length - 1))
        .map((token, index) => (queryTokens[index + 1] ? `${token} ${queryTokens[index + 1]}` : ""))
        .filter(Boolean)
     )
   ).slice(0, 6);
  const title = normalizeSearchKey(item.title);
  const fileName = normalizeSearchKey(item.fileName);
  const contributor = normalizeSearchKey(item.contributor);
  const description = normalizeSearchKey(item.description);
  const businessUnit = normalizeSearchKey(item.businessUnit);
  const documentType = normalizeSearchKey(item.documentType);
  const client = normalizeSearchKey(item.client);
  const region = normalizeSearchKey(item.region);
  const therapyArea = normalizeSearchKey(item.therapyArea);
  const diseaseArea = normalizeSearchKey(item.diseaseArea);
  const content = normalizeSearchKey(cachedContent);
  const isExactIdentifierQuery = isEmailSearch(query);

  let score = 0;

  if (title === normalizedQuery) score += 200;
  if (title.startsWith(normalizedQuery)) score += 140;
  if (title.includes(normalizedQuery)) score += 90;
  if (fileName.startsWith(normalizedQuery)) score += 90;
  if (fileName.includes(normalizedQuery)) score += 60;
  if (contributor === normalizedQuery) score += 180;
  if (contributor.includes(normalizedQuery)) score += 90;
  if (documentType === normalizedQuery) score += 80;
  else if (documentType.includes(normalizedQuery)) score += 30;

  if (businessUnit === normalizedQuery) score += 70;
  else if (businessUnit.includes(normalizedQuery)) score += 30;

  if (client === normalizedQuery) score += 70;
  else if (client.includes(normalizedQuery)) score += 30;

  if (region === normalizedQuery) score += 55;
  else if (region.includes(normalizedQuery)) score += 30;
  
  if (therapyArea === normalizedQuery || diseaseArea === normalizedQuery) score += 55;
  else if (therapyArea.includes(normalizedQuery) || diseaseArea.includes(normalizedQuery)) score += 30;

  if (isSentenceLikeSearch(query)) {
    const fullText = buildResultSearchableText(item, cachedContent);
    if (fullText.includes(normalizedQuery)) {
      score += 500;
    }
  }

  if (description.includes(normalizedQuery)) score += 90;
  if (content.includes(normalizedQuery)) score += 380;

  if (isExactIdentifierQuery) {
    if (title.includes(normalizedQuery) || fileName.includes(normalizedQuery)) score += 280;
    if (description.includes(normalizedQuery)) score += 220;
    if (content.includes(normalizedQuery)) score += 720;
  }

  queryPhrases.forEach((phrase) => {
    if (description.includes(phrase)) {
      score += 35;
    }
    if (content.includes(phrase)) {
      score += 120;
    }
  });

  let matchedTokenCount = 0;
  let contentTokenCount = 0;
  let descriptionTokenCount = 0;

  queryTokens.forEach((token) => {
    let tokenMatched = false;
    const descriptionMatches = countSearchTermMatches(description, token);
    const contentMatches = countSearchTermMatches(content, token);

    const weightedFields: Array<[string, number]> = [
      [title, 28],
      [fileName, 22],
      [contributor, 18],
      [documentType, 16],
      [businessUnit, 16],
      [client, 15],
      [region, 12],
      [therapyArea, 12],
      [diseaseArea, 12],
      [description, 15],
      [content, 6],
    ];

    weightedFields.forEach(([field, weight]) => {
      const occurrences = countSearchTermMatches(field, token);
      if (occurrences > 0) {
        score += Math.min(occurrences, 3) * weight;
        tokenMatched = true;
      }
    });

    if (tokenMatched) {
      matchedTokenCount += 1;
    }

    if (descriptionMatches > 0) {
      descriptionTokenCount += 1;
    }

    if (contentMatches > 0) {
      contentTokenCount += 1;
    }
  });

  if (matchedTokenCount === queryTokens.length && queryTokens.length > 1) {
    score += 55;
  } else {
    score += matchedTokenCount * 10;
  }

  if (descriptionTokenCount === queryTokens.length && queryTokens.length >= 3) {
    score += 120;
  }

  if (contentTokenCount === queryTokens.length && queryTokens.length >= 3) {
    score += 320;
  } else if (contentTokenCount >= 3) {
    score += contentTokenCount * 25;
  }

  return score;
};

type ExplicitSearchCondition = {
  value: string;
  quoted: boolean;
  negate: boolean;
};

const buildResultSearchableText = (item: ResultItem, cachedContent?: string): string =>
  normalizeSearchKey(
    [
      item.title,
      item.fileName,
      item.contributor,
      item.description,
      item.businessUnit,
      item.department,
      item.documentType,
      item.client,
      item.region,
      item.therapyArea,
      item.diseaseArea,
      cachedContent || "",
    ]
      .filter(Boolean)
      .join(" ||| ")
  );

const matchesSearchLiteral = (searchableText: string, literal: string, quoted = false): boolean => {
  const normalizedLiteral = normalizeSearchKey(literal);
  if (!normalizedLiteral) {
    return false;
  }

  if (quoted) {
    if (/^[a-z0-9]+$/i.test(normalizedLiteral)) {
      return countSearchTermMatches(searchableText, normalizedLiteral) > 0;
    }

    return searchableText.includes(normalizedLiteral);
  }

  if (/[\s./_-]/.test(normalizedLiteral)) {
    return searchableText.includes(normalizedLiteral);
  }

  return countSearchTermMatches(searchableText, normalizedLiteral) > 0;
};

const getSingleQuotedSearchLiteral = (value: string): string => {
  const tokens = parseSearchSyntaxTokens(normalizeSearchWhitespace(value));
  if (tokens.length !== 1) {
    return "";
  }

  const [token] = tokens;
  return token.type === "term" && token.quoted ? token.value : "";
};

const getStrictQuotedSingleTermLiteral = (value: string): string => {
  const quotedLiteral = getSingleQuotedSearchLiteral(value);
  const normalizedToken = normalizeSearchToken(quotedLiteral);
  return normalizedToken && normalizedToken.length <= 3 ? quotedLiteral : "";
};

const buildStrictQuotedSingleTermFieldValues = (item: ResultItem): string[] =>
  Array.from(
    new Set(
      [
        item.title,
        stripFileExtension(item.title),
        item.fileName,
        stripFileExtension(item.fileName),
        item.contributor,
        item.businessUnit,
        item.department,
        item.documentType,
        item.client,
        item.region,
        item.therapyArea,
        item.diseaseArea,
      ]
        .map((value) => normalizeSearchKey(value))
        .filter(Boolean)
    )
  );

const matchesStrictQuotedSingleTerm = (item: ResultItem, literal: string): boolean => {
  const normalizedLiteral = normalizeSearchKey(literal);
  if (!normalizedLiteral) {
    return false;
  }

  return buildStrictQuotedSingleTermFieldValues(item).some(
    (fieldValue) => fieldValue === normalizedLiteral
  );
};

const buildExplicitSearchConditionGroups = (value: string): ExplicitSearchCondition[][] => {
  const groups: ExplicitSearchCondition[][] = [];
  let currentGroup: ExplicitSearchCondition[] = [];
  let negateNext = false;

  parseSearchSyntaxTokens(normalizeSearchWhitespace(value)).forEach((token) => {
    if (token.type === "operator") {
      if (token.value === "OR") {
        if (currentGroup.length > 0) {
          groups.push(currentGroup);
          currentGroup = [];
        }
        negateNext = false;
        return;
      }

      if (token.value === "NOT") {
        negateNext = true;
      }
      return;
    }

    currentGroup.push({
      value: token.value,
      quoted: token.quoted,
      negate: negateNext,
    });
    negateNext = false;
  });

  if (currentGroup.length > 0) {
    groups.push(currentGroup);
  }

  return groups;
};

const matchesExplicitSearchQuery = (searchableText: string, value: string): boolean => {
  const groups = buildExplicitSearchConditionGroups(value);
  if (groups.length === 0) {
    return true;
  }

  return groups.some((group) =>
    group.every((condition) => {
      const isMatched = matchesSearchLiteral(searchableText, condition.value, condition.quoted);
      return condition.negate ? !isMatched : isMatched;
    })
  );
};

const shouldPreserveNativeSearchFallback = (value: string): boolean => {
  const normalizedValue = normalizeSearchWhitespace(value);
  if (!normalizedValue || normalizedValue === "." || normalizedValue === "*") {
    return false;
  }

  if (hasQuotedSearchSyntax(normalizedValue)) {
    return false;
  }

  if (!hasExplicitSearchSyntax(normalizedValue)) {
    const meaningfulTerms = buildMeaningfulSearchTerms(normalizedValue);
    if (meaningfulTerms.length === 1) {
      return false;
    }

    return meaningfulTerms.length <= 2;
  }

  let hasOrOperator = false;
  let hasAndOperator = false;
  let hasNotOperator = false;

  parseSearchSyntaxTokens(normalizedValue).forEach((token) => {
    if (token.type !== "operator") {
      return;
    }

    if (token.value === "OR") {
      hasOrOperator = true;
      return;
    }

    if (token.value === "AND") {
      hasAndOperator = true;
      return;
    }

    if (token.value === "NOT") {
      hasNotOperator = true;
    }
  });

  return (hasOrOperator && !hasAndOperator && !hasNotOperator) || isSimpleConjunctionSearch(normalizedValue);
};

const doesResultMatchSearchIntent = (item: ResultItem, query: string, cachedContent?: string): boolean => {
  const normalizedQuery = normalizeSearchWhitespace(query);
  if (!normalizedQuery || normalizedQuery === "." || normalizedQuery === "*") {
    return true;
  }

  const effectiveQuery = buildConversationalSearchCore(normalizedQuery) || normalizedQuery;
  const searchableText = buildResultSearchableText(item, cachedContent);
  const hasCachedContent = normalizeText(cachedContent).length > 0;
  const allowNativeSearchFallback =
    !hasCachedContent &&
    !!item.matchedByNativeSearch &&
    shouldPreserveNativeSearchFallback(effectiveQuery);
  if (!searchableText) {
    return false;
  }

  if (isEmailSearch(effectiveQuery)) {
    return matchesSearchLiteral(searchableText, effectiveQuery, true);
  }

  const strictQuotedSingleTermLiteral = getStrictQuotedSingleTermLiteral(effectiveQuery);
  if (strictQuotedSingleTermLiteral) {
    return matchesStrictQuotedSingleTerm(item, strictQuotedSingleTermLiteral);
  }

  if (hasExplicitSearchSyntax(effectiveQuery)) {
    if (matchesExplicitSearchQuery(searchableText, effectiveQuery)) {
      return true;
    }

    return allowNativeSearchFallback;
  }

  const normalizedEffectiveQuery = normalizeSearchKey(effectiveQuery);
  if (normalizedEffectiveQuery.includes(" ") && searchableText.includes(normalizedEffectiveQuery)) {
    return true;
  }

  const meaningfulTerms = buildMeaningfulSearchTerms(effectiveQuery);
  if (meaningfulTerms.length === 0) {
    return matchesSearchLiteral(searchableText, effectiveQuery);
  }

  const matchedTermCount = meaningfulTerms.filter((term) => matchesSearchLiteral(searchableText, term)).length;
  if (matchedTermCount === 0) {
    return allowNativeSearchFallback;
  }

  if (meaningfulTerms.length === 1) {
    return true;
  }

  if (meaningfulTerms.length === 2) {
    return matchedTermCount === 2 || searchableText.includes(normalizedEffectiveQuery);
  }

  const minimumMatchedTerms = getMinimumRequiredTermCoverage(effectiveQuery, meaningfulTerms.length);
  if (matchedTermCount >= minimumMatchedTerms || searchableText.includes(normalizedEffectiveQuery)) {
    return true;
  }

  return allowNativeSearchFallback;
};

const normalizeFilterValue = (value?: string | null): string =>
  cleanTaxonomyTextValue(String(value || "")).trim().toLowerCase();

const getFilterValueCandidates = (value?: string | null): string[] => {
  const candidates: string[] = [];
  const rawParts = cleanTaxonomyTextValue(String(value || ""))
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);

  rawParts.forEach((entry) => {
    const normalizedEntry = normalizeFilterValue(entry);
    if (normalizedEntry) {
      candidates.push(normalizedEntry);
    }

    entry
      .split(/\s*>\s*/)
      .map((segment) => normalizeFilterValue(segment))
      .filter(Boolean)
      .forEach((segment) => candidates.push(segment));
  });

  return candidates.filter((entry, index) => candidates.indexOf(entry) === index);
};

const getFilterCountValues = (value?: string | null): string[] => {
  const values = new Map<string, string>();
  const rawParts = cleanTaxonomyTextValue(String(value || ""))
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);

  rawParts.forEach((entry) => {
    const normalizedEntry = normalizeFilterValue(entry);
    if (normalizedEntry) {
      values.set(normalizedEntry, entry);
    }

    entry
      .split(/\s*>\s*/)
      .map((segment) => segment.trim())
      .filter(Boolean)
      .forEach((segment) => {
        const normalizedSegment = normalizeFilterValue(segment);
        if (normalizedSegment) {
          values.set(normalizedSegment, segment);
        }
      });
  });

  return Array.from(values.values());
};

const filterValueMatches = (itemValue?: string, selectedValue?: string | null): boolean => {
  const normalizedSelectedValue = normalizeFilterValue(selectedValue);
  if (!normalizedSelectedValue) {
    return true;
  }

  const normalizedItemValues = getFilterValueCandidates(itemValue);
  return normalizedItemValues.some((value) => value === normalizedSelectedValue);
};

const getMetadataFilterValue = (item: ResultItem, activeMetadataFilter?: SearchMetadataFilterKey | null): string => {
  if (!activeMetadataFilter) {
    return "";
  }

  const metadataValue = item[activeMetadataFilter];
  return typeof metadataValue === "string" ? metadataValue : "";
};

const doesMetadataFilterMatchSearch = (
  item: ResultItem,
  activeMetadataFilter: SearchMetadataFilterKey | null | undefined,
  query: string
): boolean => {
  if (!activeMetadataFilter) {
    return true;
  }

  const normalizedQuery = normalizeSearchWhitespace(query);
  if (!normalizedQuery || normalizedQuery === "." || normalizedQuery === "*") {
    return true;
  }

  const scopedText = normalizeSearchKey(getMetadataFilterValue(item, activeMetadataFilter));
  if (!scopedText) {
    return false;
  }

  const effectiveQuery = buildConversationalSearchCore(normalizedQuery) || normalizedQuery;
  const normalizedEffectiveQuery = normalizeSearchKey(effectiveQuery);
  if (!normalizedEffectiveQuery) {
    return true;
  }

  if (scopedText.includes(normalizedEffectiveQuery)) {
    return true;
  }

  const terms = buildMeaningfulSearchTerms(effectiveQuery);
  if (terms.length === 0) {
    return matchesSearchLiteral(scopedText, effectiveQuery);
  }

  return terms.every((term) => matchesSearchLiteral(scopedText, term));
};

const isActiveDocument = (item: ResultItem): boolean =>
  normalizeFilterValue(item.status || "Active") === "active";

const matchesFileTypeFilter = (
  itemFileType: string | undefined,
  selectedFileType: string | null
): boolean => {
  if (!selectedFileType) {
    return true;
  }

  const normalizedItemType = String(itemFileType || "").trim().toUpperCase();
  const normalizedSelectedType = normalizeFilterValue(selectedFileType);

  if (!normalizedItemType) {
    return false;
  }

  if (normalizedSelectedType === QUICK_FILE_TYPE_VALUES.documents) {
    return DOCUMENT_FILE_TYPES.has(normalizedItemType);
  }

  if (normalizedSelectedType === QUICK_FILE_TYPE_VALUES.video) {
    return VIDEO_FILE_TYPES.has(normalizedItemType);
  }

  if (normalizedSelectedType === QUICK_FILE_TYPE_VALUES.audio) {
    return AUDIO_FILE_TYPES.has(normalizedItemType);
  }

  return normalizeFilterValue(normalizedItemType) === normalizedSelectedType;
};

const applyDocumentFilters = (items: ResultItem[], filters: FilterState): ResultItem[] =>
  items.filter((item) => {
    if (!isActiveDocument(item)) {
      return false;
    }

    if (!matchesFileTypeFilter(item.fileType, filters.fileType)) {
      return false;
    }

    return FILTER_DEFINITIONS.every((definition) => {
      if (definition.key === "fileType") {
        return true;
      }

      const selectedValue = filters[definition.key];
      if (!selectedValue) {
        return true;
      }

      return filterValueMatches(definition.getValue(item), selectedValue);
    });
  });

const mergeResultItems = (existingItems: ResultItem[], nextItems: ResultItem[]): ResultItem[] => {
  const merged = new Map<number, ResultItem>();

  existingItems.forEach((item) => {
    merged.set(item.id, item);
  });

  nextItems.forEach((item) => {
    merged.set(item.id, item);
  });

  return Array.from(merged.values());
};

export const buildKnowledgeSearchFiltersKey = (filters?: IKnowledgeSearchRequest["filters"]): string =>
  JSON.stringify(
    Object.keys(filters || {})
      .sort()
      .map((key) => {
        const value = filters?.[key];
        return [
          key,
          Array.isArray(value)
            ? value.map((entry) => String(entry)).sort()
            : value ? String(value) : ""
        ];
      })
  );

const buildSearchResultsChangeMeta = (request: IKnowledgeSearchRequest): ISearchResultsChangeMeta => {
  const filtersKey = buildKnowledgeSearchFiltersKey(request.filters);
  const query = String(request.query || "*");
  const sort = request.sort || "relevance";
  const skip = Number(request.skip || 0);
  const top = Number(request.top || PAGE_SIZES.searchResultsPerPage);

  return {
    requestKey: JSON.stringify({ query, sort, skip, top, filtersKey }),
    query,
    sort,
    skip,
    top,
    filtersKey
  };
};

const toBackendFilterValue = (value?: string | null): string | undefined =>
  value || undefined;

export const buildKnowledgeSearchFilters = (filters: FilterState): Record<string, string | undefined> => ({
  fileExtension: filters.fileType || undefined,
  bu: toBackendFilterValue(filters.businessUnit),
  department: toBackendFilterValue(filters.department),
  documentType: toBackendFilterValue(filters.documentType),
  client: toBackendFilterValue(filters.client),
  region: toBackendFilterValue(filters.region),
  therapyArea: toBackendFilterValue(filters.therapyArea),
  diseaseArea: toBackendFilterValue(filters.diseaseArea),
});

const BACKEND_FACET_KEYS_BY_FILTER: Record<keyof FilterState, string[]> = {
  fileType: ["fileExtension", "fileType"],
  businessUnit: ["businessUnit", "bu"],
  department: ["department"],
  documentType: ["documentType"],
  client: ["client"],
  region: ["geography", "region"],
  therapyArea: ["therapyArea"],
  diseaseArea: ["diseaseArea"]
};

const normalizeBackendFacetValue = (value?: string | number | null): string =>
  cleanTaxonomyTextValue(String(value || "")).trim();

const getBackendFacetOption = (
  key: keyof FilterState,
  value: string,
  count: number
): FilterOption => {
  if (key !== "department") {
    return { title: value, value, count };
  }

  const parts = value
    .split(/\s*[:>]\s*/g)
    .map((part) => part.trim())
    .filter(Boolean);

  if (parts.length === 0) {
    return { title: value, value, count };
  }

  return {
    title: parts[parts.length - 1],
    value,
    count,
    level: key === "department" ? Math.max(0, parts.length - 1) : Math.max(0, parts.length - 2),
    parentTitle: parts.length > 1 ? parts[parts.length - 2] : undefined
  };
};

const getBackendFacetValues = (facets: IKnowledgeSearchFacets | undefined, key: keyof FilterState): IKnowledgeSearchFacetValue[] => {
  const facetKeys = BACKEND_FACET_KEYS_BY_FILTER[key] || [];
  const merged = new Map<string, IKnowledgeSearchFacetValue>();
  const facetValues = facetKeys
    .map((facetKey) => facets?.[facetKey] || [])
    .find((values) => values.length > 0) || [];

  facetValues.forEach((facet) => {
    const value = normalizeBackendFacetValue(facet.value);
    if (!value) {
      return;
    }

    const normalizedKey = normalizeFilterValue(value);
    const existing = merged.get(normalizedKey);
    merged.set(normalizedKey, {
      value,
      count: Math.max(existing?.count || 0, Number(facet.count || 0))
    });
  });

  return Array.from(merged.values());
};

const getDepartmentFacetParentValue = (option: FilterOption): string => {
  if (!option.value || Number(option.level || 0) <= 0) {
    return "";
  }

  const parts = option.value
    .split(/\s*[:>]\s*/g)
    .map((part) => part.trim())
    .filter(Boolean);

  if (parts.length <= 1) {
    return "";
  }

  return parts.slice(0, -1).join(":");
};

const compareFacetOptionAlphabetically = (left: FilterOption, right: FilterOption): number => {
  return left.title.localeCompare(right.title) || right.count - left.count;
};

const orderBackendFacetOptions = (key: keyof FilterState, options: FilterOption[]): FilterOption[] => {
  if (key !== "department") {
    return options.sort(compareFacetOptionAlphabetically);
  }

  const topLevelOptions = options.filter((option) => Number(option.level || 0) === 0);
  const childOptions = options.filter((option) => Number(option.level || 0) > 0);
  const childrenByParent = new Map<string, FilterOption[]>();
  const topLevelKeys = new Set(topLevelOptions.map((option) => normalizeFilterValue(option.value || option.title)));

  childOptions.forEach((option) => {
    const parentValue = getDepartmentFacetParentValue(option);
    const parentKey = normalizeFilterValue(parentValue);

    if (!parentKey || !topLevelKeys.has(parentKey)) {
      return;
    }

    const siblings = childrenByParent.get(parentKey) || [];
    siblings.push(option);
    childrenByParent.set(parentKey, siblings);
  });

  const orderedOptions: FilterOption[] = [];
  topLevelOptions.sort(compareFacetOptionAlphabetically).forEach((option) => {
    orderedOptions.push(option);
    const childKey = normalizeFilterValue(option.value || option.title);
    const children = childrenByParent.get(childKey) || [];
    orderedOptions.push(...children.sort(compareFacetOptionAlphabetically));
  });

  return orderedOptions;
};

const buildBackendFacetFilterGroups = (facets: IKnowledgeSearchFacets | undefined): FilterGroup[] =>
  FILTER_DEFINITIONS.map((definition) => ({
    key: definition.key,
    title: definition.title,
    children: orderBackendFacetOptions(
      definition.key,
      getBackendFacetValues(facets, definition.key)
        .filter((facet) => Number(facet.count || 0) > 0)
        .map((facet) => getBackendFacetOption(definition.key, facet.value, Number(facet.count || 0)))
    )
  })).filter((group) => group.children.length > 0);

const getServerRelativePathFromUrl = (url: string | undefined): string | undefined => {
  if (!url) {
    return undefined;
  }

  try {
    return new URL(url, window.location.origin).pathname;
  } catch {
    return url.startsWith('/') ? url : undefined;
  }
};

const mapKnowledgeSearchResultToItem = (result: IKnowledgeSearchResult, index: number): ResultItem | null => {
  const listItemId = Number(result.listItemId || result.documentId || result.id);
  if (!Number.isFinite(listItemId) || listItemId <= 0) {
    return null;
  }

  const fileName = result.fileName || result.title || "Untitled";
  const fileType = (result.fileExtension || getFileTypeFromName(fileName) || "").replace(/^\./, "").toUpperCase();
  const descriptionText = String(result.description || "").trim();
  const authors = Array.isArray(result.authors)
    ? result.authors
    : String(result.authors || "")
      .split(",")
      .map((author) => author.trim())
      .filter(Boolean);
  const publishedValue =
    result.publishedDate ||
    result.created ||
    result.createdDateTime ||
    result.modified ||
    result.lastModifiedDateTime ||
    "";
  const publishedDateValue = publishedValue && !Number.isNaN(new Date(publishedValue).getTime())
    ? new Date(publishedValue).toISOString()
    : "";
  const serverRelativeUrl = result.serverRelativeUrl || result.fileRef || getServerRelativePathFromUrl(result.webUrl) || "";
  const fileUrl = result.webUrl || (serverRelativeUrl ? `${window.location.origin}${serverRelativeUrl}` : undefined);

  return {
    id: listItemId,
    title: result.title || fileName.split(".").slice(0, -1).join(".") || "Untitled Document",
    fileName,
    contributor: authors.length > 0 ? authors.join(", ") : "Internal",
    updated: publishedValue ? formatDocumentPublishedDate(publishedValue, "en-GB", {
      day: "2-digit",
      month: "short",
      year: "numeric",
    }) || formatRelativeDate(publishedValue) || publishedValue : "",
    updatedDateValue: publishedDateValue,
    description: descriptionText,
    fileUrl,
    fileType,
    serverRelativeUrl,
    fileUniqueId: result.fileUniqueId || result.uniqueId,
    businessUnit: result.businessUnit || result.bu,
    department: result.department,
    documentType: result.documentType,
    client: result.client,
    region: result.geography || result.region,
    therapyArea: result.therapyArea,
    diseaseArea: result.diseaseArea,
    status: result.status || "Active",
    viewCount: result.views || 0,
    likeCount: result.likes || 0,
    commentCount: result.comments || 0,
    shareCount: 0,
    rank: result.rerankerScore ?? result.score ?? (1000 - index),
    matchedByNativeSearch: true,
    matchedByInstantList: false,
  };
};

const getMissingBackendSearchFields = (item: ResultItem): string[] => {
  const missingFields: string[] = [];

  if (!item.fileUniqueId) missingFields.push("fileUniqueId");
  if (!item.serverRelativeUrl) missingFields.push("serverRelativeUrl");
  if (!item.fileName) missingFields.push("fileName");
  if (!item.updatedDateValue) missingFields.push("publishedDate");
  if (!item.businessUnit) missingFields.push("businessUnit/bu");
  if (!item.department) missingFields.push("department");
  if (!item.documentType) missingFields.push("documentType");
  if (!item.client) missingFields.push("client");
  if (!item.region) missingFields.push("geography/region");
  if (!item.therapyArea) missingFields.push("therapyArea");
  if (!item.diseaseArea) missingFields.push("diseaseArea");

  return missingFields;
};

const GenericSearchDropdown: React.FC<IFilterDropdownProps> = ({
  searchText,
  draftSearchText,
  isFilterPanelOpen = false,
  activeMetadataFilter = null,
  externalFilters,
  spHttpClient,
  siteUrl,
  context,
  onSuggestionSelect,
  onSearchSubmit,
  onViewDocument,
  onSuggestionsChange,
  onResultsChange,
  onFilterGroupsChange,
  renderSuggestionsSection = true,
  fastResultsOnly = false,
  isLearner = false,
  refreshKey = 0,
  resultsPage = 1,
  resultsPageSize,
  searchSortOrder = "relevance",
}) => {
  const [activeTab] = React.useState<"documents" | "experts">("documents");
  const [documents, setDocuments] = React.useState<ResultItem[]>([]);
  const [backendTotalCount, setBackendTotalCount] = React.useState<number>(0);
  const [backendPageInfo, setBackendPageInfo] = React.useState<IKnowledgeSearchPageInfo | undefined>(undefined);
  const [backendFacetGroups, setBackendFacetGroups] = React.useState<FilterGroup[]>([]);
  const [suggestionDocuments, setSuggestionDocuments] = React.useState<ResultItem[]>([]);
  const [isDocumentsLoading, setIsDocumentsLoading] = React.useState<boolean>(false);
  const [isSearchLoading, setIsSearchLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);
  const [selectedDocumentId, setSelectedDocumentId] = React.useState<number | null>(null);
  const [showAllDocumentResults, setShowAllDocumentResults] = React.useState<boolean>(false);
  const [dismissedSuggestionQuery, setDismissedSuggestionQuery] = React.useState<string>("");
  const [filters, setFilters] = React.useState<FilterState>(EMPTY_FILTERS);
  const [openGroup, setOpenGroup] = React.useState<FilterGroupTitle | null>(null);
  const [currentDocumentPage, setCurrentDocumentPage] = React.useState<number>(1);
  const [documentSortOrder, setDocumentSortOrder] = React.useState<DocumentSortOrder>("newToOld");
  const [viewMode, setViewMode] = React.useState<DocumentViewMode>("list");
  const [shouldPrefetchDocuments] = React.useState<boolean>(true);
  const [shouldPrefetchHistory] = React.useState<boolean>(false);
  const [sharedHistory, setSharedHistory] = React.useState<SearchHistoryItem[]>([]);
  const [documentsRefreshKey, setDocumentsRefreshKey] = React.useState<number>(0);
  const [currentStartRow, setCurrentStartRow] = React.useState<number>(0);
  const [hasMoreResults, setHasMoreResults] = React.useState<boolean>(false);
  const [isResultsCleared, setIsResultsCleared] = React.useState<boolean>(false);
  const [thumbnailAttemptByDocument, setThumbnailAttemptByDocument] = React.useState<Record<number, number>>({});
  const [thumbnailLoadedByDocument, setThumbnailLoadedByDocument] = React.useState<Record<number, boolean>>({});
  const [downloadingDocumentIds, setDownloadingDocumentIds] = React.useState<Record<number, boolean>>({});
  const [usesKnowledgeSearchApiResults, setUsesKnowledgeSearchApiResults] = React.useState<boolean>(false);
  const [backendRequestMeta, setBackendRequestMeta] = React.useState<ISearchResultsChangeMeta | undefined>(undefined);
  const backendPageSize = Math.min(PAGE_SIZES.searchResultsPerPage, Math.max(1, resultsPageSize || PAGE_SIZES.searchResultsPerPage));
  const backendResultSkip = Math.max(0, resultsPage > 1 ? (Math.max(1, resultsPage) - 1) * backendPageSize : currentStartRow);
  const knowledgeSearchApiClient = React.useMemo(
    () => new KnowledgeSearchApiClient(undefined, context),
    [context]
  );
  const thumbnailWebUrl = React.useMemo(
    () => context?.pageContext?.web?.absoluteUrl || siteUrl || window.location.origin,
    [context, siteUrl]
  );

  const recordedQueriesRef = React.useRef<Set<string>>(new Set());
  const listEntityTypeRef = React.useRef<string | null>(null);
  const ensuredSearchListRef = React.useRef<boolean>(false);
  const resultsSectionRef = React.useRef<HTMLElement>(null);
  const previousCommittedSearchTextRef = React.useRef<string>("");
  const latestBackendRequestKeyRef = React.useRef<string>("");
  const latestSearchAnalyticsRef = React.useRef<SearchAnalyticsContext | null>(null);

  const getCachedDocumentContent = React.useCallback((_item: ResultItem): string => "", []);

  const experts = React.useMemo<ResultItem[]>(
    () =>
      Array.from({ length: 6 }).map((_, index) => ({
        id: index + 1,
        title: `Expert ${index + 1}`,
        fileName: `expert-${index + 1}.pdf`,
        contributor: `Role ${index + 1}`,
        updated: "Nov 20, 2025",
        updatedDateValue: new Date("2025-11-20").toISOString(),
        description: "Expert profile placeholder.",
        status: "Active",
      })),
    []
  );

  const liveSearchText = (draftSearchText ?? searchText).trim();
  const committedSearchText = searchText.trim();
  const activeSearchText = committedSearchText;
  const hasSearch = activeSearchText.length > 0;
  const effectiveFilters = React.useMemo<FilterState>(() => ({
    ...filters,
    ...(externalFilters || {}),
  }), [externalFilters, filters]);
  const knowledgeSearchFilters = React.useMemo(
    () => buildKnowledgeSearchFilters(effectiveFilters),
    [
      effectiveFilters.fileType,
      effectiveFilters.businessUnit,
      effectiveFilters.department,
      effectiveFilters.documentType,
      effectiveFilters.client,
      effectiveFilters.region,
      effectiveFilters.therapyArea,
      effectiveFilters.diseaseArea,
    ]
  );
  const serverKnowledgeSearchFilters = React.useMemo(
    () => knowledgeSearchFilters,
    [knowledgeSearchFilters]
  );
  const hasAnySelectedFilter = React.useMemo(() => getFilterValues(effectiveFilters).some(Boolean), [effectiveFilters]);
  const canClearSearchResults = hasAnySelectedFilter || hasSearch || documents.length > 0;
  const shouldLoadDocuments = hasSearch || hasAnySelectedFilter;

  const fetchDocumentStats = React.useCallback(async (docIds: number[]) => {
    const stats: Record<number, { likeCount: number; commentCount: number; viewCount: number }> = {};

    try {
      const [likesResp, commentsResp] = await Promise.all([
        spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${LIST_NAMES.documentLikes}')/items?$select=DocumentId,Id`,
          SPHttpClient.configurations.v1
        ),
        spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${LIST_NAMES.documentComments}')/items?$select=DocumentId,Id`,
          SPHttpClient.configurations.v1
        ),
      ]);

      const likesByDoc: Record<number, number> = {};
      const commentsByDoc: Record<number, number> = {};

      if (likesResp.ok) {
        const likesJson = await likesResp.json();
        (likesJson.value || []).forEach((item: any) => {
          likesByDoc[item.DocumentId] = (likesByDoc[item.DocumentId] || 0) + 1;
        });
      }

      if (commentsResp.ok) {
        const commentsJson = await commentsResp.json();
        (commentsJson.value || []).forEach((item: any) => {
          commentsByDoc[item.DocumentId] = (commentsByDoc[item.DocumentId] || 0) + 1;
        });
      }

      docIds.forEach((docId) => {
        stats[docId] = {
          likeCount: likesByDoc[docId] || 0,
          commentCount: commentsByDoc[docId] || 0,
          viewCount: 0,
        };
      });
    } catch (statsError) {
      console.warn("Unable to fetch search stats:", statsError);
      docIds.forEach((docId) => {
        stats[docId] = { likeCount: 0, commentCount: 0, viewCount: 0 };
      });
    }

    return stats;
  }, [siteUrl, spHttpClient]);

  const fetchDocumentPreviewMetadata = React.useCallback(async (
    docIds: number[]
  ): Promise<Record<number, Partial<ResultItem>>> => {
    const uniqueIds = Array.from(new Set(docIds.filter((docId) => Number.isFinite(docId) && docId > 0)));
    const metadataById: Record<number, Partial<ResultItem>> = {};

    if (uniqueIds.length === 0) {
      return metadataById;
    }

    try {
      for (const chunk of chunkNumberArray(uniqueIds, SEARCH_METADATA_CHUNK_SIZE)) {
        const filter = chunk.map((docId) => `Id eq ${docId}`).join(" or ");
        const selectFields = [
          "ID",
          "Title",
          COLUMN_NAMES.description,
          DEFAULT_SEARCH_FIELD_MAP.published,
          "Modified",
          COLUMN_NAMES.status,
          COLUMN_NAMES.fileLeafRef,
          COLUMN_NAMES.fileRef,
          "File/UniqueId",
          "File/ServerRelativeUrl",
          "File/Name",
          COLUMN_NAMES.businessUnit,
          COLUMN_NAMES.department,
          COLUMN_NAMES.documentType,
          COLUMN_NAMES.client,
          COLUMN_NAMES.geography,
          COLUMN_NAMES.therapyArea,
          COLUMN_NAMES.diseaseArea,
        ].filter(Boolean).join(",");
        const response = await spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/items` +
          `?$select=${selectFields}` +
          `&$expand=File&$filter=${encodeURIComponent(filter)}`,
          SPHttpClient.configurations.v1,
          { headers: { Accept: "application/json;odata=nometadata" } }
        );

        if (!response.ok) {
          continue;
        }

        const json = await response.json();
        const items = json.value || json.d?.results || [];
        items.forEach((item: any) => {
          const itemId = Number(item.Id);
          if (!Number.isFinite(itemId) || itemId <= 0) {
            return;
          }

          const serverRelativeUrl = item.File?.ServerRelativeUrl || item.FileRef || "";
          const fileName = item.File?.Name || item.FileLeafRef || "";
          const publishedValue = resolveDocumentPublishedValue(item, DEFAULT_SEARCH_FIELD_MAP.published) || item.Modified;
          const displayDate = formatDocumentPublishedDate(publishedValue, "en-GB", {
            day: "2-digit",
            month: "short",
            year: "numeric",
          });

          metadataById[itemId] = {
            fileUniqueId: item.File?.UniqueId || "",
            serverRelativeUrl,
            fileUrl: serverRelativeUrl ? `${window.location.origin}${serverRelativeUrl}` : undefined,
            fileName: fileName || undefined,
            fileType: fileName ? getFileTypeFromName(fileName) : undefined,
            businessUnit: extractStringValue(item[COLUMN_NAMES.businessUnit] ?? item.BU ?? item.BusinessUnit),
            department: extractStringValue(item[COLUMN_NAMES.department] ?? item.Department),
            documentType: extractStringValue(item[COLUMN_NAMES.documentType] ?? item.DocumentType),
            client: extractStringValue(item[COLUMN_NAMES.client] ?? item.Client),
            region: extractStringValue(item[COLUMN_NAMES.geography] ?? item.Geography ?? item.Region),
            therapyArea: extractStringValue(item[COLUMN_NAMES.therapyArea] ?? item.TherapyArea),
            diseaseArea: extractStringValue(item[COLUMN_NAMES.diseaseArea] ?? item.DiseaseArea),
            updated: displayDate || formatRelativeDate(String(publishedValue || "")),
            updatedDateValue: publishedValue ? new Date(publishedValue).toISOString() : undefined,
          };
        });
      }
    } catch (previewMetadataError) {
      console.warn("Unable to hydrate search preview metadata:", previewMetadataError);
    }

    return metadataById;
  }, [siteUrl, spHttpClient]);

  const ensureSearchListExists = React.useCallback(async (): Promise<boolean> => {
    ensuredSearchListRef.current = true;
    return false;
  }, []);

  const getSearchListEntityType = React.useCallback(async (): Promise<string | null> => {
    if (listEntityTypeRef.current) {
      return listEntityTypeRef.current;
    }

    try {
      const response = await spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${SEARCH_QUERY_LIST}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return null;
      }

      const json = await response.json();
      listEntityTypeRef.current = json.ListItemEntityTypeFullName || null;
      return listEntityTypeRef.current;
    } catch (entityError) {
      console.warn("Unable to fetch search query list entity type:", entityError);
      return null;
    }
  }, [siteUrl, spHttpClient]);

  React.useEffect(() => {
    return subscribeToDocumentDataChanged(() => {
      setDocumentsRefreshKey((value) => value + 1);
    });
  }, []);

  React.useEffect(() => {
    if (!shouldLoadDocuments) {
      return;
    }

    let isCancelled = false;

    const fetchDocuments = async () => {
      let activeRequestKey = "";
      try {
        const rawText = committedSearchText || '';
        const isPeriod = rawText === '.';
        const shouldShowAllActiveDocuments = isPeriod || rawText === '*';

        if (previousCommittedSearchTextRef.current !== rawText && currentStartRow !== 0) {
          previousCommittedSearchTextRef.current = rawText;
          setCurrentStartRow(0);
          return;
        }
        previousCommittedSearchTextRef.current = rawText;

        if (backendResultSkip === 0) {
          setThumbnailAttemptByDocument({});
          setThumbnailLoadedByDocument({});
          setDocuments([]);
          setBackendTotalCount(0);
          setBackendPageInfo(undefined);
          setBackendFacetGroups([]);
        }

        setUsesKnowledgeSearchApiResults(false);
        setIsDocumentsLoading(true);
        setError(null);
        setHasMoreResults(false);

        if (!knowledgeSearchApiClient.isConfigured()) {
          throw new Error("Knowledge Search API is not configured. Main search requires the Azure AI Search backend.");
        }

        const backendQuery = shouldShowAllActiveDocuments ? "*" : rawText.trim();
        const backendRequest = {
          query: backendQuery || "*",
          top: backendPageSize,
          skip: backendResultSkip,
          includeTotalCount: true,
          includeFacets: true,
          responseShape: "paged" as const,
          sort: searchSortOrder,
          filters: serverKnowledgeSearchFilters,
          analytics: {
            recordWildcardSearch: rawText.trim() === "*",
          },
        };
        const requestMeta = buildSearchResultsChangeMeta(backendRequest);
        activeRequestKey = requestMeta.requestKey;
        latestBackendRequestKeyRef.current = requestMeta.requestKey;
        setBackendRequestMeta(requestMeta);
        searchDebugLog("[Search] Backend request", backendRequest);

        const backendResponse: IKnowledgeSearchDocumentsResponse = await knowledgeSearchApiClient.searchDocuments(backendRequest);

        if (isCancelled || latestBackendRequestKeyRef.current !== requestMeta.requestKey) return;

        searchDebugLog("[Search] Backend response summary", {
          results: backendResponse.results.length,
          totalCount: backendResponse.totalCount,
          page: backendResponse.page,
          facetKeys: Object.keys(backendResponse.facets || {}),
        });
        const nextBackendFacetGroups = buildBackendFacetFilterGroups(backendResponse.facets);
        searchDebugLog("[Search] Backend facet groups mapped for UI", {
          groupKeys: nextBackendFacetGroups.map((group) => group.key),
          departmentOptionCount: nextBackendFacetGroups.find((group) => group.key === "department")?.children.length || 0,
          topDepartmentOptions: (nextBackendFacetGroups.find((group) => group.key === "department")?.children || [])
            .slice(0, 20)
            .map((option) => ({
              title: option.title,
              value: option.value || option.title,
              count: option.count,
              level: option.level || 0,
            })),
        });

        const backendDocs = backendResponse.results
          .map((item, index) => mapKnowledgeSearchResultToItem(item, index))
          .filter((item): item is ResultItem => Boolean(item));
        const docsMissingBackendMetadata = fastResultsOnly
          ? []
          : backendDocs.filter((item) => getMissingBackendSearchFields(item).length > 0);

        if (docsMissingBackendMetadata.length > 0) {
          searchDebugLog("[Search] SharePoint fallback hydration used for missing backend fields", {
            count: docsMissingBackendMetadata.length,
            sample: docsMissingBackendMetadata.slice(0, 10).map((item) => ({
              id: item.id,
              title: item.title,
              missingFields: getMissingBackendSearchFields(item),
            })),
          });
        } else {
          searchDebugLog("[Search] Backend supplied all required search display fields; SharePoint fallback hydration skipped.");
        }

        const previewMetadata = docsMissingBackendMetadata.length > 0
          ? await fetchDocumentPreviewMetadata(docsMissingBackendMetadata.map((item) => item.id))
          : {};

        if (isCancelled) return;

        const hydratedBackendDocs: ResultItem[] = backendDocs.map((item) => ({
          ...item,
          ...(previewMetadata[item.id] || {}),
          fileUniqueId: previewMetadata[item.id]?.fileUniqueId || item.fileUniqueId,
          serverRelativeUrl: previewMetadata[item.id]?.serverRelativeUrl || item.serverRelativeUrl,
          fileUrl: previewMetadata[item.id]?.fileUrl || item.fileUrl,
          fileName: previewMetadata[item.id]?.fileName || item.fileName,
          fileType: previewMetadata[item.id]?.fileType || item.fileType,
          updated: item.updated || previewMetadata[item.id]?.updated || "",
          updatedDateValue: item.updatedDateValue || previewMetadata[item.id]?.updatedDateValue || "",
          businessUnit: item.businessUnit || previewMetadata[item.id]?.businessUnit,
          department: item.department || previewMetadata[item.id]?.department,
          documentType: item.documentType || previewMetadata[item.id]?.documentType,
          client: item.client || previewMetadata[item.id]?.client,
          region: item.region || previewMetadata[item.id]?.region,
          therapyArea: item.therapyArea || previewMetadata[item.id]?.therapyArea,
          diseaseArea: item.diseaseArea || previewMetadata[item.id]?.diseaseArea,
        }));

        latestSearchAnalyticsRef.current = {
          searchRequestId: backendResponse.analytics?.searchRequestId,
          searchSessionId: backendResponse.analytics?.searchSessionId,
          startedAt: Date.now(),
          ranksByListItemId: hydratedBackendDocs.reduce<Record<number, number>>((acc, item, index) => {
            acc[item.id] = backendResultSkip + index + 1;
            return acc;
          }, {}),
        };

        setUsesKnowledgeSearchApiResults(true);
        setBackendTotalCount(Number(backendResponse.totalCount || hydratedBackendDocs.length));
        setBackendPageInfo(backendResponse.page);
        setBackendFacetGroups(nextBackendFacetGroups);
        setHasMoreResults(Boolean(backendResponse.page?.hasMore));
        searchDebugLog("[Search] Frontend render source summary", {
          documentsFromBackend: backendDocs.length,
          documentsHydratedFromSharePointFallback: docsMissingBackendMetadata.length,
          countSource: typeof backendResponse.totalCount === "number" ? "backend.totalCount" : "frontend fallback length",
          paginationSource: backendResponse.page ? "backend.page" : "frontend fallback",
          facetsSource: backendResponse.facets && Object.keys(backendResponse.facets).length > 0 ? "backend.facets" : "frontend fallback",
          thumbnailSource: "frontend URL builder using backend fileUniqueId + KM_REVIEW_HUB_DRIVE_ID",
          publishedDateSource: "backend.publishedDate; SharePoint fallback only if missing",
          authorSource: "backend.authors; no logged-in user fallback",
        });
        setDocuments((prev) => currentStartRow > 0 && backendResultSkip > 0 ? mergeResultItems(prev, hydratedBackendDocs) : hydratedBackendDocs);
        setSuggestionDocuments((prev) => currentStartRow > 0 && backendResultSkip > 0 ? mergeResultItems(prev, hydratedBackendDocs) : hydratedBackendDocs);

        if (!fastResultsOnly && hydratedBackendDocs.length > 0) {
          fetchDocumentStats(hydratedBackendDocs.map((item) => item.id)).then(stats => {
            if (latestBackendRequestKeyRef.current !== requestMeta.requestKey) return;
            setDocuments(prev => prev.map(item => ({
              ...item,
              ...(stats[item.id] || {}),
            })));
          }).catch(err => console.warn("Stats fetch failed quietly:", err));
        }
      } catch (fetchError: any) {
        if (activeRequestKey && latestBackendRequestKeyRef.current !== activeRequestKey) return;
        console.error("Error fetching search documents:", fetchError);
        setError(fetchError.message || "Failed to load search documents");
      } finally {
        if (!activeRequestKey || latestBackendRequestKeyRef.current === activeRequestKey) {
          setIsDocumentsLoading(false);
        }
      }
    };

    void fetchDocuments();

    return () => {
      isCancelled = true;
    };
  }, [committedSearchText, fastResultsOnly, fetchDocumentPreviewMetadata, fetchDocumentStats, hasSearch, knowledgeSearchApiClient, serverKnowledgeSearchFilters, shouldLoadDocuments, currentStartRow, documentsRefreshKey, refreshKey, backendPageSize, backendResultSkip, searchSortOrder]);

  React.useEffect(() => {
    const localHistoryRaw = window.localStorage.getItem(LOCAL_HISTORY_KEY);
    if (localHistoryRaw) {
      try {
        const parsed = JSON.parse(localHistoryRaw) as SearchHistoryItem[];
        if (Array.isArray(parsed)) {
          setSharedHistory(parsed);
        }
      } catch (localHistoryError) {
        console.warn("Unable to parse local search history:", localHistoryError);
      }
    }

    if (!shouldPrefetchHistory) {
      return;
    }

    const fetchSharedHistory = async () => {
      try {
        const response = await spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${SEARCH_QUERY_LIST}')/items?$select=Id,Title,QueryText,SearchCount&$orderby=Modified desc&$top=20`,
          SPHttpClient.configurations.v1
        );

        if (!response.ok) {
          return;
        }

        const json = await response.json();
        const history = (json.value || [])
          .map((item: any) => ({
            id: item.Id,
            value: String(item.QueryText || item.Title || "").trim(),
            count: Number(item.SearchCount || 0),
          }))
          .filter((item: SearchHistoryItem) => item.value.length > 0);

        if (history.length > 0) {
          setSharedHistory((current) => {
            const merged = new Map<string, SearchHistoryItem>();
            [...current, ...history].forEach((item) => {
              merged.set(normalizeText(item.value), item);
            });

            return Array.from(merged.values()).sort((left, right) => right.count - left.count);
          });
        }
      } catch (historyError) {
        console.warn("Unable to fetch shared search history:", historyError);
      }
    };

    void fetchSharedHistory();
  }, [shouldPrefetchHistory, siteUrl, spHttpClient]);

  const persistLocalHistory = React.useCallback((value: string) => {
    setSharedHistory((current) => {
      const normalized = normalizeText(value);
      const nextMap = new Map<string, SearchHistoryItem>();
      current.forEach((item) => nextMap.set(normalizeText(item.value), item));

      const existing = nextMap.get(normalized);
      nextMap.set(normalized, {
        id: existing?.id,
        value,
        count: (existing?.count || 0) + 1,
      });

      const nextHistory = Array.from(nextMap.values())
        .filter((item) => item.value.trim().length > 0)
        .sort((left, right) => right.count - left.count)
        .slice(0, 20);

      window.localStorage.setItem(LOCAL_HISTORY_KEY, JSON.stringify(nextHistory));
      return nextHistory;
    });
  }, []);

  const recordSharedQuery = React.useCallback(async (value: string) => {
    const trimmed = value.trim();
    if (!trimmed) {
      return;
    }

    const normalized = normalizeText(trimmed);
    if (recordedQueriesRef.current.has(normalized)) {
      return;
    }

    recordedQueriesRef.current.add(normalized);
    persistLocalHistory(trimmed);

    try {
      const listReady = await ensureSearchListExists();
      if (!listReady) {
        return;
      }

      const entityType = await getSearchListEntityType();
      if (!entityType) {
        return;
      }

      // Properly encode the filter to avoid 400 errors for queries with spaces
      const filterStr = `QueryText eq '${escapeODataValue(trimmed)}'`;
      const lookupResponse = await spHttpClient.get(
        `${siteUrl}/_api/web/lists/getbytitle('${SEARCH_QUERY_LIST}')/items?$select=Id,SearchCount&$filter=${encodeURIComponent(filterStr)}&$top=1`,
        SPHttpClient.configurations.v1
      );

      if (lookupResponse.ok) {
        const lookupJson = await lookupResponse.json();
        const existingItem = (lookupJson.value || [])[0];

        if (existingItem?.Id) {
          await spHttpClient.post(
            `${siteUrl}/_api/web/lists/getbytitle('${SEARCH_QUERY_LIST}')/items(${existingItem.Id})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",
              },
              body: JSON.stringify({
                __metadata: { type: entityType },
                Title: trimmed,
                QueryText: trimmed,
                SearchCount: Number(existingItem.SearchCount || 0) + 1,
              }),
            }
          );
          return;
        }
      }

      await spHttpClient.post(
        `${siteUrl}/_api/web/lists/getbytitle('${SEARCH_QUERY_LIST}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
          },
          body: JSON.stringify({
            __metadata: { type: entityType },
            Title: trimmed,
            QueryText: trimmed,
            SearchCount: 1,
          }),
        }
      );
    } catch (recordError) {
      console.warn("Unable to store shared search query:", recordError);
    }
  }, [ensureSearchListExists, getSearchListEntityType, persistLocalHistory, siteUrl, spHttpClient]);

  React.useEffect(() => {
    if (!committedSearchText) {
      return;
    }

    void recordSharedQuery(committedSearchText);
  }, [committedSearchText, recordSharedQuery]);

  const activeDocuments = React.useMemo(
    () => documents.filter((item) => isActiveDocument(item)),
    [documents]
  );

  const documentResultsBase = React.useMemo(
    () => usesKnowledgeSearchApiResults
      ? activeDocuments
      : applyDocumentFilters(activeDocuments, effectiveFilters),
    [activeDocuments, effectiveFilters, usesKnowledgeSearchApiResults]
  );

  const preliminaryRankedDocuments = React.useMemo(() => {
    const baseItems = documentResultsBase.slice();
    if (!hasSearch) {
      return baseItems.sort(
        (left, right) =>
          new Date(right.updatedDateValue).getTime() - new Date(left.updatedDateValue).getTime()
      );
    }

    const rawRankingQuery = normalizeSearchWhitespace(activeSearchText || "");
    const conversationalRankingQuery = buildConversationalSearchCore(rawRankingQuery);
    const rankingQuery = hasExplicitSearchSyntax(conversationalRankingQuery)
      ? buildListSearchLiterals(conversationalRankingQuery).join(" ")
      : conversationalRankingQuery;

    // Prefer documents with stronger real text matches while still honoring SharePoint native rank.
    return baseItems.sort((left, right) => {
      const query = rankingQuery || rawRankingQuery;
      if (!query) return (right.rank || 0) - (left.rank || 0);

      const leftContent = getCachedDocumentContent(left);
      const rightContent = getCachedDocumentContent(right);
      const leftProfile = buildSearchRelevanceProfile(left, query, leftContent);
      const rightProfile = buildSearchRelevanceProfile(right, query, rightContent);

      if (leftProfile.coverageScore !== rightProfile.coverageScore) {
        return rightProfile.coverageScore - leftProfile.coverageScore;
      }

      const leftKeywordMatches = countKeywordMatches(left, query, leftContent);
      const rightKeywordMatches = countKeywordMatches(right, query, rightContent);

      if (leftKeywordMatches !== rightKeywordMatches) {
        return rightKeywordMatches - leftKeywordMatches;
      }

      const exactPriorityDelta = getExactMatchPriority(left, query) - getExactMatchPriority(right, query);
      if (exactPriorityDelta !== 0) {
        return exactPriorityDelta;
      }

      const leftRank = Math.min(left.rank || 0, MAX_NATIVE_RANK_SCORE);
      const rightRank = Math.min(right.rank || 0, MAX_NATIVE_RANK_SCORE);
      if (leftRank !== rightRank) {
        return rightRank - leftRank;
      }

      const leftScore = scoreResult(left, query, leftContent);
      const rightScore = scoreResult(right, query, rightContent);
      if (leftScore !== rightScore) {
        return rightScore - leftScore;
      }

      // Tie-breaker: Date
      return (
        new Date(right.updatedDateValue).getTime() -
        new Date(left.updatedDateValue).getTime()
      );
    });
  }, [activeSearchText, documentResultsBase, getCachedDocumentContent, hasSearch]);

  // Legacy client-side document content parsing was archived under .archive/search-legacy. Main search now uses the Knowledge Search API only.

  const filteredItems = React.useMemo(() => {
    if (activeTab === "experts") {
      if (!liveSearchText) {
        return experts;
      }

      const normalizedSearch = normalizeText(liveSearchText);
      return experts.filter((item) =>
        [item.title, item.description, item.contributor]
          .map((value) => normalizeText(value))
          .some((value) => value.includes(normalizedSearch))
      );
    }

    if (!hasSearch) {
      return documentResultsBase;
    }

    const normalizedActiveQuery = normalizeSearchWhitespace(activeSearchText || "");
    if (!normalizedActiveQuery || normalizedActiveQuery === "." || normalizedActiveQuery === "*") {
      return preliminaryRankedDocuments;
    }

    if (usesKnowledgeSearchApiResults) {
      return documentResultsBase;
    }

    return preliminaryRankedDocuments.filter((item) => {
      if (!doesMetadataFilterMatchSearch(item, activeMetadataFilter, normalizedActiveQuery)) {
        return false;
      }

      return doesResultMatchSearchIntent(
        item,
        normalizedActiveQuery,
        getCachedDocumentContent(item)
      );
    });
  }, [
    activeSearchText,
    activeTab,
    activeMetadataFilter,
    documentResultsBase,
    experts,
    getCachedDocumentContent,
    hasSearch,
    liveSearchText,
    preliminaryRankedDocuments,
    usesKnowledgeSearchApiResults,
  ]);

  const sortedFilteredItems = React.useMemo(() => {
    if (activeTab !== "documents") {
      return filteredItems;
    }

    const getSortTime = (item: ResultItem): number => {
      const normalizedTime = new Date(item.updatedDateValue || item.updated || "").getTime();
      return Number.isNaN(normalizedTime) ? 0 : normalizedTime;
    };

    if (hasSearch) {
      if (usesKnowledgeSearchApiResults) {
        return filteredItems;
      }

      const rawRankingQuery = normalizeSearchWhitespace(activeSearchText || "");
      const conversationalRankingQuery = buildConversationalSearchCore(rawRankingQuery);
      const rankingQuery = isSentenceLikeSearch(rawRankingQuery)
        ? rawRankingQuery
        : conversationalRankingQuery || rawRankingQuery;

      return [...filteredItems].sort((left, right) => {
      const leftContent = getCachedDocumentContent(left);
      const rightContent = getCachedDocumentContent(right);
      const leftProfile = buildSearchRelevanceProfile(left, rankingQuery, leftContent);
      const rightProfile = buildSearchRelevanceProfile(right, rankingQuery, rightContent);

      if (leftProfile.coverageScore !== rightProfile.coverageScore) {
        return rightProfile.coverageScore - leftProfile.coverageScore;
      }

      const leftKeywordMatches = countKeywordMatches(left, rankingQuery, leftContent);
      const rightKeywordMatches = countKeywordMatches(right, rankingQuery, rightContent);

        if (leftKeywordMatches !== rightKeywordMatches) {
          return rightKeywordMatches - leftKeywordMatches;
        }

        const leftScore = scoreResult(left, rankingQuery, leftContent);
        const rightScore = scoreResult(right, rankingQuery, rightContent);
        if (leftScore !== rightScore) {
          return rightScore - leftScore;
        }

        const exactPriorityDelta = getExactMatchPriority(left, rankingQuery) - getExactMatchPriority(right, rankingQuery);
        if (exactPriorityDelta !== 0) {
          return exactPriorityDelta;
        }

        const leftRank = Math.min(left.rank || 0, MAX_NATIVE_RANK_SCORE);
        const rightRank = Math.min(right.rank || 0, MAX_NATIVE_RANK_SCORE);
        if (leftRank !== rightRank) {
          return rightRank - leftRank;
        }

        return getSortTime(right) - getSortTime(left);
      });
    }

    return [...filteredItems].sort((left, right) => {
      const leftTime = getSortTime(left);
      const rightTime = getSortTime(right);
      return documentSortOrder === "newToOld" ? rightTime - leftTime : leftTime - rightTime;
    });
  }, [activeSearchText, activeTab, documentSortOrder, filteredItems, getCachedDocumentContent, hasSearch, usesKnowledgeSearchApiResults]);

  const filterGroups = React.useMemo<FilterGroup[]>(() => {
    if (backendFacetGroups.length > 0) {
      return backendFacetGroups;
    }

    if (fastResultsOnly || hasSearch || hasAnySelectedFilter || usesKnowledgeSearchApiResults || isDocumentsLoading) {
      return [];
    }

    const filterSourceDocuments =
      preliminaryRankedDocuments.length > 0 ? preliminaryRankedDocuments : activeDocuments;

    return FILTER_DEFINITIONS.map((definition) => {
      const scopedFilters: FilterState = {
        ...effectiveFilters,
        [definition.key]: null,
      };

      const counts = new Map<string, number>();

      applyDocumentFilters(filterSourceDocuments, scopedFilters)
        .forEach((item) => {
          getFilterCountValues(definition.getValue(item)).forEach((value) => {
            counts.set(value, (counts.get(value) || 0) + 1);
          });
        });

      return {
        key: definition.key,
        title: definition.title,
        children: Array.from(counts.entries())
          .map(([title, count]) => ({ title, count }))
          .sort((left, right) => {
            return left.title.localeCompare(right.title) || right.count - left.count;
          }),
      };
    }).filter((group) => group.children.length > 0);
  }, [
    activeDocuments,
    backendFacetGroups,
    effectiveFilters,
    fastResultsOnly,
    hasAnySelectedFilter,
    hasSearch,
    isDocumentsLoading,
    preliminaryRankedDocuments,
    usesKnowledgeSearchApiResults
  ]);

  React.useEffect(() => {
    onFilterGroupsChange?.(filterGroups);
  }, [filterGroups, onFilterGroupsChange]);

  const suggestions = React.useMemo<SuggestionItem[]>(() => {
    if (liveSearchText.length < 1) {
      return [];
    }

    const normalizedSearch = normalizeText(liveSearchText);
    if (!normalizedSearch) {
      return [];
    }

    const suggestionSourceMap = new Map<number, ResultItem>();
    [...suggestionDocuments, ...documents].forEach((item) => {
      if (item?.id != null) {
        suggestionSourceMap.set(item.id, item);
      }
    });

    const rankedSuggestions = new Map<string, SuggestionItem & { score: number }>();
    const addSuggestion = (
      rawValue: string | undefined,
      source: SuggestionItem["source"],
      meta: string,
      scoreBase: number
    ): void => {
      const value = String(rawValue || "").trim();
      if (!value) {
        return;
      }

      const normalizedValue = normalizeText(value);
      if (!normalizedValue.includes(normalizedSearch)) {
        return;
      }

      const score = scoreBase + (normalizedValue.startsWith(normalizedSearch) ? 50 : 0) - value.length / 1000;
      const existing = rankedSuggestions.get(normalizedValue);
      if (!existing || score > existing.score) {
        rankedSuggestions.set(normalizedValue, {
          value,
          source,
          meta,
          score,
        });
      }
    };

    Array.from(suggestionSourceMap.values()).forEach((item) => {
      const title = item.title.trim();
      const fileName = item.fileName.trim();
      const matchedFields: Array<{ value?: string; meta: string; scoreBase: number }> = [
        { value: item.contributor, meta: "Author match", scoreBase: 260 },
        { value: item.businessUnit, meta: "BU match", scoreBase: 250 },
        { value: item.department, meta: "Department match", scoreBase: 250 },
        { value: item.documentType, meta: "Document type match", scoreBase: 245 },
        { value: item.client, meta: "Client match", scoreBase: 240 },
        { value: item.region, meta: "Region match", scoreBase: 235 },
        { value: item.therapyArea, meta: "Therapy area match", scoreBase: 230 },
        { value: item.diseaseArea, meta: "Disease area match", scoreBase: 230 },
        { value: item.description, meta: "Description match", scoreBase: 180 },
      ];

      addSuggestion(title, "title", "Title match", 320);
      addSuggestion(fileName, "title", "File name match", 300);

      matchedFields.forEach(({ value, meta, scoreBase }) => {
        addSuggestion(value, "title", meta, scoreBase);

        if (value && title) {
          const normalizedValue = normalizeText(String(value));
          if (normalizedValue.includes(normalizedSearch)) {
            addSuggestion(title, "title", meta, scoreBase + 20);
          }
        }
      });
    });

    const titleSuggestions = Array.from(rankedSuggestions.values())
      .sort((left, right) => {
        if (right.score !== left.score) {
          return right.score - left.score;
        }

        return left.value.localeCompare(right.value);
      })
      .slice(0, 6)
      .map(({ score: _score, ...suggestion }) => suggestion);

    const querySuggestions = sharedHistory
      .filter((item) => normalizeText(item.value).includes(normalizedSearch))
      .sort((left, right) => right.count - left.count)
      .slice(0, 6)
      .map((item) => ({
        value: item.value,
        source: "query" as const,
        count: item.count,
        meta: `${item.count || 0} searches`,
      }));

    const merged = new Map<string, SuggestionItem>();
    [...querySuggestions, ...titleSuggestions].forEach((item) => {
      const key = normalizeText(item.value);
      if (!merged.has(key)) {
        merged.set(key, item);
      }
    });

    return Array.from(merged.values()).slice(0, 8);
  }, [documents, liveSearchText, sharedHistory, suggestionDocuments]);

  React.useEffect(() => {
    onSuggestionsChange?.(suggestions);
  }, [onSuggestionsChange, suggestions]);

  const submitSuggestion = React.useCallback((query: string) => {
    const trimmed = query.trim();
    if (!trimmed) {
      return;
    }

    setDismissedSuggestionQuery(trimmed);
    onSuggestionSelect?.(trimmed);
    onSearchSubmit?.(trimmed);
    window.requestAnimationFrame(() => {
      resultsSectionRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  }, [onSearchSubmit, onSuggestionSelect]);

  const updateFilter = React.useCallback((key: keyof FilterState, value: string) => {
    setFilters((current) => {
      const nextValue = normalizeFilterValue(current[key]) === normalizeFilterValue(value) ? null : value;
      const nextFilters: FilterState = {
        ...current,
        [key]: nextValue,
      };

      if (key === "businessUnit" && normalizeFilterValue(current.businessUnit) !== normalizeFilterValue(nextValue)) {
        nextFilters.department = null;
      }

      return nextFilters;
    });
    setCurrentDocumentPage(1);
    setOpenGroup(null);
  }, []);

  const clearFilters = React.useCallback(() => {
    setFilters(EMPTY_FILTERS);
    setOpenGroup(null);
    setCurrentDocumentPage(1);
    setCurrentStartRow(0);
    setHasMoreResults(false);
    setShowAllDocumentResults(false);
    setThumbnailAttemptByDocument({});
    setThumbnailLoadedByDocument({});
    setDocuments([]);
    setSuggestionDocuments([]);
    setDismissedSuggestionQuery("");
    setIsSearchLoading(false);
    setIsDocumentsLoading(false);
    setIsResultsCleared(true);
    latestSearchAnalyticsRef.current = null;
    onSearchSubmit?.("");
  }, [onSearchSubmit]);

  const trackSearchResultAction = React.useCallback((action: "view" | "download", item: ResultItem) => {
    const analyticsContext = latestSearchAnalyticsRef.current;
    if (!analyticsContext?.searchRequestId) {
      return;
    }

    void knowledgeSearchApiClient.trackSearchAction({
      action,
      searchRequestId: analyticsContext.searchRequestId,
      searchSessionId: analyticsContext.searchSessionId,
      documentId: item.fileUniqueId || item.serverRelativeUrl || String(item.id),
      listItemId: String(item.id),
      documentTitle: item.title,
      rank: analyticsContext.ranksByListItemId[item.id],
      timeSinceSearchMs: Math.max(0, Date.now() - analyticsContext.startedAt),
      sourceSurface: "search-results",
    });
  }, [knowledgeSearchApiClient]);

  const handleView = React.useCallback((item: ResultItem) => {
    trackSearchResultAction("view", item);

    if (onViewDocument) {
      onViewDocument(item.id);
      return;
    }

    setSelectedDocumentId(item.id);
  }, [onViewDocument, trackSearchResultAction]);

  const handleCloseDetail = React.useCallback(() => {
    setSelectedDocumentId(null);
  }, []);

  const handleCloseAllDocumentResults = React.useCallback(() => {
    setShowAllDocumentResults(false);
  }, []);

  const handleDownload = React.useCallback(async (item: ResultItem) => {
    try {
      if (!item.serverRelativeUrl) {
        return;
      }
      if (downloadingDocumentIds[item.id]) {
        return;
      }
      trackSearchResultAction("download", item);
      setDownloadingDocumentIds((current) => ({ ...current, [item.id]: true }));

      const serverRelativeUrl = item.serverRelativeUrl.startsWith('/') ? item.serverRelativeUrl : `/${item.serverRelativeUrl}`;
      await downloadSharePointFile(
        context,
        siteUrl,
        serverRelativeUrl,
        item.fileName || item.title || "download"
      );
    } catch (downloadError) {
      console.error("Download failed:", downloadError);
    } finally {
      window.setTimeout(() => {
        setDownloadingDocumentIds((current) => {
          const next = { ...current };
          delete next[item.id];
          return next;
        });
      }, 1200);
    }
  }, [context, downloadingDocumentIds, siteUrl, trackSearchResultAction]);

  const shouldShowSuggestions =
    suggestions.length > 0 &&
    normalizeText(liveSearchText) !== normalizeText(dismissedSuggestionQuery);
  const totalDocumentPages = activeTab === "documents"
    ? Math.max(1, Math.ceil(sortedFilteredItems.length / DOCUMENT_RESULTS_PER_PAGE))
    : 1;
  const paginatedDocumentItems = activeTab === "documents"
    ? sortedFilteredItems.slice(
      (currentDocumentPage - 1) * DOCUMENT_RESULTS_PER_PAGE,
      currentDocumentPage * DOCUMENT_RESULTS_PER_PAGE
    )
    : [];

  React.useEffect(() => {
    onResultsChange?.(
      committedSearchText || hasAnySelectedFilter ? sortedFilteredItems : [],
      isDocumentsLoading || isSearchLoading,
      backendTotalCount,
      backendPageInfo,
      backendRequestMeta
    );
  }, [backendPageInfo, backendRequestMeta, backendTotalCount, committedSearchText, hasAnySelectedFilter, isDocumentsLoading, isSearchLoading, onResultsChange, sortedFilteredItems]);

  React.useEffect(() => {
    setCurrentDocumentPage(1);
    setCurrentStartRow(0);
  }, [committedSearchText, effectiveFilters, activeTab, searchSortOrder]);

  React.useEffect(() => {
    if (committedSearchText || hasAnySelectedFilter) {
      setIsResultsCleared(false);
    }
  }, [committedSearchText, hasAnySelectedFilter]);

  React.useEffect(() => {
    if (currentDocumentPage > totalDocumentPages) {
      setCurrentDocumentPage(totalDocumentPages);
    }
  }, [currentDocumentPage, totalDocumentPages]);

  const paginatedStartIndex = sortedFilteredItems.length === 0 ? 0 : ((currentDocumentPage - 1) * DOCUMENT_RESULTS_PER_PAGE) + 1;
  const paginatedEndIndex = activeTab === "documents"
    ? Math.min(currentDocumentPage * DOCUMENT_RESULTS_PER_PAGE, sortedFilteredItems.length)
    : sortedFilteredItems.length;
  const documentPaginationItems = React.useMemo(
    () => buildPaginationItems(currentDocumentPage, totalDocumentPages),
    [currentDocumentPage, totalDocumentPages]
  );

  React.useEffect(() => {
    const visibleThumbnailItems = (showAllDocumentResults ? filteredItems : paginatedDocumentItems).slice(0, 24);
    const preloaders: HTMLImageElement[] = [];

    visibleThumbnailItems.forEach((item) => {
      const thumbnailCandidates = getSearchThumbnailCandidates(
        item.serverRelativeUrl,
        thumbnailWebUrl,
        item.fileType,
        item.fileUrl,
        item.fileUniqueId
      );
      const activeThumbnailIndex = thumbnailAttemptByDocument[item.id] || 0;
      const activeThumbnailUrl = thumbnailCandidates[activeThumbnailIndex] || "";

      if (!activeThumbnailUrl || thumbnailLoadedByDocument[item.id]) {
        return;
      }

      const preloader = new Image();
      preloader.decoding = "async";
      preloader.onload = () => {
        setThumbnailLoadedByDocument((current) => (
          current[item.id]
            ? current
            : {
                ...current,
                [item.id]: true,
              }
        ));
      };
      preloader.onerror = () => {
        setThumbnailLoadedByDocument((current) => {
          if (!current[item.id]) {
            return current;
          }

          const nextState = { ...current };
          delete nextState[item.id];
          return nextState;
        });
        setThumbnailAttemptByDocument((current) => ({
          ...current,
          [item.id]: activeThumbnailIndex + 1,
        }));
      };
      preloader.src = activeThumbnailUrl;
      preloaders.push(preloader);
    });

    return () => {
      preloaders.forEach((preloader) => {
        preloader.onload = null;
        preloader.onerror = null;
      });
    };
  }, [
    filteredItems,
    paginatedDocumentItems,
    showAllDocumentResults,
    thumbnailWebUrl,
    thumbnailAttemptByDocument,
    thumbnailLoadedByDocument,
  ]);

  const renderSearchGridThumbnail = React.useCallback((item: ResultItem, fallbackIcon?: string): JSX.Element => {
    const thumbnailCandidates = getSearchThumbnailCandidates(
      item.serverRelativeUrl,
      thumbnailWebUrl,
      item.fileType,
      item.fileUrl,
      item.fileUniqueId
    );
    const activeThumbnailIndex = thumbnailAttemptByDocument[item.id] || 0;
    const activeThumbnailUrl = thumbnailCandidates[activeThumbnailIndex] || "";
    const isThumbnailLoaded = !!thumbnailLoadedByDocument[item.id];
    const isExcelPreview = isExcelFileType(item.fileType);
    const isAudioPreview = isAudioFileType(item.fileType);
    const isVideoPreview = isVideoFileType(item.fileType);
    const videoThumbnailSource = getAbsoluteSearchFileUrl(item.serverRelativeUrl, thumbnailWebUrl, item.fileUrl);
    const normalizedFileType = String(item.fileType || "").toUpperCase();
    const usesDocumentPreviewCrop = ["PDF", "DOC", "DOCX", "PPT", "PPTX", "XLS", "XLSX", "TXT"].indexOf(normalizedFileType) !== -1;
    const isPresentationPreview = ["PPT", "PPTX"].indexOf(normalizedFileType) !== -1;
    const isWordPreview = ["DOC", "DOCX", "TXT"].indexOf(normalizedFileType) !== -1;
    const shouldShowFallback = !isAudioPreview && !isExcelPreview && !isVideoPreview && (!activeThumbnailUrl || !isThumbnailLoaded);

    const markThumbnailLoaded = (): void => {
      setThumbnailLoadedByDocument((current) => (
        current[item.id]
          ? current
          : {
              ...current,
              [item.id]: true,
            }
      ));
    };

    const clearThumbnailLoaded = (): void => {
      setThumbnailLoadedByDocument((current) => {
        if (!current[item.id]) {
          return current;
        }

        const nextState = { ...current };
        delete nextState[item.id];
        return nextState;
      });
    };

    const advanceThumbnailCandidate = (): void => {
      if (activeThumbnailIndex < thumbnailCandidates.length - 1) {
        setThumbnailAttemptByDocument((current) => ({
          ...current,
          [item.id]: activeThumbnailIndex + 1,
        }));
      }
    };

    const fallbackContent = (
      <div
        className={styles.tileThumbnailFallback}
        style={{ display: shouldShowFallback ? "flex" : "none" }}
      >
        <div className={styles.fileTypeIcon}>{fallbackIcon || getFileTypeIcon(item.fileType)}</div>
        <span className={styles.fileType}>{fallbackIcon || item.fileType || "FILE"}</span>
      </div>
    );

    return (
      <div className={styles.tileThumbnailShell}>
        {isAudioPreview ? (
          <div className={styles.audioThumbnailContainer}>
            <div className={styles.audioThumbnailIcon}>
              <span />
              <span />
              <span />
              <span />
            </div>
            <span className={styles.audioThumbnailLabel}>AUDIO</span>
          </div>
        ) : isExcelPreview && activeThumbnailUrl ? (
          <img
            key={`${item.id}-${activeThumbnailIndex}`}
            src={activeThumbnailUrl}
            alt={`${item.title} thumbnail`}
            loading="eager"
            decoding="async"
            className={styles.excelThumbnailImage}
            onLoad={markThumbnailLoaded}
            onError={(event) => {
              clearThumbnailLoaded();

              if (item.fileUniqueId && activeThumbnailIndex === 0) {
                const sourceFallbackUrl = getExcelThumbnailUrl(
                  thumbnailWebUrl,
                  KM_REVIEW_HUB_DRIVE_ID,
                  item.fileUniqueId,
                  "source"
                );
                const largeFallbackUrl = getExcelThumbnailUrl(
                  thumbnailWebUrl,
                  KM_REVIEW_HUB_DRIVE_ID,
                  item.fileUniqueId,
                  "large"
                );

                if (event.currentTarget.src !== sourceFallbackUrl && event.currentTarget.src !== largeFallbackUrl) {
                  event.currentTarget.src = sourceFallbackUrl;
                  return;
                }

                if (event.currentTarget.src !== largeFallbackUrl) {
                  event.currentTarget.src = largeFallbackUrl;
                  return;
                }
              }

              advanceThumbnailCandidate();
            }}
          />
        ) : isExcelPreview ? (
          <div className={styles.excelPlaceholder}>
            <span className={styles.excelPlaceholderLabel}>XLS</span>
          </div>
        ) : isVideoPreview && !activeThumbnailUrl && videoThumbnailSource ? (
          <video
            className={styles.thumbnailImage}
            muted
            playsInline
            preload="metadata"
            onLoadedMetadata={(event) => {
              const video = event.currentTarget;
              video.currentTime = 1;
            }}
            onCanPlay={() => markThumbnailLoaded()}
            onError={() => clearThumbnailLoaded()}
          >
            <source src={videoThumbnailSource} type="video/mp4" />
          </video>
        ) : activeThumbnailUrl ? (
          <img
            key={`${item.id}-${activeThumbnailIndex}`}
            src={activeThumbnailUrl}
            alt={`${item.title} thumbnail`}
            className={`${styles.thumbnailImage} ${usesDocumentPreviewCrop ? styles.thumbnailDocumentImage : ""}`}
            loading="eager"
            decoding="async"
            style={isPresentationPreview ? {
              objectFit: "contain",
              objectPosition: "center center",
              background: "#ffffff"
            } : isWordPreview ? {
              imageRendering: "auto",
              filter: "contrast(1.08) saturate(1.03)",
              transform: "translateZ(0)"
            } : undefined}
            onLoad={markThumbnailLoaded}
            onError={(event) => {
              clearThumbnailLoaded();

              if (item.fileUniqueId && activeThumbnailIndex === 0) {
                const sourceFallbackUrl = getDriveItemThumbnailUrl(
                  thumbnailWebUrl,
                  KM_REVIEW_HUB_DRIVE_ID,
                  item.fileUniqueId,
                  "source"
                );
                const largeFallbackUrl = getDriveItemThumbnailUrl(
                  thumbnailWebUrl,
                  KM_REVIEW_HUB_DRIVE_ID,
                  item.fileUniqueId,
                  "large"
                );
                const image = event.currentTarget;

                if (image.src !== sourceFallbackUrl && image.src !== largeFallbackUrl) {
                  image.src = sourceFallbackUrl;
                  return;
                }

                if (image.src !== largeFallbackUrl) {
                  image.src = largeFallbackUrl;
                  return;
                }
              }

              advanceThumbnailCandidate();
            }}
          />
        ) : isVideoPreview ? (
          <div className={styles.videoThumbnailContainer}>
            <div className={styles.videoThumbnailPlaceholder}>
              <div className={styles.fileTypeIcon}>▶</div>
            </div>
          </div>
        ) : null}
        {fallbackContent}
      </div>
    );
  }, [
    thumbnailAttemptByDocument,
    thumbnailLoadedByDocument,
    thumbnailWebUrl,
  ]);

  const quickFormatOptions = React.useMemo(
    () => ["All", "Documents", "PDF", "PPTX", "XLSX", "Video", "Audio"],
    []
  );
  const activeFilterCount = React.useMemo(
    () => getFilterValues(effectiveFilters).filter(Boolean).length,
    [effectiveFilters]
  );

  const handleQuickFormatClick = React.useCallback((option: string) => {
    if (option === "All") {
      setFilters((current) => ({
        ...current,
        fileType: null,
      }));
      return;
    }

    if (option === "Documents") {
      setFilters((current) => ({
        ...current,
        fileType: QUICK_FILE_TYPE_VALUES.documents,
      }));
      return;
    }

    if (option === "Video") {
      setFilters((current) => ({
        ...current,
        fileType: QUICK_FILE_TYPE_VALUES.video,
      }));
      return;
    }

    if (option === "Audio") {
      setFilters((current) => ({
        ...current,
        fileType: QUICK_FILE_TYPE_VALUES.audio,
      }));
      return;
    }

    updateFilter("fileType", option);
  }, [updateFilter]);

  return (
    <>
      <div className={styles.workspace}>
        <div className={styles.sidebarContainer}>
          {isFilterPanelOpen && (
            <aside className={styles.filterSidebar}>
              <div className={styles.wrapper}>
                <div className={styles.filterSidebarHeading}>
                  <div className={styles.filterSidebarTitle}>Refine Results</div>
                </div>

                <nav className={styles.container}>
                  {filterGroups.length === 0 && !isDocumentsLoading && (hasSearch || hasAnySelectedFilter) && (
                    <div className={styles.emptyState}>No filters available for this search</div>
                  )}
                  {filterGroups.map((group) => {
                    const definition = FILTER_DEFINITIONS.find((item) => item.title === group.title);
                    if (!definition) {
                      return null;
                    }

                    const isGroupOpen = openGroup === group.title;

                    return (
                      <section key={group.title} className={styles.filterGroupSection}>
                        <button
                          type="button"
                          className={`${styles.parentBtn} ${isGroupOpen ? styles.rootBtnOpen : ""}`}
                          onClick={() => setOpenGroup((current) => (current === group.title ? null : group.title))}
                          aria-expanded={isGroupOpen}
                        >
                          <span className={styles.filterGroupHeading}>{group.title}</span>
                          <span className={styles.chevron} aria-hidden="true">
                            {isGroupOpen ? "▴" : "▾"}
                          </span>
                        </button>
                        {isGroupOpen && (
                          <div className={styles.childList}>
                            {group.children.map((child) => {
                              const activeValue = filters[definition.key];
                              const childValue = child.value || child.title;
                              const isActive =
                                definition.key === "fileType"
                                  ? activeValue?.toLowerCase() === childValue.toLowerCase()
                                  : activeValue === childValue;
                              const isHierarchyChild = group.key === "department" && Number(child.level || 0) > 0;

                              return (
                                <button
                                  key={`${group.title}-${childValue}`}
                                  type="button"
                                  className={`${styles.childBtn} ${isActive ? styles.childBtnActive : ""}`}
                                  style={isHierarchyChild ? { paddingLeft: 22 } : undefined}
                                  onClick={() => updateFilter(definition.key, childValue)}
                                >
                                  <span className={`${styles.filterOptionCheck} ${isActive ? styles.filterOptionCheckActive : ""}`}>
                                    {isActive ? "✓" : ""}
                                  </span>
                                  <span className={styles.filterOptionContent}>
                                    <span className={styles.childTitle}>
                                      {isHierarchyChild ? `› ${child.title}` : child.title}
                                    </span>
                                    <span className={styles.childCount}>{child.count}</span>
                                  </span>
                                </button>
                              );
                            })}
                          </div>
                        )}
                      </section>
                    );
                  })}
                </nav>

                {hasAnySelectedFilter && (
                  <div className={styles.filterSidebarActions}>
                    <button type="button" className={styles.clearFiltersPill} onClick={clearFilters}>
                      Clear all filters
                    </button>
                  </div>
                )}

                <div className={styles.filterSidebarFooter}>
                  <span className={styles.filterSidebarFooterText}>Search page filters</span>
                </div>
              </div>
            </aside>
          )}

          <main className={styles.main}>
            <section className={styles.results} ref={resultsSectionRef}>
              <section className={styles.quickFilterBar}>
                <div className={styles.quickFilterChips}>
                  {quickFormatOptions.map((option) => {
                    const isActive = option === "All"
                      ? !filters.fileType
                      : option === "Documents"
                        ? normalizeFilterValue(filters.fileType) === QUICK_FILE_TYPE_VALUES.documents
                        : option === "Video"
                          ? normalizeFilterValue(filters.fileType) === QUICK_FILE_TYPE_VALUES.video
                          : option === "Audio"
                            ? normalizeFilterValue(filters.fileType) === QUICK_FILE_TYPE_VALUES.audio
                        : normalizeFilterValue(filters.fileType) === normalizeFilterValue(option);

                    return (
                      <button
                        key={option}
                        type="button"
                        className={`${styles.quickFilterChip} ${isActive ? styles.quickFilterChipActive : ""}`}
                        onClick={() => handleQuickFormatClick(option)}
                      >
                        {option}
                      </button>
                    );
                  })}
                </div>

                <div className={styles.quickFilterMeta}>
                  <span className={styles.quickFilterLabel}>Sort:</span>
                  <select
                    className={styles.quickSortSelect}
                    value={documentSortOrder}
                    onChange={(event) => {
                      setDocumentSortOrder(event.target.value as DocumentSortOrder);
                      setCurrentDocumentPage(1);
                    }}
                    aria-label="Sort search results"
                  >
                    <option value="newToOld">New to old</option>
                    <option value="oldToNew">Old to new</option>
                  </select>
                  <button
                    type="button"
                    className={styles.quickClearButton}
                    onClick={clearFilters}
                    disabled={!canClearSearchResults}
                  >
                    Clear all
                  </button>
                </div>
              </section>

              {activeFilterCount > 0 && (
                <div className={styles.resultsSummaryBar}>
                  <span className={styles.resultsSummaryBadge}>
                    {activeFilterCount} filter{activeFilterCount === 1 ? "" : "s"} applied
                  </span>
                </div>
              )}

              {renderSuggestionsSection && shouldShowSuggestions && (
                <section className={styles.suggestionsSection}>
                  <div className={styles.suggestionsHeader}>Suggestions</div>
                  <div className={styles.suggestionsList}>
                    {suggestions.map((suggestion) => (
                      <button
                        key={`${suggestion.source}-${suggestion.value}`}
                        type="button"
                        className={styles.suggestionButton}
                        onClick={() => submitSuggestion(suggestion.value)}
                      >
                        <span className={styles.suggestionIcon}>
                          {suggestion.source === "query" ? "↺" : "⌕"}
                        </span>
                        <span className={styles.suggestionText}>{suggestion.value}</span>
                        <span className={styles.suggestionMeta}>
                          {suggestion.source === "query" ? `${suggestion.count || 0} searches` : suggestion.meta || "Title match"}
                        </span>
                      </button>
                    ))}
                  </div>
                </section>
              )}

              {(isDocumentsLoading || (isSearchLoading && filteredItems.length === 0)) && activeTab === "documents" && (
                <>
                  <div className={styles.paginationHeader}>
                    <span>Loading matches...</span>
                  </div>
                  <div className={styles.loadingState}>Loading documents...</div>
                </>
              )}

              {/* 🚩 UAT FIX: Removing technical error message display as requested */}
              {/* {error && <p style={{ color: "red" }}>Error: {error}</p>} */}

              {!isDocumentsLoading && !error && !isResultsCleared && (hasSearch || hasAnySelectedFilter) && (
                activeTab === "documents" ? (
                  sortedFilteredItems.length === 0 ? (
                    <div className={styles.emptyState}>
                      <div className={styles.emptyStateIcon}>🔍</div>
                      <div className={styles.emptyStateTitle}>No results matched your search</div>
                      <div className={styles.emptyStateDescription}>
                        We couldn't find any documents matching your current keywords or filters. 
                        Try broadening your search or clearing the filters to discover more content.
                      </div>
                      <div className={styles.emptyStateActions}>
                        {hasAnySelectedFilter && (
                          <button
                            type="button"
                            className={`${styles.emptyStateAction} ${styles.emptyStateActionPrimary}`}
                            onClick={clearFilters}
                          >
                            Clear all filters
                          </button>
                        )}
                        {suggestions.length > 0 && (
                          <button
                            type="button"
                            className={styles.emptyStateAction}
                            onClick={() => submitSuggestion(suggestions[0].value)}
                          >
                            Try "{suggestions[0].value}"
                          </button>
                        )}
                      </div>
                    </div>
                  ) : (
                    <div className={styles.documentResultsPanel}>
                      <div className={styles.searchDisplayControls}>
                        <div className={styles.viewModeToggles}>
                          <button
                            type="button"
                            className={`${styles.viewModeBtn} ${viewMode === "list" ? styles.viewModeBtnActive : ""}`}
                            onClick={() => setViewMode("list")}
                            title="List View"
                            aria-label="List View"
                          >
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01" />
                            </svg>
                          </button>
                          <button
                            type="button"
                            className={`${styles.viewModeBtn} ${viewMode === "grid" ? styles.viewModeBtnActive : ""}`}
                            onClick={() => setViewMode("grid")}
                            title="Grid View"
                            aria-label="Grid View"
                          >
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                              <rect x="3" y="3" width="7" height="7" />
                              <rect x="14" y="3" width="7" height="7" />
                              <rect x="14" y="14" width="7" height="7" />
                              <rect x="3" y="14" width="7" height="7" />
                            </svg>
                          </button>
                        </div>
                      </div>
                      <div className={styles.paginationHeader}>
                        <span>
                          {sortedFilteredItems.length > 0
                            ? `${paginatedStartIndex}-${paginatedEndIndex} of ${sortedFilteredItems.length} matches found`
                            : "0 matches found"}
                        </span>
                      </div>

                      <div className={viewMode === "grid" ? styles.documentGrid : styles.documentList}>
                        {paginatedDocumentItems.map((item) => (
                          viewMode === "grid" ? (
                          <div key={item.id} className={styles.documentGridCard}>
                            <div className={styles.gridThumbnailSection}>
                              {renderSearchGridThumbnail(item)}
                            </div>

                            <div className={styles.gridCardBody}>
                              <div className={styles.gridCardMeta}>
                                {renderAuthorLines(item.contributor || "Internal", styles.authorStack, styles.authorChip)}
                                <span className={styles.sourceType}>{item.fileType || "FILE"}</span>
                              </div>
                              <h3 className={styles.gridDocumentTitle}>{item.title}</h3>
                              <p className={styles.searchResultDescription}>
                                {buildResultPreview(
                                  item,
                                  activeSearchText,
                                  getCachedDocumentContent(item)
                                )}
                              </p>
                              <div className={styles.gridCardFooter}>
                                <span>{item.updated || "-"}</span>
                                <span className={`${styles.statusBadge} ${getStatusClassName(item.status)}`}>
                                  {getDisplayStatus(item.status) || "Active"}
                                </span>
                              </div>
                              <div className={styles.gridActionButtons}>
                                <button type="button" className={styles.viewButton} onClick={() => handleView(item)}>
                                  View
                                </button>
                                {!isLearner && (
                                  <button
                                    type="button"
                                    className={styles.downloadButton}
                                    disabled={!!downloadingDocumentIds[item.id]}
                                    onClick={() => handleDownload(item)}
                                    aria-label={`Download ${item.title}`}
                                  >
                                    {downloadingDocumentIds[item.id] ? "Downloading..." : "Download"}
                                  </button>
                                )}
                              </div>
                            </div>
                          </div>
                          ) : (
                            <article key={item.id} className={styles.documentRow}>
                              <div className={styles.listContent}>
                                <div className={styles.searchResultContent}>
                                  <h3 className={styles.searchResultTitle}>{item.title}</h3>
                                  <p className={styles.searchResultDescription}>
                                    {buildResultPreview(
                                      item,
                                      activeSearchText,
                                      getCachedDocumentContent(item)
                                    )}
                                  </p>
                                  <div className={styles.searchResultStatusLine}>
                                    <span className={styles.searchResultStatusLabel}>Status :</span>
                                    <span className={`${styles.statusBadge} ${getStatusClassName(item.status)}`}>
                                      {getDisplayStatus(item.status) || "Active"}
                                    </span>
                                  </div>
                                </div>
                                <div className={styles.searchResultDivider} aria-hidden="true" />
                                <div className={styles.searchResultMeta}>
                                  {renderAuthorLines(item.contributor || "Internal", styles.authorStack, styles.searchResultMetaStrong)}
                                  <span className={styles.sourceDate}>{item.updated || "-"}</span>
                                </div>
                                <div className={styles.searchResultActions}>
                                  <button type="button" className={styles.searchResultViewButton} onClick={() => handleView(item)}>
                                    View
                                  </button>
                                  {!isLearner && (
                                    <button
                                      type="button"
                                      className={styles.searchResultDownloadButton}
                                      disabled={!!downloadingDocumentIds[item.id]}
                                      onClick={() => handleDownload(item)}
                                      aria-label={`Download ${item.title}`}
                                    >
                                      {downloadingDocumentIds[item.id] ? "Downloading..." : "Download"}
                                    </button>
                                  )}
                                </div>
                              </div>
                            </article>
                          )
                        ))}
                      </div>

                      {hasMoreResults && (
                        <div className={styles.loadMoreContainer}>
                          <button
                            type="button"
                            className={styles.loadMoreButton}
                            onClick={() => setCurrentStartRow((prev) => prev + backendPageSize)}
                            disabled={isDocumentsLoading}
                          >
                            {isDocumentsLoading ? "Loading more..." : "Load more documents"}
                          </button>
                        </div>
                      )}

                      {totalDocumentPages > 1 && (
                        <div className={styles.paginationFooter}>
                          <button
                            type="button"
                            className={styles.pageArrowBtn}
                            onClick={() => setCurrentDocumentPage((page) => Math.max(1, page - 1))}
                            disabled={currentDocumentPage === 1}
                          >
                            Previous
                          </button>
                          <div className={styles.pageNumbers}>
                            {documentPaginationItems.map((item) => (
                              typeof item === "number" ? (
                                <button
                                  key={item}
                                  type="button"
                                  className={`${styles.pageNumberBtn} ${item === currentDocumentPage ? styles.pageNumberBtnActive : ""}`}
                                  onClick={() => setCurrentDocumentPage(item)}
                                  aria-current={item === currentDocumentPage ? "page" : undefined}
                                >
                                  {item}
                                </button>
                              ) : (
                                <span key={item} className={styles.pageEllipsis}>...</span>
                              )
                            ))}
                          </div>
                          <button
                            type="button"
                            className={styles.pageArrowBtn}
                            onClick={() => setCurrentDocumentPage((page) => Math.min(totalDocumentPages, page + 1))}
                            disabled={currentDocumentPage === totalDocumentPages}
                          >
                            Next
                          </button>
                        </div>
                      )}
                    </div>
                  )
                ) : (
                  <div className={styles.scrollContainer}>
                    <div className={styles.list}>
                      {filteredItems.length === 0 ? (
                        <div className={styles.emptyState}>No results found for the current query and filters.</div>
                      ) : (
                        filteredItems.map((item) => (
                          <article key={item.id} className={styles.card}>
                            <h4 className={styles.cardTitle}>{item.title}</h4>
                            <div className={styles.cardRow}>
                              <span className={styles.cardLabel}>Author</span>
                              <span className={styles.cardValue}>{item.contributor}</span>
                            </div>
                            <div className={styles.cardRow}>
                              <span className={styles.cardLabel}>Updated</span>
                              <span className={styles.cardValue}>{item.updated}</span>
                            </div>
                            <div className={styles.cardRow}>
                              <span className={styles.cardLabel}>Description</span>
                              <span className={styles.cardValue}>{item.description}</span>
                            </div>
                            <div className={styles.cardRow}>
                              <span className={styles.cardLabel}>File Name</span>
                              <span className={styles.cardValue}>{item.fileName}</span>
                            </div>
                          </article>
                        ))
                      )}
                    </div>
                  </div>
                )
              )}
            </section>
          </main>
        </div>
      </div>

      {showAllDocumentResults &&
        activeTab === "documents" &&
        typeof document !== "undefined" &&
        createPortal(
          <div className={styles.allResultsModal} data-testid="all-results-backdrop" onClick={(event) => {
            if ((event.target as HTMLElement).getAttribute("data-testid") === "all-results-backdrop") {
              handleCloseAllDocumentResults();
            }
          }}>
            <div className={styles.allResultsPanel}>
              <div className={styles.allResultsHeader}>
                <div>
                  <h3 className={styles.allResultsTitle}>All matching documents</h3>
                  <p className={styles.allResultsSubtitle}>{filteredItems.length} result{filteredItems.length === 1 ? "" : "s"} available</p>
                </div>
                <button type="button" className={styles.allResultsClose} onClick={handleCloseAllDocumentResults}>
                  Close
                </button>
              </div>
              <div className={styles.allResultsGrid}>
                {filteredItems.map((item) => (
                  <div key={item.id} className={styles.tile}>
                    <div className={styles.tileThumbnailSection}>
                      {renderSearchGridThumbnail(item, item.docIcon)}
                      <div className={styles.tileHeader}>
                        <span className={styles.tileDate}>{item.updated}</span>
                      </div>
                    </div>

                    <div className={styles.tileContent}>
                      <h3 className={styles.tileTitle}>{item.title}</h3>
                      <div className={styles.tileMeta}>
                        <span className={styles.metaLabel}>Format:</span>
                        <span className={styles.metaValue}>{item.fileType || "FILE"}</span>
                      </div>
                      <div className={styles.tileMeta}>
                        <span className={styles.metaLabel}>Author:</span>
                        <span className={styles.metaValue}>{item.contributor}</span>
                      </div>
                      <div className={styles.tileMeta}>
                        <span className={styles.metaLabel}>Updated:</span>
                        <span className={styles.metaValue}>{item.updated}</span>
                      </div>
                      <div className={styles.tileMeta}>
                        <span className={styles.metaLabel}>File Name:</span>
                        <span className={styles.metaValue}>{item.fileName || "Untitled"}</span>
                      </div>
                      <div className={styles.tileMeta}>
                        <span className={styles.metaLabel}>BU/SL:</span>
                        <span className={styles.metaValue}>{item.businessUnit || "Unspecified"}</span>
                      </div>
                      <p className={styles.tileAbstract}>
                        {buildResultPreview(
                          item,
                          activeSearchText,
                          getCachedDocumentContent(item)
                        )}
                      </p>
                    </div>

                    <div className={styles.tileFooter}>
                      <div className={styles.tileActions}>
                        <button type="button" className={styles.viewButton} onClick={() => handleView(item)}>
                          View
                        </button>
                        {!isLearner && (
                          <button
                            type="button"
                            className={styles.downloadButton}
                            disabled={!!downloadingDocumentIds[item.id]}
                            onClick={() => handleDownload(item)}
                            aria-label={downloadingDocumentIds[item.id] ? `Downloading ${item.title}` : `Download ${item.title}`}
                          >
                            <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                              <polyline points="7 10 12 15 17 10" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                              <line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </button>
                        )}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>,
          document.body
        )}

      {selectedDocumentId &&
        context &&
        typeof document !== "undefined" &&
        createPortal(
          <div
            className={styles.detailModal}
            data-testid="modal-backdrop"
            onClick={(event) => {
              if ((event.target as HTMLElement).getAttribute("data-testid") === "modal-backdrop") {
                handleCloseDetail();
              }
            }}
          >
            <DocumentDetailPage
              context={context}
              documentId={selectedDocumentId}
              onClose={handleCloseDetail}
              backTo="library"
              onBackToLibrary={handleCloseDetail}
            />
          </div>,
          document.body
        )}
    </>
  );
};

export default GenericSearchDropdown;
