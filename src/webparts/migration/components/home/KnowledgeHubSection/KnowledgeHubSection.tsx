import * as React from "react";
import { createPortal } from "react-dom";
import { IKnowledgeHubSectionProps } from "./IKnowledgeHubSectionProps";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { DocumentDetailPage } from "../../../pages/DocumentDetailPage/DocumentDetailPage";
import { ViewAllDocumentsPage } from "../../../pages/ViewAllDocumentsPage/ViewAllDocumentsPage";
import styles from "./KnowledgeHubSection.module.scss";
import { configurePermalinkService, syncPrettyUrlsForItems, pushPageUrl, NAV_PATHS, openAppPageInNewTab } from "../../../services/permalinkService";
import { subscribeToDocumentDataChanged } from "../../../services/documentChangeEvents";
import { SharePointSearchService, getSearchServiceInstance } from "../../../services/SharePointSearchService";
import { IChatMessage as ExternalChatMessage } from "../../../models/ChatMessage";
import { ISearchResult } from "../../../models/SearchResult";
import { KnowledgeSearchApiClient } from "../../../services/KnowledgeSearchApiClient";
import { downloadSharePointFile } from '../../../utils/fileDownload';
import {
  buildKMDataHubItemQuery,
  fetchKMDataHubReadFieldMap,
  resolveDocumentAuthor,
  resolveDocumentPublishedValue,
  splitDocumentAuthorDisplay
} from '../../../utils/documentMetadata';
import { CACHE_KEYS, COLUMN_NAMES, LIBRARY_NAMES, LIST_NAMES, TTL_MS } from '../../../config/appConfig';

/*Chatbot Section*/
import chatbotIcon from '../../../assets/chatbot-icon.png';

const isDebugLoggingEnabled = (): boolean => {
  try {
    return typeof window !== 'undefined' &&
      window.localStorage?.getItem('IKNOWLEDGE_DEBUG_LOGS') === 'true';
  } catch {
    return false;
  }
};

const debugLog = (...args: unknown[]): void => {
  if (isDebugLoggingEnabled()) {
    console.log(...args);
  }
};

interface DocumentItem {
  id: number;
  name: string;
  abstract: string;
  fileType: string;
  author: string;
  date: string;
  views?: number;
  comments?: number;
  likes?: number;
  downloads?: number;
  fileSize: string;
  serverRelativeUrl: string;
  fileRef: string;
  fileUniqueId?: string;
  previewImageUrls?: string[];
  businessUnit?: string;
  documentType?: string;
  fileName?: string;
  contentRefreshDate?: string;
  follow?: boolean;
  share?: boolean;
  bookmark?: boolean;
  sensitiveTerms?: string[];
  reviewerComments?: string;
  projectId?: string;
  versionFileName?: string;
  versionFileType?: string;
  modifiedBy?: {
    title: string;
    email: string;
    id: number;
  };
  docIcon?: string;
}

interface ChatMessage {
  sender: "user" | "bot";
  text: string;
  // optional link metadata: when present, UI will render an anchor
  title?: string;
  url?: string;
  documents?: DocumentItem[];
  // RAG additions
  citations?: ExternalChatMessage['citations'];
  historyCitations?: ExternalChatMessage['citations'];
  timestamp?: Date;
}

interface FocusedDocument {
  id: string;
  title: string;
  url?: string;
}

const CHAT_MEMORY_MAX_CHARS = 1400;
const CHAT_INPUT_MAX_CHARS = 2500;

const compactChatMemoryText = (value: string, maxLength: number): string => {
  const text = String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();

  return text.length <= maxLength ? text : text.slice(0, maxLength).trim();
};

/**
 * Keeps a compact, browser-only memory for the current chat window. It is not a
 * persisted transcript and it is not treated as factual evidence by the backend.
 */
const buildNextChatConversationSummary = (
  previousSummary: string,
  userQuestion: string,
  botMessage: ChatMessage
): string => {
  const sourceTitles = (botMessage.historyCitations || botMessage.citations || [])
    .map((source) => source.title || '')
    .filter(Boolean)
    .filter((title, index, all) => all.findIndex((candidate) => candidate.toLowerCase() === title.toLowerCase()) === index)
    .slice(0, 4);

  const turnSummary = [
    `User asked: ${compactChatMemoryText(userQuestion, 260)}`,
    `Assistant answered: ${compactChatMemoryText(botMessage.text || '', 520)}`,
    sourceTitles.length > 0 ? `Sources referenced: ${sourceTitles.join('; ')}` : ''
  ].filter(Boolean).join('\n');

  const combined = [previousSummary, turnSummary].filter(Boolean).join('\n\n');
  return combined.length <= CHAT_MEMORY_MAX_CHARS
    ? combined
    : combined.slice(combined.length - CHAT_MEMORY_MAX_CHARS).replace(/^[\s\S]*?(?=User asked:|$)/, '').trim();
};

const METRICS_CACHE_KEY = CACHE_KEYS.metricsKnowledgeHub;
const METRICS_TTL_MS = TTL_MS.metricsCache;
let fieldMapModuleCache: any = null;

const getItemFieldValue = (item: any, fieldName?: string, fallbacks: string[] = []): any => {
  const candidates = [fieldName, ...fallbacks].filter(Boolean) as string[];

  for (const candidate of candidates) {
    if (item && item[candidate] !== undefined && item[candidate] !== null) {
      return item[candidate];
    }
  }

  return undefined;
};

const normalizeTextField = (value: any): string => {
  if (Array.isArray(value)) {
    return value
      .map((entry) => String(entry?.Label || entry?.Title || entry?.Value || entry || '').trim())
      .filter(Boolean)
      .join(', ');
  }

  if (value && typeof value === 'object') {
    return String(value.Label || value.Title || value.Value || '').trim();
  }

  return String(value ?? '').trim();
};

const normalizeDescriptionText = (value: any): string => {
  const text = normalizeTextField(value);
  return !text || text === '-' ? '' : text;
};

const normalizeNumberField = (value: any): number => {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
};

const renderStatIcon = (type: 'views' | 'comments' | 'likes' | 'downloads'): React.ReactElement => {
  if (type === 'views') {
    return (
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        <circle cx="12" cy="12" r="3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    );
  }

  if (type === 'comments') {
    return (
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M21 11.5a8.5 8.5 0 0 1-8.5 8.5c-1.2 0-2.4-.25-3.47-.72L3 21l1.9-4.5A8.47 8.47 0 0 1 4 12a8.5 8.5 0 1 1 17 0Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    );
  }

  if (type === 'downloads') {
    return (
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M12 3v11" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        <path d="M8 10l4 4 4-4" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
        <path d="M4 17v1a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-1" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      </svg>
    );
  }

  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M12 21s-6.7-4.35-9.33-8.08C.7 10.13 1.3 6.2 4.6 4.42c2.3-1.24 4.9-.55 6.4 1.1 1.5-1.65 4.1-2.34 6.4-1.1 3.3 1.78 3.9 5.71 1.93 8.5C18.7 16.65 12 21 12 21Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
};
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
const getPreviewAccent = (fileType: string): string => {
  const type = fileType.toLowerCase();

  if (type === 'pptx' || type === 'ppt') {
    return 'linear-gradient(135deg, #f8fafc 0%, #dbeafe 45%, #bfdbfe 100%)';
  }

  if (type === 'pdf') {
    return 'linear-gradient(135deg, #f8fafc 0%, #e5e7eb 45%, #d1d5db 100%)';
  }

  if (type === 'docx' || type === 'doc') {
    return 'linear-gradient(135deg, #f8fafc 0%, #e0f2fe 45%, #bae6fd 100%)';
  }

  return 'linear-gradient(135deg, #f8fafc 0%, #ede9fe 45%, #ddd6fe 100%)';
};

const getPreviewLabel = (doc: DocumentItem): string => {
  const normalizedTitle = doc.name.toLowerCase();

  if (normalizedTitle.includes('battlecard')) {
    return 'Battlecard';
  }

  if (normalizedTitle.includes('regulatory')) {
    return 'Series of Contributions';
  }

  if (normalizedTitle.includes('submission') || normalizedTitle.includes('publishing')) {
    return 'Knowledge Article';
  }

  return `${doc.fileType || 'Document'} Article`;
};

const getAuthorInitials = (author: string): string => {
  if (!author || author === 'Internal') {
    return 'IN';
  }

  const parts = author.split(' ').filter(Boolean);
  return parts.slice(0, 2).map((part) => part[0]?.toUpperCase()).join('') || 'IN';
};

const buildPreviewImageUrl = (webUrl: string, serverRelativeUrl: string): string => {
  const fileUrl = `${webUrl}${serverRelativeUrl}`;
  return `${webUrl}/_layouts/15/getpreview.ashx?path=${encodeURIComponent(fileUrl)}&resolution=2`;
};

const buildGuidPreviewImageUrl = (
  webUrl: string,
  siteId: string,
  webId: string,
  fileUniqueId: string
): string => `${webUrl}/_layouts/15/getpreview.ashx?guidSite=${encodeURIComponent(siteId)}&guidWeb=${encodeURIComponent(webId)}&guidFile=${encodeURIComponent(fileUniqueId)}&resolution=2`;

const buildWopiFramePreviewUrl = (webUrl: string, serverRelativeUrl: string): string =>
  `${webUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodeURIComponent(serverRelativeUrl)}&action=imagepreview`;

const buildPreviewCandidates = (
  webUrl: string,
  serverRelativeUrl: string,
  siteId: string,
  webId: string,
  fileUniqueId?: string
): string[] => {
  const candidates: string[] = [];

  if (fileUniqueId) {
    candidates.push(buildGuidPreviewImageUrl(webUrl, siteId, webId, fileUniqueId));
  }

  candidates.push(buildPreviewImageUrl(webUrl, serverRelativeUrl));
  candidates.push(buildWopiFramePreviewUrl(webUrl, serverRelativeUrl));

  return candidates;
};

const SEARCH_SCAN_SUPPORTED_TYPES = ['PDF', 'DOCX', 'PPTX', 'XLSX', 'XLSM', 'XLS', 'XLSB', 'XLTX', 'XLTM', 'CSV', 'TXT', 'MHTML', 'MHT', 'SVG'];
const SEARCH_CONTENT_SCAN_LIMIT = 25;
const SEARCH_CONTENT_SCAN_BATCH_SIZE = 5;

const normalizeSearchText = (value: string): string =>
  (value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const escapeRegExp = (value: string): string => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

const extractKeywordTerms = (query: string): string[] => {
  const stopWords = new Set([
    'what', 'is', 'are', 'the', 'a', 'an', 'how', 'why', 'when', 'where', 'who', 'which',
    'this', 'that', 'these', 'those', 'do', 'does', 'did', 'can', 'could', 'will', 'would',
    'should', 'may', 'might', 'must', 'please', 'tell', 'me', 'about', 'find', 'show'
  ]);
  const terms: string[] = [];
  const quotedPhraseRegex = /"([^"]+)"/g;
  let match: RegExpExecArray | null;

  while ((match = quotedPhraseRegex.exec(query)) !== null) {
    const phrase = normalizeSearchText(match[1]);
    if (phrase.length > 1) {
      terms.push(phrase);
    }
  }

  const unquotedQuery = query
    .replace(quotedPhraseRegex, ' ')
    .replace(/\b(AND|OR|NOT)\b/gi, ' ');

  normalizeSearchText(unquotedQuery)
    .split(/\s+/)
    .filter(term => term.length > 1 && !stopWords.has(term))
    .forEach(term => terms.push(term));

  const normalizedFullQuery = normalizeSearchText(unquotedQuery);
  if (normalizedFullQuery.length > 2 && !terms.includes(normalizedFullQuery)) {
    terms.push(normalizedFullQuery);
  }

  return Array.from(new Set(terms));
};

const countSearchTermOccurrences = (text: string, terms: string[]): number => {
  if (!text || terms.length === 0) {
    return 0;
  }

  const normalizedText = normalizeSearchText(text);
  return terms.reduce((total, term) => {
    const normalizedTerm = normalizeSearchText(term);
    if (!normalizedTerm) {
      return total;
    }

    const regex = new RegExp(`\\b${escapeRegExp(normalizedTerm)}\\b`, 'g');
    return total + (normalizedText.match(regex) || []).length;
  }, 0);
};

const toServerRelativePath = (pathOrUrl: string): string => {
  if (!pathOrUrl) {
    return '';
  }

  try {
    if (/^https?:\/\//i.test(pathOrUrl)) {
      return new URL(pathOrUrl).pathname;
    }
  } catch (e) {
    console.warn('Unable to parse SharePoint file URL:', pathOrUrl, e);
  }

  return pathOrUrl.startsWith('/') ? pathOrUrl : `/${pathOrUrl}`;
};

/* Message Content Component - Converts URLs to clickable links */
const MessageContent: React.FC<{
  text: string;
  styles?: any;
  citations?: ExternalChatMessage['citations'];
  onCitationClick?: (citationIndex: number) => void;
}> = ({ text, styles, citations, onCitationClick }) => {
  const citationBySourceNumber = React.useMemo(() => {
    const byNumber = new Map<number, { citation: ISearchResult; visibleIndex: number }>();

    (citations || []).forEach((citation, visibleIndex) => {
      const explicitNumber = Number((citation as ISearchResult).rank);
      const sourceNumber = Number.isFinite(explicitNumber) && explicitNumber > 0
        ? explicitNumber
        : visibleIndex + 1;

      if (!byNumber.has(sourceNumber)) {
        byNumber.set(sourceNumber, { citation: citation as ISearchResult, visibleIndex });
      }
    });

    return byNumber;
  }, [citations]);

  // Regular expression to match URLs (with capturing group for split)
  const urlRegex = /(https?:\/\/[^\s)]+)/g;

  let processedText = text;
  const markdownLinks: Array<{ fullMatch: string; text: string; url: string }> = [];

  // Manual parsing to handle URLs with parentheses correctly
  let textIndex = 0;
  while (textIndex < text.length) {
    const linkStart = text.indexOf('[', textIndex);
    if (linkStart === -1) break;

    const textEnd = text.indexOf(']', linkStart + 1);
    if (textEnd === -1) break;

    const urlStart = text.indexOf('(', textEnd + 1);
    if (urlStart === -1 || urlStart !== textEnd + 1) {
      textIndex = textEnd + 1;
      continue;
    }

    let urlEnd = urlStart + 1;
    let foundUrlEnd = false;

    while (urlEnd < text.length && !foundUrlEnd) {
      if (text[urlEnd] === ')') {
        const nextChar = urlEnd + 1 < text.length ? text[urlEnd + 1] : '';
        const nextTwoChars = urlEnd + 2 < text.length ? text.substring(urlEnd + 1, urlEnd + 3) : '';

        // End of URL if followed by: end of string, space, newline, or markdown formatting
        if (nextChar === '' ||
          nextChar === ' ' ||
          nextChar === '\n' ||
          nextChar === '\r' ||
          nextTwoChars === '**' || // End of markdown bold
          (urlEnd + 1 < text.length && text[urlEnd + 1] === ' ')) {
          foundUrlEnd = true;
          break;
        }
        // If it's part of the URL (like in encoded paths), continue
        urlEnd++;
      } else if (text[urlEnd] === '\n' || text[urlEnd] === '\r') {
        // URL shouldn't contain newlines, so stop here
        break;
      } else {
        urlEnd++;
      }
    }

    if (foundUrlEnd) {
      const linkText = text.substring(linkStart + 1, textEnd);
      const linkUrl = text.substring(urlStart + 1, urlEnd);

      markdownLinks.push({
        fullMatch: text.substring(linkStart, urlEnd + 1),
        text: linkText,
        url: linkUrl
      });

      textIndex = urlEnd + 1;
    } else {
      textIndex = textEnd + 1;
    }
  }

  // Replace markdown links with placeholders
  markdownLinks.forEach((link, index) => {
    processedText = processedText.replace(link.fullMatch, `__MARKDOWN_LINK_${index}__`);
  });

  // Split text by markdown link placeholders to preserve text before/after links
  const placeholderRegex = /(__MARKDOWN_LINK_\d+__)/g;
  const parts = processedText.split(placeholderRegex);

  return (
    <div style={{ whiteSpace: 'pre-wrap' }}>
      {parts.map((part, index) => {
        // Check if this part is a markdown link placeholder
        const markdownMatch = part.match(/__MARKDOWN_LINK_(\d+)__/);
        if (markdownMatch) {
          const linkIndex = parseInt(markdownMatch[1]);
          const link = markdownLinks[linkIndex];
          if (link) {
            return (
              <a
                key={index}
                href={link.url}
                target="_blank"
                rel="noopener noreferrer"
                className={styles?.botLink || ''}
                style={!styles?.botLink ? {
                  color: '#1f72c9',
                  textDecoration: 'underline',
                  wordBreak: 'break-all',
                  cursor: 'pointer',
                  fontWeight: '500',
                  pointerEvents: 'auto',
                  display: 'inline-block',
                  margin: '2px 0'
                } : {
                  display: 'inline-flex',
                  margin: '4px 0'
                }}
                onClick={(e) => {
                  e.stopPropagation();
                }}
              >
                {link.text}
              </a>
            );
          }
        }

        // Check if this part is a URL
        if (part.startsWith('http://') || part.startsWith('https://')) {
          return (
            <a
              key={index}
              href={part}
              target="_blank"
              rel="noopener noreferrer"
              style={{
                color: '#1f72c9',
                textDecoration: 'underline',
                wordBreak: 'break-all',
                cursor: 'pointer',
                fontWeight: '500',
                pointerEvents: 'auto'
              }}
              onClick={(e) => e.stopPropagation()}
            >
              {part}
            </a>
          );
        }

        // Regular text with lightweight markdown bold support.
        if (part.trim().length > 0 || part.includes('\n')) {
          const boldParts = part.split(/(\*\*[^*]+\*\*)/g);
          return (
            <span key={index}>
              {boldParts.map((boldPart, boldIndex) => {
                const boldMatch = boldPart.match(/^\*\*([^*]+)\*\*$/);
                const textPart = boldMatch ? boldMatch[1] : boldPart.replace(/\*/g, '');
                const citationParts = textPart.split(/(\[[1-9]\d*\])/g);
                const rendered = citationParts.map((citationPart, citationIndex) => {
                  const citationMatch = citationPart.match(/^\[([1-9]\d*)\]$/);
                  if (citationMatch) {
                    const citationNumber = Number(citationMatch[1]);
                    const citationEntry = citationBySourceNumber.get(citationNumber);
                    if (citationEntry?.citation && onCitationClick) {
                      const citation = citationEntry.citation;
                      return (
                        <button
                          key={`${boldIndex}-cite-${citationIndex}`}
                          type="button"
                          aria-label={`Open reference ${citationNumber}: ${citation.title || 'source'}`}
                          title={citation.title || `Reference ${citationNumber}`}
                          onClick={(event) => {
                            event.stopPropagation();
                            onCitationClick(citationEntry.visibleIndex);
                          }}
                          style={{
                            display: 'inline-flex',
                            width: 18,
                            height: 18,
                            padding: 0,
                            margin: '0 2px',
                            border: 0,
                            background: 'transparent',
                            color: '#5a35d6',
                            verticalAlign: 'text-bottom',
                            cursor: 'pointer'
                          }}
                        >
                          <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                            <path d="M10 13a5 5 0 0 0 7.07 0l2.12-2.12a5 5 0 0 0-7.07-7.07L10.9 5.03" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            <path d="M14 11a5 5 0 0 0-7.07 0L4.81 13.12a5 5 0 0 0 7.07 7.07l1.22-1.22" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                        </button>
                      );
                    }

                    return null;
                  }

                  return <React.Fragment key={`${boldIndex}-text-${citationIndex}`}>{citationPart}</React.Fragment>;
                });

                return boldMatch
                  ? <strong key={boldIndex}>{rendered}</strong>
                  : <React.Fragment key={boldIndex}>{rendered}</React.Fragment>;
              })}
            </span>
          );
        }
        return null;
      })}
    </div>
  );
};

const TypedMessageContent: React.FC<{
  text: string;
  styles?: any;
  enableTyping?: boolean;
  onTypingComplete?: () => void;
  citations?: ExternalChatMessage['citations'];
  onCitationClick?: (citationIndex: number) => void;
}> = ({ text, styles, enableTyping = false, onTypingComplete, citations, onCitationClick }) => {
  const [visibleText, setVisibleText] = React.useState<string>(enableTyping ? '' : text);
  const onTypingCompleteRef = React.useRef(onTypingComplete);

  React.useEffect(() => {
    onTypingCompleteRef.current = onTypingComplete;
  }, [onTypingComplete]);

  React.useEffect(() => {
    let didNotifyComplete = false;
    const notifyComplete = (): void => {
      if (didNotifyComplete) {
        return;
      }

      didNotifyComplete = true;
      onTypingCompleteRef.current?.();
    };

    if (!enableTyping) {
      setVisibleText(text);
      notifyComplete();
      return;
    }

    if (!text) {
      setVisibleText('');
      notifyComplete();
      return;
    }

    let currentLength = 0;
    setVisibleText('');

    const intervalId = window.setInterval(() => {
      currentLength = Math.min(text.length, currentLength + 4);
      setVisibleText(text.slice(0, currentLength));

      if (currentLength >= text.length) {
        window.clearInterval(intervalId);
        notifyComplete();
      }
    }, 16);

    return () => window.clearInterval(intervalId);
  }, [enableTyping, text]);

  return <MessageContent text={visibleText} styles={styles} citations={citations} onCitationClick={onCitationClick} />;
};

const buildReferenceLeadIn = (): string => {
  return 'Here are the Knowledge Hub references that support this answer:';
};

const normalizeCitationText = (value: string): string =>
  (value || '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const getCitationMentionIndex = (answerText: string, citationTitle: string): number => {
  const normalizedAnswer = normalizeCitationText(answerText);
  const normalizedTitle = normalizeCitationText(citationTitle);

  if (!normalizedAnswer || !normalizedTitle) {
    return Number.MAX_SAFE_INTEGER;
  }

  const exactIndex = normalizedAnswer.indexOf(normalizedTitle);
  if (exactIndex >= 0) {
    return exactIndex;
  }

  const titleWithoutExtension = normalizedTitle.replace(/\s+(pdf|docx|pptx|xlsx|doc|ppt|xls)$/i, '').trim();
  if (titleWithoutExtension && titleWithoutExtension !== normalizedTitle) {
    const extensionlessIndex = normalizedAnswer.indexOf(titleWithoutExtension);
    if (extensionlessIndex >= 0) {
      return extensionlessIndex;
    }
  }

  return Number.MAX_SAFE_INTEGER;
};

const orderCitationsByAnswerMentions = (
  citations: ExternalChatMessage['citations'],
  answerText: string
): ExternalChatMessage['citations'] => {
  return [...(citations || [])].sort((left, right) => {
    const leftIndex = getCitationMentionIndex(answerText, left.title || '');
    const rightIndex = getCitationMentionIndex(answerText, right.title || '');
    return leftIndex - rightIndex;
  });
};
/* Document Tile Section types*/
export const KnowledgeHubSection: React.FunctionComponent<IKnowledgeHubSectionProps> = (props) => {
  const hideDownloadActions = props.isLearner === true;
  const showPublishedSection = props.showPublishedSection !== false;

  React.useEffect(() => {
    if (!props.context) {
      return;
    }

    configurePermalinkService(props.context.spHttpClient, props.context.pageContext.web.absoluteUrl);
  }, [props.context]);
  const [documents, setDocuments] = React.useState<DocumentItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [documentsRefreshKey, setDocumentsRefreshKey] = React.useState<number>(0);
  const [selectedDocumentId, setSelectedDocumentId] = React.useState<number | null>(null);
  const [selectedListTitle, setSelectedListTitle] = React.useState<string | null>(null);
  const [selectedListId, setSelectedListId] = React.useState<string>('');
  const [showViewAll, setShowViewAll] = React.useState<boolean>(false);

  // Chatbot state
  const [chatMessages, setChatMessages] = React.useState<ChatMessage[]>([]);
  const [chatInput, setChatInput] = React.useState<string>("");
  const [chatConversationSummary, setChatConversationSummary] = React.useState<string>("");
  const [chatConversationId, setChatConversationId] = React.useState<string>(() => {
    try {
      return sessionStorage.getItem('IKNOWLEDGE_CHAT_CONVERSATION_ID') || '';
    } catch {
      return '';
    }
  });
  const [isChatOpen, setIsChatOpen] = React.useState<boolean>(false);
  const [isChatMaximized, setIsChatMaximized] = React.useState<boolean>(false);
  const [isProcessing, setIsProcessing] = React.useState<boolean>(false);
  const [botFeedbackByMessage, setBotFeedbackByMessage] = React.useState<Record<number, 'like' | 'dislike' | 'copied' | undefined>>({});
  const [botTypingCompleteByMessage, setBotTypingCompleteByMessage] = React.useState<Record<number, boolean>>({});
  const messagesEndRef = React.useRef<HTMLDivElement>(null);
  const previousUrlRef = React.useRef<string | null>(null);
  const currentUserDisplayName = props.context?.pageContext?.user?.displayName || "Deepa Joseph";
  const knowledgeSearchApiClient = React.useMemo(
    () => new KnowledgeSearchApiClient(undefined, props.context),
    [props.context]
  );
  const isBotThinkingMessage = (messageText: string): boolean =>
    messageText === "Thinking..." ||
    messageText === "Searching documents..." ||
    messageText === "Checking FAQs..." ||
    messageText === "Generating response..." ||
    messageText === "Searching..." ||
    messageText === "Searching document content...";
  const setBotFeedback = (messageIndex: number, feedback: 'like' | 'dislike' | 'copied'): void => {
    setBotFeedbackByMessage((previous) => ({
      ...previous,
      [messageIndex]: previous[messageIndex] === feedback && feedback !== 'copied' ? undefined : feedback
    }));
  };

  const markBotTypingComplete = React.useCallback((messageIndex: number): void => {
    setBotTypingCompleteByMessage((previous) =>
      previous[messageIndex]
        ? previous
        : { ...previous, [messageIndex]: true }
    );
  }, []);

  // Search/chat runtime refs
  const searchServiceRef = React.useRef<SharePointSearchService | null>(null);
  const allKMDocumentsCache = React.useRef<any[] | null>(null);
  const focusedDocumentsRef = React.useRef<FocusedDocument[]>([]);
  
  // 🚩 SAFE-RAG: Cache of Active Document IDs and URLs from SharePoint
  const activeDocumentCacheRef = React.useRef<Set<string>>(new Set());
  const isRefreshingActiveStatus = React.useRef<boolean>(false);
  const activeDocumentCacheFetchedAtRef = React.useRef<number>(0);


  // No OData filter available because the Azure Search index schema
  // does not contain a Status field. Returning all retrieved documents.
  const KM_DATA_HUB_FILTER: string | undefined = undefined;

  // Initialize Services
  React.useEffect(() => {
    if (!searchServiceRef.current && props.context) {
      searchServiceRef.current = getSearchServiceInstance(
        props.context.spHttpClient,
        props.context.pageContext.web.absoluteUrl,
        LIBRARY_NAMES.kmDataHub
      );
    }
  }, [props.context]);

  /**
   * Fetches the current list of 'Active' document URLs directly from SharePoint.
   * This is used to filter Azure Search results because the index lacks a Status field.
   */
  const refreshActiveStatus = async (): Promise<Set<string>> => {
    if (!props.context) return new Set();

    const ACTIVE_STATUS_CACHE_MS = 5 * 60 * 1000;
    if (
      activeDocumentCacheRef.current.size > 0 &&
      Date.now() - activeDocumentCacheFetchedAtRef.current < ACTIVE_STATUS_CACHE_MS
    ) {
        return activeDocumentCacheRef.current;
    }
    
    // Simple lock to avoid redundant parallel refreshes
    if (isRefreshingActiveStatus.current) {
        // Wait briefly if already refreshing, or return current set
        return activeDocumentCacheRef.current;
    }

    try {
      isRefreshingActiveStatus.current = true;
      
      const siteUrl = props.context.pageContext.web.absoluteUrl;
      const libraryName = LIBRARY_NAMES.kmDataHub;
      const searchService = searchServiceRef.current || getSearchServiceInstance(
        props.context.spHttpClient,
        siteUrl,
        libraryName
      );
      let ids: number[] = [];
      try {
        const activeResult = await searchService.getDocumentIdsByStatus('Active', 0, 100, '');
        ids = activeResult.ids || [];
      } catch {
        activeDocumentCacheRef.current = new Set();
        activeDocumentCacheFetchedAtRef.current = Date.now();
        return activeDocumentCacheRef.current;
      }
      const activeItems = new Set<string>();
      ids.forEach((id) => activeItems.add(String(id)));

      activeDocumentCacheRef.current = activeItems;
      activeDocumentCacheFetchedAtRef.current = Date.now();
      return activeItems;
    } catch {
      activeDocumentCacheRef.current = new Set();
      activeDocumentCacheFetchedAtRef.current = Date.now();
      return activeDocumentCacheRef.current;
    } finally {
      isRefreshingActiveStatus.current = false;
    }
  };

  React.useEffect(() => {
    if (props.context) {
      (async () => {
        await Promise.all([
          refreshActiveStatus(),
          fetchLatestDocuments()
        ]);
      })();
    }
  }, [documentsRefreshKey, props.context]);

  React.useEffect(() => {
    return subscribeToDocumentDataChanged(() => {
      setDocumentsRefreshKey((value) => value + 1);
    });
  }, []);

  const isUrlActive = (url: string, id?: any): boolean => {
    // 🚩 STRATEGY 1: Check by SharePoint Item ID (100% Reliability)
    if (id && activeDocumentCacheRef.current.has(String(id))) {
        return true;
    }

    if (!url) return false;
    
    // 🚩 STRATEGY 2: Fallback to Path Matching
    let finalPath = url.toLowerCase();
    
    // If it's a full URL, extract the path
    try {
        if (finalPath.startsWith('http')) {
            const urlObj = new URL(finalPath);
            finalPath = urlObj.pathname.toLowerCase();
        }
    } catch(e) { 
      // Ignored intentional format fallback
    }

    // Remove any trailing slashes
    if (finalPath.endsWith('/')) finalPath = finalPath.slice(0, -1);
    
    // Decoded version for space matching
    const decodedPath = decodeURIComponent(finalPath);
    
    const isActive = activeDocumentCacheRef.current.has(finalPath) || activeDocumentCacheRef.current.has(decodedPath);
    
    return isActive;
  };

  const fetchLatestDocuments = async () => {
    if (!props.context || !searchServiceRef.current) {
      setLoading(false);
      return;
    }

    try {
      debugLog('=== FETCHING LATEST DOCUMENTS FROM SEARCH INDEX (LIBRARY AGNOSTIC) ===');

      const searchResults = await searchServiceRef.current.getLatestDocuments(20); // Fetch more to allow for filtering

      // SAFE-RAG: Filter results by SharePoint Active Status
      const activeResults = searchResults.filter(item => isUrlActive(item.url, item.id));

      debugLog(`Filtered results: ${searchResults.length} total -> ${activeResults.length} active.`);

      if (activeResults.length === 0) {
        setDocuments([]);
        setLoading(false);
        return;
      }

      const formattedItems: DocumentItem[] = activeResults.slice(0, 6).map((item: any) => {
        const fileExtension = item.url ? item.url.split('.').pop()?.toUpperCase() : (item.fileType || 'DOC');
        const fileName = normalizeTextField(getItemFieldValue(item, COLUMN_NAMES.fileLeafRef, ['fileName', COLUMN_NAMES.fileLeafRef])) || item.title;
        const name = normalizeTextField(getItemFieldValue(item, COLUMN_NAMES.title, ['title', COLUMN_NAMES.title])) || item.title;
        return {
          id: item.id || Math.floor(Math.random() * 100000),
          name,
          abstract: normalizeDescriptionText(
            getItemFieldValue(item, COLUMN_NAMES.description, ['description', COLUMN_NAMES.description])
          ),
          fileType: fileExtension,
          author: item.author || 'Internal',
          date: item.publishedDate ? new Date(item.publishedDate).toLocaleDateString('en-US', {
            month: 'short',
            day: 'numeric',
            year: 'numeric'
          }) : '',
          views: normalizeNumberField(getItemFieldValue(item, COLUMN_NAMES.views, ['views', COLUMN_NAMES.views])),
          comments: normalizeNumberField(getItemFieldValue(item, COLUMN_NAMES.comments, ['comments', COLUMN_NAMES.comments])),
          likes: normalizeNumberField(getItemFieldValue(item, COLUMN_NAMES.likes, ['likes', COLUMN_NAMES.likes])),
          downloads: normalizeNumberField(getItemFieldValue(item, COLUMN_NAMES.downloads, ['downloads', COLUMN_NAMES.downloads])),
          fileSize: '',
          serverRelativeUrl: item.url,
          fileRef: item.url,
          fileUniqueId: item.id,
          listTitle: item.listTitle || LIBRARY_NAMES.kmDataHub,
          fileName,
          docIcon: normalizeTextField(getItemFieldValue(item, COLUMN_NAMES.docIcon, ['docIcon', COLUMN_NAMES.docIcon])),
          previewImageUrls: item.url ? buildPreviewCandidates(
            props.context!.pageContext.web.absoluteUrl,
            item.url,
            String(props.context!.pageContext.site.id || ''),
            String(props.context!.pageContext.web.id || ''),
            item.id
          ) : []
        };
      });

      const engagementCounts = await fetchEngagementCounts(formattedItems.map((item) => Number(item.id) || 0));

      setDocuments(
        formattedItems.map((item) => ({
          ...item,
          views: engagementCounts[Number(item.id)]?.views || 0,
          comments: engagementCounts[Number(item.id)]?.comments || 0,
          likes: engagementCounts[Number(item.id)]?.likes || 0,
          downloads: engagementCounts[Number(item.id)]?.downloads || 0
        }))
      );
      setLoading(false);

      (async () => {
        try {
          if (!searchServiceRef.current) return;
          const allDocs = await searchServiceRef.current.getLatestDocuments(1000);

          // Only cache documents that are ACTUALLY active in SharePoint
          const filteredAllDocs = allDocs.filter(d => isUrlActive(d.url, d.id));

          allKMDocumentsCache.current = filteredAllDocs.map(d => ({
            Id: d.id,
            Title: d.title,
            [COLUMN_NAMES.fileRef]: d.url,
            [COLUMN_NAMES.status]: 'Active'
          }));
          debugLog(`Search-driven background cache populated with ${allKMDocumentsCache.current?.length} ACTIVE documents`);
        } catch (cacheErr) {
          console.warn('Background search cache fetch failed:', cacheErr);
        }
      })();
    } catch (error) {
      console.error('=== ERROR FETCHING DOCUMENTS ===', error);
      setDocuments([]);
    } finally {
      setLoading(false);
    }
  };

  const getFileTypeIcon = (fileType: string): string => {
    const type = (fileType || '').toLowerCase();
    if (type === 'pdf') {
      return '📄';
    } else if (type === 'pptx' || type === 'ppt') {
      return '📊';
    } else if (type === 'docx' || type === 'doc') {
      return '📝';
    } else if (type === 'xlsx' || type === 'xls') {
      return '📈';
    } else if (['mp4', 'mov', 'webm', 'mkv', 'avi', 'm4v', 'wmv'].indexOf(type) !== -1 || type.indexOf('video') !== -1) {
      return '▶';
    } else if (['mp3', 'wav', 'm4a', 'aac', 'flac', 'ogg'].indexOf(type) !== -1 || type.indexOf('audio') !== -1) {
      return '♪';
    }
    return '📎';
  };

  const handleView = (item: DocumentItem) => {
    if (!props.context) return;

    // 🚩 UI SAFETY: If ID is '0' or non-numeric, open the file directly to prevent Detail Page crashes
    // The Detail Page uses OData filters (Id eq ...) which require an integer.
    const isNumericId = item.id && !isNaN(Number(item.id)) && Number(item.id) > 0;
    
    if (!isNumericId) {
      debugLog(`🔗 Broad Discovery Document detected (ID: ${item.id}). Opening URL directly.`);
      const directUrl = item.serverRelativeUrl || item.fileRef || '';
      if (directUrl) {
        window.open(directUrl, '_blank');
      }
      return;
    }

    openAppPageInNewTab(NAV_PATHS.asset, { assetID: String(item.id) });
  };

  const handleCloseDetail = () => {
    setSelectedDocumentId(null);
    setSelectedListTitle(null);
    setSelectedListId('');
    previousUrlRef.current = null;
    pushPageUrl(NAV_PATHS.home);
  };

  const fetchEngagementCounts = React.useCallback(async (
    documentIds: number[]
  ): Promise<Record<number, { views: number; comments: number; likes: number; downloads: number }>> => {
    if (!props.context || documentIds.length === 0) {
      return {};
    }

    try {
      const cachedMetrics = (() => {
        try {
          const cached = sessionStorage.getItem(METRICS_CACHE_KEY);
          if (!cached) return null;
          const parsed = JSON.parse(cached);
          if (Date.now() - parsed.timestamp > METRICS_TTL_MS) return null;
          const cachedIds = Array.isArray(parsed.documentIds) ? parsed.documentIds : [];
          if (JSON.stringify(cachedIds) !== JSON.stringify(documentIds)) return null;
          return parsed.data;
        } catch {
          return null;
        }
      })();
      if (cachedMetrics) {
        return cachedMetrics;
      }

      const webUrl = props.context.pageContext.web.absoluteUrl;
      const filter = documentIds.map((id) => `DocumentId eq ${id}`).join(' or ');
      const buildRequest = (listName: string): Promise<SPHttpClientResponse | null> =>
        props.context!.spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=DocumentId,Id&$filter=${encodeURIComponent(filter)}`,
          SPHttpClient.configurations.v1
        ).catch((): null => null);

      const [viewsResponse, commentsResponse, likesResponse, downloadsResponse] = await Promise.all([
        buildRequest(LIST_NAMES.documentViews),
        buildRequest(LIST_NAMES.documentComments),
        buildRequest(LIST_NAMES.documentLikes),
        buildRequest(LIST_NAMES.documentDownloads)
      ]);

      const counts: Record<number, { views: number; comments: number; likes: number; downloads: number }> = {};
      documentIds.forEach((id) => {
        counts[id] = { views: 0, comments: 0, likes: 0, downloads: 0 };
      });

      const applyCounts = async (
        response: SPHttpClientResponse | null,
        key: 'views' | 'comments' | 'likes' | 'downloads'
      ): Promise<void> => {
        if (!response || !response.ok) {
          return;
        }

        const data = await response.json();
        (data.value || []).forEach((item: { DocumentId?: number }) => {
          if (typeof item.DocumentId !== 'number' || !counts[item.DocumentId]) {
            return;
          }

          counts[item.DocumentId][key] += 1;
        });
      };

      await Promise.all([
        applyCounts(viewsResponse, 'views'),
        applyCounts(commentsResponse, 'comments'),
        applyCounts(likesResponse, 'likes'),
        applyCounts(downloadsResponse, 'downloads')
      ]);

      try {
        sessionStorage.setItem(METRICS_CACHE_KEY, JSON.stringify({
          timestamp: Date.now(),
          documentIds,
          data: counts
        }));
      } catch {
        // Ignore session storage failures.
      }

      return counts;
    } catch {
      return {};
    }
  }, [props.context]);

  const fetchBotDocumentMetadata = React.useCallback(async (
    documentIds: number[]
  ): Promise<Record<number, Partial<DocumentItem> & { author: string; publishedDate?: string }>> => {
    if (!props.context || documentIds.length === 0) {
      return {};
    }

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      const listName = LIBRARY_NAMES.kmDataHub;
      const fieldMap = fieldMapModuleCache || await fetchKMDataHubReadFieldMap(props.context.spHttpClient, webUrl, listName);
      fieldMapModuleCache = fieldMap;
      const queryParts = buildKMDataHubItemQuery(
        fieldMap,
        [
          'Id',
          fieldMap.title || COLUMN_NAMES.title,
          fieldMap.description || COLUMN_NAMES.description,
          fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef,
          fieldMap.fileRef || COLUMN_NAMES.fileRef,
          fieldMap.status || COLUMN_NAMES.status,
          fieldMap.published || COLUMN_NAMES.published,
          fieldMap.docIcon || COLUMN_NAMES.docIcon,
          'Modified'
        ],
        []
      );

      const filter = documentIds.map((id) => `Id eq ${id}`).join(' or ');
      const response = await props.context.spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=${queryParts.select}&$expand=${queryParts.expand}&$filter=(${filter}) and ${fieldMap.status || COLUMN_NAMES.status} eq 'Active'`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        return {};
      }

      const data = await response.json();
      const metadata: Record<number, Partial<DocumentItem> & { author: string; publishedDate?: string }> = {};
      (data.value || []).forEach((item: any) => {
        if (typeof item.Id !== 'number') {
          return;
        }

        const fileName = normalizeTextField(getItemFieldValue(item, fieldMap.fileLeafRef || COLUMN_NAMES.fileLeafRef, [COLUMN_NAMES.fileLeafRef]));
        const title = normalizeTextField(getItemFieldValue(item, fieldMap.title || COLUMN_NAMES.title, [COLUMN_NAMES.title]));
        const name = title || fileName || `Document ${item.Id}`;

        metadata[item.Id] = {
          name,
          abstract: normalizeDescriptionText(
            getItemFieldValue(item, fieldMap.description || COLUMN_NAMES.description, [COLUMN_NAMES.description])
          ),
          author: resolveDocumentAuthor(item, '', fieldMap.author),
          publishedDate: resolveDocumentPublishedValue(item, fieldMap.published)?.toString(),
          fileName,
          fileRef: normalizeTextField(getItemFieldValue(item, fieldMap.fileRef || COLUMN_NAMES.fileRef, [COLUMN_NAMES.fileRef])),
          docIcon: normalizeTextField(getItemFieldValue(item, fieldMap.docIcon || COLUMN_NAMES.docIcon, [COLUMN_NAMES.docIcon]))
        };
      });

      return metadata;
    } catch {
      return {};
    }
  }, [props.context]);

  const buildVisibleDocuments = React.useCallback((
    citations: ExternalChatMessage['citations'],
    engagementCounts?: Record<number, { views: number; comments: number; likes: number; downloads: number }>,
    botDocumentMetadata?: Record<number, Partial<DocumentItem> & { author: string; publishedDate?: string }>
  ): DocumentItem[] => {
    return (citations || []).map((cite) => {
      const urlStr = cite.url || '';
      const fileExtension = urlStr ? urlStr.split('.').pop()?.toUpperCase() : 'DOC';
      const documentId = Number(cite.parentDocumentId || cite.id) || 0;
      const metadata = botDocumentMetadata?.[documentId];
      const metrics = engagementCounts?.[documentId];
      const publishedValue = metadata?.publishedDate || cite.publishedDate;

      return {
        id: documentId,
        name: metadata?.name || cite.title || 'Untitled Document',
        abstract: metadata?.abstract || cite.abstract || cite.description || '',
        fileType: fileExtension,
        author: metadata?.author || cite.author || 'Internal',
        date: publishedValue ? new Date(publishedValue).toLocaleDateString('en-US', {
          month: 'short', day: 'numeric', year: 'numeric'
        }) : '',
        views: metrics?.views ?? metadata?.views ?? cite.viewsCount ?? 0,
        comments: metrics?.comments ?? metadata?.comments ?? cite.commentsCount ?? 0,
        likes: metrics?.likes ?? metadata?.likes ?? cite.likesCount ?? 0,
        downloads: metrics?.downloads ?? metadata?.downloads ?? 0,
        businessUnit: cite.businessUnit || '',
        documentType: cite.documentType || '',
        fileRef: metadata?.fileRef || urlStr,
        serverRelativeUrl: urlStr,
        fileName: metadata?.fileName || '',
        docIcon: metadata?.docIcon,
        fileSize: '',
        listTitle: cite.listTitle || LIBRARY_NAMES.kmDataHub,
        listId: cite.listId || ''
      };
    });
  }, []);

  const hydrateLatestBotDocuments = React.useCallback(async (
    citations: ExternalChatMessage['citations']
  ): Promise<void> => {
    const visibleCitations = (citations || [])
      .filter(cite => cite && cite.title && !cite.title.toLowerCase().includes("who's who"))
      .slice(0, 5);

    const visibleDocumentIds = Array.from(new Set(
      visibleCitations
        .map((cite) => Number(cite.parentDocumentId || cite.id))
        .filter((id) => !isNaN(id) && id > 0)
    ));

    if (visibleDocumentIds.length === 0) {
      return;
    }

    const [engagementCounts, botDocumentMetadata] = await Promise.all([
      fetchEngagementCounts(visibleDocumentIds),
      fetchBotDocumentMetadata(visibleDocumentIds)
    ]);

    setChatMessages((prev) => {
      if (prev.length === 0) {
        return prev;
      }

      const updated = [...prev];
      const lastIndex = updated.length - 1;
      const lastMessage = updated[lastIndex];
      if (!lastMessage || lastMessage.sender !== 'bot' || !lastMessage.citations) {
        return prev;
      }

      updated[lastIndex] = {
        ...lastMessage,
        documents: buildVisibleDocuments(visibleCitations, engagementCounts, botDocumentMetadata)
      };

      return updated;
    });
  }, [buildVisibleDocuments, fetchBotDocumentMetadata, fetchEngagementCounts]);

  const handleDownload = async (item: DocumentItem) => {
    if (props.context && item.serverRelativeUrl) {
      const webUrl = props.context.pageContext.web.absoluteUrl;
      let serverRelativeUrl = item.serverRelativeUrl;

      // Ensure proper server relative URL format
      if (!serverRelativeUrl.startsWith('/')) {
        serverRelativeUrl = `/${serverRelativeUrl}`;
      }

      try {
        await downloadSharePointFile(
          props.context,
          webUrl,
          serverRelativeUrl,
          item.fileName || item.name
        );
      } catch (error) {
        console.error('Download error:', error);
      }
    }
  };

  const displayTiles = React.useMemo(() => documents, [documents]);

  const handleViewAll = () => {
    openAppPageInNewTab(NAV_PATHS.kmReviewHub);
  };

  const handleCloseViewAll = () => {
    setShowViewAll(false);
  };

  const handleViewDocumentFromList = (documentId: number) => {
    openAppPageInNewTab(NAV_PATHS.asset, { assetID: documentId.toString() });
  };

  const handleBackToLibrary = () => {
    setSelectedDocumentId(null);
    setSelectedListTitle(null);
    setShowViewAll(true);
  };

  /**
   * Fetches all documents across ALL libraries using Azure Search.
   * This is used when the user asks for "all documents".
   */
  const getAllDocumentLinks = async (): Promise<{ text: string; documents: DocumentItem[] }> => {
    if (!searchServiceRef.current) {
      return { text: "Search service not available.", documents: [] };
    }

    try {
      // 🚩 CRITERIA: Only show documents with Status = 'Active' (Dynamic Guardrail)
      const guardrailFilter = "status eq 'Active'";

      const items = await searchServiceRef.current.getLatestDocuments(100, guardrailFilter);

      if (items.length === 0) {
        return {
          text: "I couldn't find any documents across the libraries right now.",
          documents: []
        };
      }

      let linksText = `Here are the latest active documents across all libraries (${items.length} found):\n\n`;

      const mapToDocumentItem = (item: any): DocumentItem => {
        return {
          id: item.id,
          name: item.title,
          abstract: item.description || '',
          fileType: item.url ? item.url.split('.').pop()?.toUpperCase() || '' : 'DOC',
          author: item.author || 'Internal',
          date: item.publishedDate ? new Date(item.publishedDate).toLocaleDateString() : '',
          views: 0,
          comments: 0,
          likes: 0,
          downloads: 0,
          fileSize: '',
          serverRelativeUrl: item.url,
          fileRef: item.url,
          previewImageUrls: item.url ? buildPreviewCandidates(
            props.context!.pageContext.web.absoluteUrl,
            item.url,
            String(props.context?.pageContext.site.id || ''),
            String(props.context?.pageContext.web.id || ''),
            item.id
          ) : []
        };
      };

      const topDocuments = items.slice(0, 5).map(mapToDocumentItem);

      return { text: linksText, documents: topDocuments };
    } catch (error) {
      console.error('Error fetching all document links:', error);
      return {
        text: "I encountered an error while fetching the documents. Please try again.",
        documents: []
      };
    }
  };

  // Helper function to check if a title is a placeholder or generic text
  const isPlaceholderTitle = (title: string): boolean => {
    if (!title || title.trim().length === 0) {
      return true;
    }

    const titleLower = title.toLowerCase().trim();

    // Common placeholder patterns
    const placeholderPatterns = [
      /^lorem ipsum/i,
      /^dolor sit amet/i,
      /^powerpoint presentation$/i,
      /^presentation$/i,
      /^document$/i,
      /^untitled$/i,
      /^new document$/i,
      /^new file$/i,
      /^file$/i,
      /^diam nonummy nibh/i,
      /^euismod tincidunt/i,
      /^laoreet dolore/i,
      /^nonummy nibh euismod/i,
      /^sit amet diam/i
    ];

    // Check if title matches any placeholder pattern
    for (const pattern of placeholderPatterns) {
      if (pattern.test(titleLower)) {
        return true;
      }
    }

    // Check if title is just "Lorem Ipsum" text (common placeholder)
    if (titleLower.includes('lorem ipsum') && titleLower.includes('dolor sit')) {
      return true;
    }

    // Check if title is very short and generic (less than 3 words and common words)
    const words = titleLower.split(/\s+/).filter(w => w.length > 0);
    if (words.length <= 2) {
      const genericWords = ['presentation', 'document', 'file', 'untitled', 'new', 'powerpoint'];
      if (words.every(w => genericWords.includes(w))) {
        return true;
      }
    }

    return false;
  };

  // Helper function to clean document titles aggressively
  const cleanDocumentTitle = (title: string, fileName?: string): string => {
    if (!title || title.trim().length === 0) {
      return fileName ? fileName.replace(/\.[^.]+$/, '').trim() : 'Document';
    }

    let cleanTitle = title.replace(/\.[^.]+$/, '').trim();

    const hasUrlPatterns = cleanTitle.includes('%') ||
      cleanTitle.includes('&action=') ||
      cleanTitle.includes('?sourcedoc=') ||
      cleanTitle.includes('://') ||
      cleanTitle.match(/dashboard%20/i) ||
      cleanTitle.match(/\.pptx.*&action=/i) ||
      cleanTitle.match(/%.*%.*%/); // Multiple encoded patterns

    // If malformed and we have fileName, use fileName instead immediately
    if (hasUrlPatterns && fileName) {
      console.warn('⚠️ Title is malformed, using fileName instead. Original title:', cleanTitle);
      cleanTitle = fileName.replace(/\.[^.]+$/, '').trim();
      // Clean the fileName too in case it's also malformed
      cleanTitle = cleanTitle.replace(/%[0-9A-Fa-f]{2}/g, ' ')
        .replace(/[?&][^\s)]*/g, '')
        .replace(/[)\]]+$/, '')
        .replace(/[%&?=]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

      // If fileName is also malformed, try to extract a clean name
      if (cleanTitle.includes('%') || cleanTitle.includes('&action=')) {
        // Try to find a pattern like "Case Study-Scientific Expert..." before the malformed part
        const cleanPart = cleanTitle.split('%')[0].split('&')[0].split('?')[0].trim();
        if (cleanPart && cleanPart.length > 5) {
          cleanTitle = cleanPart;
        } else {
          // Last resort: use a generic name based on the first part
          const firstWords = cleanTitle.split(/\s+/).slice(0, 5).join(' ');
          cleanTitle = firstWords.replace(/[%&?=]/g, '').trim() || 'Document';
        }
      }
    } else {
      let decodeAttempts = 0;
      while (cleanTitle.includes('%') && decodeAttempts < 3) {
        try {
          const decoded = decodeURIComponent(cleanTitle);
          if (decoded !== cleanTitle) {
            cleanTitle = decoded;
          } else {
            break;
          }
        } catch (e) {
          break;
        }
        decodeAttempts++;
      }

      // Remove URL patterns
      cleanTitle = cleanTitle.replace(/https?:\/\/[^\s)]*/g, '').trim();
      // Remove encoded patterns (replace %20, %2F, etc. with spaces)
      cleanTitle = cleanTitle.replace(/%[0-9A-Fa-f]{2}/g, ' ').trim();
      // Remove query parameters and fragments
      cleanTitle = cleanTitle.replace(/[?&][^)\s]*/g, '').trim();
      // Remove file extension with query params
      cleanTitle = cleanTitle.replace(/\.[a-z]{3,5}[?&][^)\s]*/gi, '').trim();
      // Remove anything after patterns like "dashboard%20" or "&action="
      cleanTitle = cleanTitle.replace(/dashboard%20.*$/i, 'dashboard').trim();
      cleanTitle = cleanTitle.replace(/&action=.*$/i, '').trim();
      cleanTitle = cleanTitle.replace(/\?sourcedoc=.*$/i, '').trim();
      // Remove trailing parentheses/brackets
      cleanTitle = cleanTitle.replace(/[)\]]+$/, '').trim();
      // Remove URL-related characters
      cleanTitle = cleanTitle.replace(/[%&?=]/g, ' ').trim();
      // Normalize spaces
      cleanTitle = cleanTitle.replace(/\s+/g, ' ').trim();
    }

    // Final check: if still malformed after all cleaning, use fileName or extract clean part
    if (cleanTitle.includes('%') || cleanTitle.includes('&action=') || cleanTitle.includes('?sourcedoc=')) {
      console.warn('⚠️ Title still malformed after cleaning, attempting to extract clean part:', cleanTitle);

      // Try to extract the clean part before the malformed section
      const cleanPart = cleanTitle.split('%')[0].split('&')[0].split('?')[0].trim();
      if (cleanPart && cleanPart.length > 3) {
        cleanTitle = cleanPart;
      } else if (fileName) {
        // Use fileName and clean it
        cleanTitle = fileName.replace(/\.[^.]+$/, '').trim();
        cleanTitle = cleanTitle.replace(/%[0-9A-Fa-f]{2}/g, ' ')
          .replace(/[?&][^\s)]*/g, '')
          .replace(/[)\]]+$/, '')
          .replace(/[%&?=]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();
      } else {
        cleanTitle = 'Document';
      }
    }

    // Final validation
    if (!cleanTitle || cleanTitle.length === 0) {
      cleanTitle = 'Document';
    }

    return cleanTitle;
  };

  // Convert document references to clickable links
  const convertDocumentReferencesToLinks = async (text: string): Promise<string> => {
    if (!props.context) return text;

    const quotedReferencePattern = /(?:\*?Reference:\s*)?\*?'([^']{1,100})'\*?([.,;:!?]?)/g;

    const markdownDocumentPattern = /\*'([A-Za-z0-9_][A-Za-z0-9_\s-]{4,99}[A-Za-z0-9_])'\*/g;

    const unquotedReferencePattern = /\b([A-Z][A-Z0-9_]{4,}(?:_[A-Z0-9_]+)*)([.,;:!?]?)/g;

    const matches: Array<{ fullMatch: string; referenceName: string; punctuation: string; index: number; isQuoted: boolean }> = [];
    let match: RegExpExecArray | null;

    const isInsideLink = (textIdx: number): boolean => {
      const textUntilMatch = text.substring(0, textIdx);
      const lastOpenBracket = textUntilMatch.lastIndexOf('[');
      const lastCloseBracket = textUntilMatch.lastIndexOf(']');
      const lastOpenParen = textUntilMatch.lastIndexOf('(');
      const lastCloseParen = textUntilMatch.lastIndexOf(')');

      // If open bracket came after close bracket, we're inside [brackets]
      // If open paren came after close paren, we're inside (parens)
      return (lastOpenBracket > lastCloseBracket) || (lastOpenParen > lastCloseParen);
    };

    // Helper to check if a match is part of a URL (e.g. m365x47785454.sharepoint.com/...)
    const isPartOfUrl = (textIdx: number, matchedText: string): boolean => {
      const start = Math.max(0, textIdx - 50);
      const end = Math.min(text.length, textIdx + matchedText.length + 50);
      const surroundingText = text.substring(start, end);
      const urlRegex = /https?:\/\/[^\s]+/g;
      let urlMatch: RegExpExecArray | null;
      while ((urlMatch = urlRegex.exec(surroundingText)) !== null) {
        const urlStartInSurrounding = urlMatch.index;
        const urlEndInSurrounding = urlMatch.index + urlMatch[0].length;
        const matchStartInSurrounding = textIdx - start;
        const matchEndInSurrounding = matchStartInSurrounding + matchedText.length;

        if (matchStartInSurrounding >= urlStartInSurrounding && matchEndInSurrounding <= urlEndInSurrounding) {
          return true;
        }
      }
      return false;
    };

    // Find quoted references
    while ((match = quotedReferencePattern.exec(text)) !== null) {
      if (isInsideLink(match.index) || isPartOfUrl(match.index, match[0])) continue;

      const referenceName = match[1];
      const hasTooManySpaces = (referenceName.match(/\s/g) || []).length > 3;
      const looksLikeSentence = /\b(the|and|or|is|are|was|were|this|that|these|those)\b/i.test(referenceName);
      const hasValidChars = /^[A-Za-z0-9_\s.-]+$/.test(referenceName);
      const isTooLong = referenceName.length > 100;

      if (hasTooManySpaces || looksLikeSentence || !hasValidChars || isTooLong) continue;

      matches.push({
        fullMatch: match[0],
        referenceName: referenceName,
        punctuation: match[2] || '',
        index: match.index,
        isQuoted: true
      });
    }

    // Find markdown references
    while ((match = markdownDocumentPattern.exec(text)) !== null) {
      const isAlreadyMatched = matches.some(m =>
        m.index <= match.index &&
        m.index + m.fullMatch.length >= match.index + match[0].length
      );
      if (isAlreadyMatched || isInsideLink(match.index) || isPartOfUrl(match.index, match[0])) continue;

      matches.push({
        fullMatch: match[0],
        referenceName: match[1],
        punctuation: '',
        index: match.index,
        isQuoted: true
      });
    }

    // Find unquoted document references
    unquotedReferencePattern.lastIndex = 0;
    while ((match = unquotedReferencePattern.exec(text)) !== null) {
      const isAlreadyMatched = matches.some(m =>
        m.index <= match.index &&
        m.index + m.fullMatch.length >= match.index + match[0].length
      );

      if (isAlreadyMatched || isInsideLink(match.index) || isPartOfUrl(match.index, match[0])) continue;

      const refName = match[1];
      // Avoid matching common words that might fit the pattern
      const commonWords = ['Reference', 'Document', 'SharePoint', 'Artifacts', 'Knowledge'];
      if (commonWords.includes(refName)) continue;

      if (refName.includes('_') || (refName.match(/[A-Z]/g) || []).length >= 2) {
        matches.push({
          fullMatch: match[0],
          referenceName: refName,
          punctuation: match[2] || '',
          index: match.index,
          isQuoted: false
        });
      }
    }

    // Sort matches by index in reverse order for safe replacement
    matches.sort((a, b) => b.index - a.index);

    if (matches.length === 0) return text;

    try {
      if (!searchServiceRef.current) return text;

      debugLog('🔗 Converting document references using Azure AI Search (Library Agnostic)');

      const rawItems = await searchServiceRef.current.searchDocuments('*', KM_DATA_HUB_FILTER);
      
      // 🚩 SAFE-RAG: Only allow linking to ACTIVE documents
      const items = rawItems.filter(doc => isUrlActive(doc.url, doc.id));

      if (items.length === 0) return text;

      let processedText = text;

      for (let i = 0; i < matches.length; i++) {
        const match = matches[i];
        const referenceName = match.referenceName;
        const fullMatch = match.fullMatch;
        const punctuation = match.punctuation; // Preserve trailing punctuation
        const matchIndex = match.index; // Original index in text
        const isQuoted = match.isQuoted;

        debugLog(`🔗 Processing document reference: "${referenceName}" (quoted: ${isQuoted}) at index ${matchIndex}`);
        debugLog(`🔗 Full match: "${fullMatch}"`);

        // Normalize reference name: remove underscores, special chars, convert to lowercase Search for matching document by filename or title
        const normalizedRef = referenceName
          .replace(/[_-]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim()
          .toLowerCase();

        // Also keep original reference for exact matching
        const originalRefLower = referenceName.toLowerCase().trim();

        const webUrl = props.context?.pageContext.web.absoluteUrl || '';
        const urlObj = new URL(webUrl);
        const serverRoot = `${urlObj.protocol}//${urlObj.host}`;

        // Initialize matching variables
        let bestMatch: ISearchResult | null = null;
        let bestScore = 0;

        items.forEach((item: ISearchResult) => {
          const fileName = item.url ? item.url.split('/').pop() || '' : '';
          const title = item.title || '';

          // Normalize document names
          const normalizedFileName = fileName
            .replace(/[_-]/g, ' ')
            .replace(/\s+/g, ' ')
            .replace(/\.[^.]+$/, '') // Remove extension
            .trim()
            .toLowerCase();

          const normalizedTitle = title
            .replace(/[_-]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
            .toLowerCase();

          const fileNameLower = fileName.toLowerCase();
          const titleLower = title.toLowerCase();

          // Also try matching without removing file extension for exact matches
          const fileNameNoExt = fileNameLower.replace(/\.[^.]+$/, '');
          const titleNoExt = titleLower.replace(/\.[^.]+$/, '');

          let score = 0;
          let matchReason = '';

          // Exact match gets highest score (with or without extension)
          if (normalizedFileName === normalizedRef || normalizedTitle === normalizedRef) {
            score = 100;
            matchReason = 'normalized exact match';
          }
          // Original reference exact match (with underscores, with or without extension)
          else if (fileNameLower === originalRefLower || titleLower === originalRefLower ||
            fileNameNoExt === originalRefLower || titleNoExt === originalRefLower) {
            score = 95;
            matchReason = 'original exact match';
          }
          // Starts with match
          else if (normalizedFileName.startsWith(normalizedRef) || normalizedTitle.startsWith(normalizedRef)) {
            score = 80;
            matchReason = 'normalized starts with';
          }
          // Original reference starts with (with underscores)
          else if (fileNameLower.startsWith(originalRefLower) || titleLower.startsWith(originalRefLower)) {
            score = 75;
            matchReason = 'original starts with';
          }
          // Contains match (but require minimum length to avoid false matches)
          else if (normalizedRef.length >= 5 &&
            (normalizedFileName.includes(normalizedRef) || normalizedTitle.includes(normalizedRef))) {
            score = 60;
            matchReason = 'normalized contains';
          }
          // Original reference contains (with underscores) - LOWER THRESHOLD for known terms
          else if (originalRefLower.length >= 3 &&
            (fileNameLower.includes(originalRefLower) || titleLower.includes(originalRefLower))) {
            score = 50;
            matchReason = 'original contains';
          }
          // Try partial matching for document names with underscores 
          else if (originalRefLower.includes('_') || (originalRefLower.match(/[A-Z]/g) || []).length >= 2) {
            const refParts = originalRefLower.split('_').filter(p => p.length > 2);
            const fileParts = fileNameLower.split('_').filter(p => p.length > 2);
            const titleParts = titleLower.split('_').filter(p => p.length > 2);

            const matchingFileParts = refParts.filter(part =>
              fileParts.some(fp => fp.includes(part) || part.includes(fp))
            );
            const matchingTitleParts = refParts.filter(part =>
              titleParts.some(tp => tp.includes(part) || part.includes(tp))
            );

            if (matchingFileParts.length >= 2 || matchingTitleParts.length >= 2) {
              score = 55;
              matchReason = 'partial underscore match';
            }
          }

          if (score >= 50 && score > bestScore) {
            bestScore = score;
            bestMatch = item;
            debugLog(`✅ Found match (score: ${score}, reason: ${matchReason}):`, {
              fileName: fileName,
              title: title,
              referenceName: referenceName
            });
          }
        });

        if (!bestMatch) {
          debugLog(`⚠️ No match found with scoring. Trying exact filename match for: "${referenceName}"`);

          const exactMatch = items.find((item: ISearchResult) => {
            const fileName = (item.url ? item.url.split('/').pop() || '' : '').toLowerCase();
            const fileNameNoExt = fileName.replace(/\.[^.]+$/, '');
            const refLower = referenceName.toLowerCase();
            return fileName === refLower ||
              fileNameNoExt === refLower ||
              fileName === `${refLower}.pptx` ||
              fileName === `${refLower}.docx` ||
              fileName === `${refLower}.pdf` ||
              fileName === `${refLower}.xlsx`;
          });

          if (exactMatch) {
            bestMatch = exactMatch;
          } else {
            debugLog(`⚠️ No match found. Searched ${items.length} documents. Reference: "${referenceName}"`);
          }
        }

        const matchingDoc = bestMatch;

        if (matchingDoc) {
          let serverRelativeUrl = matchingDoc.url || '';
          const fileExtension = serverRelativeUrl.split('.').pop()?.toLowerCase() || '';
          let docUrl = '';

          if (['docx', 'doc', 'pptx', 'ppt', 'xlsx', 'xls', 'xlsm', 'xlsb', 'xltx', 'xltm'].includes(fileExtension)) {
            const encodedServerUrl = encodeURIComponent(serverRelativeUrl);
            docUrl = `${webUrl}/_layouts/15/WopiFrame.aspx?sourcedoc=${encodedServerUrl}&action=default`;
          } else {
            docUrl = `${serverRoot}${serverRelativeUrl}`;
          }

          const displayName = matchingDoc.title || (matchingDoc.url ? matchingDoc.url.split('/').pop() : '') || referenceName;
          const linkText = `[${displayName}](${docUrl})${punctuation}`; // Preserve punctuation

          let replacementText = '';

          // Check if the original match had markdown formatting (asterisks)
          const hadMarkdownFormatting = fullMatch.includes('*');

          if (isQuoted) {
            // For quoted references, preserve the quote structure
            // If it had markdown formatting, remove it and use the link
            if (fullMatch.trim().startsWith('Reference:') ||
              processedText.substring(Math.max(0, matchIndex - 20), matchIndex).trim().endsWith('Reference:')) {
              replacementText = `Reference: ${linkText}`;
            } else if (hadMarkdownFormatting) {
              // Handle cases like "*'Document_Name'*" -> "[Document_Name](url)"
              replacementText = linkText;
            } else {
              replacementText = linkText;
            }
          } else {
            // For unquoted references, just replace the document name with the link
            // Preserve any "Reference:" prefix if present
            if (fullMatch.trim().startsWith('Reference:') ||
              processedText.substring(Math.max(0, matchIndex - 20), matchIndex).trim().endsWith('Reference:')) {
              replacementText = `Reference: ${linkText}`;
            } else {
              replacementText = linkText;
            }
          }

          debugLog(`✅ Converting "${referenceName}" to clickable link: ${linkText.substring(0, 100)}...`);

          // Replace at the specific position to preserve all other text
          processedText = processedText.substring(0, matchIndex) +
            replacementText +
            processedText.substring(matchIndex + fullMatch.length);
        } else {
          debugLog(`⚠️ No matching document found for reference: "${referenceName}"`);
        }
      }

      debugLog(`🔗 Document reference conversion complete. Found ${matches.length} references, converted to links.`);

      return processedText;
    } catch (error) {
      console.error('Error converting document references to links:', error);
      return text; // Return original text on error
    }
  };

  // Chatbot handlers
  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: "smooth" });
  };

  React.useEffect(() => {
    scrollToBottom();
  }, [chatMessages]);

  const handleChatbotToggle = () => {
    setIsChatOpen((wasOpen) => {
      if (wasOpen) {
        setIsChatMaximized(false);
      }
      return !wasOpen;
    });
  };

  const handleStartNewChat = React.useCallback(() => {
    setChatMessages([]);
    setChatInput("");
    setChatConversationSummary("");
    setBotFeedbackByMessage({});
    setBotTypingCompleteByMessage({});
    setChatConversationId("");
    try {
      sessionStorage.removeItem('IKNOWLEDGE_CHAT_CONVERSATION_ID');
    } catch {
      // Session storage can be unavailable in some embedded browser contexts.
    }
  }, []);

  // --- NEW RAG FLOW IMPLEMENTATION ---

  /**
   * Optimized RAG Flow: Native SharePoint Search + Azure OpenAI
   */
  const buildBackendChatHistoryContent = React.useCallback((message: ChatMessage): string => {
    const text = (message.text || '').slice(0, 900);
    if (message.sender !== 'bot') {
      return text;
    }

    const sources = [
      ...(message.historyCitations || message.citations || []).map((source) => ({
        title: source.title || '',
        url: source.url || ''
      })),
      ...(message.documents || []).map((source) => ({
        title: source.name || source.fileName || '',
        url: source.serverRelativeUrl || source.fileRef || ''
      }))
    ]
      .filter((source) => source.title || source.url)
      .filter((source, index, all) => {
        const key = `${source.title}|${source.url}`.toLowerCase();
        return all.findIndex((candidate) => `${candidate.title}|${candidate.url}`.toLowerCase() === key) === index;
      })
      .slice(0, 5);

    if (sources.length === 0) {
      return text;
    }

    const sourceLines = sources.map((source, index) => {
      const label = source.title || source.url;
      return source.url ? `${index + 1}. ${label} - ${source.url}` : `${index + 1}. ${label}`;
    });

    return `${text}\n\nSources used:\n${sourceLines.join('\n')}`;
  }, []);

  const getChatResponseNew = async (query: string, history: ChatMessage[], updateStatus: (status: string) => void): Promise<ChatMessage> => {
    try {
      if (knowledgeSearchApiClient.isConfigured()) {
        updateStatus("Searching iKnowledge...");
        const backendHistory = history.slice(-4).map((message) => ({
          role: message.sender === "user" ? "user" : "assistant",
          content: buildBackendChatHistoryContent(message)
        }));
        const backendResponse = await knowledgeSearchApiClient.chat(
          query,
          backendHistory,
          chatConversationId || undefined,
          chatConversationSummary
        );
        const nextConversationId = backendResponse.conversationId || backendResponse.conversation?.id || '';
        if (nextConversationId && nextConversationId !== chatConversationId) {
          setChatConversationId(nextConversationId);
          try {
            sessionStorage.setItem('IKNOWLEDGE_CHAT_CONVERSATION_ID', nextConversationId);
          } catch {
            // Chat memory still works server-side for the current request.
          }
        }
        const allBackendCitations: ISearchResult[] = (backendResponse.citations || [])
          .filter((cite) => cite.kind === "document")
          .map((cite) => {
            const documentId = String(cite.listItemId || cite.documentId || cite.id || "");

            return {
              id: documentId,
              parentDocumentId: documentId,
              title: cite.title,
              content: "",
              url: cite.url,
              score: cite.score,
              semanticScore: cite.score,
              rank: cite.number,
              listTitle: LIBRARY_NAMES.kmDataHub
            };
          });

        const visibleCitations = allBackendCitations
          .slice(0, 5);

        const documentCitations = visibleCitations.filter((cite) => !String(cite.id).startsWith("people-"));
        focusedDocumentsRef.current = documentCitations
          .map((cite) => ({
            id: String(cite.parentDocumentId || cite.id || ''),
            title: cite.title,
            url: cite.url
          }))
          .filter((doc) => doc.id)
          .slice(0, 5);

        const botMessage: ChatMessage = {
          sender: "bot",
          text: backendResponse.answer,
          citations: visibleCitations,
          historyCitations: allBackendCitations.slice(0, 8),
          documents: buildVisibleDocuments(documentCitations),
          timestamp: new Date()
        };

        setChatConversationSummary((previousSummary) =>
          buildNextChatConversationSummary(previousSummary, query, botMessage)
        );

        return botMessage;
      }

      throw new Error("Knowledge Search API is not configured. Chat now requires the Azure AI Search backend and will not fall back to client-side SharePoint document parsing.");

    } catch (error) {
      console.error("🔴 Chatbot Failure:", error);
      return {
        sender: "bot",
        text: "I'm having trouble completing the answer right now. Please try again, or use the search page to review matching documents while I recover.",
        timestamp: new Date()
      };
    }
  };

  const handleSendMessage = async () => {
    const query = chatInput.trim();
    if (!query || isProcessing) return;
    if (query.length > CHAT_INPUT_MAX_CHARS) return;
    const botMessageIndex = chatMessages.length + 1;

    // Add user message
    const userMessage: ChatMessage = { sender: "user", text: query };
    setChatMessages((prev) => [...prev, userMessage]);
    setChatInput("");
    setIsProcessing(true);

    // Process message
    setIsProcessing(true);

    try {
      // Add thinking message (Status will be updated dynamically in NEW flow)
      const thinkingMessage: ChatMessage = { sender: "bot", text: "Searching..." };
      setBotTypingCompleteByMessage((previous) => ({ ...previous, [botMessageIndex]: false }));
      setChatMessages((prev) => [...prev, thinkingMessage]);

      const updateBotStatus = (_statusText: string) => {
        setChatMessages((prev) => {
          const updated = [...prev];
          if (updated.length > 0 && updated[updated.length - 1].sender === "bot") {
            updated[updated.length - 1].text = "Thinking...";
          }
          return updated;
        });
      };

      const newHistory = chatMessages.map(m => ({ sender: m.sender, text: m.text }));
      const finalBotMsg = await getChatResponseNew(query, newHistory, updateBotStatus);

      setChatMessages((prev) => {
        const updated = [...prev];
        updated[updated.length - 1] = finalBotMsg;
        return updated;
      });
      if (finalBotMsg.citations && finalBotMsg.citations.length > 0) {
        void hydrateLatestBotDocuments(finalBotMsg.citations);
      }
      setIsProcessing(false);
    } catch (error) {
      console.error("Error processing message:", error);
      const errorMessage: ChatMessage = {
        sender: "bot",
        text: "I'm sorry, I encountered an error processing your request. Please try again."
      };
      setChatMessages((prev) => {
        const updated = [...prev];
        updated[updated.length - 1] = errorMessage;
        return updated;
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleChatInputKeyPress = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSendMessage();
    }
  };

  return (
    <>
      {showPublishedSection && (
        <div className={styles.questionSection}>
          <div className={styles.sectionHeader}>
            <h2 className={styles.sectionTitle}>Recently Published</h2>
          </div>
          <div className={styles.tilesCarousel}>
            <div className={styles.tilesViewport}>
              <div className={styles.tilesContainer}>
                {loading ? (
                  <div className={styles.loading}>Loading...</div>
                ) : (
                  <>
                    {displayTiles.map((doc) => (
                      <div key={doc.id} className={styles.tile}>
                        <div className={styles.tileBody}>
                          <div className={styles.tileTopMetaCompact}>
                            <span className={styles.tileTopIconCompact}>{doc.docIcon || getFileTypeIcon(doc.fileType)}</span>
                            <span className={styles.tileTopTypeCompact}>{doc.docIcon || doc.fileType || '---'}</span>
                          </div>
                          <div className={styles.tileContentSection}>
                            <h3 className={styles.tileTitle}>{doc.name}</h3>
                            <p className={styles.tileAbstract}>{doc.abstract || ''}</p>
                            <div className={styles.tilePublishMetaCompact}>
                              {doc.fileName && (
                                <p className={styles.tilePublishTextCompact}>File name : {doc.fileName}</p>
                              )}
                              {doc.author && (
                                <div className={styles.tilePublishTextCompact}>
                                  <span>Published by :</span>
                                  {renderAuthorLines(doc.author, styles.authorLineList, styles.authorLineItem)}
                                </div>
                              )}
                              <p className={styles.tilePublishTextCompact}>Published date : {doc.date || '-'}</p>
                            </div>
                          </div>
                        </div>
                        <div className={styles.tileActions}>
                          <button
                            className={styles.viewButton}
                            onClick={() => handleView(doc)}
                          >
                            View
                          </button>
                        </div>
                        <div className={styles.tileStatsCompact}>
                          <span className={styles.tileStatItem}>
                            {renderStatIcon('views')}
                            <span>{doc.views ?? 0}</span>
                          </span>
                          <span className={styles.tileStatItem}>
                            {renderStatIcon('comments')}
                            <span>{doc.comments ?? 0}</span>
                          </span>
                          <span className={styles.tileStatItem}>
                            {renderStatIcon('likes')}
                            <span>{doc.likes ?? 0}</span>
                          </span>
                          <span className={styles.tileStatItem}>
                            {renderStatIcon('downloads')}
                            <span>{doc.downloads ?? 0}</span>
                          </span>
                        </div>
                      </div>
                    ))}
                    {!loading && (
                      <button
                        type="button"
                        className={styles.viewAllTile}
                        onClick={handleViewAll}
                        aria-label="View all published documents"
                      >
                        <span className={styles.viewAllText}>View All</span>
                      </button>
                    )}
                  </>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {typeof document !== "undefined" && createPortal(
        <div className={styles.chatPortalRoot}>
          {/* Chatbot Icon - Floating Button */}
          {!isChatOpen && (
            <div
              className={styles.chatLauncher}
              onClick={handleChatbotToggle}
              title="Open KM Assistant"
            >
              <button type="button" className={styles.chatbotButton} aria-label="Open KM Assistant">
                <img src={chatbotIcon} alt="Bot" className={styles.chatbotButtonIcon} />
              </button>
              <span className={styles.chatLauncherHint} aria-hidden="true">
                <span>Ask me</span>
              </span>
            </div>
          )}

          {/* Chat Window */}
          {isChatOpen && (
            <>
            {isChatMaximized && (
              <div className={styles.chatBackdrop} aria-hidden="true" />
            )}
            <div className={`${styles.chatWindow} ${isChatMaximized ? styles.chatWindowMaximized : ''}`}>
          <div className={styles.chatHeader}>
            <div className={styles.chatHeaderBrand}>
              <img src={chatbotIcon} alt="Bot" className={styles.chatHeaderIcon} />
              <div>
                <span className={styles.chatHeaderTitle}>KM Assistant</span>
                <span className={styles.chatHeaderStatus}><span aria-hidden="true" />Online</span>
              </div>
            </div>
            <div className={styles.chatHeaderActions}>
              <button
                className={`${styles.chatHeaderButton} ${styles.chatNewChatButton} ${isChatMaximized ? styles.chatNewChatButtonExpanded : ''}`}
                onClick={handleStartNewChat}
                aria-label="Start new chat"
                title="Start new chat"
                type="button"
                disabled={isProcessing}
              >
                <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <path d="M6.5 19h11A1.5 1.5 0 0 0 19 17.5v-6.2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M6.5 19A1.5 1.5 0 0 1 5 17.5v-11A1.5 1.5 0 0 1 6.5 5h6.2" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                  <path d="M13.2 12.1l5.25-5.25a1.55 1.55 0 0 0-2.2-2.2L11 9.9l-.65 2.85 2.85-.65Z" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                </svg>
                {isChatMaximized && <span>Start new chat</span>}
              </button>
              <button
                className={styles.chatHeaderButton}
                onClick={() => setIsChatMaximized((value) => !value)}
                aria-label={isChatMaximized ? "Restore chat" : "Maximize chat"}
                title={isChatMaximized ? "Restore chat" : "Maximize chat"}
                type="button"
              >
                {isChatMaximized ? (
                  <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                    <path d="M9 9H4M9 9V4M15 15h5M15 15v5M15 9h5M15 9V4M9 15H4M9 15v5" stroke="currentColor" strokeWidth="2.1" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                ) : (
                  <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                    <path d="M8 4H4v4M16 4h4v4M20 16v4h-4M4 16v4h4" stroke="currentColor" strokeWidth="2.1" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                )}
              </button>
              <button
                className={styles.chatHeaderButton}
                onClick={handleChatbotToggle}
                aria-label="Minimize chat"
                title="Minimize chat"
                type="button"
              >
                <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                  <path d="M7 12h10" stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" />
                </svg>
              </button>
            </div>
          </div>
          <div className={styles.messages}>
            {chatMessages.length === 0 && (
              <div className={styles.welcomeMessage}>
                <h3>Hi {currentUserDisplayName} !</h3>
                <p>Find Faster. Work Smarter.</p>
              </div>
            )}
            {chatMessages.map((msg, index) => (
              <div
                key={index}
                className={`${styles.messageContainer} ${msg.sender === "user"
                  ? styles.messageContainerUser
                  : styles.messageContainerBot
                  }`}
              >
                {msg.sender === "bot" && isBotThinkingMessage(msg.text) ? (
                  <div className={styles.thinking}>
                    <span className={styles.thinkingText}>Thinking, give me a second</span>
                    <span className={styles.thinkingDots} aria-hidden="true">
                      <span />
                      <span />
                      <span />
                    </span>
                  </div>
                ) : (
                <div
                  className={
                    msg.sender === "user" ? styles.userMsg : styles.botMsg
                  }
                >
                  <div className={msg.sender === "bot" ? styles.botAnswerText : undefined}>
                    <TypedMessageContent
                      text={msg.text}
                      styles={styles}
                      enableTyping={msg.sender === "bot" && !botTypingCompleteByMessage[index]}
                      onTypingComplete={msg.sender === "bot" ? () => markBotTypingComplete(index) : undefined}
                      citations={msg.citations}
                      onCitationClick={(citationIndex) => {
                        const citation = msg.citations?.[citationIndex];
                        const referenceDocument = msg.documents?.[citationIndex] || (citation ? buildVisibleDocuments([citation])[0] : undefined);
                        if (referenceDocument) {
                          handleView(referenceDocument);
                        }
                      }}
                    />
                  </div>

                  {msg.url && (
                    <a
                      href={msg.url}
                      target="_blank"
                      rel="noopener noreferrer"
                      className={styles.botLink}
                    >
                      {msg.title || "Open Link"}
                    </a>
                  )}
                  {msg.sender === "bot" && botTypingCompleteByMessage[index] && (
                    <div className={styles.botAnswerFooter}>
                      {msg.citations && msg.citations.length > 0 && (
                        <>
                          <p className={styles.botReferencesLeadIn}>
                            {buildReferenceLeadIn()}
                          </p>
                          <details className={styles.botReferences}>
                          <summary>
                            <span>References</span>
                            <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                              <path d="M7 10l5 5 5-5" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          </summary>
                          <div className={styles.botReferenceList}>
                            {msg.citations.map((cite, cIdx) => {
                              const referenceDocument = msg.documents?.[cIdx] || buildVisibleDocuments([cite])[0];

                              return (
                                <button
                                  key={cIdx}
                                  type="button"
                                  className={styles.botReferenceItem}
                                  onClick={() => handleView(referenceDocument)}
                                >
                                  <span className={styles.botReferenceTitle}>
                                    <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                                      <path d="M10 13a5 5 0 0 0 7.07 0l2.12-2.12a5 5 0 0 0-7.07-7.07L10.9 5.03" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                                      <path d="M14 11a5 5 0 0 0-7.07 0L4.81 13.12a5 5 0 0 0 7.07 7.07l1.22-1.22" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
                                    </svg>
                                    <span>{cite.title || `Reference ${cIdx + 1}`}</span>
                                  </span>
                                  {cite.description && (
                                    <small>{cite.description.length > 140 ? cite.description.substring(0, 140) + "..." : cite.description}</small>
                                  )}
                                </button>
                              );
                            })}
                          </div>
                          </details>
                        </>
                      )}
                      <div className={styles.botFeedbackActions} aria-label="Answer actions">
                        <button
                          type="button"
                          aria-label="Like answer"
                          title="Like answer"
                          aria-pressed={botFeedbackByMessage[index] === 'like'}
                          className={`${styles.botFeedbackButton} ${botFeedbackByMessage[index] === 'like' ? styles.botFeedbackButtonActive : ''}`}
                          onClick={() => setBotFeedback(index, 'like')}
                        >
                          <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                            <path d="M8.1 10.6l3.2-7.1a2.2 2.2 0 0 1 4.2.9v5h3.2a2.4 2.4 0 0 1 2.34 2.92l-1.45 6.55A3.1 3.1 0 0 1 16.57 21H8.1V10.6Z" fill="currentColor" opacity="0.22" />
                            <path d="M8.1 21V10.6l3.2-7.1a2.2 2.2 0 0 1 4.2.9v5h3.2a2.4 2.4 0 0 1 2.34 2.92l-1.45 6.55A3.1 3.1 0 0 1 16.57 21H8.1ZM4 10.8h2.6V21H4a2 2 0 0 1-2-2v-6.2a2 2 0 0 1 2-2Z" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                        </button>
                        <button
                          type="button"
                          aria-label="Dislike answer"
                          title="Dislike answer"
                          aria-pressed={botFeedbackByMessage[index] === 'dislike'}
                          className={`${styles.botFeedbackButton} ${botFeedbackByMessage[index] === 'dislike' ? styles.botFeedbackButtonActive : ''}`}
                          onClick={() => setBotFeedback(index, 'dislike')}
                        >
                          <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                            <path d="M8.1 13.4l3.2 7.1a2.2 2.2 0 0 0 4.2-.9v-5h3.2a2.4 2.4 0 0 0 2.34-2.92l-1.45-6.55A3.1 3.1 0 0 0 16.57 3H8.1v10.4Z" fill="currentColor" opacity="0.22" />
                            <path d="M8.1 3v10.4l3.2 7.1a2.2 2.2 0 0 0 4.2-.9v-5h3.2a2.4 2.4 0 0 0 2.34-2.92L19.59 5.13A3.1 3.1 0 0 0 16.57 3H8.1ZM4 13.2h2.6V3H4a2 2 0 0 0-2 2v6.2a2 2 0 0 0 2 2Z" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" strokeLinejoin="round" />
                          </svg>
                        </button>
                        <button
                          type="button"
                          aria-label={botFeedbackByMessage[index] === 'copied' ? "Answer copied" : "Copy answer"}
                          title={botFeedbackByMessage[index] === 'copied' ? "Answer copied" : "Copy answer"}
                          className={`${styles.botFeedbackButton} ${botFeedbackByMessage[index] === 'copied' ? styles.botFeedbackButtonCopied : ''}`}
                          onClick={() => {
                            if (navigator.clipboard) {
                              void navigator.clipboard.writeText(msg.text);
                            }
                            setBotFeedback(index, 'copied');
                          }}
                        >
                          {botFeedbackByMessage[index] === 'copied' ? (
                            <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                              <circle cx="12" cy="12" r="9" fill="currentColor" opacity="0.18" />
                              <path d="M17.8 8.3l-7.1 7.4-3.5-3.4" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" />
                            </svg>
                          ) : (
                            <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                              <rect x="8.5" y="8.5" width="11" height="11" rx="2.4" fill="currentColor" opacity="0.16" />
                              <rect x="8.5" y="8.5" width="11" height="11" rx="2.4" stroke="currentColor" strokeWidth="1.75" />
                              <path d="M5.5 15.5H4.6A2.1 2.1 0 0 1 2.5 13.4V4.6A2.1 2.1 0 0 1 4.6 2.5h8.8a2.1 2.1 0 0 1 2.1 2.1v.9" stroke="currentColor" strokeWidth="1.75" strokeLinecap="round" />
                            </svg>
                          )}
                        </button>
                      </div>
                    </div>
                  )}
                </div>
                )}
              </div>
            ))}
            <div ref={messagesEndRef} />
          </div>
          <div className={styles.inputRow}>
            <input
              type="text"
              className={styles.chatInput}
              placeholder="Type your question..."
              value={chatInput}
              onChange={(e) => setChatInput(e.target.value)}
              onKeyPress={handleChatInputKeyPress}
              disabled={isProcessing}
            />
            <button
              className={styles.sendButton}
              onClick={handleSendMessage}
              disabled={isProcessing || !chatInput.trim() || chatInput.trim().length > CHAT_INPUT_MAX_CHARS}
              aria-label="Send message"
              title="Send message"
            >
              <svg viewBox="0 0 24 24" fill="none" aria-hidden="true">
                <path d="M4 5.5l16 6.5-16 6.5 3.3-6.5L4 5.5Z" stroke="currentColor" strokeWidth="2.1" strokeLinecap="round" strokeLinejoin="round" />
                <path d="M7.4 12H20" stroke="currentColor" strokeWidth="2.1" strokeLinecap="round" />
              </svg>
            </button>
          </div>
          {chatInput.trim().length > CHAT_INPUT_MAX_CHARS && (
            <div className={styles.chatInputLimitWarning}>
              Message limit exceeded. Keep it under {CHAT_INPUT_MAX_CHARS.toLocaleString()} characters, or paste only the key paragraph or title.
            </div>
          )}
            </div>
            </>
          )}
        </div>,
        document.body
      )}

      {showViewAll && props.context && (
        <div className={styles.viewAllModal}>
          <ViewAllDocumentsPage
            context={props.context}
            onClose={handleCloseViewAll}
            onViewDocument={handleViewDocumentFromList}
          />
        </div>
      )}

      {selectedDocumentId && props.context && (
        <div className={styles.detailModal}>
          <DocumentDetailPage
            context={props.context}
            documentId={selectedDocumentId}
            listTitle={selectedListTitle || LIBRARY_NAMES.kmDataHub}
            listId={selectedListId}
            onClose={handleCloseDetail}
            backTo={(() => {
              // Retrieve backTo from sessionStorage if available
              const storedBackTo = sessionStorage.getItem(`documentBackTo_${selectedDocumentId}`);
              return storedBackTo === 'library' ? 'library' : 'home';
            })()}
            onBackToLibrary={handleBackToLibrary}
          />
        </div>
      )}
    </>
  );
};

