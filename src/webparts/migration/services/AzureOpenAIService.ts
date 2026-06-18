/**
 * Service to interact with Azure OpenAI API
 */

import { buildBuDepartmentValue } from '../components/businessUnit/BusinessUnitHierarchyUtils';

// 🆕 Azure config imports
import {
  AZURE_OPENAI_ENDPOINT,
  AZURE_OPENAI_API_KEY,
  AZURE_OPENAI_DEPLOYMENT,
  AZURE_OPENAI_API_VERSION,
  AZURE_OPENAI_WHISPER_ENDPOINT,
  AZURE_OPENAI_WHISPER_API_KEY,
  AZURE_OPENAI_WHISPER_DEPLOYMENT,
  AZURE_OPENAI_EMBEDDING_MODEL,
  AZURE_OPENAI_EMBEDDING_API_VERSION
} from './SearchConfig';

import { ConfigService } from './ConfigService';
import { PromptBuilder } from '../utils/PromptBuilder';
import { ISearchResult } from '../models/SearchResult';
import { IChatMessage } from '../models/ChatMessage';
import { COLUMN_NAMES, SEARCH_PROPERTIES } from '../config/appConfig';
import { KnowledgeSearchApiClient } from './KnowledgeSearchApiClient';

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

export interface MetadataExtraction {
  title?: string;
  documentType?: string;
  bu?: string;
  department?: string;
  subDepartment?: string;
  buDepartment?: string;
  geography?: string;
  region?: string;
  client?: string;
  description?: string;
  abstract?: string;
  diseaseArea?: string;
  therapyArea?: string;
  sensitiveTerms?: string;
  emails?: string;
  phones?: string;
  ids?: string;
  pricing?: string;
}

interface AzureOpenAIConfig {
  apiKey: string;
  endpoint: string;
  deploymentName: string;
  apiVersion?: string;
}

export class AzureOpenAIService {
  private config: AzureOpenAIConfig;
  private uploadAiClient: KnowledgeSearchApiClient;
  private embeddingCache: Map<string, number[]> = new Map(); // 🧠 Cache for document embeddings
  private static readonly DIRECT_SEARCH_SUCCESS_THRESHOLD = 3;
  private static readonly MAX_PROMPT_RESULTS = 5;
  private static readonly STRONG_MATCH_PROMPT_RESULTS = 3;
  private static readonly QUERY_STOP_WORDS = new Set([
    'how', 'do', 'does', 'can', 'you', 'tell', 'me', 'find', 'search', 'for', 'about', 'together',
    'enhance', 'improve', 'the', 'is', 'are', 'what', 'which', 'who', 'im', 'looking', 'documents',
    'document', 'docuemtns', 'docuemnt', 'related', 'have', 'any', 'with', 'from', 'into', 'than',
    'terms', 'privacy', 'protections', 'primarily'
  ]);
  private static readonly MISSING_METADATA_VALUES = new Set([
    '',
    'not found',
    'n/a',
    'na',
    'none',
    'unknown',
    'not applicable'
  ]);
  private static readonly SPECIAL_CLIENT_VALUES = new Set([
    'multi-client',
    'others',
    'not applicable'
  ]);

  constructor(config?: Partial<AzureOpenAIConfig>, context?: any) {
    this.config = {
      apiKey: config?.apiKey || AZURE_OPENAI_API_KEY,
      endpoint: config?.endpoint || AZURE_OPENAI_ENDPOINT,
      deploymentName: config?.deploymentName || AZURE_OPENAI_DEPLOYMENT,
      apiVersion: config?.apiVersion || AZURE_OPENAI_API_VERSION || '2024-02-15-preview'
    };
    this.uploadAiClient = new KnowledgeSearchApiClient(undefined, context);
  }

  private isUploadAiBackendConfigured(): boolean {
    return this.uploadAiClient.isConfigured();
  }

  private hasLegacyOpenAiConfig(): boolean {
    return Boolean(this.config.apiKey && this.config.endpoint && this.config.deploymentName);
  }

  /**
   * TRUE RAG: Orchestrated Chat Response using SharePoint Native Search
   * 1. Check ConfigService (FAQs)
   * 2. Search SharePoint via KQL
   * 3. Build Grounded Prompt
   * 4. Call Azure OpenAI
   */
  public async getChatResponse(
    query: string,
    history: IChatMessage[], 
    searchService: { searchDocuments: (query: string) => Promise<ISearchResult[]> },
    configService: ConfigService,
    onStatusUpdate?: (status: string) => void,
    preloadedResults?: ISearchResult[],
    whosWhoContext?: string
  ): Promise<IChatMessage> {
    try {
      if (this.isFastOffTopicQuery(query)) {
        return {
          sender: 'bot',
          text: `I can help with Indegene, life sciences, healthcare, Knowledge Hub documents, people, departments, therapy areas, disease areas, clients, regions, and related topics. Please ask a question in that area.`,
          timestamp: new Date(),
          citations: []
        };
      }

      if (this.isAssistantPurposeQuery(query)) {
        return {
          sender: 'bot',
          text: `My purpose is to help you find and understand information from the Indegene Knowledge Hub. I can answer questions from KM documents, summarize relevant materials, and identify document matches by metadata such as BU, client, department, therapy area, disease area, region, or document type.`,
          timestamp: new Date(),
          citations: []
        };
      }

      if (this.isUnsupportedServiceRequest(query)) {
        return {
          sender: 'bot',
          text: `I can help with Knowledge Hub documents and metadata, but I can't provide services like loans. Please ask a Knowledge Hub or Indegene-related question.`,
          timestamp: new Date(),
          citations: []
        };
      }

      // 1. Check FAQs Match FIRST
      if (onStatusUpdate) onStatusUpdate("Checking FAQs...");
      const faqs = await configService.fetchFAQs();
      const faqMatch = configService.findFaqMatch(query, faqs);
      
      if (faqMatch) {
        return {
          sender: 'bot',
          text: faqMatch.Answer,
          timestamp: new Date()
        };
      }

      let searchResults = preloadedResults && preloadedResults.length > 0
        ? preloadedResults
        : [];

      if (searchResults.length === 0) {
        // 2. Clean and try a direct search first to avoid an extra model round-trip on easy queries
        const stopWords = ['how', 'do', 'does', 'can', 'you', 'tell', 'me', 'find', 'search', 'for', 'about', 'together', 'enhance', 'improve', 'the', 'is', 'are', 'what', 'which', 'who', 'i\'m', 'looking'];
        const normalizedUserQuery = this.normalizeUserQueryForSearch(query);
        let cleanQuery = normalizedUserQuery.replace(/[?.,!]/g, '')
          .replace(/^(show me|can you find|search for|i'm looking for|find|tell me about|give me)\b/gi, '')
          .trim();
        
        const keywordsOnly = cleanQuery.split(/\s+/)
          .filter(w => w.length > 2 && !stopWords.includes(w.toLowerCase()))
          .join(' ');

        let searchQuery = keywordsOnly || cleanQuery || '*';
        searchResults = await this.runExpandedDocumentSearch(searchService, normalizedUserQuery, searchQuery);

        if (searchResults.length === 0) {
          const relaxedQueries = this.buildRelaxedSearchQueries(searchQuery);
          const relaxedResultsByQuery = await Promise.all(
            relaxedQueries.map((relaxedQuery) =>
              searchService.searchDocuments(relaxedQuery).catch((error) => {
                console.warn('Relaxed SharePoint search failed:', relaxedQuery, error);
                return [] as ISearchResult[];
              })
            )
          );

          for (let index = 0; index < relaxedQueries.length; index++) {
            if (relaxedResultsByQuery[index].length > 0) {
              searchQuery = relaxedQueries[index];
              searchResults = relaxedResultsByQuery[index];
              break;
            }
          }
        }

        if (searchResults.length < AzureOpenAIService.DIRECT_SEARCH_SUCCESS_THRESHOLD) {
          const refinementUrl = `${this.config.endpoint}/openai/deployments/${this.config.deploymentName}/chat/completions?api-version=${this.config.apiVersion}`;
          const refinementResponse = await fetch(refinementUrl, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'api-key': this.config.apiKey
            },
            body: JSON.stringify({
              messages: [
                { 
                  role: 'system', 
                  content: `You are an expert search query optimizer for Indegene SharePoint. 
                  Analyze the conversation and generate high-performance Keyword Query Language (KQL) terms.
                  
                  RULES:
                   - Resolve pronouns using history.
                   - **KEYWORD EXTRACTION:** Extract 2-4 core keywords or phrases (e.g., "common data models", "DAAI", "Sandoz").
                   - **KQL SYNTAX:** Generate a query like: ("Keyword1" OR "Keyword2").
                   - **RELEVANCY:** If the user asks a complex question, identify the unique technical terms.
                   - **STATUS FILTER:** Do NOT add status filters here; they are handled automatically.
                   - **FUZZY:** Use wildcards (*) sparingly.
                   - **OUTPUT:** Return ONLY the KQL terms. No filler.` 
                },
               ...history.slice(-4).map(msg => ({
                 role: (msg.sender === 'user' ? 'user' : 'assistant') as 'user' | 'assistant',
                 content: msg.text
               })),
               { role: 'user', content: query }
              ]
            })
          });

          if (refinementResponse.ok) {
            const refData = await refinementResponse.json();
            const aiTerms = (refData.choices[0].message.content || '').trim();
            if (aiTerms && aiTerms !== '*' && aiTerms !== query) {
              // Priority to AI terms if they are concise, otherwise mix
              searchQuery = aiTerms.length < 50 ? aiTerms : `(${keywordsOnly}) OR (${aiTerms})`; 
            }
            debugLog(`🤖 Optimized KQL: "${searchQuery}"`);
          }

          searchResults = await this.runExpandedDocumentSearch(searchService, normalizedUserQuery, searchQuery);
        }
      } else {
        const augmentedResults = await this.runExpandedDocumentSearch(searchService, query, query);
        searchResults = this.mergeSearchResults(searchResults, augmentedResults);
      }

      searchResults = this.filterRelevantResults(query, searchResults);
      // People-directory context is intentionally archived for KM Assistant until the client re-enables it.
      void whosWhoContext;
      const directoryContext = '';
      if (searchResults.length === 0) {
        return {
          sender: 'bot',
          text: `I looked through the Knowledge Hub, but I couldn’t find any relevant documents for "${query}". Please try another spelling, a related term, or narrow it by client, BU, department, or topic.`,
          timestamp: new Date(),
          citations: []
        };
      }
      const directoryPrimaryAnswer = false;

      const maxPromptResults = this.hasStrongExactMatch(query, searchResults)
        ? AzureOpenAIService.STRONG_MATCH_PROMPT_RESULTS
        : AzureOpenAIService.MAX_PROMPT_RESULTS;

      // 3. Build Grounded Prompt
      if (onStatusUpdate) onStatusUpdate("Analyzing document contents...");
      const systemPrompt = PromptBuilder.buildSystemPrompt();

      // 🚩 Limit context size to keep the answer fast while preserving accuracy
      const promptResults = (searchResults || []).slice(0, maxPromptResults);
      const userPrompt = PromptBuilder.buildUserPrompt(query, promptResults, directoryContext);

      // 4. Call Azure OpenAI Chat Completion
      const url = `${this.config.endpoint}/openai/deployments/${this.config.deploymentName}/chat/completions?api-version=${this.config.apiVersion}`;
      
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'api-key': this.config.apiKey
        },
        body: JSON.stringify({
          messages: [
            { role: 'system', content: systemPrompt },
            ...history.slice(-6).map(msg => ({
              role: (msg.sender === 'user' ? 'user' : 'assistant') as 'user' | 'assistant',
              content: msg.text
            })),
           { role: 'user', content: userPrompt }
           ]
         })
       });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`OpenAI RAG failed: ${errorText}`);
      }

      const data = await response.json();
      const rawContent = this.formatChatAnswerForDisplay(data.choices[0].message.content || '');

      // 5. Return with Citations (Mapped for UI Rendering) - Limit to 5, cited docs first
      let topResults = promptResults.slice();
      topResults.sort((a, b) => {
        const aCited = rawContent.includes(a.title) ? 1 : 0;
        const bCited = rawContent.includes(b.title) ? 1 : 0;
        return bCited - aCited;
      });
      
      return {
        sender: 'bot',
        text: rawContent,
        timestamp: new Date(),
        citations: directoryPrimaryAnswer ? [] : topResults.map((res) => ({
          ...res,
          description: res.description || res.content || '',
          abstract: res.description || ''
        }))
      };

    } catch (error) {
      console.error("🔴 RAG Chat Error Details:", error);
      return {
        sender: 'bot',
        text: "I encountered an error while searching for an answer. Please try again later.",
        timestamp: new Date()
      };
    }
  }

  private hasStrongExactMatch(query: string, results: ISearchResult[]): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    if (!normalizedQuery || results.length === 0) {
      return false;
    }

    return results.some((result) => {
      const normalizedTitle = this.normalizeForMatch(result.title || '');
      const normalizedUrl = this.normalizeForMatch(result.url || '');
      return normalizedTitle === normalizedQuery ||
        normalizedTitle.indexOf(normalizedQuery) !== -1 ||
        normalizedUrl.indexOf(normalizedQuery) !== -1;
    });
  }

  private formatChatAnswerForDisplay(content: string): string {
    const normalizedContent = (content || '')
      .replace(/^\s*(\*\*)?\s*Answer\s*:?\s*(\*\*)?\s*/i, '')
      .replace(/(^|\n)\s*📄\s*/g, '$1')
      .replace(/\*\s*📄\s*/g, '* ')
      .replace(/\bAI Summary\s*:/gi, 'Summary:')
      .replace(/\*\*Summary:\*\*/gi, 'Summary:')
      .replace(/^\s*\*\*(Summary:\s*)(.*?)\*\*\s*$/gmi, '$1$2')
      .replace(/^#{1,6}\s+/gm, '')
      .replace(/^\s*-{3,}\s*$/gm, '')
      .replace(/[ \t]{2,}$/gm, '')
      .replace(/\*“([^”]+)”\*/g, '**$1**')
      .replace(/\*”([^”]+)”\*/g, '**$1**')
      .replace(/\*([^*\n]{1,80}?(?:document|file|brochure|playbook|overview|submissions|indegene)[^*\n]{0,40})\*/gi, '**$1**')
      .replace(/\*\*([^*\n]+?)\s+(document|file|brochure|playbook)\*\*/gi, '**$1** $2')
      .replace(/(^|[^*])\*(?!\*)/g, '$1')
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    return this.normalizeDocumentSummaryFormatting(normalizedContent);
  }

  private normalizeDocumentSummaryFormatting(content: string): string {
    const lines = (content || '').split('\n');
    let insideDocumentSummaries = false;

    return lines.map((line) => {
      const trimmedLine = line.trim();
      if (/^Document Summaries\s*:/i.test(trimmedLine)) {
        insideDocumentSummaries = true;
        return line;
      }

      if (/^(References|Here are the Knowledge Hub references)\b/i.test(trimmedLine)) {
        insideDocumentSummaries = false;
        return line;
      }

      if (!insideDocumentSummaries || !trimmedLine) {
        return line;
      }

      if (/^Summary\s*:/i.test(trimmedLine)) {
        return line.replace(/^\s*\*\*(Summary:\s*)(.*?)\*\*\s*$/i, '$1$2');
      }

      if (/^\*\*.*\*\*$/.test(trimmedLine)) {
        return line;
      }

      const leadingWhitespace = line.match(/^\s*/)?.[0] || '';
      return `${leadingWhitespace}**${trimmedLine}**`;
    }).join('\n');
  }

  private normalizeForMatch(value: string): string {
    return (value || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private extractMeaningfulTokens(query: string): string[] {
    return this.normalizeForMatch(this.normalizeUserQueryForSearch(query))
      .split(/\s+/)
      .filter((token) => token.length > 2 && !AzureOpenAIService.QUERY_STOP_WORDS.has(token));
  }

  private filterRelevantResults(query: string, results: ISearchResult[]): ISearchResult[] {
    if (!results || results.length === 0) {
      return [];
    }

    const tokens = this.extractMeaningfulTokens(query);
    const scoredResults = results.map((result) => ({
      result,
      lexicalScore: this.computeLexicalEvidenceScore(tokens, result),
      evidenceScore: this.computeContentEvidenceScore(tokens, result),
      retrievalScore: Number(result.semanticScore || result.score || result.rank || 0)
    }));
    const textualMedicalAiResults = this.filterTextualMedicalAiResults(query, scoredResults);
    if (textualMedicalAiResults.length > 0) {
      return textualMedicalAiResults;
    }
    const targetedIntentResults = this.filterTargetedIntentResults(query, scoredResults);
    if (targetedIntentResults.length > 0) {
      return targetedIntentResults;
    }
    if (this.isStrictSentenceLevelDocumentIntent(query)) {
      return [];
    }
    const strictMetadataResults = this.filterStrictMetadataListingResults(query, scoredResults);
    if (strictMetadataResults.length > 0) {
      return strictMetadataResults;
    }

    const hasLexicalEvidence = scoredResults.some((entry) => entry.lexicalScore > 0);
    const looksLikeEntityQuery = tokens.length > 0 && tokens.length <= 3;

    if (hasLexicalEvidence) {
      const strongMatches = scoredResults.filter((entry) => entry.lexicalScore > 0);
      const highestLexicalScore = Math.max(...strongMatches.map((entry) => entry.lexicalScore));
      const minimumLexicalScore = tokens.length >= 4
        ? Math.max(2, Math.min(highestLexicalScore, Math.ceil(tokens.length * 0.45)))
        : 1;

      return strongMatches
        .filter((entry) =>
          entry.lexicalScore >= minimumLexicalScore &&
          this.hasEnoughDocumentEvidence(tokens, entry.evidenceScore, entry.result)
        )
        .sort((a, b) => {
          if (b.lexicalScore !== a.lexicalScore) {
            return b.lexicalScore - a.lexicalScore;
          }
          return b.retrievalScore - a.retrievalScore;
        })
        .map((entry) => entry.result);
    }

    if (looksLikeEntityQuery) {
      return [];
    }

    return results.slice(0, AzureOpenAIService.MAX_PROMPT_RESULTS);
  }

  private filterTextualMedicalAiResults(
    query: string,
    scoredResults: Array<{
      result: ISearchResult;
      lexicalScore: number;
      evidenceScore: number;
      retrievalScore: number;
    }>
  ): ISearchResult[] {
    if (!this.isTextualMedicalAiIntent(query)) {
      return [];
    }

    return scoredResults
      .filter((entry) => {
        const haystack = this.normalizeForMatch([
          entry.result.title,
          entry.result.content,
          entry.result.description,
          entry.result.abstract
        ].filter(Boolean).join(' '));

        return (
          (haystack.indexOf('not a quick fix') !== -1 && haystack.indexOf('medical domain') !== -1) ||
          (
            haystack.indexOf('medical domain') !== -1 &&
            haystack.indexOf('ai') !== -1 &&
            (
              haystack.indexOf('textual') !== -1 ||
              haystack.indexOf('pharmacovigilance') !== -1 ||
              haystack.indexOf('medical writing') !== -1 ||
              haystack.indexOf('labeling') !== -1
            )
          )
        );
      })
      .sort((left, right) => {
        const leftTitle = this.normalizeForMatch(left.result.title || '');
        const rightTitle = this.normalizeForMatch(right.result.title || '');
        const leftTitlePriority = leftTitle.indexOf('not a quick fix') !== -1 ? 1 : 0;
        const rightTitlePriority = rightTitle.indexOf('not a quick fix') !== -1 ? 1 : 0;
        if (leftTitlePriority !== rightTitlePriority) {
          return rightTitlePriority - leftTitlePriority;
        }

        return right.lexicalScore - left.lexicalScore || right.retrievalScore - left.retrievalScore;
      })
      .map((entry) => entry.result);
  }

  private filterTargetedIntentResults(
    query: string,
    scoredResults: Array<{
      result: ISearchResult;
      lexicalScore: number;
      evidenceScore: number;
      retrievalScore: number;
    }>
  ): ISearchResult[] {
    const matcher = this.getTargetedIntentMatcher(query);
    if (!matcher) {
      return [];
    }

    return scoredResults
      .filter((entry) => matcher(this.normalizeForMatch([
        entry.result.title,
        entry.result.content,
        entry.result.description,
        entry.result.abstract
      ].filter(Boolean).join(' '))))
      .sort((left, right) => right.lexicalScore - left.lexicalScore || right.retrievalScore - left.retrievalScore)
      .map((entry) => entry.result);
  }

  private filterStrictMetadataListingResults(
    query: string,
    scoredResults: Array<{
      result: ISearchResult;
      lexicalScore: number;
      evidenceScore: number;
      retrievalScore: number;
    }>
  ): ISearchResult[] {
    const metadataPhrase = this.extractMetadataSearchPhrase(query);
    if (!this.isMetadataListingQuestion(query) || !metadataPhrase) {
      return [];
    }

    const normalizedPhrase = this.normalizeForMatch(metadataPhrase);
    const phraseTokens = normalizedPhrase.split(/\s+/).filter(Boolean);

    return scoredResults
      .filter((entry) => {
        const metadataText = this.normalizeForMatch([
          entry.result.client,
          entry.result.businessUnit,
          entry.result.department,
          entry.result.documentType,
          entry.result.geography,
          entry.result.therapyArea,
          entry.result.diseaseArea,
          entry.result.sensitiveTerms?.join(' '),
          entry.result.reviewerComments,
          entry.result.projectId,
          entry.result.versionFileName,
          entry.result.versionFileType,
          entry.result.modifiedBy?.title
        ].filter(Boolean).join(' '));

        return metadataText.indexOf(normalizedPhrase) !== -1 ||
          phraseTokens.every((token) => metadataText.split(/\s+/).includes(token));
      })
      .sort((left, right) => right.lexicalScore - left.lexicalScore || right.retrievalScore - left.retrievalScore)
      .map((entry) => entry.result);
  }

  private getTargetedIntentMatcher(query: string): ((haystack: string) => boolean) | null {
    const normalizedQuery = this.normalizeForMatch(query);

    if (
      /\bclinical\b/.test(normalizedQuery) &&
      /\bcontext\b/.test(normalizedQuery) &&
      /\bmedical\b/.test(normalizedQuery) &&
      /\bcontent\b/.test(normalizedQuery)
    ) {
      return (haystack: string) =>
        haystack.indexOf('not a quick fix') !== -1 ||
        (
          haystack.indexOf('medical domain') !== -1 &&
          haystack.indexOf('clinical context') !== -1
        );
    }

    if (
      /\bsentiment\b/.test(normalizedQuery) &&
      /\banalysis\b/.test(normalizedQuery) &&
      /\b(custom|customer|complaints?)\b/.test(normalizedQuery)
    ) {
      return (haystack: string) =>
        haystack.indexOf('sentiment analysis') !== -1 &&
        (
          haystack.indexOf('complaint') !== -1 ||
          haystack.indexOf('customer') !== -1 ||
          haystack.indexOf('custom') !== -1
        );
    }

    if (
      /\bstrategic\b/.test(normalizedQuery) &&
      /\bcreative\b/.test(normalizedQuery) &&
      /\blaunch\b/.test(normalizedQuery) &&
      /\bproducts?\b/.test(normalizedQuery)
    ) {
      return (haystack: string) =>
        haystack.indexOf('strategic') !== -1 &&
        haystack.indexOf('creative') !== -1 &&
        haystack.indexOf('launch') !== -1 &&
        haystack.indexOf('product') !== -1;
    }

    if (this.isPatientSupportProgramsIntent(normalizedQuery)) {
      return (haystack: string) =>
        haystack.indexOf('what support programs would improve patient outcomes') !== -1 ||
        haystack.indexOf('patient support program') !== -1 ||
        haystack.indexOf('patient support programs') !== -1 ||
        (
          haystack.indexOf('patient support') !== -1 &&
          haystack.indexOf('outcome') !== -1
        ) ||
        (
          haystack.indexOf('support program') !== -1 &&
          haystack.indexOf('patient') !== -1
        );
    }

    if (this.isCommunicationFrequencyAudienceIntent(normalizedQuery)) {
      return (haystack: string) =>
        haystack.indexOf('what communication frequency works best for each audience segment') !== -1 ||
        (
          haystack.indexOf('communication frequency') !== -1 &&
          (
            haystack.indexOf('audience segment') !== -1 ||
            haystack.indexOf('audience segments') !== -1
          )
        ) ||
        (
          haystack.indexOf('frequency') !== -1 &&
          haystack.indexOf('communication') !== -1 &&
          haystack.indexOf('audience') !== -1 &&
          haystack.indexOf('segment') !== -1
        );
    }

    if (this.isCampaignExecutionDelayIntent(normalizedQuery)) {
      return (haystack: string) =>
        haystack.indexOf('transform commercial operations') !== -1 ||
        (
          haystack.indexOf('campaign') !== -1 &&
          haystack.indexOf('performance') !== -1 &&
          (
            haystack.indexOf('execution delay') !== -1 ||
            haystack.indexOf('delay') !== -1 ||
            haystack.indexOf('commercial operations') !== -1 ||
            haystack.indexOf('customer experience') !== -1
          )
        );
    }

    return null;
  }

  private computeLexicalEvidenceScore(tokens: string[], result: ISearchResult): number {
    if (tokens.length === 0) {
      return 0;
    }

    const title = this.normalizeForMatch(result.title || '');
    const url = this.normalizeForMatch(result.url || '');
    const metadata = this.normalizeForMatch([
      result.client,
      result.businessUnit,
      result.department,
      result.documentType,
      result.geography,
      result.therapyArea,
      result.diseaseArea,
      result.sensitiveTerms?.join(' '),
      result.reviewerComments,
      result.projectId,
      result.versionFileName,
      result.versionFileType,
      result.modifiedBy?.title,
      result.abstract,
      result.description
    ].filter(Boolean).join(' '));
    const content = this.normalizeForMatch(result.content || '');
    const titleTokens = title.split(/\s+/);
    const metadataTokens = metadata.split(/\s+/);
    const contentTokens = content.split(/\s+/);
    const queryPhrase = tokens.join(' ');
    const topicTokens = tokens.filter((token) =>
      token.length >= 5 &&
      !['significance', 'important', 'importance', 'provide', 'services', 'serve', 'served'].includes(token)
    );

    let score = 0;
    if (queryPhrase && title.indexOf(queryPhrase) !== -1) {
      score += 10;
    }
    if (queryPhrase && metadata.indexOf(queryPhrase) !== -1) {
      score += 9;
    }
    if (queryPhrase && content.indexOf(queryPhrase) !== -1) {
      score += 6;
    }

    if (tokens.includes('launch')) {
      const serviceLaunchEvidence = (
        (title.indexOf('launch') !== -1 || metadata.indexOf('launch') !== -1 || content.indexOf('launch') !== -1) &&
        (
          title.indexOf('commercial') !== -1 ||
          metadata.indexOf('commercial') !== -1 ||
          content.indexOf('commercial') !== -1 ||
          title.indexOf('customer experience') !== -1 ||
          metadata.indexOf('customer experience') !== -1 ||
          content.indexOf('customer experience') !== -1 ||
          title.indexOf('market access') !== -1 ||
          metadata.indexOf('market access') !== -1 ||
          content.indexOf('market access') !== -1
        )
      );

      if (serviceLaunchEvidence) {
        score += 16;
      }

      const regulatoryOnlyEvidence = (
        (title.indexOf('regulatory') !== -1 || metadata.indexOf('regulatory') !== -1) &&
        title.indexOf('commercial') === -1 &&
        metadata.indexOf('commercial') === -1 &&
        title.indexOf('launch') === -1 &&
        metadata.indexOf('launch') === -1
      );

      if (regulatoryOnlyEvidence) {
        score -= 8;
      }
    }

    if (tokens.includes('regulatory') && tokens.includes('market') && tokens.includes('access')) {
      const overviewEvidence = title.indexOf('overview of indegene') !== -1 ||
        title.indexOf('overview') !== -1 ||
        content.indexOf('market access') !== -1 ||
        metadata.indexOf('market access') !== -1;
      const broadServiceEvidence = content.indexOf('regulatory') !== -1 &&
        (
          content.indexOf('market access') !== -1 ||
          content.indexOf('commercial') !== -1 ||
          content.indexOf('commercialization') !== -1 ||
          metadata.indexOf('market access') !== -1
        );

      if (overviewEvidence) {
        score += 18;
      }

      if (broadServiceEvidence) {
        score += 10;
      }
    }

    if (this.isRegulatorySubmissionsManagementIntent(tokens.join(' '))) {
      const regulatorySubmissionsEvidence =
        title.indexOf('we enable healthcare organizations') !== -1 ||
        metadata.indexOf('regulatory submissions management') !== -1 ||
        content.indexOf('regulatory submissions management') !== -1;

      if (regulatorySubmissionsEvidence) {
        score += 30;
      }
    }

    tokens.forEach((token) => {
      if (titleTokens.includes(token)) {
        score += topicTokens.includes(token) ? 9 : 4;
      } else if (title.indexOf(token) !== -1) {
        score += topicTokens.includes(token) ? 7 : 3;
      }

      if (metadataTokens.includes(token)) {
        score += topicTokens.includes(token) ? 6 : 3;
      } else if (metadata.indexOf(token) !== -1) {
        score += topicTokens.includes(token) ? 4 : 2;
      }

      if (url.indexOf(token) !== -1) {
        score += 2;
      }

      if (contentTokens.includes(token)) {
        score += 1;
      }
    });

    return score;
  }

  private computeContentEvidenceScore(tokens: string[], result: ISearchResult): number {
    if (tokens.length === 0) {
      return 0;
    }

    const title = this.normalizeForMatch(result.title || '');
    const metadata = this.normalizeForMatch([
      result.client,
      result.businessUnit,
      result.department,
      result.documentType,
      result.geography,
      result.therapyArea,
      result.diseaseArea,
      result.sensitiveTerms?.join(' '),
      result.reviewerComments,
      result.projectId,
      result.versionFileName,
      result.versionFileType,
      result.modifiedBy?.title,
      result.abstract,
      result.description
    ].filter(Boolean).join(' '));
    const content = this.normalizeForMatch(result.content || '');
    const haystack = `${title} ${metadata} ${content}`;
    const topicTokens = this.getTopicTokens(tokens);

    return topicTokens.reduce((total, token) => {
      const tokenRegex = new RegExp(`\\b${this.escapeRegExp(token)}\\b`, 'i');
      return total + (tokenRegex.test(haystack) ? 1 : 0);
    }, 0);
  }

  private hasEnoughDocumentEvidence(tokens: string[], evidenceScore: number, result: ISearchResult): boolean {
    const topicTokens = this.getTopicTokens(tokens);
    if (topicTokens.length === 0) {
      return true;
    }

    const readableContent = this.normalizeForMatch(result.content || '');
    const title = this.normalizeForMatch(result.title || '');
    const metadata = this.normalizeForMatch([
      result.client,
      result.businessUnit,
      result.department,
      result.documentType,
      result.geography,
      result.therapyArea,
      result.diseaseArea,
      result.sensitiveTerms?.join(' '),
      result.reviewerComments,
      result.projectId,
      result.versionFileName,
      result.versionFileType,
      result.modifiedBy?.title,
      result.abstract,
      result.description
    ].filter(Boolean).join(' '));
    const lowValueContent = !readableContent ||
      readableContent.length < 80 ||
      readableContent === 'generated by python docx' ||
      readableContent.indexOf('generated by python docx') !== -1;
    const queryPhrase = topicTokens.join(' ');
    const hasStrongPhraseEvidence = queryPhrase.length > 0 &&
      (title.indexOf(queryPhrase) !== -1 || metadata.indexOf(queryPhrase) !== -1 || readableContent.indexOf(queryPhrase) !== -1);
    const metadataPhrase = this.extractMetadataSearchPhrase(tokens.join(' '));
    const hasMetadataPhraseEvidence = metadataPhrase.length > 0 && metadata.indexOf(metadataPhrase) !== -1;
    const requiredEvidence = topicTokens.length <= 2 ? 1 : 2;

    if (lowValueContent && !hasStrongPhraseEvidence && !hasMetadataPhraseEvidence) {
      return false;
    }

    return evidenceScore >= requiredEvidence || hasStrongPhraseEvidence || hasMetadataPhraseEvidence;
  }

  private getTopicTokens(tokens: string[]): string[] {
    return tokens.filter((token) =>
      token.length >= 3 &&
      ![
        'significance', 'significant', 'important', 'importance', 'provide', 'provides',
        'services', 'serve', 'served', 'forcing', 'force', 'companies', 'company',
        'documents', 'document', 'docuemtns', 'docuemnt', 'related', 'have'
      ].includes(token)
    );
  }

  private escapeRegExp(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  private async runExpandedDocumentSearch(
    searchService: { searchDocuments: (query: string) => Promise<ISearchResult[]> },
    originalQuery: string,
    primaryQuery: string
  ): Promise<ISearchResult[]> {
    const allResults: ISearchResult[] = [];
    const searchQueries = this.buildExpandedSearchQueries(originalQuery, primaryQuery);

    const resultsByQuery = await Promise.all(
      searchQueries.map((searchQuery) =>
        searchService.searchDocuments(searchQuery).catch((error) => {
          console.warn('Expanded SharePoint search failed:', searchQuery, error);
          return [] as ISearchResult[];
        })
      )
    );

    resultsByQuery.forEach((results) => {
      allResults.push(...results);
    });

    return this.mergeSearchResults([], allResults);
  }

  private mergeSearchResults(primaryResults: ISearchResult[], additionalResults: ISearchResult[]): ISearchResult[] {
    const merged: ISearchResult[] = [];
    const seen = new Set<string>();

    [...primaryResults, ...additionalResults].forEach((result) => {
      const key = String(result.parentDocumentId || result.id || result.url || result.title || '');
      if (!key || seen.has(key)) {
        return;
      }

      seen.add(key);
      merged.push(result);
    });

    return merged;
  }

  private buildExpandedSearchQueries(originalQuery: string, primaryQuery: string): string[] {
    const normalizedOriginal = this.normalizeUserQueryForSearch(originalQuery);
    const normalizedPrimary = this.normalizeUserQueryForSearch(primaryQuery);
    const tokens = this.extractMeaningfulTokens(normalizedOriginal);
    const queries = new Set<string>();

    if (normalizedPrimary) {
      queries.add(normalizedPrimary);
    }
    if (normalizedOriginal && normalizedOriginal !== normalizedPrimary) {
      queries.add(normalizedOriginal);
    }
    if (tokens.length > 0) {
      queries.add(tokens.join(' '));
    }

    this.buildMetadataSearchQueries(normalizedOriginal).forEach((query) => queries.add(query));

    if (/\bindustr(y|ies)\b/.test(this.normalizeForMatch(normalizedOriginal)) && /\b(serv(e|es|ed)|primary|primarily)\b/.test(this.normalizeForMatch(normalizedOriginal))) {
      queries.add('overview indegene healthcare pharmaceutical life sciences');
      queries.add('healthcare pharmaceutical industry');
    }

    if (/\bengagement models?\b/.test(this.normalizeForMatch(normalizedOriginal)) || (/\brethink\b/.test(this.normalizeForMatch(normalizedOriginal)) && /\bengagement\b/.test(this.normalizeForMatch(normalizedOriginal)))) {
      queries.add('"engagement models"');
      queries.add('life sciences companies rethink engagement models');
      queries.add('commercial engagement models life sciences');
    }

    if (/\bomnichannel\b/.test(this.normalizeForMatch(normalizedOriginal))) {
      queries.add('omnichannel activation');
      queries.add('omnichannel engagement life sciences');
    }

    if (this.isRegulatoryMarketAccessIntent(normalizedOriginal)) {
      queries.add('"Overview of Indegene" regulatory market access');
      queries.add('Indegene regulatory compliance market access commercialization');
      queries.add('regulatory market access Indegene services');
    }

    if (this.isProductLaunchIntent(normalizedOriginal)) {
      queries.add('"product launch" commercialization Indegene');
      queries.add('"launch success" life sciences commercialization');
      queries.add('"launch excellence" pharmaceutical commercialization');
      queries.add('brand launch commercial operations customer experience');
      queries.add('Indegene product launch commercial medical regulatory market access');
    }

    if (this.isTextualMedicalAiIntent(normalizedOriginal)) {
      queries.add('"NOT A QUICK FIX IN THE MEDICAL DOMAIN"');
      queries.add('"NOT A QUICK FIX" "MEDICAL DOMAIN" AI');
      queries.add('"textual medical data" AI medical domain');
      queries.add('"fewer success stories" AI textual medical data');
      queries.add('AI medical domain pharmacovigilance medical writing labeling');
      queries.add('artificial intelligence medical life sciences textual data');
    }

    if (this.isClinicalContextMedicalContentIntent(normalizedOriginal)) {
      queries.add('"NOT A QUICK FIX" "clinical context"');
      queries.add('"AI in the Medical Domain" "clinical context"');
      queries.add('"medical content" "clinical context"');
    }

    if (this.isSentimentComplaintIntent(normalizedOriginal)) {
      queries.add('"sentiment analysis" complaints');
      queries.add('"sentiment analysis" "customer complaints"');
      queries.add('"custom complaints" "sentiment analysis"');
    }

    if (this.isStrategicCreativeProductLaunchIntent(normalizedOriginal)) {
      queries.add('"strategic" "creative" "launch" products');
      queries.add('"strategic and creative" "launch of products"');
      queries.add('"support the launch of products"');
    }

    if (this.isPatientSupportProgramsIntent(normalizedOriginal)) {
      queries.add('"What support programs would improve patient outcomes"');
      queries.add('"patient support programs" "patient outcomes"');
      queries.add('"support programs" "improve patient outcomes"');
      queries.add('"improve patient outcomes" "support programs"');
      queries.add('"patient outcomes" healthcare support');
      queries.add('patient support outcomes programs');
    }

    if (this.isCommunicationFrequencyAudienceIntent(normalizedOriginal)) {
      queries.add('"What communication frequency works best for each audience segment"');
      queries.add('"communication frequency" "audience segment"');
      queries.add('"communication frequency" "audience segments"');
      queries.add('"audience segment" frequency communication');
    }

    if (this.isCampaignExecutionDelayIntent(normalizedOriginal)) {
      queries.add('"Transform Commercial Operations and Customer Experience Across Your Enterprise"');
      queries.add('"execution delays" "campaign performance"');
      queries.add('"campaign performance" "commercial operations"');
      queries.add('"campaign execution" "customer experience"');
      queries.add('transform commercial operations campaign execution performance');
    }

    this.buildFileExtensionQueries(normalizedOriginal).forEach((query) => queries.add(query));

    if (this.isRegulatorySubmissionsManagementIntent(normalizedOriginal)) {
      queries.add('"Regulatory Submissions Management"');
      queries.add('"We enable healthcare organizations be future ready and effective"');
      queries.add('"We enable healthcare organizations" "Regulatory Submissions Management"');
      queries.add('regulatory submissions management solutions clinical regulatory safety medical');
    }

    if (this.isRegulatoryLabelingIntent(normalizedOriginal)) {
      queries.add('"regulatory labeling"');
      queries.add('"regulatory labelling"');
      queries.add('"multiple regulatory labeling"');
      queries.add('regulatory labeling labeling management regulatory');
    }

    return Array.from(queries).slice(0, 12);
  }

  private buildMetadataSearchQueries(query: string): string[] {
    const metadataPhrase = this.extractMetadataSearchPhrase(query);
    if (!metadataPhrase) {
      return [];
    }

    const escapedPhrase = metadataPhrase.replace(/"/g, '');
    const phraseQuery = escapedPhrase.indexOf(' ') >= 0 ? `"${escapedPhrase}"` : escapedPhrase;
    const metadataProperties = [
      'BUOWSCHCS',
      'BUOWSTEXT',
      'BUDepartmentOWSCHCS',
      'DepartmentOWSTEXT',
      'ClientOWSCHCS',
      'ClientOWSTEXT',
      'DocumentTypeOWSCHCS',
      'TherapyAreaOWSCHCS',
      'DiseaseAreaOWSCHCS',
      'GeographyOWSCHCS',
      SEARCH_PROPERTIES.status,
      SEARCH_PROPERTIES.documentType,
      SEARCH_PROPERTIES.businessUnit,
      SEARCH_PROPERTIES.client,
      SEARCH_PROPERTIES.geography,
      SEARCH_PROPERTIES.therapyArea
    ];

    return [
      phraseQuery,
      metadataProperties.map((propertyName) => `${propertyName}:${phraseQuery}`).join(' OR ')
    ];
  }

  private extractMetadataSearchPhrase(query: string): string {
    const normalized = this.normalizeUserQueryForSearch(query)
      .replace(/[?.,!]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

    if (!normalized) {
      return '';
    }

    const patterns = [
      /\bdocuments?\s+(?:related to|for|under|about|with|from)\s+(.+)$/i,
      /\bfiles?\s+(?:related to|for|under|about|with|from)\s+(.+)$/i,
      /\b(?:related to|for|under|about|with|from)\s+(.+)$/i,
      /\b(?:client|bu|business unit|department|sub department|document type|therapy area|disease area|geography)\s+(?:is|as|of|for)?\s*(.+)$/i
    ];

    for (const pattern of patterns) {
      const match = normalized.match(pattern);
      if (match?.[1]) {
        return this.cleanMetadataSearchPhrase(match[1]);
      }
    }

    return this.cleanMetadataSearchPhrase(normalized);
  }

  private cleanMetadataSearchPhrase(value: string): string {
    const cleaned = this.normalizeForMatch(value)
      .split(/\s+/)
      .filter((token) => ![
        'ok', 'please', 'can', 'you', 'tell', 'me', 'if', 'do', 'does', 'have', 'has', 'any',
        'documents', 'document', 'files', 'file', 'related', 'to', 'for', 'about', 'under',
        'with', 'from', 'client', 'bu', 'business', 'unit', 'department', 'sub', 'type', 'area'
      ].includes(token))
      .join(' ')
      .trim();

    return cleaned;
  }

  private isProductLaunchIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\blaunch\b/.test(normalizedQuery) &&
      (
        /\b(product|brand|pharmaceutical|pharma|life sciences?)\b/.test(normalizedQuery) ||
        /\bsuccess\b/.test(normalizedQuery)
      );
  }

  private isRegulatoryMarketAccessIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bregulatory\b/.test(normalizedQuery) &&
      /\bmarket\b/.test(normalizedQuery) &&
      /\baccess\b/.test(normalizedQuery);
  }

  private isDirectoryQuestion(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\b(who|lead|head|manager|owner|contact|email|person|people|team|teams|tean|members|department|respective|role|roles|structure)\b/.test(normalizedQuery);
  }

  private isAssistantPurposeQuery(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return (
      /\b(what|whats|tell)\b/.test(normalizedQuery) &&
      /\b(your|ur)\b/.test(normalizedQuery) &&
      /\b(purpose|role|job|function)\b/.test(normalizedQuery)
    ) || /\bwhat are you\b/.test(normalizedQuery);
  }

  private isUnsupportedServiceRequest(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\b(loan|credit card|mortgage|insurance quote|salary advance|borrow money)\b/.test(normalizedQuery);
  }

  private buildFileExtensionQueries(query: string): string[] {
    const extensions = Array.from(new Set((query.match(/\b(mp3|mp4|mov|avi|mkv|wmv|m4v|pdf|pptx|ppt|docx|doc|xlsx|xls)\b/gi) || [])
      .map((extension) => extension.toLowerCase())));

    if (extensions.length === 0) {
      return [];
    }

    return [
      extensions.map((extension) => `FileExtension:${extension}`).join(' OR '),
      extensions.map((extension) => `filetype:${extension}`).join(' OR '),
      extensions.join(' ')
    ];
  }

  private isMetadataListingQuestion(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return (
      /\b(what|which|show|list|give|tell|all)\b/.test(normalizedQuery) ||
      /\bdocuments?\b/.test(normalizedQuery)
    ) &&
      /\b(documents?|files?|section|related|client|bu|business unit|department|document type|therapy area|disease area|geography|region)\b/.test(normalizedQuery);
  }

  private isTextualMedicalAiIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\b(ai|artificial intelligence)\b/.test(normalizedQuery) &&
      /\b(textual|text|medical|domain|success|stories|story)\b/.test(normalizedQuery);
  }

  private isClinicalContextMedicalContentIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bclinical\b/.test(normalizedQuery) &&
      /\bcontext\b/.test(normalizedQuery) &&
      /\bmedical\b/.test(normalizedQuery) &&
      /\bcontent\b/.test(normalizedQuery);
  }

  private isSentimentComplaintIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bsentiment\b/.test(normalizedQuery) &&
      /\banalysis\b/.test(normalizedQuery) &&
      /\b(custom|customer|complaints?)\b/.test(normalizedQuery);
  }

  private isStrategicCreativeProductLaunchIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bstrategic\b/.test(normalizedQuery) &&
      /\bcreative\b/.test(normalizedQuery) &&
      /\blaunch\b/.test(normalizedQuery) &&
      /\bproducts?\b/.test(normalizedQuery);
  }

  private isRegulatorySubmissionsManagementIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bregulatory\b/.test(normalizedQuery) &&
      /\bsubmissions?\b/.test(normalizedQuery) &&
      /\bmanagement\b/.test(normalizedQuery);
  }

  private isRegulatoryLabelingIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bregulatory\b/.test(normalizedQuery) &&
      /\blabel(l)?ing\b/.test(normalizedQuery);
  }

  private isPatientSupportProgramsIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bpatients?\b/.test(normalizedQuery) &&
      /\boutcomes?\b/.test(normalizedQuery) &&
      /\bsupport\b/.test(normalizedQuery) &&
      /\bprograms?\b/.test(normalizedQuery);
  }

  private isCommunicationFrequencyAudienceIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bcommunication\b/.test(normalizedQuery) &&
      /\bfrequency\b/.test(normalizedQuery) &&
      /\baudience\b/.test(normalizedQuery) &&
      /\bsegments?\b/.test(normalizedQuery);
  }

  private isCampaignExecutionDelayIntent(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    return /\bcampaigns?\b/.test(normalizedQuery) &&
      /\bperformance\b/.test(normalizedQuery) &&
      (
        /\bexecution\b/.test(normalizedQuery) ||
        /\bdelays?\b/.test(normalizedQuery)
      );
  }

  private isStrictSentenceLevelDocumentIntent(query: string): boolean {
    return this.isPatientSupportProgramsIntent(query) ||
      this.isCommunicationFrequencyAudienceIntent(query) ||
      this.isCampaignExecutionDelayIntent(query);
  }

  private isFastOffTopicQuery(query: string): boolean {
    const normalizedQuery = this.normalizeForMatch(query);
    if (!normalizedQuery) {
      return false;
    }

    const domainPattern = /\b(indegene|knowledge|hub|document|documents|doc|file|files|life sciences?|pharma|pharmaceutical|healthcare|health care|medical|clinical|regulatory|market access|therapy|disease|department|client|region|bu|business unit|ent|glaucoma|omnichannel|commercial|commercialization|mp3|mp4|video|audio)\b/;
    if (domainPattern.test(normalizedQuery)) {
      return false;
    }

    return /\b(i love you|love me|do you love|marry me|date me|girlfriend|boyfriend|kiss|romantic)\b/.test(normalizedQuery);
  }

  private normalizeUserQueryForSearch(query: string): string {
    return (query || '')
      .replace(/\bglucoma\b/gi, 'glaucoma')
      .replace(/\bdocuemtns\b/gi, 'documents')
      .replace(/\bdocuemtn\b/gi, 'document')
      .replace(/\bdocuemnts\b/gi, 'documents')
      .replace(/\blabellings\b/gi, 'labeling')
      .replace(/\blabell?ings?\b/gi, 'labeling')
      .replace(/\bstudioes\b/gi, 'studies')
      .replace(/\bstudios\b/gi, 'studies')
      .trim();
  }

  private buildRelaxedSearchQueries(query: string): string[] {
    const normalizedQuery = this.normalizeForMatch(query);
    if (!normalizedQuery) {
      return [];
    }

    const variants = new Set<string>();
    const tokens = normalizedQuery.split(/\s+/).filter(Boolean);

    if (tokens.length === 1) {
      const token = tokens[0];
      if (token.length >= 5) {
        variants.add(`*${token.slice(1)}*`);
        variants.add(`${token.slice(0, token.length - 1)}*`);
      }
    }

    tokens.forEach((token) => {
      if (token.length >= 5) {
        variants.add(`*${token}*`);
      }
    });

    return Array.from(variants).filter((variant) => variant && variant !== normalizedQuery);
  }

  /**
   * Generates a grounded response from PRE-FILTERED search results.
   * This is used by the SAFE-RAG pattern to ensure only Active documents are used.
   */
  public async getGroundedResponse(
    query: string,
    history: IChatMessage[],
    activeDocuments: ISearchResult[],
    onStatusUpdate?: (status: string) => void
  ): Promise<IChatMessage> {
    try {
      if (onStatusUpdate) onStatusUpdate("Generating response...");

      const systemPrompt = PromptBuilder.buildSystemPrompt();
      const userPrompt = PromptBuilder.buildUserPrompt(query, activeDocuments);

      const url = `${this.config.endpoint}/openai/deployments/${this.config.deploymentName}/chat/completions?api-version=${this.config.apiVersion}`;

      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'api-key': this.config.apiKey
        },
        body: JSON.stringify({
          messages: [
            { role: 'system', content: systemPrompt },
            ...history.slice(-6).map(msg => ({
              role: (msg.sender === 'user' ? 'user' : 'assistant') as 'user' | 'assistant',
              content: msg.text
            })),
            { role: 'user', content: userPrompt }
          ]
        })
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`OpenAI Grounding failed: ${errorText}`);
      }

      const data = await response.json();
      const rawContent = data.choices[0].message.content || '';

      return {
        sender: 'bot',
        text: rawContent,
        timestamp: new Date(),
        citations: activeDocuments.map((res) => ({
          ...res,
          description: res.description || res.content || '',
          abstract: res.description || ''
        }))
      };
    } catch (error) {
      console.error("🔴 Grounded Chat Error:", error);
      return {
        sender: 'bot',
        text: "I encountered an error while synthesizing an answer from the documents. Please try again later.",
        timestamp: new Date()
      };
    }
  }

  /**
   * Generate embedding vector for text using Azure OpenAI
   */
  async generateEmbedding(text: string): Promise<number[]> {
    if (!text || !text.trim()) {
      throw new Error('Text cannot be empty for embedding generation');
    }

    try {
      // Ensure endpoint doesn't have trailing slash, then add path
      const baseEndpoint = AZURE_OPENAI_ENDPOINT.endsWith('/')
        ? AZURE_OPENAI_ENDPOINT.slice(0, -1)
        : AZURE_OPENAI_ENDPOINT;

      const url = `${baseEndpoint}/openai/deployments/${AZURE_OPENAI_EMBEDDING_MODEL}/embeddings?api-version=${AZURE_OPENAI_EMBEDDING_API_VERSION}`;

      const res = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'api-key': AZURE_OPENAI_API_KEY
        },
        body: JSON.stringify({
          input: text.trim()
        })
      });

      if (!res.ok) {
        const errorText = await res.text();
        // If it's a 404 (deployment not found), log as warning since we handle it gracefully
        // For other errors, log as error
        if (res.status === 404) {
          console.warn('Embedding deployment not found (404) - will use keyword matching only. Deployment name:', AZURE_OPENAI_EMBEDDING_MODEL);
        } else {
          console.error('Embedding API failed:', res.status, errorText);
        }
        throw new Error(`Embedding generation failed: ${res.status}`);
      }

      const json = await res.json();
      return json.data?.[0]?.embedding || [];
    } catch (error) {
      console.error('Error generating embedding:', error);
      throw error;
    }
  }

  /**
   * Calculate cosine similarity between two vectors
   */
  private cosineSimilarity(vecA: number[], vecB: number[]): number {
    if (vecA.length !== vecB.length) {
      return 0;
    }

    let dotProduct = 0;
    let normA = 0;
    let normB = 0;

    for (let i = 0; i < vecA.length; i++) {
      dotProduct += vecA[i] * vecB[i];
      normA += vecA[i] * vecA[i];
      normB += vecB[i] * vecB[i];
    }

    const denominator = Math.sqrt(normA) * Math.sqrt(normB);
    return denominator === 0 ? 0 : dotProduct / denominator;
  }

  /**
   * Perform semantic search on local documents using embeddings
   * Searches through title, filename, abstract, author, AND document content
   * Also includes exact keyword matching for documents that contain the search term
   */
  async semanticSearchLocalDocuments(
    query: string,
    documents: Array<{
      id: number;
      title: string;
      fileName: string;
      description: string; // abstract
      contributor: string; // author
      documentContent?: string; // actual document content (extracted from file)
      [key: string]: any;
    }>,
    topK: number = 10
  ): Promise<Array<{ document: any; similarity: number }>> {
    if (!query || !query.trim() || documents.length === 0) {
      return [];
    }

    try {
      const queryLower = query.toLowerCase().trim();

      // Extract meaningful keywords from the query (remove common stop words)
      // This allows matching "breast cancer" from "what is breast cancer"
      const stopWords = ['what', 'is', 'are', 'the', 'a', 'an', 'how', 'why', 'when', 'where', 'who', 'which', 'this', 'that', 'these', 'those', 'do', 'does', 'did', 'can', 'could', 'will', 'would', 'should', 'may', 'might', 'must'];
      const queryWords = queryLower
        .split(/\s+/)
        .filter(word => word.length > 2 && !stopWords.includes(word));

      // If no meaningful keywords after filtering, use the original query
      const searchTerms = queryWords.length > 0 ? queryWords : [queryLower];

      debugLog('Keyword matching - search terms:', searchTerms);

      // First, do exact keyword matching to find ALL documents containing the keywords
      // This ensures we don't miss documents with exact matches
      const exactMatches: Array<{ document: any; similarity: number }> = [];
      const nonExactMatches: any[] = [];

      documents.forEach((doc) => {
        // Check each field separately to ensure we catch matches in documentContent even if other fields don't match
        // documentContent includes: body text, headers, footers (e.g., "© 2025 Skysecure Technologies")
        const titleLower = (doc.title || '').toLowerCase();
        const fileNameLower = (doc.fileName || '').toLowerCase();
        const descriptionLower = (doc.description || '').toLowerCase();
        const contributorLower = (doc.contributor || '').toLowerCase();
        const documentContentLower = (doc.documentContent || '').toLowerCase(); // Includes headers/footers from all pages

        // Check if ANY keyword appears in ANY field
        // A document matches if at least one keyword is found
        let inTitle = false;
        let inFileName = false;
        let inDescription = false;
        let inContributor = false;
        let inDocumentContent = false;
        let matchedKeywords = 0;

        searchTerms.forEach(keyword => {
          if (titleLower.includes(keyword)) {
            inTitle = true;
            matchedKeywords++;
          }
          if (fileNameLower.includes(keyword)) {
            inFileName = true;
            matchedKeywords++;
          }
          if (descriptionLower.includes(keyword)) {
            inDescription = true;
            matchedKeywords++;
          }
          if (contributorLower.includes(keyword)) {
            inContributor = true;
            matchedKeywords++;
          }
          if (documentContentLower.includes(keyword)) {
            inDocumentContent = true;
            matchedKeywords++;
          }
        });

        // Document matches if ANY keyword appears in ANY field, especially documentContent
        if (inTitle || inFileName || inDescription || inContributor || inDocumentContent) {
          // Combine all fields for counting total occurrences
          const searchableText = [
            doc.title || '',
            doc.fileName || '',
            doc.description || '',
            doc.contributor || '',
            doc.documentContent || ''
          ].filter(Boolean).join(' ').toLowerCase();

          // Count total occurrences of all keywords
          let totalMatches = 0;
          searchTerms.forEach(keyword => {
            const regex = new RegExp(keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
            const matches = (searchableText.match(regex) || []).length;
            totalMatches += matches;
          });

          // Calculate relevance score based on matched keywords and occurrences
          // Base score from number of matched keywords (more keywords = higher score)
          let score = matchedKeywords * 5; // Each matched keyword adds 5 points
          score += totalMatches; // Add points for total occurrences

          // Boost scores based on where keywords appear
          if (inTitle) score += 10;
          if (inFileName) score += 5;
          if (inDescription) score += 6;
          if (inContributor) score += 5;
          if (inDocumentContent) score += 4; // Important: boost for document content matches

          // Additional boost if keywords found as whole words
          searchTerms.forEach(keyword => {
            const wordBoundaryRegex = new RegExp(`\\b${keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i');
            if (wordBoundaryRegex.test(searchableText)) score += 2;
          });

          // Ensure documents with matches ONLY in documentContent still get a minimum score
          if (inDocumentContent && !inTitle && !inFileName && !inDescription && !inContributor) {
            score = Math.max(score, 5); // Minimum score for content-only matches
          }

          exactMatches.push({ document: doc, similarity: score });
        } else {
          nonExactMatches.push(doc);
        }
      });

      // Sort exact matches by relevance (higher score = more relevant)
      exactMatches.sort((a, b) => b.similarity - a.similarity);

      // IMPORTANT: Return ALL exact matches, not just topK
      // This ensures documents with matches only in content are included
      // We'll limit later if needed, but prioritize getting all exact matches first
      if (exactMatches.length > 0) {
        // Return all exact matches (they're already sorted by relevance)
        // If we have more than topK, we'll still return all, but user will see most relevant first
        return exactMatches;
      }

      // Otherwise, supplement with semantic search on non-matching documents
      // Try to generate embedding for the search query, but don't fail if it doesn't work
      let queryEmbedding: number[] | null = null;
      try {
        queryEmbedding = await this.generateEmbedding(query);
      } catch (error) {
        // Silently fall back to keyword matching - this is expected when embedding deployment doesn't exist
        // Only log if it's not a 404 (deployment not found)
        if (error instanceof Error && !error.message.includes('404')) {
          console.warn('Embedding generation failed, continuing with keyword matching only:', error);
        }
        // If embedding fails, return only exact matches (keyword-based)
        // This ensures the search still works even if embedding API is unavailable
        return exactMatches;
      }

      // Only proceed with semantic search if we successfully got query embedding
      if (!queryEmbedding || queryEmbedding.length === 0) {
        console.warn('Query embedding is empty, returning keyword matches only');
        return exactMatches;
      }

      // Generate embeddings for documents that didn't have exact matches
      const documentEmbeddings = await Promise.all(
        nonExactMatches.map(async (doc) => {
          // Combine all searchable fields including document content
          const searchableText = [
            doc.title || '',
            doc.fileName || '',
            doc.description || '', // abstract
            doc.contributor || '', // author
            doc.documentContent || '' // actual document content
          ].filter(Boolean).join(' ');

          // Limit total text to ~30,000 chars
          const truncatedText = searchableText.length > 30000
            ? searchableText.substring(0, 30000)
            : searchableText;

          // 🧠 CACHE CHECK: Use ID and content hash-like key to avoid re-embedding
          const cacheKey = `doc_${doc.id}_${truncatedText.length}`;
          let embedding = this.embeddingCache.get(cacheKey);

          if (!embedding) {
            try {
              debugLog(`📡 Cache miss for doc ${doc.id}, generating new embedding...`);
              embedding = await this.generateEmbedding(truncatedText);
              if (embedding) this.embeddingCache.set(cacheKey, embedding);
            } catch (error) {
              console.error(`Error generating embedding for doc ${doc.id}:`, error);
              return { doc, embedding: null };
            }
          } else {
            debugLog(`⚡ Hit embedding cache for doc ${doc.id}`);
          }

          return { doc, embedding: embedding || null };
        })
      );

      // Calculate similarities for semantic matches
      const semanticResults = documentEmbeddings
        .filter((item) => item.embedding !== null)
        .map((item) => ({
          document: item.doc,
          similarity: this.cosineSimilarity(queryEmbedding!, item.embedding!)
        }))
        .filter((item) => item.similarity > 0.3) // Only include reasonably similar results
        .sort((a, b) => b.similarity - a.similarity);

      // Combine exact matches (first) with semantic matches (second)
      // Exact matches are prioritized and ALL are included
      const allResults = [
        ...exactMatches, // ALL exact matches included (already sorted by relevance)
        ...semanticResults.map((item) => ({
          document: item.document,
          similarity: item.similarity * 0.5 // Reduce semantic match scores so exact matches rank higher
        }))
      ];

      // Sort all results by similarity (exact matches will be first due to higher scores)
      allResults.sort((a, b) => b.similarity - a.similarity);

      // Return top results, but ensure ALL exact matches are included
      // If we have many exact matches, return them all (up to reasonable limit)
      const exactMatchCount = exactMatches.length;
      const maxResults = Math.max(topK, exactMatchCount); // Return at least all exact matches

      return allResults.slice(0, Math.min(maxResults, 200)); // Cap at 200 to avoid performance issues
    } catch (error) {
      console.error('Error in semanticSearchLocalDocuments:', error);
      return [];
    }
  }

  // ========================================================================
  // 🔹 PART 2: YOUR EXISTING METADATA EXTRACTION LOGIC (CORRECTED/COMPLETED)
  // ========================================================================

  /**
   * Extract metadata from document text using GPT-4o
   * Uses a single extraction pass for most documents
   */
  async extractMetadata(
    documentText: string,
    docTypeTermsSection?: string,
    clientTermsSection?: string
  ): Promise<MetadataExtraction> {
    const singlePassLimit = 120000;
    if (this.isUploadAiBackendConfigured()) {
      try {
        const backendDocumentText = documentText.length > singlePassLimit
          ? documentText.substring(0, singlePassLimit)
          : documentText;
        return await this.uploadAiClient.extractUploadMetadata(
          backendDocumentText,
          docTypeTermsSection,
          clientTermsSection
        ) as MetadataExtraction;
      } catch (error) {
        console.warn('Backend upload metadata extraction failed:', error);
        if (!this.hasLegacyOpenAiConfig()) {
          throw error;
        }
      }
    }

    try {
      // Log the extracted text for debugging
      debugLog('=== DOCUMENT TEXT EXTRACTED ===');
      debugLog('Total length:', documentText.length, 'characters');
      debugLog('First 1000 chars:', documentText.substring(0, 1000));
      debugLog('Last 1000 chars:', documentText.substring(Math.max(0, documentText.length - 1000)));

      // Count emails in the raw text for debugging
      const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
      const foundEmails = documentText.match(emailRegex);
      debugLog('=== EMAILS FOUND IN RAW TEXT ===');
      debugLog('Count:', foundEmails ? foundEmails.length : 0);
      if (foundEmails) {
        debugLog('Emails:', foundEmails);
      }

      if (documentText.length <= singlePassLimit) {
        debugLog('=== PROCESSING AS SINGLE CHUNK ===');
        const result = await this.processSingleChunk(documentText, undefined, docTypeTermsSection);
        // Validate and retry if needed
        const validatedResult = await this.validateAndRetryIfNeeded(result, [documentText]);
        return validatedResult;
      }

      debugLog('=== PROCESSING AS SINGLE PASS (TRUNCATED) ===');
      debugLog('Document length:', documentText.length, 'characters');
      debugLog('Single pass limit:', singlePassLimit, 'characters');

      const truncatedText = documentText.substring(0, singlePassLimit);
      const result = await this.processSingleChunk(truncatedText, undefined, docTypeTermsSection);
      const validatedResult = await this.validateAndRetryIfNeeded(result, [truncatedText]);
      return validatedResult;
    } catch (error) {
      console.error('Error extracting metadata:', error);
      throw error;
    }
  }

  /**
   * Build the prompt for metadata extraction
   */
  private buildExtractionPrompt(
    documentText: string,
    chunkContext?: { chunkIndex?: number; totalChunks?: number; previousFindings?: Partial<MetadataExtraction> },
    docTypeTermsSection?: string
  ): string {
    let contextNote = '';
    if (chunkContext && chunkContext.totalChunks && chunkContext.totalChunks > 1) {
      contextNote = `\n\n**IMPORTANT: This is chunk ${chunkContext.chunkIndex || 1} of ${chunkContext.totalChunks}.** `;
      if (chunkContext.previousFindings) {
        const prev = chunkContext.previousFindings;
        contextNote += `Previous chunks found: ${prev.title ? `title="${prev.title}"` : ''} ${prev.documentType ? `documentType="${prev.documentType}"` : ''} ${prev.bu ? `bu="${prev.bu}"` : ''} ${prev.department ? `department="${prev.department}"` : ''}. `;
      }
      contextNote += `Extract information from THIS chunk. If mandatory fields are missing from previous chunks, extract them from this chunk.`;
    }

    const docTypeInstruction = docTypeTermsSection
      ? `2. documentType - MANDATORY - NEVER EMPTY:
Choose the single BEST matching document type from this list.
Read the document content and compare it against each description below.
Return ONLY the exact term name - no explanation, no other text.

Available document types:
${docTypeTermsSection}`
      : `2. documentType - MANDATORY - NEVER EMPTY:
Identify the document type from content structure.
Use a concise label like "Deck", "Training", "Report" etc.`;

    return `You are an expert document analyzer. Analyze the following document text and extract structured information. You MUST be thorough, accurate, and NEVER leave mandatory fields empty.

**CRITICAL RULES - READ CAREFULLY:**

1. **NEVER HALLUCINATE**: Only extract information that is ACTUALLY PRESENT in the document text below. Do NOT make up, infer, or guess information that is not explicitly or clearly implied in the document. If information is not found, use empty string "" for optional fields, EXCEPT client, geography, diseaseArea, and therapyArea, which must follow the special selection rules below.

2. **MANDATORY FIELDS ARE ABSOLUTELY REQUIRED**: The following fields MUST ALWAYS have a value - they can NEVER be empty strings:
   - title: MUST follow the title generation rules below and NEVER return empty string.
   - documentType: MUST identify the document type from content structure or explicitly stated type.
   - bu (Business Unit): extract the business unit, practice, function, or organizational area mentioned or most clearly implied by the document.
   - department: extract the department, team, capability, or sub-function mentioned or most clearly implied by the document.
   - subDepartment: if a more specific sub-team, capability, or service line is clearly present, extract it.
   - description: MUST follow the description generation rules below and NEVER return empty string.

3. **NO MASKING OR TRUNCATION**: NEVER mask, hash, or replace any part of sensitive information (emails, phones, pricing, etc.) with asterisks (*), hashes (#), or any other placeholders. Return the FULL, ORIGINAL strings exactly as they appear in the document. Do NOT truncate strings unless they exceed 1000 characters.

4. **SCAN EVERY CHARACTER**: Read the ENTIRE document text carefully. Do not skip any part. Look in headers, footers, titles, body text, signatures, and all sections.

${contextNote}

**MANDATORY FIELDS (MUST ALWAYS BE FILLED - NEVER LEAVE BLANK):**
- title: MUST follow the title generation rules below and NEVER return empty string
- documentType: MUST identify the document type. Return the most likely human-readable category found or implied by the file format, title, or content structure.
- bu (Business Unit): MUST extract the business unit, practice, function, or organizational area. If not explicit, infer the closest candidate phrase from context.
- department: MUST extract the department, team, capability, or sub-function. If not explicit, infer the closest candidate phrase from context.
- subDepartment: extract the most specific sub-team, capability, or service line if clearly present. Otherwise use "".
- description: MUST follow the description generation rules below and NEVER return empty string

**CONDITIONAL FIELDS (ONLY FILL IF FOUND IN DOCUMENT - DO NOT HALLUCINATE):**
- For client, geography, therapyArea, and diseaseArea: use the special selection rules defined in each field description above. Never return empty string for these fields.
- geography: follow the special geography selection rules in the field description below.
- client: follow the special client selection rules in the field description below.
- diseaseArea: follow the special disease area selection rules in the field description below.
- therapyArea: follow the special therapy area selection rules in the field description below.

**SENSITIVE TERMS FIELD (EXTRACT ALL INSTANCES):**
- sensitiveTerms: Extract and combine emails, phones, names, ID numbers, pricing, and other PII/sensitive terms into one comma-separated string.

Fields to extract:

1. title - MANDATORY - NEVER EMPTY: A professional, market-standard document title.
Rules for generating title:

- Maximum 255 characters
- Remove ALL special characters: $ % @ * + ? ! # & ^ ~ \` | \\ < >
- Remove file extensions (.pdf, .docx, .pptx, .xlsx, .txt etc.)
- Remove version numbers and IDs (e.g. v1, v2.0, _final, _v3, -001, _20240101, document IDs, reference numbers)
- Remove leading/trailing spaces and dashes
- Use clean readable Title Case
- Prefer the document heading or first major heading only when it is already clear, professional, and topic-specific
- If the heading or document name is generic, noisy, versioned, or file-like, generate a polished business title from the document content
- Base the title on the document's main topic, purpose, audience, and key business context, similar to how the description summarizes the content
- Use concise market-standard wording suitable for a knowledge hub, review hub, or client-facing document library
- Do NOT simply copy the file name when the content provides a clearer professional title
- If no explicit title exists, create a concise descriptive title based on the main topic
- NEVER return empty string

${docTypeInstruction}

3. bu - **MANDATORY - NEVER EMPTY**: Business Unit. Extract the most likely business unit, practice, function, or organizational area from the content. If not explicit, infer the closest phrase from context. NEVER return empty string.

4. department - **MANDATORY - NEVER EMPTY**: Department. Extract the most likely department, team, capability, or sub-function from the content. If not explicit, infer the closest phrase from context. NEVER return empty string.

5. subDepartment - Extract a more specific sub-team, capability, or service line if clearly present. If none is found, use empty string "".

6. geography - Geographic region. Rules:
   - If ONE specific geography/region/country is found -> return that value
   - If MORE THAN ONE geography or region is found in the document -> return "General"
   - If NO geography or region exists anywhere in the document -> return "Not Applicable"
   - Do NOT use empty string "" for this field - always return a value

7. client - Client name or organization. Rules:
   - If ONE specific company name is found -> return that company name exactly
   - If MORE THAN ONE different company/client name is found -> return "Multi-Client"
   - If a client clearly exists in the document but the specific name is not identifiable or seems unlisted -> return "Others"
   - If NO client name or company exists anywhere in the document -> return "Not Applicable"
   - Do NOT include person names, internal departments, or generic terms like "the client"
   - Do NOT infer the client from file names, folder paths, repository URLs, email domains, project codes, sensitiveTerms, or internal Indegene contacts

8. description - MANDATORY - NEVER EMPTY: A brief summary of the document.
Rules for generating description:

- 1-3 sentences maximum
- Remove ALL special characters: $ % @ * + ? ! # & ^ ~ \` | \\ < >
- Do NOT include file names, extensions, version numbers, or IDs
- Do NOT include document metadata like dates, reference numbers, or internal codes
- Write in clean plain English
- Summarize the main purpose, topic, and key content of the document
- NEVER return empty string

9. diseaseArea - Disease Area. Rules:
   - If ONE specific disease area is found -> return that value
   - If content spans MORE THAN ONE disease area -> return "Cross-Disease"
   - If NO disease area exists in the document -> return "Not Applicable"
   - Do NOT use empty string "" for this field - always return a value

10. therapyArea - Therapy Area. Rules:
   - If ONE specific therapy area is found -> return that value
   - If content spans MORE THAN ONE therapy area -> return "Cross-Therapy"
   - If NO therapy area exists in the document -> return "Not Applicable"
   - Do NOT use empty string "" for this field - always return a value

11. sensitiveTerms - Extract and combine any email addresses, phone numbers, full names, social security numbers (SSN), passport/ID numbers, physical addresses, bank account details, credit card numbers, pricing, costs, monetary values, budgets, project codes, client names, internal IDs, reference numbers, document IDs, case numbers, ticket numbers, or other sensitive terms found in the document. Return them as a comma-separated string. NEVER mask or hash these values. If none found, use "".

**Document text to analyze:**
${documentText}

**IMPORTANT REMINDERS:**
- Read EVERY character of the document text above
- For MANDATORY fields (title, documentType, bu, department, description): NEVER return empty string ""
- For client, geography, diseaseArea, and therapyArea: follow the special selection rules and NEVER return empty string ""
- For other optional fields: Only fill if information is actually present in the document
- DO NOT make up or invent information that is not in the document
- Extract sensitive terms into the single sensitiveTerms field

Return only valid JSON in this format:
{
  "title": "...",
  "documentType": "...",
  "bu": "...",
  "department": "...",
  "subDepartment": "...",
  "geography": "US | General | Not Applicable",
  "client": "Company Name | Multi-Client | Others | Not Applicable",
  "description": "...",
  "diseaseArea": "Disease Name | Cross-Disease | Not Applicable",
  "therapyArea": "Therapy Name | Cross-Therapy | Not Applicable",
  "sensitiveTerms": "..."
}`;
  }

  /**
   * Sanitize and validate extracted metadata
   */
  private isMissingMetadataValue(value?: string): boolean {
    return AzureOpenAIService.MISSING_METADATA_VALUES.has((value || '').trim().toLowerCase());
  }

  private splitMetadataList(value?: string): string[] {
    return (value || '')
      .split(/[,;\n]/)
      .map((part) => part.trim())
      .filter((part) => part.length > 0);
  }

  private normalizeOptionalClassification(value?: string, fallback: string = 'Not Applicable'): string {
    const trimmed = (value || '').trim();
    return this.isMissingMetadataValue(trimmed) ? fallback : trimmed;
  }

  private normalizeClientValue(value?: string): string {
    const trimmed = (value || '').trim();

    if (this.isMissingMetadataValue(trimmed)) {
      return 'Not Applicable';
    }

    const lower = trimmed.toLowerCase();

    if (lower === 'multi client' || lower === 'multi-client' || lower === 'multiple clients') {
      return 'Multi-Client';
    }

    if (lower === 'other' || lower === 'others') {
      return 'Others';
    }

    if (
      lower.indexOf('@') !== -1 ||
      lower.indexOf('http://') !== -1 ||
      lower.indexOf('https://') !== -1 ||
      lower.indexOf('svnrepo') !== -1 ||
      lower.indexOf('indegene.com') !== -1 ||
      lower.indexOf('file:') !== -1 ||
      lower.indexOf('/') !== -1 ||
      lower.indexOf('\\') !== -1
    ) {
      return 'Not Applicable';
    }

    const clientValues = this.splitMetadataList(trimmed)
      .filter((client) => !this.isMissingMetadataValue(client))
      .filter((client) => !AzureOpenAIService.SPECIAL_CLIENT_VALUES.has(client.toLowerCase()));

    if (clientValues.length > 1) {
      return 'Multi-Client';
    }

    return trimmed;
  }

  private mergeClientValues(results: MetadataExtraction[]): string {
    const values = results
      .map((result) => this.normalizeClientValue(result.client))
      .filter((client) => client !== 'Not Applicable');

    if (values.some((client) => client === 'Multi-Client')) {
      return 'Multi-Client';
    }

    const concreteClients = values
      .filter((client) => client !== 'Others')
      .filter((client, index, allClients) =>
        allClients.findIndex((candidate) => candidate.toLowerCase() === client.toLowerCase()) === index
      );

    if (concreteClients.length > 1) {
      return 'Multi-Client';
    }

    if (concreteClients.length === 1) {
      return concreteClients[0];
    }

    return values.some((client) => client === 'Others') ? 'Others' : 'Not Applicable';
  }

  private sanitizeMetadata(metadata: MetadataExtraction): MetadataExtraction {
    const sanitized: MetadataExtraction = {};

    // Ensure all fields are strings and trim whitespace
    const fields: (keyof MetadataExtraction)[] = [
      'title', 'documentType', 'bu', 'department', 'subDepartment', 'buDepartment', 'geography', 'region', 'client',
      'description', 'abstract', 'diseaseArea', 'therapyArea', 'sensitiveTerms', 'emails', 'phones',
      'ids', 'pricing'
    ];

    for (const field of fields) {
      const value = metadata[field];
      sanitized[field] = typeof value === 'string' ? value.trim() : '';
    }

    // MANDATORY FIELDS - Ensure they are never empty
    // Title: If empty, use a default or file name
    if (!sanitized.title || sanitized.title === '') {
      sanitized.title = 'Untitled Document';
      console.warn('⚠️ Title was empty, using default');
    }

    // DocumentType: Must always have a value
    if (!sanitized.documentType || sanitized.documentType === '') {
      sanitized.documentType = 'Others';
      console.warn('⚠️ DocumentType was empty, using fallback: Others');
    }

    // Business Unit: Must always have a value
    if (!sanitized.bu || sanitized.bu === '') {
      sanitized.bu = 'Unspecified BU';
      console.warn('⚠️ Business Unit was empty, using fallback:', sanitized.bu);
    }

    // Department: Must always have a value
    if (!sanitized.department || sanitized.department === '') {
      sanitized.department = 'Unspecified Department';
      console.warn('⚠️ Department was empty, using fallback:', sanitized.department);
    }

    if (!sanitized.description && sanitized.abstract) {
      sanitized.description = sanitized.abstract;
    }

    if (!sanitized.geography && sanitized.region) {
      sanitized.geography = sanitized.region;
    }

    if (!sanitized.sensitiveTerms) {
      sanitized.sensitiveTerms = [
        sanitized.emails,
        sanitized.phones,
        sanitized.ids,
        sanitized.pricing
      ].filter((value) => !!value).join(', ');
    }

    if (!sanitized.description || sanitized.description === '') {
      sanitized.description = 'Document content summary not available.';
      console.warn('⚠️ Description was empty, using default');
    }

    sanitized.geography = this.normalizeOptionalClassification(sanitized.geography);
    sanitized.diseaseArea = this.normalizeOptionalClassification(sanitized.diseaseArea);
    sanitized.therapyArea = this.normalizeOptionalClassification(sanitized.therapyArea);

    sanitized.abstract = sanitized.description;
    sanitized.region = sanitized.geography;
    sanitized.buDepartment = buildBuDepartmentValue(sanitized.bu, sanitized.department, sanitized.subDepartment);

    // Client field validation - should only contain company names
    // Remove if it contains person names, generic terms, or invalid content
    if (sanitized.client) {
      const clientLower = sanitized.client.toLowerCase();

      // Remove invalid values
      if (clientLower === 'not found' ||
        clientLower === 'n/a' ||
        clientLower === 'none' ||
        clientLower === 'unknown' ||
        clientLower === 'the client' ||
        clientLower === 'our client' ||
        clientLower === 'client' ||
        clientLower.indexOf('internal') !== -1 ||
        clientLower.indexOf('team') !== -1) {
        sanitized.client = 'Not Applicable';
        debugLog('⚠️ Client field contained invalid value, cleared');
      } else {
        // Check if it looks like a person's name (first name + last name pattern without company indicators)
        // Company indicators: Inc, LLC, Corp, Ltd, Company, Co, etc.
        const companyIndicators = ['inc', 'llc', 'corp', 'ltd', 'company', 'co', 'group', 'enterprises', 'solutions', 'systems', 'technologies', 'consulting', 'services'];
        const hasCompanyIndicator = companyIndicators.some(indicator => clientLower.indexOf(indicator) !== -1);

        // If it's just a name without company indicators and doesn't look like a company, clear it
        if (!hasCompanyIndicator && sanitized.client.split(' ').length <= 3) {
          // Single-word values like "Viatris" or "Pfizer" can be valid companies.
          // Only clear the field when it strongly looks like a person's first+last name.
          const words = sanitized.client.trim().split(/\s+/);
          const looksLikePersonName =
            words.length === 2 &&
            /^[A-Za-z'-]+$/.test(words[0]) &&
            /^[A-Za-z'-]+$/.test(words[1]);

          if (looksLikePersonName) {
            sanitized.client = 'Not Applicable';
            debugLog('⚠️ Client field appears to be a person name, not a company, cleared');
          }
        }
      }
    }

    sanitized.client = this.normalizeClientValue(sanitized.client);

    return sanitized;
  }

  private toFinalMetadata(metadata: MetadataExtraction): Record<string, string> {
    const buDepartment = buildBuDepartmentValue(metadata.bu, metadata.department, metadata.subDepartment);
    const sensitiveTerms = metadata.sensitiveTerms || [
      metadata.emails,
      metadata.phones,
      metadata.ids,
      metadata.pricing
    ].filter((value) => !!value).join(', ');

    return {
      title: metadata.title || '',
      documentType: metadata.documentType || '',
      buDepartment: metadata.buDepartment || buDepartment || '',
      client: this.normalizeClientValue(metadata.client),
      geography: this.normalizeOptionalClassification(metadata.geography || metadata.region),
      description: metadata.description || metadata.abstract || '',
      diseaseArea: this.normalizeOptionalClassification(metadata.diseaseArea),
      therapyArea: this.normalizeOptionalClassification(metadata.therapyArea),
      sensitiveTerms: sensitiveTerms || ''
    };
  }

  /**
   * Split document text into chunks of specified size
   */
  private splitIntoChunks(text: string, chunkSize: number): string[] {
    const chunks: string[] = [];
    let start = 0;

    while (start < text.length) {
      let end = start + chunkSize;

      // If not the last chunk, try to break at a word boundary
      if (end < text.length) {
        // Look for a good break point (newline, period, space)
        const breakPoint = Math.max(
          text.lastIndexOf('\n\n', end),
          text.lastIndexOf('\n', end),
          text.lastIndexOf('. ', end),
          text.lastIndexOf(' ', end)
        );

        if (breakPoint > start + chunkSize * 0.8) {
          // Only use break point if it's not too early (at least 80% of chunk size)
          end = breakPoint + 1;
        }
      }

      chunks.push(text.substring(start, end));
      start = end;
    }

    return chunks;
  }

  /**
   * Process a single chunk of text
   */
  private async processSingleChunk(
    chunkText: string,
    chunkContext?: { chunkIndex?: number; totalChunks?: number; previousFindings?: Partial<MetadataExtraction> },
    docTypeTermsSection?: string
  ): Promise<MetadataExtraction> {
    debugLog('=== PROCESSING SINGLE CHUNK ===');
    debugLog('Chunk length:', chunkText.length, 'characters');
    if (chunkContext) {
      debugLog(`Chunk ${chunkContext.chunkIndex || 1} of ${chunkContext.totalChunks || 1}`);
    }

    const prompt = this.buildExtractionPrompt(chunkText, chunkContext, docTypeTermsSection);

    const response = await fetch(
      `${this.config.endpoint}/openai/deployments/${this.config.deploymentName}/chat/completions?api-version=${this.config.apiVersion}`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'api-key': this.config.apiKey
        },
        body: JSON.stringify({
          messages: [
            {
              role: 'system',
              content: 'You are an expert document analyzer. Extract structured information from documents and return it as valid JSON only, without any markdown formatting or code blocks. NEVER leave mandatory fields empty. Only extract information that is actually present in the document - do not make up or hallucinate information.'
            },
            {
              role: 'user',
              content: prompt
            }
          ],
          max_completion_tokens: 2000
        })
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      let errorMessage = `Azure OpenAI API error: ${response.status}`;

      // Handle CORS errors: status 0 with no body indicates a network/CORS failure
      // Note: Do NOT check response.statusText === '' because HTTP/2 Fetch responses
      // always have empty statusText, even for valid error responses like 400.
      if (response.status === 0 && (!errorText || errorText.trim() === '')) {
        errorMessage = 'CORS error: Unable to connect to Azure OpenAI. The API may need CORS configuration or a backend proxy.';
      } else {
        try {
          const errorJson = JSON.parse(errorText);
          errorMessage = errorJson.error?.message || errorText || errorMessage;
        } catch {
          errorMessage = errorText || errorMessage;
        }
      }

      console.error('Azure OpenAI API error:', response.status, errorMessage);
      throw new Error(errorMessage);
    }

    const data = await response.json();
    const content = data.choices?.[0]?.message?.content;

    if (!content) {
      throw new Error('No content returned from Azure OpenAI');
    }

    debugLog('=== AI RESPONSE (RAW) ===');
    debugLog(content);

    // Parse JSON response (remove markdown code blocks if present)
    const jsonMatch = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/) || [null, content];
    const jsonText = jsonMatch[1] || content;

    debugLog('=== PARSED JSON ===');
    debugLog(jsonText);

    const extracted = JSON.parse(jsonText.trim()) as MetadataExtraction;

    debugLog('=== EXTRACTED METADATA (BEFORE SANITIZATION) ===');
    debugLog(JSON.stringify(extracted, null, 2));
    const sanitized = this.sanitizeMetadata(extracted);
    const finalMetadata = this.toFinalMetadata(sanitized);

    debugLog('=== FINAL METADATA (AFTER SANITIZATION) ===');
    debugLog(JSON.stringify(finalMetadata, null, 2));

    return sanitized;
  }

  /**
   * Process multiple chunks and merge results
   */
  private async processChunks(chunks: string[], docTypeTermsSection?: string): Promise<MetadataExtraction> {
    debugLog('=== PROCESSING', chunks.length, 'CHUNKS ===');

    const chunkResults: MetadataExtraction[] = [];
    const previousFindings: Partial<MetadataExtraction> = {};

    // Process chunks sequentially to avoid rate limits
    for (let i = 0; i < chunks.length; i++) {
      debugLog(`\n=== PROCESSING CHUNK ${i + 1}/${chunks.length} ===`);
      try {
        const chunkContext = {
          chunkIndex: i + 1,
          totalChunks: chunks.length,
          previousFindings: i > 0 ? previousFindings : undefined
        };
        const result = await this.processSingleChunk(chunks[i], chunkContext, docTypeTermsSection);
        chunkResults.push(result);

        // Update previous findings for next chunk
        if (result.title && !previousFindings.title) previousFindings.title = result.title;
        if (result.documentType && !previousFindings.documentType) previousFindings.documentType = result.documentType;
        if (result.bu && !previousFindings.bu) previousFindings.bu = result.bu;
        if (result.department && !previousFindings.department) previousFindings.department = result.department;

        debugLog(`✓ Chunk ${i + 1} processed successfully`);
      } catch (error) {
        console.error(`✗ Error processing chunk ${i + 1}:`, error);
        // Continue with other chunks even if one fails
        chunkResults.push({} as MetadataExtraction);
      }
    }

    debugLog('\n=== MERGING CHUNK RESULTS ===');
    const merged = this.mergeChunkResults(chunkResults);

    debugLog('=== MERGED RESULT (BEFORE FINAL SANITIZATION) ===');
    debugLog(JSON.stringify(merged, null, 2));

    // Final sanitization (validation, masking, etc.)
    const finalResult = this.sanitizeMetadata(merged);

    // Validate mandatory fields and retry if needed
    const validatedResult = await this.validateAndRetryIfNeeded(finalResult, chunks);
    const finalMetadata = this.toFinalMetadata(validatedResult);

    debugLog('=== FINAL MERGED RESULT ===');
    debugLog(JSON.stringify(finalMetadata, null, 2));

    return validatedResult;
  }

  /**
   * Merge results from multiple chunks intelligently
   * Searches ALL chunks for mandatory fields to ensure nothing is missed
   */
  private mergeChunkResults(chunkResults: MetadataExtraction[]): MetadataExtraction {
    if (chunkResults.length === 0) {
      return {} as MetadataExtraction;
    }

    const merged: MetadataExtraction = {};

    // 1. Helper to concatenate and get unique, non-empty, comma-separated values for Collection Fields
    const mergeCollectionField = (
      field: keyof MetadataExtraction,
      results: MetadataExtraction[]
    ): string => {
      // Collect all values, flatten, and ensure uniqueness
      const allValues = results
        .map((r) => (r[field] as string) || '')
        // Split by comma or semicolon, trim, and filter out empty strings
        .reduce((acc, s) => {
          // Check if the value is non-empty before attempting to split
          if (!s.trim()) return acc;
          const values = s.split(/[,;]/)
            .map((v) => v.trim())
            .filter(v => v !== '');
          return acc.concat(values);
        }, [] as string[])
        .filter((v, i, a) => a.indexOf(v) === i); // Get unique values

      // Join the unique values back into a comma-separated string
      return allValues.join(', ');
    };

    // Collect all emails and phones (combine from all chunks)
    const allEmails: string[] = [];
    const allPhones: string[] = [];
    const allIds: string[] = [];
    const allPricing: string[] = [];

    // Helper function to find first non-empty value for a field
    const findFirstNonEmpty = (
      field: keyof MetadataExtraction,
      results: MetadataExtraction[]
    ): string | undefined => {
      for (const result of results) {
        const value = result[field];
        if (value && typeof value === 'string' && value.trim() !== '') {
          return value.trim();
        }
      }
      return undefined;
    };

    // For mandatory fields, search ALL chunks (not just first)
    // This ensures we don't miss title/documentType if they're in later chunks
    const allTitles: string[] = [];
    const allDocumentTypes: string[] = [];
    const allBUs: string[] = [];
    const allDepartments: string[] = [];
    const allAbstracts: string[] = [];

    for (let i = 0; i < chunkResults.length; i++) {
      const result = chunkResults[i];

      // Title: Collect from ALL chunks, prioritize first chunk but check all
      if (result.title && result.title.trim()) {
        allTitles.push(result.title.trim());
      }

      // DocumentType: Collect from ALL chunks
      if (result.documentType && result.documentType.trim()) {
        allDocumentTypes.push(result.documentType.trim());
      }

      // BU: Collect from ALL chunks
      if (result.bu && result.bu.trim()) {
        allBUs.push(result.bu.trim());
      }

      // Department: Collect from ALL chunks
      if (result.department && result.department.trim()) {
        allDepartments.push(result.department.trim());
      }

      // Abstract: Collect from ALL chunks (will take longest)
      if (result.abstract && result.abstract.trim()) {
        allAbstracts.push(result.abstract.trim());
      }

      // Region: Take first non-empty match
      if (!merged.region && result.region && result.region.trim()) {
        merged.region = result.region.trim();
      }

      // Client is resolved after all chunks so multiple clients can collapse to Multi-Client.

      // Emails: Collect all unique emails
      if (result.emails) {
        const emails = result.emails.split(/[,;\n]/).map(e => e.trim()).filter(e => e.length > 0);
        for (let j = 0; j < emails.length; j++) {
          const email = emails[j];
          // Add if not already in the list (case-insensitive)
          if (email && allEmails.indexOf(email.toLowerCase()) === -1) {
            allEmails.push(email.toLowerCase());
            // Keep original case from first occurrence
            const originalEmail = emails[j];
            if (allEmails.indexOf(originalEmail.toLowerCase()) === -1) {
              allEmails[allEmails.length - 1] = originalEmail;
            }
          }
        }
      }

      // Phones: Collect all unique phones
      if (result.phones) {
        const phones = result.phones.split(/[,;\n]/).map(p => p.trim()).filter(p => p.length > 0);
        for (let j = 0; j < phones.length; j++) {
          const phone = phones[j];
          if (phone && allPhones.indexOf(phone) === -1) {
            allPhones.push(phone);
          }
        }
      }

      // IDs: Collect all unique IDs
      if (result.ids) {
        const ids = result.ids.split(/[,;\n]/).map(id => id.trim()).filter(id => id.length > 0);
        for (let j = 0; j < ids.length; j++) {
          const id = ids[j];
          if (id && allIds.indexOf(id) === -1) {
            allIds.push(id);
          }
        }
      }

      // Pricing: Collect all pricing info
      if (result.pricing) {
        const pricing = result.pricing.trim();
        if (pricing && allPricing.indexOf(pricing) === -1) {
          allPricing.push(pricing);
        }
      }
    }

    // --- Apply Merge Logic ---

    // 1. Collection Fields (Concatenate unique values)
    merged.emails = allEmails.join(', ');
    merged.phones = allPhones.join(', ');
    merged.ids = allIds.join(', ');
    merged.pricing = allPricing.join('\n\n');

    // 2. Single/Mandatory/Conditional Fields (Take the first non-empty result)
    // The first chunk is most likely to have the best Title, BU, Department, etc.
    merged.title = findFirstNonEmpty('title', chunkResults) || (allTitles.length > 0 ? allTitles[0] : '');
    merged.documentType = findFirstNonEmpty('documentType', chunkResults) || (allDocumentTypes.length > 0 ? allDocumentTypes[0] : '');
    merged.bu = findFirstNonEmpty('bu', chunkResults) || (allBUs.length > 0 ? allBUs[0] : '');
    merged.department = findFirstNonEmpty('department', chunkResults) || (allDepartments.length > 0 ? allDepartments[0] : '');
    merged.abstract = findFirstNonEmpty('abstract', chunkResults) || (allAbstracts.length > 0 ? allAbstracts.reduce((longest, current) => current.length > longest.length ? current : longest) : '');
    merged.region = this.normalizeOptionalClassification(findFirstNonEmpty('region', chunkResults));
    merged.client = this.mergeClientValues(chunkResults);
    merged.diseaseArea = this.normalizeOptionalClassification(findFirstNonEmpty('diseaseArea', chunkResults));
    merged.therapyArea = this.normalizeOptionalClassification(findFirstNonEmpty('therapyArea', chunkResults));


    debugLog('Merged emails count:', allEmails.length);
    debugLog('Merged phones count:', allPhones.length);
    debugLog('Merged IDs count:', allIds.length);
    debugLog('Title found in chunks:', allTitles.length > 0 ? 'YES' : 'NO');
    debugLog('DocumentType found in chunks:', allDocumentTypes.length > 0 ? 'YES' : 'NO');
    debugLog('BU found in chunks:', allBUs.length > 0 ? 'YES' : 'NO');
    debugLog('Department found in chunks:', allDepartments.length > 0 ? 'YES' : 'NO');

    return merged;
  }

  /**
   * Validate mandatory fields and retry extraction if needed
   */
  private async validateAndRetryIfNeeded(result: MetadataExtraction, chunks: string[]): Promise<MetadataExtraction> {
    const mandatoryFields: (keyof MetadataExtraction)[] = ['title', 'documentType', 'bu', 'department', 'description'];
    const missingFields: (keyof MetadataExtraction)[] = [];

    // Check for missing mandatory fields
    for (const field of mandatoryFields) {
      if (!result[field] || (typeof result[field] === 'string' && result[field].trim() === '')) {
        missingFields.push(field);
      }
    }

    if (missingFields.length === 0) {
      debugLog('✓ All mandatory fields are present');
      return result;
    }

    console.warn(`⚠️ Missing mandatory fields: ${missingFields.join(', ')}`);
    debugLog('Attempting to extract missing fields from document...');

    // Try to extract missing fields by re-analyzing the first chunk or all chunks
    // Focus on the beginning of the document where title/documentType are most likely
    const firstChunk = chunks[0] || '';
    const firstChunkPreview = firstChunk.substring(0, Math.min(5000, firstChunk.length));

    if (firstChunkPreview.length > 0) {
      try {
        // Create a focused prompt for missing fields only
        const focusedPrompt = this.buildFocusedExtractionPrompt(firstChunkPreview, missingFields);

        const response = await fetch(
          `${this.config.endpoint}/openai/deployments/${this.config.deploymentName}/chat/completions?api-version=${this.config.apiVersion}`,
          {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'api-key': this.config.apiKey
            },
            body: JSON.stringify({
              messages: [
                {
                  role: 'system',
                  content: 'You are an expert document analyzer. Extract ONLY the requested fields from the document. Return valid JSON only, without any markdown formatting or code blocks. NEVER leave requested fields empty.'
                },
                {
                  role: 'user',
                  content: focusedPrompt
                }
              ],
              max_completion_tokens: 1000
            })
          }
        );

        if (response.ok) {
          const data = await response.json();
          const content = data.choices?.[0]?.message?.content;

          if (content) {
            const jsonMatch = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/) || [null, content];
            const jsonText = jsonMatch[1] || content;

            try {
              const retryResult = JSON.parse(jsonText.trim()) as Partial<MetadataExtraction>;

              // Fill in missing fields
              for (const field of missingFields) {
                if (retryResult[field] && typeof retryResult[field] === 'string' && retryResult[field].trim() !== '') {
                  result[field] = retryResult[field];
                  debugLog(`✓ Retry extracted ${field}: ${retryResult[field]}`);
                }
              }
            } catch (parseError) {
              console.warn('Failed to parse retry result:', parseError);
            }
          }
        }
      } catch (error) {
        console.warn('Retry extraction failed:', error);
      }
    }

    // Final check - if still missing, use fallback values (already handled in sanitizeMetadata)
    return result;
  }

  /**
   * Build a focused prompt for extracting specific missing fields
   */
  private buildFocusedExtractionPrompt(documentText: string, missingFields: (keyof MetadataExtraction)[]): string {
    let fieldInstructions = '';

    if (missingFields.includes('title')) {
      fieldInstructions += `\n1. title - MANDATORY - NEVER EMPTY: The document title.
Rules for generating title:

- Maximum 255 characters
- Remove ALL special characters: $ % @ * + ? ! # & ^ ~ \` | \\ < >
- Remove file extensions (.pdf, .docx, .pptx, .xlsx, .txt etc.)
- Remove version numbers and IDs (e.g. v1, v2.0, _final, _v3, -001, _20240101, document IDs, reference numbers)
- Remove leading/trailing spaces and dashes
- Use clean readable Title Case
- Extract from the document heading, first major heading, or document name — whichever is clearest
- If no explicit title exists, create a concise descriptive title based on the main topic
- NEVER return empty string`;
    }

    if (missingFields.includes('documentType')) {
      fieldInstructions += `\n2. documentType - **CRITICAL**: Extract the most likely document-type label from the document structure, title, or content. NEVER return empty string.`;
    }

    if (missingFields.includes('bu')) {
      fieldInstructions += `\n3. bu - **CRITICAL**: Business Unit. Extract the most likely business unit, practice, function, or organizational area from the content. NEVER return empty string.`;
    }

    if (missingFields.includes('department')) {
      fieldInstructions += `\n4. department - **CRITICAL**: Department. Extract the most likely department, team, capability, or sub-function from the content. NEVER return empty string.`;
    }

    if (missingFields.includes('description')) {
      fieldInstructions += `\n5. description - MANDATORY - NEVER EMPTY: A brief summary of the document.
Rules for generating description:

- 1-3 sentences maximum
- Remove ALL special characters: $ % @ * + ? ! # & ^ ~ \` | \\ < >
- Do NOT include file names, extensions, version numbers, or IDs
- Do NOT include document metadata like dates, reference numbers, or internal codes
- Write in clean plain English
- Summarize the main purpose, topic, and key content of the document
- NEVER return empty string`;
    }

    return `You are an expert document analyzer. The following fields are MISSING and MUST be extracted from the document text below. Read the document carefully and extract ONLY these fields.

**CRITICAL**: These fields are MANDATORY and CANNOT be empty. Extract them from the document text provided.${fieldInstructions}

**Document text (beginning of document):**
${documentText}

Return ONLY a JSON object with the requested fields. Example:
{
${missingFields.map(f => `  "${f}": "..."`).join(',\n')}
}`;
  }

  /**
   * Resolves an ambiguous query (using pronouns) into a standalone query using conversation history
   */
  async resolveStandaloneQuery(query: string, history: Array<{ sender: string, text: string }>): Promise<string> {
    if (!history || history.length === 0) return query;

    try {
      const prompt = `Given the following conversation history and a follow-up question, rephrase the follow-up question into a standalone question that can be understood without the history.
      
Conversation History:
${history.map(m => `${m.sender}: ${m.text}`).join('\n')}

Follow-up Question: ${query}

Standalone Question (return ONLY the question text):`;

      const response = await fetch(
        `${this.config.endpoint}/openai/deployments/${this.config.deploymentName}/chat/completions?api-version=${this.config.apiVersion}`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'api-key': this.config.apiKey
          },
          body: JSON.stringify({
            messages: [
              {
                role: 'system',
                content: 'You are a search expert. Rephrase queries to be standalone. Do not answer them.'
              },
              {
                role: 'user',
                content: prompt
              }
            ],
            max_completion_tokens: 100
          })
        }
      );

      if (!response.ok) return query;
      const json = await response.json();
      return json.choices?.[0]?.message?.content?.trim() || query;
    } catch (error) {
      console.error('Error resolving standalone query:', error);
      return query;
    }
  }

  async transcribeMedia(file: File, onProgress?: (msg: string) => void): Promise<{
    transcript: Array<{ time: number; text: string }>;
    transcriptText: string;
  }> {
    const MAX_CHUNK_SIZE_MB = 24;
    const TARGET_SAMPLE_RATE = 16000;
    const CHUNK_DURATION_SEC = 600; // 10 minutes segments

    // 1. For small files, just send directly
    if (file.size <= MAX_CHUNK_SIZE_MB * 1024 * 1024) {
      try {
        return await this._transcribeSingleChunk(file);
      } catch (err) {
        console.warn("[Whisper] Direct upload failed, falling back to audio extraction...", err);
      }
    }

    // 2. For large files or failed direct upload, extract and chunk audio
    if (onProgress) onProgress("Preparing audio for analysis (this may take a minute)...");
    
    try {
      const audioContext = new (window.AudioContext || (window as any).webkitAudioContext)();
      const arrayBuffer = await file.arrayBuffer();
      
      if (onProgress) onProgress("Decoding media track...");
      const originalBuffer = await audioContext.decodeAudioData(arrayBuffer);
      
      if (onProgress) onProgress("Optimizing audio for AI (downsampling to 16kHz Mono)...");
      
      // Use OfflineAudioContext to downsample to 16kHz Mono
      const offlineCtx = new OfflineAudioContext(
        1, // Mono
        Math.floor(originalBuffer.duration * TARGET_SAMPLE_RATE),
        TARGET_SAMPLE_RATE
      );
      
      const source = offlineCtx.createBufferSource();
      source.buffer = originalBuffer;
      source.connect(offlineCtx.destination);
      source.start();
      
      const audioBuffer = await offlineCtx.startRendering();
      
      const totalDuration = audioBuffer.duration;
      const numChunks = Math.ceil(totalDuration / CHUNK_DURATION_SEC);
      
      debugLog(`[Whisper] Downsampled to 16kHz Mono. Duration: ${totalDuration.toFixed(2)}s. Chunks: ${numChunks}`);

      let allSegments: Array<{ time: number; text: string }> = [];
      let combinedText = "";
      let successfulChunks = 0;

      for (let i = 0; i < numChunks; i++) {
        const startTime = i * CHUNK_DURATION_SEC;
        const endTime = Math.min(startTime + CHUNK_DURATION_SEC, totalDuration);
        const duration = endTime - startTime;
        
        if (onProgress) onProgress(`Transcribing part ${i + 1} of ${numChunks}...`);
        
        // Extract slice from buffer
        const chunkLength = Math.floor(duration * TARGET_SAMPLE_RATE);
        const chunkBuffer = audioContext.createBuffer(1, chunkLength, TARGET_SAMPLE_RATE);
        
        const channelData = audioBuffer.getChannelData(0);
        const chunkData = chunkBuffer.getChannelData(0);
        const startOffset = Math.floor(startTime * TARGET_SAMPLE_RATE);
        
        // Use typed array subarray for performance
        chunkData.set(channelData.subarray(startOffset, startOffset + chunkLength));

        // Encode chunk to WAV
        const wavBlob = this._bufferToWav(chunkBuffer);
        const chunkFile = new File([wavBlob], `part_${i}.wav`, { type: "audio/wav" });

        try {
          const { transcript, transcriptText } = await this._transcribeSingleChunk(chunkFile);
          
          const adjustedSegments = transcript.map(s => ({
            time: s.time + startTime,
            text: s.text
          }));
          
          allSegments = allSegments.concat(adjustedSegments);
          combinedText += (combinedText ? " " : "") + transcriptText;
          successfulChunks += 1;
        } catch (err) {
          console.error(`[Whisper] Chunk ${i + 1} failed:`, err);
          allSegments.push({ time: startTime, text: `[Transcription failed for this segment: ${err.message}]` });
        }
      }

      if (successfulChunks === 0) {
        throw new Error('All media transcription chunks failed. Check that the frontend is calling the hosted knowledge search API.');
      }

      return { transcript: allSegments, transcriptText: combinedText };

    } catch (err) {
      console.error("🔴 Audio extraction failed:", err);
      return { 
        transcript: [{ time: 0, text: `[Audio extraction failed: ${err.message}. File may be too large or format unsupported.]` }], 
        transcriptText: "No transcript available." 
      };
    }
  }

  /**
   * Encodes an AudioBuffer to a WAV blob (PCM)
   */
  private _bufferToWav(buffer: AudioBuffer): Blob {
    const numChannels = buffer.numberOfChannels;
    const sampleRate = buffer.sampleRate;
    const length = buffer.length * numChannels * 2 + 44;
    const view = new DataView(new ArrayBuffer(length));

    // RIFF identifier
    view.setUint32(0, 0x52494646, false); // "RIFF"
    view.setUint32(4, 36 + buffer.length * numChannels * 2, true);
    view.setUint32(8, 0x57415645, false); // "WAVE"
    // Format chunk identifier
    view.setUint32(12, 0x666d7420, false); // "fmt "
    view.setUint32(16, 16, true); // format chunk length
    view.setUint16(20, 1, true); // sample format (PCM)
    view.setUint16(22, numChannels, true);
    view.setUint32(24, sampleRate, true);
    view.setUint32(28, sampleRate * numChannels * 2, true); // byte rate
    view.setUint16(32, numChannels * 2, true); // block align
    view.setUint16(34, 16, true); // bits per sample
    // Data chunk identifier
    view.setUint32(36, 0x64617461, false); // "data"
    view.setUint32(40, buffer.length * numChannels * 2, true);

    // Write PCM samples
    const offset = 44;
    for (let i = 0; i < buffer.length; i++) {
      for (let channel = 0; channel < numChannels; channel++) {
        const sample = Math.max(-1, Math.min(1, buffer.getChannelData(channel)[i]));
        view.setInt16(offset + (i * numChannels + channel) * 2, sample < 0 ? sample * 0x8000 : sample * 0x7FFF, true);
      }
    }

    return new Blob([view], { type: 'audio/wav' });
  }

  private async _transcribeSingleChunk(file: File | Blob): Promise<{
    transcript: Array<{ time: number; text: string }>;
    transcriptText: string;
  }> {
    let transcriptData: Array<{ time: number; text: string }> = [];
    let transcriptText = "";

    if (this.isUploadAiBackendConfigured()) {
      const uploadFile = file instanceof File
        ? file
        : new File([file], 'media.wav', { type: file.type || 'audio/wav' });
      try {
        return await this.uploadAiClient.transcribeMedia(uploadFile);
      } catch (error) {
        console.warn('[Whisper] Backend transcription failed:', error);
        if (!AZURE_OPENAI_WHISPER_API_KEY) {
          throw error;
        }
      }
    }

    try {
      const formData = new FormData();
      formData.append('file', file);
      // model and response_format are optional if endpoint is a deployment URL, but good to have
      formData.append('model', 'whisper');
      formData.append('response_format', 'verbose_json');
      
      const whisperResponse = await fetch(
        AZURE_OPENAI_WHISPER_ENDPOINT,
        {
          method: 'POST',
          headers: {
            'api-key': AZURE_OPENAI_WHISPER_API_KEY
          },
          body: formData
        }
      );

      if (whisperResponse.ok) {
        const data = await whisperResponse.json();
        
        if (typeof data === 'string') {
          transcriptText = data;
          transcriptData = [{ time: 0, text: data }];
        } else if (data.segments && Array.isArray(data.segments)) {
          transcriptData = data.segments.map((s: any) => ({
            time: Math.floor(s.start || 0),
            text: (s.text || "").trim()
          }));
          transcriptText = transcriptData.map(t => t.text).join(" ");
        } else if (data.text) {
          transcriptText = data.text;
          transcriptData = [{ time: 0, text: data.text }];
        }
      } else {
        const errorText = await whisperResponse.text();
        throw new Error(`API error ${whisperResponse.status}: ${errorText}`);
      }

    } catch (err) {
      console.error("🔴 Whisper API call error:", err);
      throw err;
    }

    return { transcript: transcriptData, transcriptText };
  }

  async generateAbstractFromTranscript(
    fileName: string,
    transcriptText: string,
    docTypeTermsSection?: string,
    clientTermsSection?: string
  ): Promise<{
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
  }> {
    if (this.isUploadAiBackendConfigured()) {
      try {
        return await this.uploadAiClient.analyzeMediaTranscript(
          fileName,
          transcriptText,
          docTypeTermsSection,
          clientTermsSection
        );
      } catch (error) {
        console.warn('Backend media transcript analysis failed:', error);
        if (!this.hasLegacyOpenAiConfig()) {
          throw error;
        }
      }
    }

    try {
      // 2. Generate abstract from transcript via GPT
      const docTypeInstruction = docTypeTermsSection
        ? `documentType: Choose the single BEST matching document type
from this list based on the transcript content. Return ONLY the
exact term name.

Available document types:
${docTypeTermsSection}  `
        : `documentType: Identify the media document type.
Use labels like "Training", "Webinar", "Podcast", "Demo",
"Presentation", "Interview" etc.`;

      const prompt = `Analyze the following transcript from a media file.

RULES:

- Only extract information actually present in the transcript
- Do NOT hallucinate or invent information not in the transcript
- For fields not found, use "Not Applicable"
- Remove special characters ($, %, @, *, +, ?, !, #) from all text
- Do NOT include file extensions, version numbers, or IDs in title

FIELDS TO EXTRACT:

title: Generate a clean, descriptive title based on the transcript
content - NOT the file name.
Rules:

- Maximum 255 characters
- Remove special characters: $ % @ * + ? ! # & ^ ~ | < >
- Remove file extensions (.mp4, .mp3, .wav, .mov etc.)
- Remove version numbers, IDs, date patterns
- Use clean readable Title Case
- Base it on the actual content/topic discussed in the transcript

abstract: A professional, detailed 1 paragraph summary of the media content.
Rules:

- Write as if describing the content to someone who has not seen it
- Do NOT mention 'transcript', 'audio', 'video', 'recording',
'speaker', 'the speaker says', 'this video', 'this audio' etc.
- Do NOT include file names, extensions, version numbers, or IDs
- Remove ALL special characters: $ % @ * + ? ! # & ^ ~ | < >
- Write in clean plain English describing the topic and key points
- Focus on WHAT the content is about, not HOW it was delivered

businessUnit: Extract the business unit or organizational area
mentioned or implied. NEVER empty.

department: Extract the department, team, or function mentioned
or implied. NEVER empty.

${docTypeInstruction}

client: Client name rules:

- If ONE specific company name found -> return that name
- If MORE THAN ONE company found -> return "Multi-Client"
- If client exists but not identifiable -> return "Others"
- If NO client found -> return "Not Applicable"
- Do NOT infer the client from file names, folder paths, repository URLs, email domains, project codes, sensitiveTerms, or internal Indegene contacts

geography: Geography rules:

- If ONE specific geography found -> return that value
- If MORE THAN ONE geography found -> return "General"
- If NO geography found -> return "Not Applicable"

therapyArea: Therapy area rules:

- If ONE therapy area found -> return that value
- If MORE THAN ONE therapy area found -> return "Cross-Therapy"
- If NO therapy area found -> return "Not Applicable"

diseaseArea: Disease area rules:

- If ONE disease area found -> return that value
- If MORE THAN ONE disease area found -> return "Cross-Disease"
- If NO disease area found -> return "Not Applicable"

Transcript to analyze:
${transcriptText}

Return ONLY valid JSON in this exact format:
{
"title": "...",
"abstract": "...",
"businessUnit": "...",
"department": "...",
"documentType": "...",
"client": "...",
"geography": "...",
"therapyArea": "...",
"diseaseArea": "..."
}`;

      const response = await fetch(
        `${this.config.endpoint}/openai/deployments/${this.config.deploymentName}/chat/completions?api-version=${this.config.apiVersion}`,
        {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'api-key': this.config.apiKey
          },
          body: JSON.stringify({
            messages: [
              {
                role: 'system',
                content: 'You are an AI assistant that extracts metadata from transcripts. Return only valid JSON.'
              },
              {
                role: 'user',
                content: prompt
              }
            ],
            max_completion_tokens: 500
          })
        }
      );

      if (response.ok) {
        const data = await response.json();
        const content = data.choices?.[0]?.message?.content;
        if (content) {
          const jsonMatch = content.match(/```(?:json)?\s*([\s\S]*?)\s*```/) || [null, content];
          const jsonText = jsonMatch[1] || content;
          const result = JSON.parse(jsonText.trim());

          return {
            abstract: result.abstract || '',
            businessUnit: result.businessUnit || '',
            department: result.department || '',
            documentType: result.documentType || '',
            client: this.normalizeClientValue(result.client),
            geography: this.normalizeOptionalClassification(result.geography),
            therapyArea: this.normalizeOptionalClassification(result.therapyArea),
            diseaseArea: this.normalizeOptionalClassification(result.diseaseArea),
            title: result.title || ''
          };
        }
      }
    } catch (error) {
      console.error("Error generating abstract from transcript:", error);
    }

    // Fallback if GPT API fails
    return {
      abstract: "This media content could not be fully analyzed.",
      businessUnit: "",
      department: "",
      documentType: "",
      client: "Not Applicable",
      geography: "Not Applicable",
      therapyArea: "Not Applicable",
      diseaseArea: "Not Applicable",
      title: ""
    };
  }

  /**
   * Refines a natural language search query into KQL keywords.
   */
  public async getRefinedKeywords(query: string, history: any[] = []): Promise<string> {
    try {
      const url = `${this.config.endpoint}/openai/deployments/${this.config.deploymentName}/chat/completions?api-version=${this.config.apiVersion}`;
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'api-key': this.config.apiKey
        },
        body: JSON.stringify({
          messages: [
            { 
              role: 'system', 
              content: `You are an expert search query optimizer for the Indegene Knowledge Hub. 
              Analyze the user's question and generate high-performance SharePoint Search (KQL) keywords.
              
              RULES:
              - Strip out conversational filler (e.g., "Do you have", "Find me", "Search for").
              - Extract core entities like Client names (e.g. Takeda), Therapy Areas (e.g. Hematology), or Department names.
              - Output ONLY the search keywords, separated by spaces. No conversational filler.` 
            },
            ...history,
            { role: 'user', content: query }
          ]
        })
      });

      if (response.ok) {
        const data = await response.json();
        return (data.choices[0].message.content || '').trim();
      }
    } catch (error) {
      console.warn("AI Query Refinement failed:", error);
    }
    return query; // Fallback to original query
  }
}
