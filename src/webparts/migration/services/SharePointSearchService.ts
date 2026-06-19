// src/webparts/migration/services/SharePointSearchService.ts
//
// Legacy SharePoint Search API helper for non-backend document surfaces.
// This bypasses ALL Azure Search index / admin-consent requirements.

import { SPHttpClient } from '@microsoft/sp-http';
import { ISearchResult } from '../models/SearchResult';
import { PromptBuilder } from '../utils/PromptBuilder';
import { SearchRankingService } from './SearchRankingService';
import { COLUMN_NAMES, LIBRARY_NAMES, SEARCH_PROPERTIES, TTL_MS } from '../config/appConfig';

const serviceInstanceCache = new Map<string, SharePointSearchService>();

interface ILibraryStatusCountSummary {
  fileCount: number;
  folderCount: number;
  statusCounts: Record<string, number>;
}

export const getDriveItemThumbnailUrl = (
  siteUrl: string,
  driveId: string,
  fileUniqueId: string,
  size: string = 'c2048x2048'
): string => {
  const cleanId = fileUniqueId.replace(/{|}/g, '');
  return `${siteUrl}/_api/v2.1/drives/${driveId}/items/${cleanId}/thumbnails/0/${size}/content`;
};

export const getExcelThumbnailUrl = getDriveItemThumbnailUrl;

export function getSearchServiceInstance(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  libraryName: string
): SharePointSearchService {
  const key = `${siteUrl}::${libraryName}`;
  if (!serviceInstanceCache.has(key)) {
    serviceInstanceCache.set(key, new SharePointSearchService(spHttpClient, siteUrl, libraryName));
  }
  return serviceInstanceCache.get(key)!;
}

export class SharePointSearchService {
  private spHttpClient: SPHttpClient;
  private siteUrl: string;
  private libraryPath: string;
  private static readonly PATH_CACHE_TTL_MS = TTL_MS.searchCache;
  private static readonly SEARCH_CACHE_TTL_MS = TTL_MS.searchCache;
  private static readonly METADATA_CACHE_TTL_MS = TTL_MS.searchCache;
  private pathCache: {
    fetchedAt: number;
    kmPath?: string;
    whosWhoPath?: string;
  } | null = null;
  private pendingPathLookup: Promise<{ kmPath?: string; whosWhoPath?: string }> | null = null;
  private searchResultCache = new Map<string, { savedAt: number; results: ISearchResult[] }>();
  private pendingSearches = new Map<string, Promise<ISearchResult[]>>();
  private metadataItemsCache: { savedAt: number; items: any[] } | null = null;
  private metadataFallbackCache = new Map<string, { savedAt: number; results: ISearchResult[] }>();
  private pendingMetadataFallbacks = new Map<string, Promise<ISearchResult[]>>();
  private rootFolderUrlCache: string | null = null;
  private fieldMapCache: any = null;
  private static libraryStatusCountCache = new Map<string, { fetchedAt: number; summary: ILibraryStatusCountSummary }>();
  private static latestDocumentIdsCache = new Map<string, { fetchedAt: number; ids: number[] }>();
  private static latestPublishedDocumentIdsCache = new Map<string, { fetchedAt: number; ids: number[] }>();
  private static authorTitleCache = new Map<number, string>();

  constructor(spHttpClient: SPHttpClient, siteUrl: string, libraryName = LIBRARY_NAMES.kmDataHub) {
    this.spHttpClient = spHttpClient;
    this.siteUrl = siteUrl;
    this.libraryPath = libraryName;
  }

  public setFieldMapCache(fieldMap: any): void {
    this.fieldMapCache = fieldMap;
  }

  public static getCachedAuthorTitle(authorId: number): string | undefined {
    return SharePointSearchService.authorTitleCache.get(authorId);
  }

  public static setCachedAuthorTitle(authorId: number, title: string): void {
    SharePointSearchService.authorTitleCache.set(authorId, title);
  }

  private async getRootFolderUrl(): Promise<string> {
    if (this.rootFolderUrlCache) return this.rootFolderUrlCache;
    try {
      const resp = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${this.libraryPath}')/RootFolder?$select=ServerRelativeUrl`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return '';
      const data = await resp.json();
      const siteOrigin = this.siteUrl.split('/sites/')[0];
      this.rootFolderUrlCache = `${siteOrigin}${data?.ServerRelativeUrl || ''}`;
      return this.rootFolderUrlCache;
    } catch {
      return '';
    }
  }

  /**
   * Strips HTML entities and tags returned inside HitHighlightedSummary.
   */
  private static cleanSummary(raw: string): string {
    return (raw || '')
      .replace(/<ddd\/>/gi, '…')
      .replace(/<c0>/gi, '')
      .replace(/<\/c0>/gi, '')
      .replace(/<[^>]+>/g, '')
      .replace(/&amp;/g, '&')
      .replace(/&quot;/g, '"')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&nbsp;/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  /**
   * Extracts a named property from a SharePoint Search Row's Cells array.
   */
  private static cell(cells: Array<{ Key: string; Value: string }>, key: string): string {
    const found = cells.find((c) => c.Key === key);
    return found ? (found.Value || '') : '';
  }

  private static escapeODataValue(value: string): string {
    return (value || '').replace(/'/g, "''");
  }

  private static escapeKqlPhrase(value: string): string {
    return (value || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"').trim();
  }

  private static normalizeStatusFilters(status?: string | string[]): string[] {
    return (Array.isArray(status) ? status : status ? [status] : [])
      .map((value) => String(value || '').trim())
      .filter(Boolean);
  }

  private static buildODataStatusCondition(fieldName: string, status?: string | string[]): string {
    const statuses = SharePointSearchService.normalizeStatusFilters(status);
    if (statuses.length === 0) return '';
    const clauses = statuses.map((value) => `${fieldName} eq '${SharePointSearchService.escapeODataValue(value)}'`);
    return clauses.length === 1 ? clauses[0] : `(${clauses.join(' or ')})`;
  }

  private static buildKqlStatusCondition(status?: string | string[]): string {
    const statuses = SharePointSearchService.normalizeStatusFilters(status);
    if (statuses.length === 0) return '';
    const clauses = statuses.map((value) => {
      const safeValue = SharePointSearchService.escapeKqlPhrase(value);
      return `${SEARCH_PROPERTIES.status}:${safeValue.includes(' ') ? `"${safeValue}"` : safeValue}`;
    });
    return clauses.length === 1 ? clauses[0] : `(${clauses.join(' OR ')})`;
  }

  private async getCachedSearchPaths(): Promise<{ kmPath?: string; whosWhoPath?: string }> {
    if (
      this.pathCache &&
      Date.now() - this.pathCache.fetchedAt < SharePointSearchService.PATH_CACHE_TTL_MS
    ) {
      return {
        kmPath: this.pathCache.kmPath,
        whosWhoPath: this.pathCache.whosWhoPath
      };
    }

    if (this.pendingPathLookup) {
      return this.pendingPathLookup;
    }

    this.pendingPathLookup = this.fetchSearchPaths().finally(() => {
      this.pendingPathLookup = null;
    });

    return this.pendingPathLookup;
  }

  private async fetchSearchPaths(): Promise<{ kmPath?: string; whosWhoPath?: string }> {
    const [kmRes, whosWhoRes, whosWhoDirRes] = await Promise.all([
      this.spHttpClient.get(`${this.siteUrl}/_api/web/lists/getbytitle('${this.libraryPath}')/RootFolder?$select=ServerRelativeUrl`, SPHttpClient.configurations.v1),
      this.spHttpClient.get(`${this.siteUrl}/_api/web/lists/getbytitle('Who''s Who')/RootFolder?$select=ServerRelativeUrl`, SPHttpClient.configurations.v1).catch(() => null),
      this.spHttpClient.get(`${this.siteUrl}/_api/web/lists/getbytitle('Who''s Who Directory')/RootFolder?$select=ServerRelativeUrl`, SPHttpClient.configurations.v1).catch(() => null)
    ]);

    const siteOrigin = this.siteUrl.split('/sites/')[0];
    let kmPath = '';
    let whosWhoPath = '';

    if (kmRes.ok) {
      const kmJson = await kmRes.json();
      if (kmJson.ServerRelativeUrl) {
        kmPath = `${siteOrigin}${kmJson.ServerRelativeUrl}*`;
      }
    }

    const whosWhoListRes = whosWhoRes?.ok ? whosWhoRes : whosWhoDirRes;
    if (whosWhoListRes?.ok) {
      const whosWhoJson = await whosWhoListRes.json();
      if (whosWhoJson.ServerRelativeUrl) {
        whosWhoPath = `${siteOrigin}${whosWhoJson.ServerRelativeUrl}*`;
      }
    } else {
      try {
        const listLookupRes = await this.spHttpClient.get(
          `${this.siteUrl}/_api/web/lists?$filter=substringof('Who',Title) or substringof('People',Title)&$select=RootFolder/ServerRelativeUrl,Title&$expand=RootFolder`,
          SPHttpClient.configurations.v1
        );
        if (listLookupRes.ok) {
          const listLookupJson = await listLookupRes.json();
          const foundList = (listLookupJson.value || [])[0];
          if (foundList && foundList.RootFolder?.ServerRelativeUrl) {
            whosWhoPath = `${siteOrigin}${foundList.RootFolder.ServerRelativeUrl}*`;
          }
        }
      } catch (e) {
        console.warn("Personnel list dynamic lookup failed:", e);
      }
    }

    this.pathCache = {
      fetchedAt: Date.now(),
      kmPath,
      whosWhoPath
    };

    return { kmPath, whosWhoPath };
  }

  private normalizeForMatch(value: string): string {
    return (value || '')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private extractMetadataSearchPhrase(query: string): string {
    const normalized = (query || '')
      .replace(/[?.,!]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();

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
    return this.normalizeForMatch(value)
      .split(/\s+/)
      .filter((token) => ![
        'ok', 'please', 'can', 'you', 'tell', 'me', 'if', 'do', 'does', 'have', 'has', 'any',
        'documents', 'document', 'files', 'file', 'related', 'to', 'for', 'about', 'under',
        'with', 'from', 'client', 'bu', 'business', 'unit', 'department', 'sub', 'type', 'area'
      ].includes(token))
      .join(' ')
      .trim();
  }

  private extractFileExtensions(query: string): string[] {
    return Array.from(new Set(((query || '').match(/\b(mp3|mp4|mov|avi|mkv|wmv|m4v|pdf|pptx|ppt|docx|doc|xlsx|xls)\b/gi) || [])
      .map((extension) => extension.toUpperCase())));
  }

  private stringifyListField(value: any): string {
    if (!value) {
      return '';
    }

    if (typeof value === 'string') {
      return value;
    }

    if (Array.isArray(value)) {
      return value.map((entry) => this.stringifyListField(entry)).filter(Boolean).join('; ');
    }

    if (typeof value === 'object') {
      return value.Label || value.Title || value.LookupValue || value.Name || value.Value || '';
    }

    return String(value);
  }

  private stringifyListFieldArray(value: any): string[] {
    return this.stringifyListField(value)
      .split(/[;,]/)
      .map((entry) => entry.trim())
      .filter(Boolean);
  }

  private numberListField(value: any): number {
    const parsed = Number(this.stringifyListField(value));
    return Number.isFinite(parsed) ? parsed : 0;
  }

  private booleanListField(value: any): boolean {
    if (typeof value === 'boolean') return value;
    if (typeof value === 'number') return value > 0;
    const normalized = this.stringifyListField(value).trim().toLowerCase();
    return normalized === 'true' || normalized === 'yes' || normalized === '1';
  }

  private static cellNumber(cells: Array<{ Key: string; Value: string }>, ...keys: string[]): number {
    for (const key of keys) {
      const parsed = Number(SharePointSearchService.cell(cells, key));
      if (Number.isFinite(parsed)) {
        return parsed;
      }
    }

    return 0;
  }

  private static cellBoolean(cells: Array<{ Key: string; Value: string }>, ...keys: string[]): boolean {
    for (const key of keys) {
      const normalized = SharePointSearchService.cell(cells, key).trim().toLowerCase();
      if (normalized === 'true' || normalized === 'yes' || normalized === '1') return true;
    }

    return false;
  }

  private static cellArray(cells: Array<{ Key: string; Value: string }>, ...keys: string[]): string[] {
    for (const key of keys) {
      const value = SharePointSearchService.cell(cells, key);
      if (value) {
        return value.split(/[;,]/).map((entry) => entry.trim()).filter(Boolean);
      }
    }

    return [];
  }

  private scoreListItemMetadata(item: any, phrase: string, phraseTokens: string[]): number {
    const title = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.title]));
    const description = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.description]));
    const businessUnit = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.businessUnit]));
    const department = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.department]));
    const documentType = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.documentType]));
    const client = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.client]));
    const geography = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.geography]));
    const therapyArea = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.therapyArea]));
    const diseaseArea = this.normalizeForMatch(this.stringifyListField(item[COLUMN_NAMES.diseaseArea]));
    const normalizedPhrase = this.normalizeForMatch(phrase);
    const metadataText = [
      businessUnit,
      department,
      documentType,
      client,
      geography,
      therapyArea,
      diseaseArea
    ].join(' ');

    let metadataScore = 0;
    if (metadataText.indexOf(normalizedPhrase) !== -1) metadataScore += 100;

    phraseTokens.forEach((token) => {
      if (businessUnit.split(/\s+/).includes(token)) metadataScore += 18;
      if (client.split(/\s+/).includes(token)) metadataScore += 18;
      if (documentType.split(/\s+/).includes(token)) metadataScore += 18;
      if (department.split(/\s+/).includes(token)) metadataScore += 14;
      if (therapyArea.split(/\s+/).includes(token)) metadataScore += 14;
      if (diseaseArea.split(/\s+/).includes(token)) metadataScore += 14;
      if (geography.split(/\s+/).includes(token)) metadataScore += 14;
    });

    if (metadataScore === 0) {
      return 0;
    }

    let score = metadataScore;
    if (title.indexOf(normalizedPhrase) !== -1) score += 12;
    if (description.indexOf(normalizedPhrase) !== -1) score += 8;

    return score;
  }

  private async getCachedMetadataItems(selectFields: string[]): Promise<any[]> {
    if (
      this.metadataItemsCache &&
      Date.now() - this.metadataItemsCache.savedAt < SharePointSearchService.METADATA_CACHE_TTL_MS
    ) {
      return this.metadataItemsCache.items;
    }

    const baseUrl =
      `${this.siteUrl}/_api/web/lists/getbytitle('${SharePointSearchService.escapeODataValue(this.libraryPath)}')/items` +
      `?$select=${selectFields.join(',')}`;
    let pageUrl = `${baseUrl}&$top=500`;
    const allItems: any[] = [];

    while (pageUrl) {
      const response = await this.spHttpClient.get(pageUrl, SPHttpClient.configurations.v1);
      if (!response.ok) {
        break;
      }

      const json = await response.json();
      const items = json?.value || json?.d?.results || [];
      allItems.push(...items);
      pageUrl = json?.['@odata.nextLink'] || json?.d?.__next || '';
    }

    this.metadataItemsCache = { savedAt: Date.now(), items: allItems };
    return allItems;
  }

  private async searchMetadataListItems(query: string): Promise<ISearchResult[]> {
    const phrase = this.extractMetadataSearchPhrase(query);
    const extensions = this.extractFileExtensions(query);
    if ((!phrase || phrase.length < 2) && extensions.length === 0) {
      return [];
    }

    const cacheKey = `${this.normalizeForMatch(phrase)}|${extensions.join(',')}`;
    const cachedFallback = this.metadataFallbackCache.get(cacheKey);
    if (cachedFallback && Date.now() - cachedFallback.savedAt < SharePointSearchService.METADATA_CACHE_TTL_MS) {
      return cachedFallback.results;
    }

    const pendingFallback = this.pendingMetadataFallbacks.get(cacheKey);
    if (pendingFallback) {
      return pendingFallback;
    }

    const fallbackPromise = this.searchMetadataListItemsUncached(phrase, extensions).then((results) => {
      this.metadataFallbackCache.set(cacheKey, { savedAt: Date.now(), results });
      return results;
    }).finally(() => {
      this.pendingMetadataFallbacks.delete(cacheKey);
    });

    this.pendingMetadataFallbacks.set(cacheKey, fallbackPromise);
    return fallbackPromise;
  }

  private async searchMetadataListItemsUncached(phrase: string, extensions: string[] = []): Promise<ISearchResult[]> {

    const selectFields = [
      'Id',
      COLUMN_NAMES.title,
      COLUMN_NAMES.description,
      COLUMN_NAMES.fileRef,
      COLUMN_NAMES.fileLeafRef,
      COLUMN_NAMES.status,
      COLUMN_NAMES.businessUnit,
      COLUMN_NAMES.department,
      COLUMN_NAMES.documentType,
      COLUMN_NAMES.client,
      COLUMN_NAMES.geography,
      COLUMN_NAMES.therapyArea,
      COLUMN_NAMES.diseaseArea,
      COLUMN_NAMES.contentRefreshDate,
      COLUMN_NAMES.views,
      COLUMN_NAMES.likes,
      COLUMN_NAMES.comments,
      COLUMN_NAMES.downloads,
      COLUMN_NAMES.follow,
      COLUMN_NAMES.share,
      COLUMN_NAMES.bookmark,
      COLUMN_NAMES.sensitiveTerms,
      COLUMN_NAMES.reviewerComments,
      COLUMN_NAMES.projectId,
      COLUMN_NAMES.versionFileName,
      COLUMN_NAMES.versionFileType,
      COLUMN_NAMES.docIcon,
      `${COLUMN_NAMES.modifiedBy}/Title`,
      `${COLUMN_NAMES.modifiedBy}/EMail`,
      `${COLUMN_NAMES.modifiedBy}/Id`
    ];

    try {
      const siteOrigin = this.siteUrl.split('/sites/')[0];
      const phraseTokens = this.normalizeForMatch(phrase).split(/\s+/).filter(Boolean);
      const metadataItems = await this.getCachedMetadataItems(selectFields);
      return metadataItems
        .filter((item: any) => {
          const status = this.stringifyListField(item[COLUMN_NAMES.status]).toLowerCase();
          return !status || status.indexOf('active') !== -1;
        })
        .map((item: any) => {
          const extensionScore = extensions.includes((this.stringifyListField(item[COLUMN_NAMES.fileLeafRef]).split('.').pop() || '').toUpperCase()) ? 120 : 0;
          return { item, score: extensionScore + this.scoreListItemMetadata(item, phrase, phraseTokens) };
        })
        .filter(({ score }: { score: number }) => score > 0)
        .sort((left: { score: number }, right: { score: number }) => right.score - left.score)
        .slice(0, 10)
        .map((item: any) => {
          const listItem = item.item;
          const businessUnit = this.stringifyListField(listItem[COLUMN_NAMES.businessUnit]);
          const department = this.stringifyListField(listItem[COLUMN_NAMES.department]);
          const documentType = this.stringifyListField(listItem[COLUMN_NAMES.documentType]);
          const client = this.stringifyListField(listItem[COLUMN_NAMES.client]);
          const geography = this.stringifyListField(listItem[COLUMN_NAMES.geography]);
          const therapyArea = this.stringifyListField(listItem[COLUMN_NAMES.therapyArea]);
          const diseaseArea = this.stringifyListField(listItem[COLUMN_NAMES.diseaseArea]);
          const serverRelativeUrl = this.stringifyListField(listItem[COLUMN_NAMES.fileRef]);
          const modifiedBy = listItem[COLUMN_NAMES.modifiedBy];
          const url = serverRelativeUrl && serverRelativeUrl.startsWith('http')
            ? serverRelativeUrl
            : `${siteOrigin}${serverRelativeUrl}`;
          const metadataHeader = `[METADATA] Department: ${department || 'N/A'}, Client: ${client || 'N/A'}, BU: ${businessUnit || 'N/A'}, Document Type: ${documentType || 'N/A'}, Therapy Area: ${therapyArea || 'N/A'}, Disease Area: ${diseaseArea || 'N/A'}, Geography: ${geography || 'N/A'}\n\n`;
          const description = this.stringifyListField(listItem[COLUMN_NAMES.description]);
          const title = this.stringifyListField(listItem[COLUMN_NAMES.title]) || this.stringifyListField(listItem[COLUMN_NAMES.fileLeafRef]) || 'Untitled';

          return {
            id: String(listItem.Id || 0),
            title,
            content: metadataHeader + PromptBuilder.truncateContent(description || title, 4000),
            url,
            author: 'Internal',
            fileType: (this.stringifyListField(listItem[COLUMN_NAMES.fileLeafRef]).split('.').pop() || 'DOC').toUpperCase(),
            description,
            rank: 1000,
            status: 'Active',
            listTitle: this.libraryPath,
            fileName: this.stringifyListField(listItem[COLUMN_NAMES.fileLeafRef]),
            contentRefreshDate: this.stringifyListField(listItem[COLUMN_NAMES.contentRefreshDate]),
            viewsCount: this.numberListField(listItem[COLUMN_NAMES.views]),
            likesCount: this.numberListField(listItem[COLUMN_NAMES.likes]),
            commentsCount: this.numberListField(listItem[COLUMN_NAMES.comments]),
            downloadsCount: this.numberListField(listItem[COLUMN_NAMES.downloads]),
            downloads: this.numberListField(listItem[COLUMN_NAMES.downloads]),
            follow: this.booleanListField(listItem[COLUMN_NAMES.follow]),
            share: this.booleanListField(listItem[COLUMN_NAMES.share]),
            bookmark: this.booleanListField(listItem[COLUMN_NAMES.bookmark]),
            sensitiveTerms: this.stringifyListFieldArray(listItem[COLUMN_NAMES.sensitiveTerms]),
            reviewerComments: this.stringifyListField(listItem[COLUMN_NAMES.reviewerComments]),
            projectId: this.stringifyListField(listItem[COLUMN_NAMES.projectId]),
            versionFileName: this.stringifyListField(listItem[COLUMN_NAMES.versionFileName]),
            versionFileType: this.stringifyListField(listItem[COLUMN_NAMES.versionFileType]),
            modifiedBy: modifiedBy ? {
              title: this.stringifyListField(modifiedBy.Title),
              email: this.stringifyListField(modifiedBy.EMail),
              id: Number(modifiedBy.Id) || 0
            } : undefined,
            docIcon: this.stringifyListField(listItem[COLUMN_NAMES.docIcon]),
            businessUnit,
            documentType,
            department,
            geography,
            diseaseArea,
            therapyArea,
            client
          } as ISearchResult;
        });
    } catch (error) {
      console.warn('SharePoint metadata list fallback failed:', error);
      return [];
    }
  }

  private scoreResultAgainstQuery(query: string, result: ISearchResult): number {
    const normalizedQuery = this.normalizeForMatch(query);
    const normalizedTitle = this.normalizeForMatch(result.title || '');
    const normalizedUrl = this.normalizeForMatch(result.url || '');
    const normalizedBusinessUnit = this.normalizeForMatch(result.businessUnit || '');
    const normalizedDocumentType = this.normalizeForMatch(result.documentType || '');
    const normalizedContent = this.normalizeForMatch(result.content || '');
    let score = Number((result as any).rank || 0);

    if (!normalizedQuery || normalizedQuery === '*') {
      return score;
    }

    if (normalizedTitle === normalizedQuery) {
      score += 200;
    } else if (normalizedTitle.indexOf(normalizedQuery) !== -1) {
      score += 120;
    }

    if (normalizedUrl.indexOf(normalizedQuery) !== -1) {
      score += 100;
    }

    if (normalizedBusinessUnit === normalizedQuery || normalizedDocumentType === normalizedQuery) {
      score += 75;
    }

    if (normalizedContent.indexOf(normalizedQuery) !== -1) {
      score += 20;
    }

    normalizedQuery.split(/\s+/).filter(Boolean).forEach((token) => {
      if (token.length < 2) {
        return;
      }

      if (normalizedTitle.split(/\s+/).includes(token)) {
        score += 30;
      }

      if (normalizedUrl.split(/\s+/).includes(token)) {
        score += 20;
      }
    });

    return score;
  }

  /**
   * 🚩 NEW: Extracts pure keywords from a KQL query by stripping operators and property names.
   * This is used for the relevancy filter to ensure we don't look for "Title:" inside text.
   */
  private static extractPureKeywords(kql: string): string[] {
    if (!kql) return [];
    const stopWords = ['can', 'you', 'tell', 'me', 'who', 'is', 'the', 'in', 'at', 'on', 'with', 'for', 'what', 'are', 'about', 'and', 'or', 'any', 'some'];
    
    // 1. Remove property filters like Title:"..." or (Description:"...")
    let clean = kql.replace(/[a-z0-9]+:"[^"]*"/gi, (match) => {
      const parts = match.split(':');
      return parts.length > 1 ? parts[1].replace(/"/g, '') : match;
    });

    // 2. Remove isolated field names and operators
    clean = clean.replace(/\b(Title|Description|AllText|AllContacts|ServiceLine|OR|AND|NOT|Path)\b/gi, ' ');
    
    // 3. Remove KQL special characters
    clean = clean.replace(/[:"()*]/g, ' ');

    return clean.toLowerCase()
      .split(/\s+/)
      .filter(w => w.length > 1 && !stopWords.includes(w));
  }

  /**
   * Uses /_api/search/query to retrieve the top-N most relevant documents.
   */
  public async searchDocuments(query: string, filter?: string): Promise<ISearchResult[]> {
    const searchCacheKey = `${query || '*'}|${filter || ''}`;
    const cachedSearch = this.searchResultCache.get(searchCacheKey);
    if (cachedSearch && Date.now() - cachedSearch.savedAt < SharePointSearchService.SEARCH_CACHE_TTL_MS) {
      return cachedSearch.results;
    }

    const pendingSearch = this.pendingSearches.get(searchCacheKey);
    if (pendingSearch) {
      return pendingSearch;
    }

    const searchPromise = this.searchDocumentsUncached(query, filter).then((results) => {
      this.searchResultCache.set(searchCacheKey, { savedAt: Date.now(), results });
      return results;
    }).finally(() => {
      this.pendingSearches.delete(searchCacheKey);
    });

    this.pendingSearches.set(searchCacheKey, searchPromise);
    return searchPromise;
  }

  private async searchDocumentsUncached(query: string, filter?: string): Promise<ISearchResult[]> {
    const execute = async (searchQuery: string, useBarebones = false): Promise<ISearchResult[]> => {
      try {
        const { kmPath, whosWhoPath } = await this.getCachedSearchPaths();
        const safeQuery = (searchQuery || '*').replace(/'/g, "''");
        let pathScope = '';
        const paths: string[] = [];

        if (kmPath) {
          paths.push(`Path:"${kmPath}"`);
        }

        if (whosWhoPath) {
          paths.push(`Path:"${whosWhoPath}"`);
        }
        // 🚩 INTELLIGENT SCOPE: For general knowledge queries, only search the KM Library.
        // Only include Who's Who if the query looks like a personnel search.
        const isPersonnelSearch = /\b(who|is|team|department|person|contact|email|profile)\b/i.test(safeQuery);
        
        if (paths.length > 0) {
          const filteredPaths = isPersonnelSearch 
            ? paths 
            : paths.filter(p => !p.toLowerCase().includes('who'));
          
          if (filteredPaths.length > 0) {
            pathScope = `(${filteredPaths.join(' OR ')})`;
          }
        }

        const selectProps = useBarebones
          ? [SEARCH_PROPERTIES.title, SEARCH_PROPERTIES.path, COLUMN_NAMES.description, SEARCH_PROPERTIES.rank, SEARCH_PROPERTIES.listItemId, 'ListID']
          : [
              SEARCH_PROPERTIES.title, 'TitleName', SEARCH_PROPERTIES.path, 'HitHighlightedSummary', COLUMN_NAMES.description,
              'Author', COLUMN_NAMES.author, 'DisplayAuthor', 'AuthorOWSUSER', 'EditorOWSUSER', 'ModifiedBy', COLUMN_NAMES.modifiedBy,
              'FileExtension', 'DocId', SEARCH_PROPERTIES.listItemId, 'ListID', 'LastModifiedTime', SEARCH_PROPERTIES.write, SEARCH_PROPERTIES.rank,
              'AllContactsOWSMTXT', 'AllEmailsOWSMTXT', 'AllTextOWSMTXT', 'ServiceLineOWSTEXT',
              'BUDepartmentOWSCHCS', 'TherapyAreaOWSCHCS', 'DiseaseAreaOWSCHCS', 'ClientOWSCHCS', 'GeographyOWSCHCS', 'DocumentTypeOWSCHCS',
              'BUOWSCHCS', 'BUOWSTEXT', 'DepartmentOWSTEXT', 'ClientOWSTEXT', SEARCH_PROPERTIES.status, SEARCH_PROPERTIES.documentType, SEARCH_PROPERTIES.businessUnit, SEARCH_PROPERTIES.client, SEARCH_PROPERTIES.geography, SEARCH_PROPERTIES.therapyArea,
              COLUMN_NAMES.views, COLUMN_NAMES.likes, 'ViewsOWSNMBR', 'LikesOWSNMBR', COLUMN_NAMES.documentType, COLUMN_NAMES.diseaseArea, COLUMN_NAMES.therapyArea, COLUMN_NAMES.department,
              COLUMN_NAMES.author, 'Author0OWSTEXT', 'Author0OWSUSER', COLUMN_NAMES.status, 'StatusOWSCHCS',
              COLUMN_NAMES.fileLeafRef, COLUMN_NAMES.contentRefreshDate, COLUMN_NAMES.comments, COLUMN_NAMES.downloads,
              COLUMN_NAMES.follow, COLUMN_NAMES.share, COLUMN_NAMES.bookmark, COLUMN_NAMES.sensitiveTerms,
              COLUMN_NAMES.reviewerComments, COLUMN_NAMES.projectId, COLUMN_NAMES.versionFileName,
              COLUMN_NAMES.versionFileType, COLUMN_NAMES.docIcon, COLUMN_NAMES.modifiedBy
            ]; 

        
        // 🚩 UAT REQUIREMENT: Only fetch ACTIVE or UNDER REVIEW documents
        // We expand the filter based on user feedback that "Under Review" files must be discoverable.
        // 🚩 KEEP KQL SIMPLE — complex filters like statusFilter in KQL cause SharePoint to silently return 0 results.
        // Active/Under Review filtering and Who's Who exclusion are applied CLIENT-SIDE after results arrive.
        const fullQuery = pathScope ? `(${safeQuery}) AND ${pathScope}` : safeQuery;

        const searchUrl =
          `${this.siteUrl}/_api/search/query` +
          `?querytext='${fullQuery}'` +
          `&selectproperties='${selectProps.join(',')}'` +
          `&rowlimit=100` +
          `&enablequeryrules=true` +
          `&clienttype='ContentSearchRegular'`;

        const response = await this.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
        
        // 🚩 SELF-HEALING: If primary fails with 500, immediately retry with barebones
        if (!response.ok) {
          if (!useBarebones) {
            return await execute(searchQuery, true);
          }
          return [];
        }

        const json = await response.json();
        const rows: any[] = json?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
        const allMappedResults: ISearchResult[] = rows.map((row, idx) => {
          const cells: Array<{ Key: string; Value: string }> = row.Cells || [];
          
          const summary = SharePointSearchService.cleanSummary(SharePointSearchService.cell(cells, 'HitHighlightedSummary'));
          const title = SharePointSearchService.cell(cells, SEARCH_PROPERTIES.title);
          const description = SharePointSearchService.cell(cells, COLUMN_NAMES.description);

          // 🚩 METADATA EXTRACTION: Map the Managed Properties confirmed by the user
          // We check for OWSCHCS (Taxonomy/Choice), OWSTEXT (Text), and internal variants.
          const department = SharePointSearchService.cell(cells, 'BUDepartmentOWSCHCS') || 
                             SharePointSearchService.cell(cells, 'DepartmentOWSTEXT') || 
                             SharePointSearchService.cell(cells, COLUMN_NAMES.department) ||
                             SharePointSearchService.cell(cells, 'Department') ||
                             SharePointSearchService.cell(cells, SEARCH_PROPERTIES.status) ||
                             SharePointSearchService.cell(cells, SEARCH_PROPERTIES.documentType);

          const client = SharePointSearchService.cell(cells, 'ClientOWSCHCS') || 
                         SharePointSearchService.cell(cells, 'ClientOWSTEXT') ||
                         SharePointSearchService.cell(cells, 'Client') ||
                         SharePointSearchService.cell(cells, SEARCH_PROPERTIES.businessUnit);

          const businessUnit = SharePointSearchService.cell(cells, 'BUOWSCHCS') || 
                               SharePointSearchService.cell(cells, 'BUOWSTEXT') ||
                               SharePointSearchService.cell(cells, COLUMN_NAMES.businessUnit) ||
                               SharePointSearchService.cell(cells, SEARCH_PROPERTIES.client);

          const docType = SharePointSearchService.cell(cells, 'DocumentTypeOWSCHCS') || 
                          SharePointSearchService.cell(cells, 'DocumentType') ||
                          SharePointSearchService.cell(cells, SEARCH_PROPERTIES.geography);

          const therapyArea = SharePointSearchService.cell(cells, 'TherapyAreaOWSCHCS') || 
                              SharePointSearchService.cell(cells, 'TherapyArea') ||
                              SharePointSearchService.cell(cells, SEARCH_PROPERTIES.therapyArea);

          const diseaseArea = SharePointSearchService.cell(cells, 'DiseaseAreaOWSCHCS') || 
                              SharePointSearchService.cell(cells, 'DiseaseArea');

          // 🚩 PERSONNEL & TEAM METADATA: Map exact internal names from Who's Who
          const allText = SharePointSearchService.cell(cells, 'AllTextOWSMTXT') || 
                          SharePointSearchService.cell(cells, 'AllText');
          const contacts = SharePointSearchService.cell(cells, 'AllContactsOWSMTXT') || 
                           SharePointSearchService.cell(cells, 'AllContacts');
          const emails = SharePointSearchService.cell(cells, 'AllEmailsOWSMTXT') || 
                         SharePointSearchService.cell(cells, 'AllEmails');
          const serviceLine = SharePointSearchService.cell(cells, 'ServiceLineOWSTEXT') || 
                              SharePointSearchService.cell(cells, 'ServiceLine');

          const personBio = [contacts, emails, allText, serviceLine, description].filter(Boolean).join('\n');
          
          // 🚩 AI GROUNDING: Prepend metadata to content so the AI "sees" it during reasoning
          const metadataHeader = `[METADATA] Department: ${department || 'N/A'}, Client: ${client || 'N/A'}, BU: ${businessUnit || 'N/A'}, Document Type: ${docType || 'N/A'}, Therapy Area: ${therapyArea || 'N/A'}, Disease Area: ${diseaseArea || 'N/A'}\n\n`;
          
          const rawContent = summary || personBio || description || title || 'No content preview available.';
          const content = metadataHeader + PromptBuilder.truncateContent(rawContent, 4000); 

          // 🚩 DYNAMIC RESOLUTION: Prioritize Friendly Display names and clean TitleName
          const rawTitle = SharePointSearchService.cell(cells, 'TitleName') || title;
          const cleanTitle = (rawTitle && rawTitle !== '-' && rawTitle.trim() !== '') ? rawTitle : (title || 'Untitled');

          const docUrl = SharePointSearchService.cell(cells, 'Path');
          const lowerPath = (docUrl || '').toLowerCase();
          const rawAuthor = SharePointSearchService.cell(cells, 'Author') || 
                            SharePointSearchService.cell(cells, COLUMN_NAMES.author) ||
                            SharePointSearchService.cell(cells, 'Author0OWSTEXT') ||
                            SharePointSearchService.cell(cells, 'Author0OWSUSER') ||
                            SharePointSearchService.cell(cells, 'DisplayAuthor') || 
                            SharePointSearchService.cell(cells, 'AuthorOWSUSER') || 
                            SharePointSearchService.cell(cells, 'EditorOWSUSER') ||
                            SharePointSearchService.cell(cells, 'ModifiedBy') ||
                            SharePointSearchService.cell(cells, COLUMN_NAMES.modifiedBy) ||
                            SharePointSearchService.cell(cells, 'CreatedBy') ||
                            SharePointSearchService.cell(cells, 'LastModifiedBy') ||
                            // Catch-all: find any cell that looks like it contains an author
                            (cells.find(c => c.Key.toLowerCase().includes('author') || c.Key.toLowerCase().includes('editor'))?.Value) ||
                            (lowerPath.includes("/lists/who") ? contacts : '');
          
          let cleanAuthor = (rawAuthor || '').replace(/i:0#.f\|membership\|/gi, '').trim();
          
          // Handle ID;#Name or Name;#ID patterns
          if (cleanAuthor.includes(';#')) {
            const parts = cleanAuthor.split(';#').filter(p => p.length > 0);
            // If first part is a number, take second part. Otherwise take first.
            cleanAuthor = (/^\d+$/.test(parts[0])) ? (parts[1] || parts[0]) : parts[0];
          } else if (cleanAuthor.includes(';')) {
             // Handle simple semicolon separation
             cleanAuthor = cleanAuthor.split(';')[0].trim();
          }

          // Pattern Match: "email@domain.com | Display Name | GUID" or "email@domain.com, Display Name"
          if (cleanAuthor.includes('|') || cleanAuthor.includes(',')) {
            const parts = cleanAuthor.split(/[|,]/).map(p => p.trim()).filter(p => p.length > 0);
            
            // Priority 1: Find something that looks like a name and NOT a login/GUID
            const likelyName = parts.find(p => !p.includes('@') && !/^[0-9a-f-]{36}$/i.test(p) && p.length > 2 && !p.includes('\\'));
            
            if (likelyName) {
              cleanAuthor = likelyName;
            } else {
              // Fallback to first part if it's not obviously a technical string
              cleanAuthor = parts[0] || '';
            }
          }

          // 🚩 DYNAMIC LIST DISCOVERY: Extract list name from path to avoid "ID not found" errors
          // e.g. .../sites/InSkytes/KM%20Review%20Hub/test.pdf -> KM Review Hub
          let listTitle = LIBRARY_NAMES.kmDataHub;
          try {
            const pathParts = docUrl.split('/');
            const formsIndex = pathParts.findIndex(p => p.toLowerCase() === 'forms');
            const listsIndex = pathParts.findIndex(p => p.toLowerCase() === 'lists');
            
            if (formsIndex > 0) {
              listTitle = decodeURIComponent(pathParts[formsIndex - 1]);
            } else if (listsIndex > 0 && pathParts.length > listsIndex + 1) {
              listTitle = decodeURIComponent(pathParts[listsIndex + 1]);
            } else {
              // 🚩 SMART EXTRACTION: If no /Forms/ or /Lists/, the library is usually the 2nd segment after the site root
              // Site Root: .../sites/InSkytes/ -> segments [..., 'sites', 'InSkytes']
              // Document: .../sites/InSkytes/LibraryName/File.pdf -> segments [..., 'sites', 'InSkytes', 'LibraryName', 'File.pdf']
              const siteIndex = pathParts.findIndex(p => p.toLowerCase() === 'sites');
              if (siteIndex !== -1 && pathParts.length > siteIndex + 3) {
                listTitle = decodeURIComponent(pathParts[siteIndex + 2]);
              }
            }
          } catch (e) {
             console.warn("Failed to extract list title from path:", docUrl);
          }

          // 🚩 ID REFINEMENT: Prefer ListItemID for detail page compatibility. 
          // Standard SharePoint IDs are small integers (Int32). 
          // ⚠️ SAFETY: Use regex length check to avoid JS precision loss on 19-digit WorkIds.
          const rawId = SharePointSearchService.cell(cells, SEARCH_PROPERTIES.listItemId);
          const isStandardId = rawId && /^\d{1,10}$/.test(rawId) && Number(rawId) < 2147483647;
          
          const numericId = isStandardId ? rawId : '0';
          const fileName = SharePointSearchService.cell(cells, COLUMN_NAMES.fileLeafRef) || cleanTitle;
          const modifiedByTitle = SharePointSearchService.cell(cells, COLUMN_NAMES.modifiedBy) ||
                                  SharePointSearchService.cell(cells, COLUMN_NAMES.modifiedBy) ||
                                  SharePointSearchService.cell(cells, 'ModifiedBy');

          return {
            id: numericId,
            title: cleanTitle,
            content: content,
            url: docUrl,
            author: cleanAuthor || 'Internal',
            publishedDate: SharePointSearchService.cell(cells, 'LastModifiedTime') || 
                           SharePointSearchService.cell(cells, SEARCH_PROPERTIES.write) ||
                           SharePointSearchService.cell(cells, COLUMN_NAMES.created) ||
                           SharePointSearchService.cell(cells, COLUMN_NAMES.modified),
            fileType: (SharePointSearchService.cell(cells, 'FileExtension') || 'DOC').toUpperCase(),
            description: summary || description,
            rank: Number(SharePointSearchService.cell(cells, SEARCH_PROPERTIES.rank)) || 0,
            status: 'Active',
            listTitle: listTitle,
            listId: SharePointSearchService.cell(cells, 'ListID'),
            fileName,
            contentRefreshDate: SharePointSearchService.cell(cells, COLUMN_NAMES.contentRefreshDate),
            commentsCount: SharePointSearchService.cellNumber(cells, COLUMN_NAMES.comments, COLUMN_NAMES.comments, 'CommentsOWSNMBR'),
            downloadsCount: SharePointSearchService.cellNumber(cells, COLUMN_NAMES.downloads, COLUMN_NAMES.downloads, 'DownloadsOWSNMBR'),
            downloads: SharePointSearchService.cellNumber(cells, COLUMN_NAMES.downloads, COLUMN_NAMES.downloads, 'DownloadsOWSNMBR'),
            follow: SharePointSearchService.cellBoolean(cells, COLUMN_NAMES.follow, COLUMN_NAMES.follow),
            share: SharePointSearchService.cellBoolean(cells, COLUMN_NAMES.share, COLUMN_NAMES.share),
            bookmark: SharePointSearchService.cellBoolean(cells, COLUMN_NAMES.bookmark, COLUMN_NAMES.bookmark),
            sensitiveTerms: SharePointSearchService.cellArray(cells, COLUMN_NAMES.sensitiveTerms, COLUMN_NAMES.sensitiveTerms),
            reviewerComments: SharePointSearchService.cell(cells, COLUMN_NAMES.reviewerComments),
            projectId: SharePointSearchService.cell(cells, COLUMN_NAMES.projectId),
            versionFileName: SharePointSearchService.cell(cells, COLUMN_NAMES.versionFileName),
            versionFileType: SharePointSearchService.cell(cells, COLUMN_NAMES.versionFileType),
            modifiedBy: modifiedByTitle ? {
              title: modifiedByTitle,
              email: '',
              id: 0
            } : undefined,
            docIcon: SharePointSearchService.cell(cells, COLUMN_NAMES.docIcon),
            businessUnit: businessUnit,
            documentType: docType,
            department: department,
            geography: SharePointSearchService.cell(cells, 'GeographyOWSCHCS') ||
                       SharePointSearchService.cell(cells, COLUMN_NAMES.geography) ||
                       SharePointSearchService.cell(cells, SEARCH_PROPERTIES.therapyArea),
            diseaseArea: diseaseArea,
            therapyArea: therapyArea,
            client: client,
            viewsCount: Number(SharePointSearchService.cell(cells, 'ViewsOWSNMBR')) || Number(SharePointSearchService.cell(cells, COLUMN_NAMES.views)) || 0,
            likesCount: Number(SharePointSearchService.cell(cells, 'LikesOWSNMBR')) || Number(SharePointSearchService.cell(cells, COLUMN_NAMES.likes)) || 0,
            // Store raw status for client-side filtering below
            _rawStatus: (
              SharePointSearchService.cell(cells, 'StatusOWSCHCS') ||
              SharePointSearchService.cell(cells, COLUMN_NAMES.status) ||
              SharePointSearchService.cell(cells, SEARCH_PROPERTIES.status) || ''
            ).toLowerCase()
          } as ISearchResult;
        });

        // 🚩 CLIENT-SIDE FILTER 1: Status — Allow Active or Under Review only.
        // If the Status field is not indexed, _rawStatus will be '' — let those through.
        const allowedStatuses = ['active'];
        let filteredResults = allMappedResults.filter(r => {
          const s = ((r as any)._rawStatus || '').toLowerCase();
          return s === '' || allowedStatuses.some(allowed => s.includes(allowed));
        });

        // 🚩 CLIENT-SIDE FILTER 2: Who's Who — exclude personnel files for non-personnel queries.
        const isPersonnelQuery = /\b(who is|who are|head of|lead|manager|director|contact|email|team member)\b/i.test(safeQuery);
        if (!isPersonnelQuery) {
          filteredResults = filteredResults.filter(r => {
            const lp = (r.url || '').toLowerCase();
            const lt = (r.title || '').toLowerCase();
            return !lp.includes('/lists/who') &&
                   !lp.includes('who%27s%20who') &&
                   !lt.includes("who's who");
          });
        }

        const rankedResults = SearchRankingService.rankResults(safeQuery, filteredResults);
        return rankedResults.length > 0
          ? rankedResults
          : filteredResults.sort((a, b) => {
            const scoreA = this.scoreResultAgainstQuery(safeQuery, a);
            const scoreB = this.scoreResultAgainstQuery(safeQuery, b);
            const dateA = new Date(a.publishedDate || 0).getTime();
            const dateB = new Date(b.publishedDate || 0).getTime();
            return scoreB !== scoreA ? scoreB - scoreA : dateB - dateA;
          });
      } catch (err) {
        if (!useBarebones) {
          return await execute(searchQuery, true);
        }
        console.error('SharePointSearchService.execute terminal error:', err);
        return [];
      }
    };

    const metadataResultsPromise = this.searchMetadataListItems(query);

    // 1. Try AI-Refined primary search first (Path Restricted to Library + List)
    let results = await execute(query);

    // 2. 🚩 FALLBACK: Broad Site Discovery
    // If the library-restricted search returns very few or irrelevant results, 
    // broaden the scope to the whole site for better discovery.
    if (results.length < 2 && query && query !== '*') {
      const siteWideResults = await execute(query, true); // Barebones = No Path restriction
      
      // Merge results, prioritizing the site-wide discovery if it found better candidates
      if (siteWideResults.length > 0) {
        // Simple deduplication by URL
        const existingUrls = new Set(results.map(r => r.url));
        const newResults = siteWideResults.filter(r => !existingUrls.has(r.url));
        results = [...results, ...newResults].slice(0, 10);
      }
    }

    const metadataResults = await metadataResultsPromise;
    if (metadataResults.length > 0) {
      const existingKeys = new Set(metadataResults.map((result) => String(result.id || result.url || result.title)));
      const remainingResults = results.filter((result) =>
        !existingKeys.has(String(result.id || result.url || result.title))
      );
      results = [...metadataResults, ...remainingResults].slice(0, 10);
    }

    return results;
  }

  public async getDocumentCount(
    status: string,
    category?: string
  ): Promise<number> {
    try {
      const absUrl = await this.getRootFolderUrl();
      if (!absUrl) return 0;
      const conditions = [`Path:"${absUrl}/*"`];
      if (status) conditions.push(`${SEARCH_PROPERTIES.status}:${status}`);
      if (category) conditions.push(`${SEARCH_PROPERTIES.documentType}:"${SharePointSearchService.escapeKqlPhrase(category)}"`);
      const querytext = conditions.join(' AND ');
      const url =
        `${this.siteUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(querytext)}'` +
        `&rowlimit=1` +
        `&selectproperties='${SEARCH_PROPERTIES.listItemId}'` +
        `&clienttype='ContentSearchRegular'`;
      const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!resp.ok) return 0;
      const data = await resp.json();
      return data?.PrimaryQueryResult?.RelevantResults?.TotalRows ?? 0;
    } catch {
      return 0;
    }
  }

  public async getLibraryRootFolderUrl(): Promise<string> {
    return this.getRootFolderUrl();
  }

  public async getDocumentCountOnly(status: string): Promise<number> {
    return this.getDocumentCount(status);
  }

  public async getLibraryItemCount(): Promise<number> {
    try {
      const listTitle = SharePointSearchService.escapeODataValue(decodeURIComponent(this.libraryPath));
      const url =
        `${this.siteUrl}/_api/web/lists/getbytitle('${listTitle}')?$select=ItemCount`;
      const resp = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=verbose' } }
      );
      if (!resp.ok) {
        return 0;
      }
      const json = await resp.json();
      const count = Number(json?.d?.ItemCount || json?.ItemCount || 0) || 0;
      return count;
    } catch {
      return 0;
    }
  }

  public async getLibraryStatusCounts(): Promise<ILibraryStatusCountSummary | null> {
    const cacheKey = `${this.siteUrl}|${this.libraryPath}|statusCounts`;
    const cached = SharePointSearchService.libraryStatusCountCache.get(cacheKey);
    if (cached && Date.now() - cached.fetchedAt < SharePointSearchService.SEARCH_CACHE_TTL_MS) {
      return cached.summary;
    }

    try {
      const maxIdResp = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${this.libraryPath}')/items` +
          `?$select=ID&$orderby=ID desc&$top=1`,
        SPHttpClient.configurations.v1
      );
      if (!maxIdResp.ok) {
        return null;
      }

      const maxIdJson = await maxIdResp.json();
      const maxId = Number((maxIdJson.value || maxIdJson?.d?.results || [])[0]?.ID || 0);
      if (!maxId) {
        const emptySummary = { fileCount: 0, folderCount: 0, statusCounts: {} };
        SharePointSearchService.libraryStatusCountCache.set(cacheKey, { fetchedAt: Date.now(), summary: emptySummary });
        return emptySummary;
      }

      const summary: ILibraryStatusCountSummary = {
        fileCount: 0,
        folderCount: 0,
        statusCounts: {}
      };
      const batchSize = 2000;

      for (let startId = 1; startId <= maxId; startId += batchSize) {
        const endId = startId + batchSize - 1;
        let pageUrl =
          `${this.siteUrl}/_api/web/lists/getbytitle('${this.libraryPath}')/items` +
            `?$select=ID,${COLUMN_NAMES.status},FSObjType` +
            `&$filter=${encodeURIComponent(`ID ge ${startId} and ID le ${endId}`)}` +
            `&$top=500`;

        while (pageUrl) {
          const pageResp = await this.spHttpClient.get(
            pageUrl,
            SPHttpClient.configurations.v1
          );

          if (!pageResp.ok) {
            break;
          }

          const pageJson = await pageResp.json();
          (pageJson.value || pageJson?.d?.results || []).forEach((item: any) => {
            if (Number(item?.FSObjType) === 1) {
              summary.folderCount += 1;
              return;
            }

            summary.fileCount += 1;
            const status = String(item?.[COLUMN_NAMES.status] || '(blank)').trim() || '(blank)';
            summary.statusCounts[status] = (summary.statusCounts[status] || 0) + 1;
          });
          pageUrl = pageJson['@odata.nextLink'] || pageJson?.d?.__next || '';
        }
      }

      SharePointSearchService.libraryStatusCountCache.set(cacheKey, { fetchedAt: Date.now(), summary });
      return summary;
    } catch {
      return null;
    }
  }

  public async getStatusCountsBySearch(): Promise<ILibraryStatusCountSummary> {
    try {
      const absUrl = await this.getRootFolderUrl();
      const statuses = ['Active', 'Under Review', 'Reject', 'Archive'];

      const countResults = await Promise.all(
        statuses.map(async (status) => {
          try {
            const conditions: string[] = [`Path:"${absUrl}/*"`];
            const statusValue = status.includes(' ')
              ? `"${status}"`
              : status;
            conditions.push(`${SEARCH_PROPERTIES.status}:${statusValue}`);
            const querytext = conditions.join(' AND ');

            const countUrl =
              `${this.siteUrl}/_api/search/query` +
              `?querytext='${encodeURIComponent(querytext)}'` +
              `&rowlimit=1` +
              `&trimduplicates=false`;

            const resp = await fetch(countUrl, {
              credentials: 'same-origin',
              headers: {
                'Accept': 'application/json;odata=verbose'
              }
            });
            if (!resp.ok) return { status, count: 0 };
            const json = await resp.json();
            const count =
              json?.d?.query?.PrimaryQueryResult?.RelevantResults?.TotalRows || 0;
            return { status, count };
          } catch {
            return { status, count: 0 };
          }
        })
      );

      const statusCounts: Record<string, number> = {};
      let fileCount = 0;
      countResults.forEach(({ status, count }) => {
        statusCounts[status] = count;
        fileCount += count;
      });

      return {
        fileCount,
        folderCount: 0,
        statusCounts
      };
    } catch {
      return { fileCount: 0, folderCount: 0, statusCounts: {} };
    }
  }

  public async getListItemCount(
    status?: string | string[],
    category?: string,
    userId?: number,
    authorField: string = COLUMN_NAMES.author
  ): Promise<number> {
    try {
      if (userId) {
        const authorIdFields = ['AuthorId'];
        const ids = new Set<number>();

        for (const authorIdField of authorIdFields) {
          const filters = [`${authorIdField} eq ${userId}`];
          const statusCondition = SharePointSearchService.buildODataStatusCondition(COLUMN_NAMES.status, status);
          if (statusCondition) {
            filters.push(statusCondition);
          }

          try {
            let pageUrl =
              `${this.siteUrl}/_api/web/lists/getbytitle('${this.libraryPath}')/items` +
              `?$select=ID` +
              `&$filter=${encodeURIComponent(filters.join(' and '))}` +
              `&$top=500`;

            while (pageUrl) {
              const resp = await this.spHttpClient.get(pageUrl, SPHttpClient.configurations.v1);
              if (!resp.ok) {
                break;
              }

              const json = await resp.json();
              (json.value || json?.d?.results || []).forEach((item: any) => {
                const id = Number(item.ID || item.Id);
                if (id > 0) ids.add(id);
              });
              pageUrl = json['@odata.nextLink'] || json?.d?.__next || '';
            }
          } catch {
            continue;
          }
        }

        return ids.size;
      }

      const absUrl = await this.getRootFolderUrl();
      if (!absUrl) return 0;
      const conditions: string[] = [
        `Path:"${absUrl}/*"`
      ];
      const statusCondition = SharePointSearchService.buildKqlStatusCondition(status);
      if (statusCondition) {
        conditions.push(statusCondition);
      }
      if (category) {
        conditions.push(`${SEARCH_PROPERTIES.documentType}:"${SharePointSearchService.escapeKqlPhrase(category)}"`);
      }

      const querytext = conditions.join(' AND ');
      const countUrl =
        `${this.siteUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(querytext)}'` +
        `&rowlimit=1` +
        `&selectproperties='${SEARCH_PROPERTIES.listItemId}'` +
        `&trimduplicates=false` +
        `&clienttype='ContentSearchRegular'`;

      const resp = await this.spHttpClient.get(
        countUrl,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) {
        return 0;
      }
      const json = await resp.json();
      const totalRows =
        json?.d?.query?.PrimaryQueryResult?.RelevantResults?.TotalRows ||
        json?.PrimaryQueryResult?.RelevantResults?.TotalRows ||
        0;
      return Number(totalRows) || 0;
    } catch {
      return 0;
    }
  }

  // Get document IDs from Search API with status filter.
  public async getDocumentIdsByStatus(
    status: string | string[],
    startRow: number = 0,
    rowLimit: number = 500,
    sortField: string = SEARCH_PROPERTIES.write,
    sortDirection: 'ascending' | 'descending' = 'descending'
  ): Promise<{ ids: number[]; totalRows: number }> {
    try {
      const absUrl = await this.getRootFolderUrl();
      if (!absUrl) return { ids: [], totalRows: 0 };

      const conditions: string[] = [
        `Path:"${absUrl}/*"`
      ];
      const statusCondition = SharePointSearchService.buildKqlStatusCondition(status);
      if (statusCondition) {
        conditions.push(statusCondition);
      }

      const querytext = conditions.join(' AND ');
      const safeRowLimit = Math.min(Math.max(rowLimit, 1), 500);
      const idsUrl =
        `${this.siteUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(querytext)}'` +
        `&selectproperties='${SEARCH_PROPERTIES.listItemId},${SEARCH_PROPERTIES.path}'` +
        `&rowlimit=${safeRowLimit}` +
        `&startrow=${startRow}` +
        `${sortField ? `&sortlist='${sortField}:${sortDirection}'` : ''}` +
        `&trimduplicates=false` +
        `&clienttype='ContentSearchRegular'`;
      const idsResp = await this.spHttpClient.get(idsUrl, SPHttpClient.configurations.v1);
      if (!idsResp.ok) return { ids: [], totalRows: 0 };
      const idsData = await idsResp.json();
      const totalRows = idsData?.PrimaryQueryResult?.RelevantResults?.TotalRows ?? 0;
      const rows = idsData?.PrimaryQueryResult?.RelevantResults?.Table?.Rows ?? [];

      const ids: number[] = rows
        .map((row: any) => {
          const cells = row.Cells || [];
          const obj: Record<string, string> = {};
          cells.forEach((c: any) => { obj[c.Key] = c.Value; });
          return Number(obj[SEARCH_PROPERTIES.listItemId]);
        })
        .filter((id: number) => Number.isFinite(id) && id > 0);

      return { ids, totalRows };
    } catch {
      return { ids: [], totalRows: 0 };
    }
  }

  public async getDocumentIdsByStatusRest(
    status: string | string[],
    startRow: number = 0,
    rowLimit: number = 30,
    statusField: string = COLUMN_NAMES.status,
    sortField: string = COLUMN_NAMES.created,
    sortDirection: 'asc' | 'desc' = 'desc'
  ): Promise<{ ids: number[]; totalRows: number }> {
    try {
      const statusCondition = SharePointSearchService.buildODataStatusCondition(statusField, status);
      if (!statusCondition) {
        return { ids: [], totalRows: 0 };
      }

      const safeRowLimit = Math.min(Math.max(rowLimit, 1), 500);
      const requestedEnd = startRow + safeRowLimit;
      const listTitle = SharePointSearchService.escapeODataValue(decodeURIComponent(this.libraryPath));
      const safeSortField = sortField || COLUMN_NAMES.created;
      let pageUrl =
        `${this.siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items` +
        `?$select=ID` +
        `&$filter=${encodeURIComponent(statusCondition)}` +
        `&$orderby=${safeSortField} ${sortDirection},ID ${sortDirection}` +
        `&$top=500`;
      const orderedIds: number[] = [];

      while (pageUrl && orderedIds.length < requestedEnd) {
        const resp = await this.spHttpClient.get(pageUrl, SPHttpClient.configurations.v1);
        if (!resp.ok) {
          break;
        }

        const json = await resp.json();
        (json.value || json?.d?.results || []).forEach((item: any) => {
          const id = Number(item.ID || item.Id);
          if (id > 0) {
            orderedIds.push(id);
          }
        });
        pageUrl = json['@odata.nextLink'] || json?.d?.__next || '';
      }

      return {
        ids: orderedIds.slice(startRow, requestedEnd),
        totalRows: orderedIds.length
      };
    } catch {
      return { ids: [], totalRows: 0 };
    }
  }

  public async getActiveSearchIdsByIds(ids: number[], category?: string, extraConditions: string[] = []): Promise<number[] | null> {
    const uniqueIds = Array.from(new Set(ids.filter((id) => Number.isFinite(id) && id > 0)));
    if (uniqueIds.length === 0) return [];

    try {
      const absUrl = await this.getRootFolderUrl();
      if (!absUrl) return [];

      const chunks: number[][] = [];
      const chunkSize = 30;
      for (let i = 0; i < uniqueIds.length; i += chunkSize) {
        chunks.push(uniqueIds.slice(i, i + chunkSize));
      }

      const chunkResults = await Promise.all(chunks.map(async (chunk) => {
        const idQuery = chunk.map((id) => `${SEARCH_PROPERTIES.listItemId}:${id}`).join(' OR ');
        const conditions = [
          `Path:"${absUrl}/*"`,
          `${SEARCH_PROPERTIES.status}:Active`,
          `(${idQuery})`
        ];
        if (category) {
          conditions.push(`${SEARCH_PROPERTIES.documentType}:"${SharePointSearchService.escapeKqlPhrase(category)}"`);
        }
        conditions.push(...extraConditions.filter(Boolean));
        const querytext = conditions.join(' AND ');
        const url =
          `${this.siteUrl}/_api/search/query` +
          `?querytext='${encodeURIComponent(querytext)}'` +
          `&rowlimit=${chunk.length}` +
          `&selectproperties='${SEARCH_PROPERTIES.listItemId}'` +
          `&trimduplicates=false` +
          `&clienttype='ContentSearchRegular'`;

        const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!resp.ok) {
          return null;
        }

        const json = await resp.json();
        const relevantResults =
          json?.d?.query?.PrimaryQueryResult?.RelevantResults ||
          json?.PrimaryQueryResult?.RelevantResults;
        const rows = relevantResults?.Table?.Rows?.results || relevantResults?.Table?.Rows || [];
        return rows
          .map((row: any) => {
            const cells = row.Cells?.results || row.Cells || [];
            return Number(SharePointSearchService.cell(cells, SEARCH_PROPERTIES.listItemId));
          })
          .filter((id: number) => Number.isFinite(id) && id > 0);
      }));

      if (chunkResults.some((result) => result === null)) return null;
      return chunkResults.reduce((acc: number[], result) => acc.concat(result || []), []);
    } catch {
      return null;
    }
  }

  public async getSearchIdsByIds(ids: number[], status?: string | string[]): Promise<number[] | null> {
    const uniqueIds = Array.from(new Set(ids.filter((id) => Number.isFinite(id) && id > 0)));
    if (uniqueIds.length === 0) return [];

    try {
      const absUrl = await this.getRootFolderUrl();
      if (!absUrl) return [];

      const chunks: number[][] = [];
      const chunkSize = 30;
      for (let i = 0; i < uniqueIds.length; i += chunkSize) {
        chunks.push(uniqueIds.slice(i, i + chunkSize));
      }

      const chunkResults = await Promise.all(chunks.map(async (chunk) => {
        const idQuery = chunk.map((id) => `${SEARCH_PROPERTIES.listItemId}:${id}`).join(' OR ');
        const conditions = [
          `Path:"${absUrl}/*"`,
          `(${idQuery})`
        ];
        const statusCondition = SharePointSearchService.buildKqlStatusCondition(status);
        if (statusCondition) {
          conditions.push(statusCondition);
        }

        const querytext = conditions.join(' AND ');
        const url =
          `${this.siteUrl}/_api/search/query` +
          `?querytext='${encodeURIComponent(querytext)}'` +
          `&rowlimit=${chunk.length}` +
          `&selectproperties='${SEARCH_PROPERTIES.listItemId}'` +
          `&trimduplicates=false` +
          `&clienttype='ContentSearchRegular'`;

        const resp = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
        if (!resp.ok) {
          return null;
        }

        const json = await resp.json();
        const relevantResults =
          json?.d?.query?.PrimaryQueryResult?.RelevantResults ||
          json?.PrimaryQueryResult?.RelevantResults;
        const rows = relevantResults?.Table?.Rows?.results || relevantResults?.Table?.Rows || [];
        return rows
          .map((row: any) => {
            const cells = row.Cells?.results || row.Cells || [];
            return Number(SharePointSearchService.cell(cells, SEARCH_PROPERTIES.listItemId));
          })
          .filter((id: number) => Number.isFinite(id) && id > 0);
      }));

      if (chunkResults.some((result) => result === null)) return null;
      return chunkResults.reduce((acc: number[], result) => acc.concat(result || []), []);
    } catch {
      return null;
    }
  }

  public async getDocumentIdsByAuthor(
    userId: number,
    authorField: string,
    startRow: number = 0,
    rowLimit: number = 30,
    status?: string | string[],
    sortDirection: 'asc' | 'desc' = 'desc',
  ): Promise<{ ids: number[]; totalRows: number }> {
    try {
      const authorIdFields = ['AuthorId'];
      const rowsById = new Map<number, number>();

      for (const authorIdField of authorIdFields) {
        const filters = [`${authorIdField} eq ${userId}`];
        const statusCondition = SharePointSearchService.buildODataStatusCondition(COLUMN_NAMES.status, status);
        if (statusCondition) {
          filters.push(statusCondition);
        }

        try {
          let pageUrl =
            `${this.siteUrl}/_api/web/lists/getbytitle('${this.libraryPath}')/items` +
            `?$select=ID,${COLUMN_NAMES.created}` +
            `&$filter=${encodeURIComponent(filters.join(' and '))}` +
            `&$orderby=${COLUMN_NAMES.created} ${sortDirection}` +
            `&$top=500`;

          while (pageUrl) {
            const pageResp = await this.spHttpClient.get(pageUrl, SPHttpClient.configurations.v1);
            if (!pageResp.ok) {
              console.warn(`Author field ${authorIdField} query failed, skipping`, pageResp.status);
              break;
            }

            const pageData = await pageResp.json();
            (pageData.value || pageData?.d?.results || []).forEach((item: any) => {
              const id = Number(item.ID || item.Id);
              if (id > 0) {
                rowsById.set(id, new Date(item[COLUMN_NAMES.created] || 0).getTime() || 0);
              }
            });
            pageUrl = pageData['@odata.nextLink'] || pageData?.d?.__next || '';
          }
        } catch (error) {
          console.warn(`Author field ${authorIdField} query failed, skipping`, error);
        }
      }

      const pageIds = Array.from(rowsById.entries())
        .sort((left, right) => sortDirection === 'asc' ? left[1] - right[1] : right[1] - left[1])
        .map(([id]) => id)
        .slice(startRow, startRow + rowLimit);

      return { ids: pageIds, totalRows: pageIds.length };
    } catch {
      return { ids: [], totalRows: 0 };
    }
  }

  public async getDocumentIdsByAuthor0Search(
    userEmail: string,
    sortDirection: string = 'descending',
    startRow: number = 0,
    rowLimit: number = 30,
    status?: string | string[]
  ): Promise<number[]> {
    try {
      const absUrl = await this.getRootFolderUrl();
      const normalizedEmail = (userEmail || '').trim();
      if (!absUrl || !normalizedEmail) return [];

      const searchSortDirection = sortDirection.toLowerCase() === 'ascending' || sortDirection.toLowerCase() === 'asc'
        ? 'ascending'
        : 'descending';
      const author0ClaimsLogin = `i:0#.f|membership|${normalizedEmail}`;
      const conditions = [
        `Path:"${absUrl}/*"`,
        `Author0OWSUSER:"${SharePointSearchService.escapeKqlPhrase(author0ClaimsLogin)}"`
      ];
      const statusCondition = SharePointSearchService.buildKqlStatusCondition(status);
      if (statusCondition) {
        conditions.push(statusCondition);
      }
      const querytext = conditions.join(' AND ');
      const safeRowLimit = Math.min(Math.max(rowLimit, 1), 500);
      const sortDirectionValue = searchSortDirection === 'descending' ? 1 : 0;
      const searchUrl = `${this.siteUrl}/_api/search/postquery`;
      const body = JSON.stringify({
        request: {
          __metadata: { type: 'Microsoft.Office.Server.Search.REST.SearchRequest' },
          Querytext: querytext,
          SelectProperties: { results: [SEARCH_PROPERTIES.listItemId] },
          RowLimit: safeRowLimit,
          StartRow: startRow,
          SortList: { results: [{ Property: 'Created', Direction: sortDirectionValue }] },
          TrimDuplicates: false,
          ClientType: 'ContentSearchRegular'
        }
      });

      const idsResp = await this.spHttpClient.post(
        searchUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body
        }
      );
      if (!idsResp.ok) {
        return [];
      }

      const json = await idsResp.json();
      const relevantResults =
        json?.d?.postquery?.PrimaryQueryResult?.RelevantResults ||
        json?.postquery?.PrimaryQueryResult?.RelevantResults ||
        json?.d?.query?.PrimaryQueryResult?.RelevantResults ||
        json?.PrimaryQueryResult?.RelevantResults;
      const rows =
        json?.d?.postquery?.PrimaryQueryResult?.RelevantResults?.Table?.Rows?.results ||
        relevantResults?.Table?.Rows?.results ||
        relevantResults?.Table?.Rows ||
        [];

      return rows
        .map((row: any) => {
          const cells = row.Cells?.results || row.Cells || [];
          return Number(SharePointSearchService.cell(cells, SEARCH_PROPERTIES.listItemId));
        })
        .filter((id: number) => Number.isFinite(id) && id > 0);
    } catch {
      return [];
    }
  }

  public async getAuthor0DocumentCount(userEmail: string, status?: string | string[]): Promise<number> {
    try {
      const absUrl = await this.getRootFolderUrl();
      const normalizedEmail = (userEmail || '').trim();
      if (!absUrl || !normalizedEmail) return 0;

      const author0ClaimsLogin = `i:0#.f|membership|${normalizedEmail}`;
      const conditions = [
        `Path:"${absUrl}/*"`,
        `Author0OWSUSER:"${SharePointSearchService.escapeKqlPhrase(author0ClaimsLogin)}"`
      ];
      const statusCondition = SharePointSearchService.buildKqlStatusCondition(status);
      if (statusCondition) {
        conditions.push(statusCondition);
      }
      const querytext = conditions.join(' AND ');
      const searchUrl = `${this.siteUrl}/_api/search/postquery`;
      const body = JSON.stringify({
        request: {
          __metadata: { type: 'Microsoft.Office.Server.Search.REST.SearchRequest' },
          Querytext: querytext,
          SelectProperties: { results: [SEARCH_PROPERTIES.listItemId] },
          RowLimit: 1,
          StartRow: 0,
          SortList: { results: [{ Property: 'Created', Direction: 1 }] },
          TrimDuplicates: false,
          ClientType: 'ContentSearchRegular'
        }
      });

      const resp = await this.spHttpClient.post(
        searchUrl,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body
        }
      );
      if (!resp.ok) {
        return 0;
      }

      const json = await resp.json();
      const relevantResults =
        json?.d?.postquery?.PrimaryQueryResult?.RelevantResults ||
        json?.postquery?.PrimaryQueryResult?.RelevantResults ||
        json?.d?.query?.PrimaryQueryResult?.RelevantResults ||
        json?.PrimaryQueryResult?.RelevantResults;

      return Number(relevantResults?.TotalRows || relevantResults?.TotalRowsIncludingDuplicates || 0) || 0;
    } catch {
      return 0;
    }
  }

  // Get latest document IDs without any filter so new uncrawled documents are included.
  public async getLatestDocumentIds(
    count: number = 100
  ): Promise<number[]> {
    const cacheKey = `${this.siteUrl}::${count}`;
    const cached = SharePointSearchService.latestDocumentIdsCache.get(cacheKey);
    if (cached && Date.now() - cached.fetchedAt < 5 * 60 * 1000) {
      return cached.ids;
    }

    try {
      const resp = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${this.libraryPath}')/items` +
        `?$select=ID` +
        `&$orderby=ID desc` +
        `&$top=${count}`,
        SPHttpClient.configurations.v1
      );
      if (!resp.ok) return [];
      const data = await resp.json();
      const result = (data.value || [])
        .map((item: any) => Number(item.ID || item.Id))
        .filter((id: number) => id > 0);
      SharePointSearchService.latestDocumentIdsCache.set(cacheKey, { fetchedAt: Date.now(), ids: result });
      return result;
    } catch {
      return [];
    }
  }

  // Get latest active document IDs directly from the list so recently activated items
  // do not wait for SharePoint Search indexing. Requires Status and Published to be indexed.
  public async getLatestPublishedDocumentIdsByStatus(
    status: string = 'Active',
    count: number = 15,
    publishedField: string = COLUMN_NAMES.published,
    statusField: string = COLUMN_NAMES.status
  ): Promise<number[]> {
    const safeCount = Math.min(Math.max(count, 1), 50);
    const safeStatusField = statusField || COLUMN_NAMES.status;
    const safePublishedField = publishedField || COLUMN_NAMES.published;
    const cacheKey = [
      this.siteUrl,
      this.libraryPath,
      safeStatusField,
      status,
      safePublishedField,
      safeCount
    ].join('::');
    const cached = SharePointSearchService.latestPublishedDocumentIdsCache.get(cacheKey);
    if (cached && Date.now() - cached.fetchedAt < 30 * 1000) {
      return cached.ids;
    }

    try {
      const filter = [
        `${safeStatusField} eq '${SharePointSearchService.escapeODataValue(status)}'`,
        `${safePublishedField} ne null`
      ].join(' and ');
      const resp = await this.spHttpClient.get(
        `${this.siteUrl}/_api/web/lists/getbytitle('${SharePointSearchService.escapeODataValue(this.libraryPath)}')/items` +
        `?$select=ID` +
        `&$filter=${encodeURIComponent(filter)}` +
        `&$orderby=${safePublishedField} desc,ID desc` +
        `&$top=${safeCount}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=verbose' } }
      );
      if (!resp.ok) {
        return [];
      }

      const data = await resp.json();
      const result = (data.value || data?.d?.results || [])
        .map((item: any) => Number(item.ID || item.Id))
        .filter((id: number) => id > 0);
      SharePointSearchService.latestPublishedDocumentIdsCache.set(cacheKey, { fetchedAt: Date.now(), ids: result });
      return result;
    } catch {
      return [];
    }
  }

  // Fetch full document details by IDs via REST.
  public async getDocumentDetailsByIds(
    ids: number[],
    queryParts: { select: string; expand: string }
  ): Promise<any[]> {
    if (ids.length === 0) return [];
    try {
      const webUrl = this.siteUrl;
      const libraryName = this.libraryPath;
      const detailBatchSize = 20;
      const chunks: number[][] = [];
      for (let i = 0; i < ids.length; i += detailBatchSize) {
        chunks.push(ids.slice(i, i + detailBatchSize));
      }

      const fetchChunk = async (chunk: number[]): Promise<any[]> => {
        try {
          const idFilter = chunk.map((id) => `ID eq ${id}`).join(' or ');
          const buildItemsUrl = (select: string, expand?: string): string =>
            `${webUrl}/_api/web/lists/getbytitle('${libraryName}')/items` +
            `?$select=${select}` +
            `${expand ? `&$expand=${expand}` : ''}` +
            `&$filter=${encodeURIComponent(idFilter)}` +
            `&$top=${chunk.length}`;

          const fetchItems = async (select: string, expand?: string): Promise<Response> =>
            this.spHttpClient.get(
              buildItemsUrl(select, expand),
              SPHttpClient.configurations.v1
            ) as unknown as Promise<Response>;

          const resp = await fetchItems(queryParts.select, queryParts.expand);
          if (resp.ok) {
            const data = await resp.json();
            return data.value || data?.d?.results || [];
          }

          if (chunk.length > 1) {
            const midpoint = Math.ceil(chunk.length / 2);
            const [leftItems, rightItems] = await Promise.all([
              fetchChunk(chunk.slice(0, midpoint)),
              fetchChunk(chunk.slice(midpoint))
            ]);
            return leftItems.concat(rightItems);
          }

          return [];
        } catch {
          if (chunk.length > 1) {
            const midpoint = Math.ceil(chunk.length / 2);
            const [leftItems, rightItems] = await Promise.all([
              fetchChunk(chunk.slice(0, midpoint)),
              fetchChunk(chunk.slice(midpoint))
            ]);
            return leftItems.concat(rightItems);
          }

          return [];
        }
      };

      const chunkResults = await Promise.all(
        chunks.map((chunk) => fetchChunk(chunk))
      );
      return chunkResults.reduce((acc: any[], items: any[]) => acc.concat(items), []);
    } catch {
      return [];
    }
  }

  public async getLatestDocuments(top = 5, filter?: string): Promise<ISearchResult[]> {
    try {
      let pathScope = '';
      const absUrl = await this.getRootFolderUrl();
      if (absUrl) {
        pathScope = `Path:"${absUrl}/*"`;
      }

      const queryParts = [
        pathScope,
        `${SEARCH_PROPERTIES.status}:Active`
      ].filter(Boolean);
      const fullQuery = queryParts.join(' AND ');
      const safeTop = Math.min(Math.max(top, 1), 500);
      const selectProps = [
        SEARCH_PROPERTIES.listItemId,
        SEARCH_PROPERTIES.path,
        SEARCH_PROPERTIES.title,
        SEARCH_PROPERTIES.status,
        SEARCH_PROPERTIES.documentType,
        SEARCH_PROPERTIES.published,
        COLUMN_NAMES.fileLeafRef
      ].join(',');

      const searchUrl =
        `${this.siteUrl}/_api/search/query` +
        `?querytext='${encodeURIComponent(fullQuery)}'` +
        `&selectproperties='${selectProps}'` +
        `&rowlimit=${safeTop}` +
        `&sortlist='${SEARCH_PROPERTIES.published}:descending'` +
        `&clienttype='ContentSearchRegular'`;

      const response = await this.spHttpClient.get(searchUrl, SPHttpClient.configurations.v1);
      if (!response.ok) return [];

      const json = await response.json();
      const rows: any[] = json?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];

      return rows.map((row, idx) => {
        const cells: Array<{ Key: string; Value: string }> = row.Cells || [];
        const title = SharePointSearchService.cell(cells, SEARCH_PROPERTIES.title);
        const path = SharePointSearchService.cell(cells, SEARCH_PROPERTIES.path);
        const fileName = SharePointSearchService.cell(cells, COLUMN_NAMES.fileLeafRef);
        const extensionSource = fileName || path || '';
        const fileType = extensionSource.indexOf('.') !== -1
          ? extensionSource.split('.').pop() || ''
          : '';
        return {
          id: SharePointSearchService.cell(cells, SEARCH_PROPERTIES.listItemId) || String(idx),
          title: title || 'Untitled',
          content: title,
          url: path,
          author: 'Internal',
          publishedDate: SharePointSearchService.cell(cells, SEARCH_PROPERTIES.published),
          fileType: fileType.toUpperCase(),
          description: SharePointSearchService.cell(cells, COLUMN_NAMES.description),
          status: SharePointSearchService.cell(cells, SEARCH_PROPERTIES.status) || 'Active',
          [SEARCH_PROPERTIES.status]: SharePointSearchService.cell(cells, SEARCH_PROPERTIES.status) || 'Active',
          fileName,
          contentRefreshDate: SharePointSearchService.cell(cells, COLUMN_NAMES.contentRefreshDate),
          viewsCount: SharePointSearchService.cellNumber(cells, COLUMN_NAMES.views, COLUMN_NAMES.views, 'ViewsOWSNMBR'),
          likesCount: SharePointSearchService.cellNumber(cells, COLUMN_NAMES.likes, COLUMN_NAMES.likes, 'LikesOWSNMBR'),
          commentsCount: SharePointSearchService.cellNumber(cells, COLUMN_NAMES.comments, COLUMN_NAMES.comments, 'CommentsOWSNMBR'),
          downloadsCount: SharePointSearchService.cellNumber(cells, COLUMN_NAMES.downloads, COLUMN_NAMES.downloads, 'DownloadsOWSNMBR'),
          downloads: SharePointSearchService.cellNumber(cells, COLUMN_NAMES.downloads, COLUMN_NAMES.downloads, 'DownloadsOWSNMBR'),
          tags: [],
          score: 0
        } as ISearchResult;
      });
    } catch {
      return [];
    }
  }

}
