import { SPHttpClient } from "@microsoft/sp-http";

interface IWhosWhoItem {
  Id?: number;
  Title?: string;
  AllContacts?: string;
  AllEmails?: string;
  AllText?: string;
  Description?: string;
  SectionTitles?: string;
  ContentIDs?: string;
}

export interface IWhosWhoContext {
  context: string;
  matchCount: number;
}

export class WhosWhoService {
  private static readonly LIST_TITLE = "Who's Who";
  private static readonly CACHE_TTL_MS = 10 * 60 * 1000;
  private readonly cache = new Map<string, { savedAt: number; value: IWhosWhoContext }>();
  private readonly pendingSearches = new Map<string, Promise<IWhosWhoContext>>();
  private itemsCache: { savedAt: number; items: IWhosWhoItem[] } | null = null;
  private pendingItemsFetch: Promise<IWhosWhoItem[]> | null = null;

  constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly siteUrl: string
  ) {}

  public static shouldSearch(query: string): boolean {
    const normalized = WhosWhoService.normalize(query);
    if (!normalized) {
      return false;
    }

    return /\b(who'?s who|who is|who are|who handles|who owns|who leads|who manages|lead|led|head|manager|owner|contact|contacts|email|mail id|role|roles|directory|team members?|people|person|stakeholder)\b/.test(normalized) ||
      /\b(insights?|details?|tell me more|more about|information about|info about)\b/.test(normalized) ||
      /\b(his|her|their|its)\s+(team|teams|tean|tema)\b/.test(normalized) ||
      /\b(department|function|area|enterprise|commercial|omnichannel|activation)\b.*\b(team|teams|tean|tema|members?|structure|lead|led|head)\b/.test(normalized) ||
      /\b(team|teams|tean|tema)\b.*\b(enterprise|commercial|omnichannel|activation)\b/.test(normalized) ||
      /\b(respective|their|its)\s+(team|teams|tean|tema)\b/.test(normalized) ||
      /\bwhat does\b.*\b(team|teams|tean|tema)\b/.test(normalized) ||
      /\bwhich\b.*\b(team|teams|tean|tema)\b/.test(normalized);
  }

  public async search(query: string): Promise<IWhosWhoContext> {
    const cacheKey = WhosWhoService.normalize(query);
    const cached = this.cache.get(cacheKey);
    if (cached && Date.now() - cached.savedAt < WhosWhoService.CACHE_TTL_MS) {
      return cached.value;
    }

    const pendingSearch = this.pendingSearches.get(cacheKey);
    if (pendingSearch) {
      return pendingSearch;
    }

    const searchPromise = this.executeSearch(query, cacheKey).finally(() => {
      this.pendingSearches.delete(cacheKey);
    });

    this.pendingSearches.set(cacheKey, searchPromise);
    return searchPromise;
  }

  private async executeSearch(query: string, cacheKey: string): Promise<IWhosWhoContext> {
    const terms = this.extractQueryTerms(query);
    if (terms.length === 0) {
      const emptyValue = { context: "", matchCount: 0 };
      this.cache.set(cacheKey, { savedAt: Date.now(), value: emptyValue });
      return emptyValue;
    }

    const items = await this.fetchItems(terms);
    const rankedItems = items
      .map((item) => ({ item, score: this.scoreItem(item, terms) }))
      .filter(({ score }) => score > 0)
      .sort((left, right) => right.score - left.score)
      .slice(0, 5)
      .map(({ item }) => item);

    const value = {
      context: this.buildContext(rankedItems),
      matchCount: rankedItems.length
    };

    this.cache.set(cacheKey, { savedAt: Date.now(), value });
    return value;
  }

  private async fetchItems(terms: string[]): Promise<IWhosWhoItem[]> {
    const cachedItems = await this.getCachedItems();
    if (cachedItems.length > 0) {
      return cachedItems.filter((item) => this.scoreItem(item, terms) > 0);
    }

    return [];
  }

  private async getCachedItems(): Promise<IWhosWhoItem[]> {
    if (this.itemsCache && Date.now() - this.itemsCache.savedAt < WhosWhoService.CACHE_TTL_MS) {
      return this.itemsCache.items;
    }

    if (this.pendingItemsFetch) {
      return this.pendingItemsFetch;
    }

    this.pendingItemsFetch = this.fetchAllItems().then((items) => {
      this.itemsCache = { savedAt: Date.now(), items };
      return items;
    }).finally(() => {
      this.pendingItemsFetch = null;
    });

    return this.pendingItemsFetch;
  }

  private async fetchAllItems(): Promise<IWhosWhoItem[]> {
    const selectFields = "Id,Title,AllContacts,AllEmails,AllText,Description,SectionTitles,ContentIDs";
    const listPath = `/_api/web/lists/getbytitle('${WhosWhoService.escapeODataValue(WhosWhoService.LIST_TITLE)}')/items`;
    const url = `${this.siteUrl}${listPath}?$select=${selectFields}&$top=5000`;
    return this.fetchJsonItems(url);
  }

  private async fetchJsonItems(url: string): Promise<IWhosWhoItem[]> {
    try {
      const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) {
        return [];
      }

      const json = await response.json();
      return json.value || [];
    } catch (error) {
      console.warn("Who's Who lookup failed:", error);
      return [];
    }
  }

  private extractQueryTerms(query: string): string[] {
    const stopWords = new Set([
      "what", "which", "who", "whos", "whose", "does", "do", "did", "the", "and", "for", "with",
      "from", "about", "give", "show", "tell", "me", "team", "teams", "role", "roles", "contact",
      "contacts", "email", "lead", "head", "manager", "handles", "handle", "owner", "directory",
      "insight", "insights", "detail", "details", "more", "information", "info", "his", "her", "their", "its", "tean", "tema",
      "complete", "full", "respective", "previous", "context", "yes", "well"
    ]);

    return Array.from(new Set(
      WhosWhoService.normalize(query)
        .split(/\s+/)
        .filter((term) => term.length >= 3 && !stopWords.has(term))
    )).slice(0, 8);
  }

  private scoreItem(item: IWhosWhoItem, terms: string[]): number {
    const title = WhosWhoService.normalize(item.Title || "");
    const contacts = WhosWhoService.normalize(item.AllContacts || "");
    const emails = WhosWhoService.normalize(item.AllEmails || "");
    const allText = WhosWhoService.normalize(item.AllText || "");
    const description = WhosWhoService.normalize(item.Description || "");
    const sections = WhosWhoService.normalize(item.SectionTitles || "");

    return terms.reduce((score, term) => {
      let nextScore = score;
      if (title.indexOf(term) >= 0) nextScore += 12;
      if (contacts.indexOf(term) >= 0) nextScore += 10;
      if (emails.indexOf(term) >= 0) nextScore += 10;
      if (sections.indexOf(term) >= 0) nextScore += 8;
      if (allText.indexOf(term) >= 0) nextScore += 6;
      if (description.indexOf(term) >= 0) nextScore += 5;
      return nextScore;
    }, 0);
  }

  private buildContext(items: IWhosWhoItem[]): string {
    if (items.length === 0) {
      return "";
    }

    return items.map((item, index) => [
      `=== WHO'S WHO RECORD ${index + 1} ===`,
      `TITLE: ${item.Title || "Untitled"}`,
      `CONTACTS: ${this.cleanValue(item.AllContacts)}`,
      `EMAILS: ${this.cleanValue(item.AllEmails)}`,
      `ROLE/TEXT: ${this.cleanValue(item.AllText)}`,
      `DESCRIPTION: ${this.cleanValue(item.Description)}`,
      `SECTION TITLES: ${this.cleanValue(item.SectionTitles)}`,
      `CONTENT IDS: ${this.cleanValue(item.ContentIDs)}`,
      `=== END WHO'S WHO RECORD ${index + 1} ===`
    ].join("\n")).join("\n\n");
  }

  private cleanValue(value?: string): string {
    return (value || "Not listed")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  private static normalize(value: string): string {
    return (value || "")
      .toLowerCase()
      .replace(/[^a-z0-9@._-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  private static escapeODataValue(value: string): string {
    return (value || "").replace(/'/g, "''");
  }
}
