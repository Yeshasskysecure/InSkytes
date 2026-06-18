import { ISearchResult } from '../models/SearchResult';

type TQueryTerm = {
  value: string;
  quoted: boolean;
};

type TParsedQuery = {
  original: string;
  normalized: string;
  positiveTerms: TQueryTerm[];
  requiredTerms: TQueryTerm[];
  optionalTerms: TQueryTerm[];
  excludedTerms: TQueryTerm[];
  exactPhrases: string[];
  meaningfulTerms: string[];
  isBoolean: boolean;
  isSentence: boolean;
};

const STOP_WORDS = new Set([
  'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'can', 'do', 'does', 'for', 'from',
  'have', 'how', 'if', 'in', 'is', 'it', 'me', 'of', 'on', 'or', 'related', 'that',
  'the', 'this', 'to', 'what', 'when', 'where', 'who', 'with', 'you'
]);

export class SearchRankingService {
  public static rankResults(query: string, results: ISearchResult[]): ISearchResult[] {
    const parsed = this.parseQuery(query);
    if (!parsed.normalized || parsed.normalized === '*') {
      return results;
    }

    return results
      .map((result) => ({
        result,
        ranking: this.scoreResult(parsed, result)
      }))
      .filter(({ ranking }) => !ranking.notTermExcluded && ranking.booleanMatched)
      .sort((left, right) => {
        if (right.ranking.finalRankScore !== left.ranking.finalRankScore) {
          return right.ranking.finalRankScore - left.ranking.finalRankScore;
        }

        const rightSemantic = Number(right.result.semanticScore || right.result.score || right.result.rank || 0);
        const leftSemantic = Number(left.result.semanticScore || left.result.score || left.result.rank || 0);
        if (rightSemantic !== leftSemantic) {
          return rightSemantic - leftSemantic;
        }

        return new Date(right.result.publishedDate || 0).getTime() - new Date(left.result.publishedDate || 0).getTime();
      })
      .map(({ result, ranking }) => ({
        ...result,
        searchRankDebug: {
          titleMatchCount: ranking.titleMatchCount,
          metadataMatchCount: ranking.metadataMatchCount,
          abstractMatchCount: ranking.abstractMatchCount,
          contentMatchCount: ranking.contentMatchCount,
          exactPhraseMatch: ranking.exactPhraseMatch,
          fullSentenceMatch: ranking.fullSentenceMatch,
          notTermExcluded: ranking.notTermExcluded,
          semanticScore: Number(result.semanticScore || result.score || result.rank || 0),
          finalRankScore: ranking.finalRankScore
        }
      }));
  }

  private static scoreResult(parsed: TParsedQuery, result: ISearchResult): {
    titleMatchCount: number;
    metadataMatchCount: number;
    abstractMatchCount: number;
    contentMatchCount: number;
    exactPhraseMatch: boolean;
    fullSentenceMatch: boolean;
    notTermExcluded: boolean;
    booleanMatched: boolean;
    finalRankScore: number;
  } {
    const titleText = this.normalize([result.title].join(' '));
    const metadataText = this.normalize([
      result.businessUnit,
      result.department,
      result.documentType,
      result.client,
      result.geography,
      result.therapyArea,
      result.diseaseArea,
      result.author
    ].filter(Boolean).join(' '));
    const abstractText = this.normalize([result.description, result.abstract].filter(Boolean).join(' '));
    const contentText = this.normalize(result.content || '');
    const allText = `${titleText} ${metadataText} ${abstractText} ${contentText}`;

    const notTermExcluded = parsed.excludedTerms.some((term) => this.matchesLiteral(allText, term));
    const requiredMatched = parsed.requiredTerms.every((term) => this.matchesLiteral(allText, term));
    const optionalMatched = parsed.optionalTerms.length === 0 || parsed.optionalTerms.some((term) => this.matchesLiteral(allText, term));
    const positiveMatched = parsed.positiveTerms.length === 0 || parsed.positiveTerms.some((term) => this.matchesLiteral(allText, term));
    const booleanMatched = parsed.isBoolean ? requiredMatched && optionalMatched : positiveMatched;

    const termsForCounting = parsed.positiveTerms.length > 0
      ? parsed.positiveTerms
      : parsed.meaningfulTerms.map((value) => ({ value, quoted: false }));

    const titleMatchCount = this.countTerms(titleText, termsForCounting);
    const metadataMatchCount = this.countTerms(metadataText, termsForCounting);
    const abstractMatchCount = this.countTerms(abstractText, termsForCounting);
    const contentMatchCount = this.countTerms(contentText, termsForCounting);
    const exactPhraseMatch = parsed.exactPhrases.some((phrase) =>
      titleText.indexOf(phrase) !== -1 ||
      metadataText.indexOf(phrase) !== -1 ||
      abstractText.indexOf(phrase) !== -1 ||
      contentText.indexOf(phrase) !== -1
    );
    const fullSentenceMatch = parsed.isSentence && (
      titleText.indexOf(parsed.normalized) !== -1 ||
      metadataText.indexOf(parsed.normalized) !== -1 ||
      abstractText.indexOf(parsed.normalized) !== -1 ||
      contentText.indexOf(parsed.normalized) !== -1
    );

    let finalRankScore = 0;
    if (titleText === parsed.normalized) finalRankScore += 100000;
    if (titleText.indexOf(parsed.normalized) !== -1) finalRankScore += 60000;
    if (metadataText.indexOf(parsed.normalized) !== -1) finalRankScore += 45000;
    if (fullSentenceMatch) finalRankScore += 90000;
    if (exactPhraseMatch) finalRankScore += 70000;

    finalRankScore += titleMatchCount * 8000;
    finalRankScore += metadataMatchCount * 6000;
    finalRankScore += abstractMatchCount * 3000;
    finalRankScore += contentMatchCount * 1200;
    finalRankScore += Number(result.semanticScore || result.score || result.rank || 0);

    return {
      titleMatchCount,
      metadataMatchCount,
      abstractMatchCount,
      contentMatchCount,
      exactPhraseMatch,
      fullSentenceMatch,
      notTermExcluded,
      booleanMatched,
      finalRankScore
    };
  }

  private static parseQuery(query: string): TParsedQuery {
    const original = String(query || '').trim();
    const normalized = this.normalize(original);
    const tokens = this.parseTokens(original);
    const isBoolean = tokens.some((token) => token.type === 'operator');
    const exactPhrases = tokens
      .filter((token) => token.type === 'term' && token.quoted)
      .map((token) => this.normalize(token.value))
      .filter(Boolean);
    const meaningfulTerms = this.tokenize(normalized).filter((token) => !STOP_WORDS.has(token));
    const isSentence = meaningfulTerms.length >= 5 && original.length >= 35 && !isBoolean;

    const requiredTerms: TQueryTerm[] = [];
    const optionalTerms: TQueryTerm[] = [];
    const excludedTerms: TQueryTerm[] = [];
    const positiveTerms: TQueryTerm[] = [];
    let nextOperator: 'AND' | 'OR' | 'NOT' | null = null;

    const addPositive = (term: TQueryTerm): void => {
      positiveTerms.push(term);
      if (nextOperator === 'OR') {
        optionalTerms.push(term);
      } else {
        requiredTerms.push(term);
      }
    };

    tokens.forEach((token) => {
      if (token.type === 'operator') {
        if (token.value === 'OR') {
          const previousRequired = requiredTerms.pop();
          if (previousRequired) {
            optionalTerms.push(previousRequired);
          }
        }
        nextOperator = token.value as 'AND' | 'OR' | 'NOT';
        return;
      }

      const term = { value: token.value, quoted: token.quoted };
      if (nextOperator === 'NOT') {
        excludedTerms.push(term);
      } else {
        addPositive(term);
      }
      nextOperator = null;
    });

    if (!isBoolean && positiveTerms.length === 0) {
      meaningfulTerms.forEach((value) => positiveTerms.push({ value, quoted: false }));
    }

    return {
      original,
      normalized,
      positiveTerms,
      requiredTerms,
      optionalTerms,
      excludedTerms,
      exactPhrases,
      meaningfulTerms,
      isBoolean,
      isSentence
    };
  }

  private static parseTokens(value: string): Array<{ type: 'term'; value: string; quoted: boolean } | { type: 'operator'; value: string }> {
    const tokens: Array<{ type: 'term'; value: string; quoted: boolean } | { type: 'operator'; value: string }> = [];
    const pattern = /"([^"]+)"|(\S+)/g;
    let match: RegExpExecArray | null;

    while ((match = pattern.exec(value)) !== null) {
      const quoted = String(match[1] || '').trim();
      const raw = String(match[2] || '').trim();
      if (quoted) {
        tokens.push({ type: 'term', value: quoted, quoted: true });
        continue;
      }

      const operator = raw.toUpperCase();
      if (operator === 'AND' || operator === 'OR' || operator === 'NOT') {
        tokens.push({ type: 'operator', value: operator });
      } else if (raw) {
        tokens.push({ type: 'term', value: raw, quoted: false });
      }
    }

    return tokens;
  }

  private static countTerms(haystack: string, terms: TQueryTerm[]): number {
    return terms.reduce((total, term) => total + this.countLiteral(haystack, term), 0);
  }

  private static matchesLiteral(haystack: string, term: TQueryTerm): boolean {
    return this.countLiteral(haystack, term) > 0;
  }

  private static countLiteral(haystack: string, term: TQueryTerm): number {
    const normalizedTerm = this.normalize(term.value);
    if (!haystack || !normalizedTerm) {
      return 0;
    }

    if (term.quoted || normalizedTerm.indexOf(' ') !== -1) {
      return this.countOccurrences(haystack, normalizedTerm);
    }

    const matches = haystack.match(new RegExp(`\\b${this.escapeRegExp(normalizedTerm)}\\b`, 'g'));
    return matches ? matches.length : 0;
  }

  private static countOccurrences(haystack: string, needle: string): number {
    let count = 0;
    let index = 0;
    while (index < haystack.length) {
      const matchIndex = haystack.indexOf(needle, index);
      if (matchIndex === -1) {
        break;
      }
      count += 1;
      index = matchIndex + needle.length;
    }
    return count;
  }

  private static tokenize(value: string): string[] {
    return Array.from(new Set(value.split(/[^a-z0-9]+/i).filter(Boolean)));
  }

  private static normalize(value: string): string {
    return String(value || '')
      .toLowerCase()
      .replace(/[_-]+/g, ' ')
      .replace(/[^\w\s.]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  private static escapeRegExp(value: string): string {
    return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
}
