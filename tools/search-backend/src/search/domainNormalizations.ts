import fs = require('fs');
import path = require('path');

declare const __dirname: string;

interface PhraseNormalizationRule {
  pattern: string;
  replacement: string;
}

interface AcronymCoreRule {
  pattern: string;
  core: string;
}

interface SearchDomainNormalizationsConfig {
  phraseNormalizations?: PhraseNormalizationRule[];
  retrievalNormalizations?: PhraseNormalizationRule[];
  acronymSearchCores?: AcronymCoreRule[];
}

let cachedConfig: SearchDomainNormalizationsConfig | undefined;

const configPath = (): string => path.resolve(__dirname, '..', '..', 'config', 'search-domain-normalizations.json');

const readConfig = (): SearchDomainNormalizationsConfig => {
  if (cachedConfig) return cachedConfig;
  cachedConfig = JSON.parse(fs.readFileSync(configPath(), 'utf8')) as SearchDomainNormalizationsConfig;
  return cachedConfig;
};

const toRegExp = (pattern: string, flags: string): RegExp | undefined => {
  try {
    return new RegExp(pattern, flags);
  } catch {
    return undefined;
  }
};

export const normalizeDomainSearchPhrase = (query: string): string =>
  (readConfig().phraseNormalizations || []).reduce((value, rule) => {
    const pattern = toRegExp(rule.pattern, 'gi');
    return pattern ? value.replace(pattern, rule.replacement) : value;
  }, String(query || ''));

export const normalizeDomainRetrievalPhrase = (query: string): string =>
  (readConfig().retrievalNormalizations || []).reduce((value, rule) => {
    const pattern = toRegExp(rule.pattern, 'gi');
    return pattern ? value.replace(pattern, rule.replacement) : value;
  }, String(query || ''));

export const domainAcronymCoreQuery = (query: string): string => {
  const cores = (readConfig().acronymSearchCores || [])
    .filter((entry) => {
      const pattern = toRegExp(entry.pattern, 'i');
      return Boolean(pattern && pattern.test(query));
    })
    .map((entry) => entry.core);

  return Array.from(new Set(cores)).join(' ');
};
