import type { SearchEnv } from '../runtime/env';

export interface DepartmentTaxonomyPath {
  bu: string;
  parts: string[];
}

export const comparableTaxonomyValue = (value: string): string =>
  String(value || '')
    .normalize('NFKC')
    .replace(/＆/g, '&')
    .replace(/\s*&\s*/g, ' & ')
    .replace(/\s*\(\s*/g, '(')
    .replace(/\s*\)\s*/g, ')')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[\s_-]+/g, ' ')
    .trim();

const cleanTaxonomyPart = (value: string): string =>
  String(value || '')
    .normalize('NFKC')
    .replace(/＆/g, '&')
    .replace(/\s*&\s*/g, ' & ')
    .replace(/\s*\(\s*/g, '(')
    .replace(/\s*\)\s*/g, ')')
    .replace(/\s+/g, ' ')
    .trim();

const splitConfigPath = (value: string): string[] =>
  String(value || '')
    .split(/\s*(?:\||>|›|＞|:)\s*/g)
    .map(cleanTaxonomyPart)
    .filter(Boolean);

export const canonicalDepartmentParts = (
  _env: SearchEnv,
  parts: string[],
  _buValues: string[]
): string[] => {
  const cleanedParts = parts.map(cleanTaxonomyPart).filter(Boolean);
  if (cleanedParts.length === 0) return [];
  return cleanedParts;
};

export const legacyDepartmentAliasesForCanonicalPath = (
  _env: SearchEnv,
  normalizedPath: string
): string[] => {
  const parts = splitConfigPath(normalizedPath);
  if (parts.length === 0) return [];
  return [];
};
