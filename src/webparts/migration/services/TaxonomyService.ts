import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { findBestMatch } from './ValidationConstants';
import { COLUMN_NAMES, TERM_SET_IDS } from '../config/appConfig';

export interface ITaxonomyTerm {
  id: string;
  label: string;
  path: string;
  level?: number;
  parentId?: string;
}

export interface ITaxonomyTreeNode {
  key: string;
  label: string;
  path: string;
  term: ITaxonomyTerm | null;
  children: ITaxonomyTreeNode[];
}

export interface ITaxonomyFieldConfig {
  fieldInternalName: string;
  termSetId: string;
}

export interface ITaxonomyFieldOptions {
  buDepartment: ITaxonomyTerm[];
  documentType: ITaxonomyTerm[];
  client: ITaxonomyTerm[];
  geography: ITaxonomyTerm[];
  diseaseArea: ITaxonomyTerm[];
  therapyArea: ITaxonomyTerm[];
}

export const TAXONOMY_FIELD_CONFIGS: Record<keyof ITaxonomyFieldOptions, ITaxonomyFieldConfig> = {
  buDepartment: {
    fieldInternalName: '',
    termSetId: TERM_SET_IDS.buDepartment
  },
  client: {
    fieldInternalName: COLUMN_NAMES.client,
    termSetId: TERM_SET_IDS.client
  },
  diseaseArea: {
    fieldInternalName: COLUMN_NAMES.diseaseArea,
    termSetId: TERM_SET_IDS.diseaseArea
  },
  documentType: {
    fieldInternalName: COLUMN_NAMES.documentType,
    termSetId: TERM_SET_IDS.documentType
  },
  geography: {
    fieldInternalName: COLUMN_NAMES.geography,
    termSetId: TERM_SET_IDS.geography
  },
  therapyArea: {
    fieldInternalName: COLUMN_NAMES.therapyArea,
    termSetId: TERM_SET_IDS.therapyArea
  }
};

const normalizeTermValue = (value?: string): string =>
  (value || '').trim().toLowerCase();

const tokenizeValue = (value?: string): string[] =>
  normalizeTermValue(value)
    .replace(/[^a-z0-9]+/g, ' ')
    .split(/\s+/)
    .filter((token) => token.length > 1);

const scoreCandidateMatch = (value: string, candidate: string): number => {
  const normalizedValue = normalizeTermValue(value);
  const normalizedCandidate = normalizeTermValue(candidate);

  if (!normalizedValue || !normalizedCandidate) {
    return 0;
  }

  if (normalizedValue === normalizedCandidate) {
    return 1000;
  }

  let score = 0;

  if (normalizedCandidate.indexOf(normalizedValue) !== -1 || normalizedValue.indexOf(normalizedCandidate) !== -1) {
    score += 400;
  }

  const valueTokens = tokenizeValue(normalizedValue);
  const candidateTokens = tokenizeValue(normalizedCandidate);

  for (let index = 0; index < valueTokens.length; index++) {
    const valueToken = valueTokens[index];

    for (let candidateIndex = 0; candidateIndex < candidateTokens.length; candidateIndex++) {
      const candidateToken = candidateTokens[candidateIndex];

      if (valueToken === candidateToken) {
        score += 120;
      } else if (valueToken.indexOf(candidateToken) !== -1 || candidateToken.indexOf(valueToken) !== -1) {
        score += 60;
      }
    }
  }

  return score;
};

const getPrimaryLabel = (labels: Array<{ name?: string; languageTag?: string }>): string => {
  if (!labels || labels.length === 0) {
    return '';
  }

  for (let index = 0; index < labels.length; index++) {
    if (labels[index].name) {
      return labels[index].name || '';
    }
  }

  return '';
};

const normalizePath = (path?: string, label?: string): string =>
  String(path || label || '')
    .replace(/;/g, ' > ')
    .replace(/\|/g, ' > ')
    .replace(/\\/g, ' > ')
    .replace(/\s*>\s*/g, ' > ')
    .trim();

const getPathSegments = (path?: string, label?: string): string[] =>
  normalizePath(path, label)
    .split(' > ')
    .map((segment) => segment.trim())
    .filter((segment) => segment.length > 0);

export const getTaxonomyPathSegments = (term?: ITaxonomyTerm | null): string[] =>
  getPathSegments(term?.path, term?.label);

const parseTerms = (items: any[]): ITaxonomyTerm[] => {
  const seen: Record<string, boolean> = {};
  const terms: ITaxonomyTerm[] = [];
  const termById: Record<string, ITaxonomyTerm> = {};
  const termByPath: Record<string, ITaxonomyTerm> = {};

  for (let index = 0; index < items.length; index++) {
    const item = items[index];
    const id = item?.id || item?.Id;
    const label = getPrimaryLabel(item?.labels || item?.Labels || []);
    const path = item?.path || item?.PathOfTerm || label;
    const parentId = item?.parent?.id || item?.Parent?.Id || item?.parentId || item?.ParentId;

    if (!id || !label || seen[id]) {
      continue;
    }

    seen[id] = true;
    const term: ITaxonomyTerm = {
      id: String(id),
      label,
      path: normalizePath(path, label),
      parentId: parentId ? String(parentId) : undefined,
      level: 0
    };
    terms.push(term);
    termById[term.id] = term;
    termByPath[term.path] = term;
  }

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index];
    const segments = getPathSegments(term.path, term.label);

    if (!term.parentId && segments.length > 1) {
      const parentPath = segments.slice(0, segments.length - 1).join(' > ');
      if (termByPath[parentPath]) {
        term.parentId = termByPath[parentPath].id;
      }
    }
  }

  const computeLevel = (term: ITaxonomyTerm): number => {
    const segments = getPathSegments(term.path, term.label);

    if (term.parentId && termById[term.parentId]) {
      const parent = termById[term.parentId];
      if (typeof parent.level === 'number' && parent.level > 0) {
        return parent.level + 1;
      }

      parent.level = computeLevel(parent);
      return (parent.level || 0) + 1;
    }

    return segments.length > 0 ? segments.length - 1 : 0;
  };

  for (let index = 0; index < terms.length; index++) {
    terms[index].level = computeLevel(terms[index]);
  }

  terms.sort((left, right) => {
    if ((left.level || 0) !== (right.level || 0)) {
      return (left.level || 0) - (right.level || 0);
    }

    return left.path.localeCompare(right.path);
  });

  return terms;
};

const fetchTermsForSet = async (context: WebPartContext, termSetId: string): Promise<ITaxonomyTerm[]> => {
  const webUrl = context.pageContext.web.absoluteUrl;
  let requestUrl = `${webUrl}/_api/v2.1/termstore/sets/${termSetId}/terms`;
  const allItems: any[] = [];

  while (requestUrl) {
    const response = await context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata.metadata=minimal'
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to fetch terms for set ${termSetId}: ${response.status} ${errorText}`);
    }

    const json = await response.json();
    allItems.push(...(json.value || []));
    requestUrl = json['@odata.nextLink'] || '';
  }

  const parsedTerms = parseTerms(allItems);

  if (termSetId === TAXONOMY_FIELD_CONFIGS.buDepartment.termSetId) {
    console.log('BU term store debug:', parsedTerms.map((term) => ({
      id: term.id,
      label: term.label,
      path: term.path,
      parentId: term.parentId || null,
      level: term.level || 0
    })));
  }

  return parsedTerms;
};

export const fetchAllTaxonomyOptions = async (context: WebPartContext): Promise<ITaxonomyFieldOptions> => {
  const [
    buDepartment,
    documentType,
    client,
    geography,
    diseaseArea,
    therapyArea
  ] = await Promise.all([
    fetchTermsForSet(context, TAXONOMY_FIELD_CONFIGS.buDepartment.termSetId),
    fetchTermsForSet(context, TAXONOMY_FIELD_CONFIGS.documentType.termSetId),
    fetchTermsForSet(context, TAXONOMY_FIELD_CONFIGS.client.termSetId),
    fetchTermsForSet(context, TAXONOMY_FIELD_CONFIGS.geography.termSetId),
    fetchTermsForSet(context, TAXONOMY_FIELD_CONFIGS.diseaseArea.termSetId),
    fetchTermsForSet(context, TAXONOMY_FIELD_CONFIGS.therapyArea.termSetId)
  ]);

  return {
    buDepartment,
    documentType,
    client,
    geography,
    diseaseArea,
    therapyArea
  };
};

export const buildTaxonomyTree = (terms: ITaxonomyTerm[]): ITaxonomyTreeNode[] => {
  const nodeByKey: Record<string, ITaxonomyTreeNode> = {};
  const roots: ITaxonomyTreeNode[] = [];

  const ensureNode = (segments: string[]): ITaxonomyTreeNode => {
    const key = segments.join(' > ');
    if (!nodeByKey[key]) {
      nodeByKey[key] = {
        key,
        label: segments[segments.length - 1] || '',
        path: key,
        term: null,
        children: []
      };
    }

    return nodeByKey[key];
  };

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index];
    const segments = getTaxonomyPathSegments(term);

    if (segments.length === 0) {
      continue;
    }

    for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex++) {
      const currentSegments = segments.slice(0, segmentIndex + 1);
      const currentNode = ensureNode(currentSegments);

      if (segmentIndex === segments.length - 1) {
        currentNode.term = term;
      }

      if (segmentIndex === 0) {
        if (!roots.some((rootNode) => rootNode.key === currentNode.key)) {
          roots.push(currentNode);
        }
        continue;
      }

      const parentNode = ensureNode(segments.slice(0, segmentIndex));
      if (!parentNode.children.some((childNode) => childNode.key === currentNode.key)) {
        parentNode.children.push(currentNode);
      }
    }
  }

  const sortNodes = (nodes: ITaxonomyTreeNode[]): void => {
    nodes.sort((left, right) => left.label.localeCompare(right.label));
    for (let index = 0; index < nodes.length; index++) {
      sortNodes(nodes[index].children);
    }
  };

  sortNodes(roots);
  return roots;
};

export const matchTaxonomyTerm = (value: string, terms: ITaxonomyTerm[]): ITaxonomyTerm | null => {
  const normalizedValue = normalizeTermValue(value);

  if (!normalizedValue) {
    return null;
  }

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index];
    if (normalizeTermValue(term.label) === normalizedValue || normalizeTermValue(term.path) === normalizedValue) {
      return term;
    }
  }

  let bestScoredTerm: ITaxonomyTerm | null = null;
  let bestScore = 0;

  for (let index = 0; index < terms.length; index++) {
    const term = terms[index];
    const labelScore = scoreCandidateMatch(value, term.label);
    const pathScore = scoreCandidateMatch(value, term.path);
    const score = Math.max(labelScore, pathScore);

    if (score > bestScore) {
      bestScore = score;
      bestScoredTerm = term;
    }
  }

  if (bestScoredTerm && bestScore >= 120) {
    return bestScoredTerm;
  }

  const bestPathMatch = findBestMatch(value, terms.map((term) => term.path));
  if (bestPathMatch) {
    for (let index = 0; index < terms.length; index++) {
      if (terms[index].path === bestPathMatch) {
        return terms[index];
      }
    }
  }

  const bestLabelMatch = findBestMatch(value, terms.map((term) => term.label));
  if (bestLabelMatch) {
    for (let index = 0; index < terms.length; index++) {
      if (terms[index].label === bestLabelMatch) {
        return terms[index];
      }
    }
  }

  return null;
};
