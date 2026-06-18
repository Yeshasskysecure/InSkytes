import {
  canonicalDepartmentParts,
  legacyDepartmentAliasesForCanonicalPath
} from './departmentTaxonomy';

export interface DocumentSearchFilters {
  status?: string;
  bu?: string[];
  department?: string[];
  diseaseArea?: string[];
  therapyArea?: string[];
  client?: string[];
  region?: string[];
  documentType?: string[];
  fileExtension?: string[];
  authors?: string[];
}

const escapeODataString = (value: string): string => value.replace(/'/g, "''");

const buildEqualsAny = (fieldName: string, values: string[] | undefined): string | undefined => {
  const cleanedValues = (values || []).map((value) => value.trim()).filter(Boolean);
  if (cleanedValues.length === 0) {
    return undefined;
  }

  const clauses = cleanedValues.map((value) => `${fieldName} eq '${escapeODataString(value)}'`);
  return clauses.length === 1 ? clauses[0] : `(${clauses.join(' or ')})`;
};

const escapeSearchPhrase = (value: string): string =>
  value.replace(/\\/g, '\\\\').replace(/"/g, '\\"');

const buildSearchableMetadataAny = (fieldName: string, values: string[] | undefined): string | undefined => {
  const cleanedValues = (values || []).map((value) => value.trim()).filter(Boolean);
  if (cleanedValues.length === 0) {
    return undefined;
  }

  const clauses = cleanedValues.map((value) => {
    const escaped = escapeODataString(value);
    const phrase = escapeODataString(`"${escapeSearchPhrase(value)}"`);
    return `(${fieldName} eq '${escaped}' or search.ismatch('${phrase}', '${fieldName}', 'full', 'all'))`;
  });

  return clauses.length === 1 ? clauses[0] : `(${clauses.join(' or ')})`;
};

const buildCollectionAny = (
  fieldName: string,
  variableName: string,
  values: string[] | undefined
): string | undefined => {
  const cleanedValues = (values || []).map((value) => value.trim()).filter(Boolean);
  if (cleanedValues.length === 0) {
    return undefined;
  }

  const clauses = cleanedValues.map((value) => `${variableName} eq '${escapeODataString(value)}'`);
  return `${fieldName}/any(${variableName}: ${clauses.join(' or ')})`;
};

const normalizedFacetFiltersEnabled = (): boolean =>
  /^(1|true|yes)$/i.test(String((process as any).env?.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED || ''));

const canonicalDepartmentFacetsEnabled = (): boolean =>
  /^(1|true|yes)$/i.test(String((process as any).env?.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED || ''));

const normalizeFilterPathValue = (value: string): string =>
  value
    .normalize('NFKC')
    .replace(/＆/g, '&')
    .replace(/\s*&\s*/g, ' & ')
    .replace(/\s+/g, ' ')
    .split(/\s*(?::|>|›|＞)\s*/g)
    .map((part) => part.trim())
    .filter(Boolean)
    .join(':');

const normalizeFilterValues = (values: string[] | undefined): string[] | undefined => {
  const normalized = (values || [])
    .map((value) => normalizeFilterPathValue(String(value || '').trim()))
    .filter(Boolean);

  return normalized.length > 0 ? Array.from(new Set(normalized)) : undefined;
};

const expandDepartmentPathValue = (normalized: string): string[] => {
  return normalized ? [normalized, ...legacyDepartmentAliasesForCanonicalPath((process as any).env || {}, normalized)] : [];
};

const expandDepartmentFilterValues = (
  values: string[] | undefined,
  businessUnits: string[] | undefined
): string[] | undefined => {
  const expanded: string[] = [];
  const env = (process as any).env || {};
  const buValues = normalizeFilterValues(businessUnits) || [];

  (values || []).forEach((value) => {
    const normalized = normalizeFilterPathValue(String(value || '').trim());
    if (!normalized) return;

    // Canonical department facets strip BU prefixes for display, while older
    // UI fallbacks can still send full hierarchy paths. Keep the original
    // value and its suffix so stale clients do not turn valid child filters
    // into zero-result filters.
    expanded.push(...expandDepartmentPathValue(normalized));

    const canonicalParts = canonicalDepartmentParts(env, normalized.split(':').filter(Boolean), buValues);
    if (canonicalParts.length > 0) {
      expanded.push(...expandDepartmentPathValue(canonicalParts.join(':')));
    }
  });

  const unique = Array.from(new Set(expanded.filter(Boolean)));
  return unique.length > 0 ? unique : undefined;
};

const buildMetadataFilter = (
  legacyFieldName: string,
  normalizedFieldName: string,
  variableName: string,
  values: string[] | undefined,
  normalizeValues: (values: string[] | undefined) => string[] | undefined = normalizeFilterValues
): string | undefined => normalizedFacetFiltersEnabled()
  ? buildCollectionAny(normalizedFieldName, variableName, normalizeValues(values))
  : buildSearchableMetadataAny(legacyFieldName, values);

export const buildDocumentFilter = (filters: DocumentSearchFilters = {}): string => {
  const departmentFilterValues = expandDepartmentFilterValues(filters.department, filters.bu);
  const clauses = [
    `status eq '${escapeODataString(filters.status || 'Active')}'`,
    buildMetadataFilter('bu', 'buFilterValues', 'buValue', filters.bu),
    buildMetadataFilter('department', 'departmentFilterValues', 'departmentValue', departmentFilterValues, (values) => values),
    buildMetadataFilter('diseaseArea', 'diseaseAreaFilterValues', 'diseaseAreaValue', filters.diseaseArea),
    buildMetadataFilter('therapyArea', 'therapyAreaFilterValues', 'therapyAreaValue', filters.therapyArea),
    buildMetadataFilter('client', 'clientFilterValues', 'clientValue', filters.client),
    buildMetadataFilter('region', 'regionFilterValues', 'regionValue', filters.region),
    buildMetadataFilter('documentType', 'documentTypeFilterValues', 'documentTypeValue', filters.documentType),
    buildEqualsAny('fileExtension', filters.fileExtension),
    buildCollectionAny('authors', 'author', filters.authors)
  ].filter((clause): clause is string => Boolean(clause));

  return clauses.join(' and ');
};
