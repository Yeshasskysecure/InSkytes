import { safeSearchKey, SearchEnv } from './env';
import {
  KnowledgeDocumentIndexRecord,
  PeopleIndexRecord
} from '../search/indexRecords';
import { canonicalDepartmentParts } from '../search/departmentTaxonomy';

type Fields = Record<string, unknown>;

export const parseListValue = (value: unknown): string => {
  if (!value) return '';
  if (typeof value === 'string') return value;
  if (Array.isArray(value)) return value.map(parseListValue).filter(Boolean).join('; ');
  if (typeof value === 'object') {
    const record = value as Record<string, unknown>;
    return String(record.Label || record.LookupValue || record.Title || record.Name || record.Email || record.value || '');
  }
  return String(value);
};

const parseAuthorList = (value: unknown): string[] => parseListValue(value)
  .split(/[;,]/)
  .map((item) => item.trim())
  .filter(Boolean);

const placeholderMetadataValues = new Set([
  'n.a.',
  'na',
  'not-applicable',
  'not_applicable',
  'none',
  '(none)',
  'null',
  '-'
]);

const normalizeMetadataValue = (value: unknown): string => parseListValue(value)
  .split(';')
  .map((item) => {
    const trimmed = item.trim();
    const normalized = trimmed.toLowerCase().replace(/[\s_-]+/g, ' ');
    return normalized === 'not applicable' || normalized === 'n/a' ? 'Not Applicable' : trimmed;
  })
  .filter((item) => {
    if (!item) return false;
    const normalized = item.toLowerCase().replace(/\s+/g, ' ');
    return !placeholderMetadataValues.has(normalized);
  })
  .join('; ');

const uniqueValues = (values: string[]): string[] =>
  values.filter((value, index, entries) => Boolean(value) && entries.indexOf(value) === index);

const cleanMetadataToken = (value: string): string => {
  const withoutTaxonomyId = value
    .replace(/^-?\d+#/g, '')
    .split('|')[0]
    .normalize('NFKC')
    .replace(/＆/g, '&')
    .replace(/\s*&\s*/g, ' & ')
    .replace(/\s+/g, ' ')
    .trim();
  const normalized = withoutTaxonomyId.toLowerCase().replace(/[\s_-]+/g, ' ');
  return normalized === 'not applicable' || normalized === 'n/a' ? 'Not Applicable' : withoutTaxonomyId;
};

const comparableMetadataToken = (value: string): string =>
  cleanMetadataToken(value)
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[\s_-]+/g, ' ')
    .trim();

const metadataFilterValues = (metadataValue: string | undefined): string[] =>
  uniqueValues(
    String(metadataValue || '')
      .replace(/;#/g, ';')
      .split(';')
      .map((item) => cleanMetadataToken(item.trim()))
      .filter((item) => {
        if (!item) return false;
        const normalized = item.toLowerCase().replace(/\s+/g, ' ');
        return !placeholderMetadataValues.has(normalized);
      })
  );

const splitDepartmentPath = (value: string): string[] =>
  String(value || '')
    .normalize('NFKC')
    .split(/\s*(?::|>|›|＞)\s*/g)
    .map((part) => cleanMetadataToken(part))
    .filter(Boolean);

const departmentHierarchyValues = (
  metadataValue: string | undefined,
  businessUnit: string | undefined,
  env: SearchEnv
): string[] => {
  const values = metadataFilterValues(metadataValue);
  const buValues = metadataFilterValues(businessUnit);
  const buComparableValues = new Set(buValues.map(comparableMetadataToken));
  const expanded: string[] = [];

  values.forEach((value) => {
    const parts = splitDepartmentPath(value);
    if (parts.length === 0) return;

    const firstPartComparable = comparableMetadataToken(parts[0]);
    const shouldStripBusinessUnitPrefix = parts.length > 1 &&
      buComparableValues.has(firstPartComparable);
    const rawDepartmentParts = shouldStripBusinessUnitPrefix
      ? parts.slice(1)
      : parts;
    const departmentParts = canonicalDepartmentParts(env, rawDepartmentParts, buValues);

    if (departmentParts.length === 0) return;

    for (let index = 1; index <= departmentParts.length; index += 1) {
      expanded.push(departmentParts.slice(0, index).join(':'));
    }
  });

  return uniqueValues(expanded);
};

const departmentFilterValues = (
  metadataValue: string | undefined,
  businessUnit: string | undefined,
  env: SearchEnv
): string[] => {
  const filterValues = departmentHierarchyValues(metadataValue, businessUnit, env);
  const legacyPathValues = metadataFilterValues(metadataValue)
    .map(splitDepartmentPath)
    .filter((parts) => parts.length > 1)
    .map((parts) => parts.join(':'));

  return uniqueValues([...filterValues, ...legacyPathValues]);
};

const departmentFacetValues = (
  metadataValue: string | undefined,
  businessUnit: string | undefined,
  env: SearchEnv
): string[] =>
  departmentHierarchyValues(metadataValue, businessUnit, env);

const firstFieldValue = (fields: Fields, names: string[]): unknown => {
  for (const name of names) {
    if (!name) {
      continue;
    }
    const value = fields[name];
    if (value !== undefined && value !== null && parseListValue(value).trim()) {
      return value;
    }
  }

  return undefined;
};

const configuredField = (
  env: SearchEnv,
  fields: Fields,
  envKey: string,
  fallbackNames: string[]
): unknown => firstFieldValue(fields, [env[envKey], ...fallbackNames].filter(Boolean) as string[]);

const whosWhoFieldAliases = {
  contacts: ['AllContacts', 'field_1'],
  emails: ['AllEmails', 'field_2'],
  allText: ['AllText', 'field_3'],
  contentIds: ['ContentIDs', 'field_7'],
  description: ['Description', 'field_11'],
  sectionTitles: ['SectionTitles', 'field_24'],
  serviceLine: ['ServiceLine', 'field_25']
};

const extensionFromName = (fileName: string): string => {
  const match = String(fileName || '').match(/\.([^.]+)$/);
  return match ? match[1].toUpperCase() : 'FILE';
};

const serverRelativeFromWebUrl = (webUrl: string): string => {
  try {
    return new URL(webUrl).pathname;
  } catch {
    return '';
  }
};

const uniqueIdFromDriveItem = (driveItem: any): string => {
  const sharePointIds = driveItem.sharepointIds || driveItem.parentReference?.sharepointIds || {};
  const explicitUniqueId = parseListValue(sharePointIds.listItemUniqueId);
  if (explicitUniqueId) return explicitUniqueId;

  const tagValue = parseListValue(driveItem.eTag || driveItem.cTag);
  const match = tagValue.match(/\{([0-9a-f-]{36})\}/i);
  return match ? match[1].toLowerCase() : '';
};

export const mapDocumentRecord = (
  env: SearchEnv,
  item: any,
  extractedText: string
): KnowledgeDocumentIndexRecord => {
  const now = new Date().toISOString();
  const fields: Fields = item.fields || {};
  const driveItem = item.driveItem || {};
  const fileName = parseListValue(configuredField(env, fields, 'SP_FIELD_FILE_NAME', ['FileLeafRef'])) || driveItem.name || `item-${item.id}`;
  const title = parseListValue(configuredField(env, fields, 'SP_FIELD_TITLE', ['Title'])) || fileName;
  const description = parseListValue(configuredField(env, fields, 'SP_FIELD_DESCRIPTION', ['Description']));
  const status = parseListValue(configuredField(env, fields, 'SP_FIELD_STATUS', ['Status'])) || 'Active';
  const documentId = driveItem.id || item.id || fileName;
  const fileUniqueId =
    parseListValue(fields.UniqueId || fields.FileUniqueId || fields.GUID) ||
    uniqueIdFromDriveItem(driveItem) ||
    '';
  const webUrl = driveItem.webUrl || parseListValue(configuredField(env, fields, 'SP_FIELD_URL', ['URL'])) || '';
  const fileRef = parseListValue(fields.FileRef) || serverRelativeFromWebUrl(webUrl);
  const createdDateTime = item.createdDateTime || parseListValue(configuredField(env, fields, 'SP_FIELD_CREATED', ['Created'])) || undefined;
  const lastModifiedDateTime =
    parseListValue(configuredField(env, fields, 'SP_FIELD_EDITED_DATE', ['Edited'])) ||
    item.lastModifiedDateTime ||
    parseListValue(configuredField(env, fields, 'SP_FIELD_MODIFIED', ['Modified'])) ||
    undefined;
  const contentPreview = (extractedText || description || title).slice(0, 500);
  const bu = normalizeMetadataValue(configuredField(env, fields, 'SP_FIELD_BU', ['BU']));
  const department = normalizeMetadataValue(configuredField(env, fields, 'SP_FIELD_DEPARTMENT', ['Department_x0020__x002f__x0020_Sub_x0020_Department', 'Department']));
  const diseaseArea = normalizeMetadataValue(configuredField(env, fields, 'SP_FIELD_DISEASE_AREA', ['Disease_x0020_Area']));
  const therapyArea = normalizeMetadataValue(configuredField(env, fields, 'SP_FIELD_THERAPY_AREA', ['Therapy_x0020_Area']));
  const client = normalizeMetadataValue(configuredField(env, fields, 'SP_FIELD_CLIENT', ['Client']));
  const region = normalizeMetadataValue(configuredField(env, fields, 'SP_FIELD_GEOGRAPHY', ['Geography', 'Region']));
  const documentType = normalizeMetadataValue(configuredField(env, fields, 'SP_FIELD_DOCUMENT_TYPE', ['Document_x0020_Type']));

  return {
    id: safeSearchKey(documentId),
    siteId: env.SHAREPOINT_SITE_ID,
    driveId: env.SHAREPOINT_LIBRARY_DRIVE_ID,
    listId: env.SHAREPOINT_LIBRARY_LIST_ID,
    listItemId: String(item.id || ''),
    driveItemId: driveItem.id || '',
    fileUniqueId,
    uniqueId: fileUniqueId,
    fileRef,
    serverRelativeUrl: fileRef,
    fileName,
    title,
    webUrl,
    fileExtension: extensionFromName(fileName),
    mimeType: (driveItem.file && driveItem.file.mimeType) || '',
    status,
    bu,
    buFilterValues: metadataFilterValues(bu),
    department,
    departmentFilterValues: departmentFilterValues(department, bu, env),
    departmentFacetValues: departmentFacetValues(department, bu, env),
    diseaseArea,
    diseaseAreaFilterValues: metadataFilterValues(diseaseArea),
    therapyArea,
    therapyAreaFilterValues: metadataFilterValues(therapyArea),
    client,
    clientFilterValues: metadataFilterValues(client),
    region,
    regionFilterValues: metadataFilterValues(region),
    documentType,
    documentTypeFilterValues: metadataFilterValues(documentType),
    authors: parseAuthorList(configuredField(env, fields, 'SP_FIELD_AUTHOR', ['Author0', 'Author'])),
    modifiedBy: parseListValue(configuredField(env, fields, 'SP_FIELD_EDITOR', ['Editor0', 'Editor'])),
    createdDateTime,
    created: createdDateTime,
    lastModifiedDateTime,
    modified: lastModifiedDateTime,
    publishedDate: parseListValue(configuredField(env, fields, 'SP_FIELD_PUBLISHED', ['Published'])) || undefined,
    contentRefreshDate: parseListValue(configuredField(env, fields, 'SP_FIELD_CONTENT_REFRESH_DATE', ['ContentRefreshDate'])) || undefined,
    description,
    contentPreview,
    projectId: parseListValue(configuredField(env, fields, 'SP_FIELD_PROJECT_ID', ['ProjectId', 'ProjectID', 'ProjectId0'])),
    sensitiveTerms: parseListValue(configuredField(env, fields, 'SP_FIELD_SENSITIVE_TERMS', ['SensitiveTerms'])),
    version: parseListValue(configuredField(env, fields, 'SP_FIELD_VERSION', ['Version', '_UIVersionString'])),
    indexedAt: now
  };
};

export const mapDocumentMetadataRecord = (env: SearchEnv, item: any): KnowledgeDocumentIndexRecord =>
  mapDocumentRecord(env, item, '');

export const mapPeopleRecords = (items: any[]): PeopleIndexRecord[] => {
  const now = new Date().toISOString();
  return items.map((item, index) => {
    const fields: Fields = item.fields || {};
    const name = parseListValue(fields.Title || fields.PersonName || fields.Name) || `Person ${index + 1}`;
    const contacts = parseListValue(firstFieldValue(fields, whosWhoFieldAliases.contacts));
    const emails = parseListValue(firstFieldValue(fields, whosWhoFieldAliases.emails));
    const allText = parseListValue(firstFieldValue(fields, whosWhoFieldAliases.allText));
    const contentIds = parseListValue(firstFieldValue(fields, whosWhoFieldAliases.contentIds));
    const description = parseListValue(firstFieldValue(fields, whosWhoFieldAliases.description));
    const sectionTitles = parseListValue(firstFieldValue(fields, whosWhoFieldAliases.sectionTitles));
    const serviceLine = parseListValue(firstFieldValue(fields, whosWhoFieldAliases.serviceLine));

    return {
      id: safeSearchKey(String(item.id || index + 1)),
      personName: name,
      email: parseListValue(fields.Email) || emails,
      contacts,
      allEmails: emails,
      allText,
      description,
      sectionTitles,
      serviceLine,
      contentIds,
      role: parseListValue(fields.Role || fields.JobTitle) || allText || description,
      team: parseListValue(fields.Team) || serviceLine || sectionTitles,
      bu: parseListValue(fields.BU || fields.BusinessUnit),
      region: parseListValue(fields.Region || fields.Geography),
      manager: parseListValue(fields.Manager),
      skills: parseListValue(fields.Skills || contentIds || sectionTitles).split(/[;,]/).map((x) => x.trim()).filter(Boolean),
      active: fields.Active === undefined ? true : fields.Active !== false,
      listItemUrl: item.webUrl || '',
      lastModifiedDateTime: item.lastModifiedDateTime || parseListValue(fields.Modified) || undefined,
      indexedAt: now
    };
  });
};
