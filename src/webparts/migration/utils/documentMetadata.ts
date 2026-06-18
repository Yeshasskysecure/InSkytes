import { SPHttpClient } from '@microsoft/sp-http';
import { COLUMN_NAMES } from '../config/appConfig';

interface IKMDataHubReadFieldMap {
  author: string;
  isPersonField?: boolean;
  title: string;
  description: string;
  published: string;
  status: string;
  url: string;
  contentRefreshDate?: string;
  views?: string;
  likes?: string;
  comments?: string;
  downloads?: string;
  follow?: string;
  share?: string;
  bookmark?: string;
  sensitiveTerms?: string;
  reviewerComments?: string;
  projectId?: string;
  versionFileName?: string;
  versionFileType?: string;
  edited?: string;
  editedBy?: string;
  modifiedBy?: string;
  docIcon?: string;
  fileLeafRef?: string;
  fileRef?: string;
  businessUnit?: string;
  department?: string;
}

const kmDataHubReadFieldMapCache: Record<string, IKMDataHubReadFieldMap> = {};

const extractAuthorValue = (value: any): string => {
  if (typeof value === 'string' && value.trim()) {
    const normalizedStringValue = value
      .split(/;#|;/)
      .map((entry) => entry.trim())
      .filter(Boolean)
      .filter((entry, index, array) => array.indexOf(entry) === index)
      .join(', ');

    return normalizedStringValue || value.trim();
  }

  if (Array.isArray(value) && value.length > 0) {
    const joinedAuthors = value
      .map((entry: unknown) => {
        if (typeof entry === 'string') {
          return entry.trim();
        }

        if (entry && typeof entry === 'object') {
          const personEntry = entry as {
            Title?: string;
            Name?: string;
            Email?: string;
            EMail?: string;
            UserPrincipalName?: string;
            LoginName?: string;
          };
          const resolvedEntry =
            personEntry.Title ||
            personEntry.Name ||
            personEntry.Email ||
            personEntry.EMail ||
            personEntry.UserPrincipalName ||
            personEntry.LoginName ||
            '';

          return typeof resolvedEntry === 'string' ? resolvedEntry.trim() : '';
        }

        return '';
      })
      .filter(Boolean)
      .filter((entry, index, array) => array.indexOf(entry) === index)
      .join(', ');

    if (joinedAuthors) {
      return joinedAuthors;
    }
  }

  if (value?.Title) {
    return value.Title;
  }

  if (value?.Name) {
    return value.Name;
  }

  return '';
};

const extractPersonDisplayValues = (value: any): string[] => {
  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.reduce<string[]>((entries, entry) => {
      entries.push(...extractPersonDisplayValues(entry));
      return entries;
    }, []);
  }

  if (typeof value === 'string') {
    return value
      .split(/[;,]/)
      .map((entry) => entry.trim())
      .filter(Boolean);
  }

  if (typeof value === 'object') {
    return [
      value.Title,
      value.Name,
      value.Email,
      value.EMail
    ]
      .filter((entry): entry is string => typeof entry === 'string' && entry.trim().length > 0)
      .map((entry) => entry.trim());
  }

  return [];
};

export const fetchKMDataHubReadFieldMap = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  libraryName: string
): Promise<IKMDataHubReadFieldMap> => {
  const cacheKey = `${webUrl}::${libraryName}`;
  if (kmDataHubReadFieldMapCache[cacheKey]) {
    return kmDataHubReadFieldMapCache[cacheKey];
  }

  const defaultFieldMap: IKMDataHubReadFieldMap = {
    author: COLUMN_NAMES.author,
    isPersonField: true,
    title: COLUMN_NAMES.title,
    description: COLUMN_NAMES.description,
    published: COLUMN_NAMES.published,
    status: COLUMN_NAMES.status,
    url: COLUMN_NAMES.url,
    contentRefreshDate: COLUMN_NAMES.contentRefreshDate,
    views: COLUMN_NAMES.views,
    likes: COLUMN_NAMES.likes,
    comments: COLUMN_NAMES.comments,
    downloads: COLUMN_NAMES.downloads,
    follow: COLUMN_NAMES.follow,
    share: COLUMN_NAMES.share,
    bookmark: COLUMN_NAMES.bookmark,
    sensitiveTerms: COLUMN_NAMES.sensitiveTerms,
    reviewerComments: COLUMN_NAMES.reviewerComments,
    projectId: COLUMN_NAMES.projectId,
    versionFileName: COLUMN_NAMES.versionFileName,
    versionFileType: COLUMN_NAMES.versionFileType,
    edited: COLUMN_NAMES.edited,
    editedBy: COLUMN_NAMES.editedBy,
    modifiedBy: COLUMN_NAMES.modifiedBy,
    docIcon: COLUMN_NAMES.docIcon,
    fileLeafRef: COLUMN_NAMES.fileLeafRef,
    fileRef: COLUMN_NAMES.fileRef
  };

  try {
    const response = await spHttpClient.get(
      `${webUrl}/_api/web/lists/getbytitle('${libraryName}')/fields?$select=Title,InternalName,FieldTypeKind`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata.metadata=minimal'
        }
      }
    );

    if (!response.ok) {
      kmDataHubReadFieldMapCache[cacheKey] = defaultFieldMap;
      return defaultFieldMap;
    }

    const json = await response.json();
    const fields = (json?.d?.results || json?.value || []) as Array<{
      Title?: string;
      InternalName?: string;
      FieldTypeKind?: number;
    }>;

    const normalize = (value?: string): string => (value || '').trim().toLowerCase();
    const findInternalName = (displayName: string, fallback: string): string => {
      const matchedField = fields.find((field) =>
        normalize(field.Title) === normalize(displayName) ||
        normalize(field.InternalName) === normalize(fallback)
      );

      return matchedField?.InternalName || fallback;
    };

    // Prefer the dedicated Author column first. Only fall back to other custom fields if Author0 is unavailable.
    const customAuthorField =
      fields.find((field) => field.InternalName === COLUMN_NAMES.author) ||
      fields.find((field) => field.InternalName !== 'Author' && normalize(field.Title) === 'author') ||
      fields.find((field) => field.InternalName !== 'Author' && normalize(field.Title) === 'contributor');
    const authorInternalName = customAuthorField?.InternalName || findInternalName('Author', COLUMN_NAMES.author);
    
    // 🚩 UAT FIX: Ensure we detect the correct internal name AND type for expansion
    const authorField = fields.find(f => f.InternalName === authorInternalName);
    const isPersonField = authorField?.FieldTypeKind === 20 || authorField?.FieldTypeKind === 11; // 20: User, 11: Multi-User

    const resolvedFieldMap: IKMDataHubReadFieldMap = {
      author: authorInternalName,
      isPersonField: isPersonField,
      title: findInternalName(COLUMN_NAMES.title, COLUMN_NAMES.title),
      description: findInternalName(COLUMN_NAMES.description, COLUMN_NAMES.description),
      published: findInternalName(COLUMN_NAMES.published, COLUMN_NAMES.published),
      status: findInternalName(COLUMN_NAMES.status, COLUMN_NAMES.status),
      url: findInternalName('URL', COLUMN_NAMES.url),
      contentRefreshDate: findInternalName('Content Refresh Date', COLUMN_NAMES.contentRefreshDate),
      views: findInternalName(COLUMN_NAMES.views, COLUMN_NAMES.views),
      likes: findInternalName(COLUMN_NAMES.likes, COLUMN_NAMES.likes),
      comments: findInternalName(COLUMN_NAMES.comments, COLUMN_NAMES.comments),
      downloads: findInternalName(COLUMN_NAMES.downloads, COLUMN_NAMES.downloads),
      follow: findInternalName(COLUMN_NAMES.follow, COLUMN_NAMES.follow),
      share: findInternalName(COLUMN_NAMES.share, COLUMN_NAMES.share),
      bookmark: findInternalName(COLUMN_NAMES.bookmark, COLUMN_NAMES.bookmark),
      sensitiveTerms: findInternalName('Sensitive Terms', COLUMN_NAMES.sensitiveTerms),
      reviewerComments: findInternalName('Reviewer Comments', COLUMN_NAMES.reviewerComments),
      projectId: findInternalName('Project ID', COLUMN_NAMES.projectId),
      versionFileName: findInternalName('Version File Name', COLUMN_NAMES.versionFileName),
      versionFileType: findInternalName('Version File Type', COLUMN_NAMES.versionFileType),
      edited: findInternalName('Edited', COLUMN_NAMES.edited),
      editedBy: findInternalName('Editor', COLUMN_NAMES.editedBy),
      modifiedBy: findInternalName('Modified By', COLUMN_NAMES.modifiedBy),
      docIcon: findInternalName('Doc Icon', COLUMN_NAMES.docIcon),
      fileLeafRef: findInternalName(COLUMN_NAMES.fileLeafRef, COLUMN_NAMES.fileLeafRef),
      fileRef: findInternalName(COLUMN_NAMES.fileRef, COLUMN_NAMES.fileRef),
      businessUnit: findInternalName('BU', COLUMN_NAMES.businessUnit),
      department: findInternalName('Department / Sub Department', COLUMN_NAMES.department)
    };

    kmDataHubReadFieldMapCache[cacheKey] = resolvedFieldMap;
    return resolvedFieldMap;
  } catch (err) {
    console.warn("Error fetching field map, using defaults:", err);
    kmDataHubReadFieldMapCache[cacheKey] = defaultFieldMap;
    return defaultFieldMap;
  }
};

export const buildKMDataHubItemQuery = (
  fieldMap: IKMDataHubReadFieldMap,
  baseSelectFields: string[],
  baseExpandFields: string[] = []
): { select: string; expand: string } => {
  const selectFields = [...baseSelectFields];
  const expandFields = [...baseExpandFields];

  if (fieldMap.author) {
    selectFields.push(`${fieldMap.author}/Title`);
    selectFields.push(`${fieldMap.author}Id`);
    if (!expandFields.includes(fieldMap.author)) {
      expandFields.push(fieldMap.author);
    }
  }

  // 🚩 UAT FIX: Always include system 'Author' as a robust fallback.
  selectFields.push("Author/Title");
  if (!expandFields.includes("Author")) {
    expandFields.push("Author");
  }

  if (fieldMap.published) {
    selectFields.push(fieldMap.published);
  }

  return {
    select: Array.from(new Set(selectFields)).join(','),
    expand: Array.from(new Set(expandFields)).join(',')
  };
};

export const resolveDocumentAuthor = (
  item: any,
  fallbackAuthor: string = 'Internal',
  dynamicAuthorFieldInternalName?: string
): string => {
  const authorCandidates = [
    dynamicAuthorFieldInternalName ? item?.[dynamicAuthorFieldInternalName] : undefined,
    item?.Author0,
    item?.Contributor,
    item?.DocAuthor,
    item?.Author
  ];

  for (const candidate of authorCandidates) {
    const resolvedAuthor = extractAuthorValue(candidate);
    if (resolvedAuthor) {
      return resolvedAuthor;
    }
  }

  return fallbackAuthor;
};

export const splitDocumentAuthorDisplay = (author: string | undefined | null): string[] => {
  const normalizedValue = String(author || '').trim();
  if (!normalizedValue) {
    return [];
  }

  return normalizedValue
    .split(',')
    .map((entry) => entry.trim())
    .filter(Boolean)
    .filter((entry, index, entries) => entries.indexOf(entry) === index);
};

export const resolveDocumentPublishedValue = (
  item: any,
  dynamicPublishedFieldInternalName?: string
): string | Date | undefined =>
  (dynamicPublishedFieldInternalName ? item?.[dynamicPublishedFieldInternalName] : undefined) ||
  item?.Published ||
  item?.OData__Published ||
  item?.FirstPublishedDate ||
  item?.TimeStamp ||
  item?.Created;

export const formatDocumentPublishedDate = (
  value: string | Date | undefined,
  locale: string,
  options: Intl.DateTimeFormatOptions
): string => {
  if (!value) {
    return '';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }

  return date.toLocaleDateString(locale, options);
};
