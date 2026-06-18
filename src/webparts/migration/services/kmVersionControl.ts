import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { COLUMN_NAMES, LIBRARY_NAMES } from '../config/appConfig';
import { BUSINESS_UNIT_HIERARCHY } from '../components/businessUnit/RenameBusinessUnitDialog/BusinessUnitHierarchy';

const LIBRARY_NAME = LIBRARY_NAMES.kmDataHub;

const escapeODataString = (value: string): string => (value || '').replace(/'/g, "''");
const formatSharePointDateFieldValue = (date: Date = new Date()): string => {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'Asia/Kolkata',
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    hour12: true
  }).formatToParts(date);
  const getPart = (type: string): string =>
    parts.find((part) => part.type === type)?.value || '';

  return `${getPart('month')}/${getPart('day')}/${getPart('year')} ${getPart('hour')}:${getPart('minute')} ${getPart('dayPeriod')}`;
};
const encodeFileRef = (fileRef: string): string => {
  try {
    const decoded = decodeURIComponent(fileRef);
    return encodeURIComponent(decoded);
  } catch {
    return encodeURIComponent(fileRef);
  }
};

// Helpers

export const getFileRef = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number
): Promise<string> => {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
    `/items(${itemId})?$select=FileRef`;
  const res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1
  );
  const data = await res.json();
  const fileRef = data?.FileRef || data?.d?.FileRef;
  if (!fileRef) throw new Error(`FileRef not found for item ${itemId}`);
  return fileRef as string;
};

// Checkout

export const checkoutFile = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  fileRef: string
): Promise<void> => {
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativeUrl` +
    `('${encodeFileRef(fileRef)}')/checkout()`;
  console.log('checkoutFile: calling', url);

  for (let attempt = 1; attempt <= 5; attempt++) {
    const response = await spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );
    console.log('checkoutFile: response status', response.status, 'attempt', attempt);

    if (response.ok) {
      console.log('checkoutFile: success for', fileRef);
      return;
    }

    const text = await response.text();
    if (
      text.includes('already checked out') ||
      text.includes('checked out by')
    ) {
      console.warn('checkoutFile: already checked out, continuing.');
      return;
    }

    if (response.status === 423 && attempt < 5) {
      console.warn('checkoutFile: file locked, retrying:', text);
      await new Promise(r => setTimeout(r, 1000));
      continue;
    }

    console.error('checkoutFile failed:', response.status, text);
    throw new Error(`checkoutFile failed: ${response.status} ${text}`);
  }
};

// Checkin - overwrite (no new version)

export const checkinOverwrite = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  fileRef: string,
  comment: string = 'Metadata update'
): Promise<void> => {
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativeUrl` +
    `('${encodeFileRef(fileRef)}')/checkin(comment='${escapeODataString(comment)}',checkintype=2)`;
  console.log('checkinOverwrite: calling', url);
  const response = await spHttpClient.post(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: { Accept: 'application/json;odata=nometadata' }
    }
  );
  console.log('checkinOverwrite: response status', response.status);
  if (!response.ok) {
    const text = await response.text();
    if (
      response.status === 423 ||
      text.includes('not checked out') ||
      text.includes('SPFileCheckOutException')
    ) {
      console.log(
        'checkinOverwrite: 423 - file already checked in by SP.'
      );
      return;
    }
    console.error('checkinOverwrite failed:', response.status, text);
    throw new Error(
      `checkinOverwrite failed: ${response.status} ${text}`
    );
  }
  console.log('checkinOverwrite: success');
};

// Checkin - major (new version)

export const checkinMajor = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  fileRef: string,
  comment: string = 'New file version'
): Promise<void> => {
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativeUrl` +
    `('${encodeFileRef(fileRef)}')/checkin(comment='${escapeODataString(comment)}',checkintype=1)`;
  console.log('checkinMajor: calling', url);
  const response = await spHttpClient.post(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: { Accept: 'application/json;odata=nometadata' }
    }
  );
  console.log('checkinMajor: response status', response.status);
  if (!response.ok) {
    const text = await response.text();
    console.error('checkinMajor failed:', response.status, text);
    throw new Error(
      `checkinMajor failed: ${response.status} ${text}`
    );
  }
  console.log('checkinMajor: success');
};

// Undo checkout (on error)

export const undoCheckout = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  fileRef: string
): Promise<void> => {
  try {
    const url =
      `${webUrl}/_api/web/GetFileByServerRelativeUrl` +
      `('${encodeFileRef(fileRef)}')/undocheckout()`;
    const response = await spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (response.status === 423) {
      console.warn(
        'undoCheckout: file locked by SP auto-checkout (423),' +
        ' cannot undo. File will release automatically.'
      );
      return;
    }
  } catch (e) {
    console.warn('undoCheckout failed silently:', e);
  }
};

// Update without version (metadata, status, counters)

export const updateWithoutVersion = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number,
  updateFn: () => Promise<void>
): Promise<void> => {
  const fileRef = await getFileRef(spHttpClient, webUrl, itemId);
  await checkoutFile(spHttpClient, webUrl, fileRef);
  try {
    await updateFn();
    await checkinOverwrite(spHttpClient, webUrl, fileRef, 'Metadata update');
  } catch (err) {
    await undoCheckout(spHttpClient, webUrl, fileRef);
    throw err;
  }
};

// Upload new file version (file replace only)

export const uploadNewFileVersion = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number,
  uploadFn: () => Promise<void>
): Promise<void> => {
  const fileRef = await getFileRef(spHttpClient, webUrl, itemId);
  await checkoutFile(spHttpClient, webUrl, fileRef);
  try {
    await uploadFn();
    await checkinMajor(spHttpClient, webUrl, fileRef, 'New file version uploaded');
  } catch (err) {
    await undoCheckout(spHttpClient, webUrl, fileRef);
    throw err;
  }
};

export const checkinNewFile = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  fileRef: string,
  comment: string = 'Initial upload'
): Promise<void> => {
  // For brand new files, SharePoint auto-checks them out on creation.
  await checkinMajor(spHttpClient, webUrl, fileRef, comment);
};

export const saveMetadataAndCheckin = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  fileRef: string,
  metadataFn: () => Promise<void>
): Promise<void> => {
  if (!fileRef) {
    throw new Error('saveMetadataAndCheckin: fileRef is empty');
  }

  // STEP 1: Save all metadata while file is in SP auto-lock.
  // ValidateUpdateListItem works on auto-locked new files
  // and implicitly checks the file in after saving.
  // This creates v1.0 with all correct metadata in one step.
  console.log('saveMetadataAndCheckin: step1 saving metadata');
  try {
    await metadataFn();
    console.log('saveMetadataAndCheckin: step1 DONE - metadata saved');
  } catch (metaErr) {
    console.error('saveMetadataAndCheckin: metadata save failed:', metaErr);
    throw metaErr;
  }

  // Small wait for SP to process
  await new Promise(r => setTimeout(r, 300));

  // STEP 2: Attempt checkinMajor to ensure file is
  // properly checked in as Major version.
  // If ValidateUpdateListItem already checked it in
  // (423 response) - that is fine, file is already
  // at v1.0 with correct metadata. Treat 423 as success.
  console.log('saveMetadataAndCheckin: step2 checkinMajor -> v1.0');
  try {
    await checkinMajor(
      spHttpClient,
      webUrl,
      fileRef,
      'Initial upload'
    );
    console.log('saveMetadataAndCheckin: checkinMajor success');
  } catch (checkinErr: any) {
    const msg = checkinErr?.message || String(checkinErr);
    if (
      msg.includes('423') ||
      msg.includes('not checked out') ||
      msg.includes('SPFileCheckOutException')
    ) {
      console.log(
        'saveMetadataAndCheckin: checkinMajor 423 - ' +
        'file already checked in by SP. v1.0 created'
      );
      return;
    }
    console.error('saveMetadataAndCheckin: unexpected checkinMajor error:', checkinErr);
    throw checkinErr;
  }
};

// Get version history

export interface IVersionEntry {
  id: number;
  versionLabel: string;
  created: string;
  contentRefreshDate?: string;
  createdBy: string;
  createdByEmail: string;
  size: number;
  checkinComment: string;
  isCurrentVersion: boolean;
  fileRef: string;
  versionUrl: string;
  versionRawUrl?: string;
  fileName: string;
  fileType: string;
  itemId?: number;
}

export interface IVersionDetail {
  versionLabel: string;
  versionId: number;
  isCurrentVersion: boolean;
  created: string;
  modifiedDate: string;
  createdDate: string;
  title: string;
  status: string;
  fileName: string;
  fileType: string;
  description: string;
  author: string;
  authorEmail: string;
  modifiedBy: string;
  modifiedByEmail: string;
  createdBy: string;
  createdByEmail: string;
  published: string;
  sensitiveTerms: string;
  contentRefreshDate: string;
  projectId: string;
  documentType: string;
  geography: string;
  client: string;
  diseaseArea: string;
  therapyArea: string;
  bu: string;
  department: string;
  url: string;
}

const getVersionFieldText = (field: any): string => {
  if (!field) return '-';
  if (typeof field === 'string') {
    const normalizedTerms = field
      .split(/;#|[;]/)
      .map((entry) => entry.trim())
      .filter(Boolean)
      .filter((entry) => !/^-?\d+$/.test(entry.replace(/^#/, '')))
      .map((entry) => entry.replace(/^#/, ''))
      .map((entry) => entry.split('|')[0].trim())
      .filter(Boolean);
    return normalizedTerms.length > 0 ? normalizedTerms.join(', ') : '-';
  }
  if (typeof field === 'number' || typeof field === 'boolean') return String(field);
  if (Array.isArray(field)) {
    const labels = field.map(getVersionFieldText).filter((value) => value && value !== '-');
    return labels.length > 0 ? labels.join(', ') : '-';
  }

  const results = field?.results || field?.value || field?.Values || [];
  if (Array.isArray(results) && results.length > 0) {
    const labels = results
      .map(getVersionFieldText)
      .filter(Boolean);
    return labels.length > 0 ? labels.join(', ') : '-';
  }

  const knownValue =
    field.Label ||
    field.LookupValue ||
    field.Title ||
    field.Name ||
    field.Value ||
    field.Term ||
    field.Url ||
    field.Description ||
    field.EMail ||
    field.Email ||
    '';

  if (knownValue) return String(knownValue).trim() || '-';

  const primitiveValues = Object.keys(field)
    .filter((key) => key !== '__metadata' && key !== 'results')
    .map((key) => field[key])
    .filter((value) => typeof value === 'string' || typeof value === 'number')
    .map((value) => String(value).trim())
    .filter(Boolean);

  return primitiveValues.length > 0 ? primitiveValues.join(', ') : '-';
};

const KM_DATA_HUB_DEPARTMENT_VERSION_FIELD = COLUMN_NAMES.department;

type TaxonomyTextFieldAliases = Record<string, string[]>;

const taxonomyTextFieldAliasCache: Record<string, Promise<TaxonomyTextFieldAliases>> = {};
const TAXONOMY_FIELD_TITLE_ALIASES: Record<string, string[]> = {
  [COLUMN_NAMES.businessUnit]: ['BU', 'Business Unit'],
  [COLUMN_NAMES.department]: [
    'Department / Sub Department',
    'Department/Sub Department',
    'Department',
    'Sub Department'
  ],
  [COLUMN_NAMES.documentType]: ['Document Type'],
  [COLUMN_NAMES.client]: ['Client'],
  [COLUMN_NAMES.geography]: ['Geography'],
  [COLUMN_NAMES.diseaseArea]: ['Disease Area'],
  [COLUMN_NAMES.therapyArea]: ['Therapy Area']
};

const normalizeGuid = (value?: string): string =>
  (value || '').replace(/[{}]/g, '').toLowerCase();

const normalizeFieldName = (value?: string): string =>
  String(value || '').replace(/\s+/g, ' ').trim().toLowerCase();

const resolveContentRefreshDate = (
  item: Record<string, any>,
  textValues: Record<string, any> = {}
): string => (
  textValues[COLUMN_NAMES.contentRefreshDate] ||
  textValues.ContentRefreshDate ||
  textValues.Content_x0020_Refresh_x0020_Date ||
  item[COLUMN_NAMES.contentRefreshDate] ||
  item.ContentRefreshDate ||
  item.Content_x0020_Refresh_x0020_Date ||
  item.Content_x005f_x0020_x005f_Refresh_x005f_x0020_x005f_Date ||
  ''
);

const uniqueValues = (values: string[]): string[] =>
  values.filter((value, index, allValues) => value && allValues.indexOf(value) === index);

const getTaxonomyTextFieldAliases = async (
  spHttpClient: SPHttpClient,
  webUrl: string
): Promise<TaxonomyTextFieldAliases> => {
  const cacheKey = `${webUrl}|${LIBRARY_NAME}`;
  if (!taxonomyTextFieldAliasCache[cacheKey]) {
    taxonomyTextFieldAliasCache[cacheKey] = (async () => {
      try {
        const fieldsRes = await spHttpClient.get(
          `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
          `/fields?$select=Title,InternalName,Hidden,Id,TextField&$top=5000`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=verbose'
            }
          }
        );

        if (!fieldsRes.ok) {
          return {};
        }

        const fieldsJson = await fieldsRes.json();
        const fields = fieldsJson?.d?.results || fieldsJson?.value || [];
        const fieldMap: Record<string, any> = {};
        const fieldByTitle: Record<string, any> = {};
        const fieldById: Record<string, any> = {};

        fields.forEach((field: any) => {
          if (field?.InternalName) {
            fieldMap[field.InternalName] = field;
          }
          if (field?.Title) {
            fieldByTitle[normalizeFieldName(field.Title)] = field;
          }
          const id = normalizeGuid(field?.Id);
          if (id) {
            fieldById[id] = field;
          }
        });

        const taxonomyFields = [
          COLUMN_NAMES.businessUnit,
          COLUMN_NAMES.department,
          COLUMN_NAMES.documentType,
          COLUMN_NAMES.client,
          COLUMN_NAMES.geography,
          COLUMN_NAMES.diseaseArea,
          COLUMN_NAMES.therapyArea
        ];

        return taxonomyFields.reduce<TaxonomyTextFieldAliases>((aliases, internalName) => {
          const configuredTitleAliases = TAXONOMY_FIELD_TITLE_ALIASES[internalName] || [];
          const taxonomyField =
            fieldMap[internalName] ||
            configuredTitleAliases
              .map((title) => fieldByTitle[normalizeFieldName(title)])
              .filter(Boolean)[0];
          const textFieldId = normalizeGuid(taxonomyField?.TextField);
          const textFieldInternalName = textFieldId ? fieldById[textFieldId]?.InternalName : '';
          const resolvedInternalName = taxonomyField?.InternalName || '';
          aliases[internalName] = uniqueValues([
            `${internalName}_0`,
            resolvedInternalName,
            resolvedInternalName ? `${resolvedInternalName}_0` : '',
            textFieldInternalName,
            ...configuredTitleAliases
          ]);
          return aliases;
        }, {});
      } catch (err) {
        console.warn('Version detail taxonomy text field lookup failed:', err);
        return {};
      }
    })();
  }

  return taxonomyTextFieldAliasCache[cacheKey];
};

const mergeTaxonomyTextValues = (
  target: Record<string, any>,
  source: Record<string, any>,
  aliasesByField: TaxonomyTextFieldAliases
): void => {
  Object.keys(aliasesByField).forEach((internalName) => {
    const candidateNames = uniqueValues([
      internalName,
      ...(aliasesByField[internalName] || [])
    ]);
    const currentValue = target[internalName];
    const hasCurrentValue = getVersionFieldText(currentValue) !== '-';

    if (hasCurrentValue) {
      return;
    }

    for (const candidateName of candidateNames) {
      const candidateValue = source[candidateName];
      if (getVersionFieldText(candidateValue) !== '-') {
        target[internalName] = candidateValue;
        return;
      }
    }
  });
};

const splitVersionTaxonomyLabels = (field: any): string[] => {
  if (!field) return [];

  if (typeof field === 'string') {
    return field
      .split(/;#|[;\r\n]+/)
      .map((entry) => entry.trim())
      .filter(Boolean)
      .filter((entry) => !/^-?\d+$/.test(entry.replace(/^#/, '')))
      .map((entry) => entry.replace(/^#/, ''))
      .map((entry) => entry.split('|')[0].trim())
      .filter(Boolean);
  }

  if (Array.isArray(field)) {
    return field.reduce<string[]>((labels, entry) => {
      labels.push(...splitVersionTaxonomyLabels(entry));
      return labels;
    }, []);
  }

  const results = field?.results || field?.value || field?.Values || [];
  if (Array.isArray(results) && results.length > 0) {
    return splitVersionTaxonomyLabels(results);
  }

  const knownValue =
    field.Label ||
    field.LookupValue ||
    field.Title ||
    field.Name ||
    field.Value ||
    field.Term ||
    '';

  return knownValue ? [String(knownValue).trim()].filter(Boolean) : [];
};

const getVersionMetadataLabels = (
  version: any,
  textValues: Record<string, any>,
  internalName: string,
  aliases: string[] = []
): string[] => {
  const fieldAliases = uniqueValues([
    `${internalName}_0`,
    ...aliases
  ]);
  const candidates = [
    textValues?.[internalName],
    version?.[internalName],
    ...fieldAliases.reduce<any[]>((values, alias) => {
      values.push(textValues?.[alias], version?.[alias]);
      return values;
    }, [])
  ];

  return uniqueValues(
    candidates.reduce<string[]>((labels, candidate) => {
      labels.push(...splitVersionTaxonomyLabels(candidate));
      return labels;
    }, [])
  );
};

const normalizeTaxonomyLabel = (value: string): string =>
  String(value || '').trim().toLowerCase();

const parseVersionDepartmentLabel = (
  label: string,
  buLabels: string[]
): { department: string; subDepartment?: string } | null => {
  const rawSegments = label
    .split(/\s*[:>]\s*/)
    .map((segment) => segment.trim())
    .filter(Boolean);

  if (rawSegments.length === 0) {
    return null;
  }

  const normalizedBuLabels = buLabels.map(normalizeTaxonomyLabel);
  const departmentSegments =
    rawSegments.length > 1 && normalizedBuLabels.indexOf(normalizeTaxonomyLabel(rawSegments[0])) >= 0
      ? rawSegments.slice(1)
      : rawSegments;

  if (departmentSegments.length === 0) {
    return null;
  }

  if (departmentSegments.length > 1) {
    return {
      department: departmentSegments[0],
      subDepartment: departmentSegments.slice(1).join(': ')
    };
  }

  const leafLabel = departmentSegments[0];
  const normalizedLeafLabel = normalizeTaxonomyLabel(leafLabel);
  for (const businessUnit of buLabels) {
    const businessUnitDepartments = BUSINESS_UNIT_HIERARCHY[businessUnit] || {};
    const matchedDepartment = Object.keys(businessUnitDepartments).find(
      (departmentName) => normalizeTaxonomyLabel(departmentName) === normalizedLeafLabel
    );
    if (matchedDepartment) {
      return {
        department: matchedDepartment
      };
    }

    const matchedParentDepartment = Object.keys(businessUnitDepartments).find((departmentName) =>
      (businessUnitDepartments[departmentName] || []).some(
        (subDepartment) => normalizeTaxonomyLabel(subDepartment) === normalizedLeafLabel
      )
    );
    if (matchedParentDepartment) {
      const matchedSubDepartment = (businessUnitDepartments[matchedParentDepartment] || []).find(
        (subDepartment) => normalizeTaxonomyLabel(subDepartment) === normalizedLeafLabel
      );
      return {
        department: matchedParentDepartment,
        subDepartment: matchedSubDepartment || leafLabel
      };
    }
  }

  return {
    department: leafLabel
  };
};

const formatVersionDepartmentText = (
  departmentLabels: string[],
  buLabels: string[]
): string => {
  const departmentMap = new Map<string, Set<string>>();

  departmentLabels.forEach((label) => {
    const parsedLabel = parseVersionDepartmentLabel(label, buLabels);
    if (!parsedLabel?.department) {
      return;
    }

    if (!departmentMap.has(parsedLabel.department)) {
      departmentMap.set(parsedLabel.department, new Set<string>());
    }

    if (parsedLabel.subDepartment) {
      departmentMap.get(parsedLabel.department)?.add(parsedLabel.subDepartment);
    }
  });

  const formattedDepartments = Array.from(departmentMap.entries()).map(([department, subDepartments]) => {
    const subDepartmentList = Array.from(subDepartments);
    return subDepartmentList.length > 0
      ? `${department}: ${subDepartmentList.join(', ')}`
      : department;
  });

  return formattedDepartments.join('; ');
};

const getVersionDepartmentText = (
  version: any,
  textValues: Record<string, any> = {}
): string => {
  const departmentAliases = [
    'Department / Sub Department',
    'Department/Sub Department',
    'Department',
    'Sub Department',
    'Department_x0020_Sub_x0020_Department',
    'Department_x002f_Sub_x0020_Department'
  ];
  const buLabels = getVersionMetadataLabels(version, textValues, COLUMN_NAMES.businessUnit, ['Business Unit', 'BU']);
  const departmentLabels = getVersionMetadataLabels(version, textValues, KM_DATA_HUB_DEPARTMENT_VERSION_FIELD, departmentAliases);
  const resolvedDepartmentText = formatVersionDepartmentText(departmentLabels, buLabels);

  if (resolvedDepartmentText) {
    return resolvedDepartmentText;
  }

  if (buLabels.length > 1) {
    return formatVersionDepartmentText(buLabels.slice(1), buLabels.slice(0, 1)) || buLabels.slice(1).join('; ');
  }

  return getVersionMetadataText(version, textValues, KM_DATA_HUB_DEPARTMENT_VERSION_FIELD, departmentAliases);
};

const getVersionMetadataText = (
  version: any,
  textValues: Record<string, any>,
  internalName: string,
  aliases: string[] = []
): string => {
  const fieldAliases = [
    `${internalName}_0`,
    ...aliases
  ];
  const candidates = [
    textValues?.[internalName],
    version?.[internalName],
    ...fieldAliases.reduce<any[]>((values, alias) => {
      values.push(textValues?.[alias], version?.[alias]);
      return values;
    }, [])
  ];

  for (const candidate of candidates) {
    const resolvedText = getVersionFieldText(candidate);
    if (resolvedText && resolvedText !== '-') {
      return resolvedText;
    }
  }

  return '-';
};

const getVersionPerson = (field: any): { name: string; email: string } => {
  if (!field) return { name: 'Unknown', email: '' };
  const values = field?.results || field?.value || (Array.isArray(field) ? field : [field]);
  const first = Array.isArray(values) ? values[0] : values;
  if (!first) return { name: 'Unknown', email: '' };
  if (typeof first === 'string') return { name: first, email: '' };
  return {
    name: first.LookupValue || first.Title || first.Name || first.DisplayName || first.Email || first.EMail || 'Unknown',
    email: first.Email || first.EMail || first.UserPrincipalName || ''
  };
};

const mapVersionDetail = (
  v: any,
  versionId: number,
  textValues: Record<string, any> = {}
): IVersionDetail => {
  const customAuthor = getVersionPerson(v.Author0);
  const editedValue = textValues[COLUMN_NAMES.edited] || v[COLUMN_NAMES.edited];
  const editedByValue =
    v[COLUMN_NAMES.editedBy] ||
    textValues[COLUMN_NAMES.editedBy] ||
    v.Editor ||
    v.ModifiedBy ||
    v.Modified_x0020_By;
  const modifiedBy = getVersionPerson(editedByValue);
  const createdBy = getVersionPerson(v.Author || v.CreatedBy || v.Created_x0020_By);

  return {
    versionLabel: v.VersionLabel || '',
    versionId: v.VersionId || versionId,
    isCurrentVersion: v.IsCurrentVersion || false,
    created:
      textValues[COLUMN_NAMES.created] ||
      v.Created_x0020_Date ||
      v.CreatedDate ||
      v[COLUMN_NAMES.created] ||
      '',
    modifiedDate:
      editedValue ||
      textValues.Modified ||
      v.Modified ||
      textValues.Created ||
      v.Created ||
      '',
    createdDate:
      textValues[COLUMN_NAMES.created] ||
      v.Created_x0020_Date ||
      v.CreatedDate ||
      v[COLUMN_NAMES.created] ||
      '',
    title: v.Title || '-',
    status: v.Status || '-',
    fileName: v.VersionFileName || v.FileLeafRef || '-',
    fileType:
      v.VersionFileType ||
      v.File_x005f_x0020_x005f_Type ||
      '-',
    description: v.Description || v.Abstract || '-',
    author: customAuthor.name,
    authorEmail: customAuthor.email,
    modifiedBy: modifiedBy.name,
    modifiedByEmail: modifiedBy.email,
    createdBy: createdBy.name,
    createdByEmail: createdBy.email,
    published: textValues.Published || v.Published || '',
    sensitiveTerms: getVersionFieldText(v.SensitiveTerms || v.Sensitive_x0020_Terms),
    contentRefreshDate: resolveContentRefreshDate(v, textValues),
    projectId: getVersionMetadataText(v, textValues, COLUMN_NAMES.projectId, ['ProjectID', 'Project_x0020_Id', 'Project_x0020_ID']),
    documentType: getVersionMetadataText(v, textValues, COLUMN_NAMES.documentType, ['Document Type', 'Document_Type', 'Document_x005f_x0020_x005f_Type']),
    geography: getVersionMetadataText(v, textValues, COLUMN_NAMES.geography, ['Geography']),
    client: getVersionMetadataText(v, textValues, COLUMN_NAMES.client, ['Client']),
    diseaseArea: getVersionMetadataText(v, textValues, COLUMN_NAMES.diseaseArea, ['Disease Area', 'Disease_Area', 'Disease_x005f_x0020_x005f_Area']),
    therapyArea: getVersionMetadataText(v, textValues, COLUMN_NAMES.therapyArea, ['Therapy Area', 'Therapy_Area', 'Therapy_x005f_x0020_x005f_Area']),
    bu: getVersionMetadataText(v, textValues, COLUMN_NAMES.businessUnit, ['Business Unit', 'BU']),
    department: getVersionDepartmentText(v, textValues),
    url: v.URL?.Url || '-'
  };
};

export const fetchVersionDetail = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number,
  versionId: number
): Promise<IVersionDetail> => {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
    `/items(${itemId})/versions(${versionId})`;

  const res = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata=verbose'
      }
    }
  );

  const data = await res.json();
  const v = data?.d || data;
  let textValues: Record<string, any> = {};
  try {
    const textUrl =
      `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
      `/items(${itemId})/versions(${versionId})/FieldValuesAsText`;
    const textRes = await spHttpClient.get(
      textUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose'
        }
      }
    );
    if (textRes.ok) {
      const textData = await textRes.json();
      textValues = textData?.d || textData || {};
    }
  } catch (textErr) {
    console.warn('Version detail FieldValuesAsText load failed:', textErr);
  }

  try {
    const taxonomyTextFieldAliases = await getTaxonomyTextFieldAliases(spHttpClient, webUrl);
    mergeTaxonomyTextValues(textValues, textValues, taxonomyTextFieldAliases);
    mergeTaxonomyTextValues(textValues, v, taxonomyTextFieldAliases);
  } catch (taxonomyAliasErr) {
    console.warn('Version detail taxonomy alias merge failed:', taxonomyAliasErr);
  }

  if (!textValues[COLUMN_NAMES.created]) {
    try {
      const [currentCreatedRes, currentTextRes] = await Promise.all([
      spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
        `/items(${itemId})?$select=${COLUMN_NAMES.created}`,
        SPHttpClient.configurations.v1
      ),
      spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
        `/items(${itemId})/FieldValuesAsText`,
        SPHttpClient.configurations.v1
      )
      ]);
      const currentCreatedJson = currentCreatedRes.ok ? await currentCreatedRes.json() : {};
      const currentCreatedItem = currentCreatedJson?.d || currentCreatedJson || {};
      const currentTextJson = currentTextRes.ok ? await currentTextRes.json() : {};
      const currentTextValues = currentTextJson?.d || currentTextJson || {};
      const currentCreatedValue =
        currentTextValues[COLUMN_NAMES.created] ||
        currentCreatedItem[COLUMN_NAMES.created] ||
        '';

      if (currentCreatedValue) {
        textValues[COLUMN_NAMES.created] = currentCreatedValue;
      }
    } catch (currentCreatedErr) {
      console.warn('Version detail current created fallback failed:', currentCreatedErr);
    }
  }

  return mapVersionDetail(v, versionId, textValues);
};

export const fetchCurrentVersionDetail = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number
): Promise<IVersionDetail | null> => {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
    `/items(${itemId})/versions` +
    `?$filter=IsCurrentVersion eq true&$top=1`;

  const res = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata=verbose'
      }
    }
  );
  const data = await res.json();
  const results = data?.d?.results || data?.value || [];
  if (results.length === 0) return null;

  const versionId = results[0].VersionId || results[0].ID || results[0].Id || 0;
  return fetchVersionDetail(spHttpClient, webUrl, itemId, versionId);
};

export const getVersionHistory = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number
): Promise<IVersionEntry[]> => {
  const fileRef = await getFileRef(spHttpClient, webUrl, itemId);

  // STEP 1: Fetch old versions. SharePoint /Versions excludes the current version.
  const versionsUrl =
    `${webUrl}/_api/web/GetFileByServerRelativeUrl` +
    `('${encodeFileRef(fileRef)}')/Versions` +
    `?$select=ID,VersionLabel,Created,CreatedBy/Title,` +
    `CreatedBy/Email,Size,CheckInComment,Url` +
    `&$expand=CreatedBy`;
  const versionsRes: SPHttpClientResponse = await spHttpClient.get(
    versionsUrl,
    SPHttpClient.configurations.v1
  );
  const versionsData = await versionsRes.json();

  const oldVersions = (
    versionsData?.value ||
    versionsData?.d?.results ||
    []
  );

  // STEP 2: Fetch list item versions to get file name/type per version.
  const listVersionsUrl =
    `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
    `/items(${itemId})/versions` +
    `?$select=VersionId,VersionLabel,FileLeafRef,` +
    `File_x005f_x0020_x005f_Type,` +
    `VersionFileName,VersionFileType,${COLUMN_NAMES.contentRefreshDate}`;

  const listVersionsRes = await spHttpClient.get(
    listVersionsUrl,
    SPHttpClient.configurations.v1,
    {
      headers: { Accept: 'application/json;odata=verbose' }
    }
  );
  const listVersionsData = await listVersionsRes.json();
  const listVersionItems =
    listVersionsData?.d?.results ||
    listVersionsData?.value ||
    [];

  const versionMetaMap: Record<string, {
    fileName: string;
    fileType: string;
    contentRefreshDate: string;
  }> = {};

  listVersionItems.forEach((lv: any) => {
    const label = lv.VersionLabel || '';
    const fileName =
      lv.VersionFileName ||
      lv.FileLeafRef ||
      '';
    const fileType =
      lv.VersionFileType ||
      lv.File_x005f_x0020_x005f_Type ||
      fileName.split('.').pop() ||
      '';
    versionMetaMap[label] = {
      fileName: fileName,
      fileType: fileType,
      contentRefreshDate: resolveContentRefreshDate(lv)
    };
  });

  console.log('Version metadata map:', versionMetaMap);

  // STEP 3: Fetch the current version from the file itself.
  const currentUrl =
    `${webUrl}/_api/web/GetFileByServerRelativeUrl` +
    `('${encodeFileRef(fileRef)}')` +
    `?$select=UIVersionLabel,TimeLastModified,Size,` +
    `CheckInComment,Author/Title,Author/EMail,Name` +
    `&$expand=Author`;
  const currentRes = await spHttpClient.get(
    currentUrl,
    SPHttpClient.configurations.v1
  );
  const currentData = await currentRes.json();
  const currentVersionLabel =
    currentData?.UIVersionLabel ||
    currentData?.d?.UIVersionLabel ||
    '';
  const currentModified =
    currentData?.TimeLastModified ||
    currentData?.d?.TimeLastModified ||
    '';
  const currentSize =
    currentData?.Size ||
    currentData?.d?.Size ||
    0;
  const currentCheckinComment =
    currentData?.CheckInComment ||
    currentData?.d?.CheckInComment ||
    '';
  const currentAuthorTitle =
    currentData?.Author?.Title ||
    currentData?.d?.Author?.Title ||
    'Unknown';
  const currentAuthorEmail =
    currentData?.Author?.EMail ||
    currentData?.d?.Author?.EMail ||
    '';
  const currentFileName =
    currentData?.Name ||
    currentData?.d?.Name ||
    fileRef.split('/').pop() ||
    '';
  let currentContentRefreshDate = '';
  try {
    const [currentItemRes, currentTextRes] = await Promise.all([
      spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
        `/items(${itemId})?$select=${COLUMN_NAMES.contentRefreshDate}`,
        SPHttpClient.configurations.v1
      ),
      spHttpClient.get(
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
        `/items(${itemId})/FieldValuesAsText`,
        SPHttpClient.configurations.v1
      )
    ]);
    const currentItemJson = currentItemRes.ok ? await currentItemRes.json() : {};
    const currentItem = currentItemJson?.d || currentItemJson || {};
    const currentTextJson = currentTextRes.ok ? await currentTextRes.json() : {};
    const currentTextValues = currentTextJson?.d || currentTextJson || {};
    currentContentRefreshDate = resolveContentRefreshDate(currentItem, currentTextValues);
  } catch (currentContentRefreshErr) {
    console.warn('Version history current content refresh date load failed:', currentContentRefreshErr);
  }

  const currentEntry: IVersionEntry = {
    id: -1,
    itemId: itemId,
    versionLabel: currentVersionLabel,
    created: currentModified,
    contentRefreshDate: currentContentRefreshDate,
    createdBy: currentAuthorTitle,
    createdByEmail: currentAuthorEmail,
    size: currentSize,
    checkinComment: currentCheckinComment,
    isCurrentVersion: true,
    fileRef: fileRef,
    versionUrl:
      `${webUrl}/_layouts/15/WopiFrame.aspx` +
      `?sourcedoc=${encodeURIComponent(fileRef)}&action=view`,
    fileName: currentFileName,
    fileType: currentFileName.split('.').pop() || ''
  };

  const oldEntries: IVersionEntry[] = oldVersions
    .filter((v: any) => /^\d+\.0$/.test(v.VersionLabel || ''))
    .map((v: any) => {
      const vId = v.ID ?? v.Id ?? 0;
      const meta = versionMetaMap[v.VersionLabel] || {
        fileName: '',
        fileType: '',
        contentRefreshDate: ''
      };
      return {
      id: vId,
      itemId: itemId,
      versionLabel: v.VersionLabel as string,
      created: v.Created,
      contentRefreshDate: meta.contentRefreshDate,
      createdBy: v.CreatedBy?.Title || 'Unknown',
      createdByEmail: v.CreatedBy?.Email || '',
      size: v.Size || 0,
      checkinComment: v.CheckInComment || '',
      isCurrentVersion: false,
      fileRef: fileRef,
      fileName: meta.fileName || fileRef.split('/').pop() || '',
      fileType: meta.fileType || '',
      versionUrl: (() => {
        const vtiPath = v.Url || '';
        let resolvedVersionUrl = '';
        if (!vtiPath) {
          resolvedVersionUrl = `${webUrl}/_layouts/15/WopiFrame.aspx` +
            `?sourcedoc=${encodeURIComponent(fileRef)}&action=view`;
          console.log(
            'Version', v.VersionLabel,
            'vtiPath:', v.Url,
            'versionUrl:', resolvedVersionUrl
          );
          return resolvedVersionUrl;
        }

        let siteRelativePath = '';
        try {
          siteRelativePath = new URL(webUrl).pathname.replace(/\/$/, '');
        } catch {
          siteRelativePath = webUrl.replace(/^https?:\/\/[^/]+/, '').replace(/\/$/, '');
        }
        const vtiServerRelative =
          `${siteRelativePath}/${String(vtiPath).replace(/^\//, '')}`;
        const ext = (
          meta.fileType ||
          String(vtiPath).split('.').pop() ||
          fileRef.split('.').pop() ||
          ''
        ).toLowerCase();
        const officeExts = [
          'pptx', 'ppt', 'docx', 'doc',
          'xlsx', 'xls', 'xlsm', 'xlsb'
        ];

        if (officeExts.includes(ext)) {
          const origin = webUrl.split('/sites/')[0];
          resolvedVersionUrl = `${webUrl}/_layouts/15/WopiFrame.aspx` +
            `?sourcedoc=${encodeURIComponent(`${origin}${vtiServerRelative}`)}` +
            `&action=view`;
        } else {
          const origin = webUrl.split('/sites/')[0];
          resolvedVersionUrl = `${origin}${vtiServerRelative}`;
        }

        console.log(
          'Version', v.VersionLabel,
          'vtiPath:', v.Url,
          'versionUrl:', resolvedVersionUrl
        );

        return resolvedVersionUrl;
      })(),
      versionRawUrl: v.Url || ''
      };
    });

  return [currentEntry, ...oldEntries];
};

export const restoreVersion = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  fileRef: string,
  versionLabel: string,
  restoredItemId?: number
): Promise<void> => {
  let targetRestoredName = '';
  let targetRestoredType = '';

  if (restoredItemId) {
    try {
      const targetMetaUrl =
        `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
        `/items(${restoredItemId})/versions` +
        `?$select=VersionId,VersionLabel,VersionFileName,` +
        `VersionFileType,FileLeafRef`;
      const targetMetaRes = await spHttpClient.get(
        targetMetaUrl,
        SPHttpClient.configurations.v1,
        {
          headers: { Accept: 'application/json;odata=verbose' }
        }
      );
      const targetMetaData = await targetMetaRes.json();
      const targetItems =
        targetMetaData?.d?.results ||
        targetMetaData?.value ||
        [];
      const targetVersion = targetItems.find((entry: any) =>
        String(entry.VersionLabel || '') === String(versionLabel || '')
      );

      targetRestoredName =
        targetVersion?.VersionFileName ||
        targetVersion?.FileLeafRef ||
        '';
      targetRestoredType =
        targetVersion?.VersionFileType ||
        targetRestoredName.split('.').pop() ||
        '';

      console.log(
        'Restore: target version file metadata:',
        { targetRestoredName, targetRestoredType }
      );
    } catch (targetMetaErr) {
      console.warn('Restore: could not fetch target version file metadata:', targetMetaErr);
    }
  }

  let activeFileRef = fileRef;
  await checkoutFile(spHttpClient, webUrl, fileRef);

  try {
    const url =
      `${webUrl}/_api/web/GetFileByServerRelativeUrl` +
      `('${encodeFileRef(fileRef)}')/Versions` +
      `/RestoreByLabel(versionlabel='${encodeURIComponent(versionLabel)}')`;

    console.log('Restore: calling RestoreByLabel', url);

    const response = await spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        }
      }
    );

    console.log('Restore: response status', response.status);

    if (!response.ok) {
      const text = await response.text();
      console.error('RestoreByLabel failed:', response.status, text);
      await undoCheckout(spHttpClient, webUrl, fileRef);
      throw new Error(`Restore failed: ${response.status} ${text}`);
    }

    if (targetRestoredName) {
      const currentFileName = activeFileRef.split('/').pop() || '';
      const isSameFile =
        targetRestoredName.toLowerCase() === currentFileName.toLowerCase();

      if (!isSameFile) {
        const folderRef = activeFileRef.substring(0, activeFileRef.lastIndexOf('/'));
        const restoredFileRef = `${folderRef}/${targetRestoredName}`;
        const moveResponse = await spHttpClient.post(
          `${webUrl}/_api/web/GetFileByServerRelativeUrl('${escapeODataString(activeFileRef)}')` +
          `/MoveTo(newUrl='${escapeODataString(restoredFileRef)}',flags=1)`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: 'application/json;odata=nometadata'
            }
          }
        );

        if (!moveResponse.ok) {
          const moveText = await moveResponse.text();
          throw new Error(`Restore file rename failed: ${moveResponse.status} ${moveText}`);
        }

        activeFileRef = restoredFileRef;
        console.log('Restore: actual file renamed to restored version name:', activeFileRef);
      }
    }

    await checkinMajor(
      spHttpClient,
      webUrl,
      activeFileRef,
      `Restored to version ${versionLabel}`
    );

    console.log('Restore: checkinMajor complete');

    if (restoredItemId) {
      try {
        await new Promise((resolve) => setTimeout(resolve, 500));

        const restoredMetaUrl =
          `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
          `/items(${restoredItemId})/versions` +
          `?$select=VersionId,VersionLabel,VersionFileName,` +
          `VersionFileType,FileLeafRef,IsCurrentVersion` +
          `&$filter=IsCurrentVersion eq true&$top=1`;

        const restoredMetaRes = await spHttpClient.get(
          restoredMetaUrl,
          SPHttpClient.configurations.v1,
          {
            headers: { Accept: 'application/json;odata=verbose' }
          }
        );

        const restoredMetaData = await restoredMetaRes.json();
        const restoredItems =
          restoredMetaData?.d?.results ||
          restoredMetaData?.value ||
          [];

        if (restoredItems.length > 0) {
          const restoredMeta = restoredItems[0];
          const restoredName =
            targetRestoredName ||
            restoredMeta.VersionFileName ||
            restoredMeta.FileLeafRef ||
            '';
          const restoredType =
            targetRestoredType ||
            restoredMeta.VersionFileType ||
            restoredName.split('.').pop() ||
            '';

          console.log(
            'Restore: restored version metadata:',
            { restoredName, restoredType }
          );

          if (restoredName) {
            let existingTitle = '';
            try {
              const titleResp = await spHttpClient.get(
                `${webUrl}/_api/web/lists/getbytitle` +
                `('${LIBRARY_NAME}')/items(${restoredItemId})` +
                `?$select=Title`,
                SPHttpClient.configurations.v1
              );
              if (titleResp.ok) {
                const titleData = await titleResp.json();
                existingTitle =
                  titleData?.d?.Title ||
                  titleData?.Title ||
                  '';
              }
            } catch {
              console.warn('Restore: could not fetch title');
            }

            await updateWithoutVersion(
              spHttpClient,
              webUrl,
              restoredItemId,
              async () => {
                const metaResponse = await spHttpClient.post(
                  `${webUrl}/_api/web/lists/getbytitle` +
                  `('${LIBRARY_NAME}')/items(${restoredItemId})` +
                  `/ValidateUpdateListItem`,
                  SPHttpClient.configurations.v1,
                  {
                    headers: {
                      Accept: 'application/json;odata=verbose',
                      'Content-Type': 'application/json;odata=verbose',
                      'odata-version': ''
                    },
                    body: JSON.stringify({
                      formValues: [
                        {
                          FieldName: COLUMN_NAMES.title,
                          FieldValue: existingTitle || ''
                        },
                        {
                          FieldName: COLUMN_NAMES.status,
                          FieldValue: 'Under Review'
                        },
                        {
                          FieldName: COLUMN_NAMES.published,
                          FieldValue: ''
                        },
                        {
                          FieldName: COLUMN_NAMES.versionFileName,
                          FieldValue: restoredName
                        },
                        {
                          FieldName: COLUMN_NAMES.versionFileType,
                          FieldValue: restoredType
                        },
                        {
                          FieldName: COLUMN_NAMES.contentRefreshDate,
                          FieldValue: formatSharePointDateFieldValue()
                        }
                      ],
                      bNewDocumentUpdate: false
                    })
                  }
                );

                if (metaResponse.ok) {
                  console.log(
                    'Restore metadata updated',
                    'VersionFileName:',
                    restoredName,
                    'VersionFileType:',
                    restoredType,
                    'Status -> Under Review',
                    'Published -> cleared',
                    'Title preserved:',
                    existingTitle
                  );
                } else {
                  const errText = await metaResponse.text();
                  console.warn(
                    'Restore metadata update failed:',
                    metaResponse.status,
                    errText
                  );
                }
              }
            );
          }
        }
      } catch (metaErr) {
        console.warn(
          'Restore: metadata update after restore failed:',
          metaErr
        );
      }
    }

    console.log('Restore complete version:', versionLabel);
  } catch (err) {
    try {
      await undoCheckout(spHttpClient, webUrl, activeFileRef);
    } catch {
      // Ignore cleanup errors.
    }
    throw err;
  }
};

export const deleteVersion = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  fileRef: string,
  versionId: number
): Promise<void> => {
  const url =
    `${webUrl}/_api/web/GetFileByServerRelativeUrl` +
    `('${encodeFileRef(fileRef)}')/Versions(${versionId})/DeleteObject()`;
  const response = await spHttpClient.post(
    url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'X-HTTP-Method': 'DELETE',
        'IF-MATCH': '*'
      }
    }
  );
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Delete version failed: ${response.status} ${text}`);
  }
};

// Get current version

export const getCurrentVersion = async (
  spHttpClient: SPHttpClient,
  webUrl: string,
  itemId: number
): Promise<string> => {
  const url =
    `${webUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')` +
    `/items(${itemId})?$select=OData__UIVersionString`;
  const res: SPHttpClientResponse = await spHttpClient.get(
    url,
    SPHttpClient.configurations.v1
  );
  const data = await res.json();
  return (data?.OData__UIVersionString as string) || (data?.d?.OData__UIVersionString as string) || '1.0';
};
