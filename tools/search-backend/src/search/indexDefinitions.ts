export interface AzureSearchFieldDefinition {
  name: string;
  type: string;
  key?: boolean;
  searchable?: boolean;
  filterable?: boolean;
  sortable?: boolean;
  facetable?: boolean;
  retrievable?: boolean;
  dimensions?: number;
  vectorSearchProfile?: string;
}

export interface AzureSearchIndexDefinition {
  name: string;
  fields: AzureSearchFieldDefinition[];
  vectorSearch?: {
    algorithms: Array<{
      name: string;
      kind: 'hnsw';
      hnswParameters: {
        metric: 'cosine';
        m: number;
        efConstruction: number;
        efSearch: number;
      };
    }>;
    profiles: Array<{
      name: string;
      algorithm: string;
    }>;
  };
  semantic?: {
    configurations: Array<{
      name: string;
      prioritizedFields: {
        titleField?: { fieldName: string };
        prioritizedContentFields?: Array<{ fieldName: string }>;
        prioritizedKeywordsFields?: Array<{ fieldName: string }>;
      };
    }>;
  };
}

const stringField = (
  name: string,
  options: Omit<AzureSearchFieldDefinition, 'name' | 'type'> = {}
): AzureSearchFieldDefinition => ({
  name,
  type: 'Edm.String',
  retrievable: true,
  ...options
});

const stringCollectionField = (
  name: string,
  options: Omit<AzureSearchFieldDefinition, 'name' | 'type'> = {}
): AzureSearchFieldDefinition => ({
  name,
  type: 'Collection(Edm.String)',
  retrievable: true,
  ...options
});

const dateField = (
  name: string,
  options: Omit<AzureSearchFieldDefinition, 'name' | 'type'> = {}
): AzureSearchFieldDefinition => ({
  name,
  type: 'Edm.DateTimeOffset',
  retrievable: true,
  ...options
});

const numberField = (
  name: string,
  options: Omit<AzureSearchFieldDefinition, 'name' | 'type'> = {}
): AzureSearchFieldDefinition => ({
  name,
  type: 'Edm.Int32',
  retrievable: true,
  ...options
});

export interface SearchIndexDefinitionOptions {
  documentsIndexName: string;
  chunksIndexName: string;
  peopleIndexName: string;
  vectorDimensions: number;
  vectorProfileName: string;
  vectorAlgorithmName: string;
  semanticConfigName: string;
}

const longSearchableTextField = (name: string) =>
  stringField(name, {
    searchable: true,
    filterable: false,
    sortable: false,
    facetable: false
  });

export const buildDocumentsIndexDefinition = (
  options: SearchIndexDefinitionOptions
): AzureSearchIndexDefinition => ({
  name: options.documentsIndexName,
  fields: [
    stringField('id', { key: true }),
    stringField('siteId', { filterable: true }),
    stringField('driveId', { filterable: true }),
    stringField('listId', { filterable: true }),
    stringField('listItemId', { filterable: true }),
    stringField('driveItemId', { filterable: true }),
    stringField('fileUniqueId', { filterable: true }),
    stringField('uniqueId', { filterable: true }),
    stringField('fileRef'),
    stringField('serverRelativeUrl'),
    stringField('fileName', { searchable: true, filterable: true, sortable: true }),
    stringField('title', { searchable: true, filterable: true, sortable: true }),
    stringField('webUrl'),
    stringField('fileExtension', { filterable: true, facetable: true }),
    stringField('mimeType', { filterable: true }),
    stringField('status', { filterable: true }),
    stringField('bu', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('buFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('department', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('departmentFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('departmentFacetValues', { searchable: true, filterable: true, facetable: true }),
    stringField('diseaseArea', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('diseaseAreaFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('therapyArea', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('therapyAreaFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('client', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('clientFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('region', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('regionFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('documentType', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('documentTypeFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('authors', { searchable: true, filterable: true, facetable: true }),
    stringField('modifiedBy', { searchable: true, filterable: true }),
    dateField('createdDateTime', { filterable: true, sortable: true }),
    dateField('created', { filterable: true, sortable: true }),
    dateField('lastModifiedDateTime', { filterable: true, sortable: true }),
    dateField('modified', { filterable: true, sortable: true }),
    dateField('publishedDate', { filterable: true, sortable: true }),
    dateField('contentRefreshDate', { filterable: true, sortable: true }),
    longSearchableTextField('description'),
    stringField('contentPreview', { searchable: true }),
    stringField('projectId', { searchable: true, filterable: true }),
    stringField('sensitiveTerms', { searchable: true }),
    stringField('version', { filterable: true }),
    dateField('indexedAt', { filterable: true, sortable: true })
  ]
});

export const buildChunksIndexDefinition = (
  options: SearchIndexDefinitionOptions
): AzureSearchIndexDefinition => ({
  name: options.chunksIndexName,
  fields: [
    stringField('id', { key: true }),
    stringField('documentId', { filterable: true }),
    stringField('siteId', { filterable: true }),
    stringField('driveId', { filterable: true }),
    stringField('listId', { filterable: true }),
    stringField('listItemId', { filterable: true }),
    stringField('title', { searchable: true, filterable: true }),
    stringField('fileName', { searchable: true, filterable: true }),
    stringField('webUrl'),
    stringField('status', { filterable: true }),
    stringField('bu', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('buFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('department', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('departmentFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('departmentFacetValues', { searchable: true, filterable: true, facetable: true }),
    stringField('diseaseArea', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('diseaseAreaFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('therapyArea', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('therapyAreaFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('client', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('clientFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('region', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('regionFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringField('documentType', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('documentTypeFilterValues', { searchable: true, filterable: true, facetable: true }),
    stringCollectionField('authors', { searchable: true, filterable: true, facetable: true }),
    numberField('chunkOrdinal', { filterable: true, sortable: true }),
    longSearchableTextField('description'),
    stringField('chunkText', { searchable: true }),
    {
      name: 'chunkVector',
      type: 'Collection(Edm.Single)',
      searchable: true,
      retrievable: false,
      dimensions: options.vectorDimensions,
      vectorSearchProfile: options.vectorProfileName
    },
    numberField('pageNumber', { filterable: true, sortable: true }),
    stringField('sectionTitle', { searchable: true, filterable: true }),
    stringField('sourceKind', { filterable: true, facetable: true }),
    dateField('lastModifiedDateTime', { filterable: true, sortable: true }),
    dateField('indexedAt', { filterable: true, sortable: true })
  ],
  vectorSearch: {
    algorithms: [
      {
        name: options.vectorAlgorithmName,
        kind: 'hnsw',
        hnswParameters: {
          metric: 'cosine',
          m: 4,
          efConstruction: 400,
          efSearch: 500
        }
      }
    ],
    profiles: [
      {
        name: options.vectorProfileName,
        algorithm: options.vectorAlgorithmName
      }
    ]
  },
  semantic: {
    configurations: [
      {
        name: options.semanticConfigName,
        prioritizedFields: {
          titleField: { fieldName: 'title' },
          prioritizedContentFields: [
            { fieldName: 'chunkText' },
            { fieldName: 'description' }
          ],
          prioritizedKeywordsFields: [
            { fieldName: 'bu' },
            { fieldName: 'department' },
            { fieldName: 'diseaseArea' },
            { fieldName: 'therapyArea' },
            { fieldName: 'client' },
            { fieldName: 'region' },
            { fieldName: 'documentType' }
          ]
        }
      }
    ]
  }
});

export const buildPeopleIndexDefinition = (
  options: SearchIndexDefinitionOptions
): AzureSearchIndexDefinition => ({
  name: options.peopleIndexName,
  fields: [
    stringField('id', { key: true }),
    stringField('personName', { searchable: true, filterable: true, sortable: true }),
    stringField('email', { searchable: true, filterable: true }),
    longSearchableTextField('contacts'),
    longSearchableTextField('allEmails'),
    longSearchableTextField('allText'),
    longSearchableTextField('description'),
    longSearchableTextField('sectionTitles'),
    stringField('serviceLine', { searchable: true, filterable: true, facetable: true }),
    longSearchableTextField('contentIds'),
    longSearchableTextField('role'),
    stringField('team', { searchable: true, filterable: true, facetable: true }),
    stringField('bu', { searchable: true, filterable: true, facetable: true }),
    stringField('region', { searchable: true, filterable: true, facetable: true }),
    stringField('manager', { searchable: true, filterable: true }),
    stringCollectionField('skills', { searchable: true, filterable: true, facetable: true }),
    { name: 'active', type: 'Edm.Boolean', filterable: true, retrievable: true },
    stringField('listItemUrl'),
    dateField('lastModifiedDateTime', { filterable: true, sortable: true }),
    dateField('indexedAt', { filterable: true, sortable: true })
  ]
});

export const buildAllIndexDefinitions = (
  options: SearchIndexDefinitionOptions
): AzureSearchIndexDefinition[] => [
  buildDocumentsIndexDefinition(options),
  buildChunksIndexDefinition(options),
  buildPeopleIndexDefinition(options)
];
