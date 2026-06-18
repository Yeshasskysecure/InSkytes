import { DocumentSearchFilters, buildDocumentFilter } from './filters';

export interface HybridSearchRequestOptions {
  query: string;
  queryVector: number[];
  filters?: DocumentSearchFilters;
  semanticConfiguration: string;
  vectorField: string;
  k?: number;
  top?: number;
}

export interface AzureHybridSearchRequest {
  search: string;
  vectorQueries: Array<{
    kind: 'vector';
    vector: number[];
    fields: string;
    k: number;
    exhaustive: boolean;
  }>;
  vectorFilterMode: 'preFilter';
  queryType: 'semantic';
  semanticConfiguration: string;
  captions: 'extractive';
  answers: 'extractive';
  filter: string;
  select: string;
  top: number;
}

export const buildHybridChunkSearchRequest = (
  options: HybridSearchRequestOptions
): AzureHybridSearchRequest => {
  if (options.queryVector.length === 0) {
    throw new Error('Hybrid search requires a non-empty query vector.');
  }

  return {
    search: options.query.trim() || '*',
    vectorQueries: [
      {
        kind: 'vector',
        vector: options.queryVector,
        fields: options.vectorField,
        k: options.k || 50,
        exhaustive: false
      }
    ],
    vectorFilterMode: 'preFilter',
    queryType: 'semantic',
    semanticConfiguration: options.semanticConfiguration,
    captions: 'extractive',
    answers: 'extractive',
    filter: buildDocumentFilter(options.filters),
    select: [
      'id',
      'documentId',
      'title',
      'fileName',
      'webUrl',
      'description',
      'chunkText',
      'pageNumber',
      'sectionTitle',
      'sourceKind',
      'status',
      'bu',
      'department',
      'diseaseArea',
      'therapyArea',
      'client',
      'region',
      'documentType',
      'authors',
      'lastModifiedDateTime'
    ].join(','),
    top: options.top || 20
  };
};
