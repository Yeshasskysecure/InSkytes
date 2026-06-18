export type ExtractionStrategy = 'local-first' | 'doc-intel-first';

export interface SearchEnvironment {
  graphTenantId: string;
  graphClientId: string;
  graphClientSecret: string;
  sharePointHostname: string;
  sharePointSitePath: string;
  sharePointSiteId: string;
  sharePointLibraryListId: string;
  sharePointLibraryDriveId: string;
  sharePointWhosWhoListId: string;
  azureSearchEndpoint: string;
  azureSearchApiVersion: string;
  azureSearchAdminKey: string;
  azureSearchQueryKey: string;
  documentsIndexName: string;
  chunksIndexName: string;
  peopleIndexName: string;
  vectorProfileName: string;
  vectorAlgorithmName: string;
  semanticConfigName: string;
  vectorFieldName: string;
  vectorDimensions: number;
  hybridK: number;
  openAiEndpoint: string;
  openAiApiKey: string;
  openAiApiVersion: string;
  openAiUseLegacyApi: boolean;
  chatDeployment: string;
  embeddingDeployment: string;
  embeddingDimensions: number;
  extractionStrategy: ExtractionStrategy;
  documentIntelligenceEndpoint?: string;
  documentIntelligenceKey?: string;
  serviceBusConnectionString?: string;
  databaseUrl?: string;
}

export const requiredEnvironmentKeys = [
  'GRAPH_TENANT_ID',
  'GRAPH_CLIENT_ID',
  'GRAPH_CLIENT_SECRET',
  'SHAREPOINT_HOSTNAME',
  'SHAREPOINT_SITE_PATH',
  'SHAREPOINT_SITE_ID',
  'SHAREPOINT_LIBRARY_LIST_ID',
  'SHAREPOINT_LIBRARY_DRIVE_ID',
  'SHAREPOINT_WHOSWHO_LIST_ID',
  'AZURE_SEARCH_ENDPOINT',
  'AZURE_SEARCH_API_VERSION',
  'AZURE_SEARCH_ADMIN_KEY',
  'AZURE_SEARCH_QUERY_KEY',
  'AZURE_SEARCH_DOCUMENTS_INDEX',
  'AZURE_SEARCH_CHUNKS_INDEX',
  'AZURE_SEARCH_PEOPLE_INDEX',
  'AZURE_SEARCH_VECTOR_PROFILE',
  'AZURE_SEARCH_VECTOR_ALGORITHM',
  'AZURE_SEARCH_SEMANTIC_CONFIG',
  'AZURE_SEARCH_VECTOR_FIELD',
  'AZURE_SEARCH_VECTOR_DIMENSIONS',
  'AZURE_OPENAI_ENDPOINT',
  'AZURE_OPENAI_API_KEY',
  'AZURE_OPENAI_API_VERSION',
  'AZURE_OPENAI_CHAT_DEPLOYMENT',
  'AZURE_OPENAI_EMBEDDING_DEPLOYMENT',
  'AZURE_OPENAI_EMBEDDING_DIMENSIONS'
] as const;

export const optionalPhaseTwoEnvironmentKeys = [
  'AZURE_DOC_INTEL_ENDPOINT',
  'AZURE_DOC_INTEL_KEY',
  'AZURE_SERVICEBUS_CONNECTION_STRING',
  'DATABASE_URL',
  'WEBHOOK_BASE_URL',
  'GRAPH_WEBHOOK_CLIENT_STATE'
] as const;
