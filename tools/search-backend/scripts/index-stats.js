const { loadEnv } = require('./lib/env');
const { searchIndex } = require('./lib/azureSearchClient');

const envPath = process.argv[2] || 'config/search.env';

const facetFields = [
  'fileExtension',
  'bu',
  'department',
  'client',
  'region',
  'therapyArea',
  'diseaseArea',
  'documentType'
];

const topFacetValues = (facets, field) => (facets[field] || [])
  .slice(0, 12)
  .map((item) => ({ value: item.value, count: item.count }));

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  const [documents, chunks, people, placeholderSearch] = await Promise.all([
    searchIndex(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, {
      search: '*',
      filter: "status eq 'Active'",
      count: true,
      top: 0,
      facets: facetFields
    }),
    searchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
      search: '*',
      filter: "status eq 'Active'",
      count: true,
      top: 0,
      facets: ['sourceKind']
    }),
    searchIndex(env, env.AZURE_SEARCH_PEOPLE_INDEX, {
      search: '*',
      filter: 'active eq true',
      count: true,
      top: 0
    }),
    searchIndex(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, {
      search: '"not applicable"',
      filter: "status eq 'Active'",
      count: true,
      top: 5,
      select: 'title,fileName,client,region,therapyArea,diseaseArea,documentType'
    })
  ]);

  const facets = documents['@search.facets'] || {};
  const output = {
    env: absolutePath,
    counts: {
      documents: documents['@odata.count'] || 0,
      chunks: chunks['@odata.count'] || 0,
      people: people['@odata.count'] || 0
    },
    averageChunksPerDocument: documents['@odata.count']
      ? Number(((chunks['@odata.count'] || 0) / documents['@odata.count']).toFixed(2))
      : 0,
    facets: Object.fromEntries(facetFields.map((field) => [field, topFacetValues(facets, field)])),
    sourceKinds: chunks['@search.facets'] || {},
    placeholderSearch: {
      query: '"not applicable"',
      count: placeholderSearch['@odata.count'] || 0,
      examples: (placeholderSearch.value || []).map((item) => ({
        title: item.title,
        fileName: item.fileName,
        client: item.client,
        region: item.region,
        therapyArea: item.therapyArea,
        diseaseArea: item.diseaseArea,
        documentType: item.documentType
      }))
    }
  };

  console.log(JSON.stringify(output, null, 2));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
