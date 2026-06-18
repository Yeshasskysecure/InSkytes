const { loadEnv } = require('./lib/env');
const { deleteDocuments, searchIndex } = require('./lib/azureSearchClient');

const envPath = process.argv[2] || 'config/search.env';
const confirmed = process.argv.includes('--yes-purge-non-active');
const batchSize = 1000;

const purgeIndex = async (env, indexName, keyField, selectFields) => {
  let deleted = 0;

  while (true) {
    const result = await searchIndex(env, indexName, {
      search: '*',
      filter: "status ne 'Active'",
      select: selectFields,
      top: batchSize
    });
    const keys = (result.value || []).map((item) => item[keyField]).filter(Boolean);

    if (keys.length === 0) {
      break;
    }

    await deleteDocuments(env, indexName, keyField, keys);
    deleted += keys.length;
    console.log(`Deleted ${keys.length} non-active records from ${indexName}; total ${deleted}`);
  }

  return deleted;
};

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  console.log(`Using env: ${absolutePath}`);

  if (!confirmed) {
    throw new Error([
      'Refusing to purge non-active search records without explicit confirmation.',
      'Re-run with --yes-purge-non-active after confirming the target env and indexes.'
    ].join(' '));
  }

  const documentsDeleted = await purgeIndex(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, 'id', 'id,status,title,fileName');
  const chunksDeleted = await purgeIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, 'id', 'id,status,documentId,title,fileName');

  console.log(JSON.stringify({
    documentsDeleted,
    chunksDeleted
  }, null, 2));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
