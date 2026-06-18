const { loadEnv } = require('./lib/env');
const { createOrUpdateIndexes, deleteIndexIfExists } = require('./lib/azureSearchClient');
const { buildIndexDefinitions } = require('./lib/indexDefinitions');

const run = async () => {
  const envPath = process.argv[2] || 'config/search.env';
  const { env, absolutePath } = loadEnv(envPath);
  const definitions = buildIndexDefinitions(env);
  const isDevTarget = /\.dev\.env$/i.test(envPath)
    || definitions.every((definition) => String(definition.name || '').toLowerCase().includes('dev'));
  const confirmed = process.argv.includes('--yes-delete-indexes');

  if (!isDevTarget && !confirmed) {
    throw new Error('Refusing to reset non-dev search indexes. Re-run with --yes-delete-indexes only when you intentionally want to delete and recreate these indexes.');
  }

  console.log(`Using env: ${absolutePath}`);
  console.log('Deleting indexes if they exist...');

  for (const definition of definitions) {
    await deleteIndexIfExists(env, definition.name);
    console.log(`Deleted/absent: ${definition.name}`);
  }

  console.log('Creating indexes...');
  await createOrUpdateIndexes(env, definitions);
  definitions.forEach((definition) => console.log(`Created: ${definition.name}`));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
