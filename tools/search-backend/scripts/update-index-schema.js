const { loadEnv } = require('./lib/env');
const { createOrUpdateIndexes, getIndex } = require('./lib/azureSearchClient');
const { buildIndexDefinitions } = require('./lib/indexDefinitions');

const fixedFieldAttributes = [
  'type',
  'key',
  'searchable',
  'filterable',
  'sortable',
  'facetable',
  'dimensions',
  'vectorSearchProfile'
];

const compareRequiredFieldCapabilities = (existingField, desiredField) => {
  const differences = [];

  fixedFieldAttributes.forEach((attribute) => {
    const desiredValue = desiredField[attribute];
    if (desiredValue === undefined || desiredValue === false) {
      return;
    }

    if (desiredValue === true) {
      if (existingField[attribute] !== true) {
        differences.push(`${attribute}: existing=${existingField[attribute] === undefined ? '(unset)' : existingField[attribute]}, required=true`);
      }
      return;
    }

    if (existingField[attribute] !== desiredValue) {
      differences.push(`${attribute}: existing=${existingField[attribute] === undefined ? '(unset)' : existingField[attribute]}, required=${desiredValue}`);
    }
  });

  return differences;
};

const inspectIndexUpdate = async (env, definition) => {
  try {
    const existing = await getIndex(env, definition.name);
    const existingFields = new Map((existing.fields || []).map((field) => [field.name, field]));
    const desiredFields = new Map((definition.fields || []).map((field) => [field.name, field]));
    const additions = [];
    const blocked = [];
    const mergedFields = [...(existing.fields || [])];

    desiredFields.forEach((desiredField, name) => {
      const existingField = existingFields.get(name);
      if (!existingField) {
        additions.push(name);
        mergedFields.push(desiredField);
        return;
      }

      const differences = compareRequiredFieldCapabilities(existingField, desiredField);
      if (differences.length > 0) {
        blocked.push(`${name} is missing required capabilities (${differences.join('; ')})`);
      }
    });

    const mergedDefinition = {
      ...existing,
      fields: mergedFields
    };

    if (definition.vectorSearch && !mergedDefinition.vectorSearch) {
      mergedDefinition.vectorSearch = definition.vectorSearch;
    }

    if (definition.semantic) {
      mergedDefinition.semantic = definition.semantic;
    }

    return { name: definition.name, exists: true, additions, blocked, definition: mergedDefinition };
  } catch (error) {
    if (error.statusCode === 404 || String(error.message || '').includes('404')) {
      return {
        name: definition.name,
        exists: false,
        additions: (definition.fields || []).map((field) => field.name),
        blocked: [],
        definition
      };
    }

    throw error;
  }
};

const run = async () => {
  const envPath = process.argv[2] || 'config/search.env';
  const dryRun = process.argv.includes('--dry-run');
  const { env, absolutePath } = loadEnv(envPath);
  const definitions = buildIndexDefinitions(env);

  console.log(`Using env: ${absolutePath}`);
  console.log('Preflighting Azure Search index schemas without deleting documents...');

  const inspections = [];
  for (const definition of definitions) {
    inspections.push(await inspectIndexUpdate(env, definition));
  }

  const blocked = inspections.flatMap((inspection) =>
    inspection.blocked.map((message) => `${inspection.name}: ${message}`)
  );

  if (blocked.length > 0) {
    throw new Error([
      'Refusing to update index schemas because the local definitions include changes Azure Search treats as rebuild-only.',
      ...blocked.map((message) => `- ${message}`)
    ].join('\n'));
  }

  inspections.forEach((inspection) => {
    const action = inspection.exists ? 'update' : 'create';
    const additions = inspection.additions.length > 0 ? inspection.additions.join(', ') : '(none)';
    console.log(`${inspection.name}: ${action}; new fields: ${additions}`);
  });

  if (dryRun) {
    console.log('Dry run only; no schema updates applied.');
    return;
  }

  console.log('Applying additive schema updates without allowIndexDowntime...');
  await createOrUpdateIndexes(env, inspections.map((inspection) => inspection.definition), { allowIndexDowntime: false });
  inspections.forEach((inspection) => console.log(`Updated schema: ${inspection.name}`));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
