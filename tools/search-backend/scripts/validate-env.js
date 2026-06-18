const fs = require('fs');
const path = require('path');

const envPath = process.argv[2] || path.join('config', 'search.env');
const absolutePath = path.resolve(envPath);

const requiredKeys = [
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
  'AZURE_SEARCH_VECTOR_FIELD',
  'AZURE_SEARCH_VECTOR_DIMENSIONS',
  'AZURE_OPENAI_ENDPOINT',
  'AZURE_OPENAI_API_KEY',
  'AZURE_OPENAI_API_VERSION',
  'AZURE_OPENAI_CHAT_DEPLOYMENT',
  'AZURE_OPENAI_EMBEDDING_DEPLOYMENT',
  'AZURE_OPENAI_EMBEDDING_DIMENSIONS'
];

const optionalPhaseTwoKeys = [
  'AZURE_DOC_INTEL_ENDPOINT',
  'AZURE_DOC_INTEL_KEY',
  'AZURE_SERVICEBUS_CONNECTION_STRING',
  'DATABASE_URL',
  'WEBHOOK_BASE_URL',
  'GRAPH_WEBHOOK_CLIENT_STATE'
];

const parseEnv = (content) => {
  const values = {};
  content.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) {
      return;
    }

    const equalsIndex = trimmed.indexOf('=');
    if (equalsIndex === -1) {
      return;
    }

    const key = trimmed.slice(0, equalsIndex).trim();
    const value = trimmed.slice(equalsIndex + 1).trim();
    values[key] = value;
  });
  return values;
};

if (!fs.existsSync(absolutePath)) {
  console.error(`Env file not found: ${absolutePath}`);
  process.exit(1);
}

const env = parseEnv(fs.readFileSync(absolutePath, 'utf8'));
const missingRequired = requiredKeys.filter((key) => !env[key]);
const missingOptional = optionalPhaseTwoKeys.filter((key) => !env[key]);
const vectorDimensions = Number(env.AZURE_SEARCH_VECTOR_DIMENSIONS || 0);
const embeddingDimensions = Number(env.AZURE_OPENAI_EMBEDDING_DIMENSIONS || 0);
const issues = [];

if (env.AZURE_SEARCH_ENDPOINT && !/^https:\/\/[^/]+\.search\.windows\.net\/?$/.test(env.AZURE_SEARCH_ENDPOINT)) {
  issues.push('AZURE_SEARCH_ENDPOINT must look like https://<service>.search.windows.net');
}

if (vectorDimensions !== 3072) {
  issues.push('AZURE_SEARCH_VECTOR_DIMENSIONS should be 3072 for text-embedding-3-large.');
}

if (embeddingDimensions !== vectorDimensions) {
  issues.push('AZURE_OPENAI_EMBEDDING_DIMENSIONS must match AZURE_SEARCH_VECTOR_DIMENSIONS.');
}

console.log(`Checked ${absolutePath}`);
console.log(`Required values missing: ${missingRequired.length}`);
missingRequired.forEach((key) => console.log(`- ${key}`));
console.log(`Optional phase 2 values missing: ${missingOptional.length}`);
missingOptional.forEach((key) => console.log(`- ${key}`));

if (issues.length > 0) {
  console.log('Configuration issues:');
  issues.forEach((issue) => console.log(`- ${issue}`));
}

if (missingRequired.length > 0 || issues.length > 0) {
  process.exit(1);
}

console.log('Search environment shape is valid. Secrets were not printed.');
