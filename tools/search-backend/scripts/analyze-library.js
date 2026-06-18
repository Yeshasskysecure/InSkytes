const { loadEnv, toNumber } = require('./lib/env');
const { fetchLibraryItemsPage, getGraphToken } = require('./lib/graphClient');
const { mapDocumentMetadataRecord } = require('./lib/mappers');
const { withRetry } = require('./lib/retry');

const envPath = process.argv[2] || 'config/search.env';
const maxItems = toNumber(process.argv[3], 100000);
const progressEvery = toNumber(process.env.SEARCH_ANALYZE_PROGRESS_EVERY, 1000);

const addCount = (bucket, key) => {
  const normalizedKey = key || '(blank)';
  bucket[normalizedKey] = (bucket[normalizedKey] || 0) + 1;
};

const addMetadataValues = (bucket, value) => {
  const values = String(value || '')
    .split(';')
    .map((item) => item.trim())
    .filter(Boolean);

  if (values.length === 0) {
    addCount(bucket, '(blank)');
    return;
  }

  values.forEach((item) => addCount(bucket, item));
};

const extensionFromFileName = (fileName) => {
  const match = String(fileName || '').match(/\.([^.]+)$/);
  return match ? match[1].toLowerCase() : '(none)';
};

const roughChunkEstimate = (extension, activeCount) => {
  if (['mp4', 'mov', 'avi', 'mkv', 'wmv', 'mp3', 'wav', 'm4a'].includes(extension)) return 0;
  if (extension === 'pdf') return activeCount * 6;
  if (['pptx', 'ppt'].includes(extension)) return activeCount * 8;
  if (['docx', 'doc'].includes(extension)) return activeCount * 5;
  if (['xlsx', 'xlsm', 'xls', 'xlsb'].includes(extension)) return activeCount * 4;
  return activeCount * 2;
};

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  const token = await getGraphToken(env);
  const summary = {
    totalSampled: 0,
    byStatus: {},
    byExtension: {},
    activeByExtension: {},
    metadataCompleteness: {
      title: 0,
      bu: 0,
      department: 0,
      documentType: 0,
      client: 0,
      region: 0,
      diseaseArea: 0,
      therapyArea: 0,
      authors: 0
    },
    activeMetadataValues: {
      bu: {},
      department: {},
      documentType: {},
      client: {},
      region: {},
      therapyArea: {},
      diseaseArea: {}
    },
    estimatedActiveChunks: 0
  };

  let nextLink = '';
  while (summary.totalSampled < maxItems) {
    const page = await withRetry(
      () => fetchLibraryItemsPage(env, token, Math.min(200, maxItems - summary.totalSampled), nextLink || undefined),
      { retries: 6, baseDelayMs: 1500, label: `analyze library page after ${summary.totalSampled}` }
    );
    const items = page.value || [];
    for (const item of items) {
      const record = mapDocumentMetadataRecord(env, item);
      const extension = extensionFromFileName(record.fileName);
      const active = String(record.status || '').toLowerCase() === 'active';

      summary.totalSampled += 1;
      addCount(summary.byStatus, record.status);
      addCount(summary.byExtension, extension);
      if (active) addCount(summary.activeByExtension, extension);
      if (active) {
        addMetadataValues(summary.activeMetadataValues.bu, record.bu);
        addMetadataValues(summary.activeMetadataValues.department, record.department);
        addMetadataValues(summary.activeMetadataValues.documentType, record.documentType);
        addMetadataValues(summary.activeMetadataValues.client, record.client);
        addMetadataValues(summary.activeMetadataValues.region, record.region);
        addMetadataValues(summary.activeMetadataValues.therapyArea, record.therapyArea);
        addMetadataValues(summary.activeMetadataValues.diseaseArea, record.diseaseArea);
      }

      if (record.title) summary.metadataCompleteness.title += 1;
      if (record.bu) summary.metadataCompleteness.bu += 1;
      if (record.department) summary.metadataCompleteness.department += 1;
      if (record.documentType) summary.metadataCompleteness.documentType += 1;
      if (record.client) summary.metadataCompleteness.client += 1;
      if (record.region) summary.metadataCompleteness.region += 1;
      if (record.diseaseArea) summary.metadataCompleteness.diseaseArea += 1;
      if (record.therapyArea) summary.metadataCompleteness.therapyArea += 1;
      if (record.authors.length > 0) summary.metadataCompleteness.authors += 1;
    }

    nextLink = page['@odata.nextLink'];
    if (summary.totalSampled > 0 && summary.totalSampled % progressEvery === 0) {
      console.error(`Analyzed ${summary.totalSampled} items...`);
    }
    if (!nextLink || items.length === 0) break;
  }

  Object.entries(summary.activeByExtension).forEach(([extension, count]) => {
    summary.estimatedActiveChunks += roughChunkEstimate(extension, count);
  });

  console.log(`Using env: ${absolutePath}`);
  console.log(JSON.stringify(summary, null, 2));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
