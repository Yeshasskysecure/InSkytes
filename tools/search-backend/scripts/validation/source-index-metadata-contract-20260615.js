const fs = require('fs');
const path = require('path');
const { loadEnv, requireValue, toNumber } = require('../lib/env');
const { searchIndex } = require('../lib/azureSearchClient');
const { fetchLibraryItemsPage, getGraphToken } = require('../lib/graphClient');
const { mapDocumentMetadataRecord } = require('../lib/mappers');
const { withRetry } = require('../lib/retry');

const envPath = process.argv[2] || 'config/search.env';
const pageSize = Math.min(Math.max(toNumber(process.env.SOURCE_INDEX_AUDIT_PAGE_SIZE, 200), 1), 1000);
const sampleLimit = Math.min(Math.max(toNumber(process.env.SOURCE_INDEX_AUDIT_SAMPLE_LIMIT, 25), 1), 100);

const fieldContracts = [
  { name: 'bu', scalar: 'bu', collections: ['buFilterValues'] },
  { name: 'department', scalar: 'department', collections: ['departmentFilterValues', 'departmentFacetValues'] },
  { name: 'documentType', scalar: 'documentType', collections: ['documentTypeFilterValues'] },
  { name: 'client', scalar: 'client', collections: ['clientFilterValues'] },
  { name: 'region', scalar: 'region', collections: ['regionFilterValues'] },
  { name: 'therapyArea', scalar: 'therapyArea', collections: ['therapyAreaFilterValues'] },
  { name: 'diseaseArea', scalar: 'diseaseArea', collections: ['diseaseAreaFilterValues'] }
];

const activeStatus = (value) => String(value || '').trim().toLowerCase() === 'active';

const sortedArray = (value) => Array.isArray(value)
  ? value.map((item) => String(item || '')).filter(Boolean).sort()
  : [];

const arraysEqual = (left, right) => {
  const normalizedLeft = sortedArray(left);
  const normalizedRight = sortedArray(right);
  return normalizedLeft.length === normalizedRight.length &&
    normalizedLeft.every((value, index) => value === normalizedRight[index]);
};

const countValues = (records, fieldName) => {
  const counts = new Map();
  for (const record of records) {
    const values = Array.isArray(record[fieldName])
      ? Array.from(new Set(record[fieldName].filter(Boolean)))
      : (record[fieldName] ? [record[fieldName]] : []);
    for (const value of values) {
      counts.set(value, (counts.get(value) || 0) + 1);
    }
  }
  return counts;
};

const topCounts = (counts, limit = 20) =>
  Array.from(counts.entries())
    .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0]))
    .slice(0, limit)
    .map(([value, count]) => ({ value, count }));

const diffCounts = (sourceCounts, indexCounts) => {
  const values = new Set([...sourceCounts.keys(), ...indexCounts.keys()]);
  return Array.from(values)
    .map((value) => ({
      value,
      source: sourceCounts.get(value) || 0,
      index: indexCounts.get(value) || 0,
      delta: (indexCounts.get(value) || 0) - (sourceCounts.get(value) || 0)
    }))
    .filter((entry) => entry.delta !== 0)
    .sort((left, right) => Math.abs(right.delta) - Math.abs(left.delta) || left.value.localeCompare(right.value));
};

const fetchSharePointSource = async (env) => {
  const token = await getGraphToken(env);
  const recordsById = new Map();
  let scanned = 0;
  let active = 0;
  let nextLink;

  for (;;) {
    const page = await withRetry(
      () => fetchLibraryItemsPage(env, token, pageSize, nextLink),
      { retries: 8, baseDelayMs: 3000, maxDelayMs: 60000, label: `read SharePoint page after ${scanned}` }
    );
    const items = page.value || [];

    for (const item of items) {
      scanned += 1;
      const record = mapDocumentMetadataRecord(env, item);
      if (!activeStatus(record.status)) continue;
      active += 1;
      recordsById.set(record.id, record);
    }

    console.log(`SharePoint source scanned ${scanned}; active mapped ${active}`);
    nextLink = page['@odata.nextLink'];
    if (!nextLink) break;
  }

  return { scanned, active, recordsById };
};

const fetchIndexDocuments = async (env, indexName) => {
  const selectFields = [
    'id',
    'title',
    'fileName',
    'status',
    ...fieldContracts.flatMap((contract) => [contract.scalar, ...contract.collections])
  ];
  const records = [];
  let skip = 0;

  for (;;) {
    const response = await withRetry(
      () => searchIndex(env, indexName, {
        search: '*',
        filter: "status eq 'Active'",
        select: selectFields.join(','),
        top: pageSize,
        skip,
        count: true
      }),
      { retries: 5, baseDelayMs: 1500, maxDelayMs: 30000, label: `read ${indexName} at skip ${skip}` }
    );
    const items = response.value || [];
    records.push(...items);
    const total = Number(response['@odata.count'] || 0);
    console.log(`${indexName} scanned ${records.length}${total ? `/${total}` : ''}`);
    if (items.length < pageSize) break;
    skip += items.length;
  }

  return records;
};

const compareRecords = (sourceById, indexRecords) => {
  const mismatchStats = {};
  const mismatchSamples = [];
  const indexIds = new Set(indexRecords.map((record) => record.id));
  const missingInSource = [];
  const missingInIndex = [];

  for (const contract of fieldContracts) {
    mismatchStats[contract.name] = {
      scalarMismatches: 0,
      collectionMismatches: Object.fromEntries(contract.collections.map((field) => [field, 0]))
    };
  }

  for (const indexRecord of indexRecords) {
    const sourceRecord = sourceById.get(indexRecord.id);
    if (!sourceRecord) {
      if (missingInSource.length < sampleLimit) {
        missingInSource.push({ id: indexRecord.id, title: indexRecord.title || indexRecord.fileName || '' });
      }
      continue;
    }

    for (const contract of fieldContracts) {
      const sourceScalar = String(sourceRecord[contract.scalar] || '');
      const indexScalar = String(indexRecord[contract.scalar] || '');
      if (sourceScalar !== indexScalar) {
        mismatchStats[contract.name].scalarMismatches += 1;
        if (mismatchSamples.length < sampleLimit) {
          mismatchSamples.push({
            id: indexRecord.id,
            title: indexRecord.title || indexRecord.fileName || '',
            field: contract.scalar,
            source: sourceScalar,
            index: indexScalar
          });
        }
      }

      for (const collectionField of contract.collections) {
        if (arraysEqual(sourceRecord[collectionField], indexRecord[collectionField])) continue;
        mismatchStats[contract.name].collectionMismatches[collectionField] += 1;
        if (mismatchSamples.length < sampleLimit) {
          mismatchSamples.push({
            id: indexRecord.id,
            title: indexRecord.title || indexRecord.fileName || '',
            field: collectionField,
            source: sortedArray(sourceRecord[collectionField]),
            index: sortedArray(indexRecord[collectionField])
          });
        }
      }
    }
  }

  for (const [id, sourceRecord] of sourceById.entries()) {
    if (!indexIds.has(id) && missingInIndex.length < sampleLimit) {
      missingInIndex.push({ id, title: sourceRecord.title || sourceRecord.fileName || '' });
    }
  }

  return { mismatchStats, mismatchSamples, missingInSource, missingInIndex };
};

const buildFacetComparisons = (sourceRecords, indexRecords) => {
  const comparisons = {};
  for (const contract of fieldContracts) {
    for (const fieldName of contract.collections) {
      const sourceCounts = countValues(sourceRecords, fieldName);
      const indexCounts = countValues(indexRecords, fieldName);
      const diffs = diffCounts(sourceCounts, indexCounts);
      comparisons[fieldName] = {
        sourceDistinctValues: sourceCounts.size,
        indexDistinctValues: indexCounts.size,
        diffCount: diffs.length,
        topSource: topCounts(sourceCounts),
        topIndex: topCounts(indexCounts),
        largestDiffs: diffs.slice(0, 25)
      };
    }
  }
  return comparisons;
};

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  const documentsIndex = requireValue(env, 'AZURE_SEARCH_DOCUMENTS_INDEX');

  console.log(`Using env: ${absolutePath}`);
  console.log('Auditing SharePoint source metadata against Azure Search document index. Read-only.');

  const source = await fetchSharePointSource(env);
  const indexRecords = await fetchIndexDocuments(env, documentsIndex);
  const sourceRecords = Array.from(source.recordsById.values()).filter((record) =>
    indexRecords.some((indexRecord) => indexRecord.id === record.id)
  );
  const comparison = compareRecords(source.recordsById, indexRecords);
  const facetComparisons = buildFacetComparisons(sourceRecords, indexRecords);

  const summary = {
    sourceItemsScanned: source.scanned,
    sourceActiveMapped: source.active,
    indexedActiveDocuments: indexRecords.length,
    indexedDocumentsMissingFromSharePoint: comparison.missingInSource.length,
    sourceActiveMissingFromIndexSampleCount: comparison.missingInIndex.length,
    mismatchStats: comparison.mismatchStats,
    mismatchSampleCount: comparison.mismatchSamples.length,
    facetComparisons,
    samples: {
      metadataMismatches: comparison.mismatchSamples,
      indexedDocumentsMissingFromSharePoint: comparison.missingInSource,
      sourceActiveMissingFromIndex: comparison.missingInIndex
    }
  };

  const outputDir = path.join('tools', 'search-backend', '.inventory');
  fs.mkdirSync(outputDir, { recursive: true });
  const jsonPath = path.join(outputDir, 'source-index-metadata-contract-20260615.json');
  fs.writeFileSync(jsonPath, `${JSON.stringify(summary, null, 2)}\n`);

  console.log(JSON.stringify({
    jsonPath,
    sourceItemsScanned: summary.sourceItemsScanned,
    sourceActiveMapped: summary.sourceActiveMapped,
    indexedActiveDocuments: summary.indexedActiveDocuments,
    mismatchStats: summary.mismatchStats,
    facetDiffCounts: Object.fromEntries(
      Object.entries(summary.facetComparisons).map(([key, value]) => [key, value.diffCount])
    ),
    samples: summary.samples
  }, null, 2));
};

run().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
