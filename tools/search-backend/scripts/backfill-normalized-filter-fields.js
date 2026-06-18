const { loadEnv, requireValue, toNumber } = require('./lib/env');
const { searchIndex, uploadDocuments } = require('./lib/azureSearchClient');
const { canonicalDepartmentParts } = require('./lib/departmentTaxonomy');

const envPath = process.argv[2] || 'config/search.env';
const dryRun = process.argv.includes('--dry-run');
const pageSize = Math.min(Math.max(toNumber(process.env.NORMALIZED_FILTER_BACKFILL_PAGE_SIZE, 1000), 1), 1000);

const placeholderMetadataValues = new Set([
  'n.a.',
  'na',
  'not-applicable',
  'not_applicable',
  'none',
  '(none)',
  'null',
  '-'
]);

const uniqueValues = (values) =>
  values.filter((value, index, entries) => Boolean(value) && entries.indexOf(value) === index);

const cleanMetadataToken = (value) => {
  const withoutTaxonomyId = String(value || '')
    .replace(/^-?\d+#/g, '')
    .split('|')[0]
    .normalize('NFKC')
    .replace(/＆/g, '&')
    .replace(/\s*&\s*/g, ' & ')
    .replace(/\s+/g, ' ')
    .trim();
  const normalized = withoutTaxonomyId.toLowerCase().replace(/[\s_-]+/g, ' ');
  return normalized === 'not applicable' || normalized === 'n/a' ? 'Not Applicable' : withoutTaxonomyId;
};

const comparableMetadataToken = (value) =>
  cleanMetadataToken(value)
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[\s_-]+/g, ' ')
    .trim();

const metadataFilterValues = (metadataValue) =>
  uniqueValues(
    String(metadataValue || '')
      .replace(/;#/g, ';')
      .split(';')
      .map((item) => cleanMetadataToken(item.trim()))
      .filter((item) => {
        if (!item) return false;
        const normalized = item.toLowerCase().replace(/\s+/g, ' ');
        return !placeholderMetadataValues.has(normalized);
      })
  );

const splitDepartmentPath = (value) =>
  String(value || '')
    .normalize('NFKC')
    .split(/\s*(?::|>|›|＞)\s*/g)
    .map((part) => cleanMetadataToken(part))
    .filter(Boolean);

const departmentHierarchyValues = (metadataValue, businessUnit, env) => {
  const values = metadataFilterValues(metadataValue);
  const buValues = metadataFilterValues(businessUnit);
  const buComparableValues = new Set(buValues.map(comparableMetadataToken));
  const expanded = [];

  values.forEach((value) => {
    const parts = splitDepartmentPath(value);
    if (parts.length === 0) return;

    const firstPartComparable = comparableMetadataToken(parts[0]);
    const shouldStripBusinessUnitPrefix = parts.length > 1 &&
      buComparableValues.has(firstPartComparable);
    const rawDepartmentParts = shouldStripBusinessUnitPrefix
      ? parts.slice(1)
      : parts;
    const departmentParts = canonicalDepartmentParts(env, rawDepartmentParts, buValues);

    if (departmentParts.length === 0) return;

    for (let index = 1; index <= departmentParts.length; index += 1) {
      expanded.push(departmentParts.slice(0, index).join(':'));
    }
  });

  return uniqueValues(expanded);
};

const departmentFilterValues = (metadataValue, businessUnit, env) => {
  const filterValues = departmentHierarchyValues(metadataValue, businessUnit, env);
  const legacyPathValues = metadataFilterValues(metadataValue)
    .map(splitDepartmentPath)
    .filter((parts) => parts.length > 1)
    .map((parts) => parts.join(':'));

  return uniqueValues([...filterValues, ...legacyPathValues]);
};

const departmentFacetValues = (metadataValue, businessUnit, env) =>
  departmentHierarchyValues(metadataValue, businessUnit, env);

const normalizedFilterPatch = (record, env) => ({
  id: record.id,
  buFilterValues: metadataFilterValues(record.bu),
  departmentFilterValues: departmentFilterValues(record.department, record.bu, env),
  departmentFacetValues: departmentFacetValues(record.department, record.bu, env),
  diseaseAreaFilterValues: metadataFilterValues(record.diseaseArea),
  therapyAreaFilterValues: metadataFilterValues(record.therapyArea),
  clientFilterValues: metadataFilterValues(record.client),
  regionFilterValues: metadataFilterValues(record.region),
  documentTypeFilterValues: metadataFilterValues(record.documentType)
});

const backfillIndex = async (env, indexName) => {
  const select = 'id,bu,department,diseaseArea,therapyArea,client,region,documentType';
  let skip = 0;
  const records = [];
  let uploaded = 0;

  for (;;) {
    const response = await searchIndex(env, indexName, {
      search: '*',
      select,
      top: pageSize,
      skip,
      count: true
    });

    const items = response.value || [];
    if (items.length === 0) break;

    records.push(...items);

    const total = Number(response['@odata.count'] || 0);
    console.log(`${indexName}: scanned ${records.length}${total ? `/${total}` : ''}`);

    if (items.length < pageSize) break;
    skip += items.length;
  }

  const patches = records.map((record) => normalizedFilterPatch(record, env));
  console.log(`${indexName}: prepared ${patches.length} normalized filter patches`);

  if (!dryRun) {
    for (let index = 0; index < patches.length; index += pageSize) {
      const batch = patches.slice(index, index + pageSize);
      await uploadDocuments(env, indexName, batch);
      uploaded += batch.length;
      console.log(`${indexName}: uploaded ${uploaded}/${patches.length}`);
    }
  }

  return { indexName, processed: patches.length, uploaded: dryRun ? 0 : uploaded };
};

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  const documentsIndex = requireValue(env, 'AZURE_SEARCH_DOCUMENTS_INDEX');
  const chunksIndex = requireValue(env, 'AZURE_SEARCH_CHUNKS_INDEX');

  console.log(`Using env: ${absolutePath}`);
  console.log(`Mode: ${dryRun ? 'dry-run' : 'apply'}`);
  console.log('Backfilling normalized filter fields from existing Azure Search metadata only.');

  const results = [];
  results.push(await backfillIndex(env, documentsIndex));
  results.push(await backfillIndex(env, chunksIndex));

  console.log(JSON.stringify({ dryRun, results }, null, 2));
};

run().catch((error) => {
  console.error(error.message || error);
  process.exit(1);
});
