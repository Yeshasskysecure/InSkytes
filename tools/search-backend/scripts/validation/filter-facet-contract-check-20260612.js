const { loadEnv } = require('../../dist/runtime/env');
const { searchDocumentsPaged } = require('../../dist/api/knowledgeSearchApi');

process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED =
  process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED || 'true';
process.env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED =
  process.env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED || 'true';

const envPath = process.argv[2] || 'config/search.env';
const maxFacetValues = Number(process.argv[3] || 80);

const queries = [
  '*',
  'PRMA & HEOR',
  'PRMA ＆ HEOR',
  'capability',
  'oncology',
  'Generative AI',
  'agentic ai',
  'market access',
  'pharmacovigilance',
  'customer centricity',
  'ServiceNow',
  'risk audit and compliance',
  'no-such-ikn-query-zzzxxy'
];

const usefulFacetValue = (value) => {
  const normalized = String(value || '').trim().toLowerCase();
  return normalized && !['not applicable', 'n/a', 'na', 'none', 'null'].includes(normalized);
};

const search = (env, query, filters = {}) =>
  searchDocumentsPaged(env, {
    query,
    top: 1,
    skip: 0,
    includeFacets: true,
    includeTotalCount: true,
    filters
  }, { log: () => {} });

const compareCount = (expected, actual) => Number(expected) === Number(actual);

const run = async () => {
  const { env } = loadEnv(envPath);
  Object.assign(process.env, env);
  const failures = [];
  const summaries = [];

  for (const query of queries) {
    const base = await search(env, query);
    const baseDepartments = (base.facets.department || [])
      .filter((item) => usefulFacetValue(item.value))
      .slice(0, maxFacetValues);
    const baseBus = (base.facets.bu || [])
      .filter((item) => usefulFacetValue(item.value))
      .slice(0, maxFacetValues);

    for (const department of baseDepartments) {
      const filtered = await search(env, query, { department: [department.value] });
      if (!compareCount(department.count, filtered.totalCount)) {
        failures.push({
          query,
          scope: 'department',
          value: department.value,
          expected: department.count,
          actual: filtered.totalCount
        });
      }
    }

    for (const bu of baseBus) {
      const buFiltered = await search(env, query, { bu: [bu.value] });
      if (!compareCount(bu.count, buFiltered.totalCount)) {
        failures.push({
          query,
          scope: 'businessUnit',
          value: bu.value,
          expected: bu.count,
          actual: buFiltered.totalCount
        });
      }

      const departments = (buFiltered.facets.department || [])
        .filter((item) => usefulFacetValue(item.value))
        .slice(0, maxFacetValues);

      for (const department of departments) {
        const bothFiltered = await search(env, query, {
          bu: [bu.value],
          department: [department.value]
        });
        if (!compareCount(department.count, bothFiltered.totalCount)) {
          failures.push({
            query,
            scope: 'businessUnit+department',
            businessUnit: bu.value,
            department: department.value,
            expected: department.count,
            actual: bothFiltered.totalCount
          });
        }
      }
    }

    summaries.push({
      query,
      totalCount: base.totalCount,
      checkedDepartments: baseDepartments.length,
      checkedBusinessUnits: baseBus.length,
      topDepartments: baseDepartments.slice(0, 12),
      topBusinessUnits: baseBus.slice(0, 12)
    });
  }

  console.log(JSON.stringify({
    normalizedFacetsEnabled: process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED,
    canonicalDepartmentFacetsEnabled: process.env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED,
    maxFacetValues,
    checkedQueries: queries.length,
    failureCount: failures.length,
    failures: failures.slice(0, 50),
    summaries
  }, null, 2));

  if (failures.length > 0) {
    process.exitCode = 1;
  }
};

run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
