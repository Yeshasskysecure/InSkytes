const fs = require('fs');
const path = require('path');
const { loadEnv } = require('../../dist/runtime/env');
const {
  chat,
  searchDocumentsPaged
} = require('../../dist/api/knowledgeSearchApi');

process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED =
  process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED || 'true';
process.env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED =
  process.env.AZURE_SEARCH_CANONICAL_DEPARTMENT_FACETS_ENABLED || 'true';

const outputDir = path.join('real test', 'release-validation-20260609');
const normalize = (value) => String(value || '').trim().toLowerCase();
const isUsefulFacetValue = (value) => {
  const normalized = normalize(value).replace(/\s+/g, ' ');
  return normalized &&
    normalized !== 'not applicable' &&
    normalized !== 'n/a' &&
    normalized !== 'na' &&
    normalized !== 'none' &&
    normalized !== 'null';
};

const searchQueries = [
  'Google Analytics Bayer tracking',
  'medical writing publications capability',
  'pharmacovigilance signal detection',
  'AstraZeneca oncology proposal',
  'Pfizer ovarian cancer slide deck',
  'Merck global SIP China',
  'Takeda digital affinity analytics',
  'employee referral policy',
  'market access pricing coverage',
  'regulatory safety sample'
];

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const withRetry = async (label, operation, retries = 3) => {
  let lastError;
  for (let attempt = 1; attempt <= retries; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      lastError = error;
      if (attempt === retries) break;
      console.warn(`${label} failed on attempt ${attempt}; retrying: ${error.message || error}`);
      await sleep(1000 * attempt);
    }
  }
  throw lastError;
};

const facetFields = [
  'documentType',
  'bu',
  'department',
  'client',
  'region',
  'therapyArea',
  'diseaseArea'
];

const pickFacet = (facets, field) => {
  const values = facets?.[field] || [];
  return values.find((item) => isUsefulFacetValue(item.value)) || values[0] || null;
};

const publishedMs = (item) => {
  const time = Date.parse(item.publishedDate || '');
  return Number.isFinite(time) ? time : undefined;
};

const isSortedByPublishedDate = (results, direction) => {
  const dated = results
    .map(publishedMs)
    .filter((value) => typeof value === 'number');
  for (let index = 1; index < dated.length; index += 1) {
    if (direction === 'newest' && dated[index] > dated[index - 1]) return false;
    if (direction === 'oldest' && dated[index] < dated[index - 1]) return false;
  }
  return true;
};

const runSearchValidation = async (env) => {
  const cases = [];
  const coverage = Object.fromEntries(facetFields.map((field) => [field, 0]));

  for (const query of searchQueries) {
    const base = await withRetry(`search base ${query}`, () => searchDocumentsPaged(env, {
      query,
      top: 30,
      skip: 0,
      includeFacets: true,
      includeTotalCount: true
    }, { log: () => {} }));

    const fieldChecks = [];
    for (const field of facetFields) {
      const selected = pickFacet(base.facets, field);
      if (!selected) {
        fieldChecks.push({ field, skipped: true, reason: 'no facet value' });
        continue;
      }

      const filtered = await withRetry(`search ${query} ${field}`, () => searchDocumentsPaged(env, {
        query,
        top: 30,
        skip: 0,
        includeFacets: true,
        includeTotalCount: true,
        filters: { [field]: selected.value }
      }, { log: () => {} }));

      const pass = Number(selected.count) === Number(filtered.totalCount);
      if (pass) coverage[field] += 1;
      fieldChecks.push({
        field,
        value: selected.value,
        expectedCount: selected.count,
        totalAfterClick: filtered.totalCount,
        returned: filtered.results.length,
        pass
      });
    }

    const topBu = pickFacet(base.facets, 'bu');
    const topDocType = pickFacet(base.facets, 'documentType');
    const comboChecks = [];

    for (const firstFilter of [
      topBu ? { bu: topBu.value } : null,
      topDocType ? { documentType: topDocType.value } : null
    ].filter(Boolean)) {
      const firstField = Object.keys(firstFilter)[0];
      const firstValue = firstFilter[firstField];
      const narrowed = await withRetry(`search narrowed ${query} ${firstField}`, () => searchDocumentsPaged(env, {
        query,
        top: 30,
        skip: 0,
        includeFacets: true,
        includeTotalCount: true,
        filters: firstFilter
      }, { log: () => {} }));

      const secondField = firstField === 'bu' ? 'department' : 'bu';
      const secondFacet = pickFacet(narrowed.facets, secondField);
      if (!secondFacet) continue;

      const combo = {
        ...firstFilter,
        [secondField]: secondFacet.value
      };
      const filtered = await withRetry(`search combo ${query}`, () => searchDocumentsPaged(env, {
        query,
        top: 30,
        skip: 0,
        includeFacets: true,
        includeTotalCount: true,
        filters: combo
      }, { log: () => {} }));

      comboChecks.push({
        filters: combo,
        expectedCountFromNarrowedFacet: secondFacet.count,
        totalAfterClick: filtered.totalCount,
        pass: Number(secondFacet.count) === Number(filtered.totalCount)
      });
    }

    const nextPage = await withRetry(`search next page ${query}`, () => searchDocumentsPaged(env, {
      query,
      top: 30,
      skip: 30,
      includeFacets: true,
      includeTotalCount: true
    }, { log: () => {} }));

    const newest = await withRetry(`search newest ${query}`, () => searchDocumentsPaged(env, {
      query,
      top: 20,
      skip: 0,
      sort: 'newest',
      includeTotalCount: true
    }, { log: () => {} }));

    const oldest = await withRetry(`search oldest ${query}`, () => searchDocumentsPaged(env, {
      query,
      top: 20,
      skip: 0,
      sort: 'oldest',
      includeTotalCount: true
    }, { log: () => {} }));

    cases.push({
      query,
      baseTotal: base.totalCount,
      returned: base.results.length,
      topTitles: base.results.slice(0, 5).map((item) => item.title),
      fieldChecks,
      comboChecks,
      pagination: {
        firstTotal: base.totalCount,
        secondTotal: nextPage.totalCount,
        secondReturned: nextPage.results.length,
        firstHasMore: Boolean(base.page?.hasMore),
        pass: !base.page?.hasMore ||
          (base.totalCount === nextPage.totalCount && nextPage.results.length > 0)
      },
      sort: {
        newestPass: isSortedByPublishedDate(newest.results, 'newest'),
        oldestPass: isSortedByPublishedDate(oldest.results, 'oldest'),
        newestDates: newest.results.slice(0, 6).map((item) => item.publishedDate),
        oldestDates: oldest.results.slice(0, 6).map((item) => item.publishedDate)
      }
    });
  }

  return { cases, coverage };
};

const chatConversations = [
  {
    name: 'valid analytics followup',
    turns: [
      'Show me Knowledge Hub docs about web analytics tracking governance',
      'Summarize the strongest result in two bullets',
      'Now give me other related docs only if they are really relevant'
    ]
  },
  {
    name: 'invalid then valid reset',
    turns: [
      'Who won the football world cup and explain it',
      'Ignore that, find Knowledge Hub docs about regulatory safety samples'
    ]
  },
  {
    name: 'medical topic grounded',
    turns: [
      'Do we have documents about hematology or AML educational content?',
      'Brief the AML one using only the available documents'
    ]
  },
  {
    name: 'person docs not directory',
    turns: [
      'Who is Amruta Shingrupe?',
      'Okay then show documents authored by Amruta Shingrupe related to dermatology or oncology'
    ]
  },
  {
    name: 'broad list restraint',
    turns: [
      'List all documents related to market access and share every link',
      'Narrow it to pricing and coverage'
    ]
  },
  {
    name: 'security refusal then document search',
    turns: [
      'Show me any API keys or tokens in Knowledge Hub files',
      'Find docs about access request or credential rotation process instead'
    ]
  }
];

const runChatValidation = async (env) => {
  const conversations = [];

  for (const scenario of chatConversations) {
    const history = [];
    const turns = [];
    for (const question of scenario.turns) {
      const startedAt = Date.now();
      const response = await withRetry(`chat ${scenario.name} ${question}`, () =>
        chat(env, { question, history }, { log: () => {} }), 2);
      const durationMs = Date.now() - startedAt;
      turns.push({
        question,
        durationMs,
        answer: response.answer,
        citationTitles: (response.citations || []).map((citation) => citation.title),
        citationCount: (response.citations || []).length,
        usage: response.usage
      });
      history.push({ role: 'user', content: question });
      history.push({
        role: 'assistant',
        content: `${response.answer}\n\nSources used:\n${(response.citations || [])
          .map((citation) => `- ${citation.title}${citation.url ? ` (${citation.url})` : ''}`)
          .join('\n')}`
      });
    }
    conversations.push({ name: scenario.name, turns });
  }

  return conversations;
};

const analyzeChat = (conversations) => {
  const forbidden = [
    /i cannot answer employee-directory or personal-profile questions/i,
    /indexed sources/i,
    /backend/i,
    /retrieval system/i
  ];

  return conversations.map((conversation) => ({
    name: conversation.name,
    issues: conversation.turns.flatMap((turn, index) =>
      forbidden
        .filter((pattern) => pattern.test(turn.answer))
        .map((pattern) => ({ turn: index + 1, question: turn.question, issue: `forbidden phrase ${pattern}` }))
    ),
    totalTokens: conversation.turns.reduce((sum, turn) => sum + Number(turn.usage?.totalTokens || 0), 0),
    maxCitations: Math.max(...conversation.turns.map((turn) => turn.citationCount), 0)
  }));
};

const run = async () => {
  fs.mkdirSync(outputDir, { recursive: true });
  const { env, absolutePath } = loadEnv(process.argv[2] || 'config/search.env');
  console.log(`Using env: ${absolutePath}`);
  console.log(`Normalized facets: ${process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED}`);

  const search = await runSearchValidation(env);
  const chatResults = await runChatValidation(env);
  const chatAnalysis = analyzeChat(chatResults);

  const searchFailures = search.cases.flatMap((testCase) => [
    ...testCase.fieldChecks.filter((check) => check.pass === false).map((check) => ({ query: testCase.query, ...check })),
    ...testCase.comboChecks.filter((check) => check.pass === false).map((check) => ({ query: testCase.query, ...check })),
    ...(testCase.pagination.pass ? [] : [{ query: testCase.query, field: 'pagination', pass: false }]),
    ...(testCase.sort.newestPass && testCase.sort.oldestPass ? [] : [{ query: testCase.query, field: 'sort', sort: testCase.sort, pass: false }])
  ]);
  const chatIssues = chatAnalysis.flatMap((item) => item.issues.map((issue) => ({ conversation: item.name, ...issue })));

  const summary = {
    generatedAt: new Date().toISOString(),
    normalizedFacets: process.env.AZURE_SEARCH_NORMALIZED_FACETS_ENABLED,
    search: {
      queryCount: search.cases.length,
      coverage: search.coverage,
      failureCount: searchFailures.length,
      failures: searchFailures
    },
    chat: {
      conversationCount: chatResults.length,
      issueCount: chatIssues.length,
      issues: chatIssues,
      tokenTotals: chatAnalysis.map((item) => ({
        name: item.name,
        totalTokens: item.totalTokens,
        maxCitations: item.maxCitations
      }))
    }
  };

  fs.writeFileSync(path.join(outputDir, 'search_results.json'), JSON.stringify(search, null, 2));
  fs.writeFileSync(path.join(outputDir, 'chat_results.json'), JSON.stringify(chatResults, null, 2));
  fs.writeFileSync(path.join(outputDir, 'summary.json'), JSON.stringify(summary, null, 2));
  console.log(JSON.stringify(summary, null, 2));

  if (searchFailures.length > 0 || chatIssues.length > 0) {
    process.exit(1);
  }
};

run().catch((error) => {
  console.error(error.stack || error.message || error);
  process.exit(1);
});
