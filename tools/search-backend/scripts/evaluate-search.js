const { loadEnv } = require('./lib/env');
const {
  chat,
  searchDocuments,
  searchPeople,
  semanticSearch
} = require('../dist/api/knowledgeSearchApi');

const defaultCases = [
  {
    name: 'active metadata wildcard',
    type: 'generic',
    query: '*',
    expectAny: true,
    expectedAllStatus: 'Active'
  },
  {
    name: 'exact title search',
    type: 'generic',
    query: 'SOP For Indegene Service Operations',
    expectedTitleIncludes: 'SOP For Indegene Service Operations',
    expectedAllStatus: 'Active'
  },
  {
    name: 'metadata document type filter',
    type: 'generic',
    query: 'CMMI',
    filters: { documentType: 'Certificate' },
    expectedTitleIncludes: 'CMMI',
    expectedAllStatus: 'Active',
    expectedAnyFieldIncludes: { field: 'documentType', value: 'Certificate' }
  },
  {
    name: 'generic no-result query',
    type: 'generic',
    query: 'zzzxxyy918273qwerty',
    expectedNoResults: true
  },
  {
    name: 'semantic certificate sentence',
    type: 'semantic',
    query: 'CMMI Appraisal Results Record Indegene Limited ECS EMS Omnichannel Activation',
    expectedTitleIncludes: 'CMMI',
    expectedAllStatus: 'Active'
  },
  {
    name: 'semantic author query',
    type: 'semantic',
    query: 'Bushra Hameed CMMI certificate',
    expectedTitleIncludes: 'CMMI',
    expectedAllStatus: 'Active'
  },
  {
    name: 'people directory wildcard',
    type: 'people',
    query: '*',
    expectAny: true
  },
  {
    name: 'chat greeting handled without citations',
    type: 'chat',
    query: 'hello',
    expectedAnswerIncludes: 'Ask me for a document',
    expectedCitationCount: 0
  },
  {
    name: 'chat unsupported action handled safely',
    type: 'chat',
    query: 'Can you delete this document for me?',
    expectedAnswerIncludes: 'cannot edit',
    expectedCitationCount: 0
  },
  {
    name: 'chat ambiguous followup asks for subject',
    type: 'chat',
    query: 'what about that?',
    expectedAnswerIncludes: 'one more detail',
    expectedCitationCount: 0
  },
  {
    name: 'chat no evidence handled gracefully',
    type: 'chat',
    query: 'zzzxxyy918273qwerty',
    expectedAnswerIncludes: 'enough evidence',
    expectedCitationCount: 0
  },
  {
    name: 'chat people question grounded without visible citations',
    type: 'chat',
    query: 'Who leads Omnichannel Activation?',
    expectedAnswerIncludes: 'Omnichannel',
    expectedCitationCount: 0
  },
  {
    name: 'chat follow-up uses recent conversation context without visible people citations',
    type: 'chat',
    query: 'what about that team?',
    history: [
      { role: 'user', content: 'Who leads Omnichannel Activation?' },
      { role: 'assistant', content: 'Omnichannel Activation is led by Gurpinder Singh.' }
    ],
    expectedAnswerIncludes: 'Omnichannel',
    expectedCitationCount: 0
  }
];

const runGeneric = async (env, testCase) => {
  return searchDocuments(env, {
    query: testCase.query,
    filters: testCase.filters,
    top: 5
  });
};

const runSemantic = async (env, testCase) => {
  return semanticSearch(env, {
    query: testCase.query,
    filters: testCase.filters,
    top: 5
  });
};

const runPeople = async (env, testCase) => {
  return searchPeople(env, {
    query: testCase.query,
    top: 5
  });
};

const passCase = (testCase, results) => {
  if (testCase.type === 'chat') {
    const answer = String(results.answer || '');
    const answerLower = answer.toLowerCase();
    const citations = results.citations || [];
    if (testCase.expectedAnswerIncludes && !answer.toLowerCase().includes(testCase.expectedAnswerIncludes.toLowerCase())) {
      return false;
    }
    if (Array.isArray(testCase.expectedAllAnswerIncludes)) {
      const allMatched = testCase.expectedAllAnswerIncludes
        .every((value) => answerLower.includes(String(value).toLowerCase()));
      if (!allMatched) return false;
    }
    if (Array.isArray(testCase.forbiddenAnswerIncludes)) {
      const forbiddenMatched = testCase.forbiddenAnswerIncludes
        .some((value) => answerLower.includes(String(value).toLowerCase()));
      if (forbiddenMatched) return false;
    }
    if (typeof testCase.expectedCitationCount === 'number' && citations.length !== testCase.expectedCitationCount) {
      return false;
    }
    if (typeof testCase.minimumCitationCount === 'number' && citations.length < testCase.minimumCitationCount) {
      return false;
    }
    if (testCase.expectedCitationTitleIncludes) {
      const expectedTitles = Array.isArray(testCase.expectedCitationTitleIncludes)
        ? testCase.expectedCitationTitleIncludes
        : [testCase.expectedCitationTitleIncludes];
      const citationTitles = citations.map((citation) => String(citation.title || '').toLowerCase()).join('\n');
      const allTitlesMatched = expectedTitles
        .every((value) => citationTitles.includes(String(value).toLowerCase()));
      if (!allTitlesMatched) return false;
    }
    if (testCase.forbidBracketCitations && /\[(?:\d+|[1-9]\d*\s*[-–]\s*[1-9]\d*)\]/.test(answer)) {
      return false;
    }
    return true;
  }

  const values = Array.isArray(results) ? results : (results.value || []);
  if (testCase.expectedNoResults) return values.length === 0;
  if (testCase.expectedAllStatus) {
    const allStatusMatched = values.every((item) => String(item.status || '') === testCase.expectedAllStatus);
    if (!allStatusMatched) return false;
  }
  if (testCase.expectedAnyFieldIncludes) {
    const { field, value } = testCase.expectedAnyFieldIncludes;
    const matched = values.some((item) => String(item[field] || '').toLowerCase().includes(String(value).toLowerCase()));
    if (!matched) return false;
  }
  if (testCase.expectAny) return values.length > 0;
  if (testCase.expectedTitleIncludes) {
    return values.some((item) => String(item.title || item.fileName || '').toLowerCase().includes(testCase.expectedTitleIncludes.toLowerCase()));
  }
  return false;
};

const loadCases = () => {
  const casesPath = process.argv[3];
  if (!casesPath) {
    return defaultCases;
  }

  const fs = require('fs');
  const path = require('path');
  const absolutePath = path.resolve(casesPath);
  const parsed = JSON.parse(fs.readFileSync(absolutePath, 'utf8'));
  if (!Array.isArray(parsed)) {
    throw new Error(`Evaluation cases file must contain a JSON array: ${absolutePath}`);
  }
  console.log(`Using evaluation cases: ${absolutePath}`);
  return parsed;
};

const run = async () => {
  const { env, absolutePath } = loadEnv(process.argv[2] || 'config/search.env');
  console.log(`Using env: ${absolutePath}`);
  const testCases = loadCases();

  const results = [];
  for (const testCase of testCases) {
    const startedAt = Date.now();
    const searchResults = testCase.type === 'generic'
      ? await runGeneric(env, testCase)
      : testCase.type === 'semantic'
        ? await runSemantic(env, testCase)
        : testCase.type === 'people'
          ? await runPeople(env, testCase)
          : await chat(env, { question: testCase.query, history: testCase.history });
    const durationMs = Date.now() - startedAt;

    results.push({
      name: testCase.name,
      type: testCase.type,
      passed: passCase(testCase, searchResults),
      durationMs,
      top: testCase.type === 'chat'
        ? {
          answerPreview: String(searchResults.answer || '').slice(0, 180),
          citationCount: (searchResults.citations || []).length
        }
        : (Array.isArray(searchResults) ? searchResults : (searchResults.value || [])).slice(0, 5).map((item) => ({
          title: item.title || item.personName || item.fileName,
          score: item['@search.score'],
          rerankerScore: item['@search.rerankerScore'],
          status: item.status,
          documentType: item.documentType,
          authors: item.authors
        }))
    });
  }

  const passed = results.filter((result) => result.passed).length;
  console.log(JSON.stringify({ passed, total: results.length, results }, null, 2));
  if (passed !== results.length) process.exit(1);
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
