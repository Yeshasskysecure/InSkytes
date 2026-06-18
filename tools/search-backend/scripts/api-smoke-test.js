const { loadEnv } = require('../dist/runtime/env');
const {
  chat,
  searchDocuments,
  searchPeople,
  semanticSearch
} = require('../dist/api/knowledgeSearchApi');

const envPath = process.argv[2] || 'config/search.env';

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  const [documents, semantic, people, chatResult] = await Promise.all([
    searchDocuments(env, { query: '*', top: 3 }),
    semanticSearch(env, { query: 'CMMI Appraisal Results Record Indegene Limited', top: 3 }),
    searchPeople(env, { query: 'Omnichannel Activation', top: 3 }),
    chat(env, { question: 'Who leads Omnichannel Activation?' })
  ]);

  console.log(`Using env: ${absolutePath}`);
  console.log(JSON.stringify({
    documents: documents.map((item) => ({ title: item.title, status: item.status })),
    semantic: semantic.map((item) => ({ title: item.title, status: item.status, rerankerScore: item.rerankerScore })),
    people: people.map((item) => ({ personName: item.personName, score: item['@search.score'] })),
    chat: {
      answerPreview: chatResult.answer.slice(0, 300),
      citationCount: chatResult.citations.length
    }
  }, null, 2));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
