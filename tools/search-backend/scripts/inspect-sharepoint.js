const { loadEnv } = require('./lib/env');
const { getGraphToken } = require('./lib/graphClient');

const graphBaseUrl = 'https://graph.microsoft.com/v1.0';

const graphGet = async (token, url) => {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: 'application/json'
    }
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph request failed ${response.status}: ${text.slice(0, 700)}`);
  }

  return response.json();
};

const fieldType = (column) => {
  const knownTypes = [
    'choice',
    'boolean',
    'calculated',
    'currency',
    'dateTime',
    'hyperlinkOrPicture',
    'lookup',
    'number',
    'personOrGroup',
    'term',
    'text',
    'thumbnail'
  ];

  return knownTypes.find((type) => column[type]) || 'other';
};

const summarizeLibraryItems = (items) => {
  const status = {};
  const extensions = {};
  const authors = {};

  (items.value || []).forEach((item) => {
    const fields = item.fields || {};
    const statusValue = String(fields.Status || '(blank)');
    status[statusValue] = (status[statusValue] || 0) + 1;

    const fileName = String(fields.FileLeafRef || '');
    const extension = (fileName.split('.').pop() || '(none)').toLowerCase();
    extensions[extension] = (extensions[extension] || 0) + 1;

    const author = String(fields.Author0 || fields.Author || '(blank)');
    authors[author] = (authors[author] || 0) + 1;
  });

  return {
    sampleCount: (items.value || []).length,
    status,
    extensions,
    authorValueExamples: Object.keys(authors).slice(0, 8)
  };
};

const run = async () => {
  const { env, absolutePath } = loadEnv(process.argv[2] || 'config/search.env');
  const token = await getGraphToken(env);
  const siteId = encodeURIComponent(env.SHAREPOINT_SITE_ID);
  const libraryListId = encodeURIComponent(env.SHAREPOINT_LIBRARY_LIST_ID);
  const whosWhoListId = encodeURIComponent(env.SHAREPOINT_WHOSWHO_LIST_ID);

  const [library, libraryColumns, libraryItems, whosWho, whosWhoColumns, whosWhoItems] = await Promise.all([
    graphGet(token, `${graphBaseUrl}/sites/${siteId}/lists/${libraryListId}?$select=id,displayName,webUrl,list`),
    graphGet(token, `${graphBaseUrl}/sites/${siteId}/lists/${libraryListId}/columns?$select=name,displayName,hidden,readOnly,columnGroup,choice,text,personOrGroup,lookup,term,dateTime,number,boolean,hyperlinkOrPicture`),
    graphGet(token, `${graphBaseUrl}/sites/${siteId}/lists/${libraryListId}/items?$expand=fields&$top=50`),
    graphGet(token, `${graphBaseUrl}/sites/${siteId}/lists/${whosWhoListId}?$select=id,displayName,webUrl,list`),
    graphGet(token, `${graphBaseUrl}/sites/${siteId}/lists/${whosWhoListId}/columns?$select=name,displayName,hidden,readOnly,columnGroup,choice,text,personOrGroup,lookup,term,dateTime,number,boolean,hyperlinkOrPicture`),
    graphGet(token, `${graphBaseUrl}/sites/${siteId}/lists/${whosWhoListId}/items?$expand=fields&$top=20`)
  ]);

  const visibleColumns = (columns) => (columns.value || [])
    .filter((column) => !column.hidden)
    .map((column) => ({
      name: column.name,
      displayName: column.displayName,
      type: fieldType(column),
      readOnly: Boolean(column.readOnly)
    }));

  console.log(`Using env: ${absolutePath}`);
  console.log(JSON.stringify({
    library: {
      displayName: library.displayName,
      webUrl: library.webUrl,
      itemCount: library.list && library.list.itemCount,
      sample: summarizeLibraryItems(libraryItems),
      visibleColumns: visibleColumns(libraryColumns)
    },
    whosWho: {
      displayName: whosWho.displayName,
      webUrl: whosWho.webUrl,
      itemCount: whosWho.list && whosWho.list.itemCount,
      sampleCount: (whosWhoItems.value || []).length,
      visibleColumns: visibleColumns(whosWhoColumns)
    }
  }, null, 2));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
