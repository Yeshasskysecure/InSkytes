const { spawnSync } = require('child_process');
const path = require('path');
const { loadEnv, safeSearchKey, toNumber } = require('./lib/env');
const {
  deleteDriveItem,
  fetchListItemByDriveItemId,
  getGraphToken,
  updateDriveItem,
  updateDriveItemContent,
  updateListItemFields,
  uploadDriveItemContentByPath
} = require('./lib/graphClient');
const { searchIndex } = require('./lib/azureSearchClient');
const { withRetry } = require('./lib/retry');

const envPath = process.argv[2] || 'config/search.dev.env';
const confirmed = process.argv.includes('--yes');
const keepFile = process.argv.includes('--keep-file');

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const odataEscape = (value) => String(value || '').replace(/'/g, "''");

const runDelta = (mode, label, extraEnv) => {
  console.log(`\n=== Delta ${label} ===`);
  const result = spawnSync(
    process.execPath,
    ['tools/search-backend/scripts/delta-sync.js', envPath, mode, 'all'],
    {
      cwd: process.cwd(),
      env: { ...process.env, ...extraEnv },
      encoding: 'utf8',
      maxBuffer: 20 * 1024 * 1024
    }
  );

  if (result.stdout) process.stdout.write(result.stdout);
  if (result.stderr) process.stderr.write(result.stderr);
  if (result.status !== 0) {
    throw new Error(`Delta ${label} failed with exit code ${result.status}`);
  }
};

const findDocumentByFileName = async (env, fileName) => {
  const result = await searchIndex(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, {
    search: '*',
    filter: `fileName eq '${odataEscape(fileName)}'`,
    select: 'id,title,fileName,status,contentPreview,lastModifiedDateTime,indexedAt',
    top: 5
  });
  return (result.value || [])[0];
};

const countChunks = async (env, documentId) => {
  const result = await searchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
    search: '*',
    filter: `documentId eq '${odataEscape(documentId)}'`,
    select: 'id,documentId,title,fileName,chunkText,indexedAt',
    top: 1000
  });
  return result.value || [];
};

const waitFor = async (label, assertion, timeoutMs, intervalMs) => {
  const startedAt = Date.now();
  let lastError;
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const value = await assertion();
      console.log(`PASS: ${label}`);
      return value;
    } catch (error) {
      lastError = error;
      await sleep(intervalMs);
    }
  }
  throw new Error(`Timed out waiting for ${label}: ${lastError ? lastError.message : 'no details'}`);
};

const assertIndexed = async (env, fileName, expectedText) => waitFor(
  `${fileName} indexed${expectedText ? ` with text "${expectedText}"` : ''}`,
  async () => {
    const document = await findDocumentByFileName(env, fileName);
    if (!document) throw new Error('document not found');
    if (document.status !== 'Active') throw new Error(`document status is ${document.status}`);
    const chunks = await countChunks(env, document.id);
    if (chunks.length === 0) throw new Error('chunks not found');
    if (expectedText) {
      const combinedText = `${document.contentPreview || ''}\n${chunks.map((chunk) => chunk.chunkText || '').join('\n')}`.toLowerCase();
      if (!combinedText.includes(expectedText.toLowerCase())) {
        throw new Error('expected text not found in document/chunks');
      }
    }
    return { document, chunks };
  },
  120000,
  5000
);

const assertNotIndexed = async (env, fileName) => waitFor(
  `${fileName} removed from documents index`,
  async () => {
    const document = await findDocumentByFileName(env, fileName);
    if (document) throw new Error(`still indexed as ${document.id}`);
    return true;
  },
  120000,
  5000
);

const updateFieldsBestEffort = async (env, token, listItemId, fields, label) => {
  try {
    await withRetry(
      () => updateListItemFields(env, token, listItemId, fields),
      { retries: 4, baseDelayMs: 1500, label }
    );
    return true;
  } catch (error) {
    console.warn(`WARN: ${label} failed: ${error.message}`);
    return false;
  }
};

const run = async () => {
  if (!confirmed) {
    throw new Error([
      'This script mutates the configured SharePoint library and Azure Search dev indexes.',
      'Run again with --yes only against dev env, for example:',
      'npm run search:crud:test:dev -- --yes'
    ].join('\n'));
  }

  const { env, absolutePath } = loadEnv(envPath);
  const hostname = String(env.SHAREPOINT_HOSTNAME || '').toLowerCase();
  if (!hostname.includes('development')) {
    throw new Error(`Refusing to run destructive CRUD test against non-dev SharePoint host: ${env.SHAREPOINT_HOSTNAME || '(missing)'}`);
  }

  const token = await getGraphToken(env);
  const runId = new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14);
  const originalFileName = `crud-sync-test-${runId}.txt`;
  const renamedFileName = `crud-sync-test-${runId}-renamed.txt`;
  const markerV1 = `CRUD_SYNC_TEST_${runId}_VERSION_1`;
  const markerV2 = `CRUD_SYNC_TEST_${runId}_VERSION_2`;
  const activeStatus = env.CRUD_TEST_ACTIVE_STATUS || 'Active';
  const inactiveStatus = env.CRUD_TEST_INACTIVE_STATUS || 'Inactive';
  const inventoryPath = path.resolve('tools', 'search-backend', '.inventory', `crud-sync-test-${safeSearchKey(runId)}.json`);
  const childEnv = {
    SEARCH_INVENTORY_PATH: inventoryPath,
    SYNC_DELTA_SOURCE_KEY: `crud-sync-test-${safeSearchKey(runId)}`
  };

  let driveItemId = '';
  let currentFileName = originalFileName;

  console.log(`Using env: ${absolutePath}`);
  console.log(`Using isolated CRUD test inventory: ${inventoryPath}`);
  console.log(`Test file: ${originalFileName}`);

  try {
    runDelta('checkpoint', 'checkpoint before test file creation', childEnv);

    console.log('\n=== Create file ===');
    const created = await withRetry(
      () => uploadDriveItemContentByPath(env, token, originalFileName, [
        markerV1,
        'This is a controlled dev-only CRUD sync test document.',
        'It exists to verify SharePoint create, update, rename, status change, and delete propagation into Azure AI Search.',
        'The content is intentionally longer than fifty characters so local text extraction creates chunks.'
      ].join('\n')),
      { retries: 4, baseDelayMs: 1500, label: 'upload CRUD test file' }
    );
    driveItemId = created.id;

    const createdListItem = await withRetry(
      () => fetchListItemByDriveItemId(env, token, driveItemId),
      { retries: 6, baseDelayMs: 1500, label: 'fetch created CRUD test list item' }
    );
    await updateFieldsBestEffort(env, token, createdListItem.id, {
      Title: `CRUD Sync Test ${runId}`,
      Status: activeStatus
    }, 'set created CRUD test metadata active');

    runDelta('delta', 'after create', childEnv);
    const createdIndexed = await assertIndexed(env, originalFileName, markerV1);

    console.log('\n=== Update content ===');
    await withRetry(
      () => updateDriveItemContent(env, token, driveItemId, [
        markerV2,
        'This is the updated version of the controlled dev-only CRUD sync test document.',
        'The Azure Search chunks should contain VERSION_2 after delta catch-up.',
        'The old text should no longer be the only indexed representation.'
      ].join('\n')),
      { retries: 4, baseDelayMs: 1500, label: 'update CRUD test file content' }
    );
    runDelta('delta', 'after content update', childEnv);
    await assertIndexed(env, originalFileName, markerV2);

    console.log('\n=== Update metadata/title ===');
    const currentListItem = await withRetry(
      () => fetchListItemByDriveItemId(env, token, driveItemId),
      { retries: 6, baseDelayMs: 1500, label: 'fetch CRUD test list item before metadata update' }
    );
    await updateFieldsBestEffort(env, token, currentListItem.id, {
      Title: `CRUD Sync Test ${runId} Updated Title`,
      Status: activeStatus
    }, 'update CRUD test metadata title');
    runDelta('delta', 'after metadata update', childEnv);
    await waitFor('metadata title updated in documents index', async () => {
      const document = await findDocumentByFileName(env, currentFileName);
      if (!document) throw new Error('document not found');
      if (!String(document.title || '').includes('Updated Title')) throw new Error(`title is "${document.title}"`);
      return document;
    }, 120000, 5000);

    console.log('\n=== Rename file ===');
    await withRetry(
      () => updateDriveItem(env, token, driveItemId, { name: renamedFileName }),
      { retries: 4, baseDelayMs: 1500, label: 'rename CRUD test file' }
    );
    currentFileName = renamedFileName;
    runDelta('delta', 'after rename', childEnv);
    await assertIndexed(env, renamedFileName, markerV2);
    await assertNotIndexed(env, originalFileName);

    console.log('\n=== Active to inactive/non-active ===');
    const renamedListItem = await withRetry(
      () => fetchListItemByDriveItemId(env, token, driveItemId),
      { retries: 6, baseDelayMs: 1500, label: 'fetch renamed CRUD test list item' }
    );
    const inactiveApplied = await updateFieldsBestEffort(env, token, renamedListItem.id, {
      Status: inactiveStatus
    }, `set CRUD test status ${inactiveStatus}`);
    if (inactiveApplied) {
      runDelta('delta', 'after status non-active', childEnv);
      await assertNotIndexed(env, renamedFileName);

      console.log('\n=== Reactivate ===');
      await updateFieldsBestEffort(env, token, renamedListItem.id, {
        Status: activeStatus
      }, `set CRUD test status ${activeStatus}`);
      runDelta('delta', 'after reactivation', childEnv);
      await assertIndexed(env, renamedFileName, markerV2);
    } else {
      console.warn('WARN: Non-active status transition was skipped because SharePoint rejected the configured inactive status value.');
    }

    if (!keepFile) {
      console.log('\n=== Delete file ===');
      await withRetry(
        () => deleteDriveItem(env, token, driveItemId),
        { retries: 4, baseDelayMs: 1500, label: 'delete CRUD test file' }
      );
      driveItemId = '';
      runDelta('delta', 'after delete', childEnv);
      await assertNotIndexed(env, renamedFileName);
    }

    console.log(JSON.stringify({
      ok: true,
      env: absolutePath,
      originalFileName,
      renamedFileName,
      createdDocumentId: createdIndexed.document.id,
      createdChunkCount: createdIndexed.chunks.length,
      inactiveStatusTested: inactiveApplied,
      deletedAtEnd: !keepFile,
      inventoryPath
    }, null, 2));
  } finally {
    if (driveItemId && !keepFile) {
      try {
        await deleteDriveItem(env, token, driveItemId);
      } catch {
        // Best-effort cleanup only; the main delete path above is asserted.
      }
    }
  }
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
