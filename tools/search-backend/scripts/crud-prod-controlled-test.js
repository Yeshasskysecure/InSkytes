const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const { loadEnv, safeSearchKey } = require('./lib/env');
const {
  deleteDriveItem,
  fetchListItemByDriveItemId,
  getGraphToken,
  updateDriveItemContent,
  updateListItemFields,
  uploadDriveItemContentByPath
} = require('./lib/graphClient');
const { searchIndex } = require('./lib/azureSearchClient');
const { withRetry } = require('./lib/retry');

const envPath = process.argv[2] || 'config/search.env';
const confirmed = process.argv.includes('--yes') && process.argv.includes('--prod-controlled');
const keepFiles = process.argv.includes('--keep-files');
const skipInactive = process.argv.includes('--skip-inactive');
const manualPrepare = process.argv.includes('--manual-prepare');
const manualVerifyNonActive = process.argv.includes('--manual-verify-nonactive');
const manualVerifyCreate = process.argv.includes('--manual-verify-create');
const manualVerifyMetadata = process.argv.includes('--manual-verify-metadata');
const manualVerifyContent = process.argv.includes('--manual-verify-content');
const manualVerifyDelete = process.argv.includes('--manual-verify-delete');
const manualMode = manualPrepare || manualVerifyNonActive || manualVerifyCreate || manualVerifyMetadata || manualVerifyContent || manualVerifyDelete;

const argValue = (name) => {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : '';
};

const testFiles = [
  'Agentic_AI_Holy_Book.pdf',
  'Payment_Tracking_Agent_Approach_Plan.pdf',
  'AI_Meeting_Platform_Feasibility_Reference.docx'
];

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const odataEscape = (value) => String(value || '').replace(/'/g, "''");
const runIdFromDate = () => new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14);
const prefixForRun = (runId) => `iknowledge-crud-prod-test-${runId}-`;
const inventoryPathForRun = (runId) => path.resolve('tools', 'search-backend', '.inventory', `crud-prod-controlled-${safeSearchKey(runId)}.json`);
const statePathForRun = (runId) => path.resolve('C:/tmp', `iknowledge-crud-prod-controlled-${runId}.json`);
const childEnvForRun = (runId) => ({
  SEARCH_INVENTORY_PATH: inventoryPathForRun(runId),
  SYNC_DELTA_SOURCE_KEY: `crud-prod-controlled-${safeSearchKey(runId)}`
});

const manualItemsForRun = (runId) => testFiles.map((originalName) => ({
  originalName,
  fileName: `${prefixForRun(runId)}${originalName}`
}));

const readManualState = (runId) => {
  const statePath = statePathForRun(runId);
  if (!fs.existsSync(statePath)) {
    throw new Error(`Manual CRUD state not found: ${statePath}. Run --manual-prepare or --manual-verify-create first.`);
  }
  return JSON.parse(fs.readFileSync(statePath, 'utf8'));
};

const writeManualState = (state) => {
  const reportPath = statePathForRun(state.runId);
  fs.writeFileSync(reportPath, JSON.stringify({ ...state, updatedAt: new Date().toISOString() }, null, 2));
  return reportPath;
};

const makeManualUploadCopies = (runId) => {
  const baseDir = path.resolve('C:/tmp', `iknowledge-crud-upload-${runId}`);
  const uploadDir = path.join(baseDir, '01-upload-these');
  const replacementDir = path.join(baseDir, '02-optional-content-replacement');
  fs.mkdirSync(uploadDir, { recursive: true });
  fs.mkdirSync(replacementDir, { recursive: true });

  const items = manualItemsForRun(runId);
  for (const item of items) {
    fs.copyFileSync(path.resolve(item.originalName), path.join(uploadDir, item.fileName));
  }

  const firstPdf = items.find((item) => item.originalName.endsWith('.pdf'));
  const replacementPdf = testFiles.find((fileName) => fileName.endsWith('.pdf') && fileName !== (firstPdf && firstPdf.originalName));
  if (firstPdf && replacementPdf) {
    fs.copyFileSync(path.resolve(replacementPdf), path.join(replacementDir, firstPdf.fileName));
  }

  return { baseDir, uploadDir, replacementDir, items, contentTarget: firstPdf ? firstPdf.fileName : '' };
};

const mimeTypeFor = (fileName) => {
  const ext = path.extname(fileName).toLowerCase();
  if (ext === '.pdf') return 'application/pdf';
  if (ext === '.docx') return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  return 'application/octet-stream';
};

const runDelta = (mode, label, childEnv) => {
  console.log(`\n=== Delta ${label} ===`);
  const result = spawnSync(
    process.execPath,
    ['tools/search-backend/scripts/delta-sync.js', envPath, mode, 'all'],
    {
      cwd: process.cwd(),
      env: { ...process.env, ...childEnv },
      encoding: 'utf8',
      maxBuffer: 50 * 1024 * 1024
    }
  );

  if (result.stdout) process.stdout.write(result.stdout);
  if (result.stderr) process.stderr.write(result.stderr);
  if (result.status !== 0) {
    throw new Error(`Delta ${label} failed with exit code ${result.status}`);
  }

  return `${result.stdout || ''}\n${result.stderr || ''}`;
};

const findDocumentByFileName = async (env, fileName) => {
  const result = await searchIndex(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, {
    search: '*',
    filter: `fileName eq '${odataEscape(fileName)}'`,
    select: 'id,listItemId,title,fileName,status,contentPreview,lastModifiedDateTime,indexedAt',
    top: 5
  });
  return (result.value || [])[0];
};

const countChunks = async (env, documentId) => {
  const result = await searchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
    search: '*',
    filter: `documentId eq '${odataEscape(documentId)}'`,
    select: 'id,documentId,title,fileName,chunkOrdinal,indexedAt',
    top: 1000
  });
  return result.value || [];
};

const waitFor = async (label, assertion, timeoutMs = 180000, intervalMs = 5000) => {
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

const assertIndexed = async (env, fileName, { requireChunks = true, titleIncludes = '' } = {}) => waitFor(
  `${fileName} indexed${requireChunks ? ' with chunks' : ''}`,
  async () => {
    const document = await findDocumentByFileName(env, fileName);
    if (!document) throw new Error('document not found');
    if (String(document.status || '').toLowerCase() !== 'active') {
      throw new Error(`document status is ${document.status}`);
    }
    if (titleIncludes && !String(document.title || '').includes(titleIncludes)) {
      throw new Error(`title "${document.title}" does not include "${titleIncludes}"`);
    }
    const chunks = await countChunks(env, document.id);
    if (requireChunks && chunks.length === 0) throw new Error('chunks not found');
    return { document, chunks };
  }
);

const assertNotIndexed = async (env, fileName) => waitFor(
  `${fileName} removed from documents index`,
  async () => {
    const document = await findDocumentByFileName(env, fileName);
    if (document) throw new Error(`still indexed as ${document.id}`);
    return true;
  }
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

const cleanupFiles = async (env, token, uploaded, childEnv) => {
  const remaining = uploaded.filter((item) => item.driveItemId);
  if (remaining.length === 0) return;

  console.log('\n=== Best-effort cleanup ===');
  for (const item of remaining) {
    try {
      await deleteDriveItem(env, token, item.driveItemId);
      item.driveItemId = '';
      console.log(`Cleanup deleted SharePoint file: ${item.fileName}`);
    } catch (error) {
      console.warn(`WARN: cleanup failed for ${item.fileName}: ${error.message}`);
    }
  }

  try {
    runDelta('delta', 'cleanup after failed controlled test', childEnv);
  } catch (error) {
    console.warn(`WARN: cleanup delta failed: ${error.message}`);
  }
};

const printManualNextSteps = (state, nextStep) => {
  console.log(JSON.stringify({
    mode: 'manual-frontend-upload',
    runId: state.runId,
    prefix: state.prefix,
    uploadDir: state.uploadDir,
    replacementDir: state.replacementDir,
    contentTarget: state.contentTarget,
    statePath: statePathForRun(state.runId),
    inventoryPath: state.inventoryPath,
    nextStep,
    files: state.items.map((item) => item.fileName)
  }, null, 2));
};

const runManualPrepare = async (env, absolutePath) => {
  const runId = argValue('--run-id') || runIdFromDate();
  const prefix = prefixForRun(runId);
  const childEnv = childEnvForRun(runId);
  const prepared = makeManualUploadCopies(runId);

  console.log(`Using env: ${absolutePath}`);
  console.log(`Using isolated CRUD inventory: ${childEnv.SEARCH_INVENTORY_PATH}`);
  console.log(`Test prefix: ${prefix}`);
  console.log('Manual mode will not mutate SharePoint directly. Upload/delete happens only through your frontend/admin session.');

  const checkpointOutput = runDelta('checkpoint', 'manual checkpoint before frontend upload', childEnv);
  const state = {
    ok: true,
    mode: 'manual-frontend-upload',
    phase: 'prepared',
    env: absolutePath,
    runId,
    prefix,
    inventoryPath: childEnv.SEARCH_INVENTORY_PATH,
    uploadDir: prepared.uploadDir,
    replacementDir: prepared.replacementDir,
    contentTarget: prepared.contentTarget,
    items: prepared.items,
    observations: [{
      step: 'checkpoint',
      checkpointLogged: checkpointOutput.includes('Delta sync summary') || checkpointOutput.includes('"mode": "checkpoint"')
    }]
  };

  writeManualState(state);
  printManualNextSteps(state, [
    'Upload every file from uploadDir through the frontend upload flow.',
    'Do not rename them in the frontend. The iknowledge-crud-prod-test-* prefix is the safety boundary.',
    'The frontend upload path normally saves new files as Under Review.',
    `Then run: npm run search:crud:test:prod-controlled -- --yes --prod-controlled --manual-verify-nonactive --run-id ${runId}`,
    'After that, approve/change only these prefixed files to Active.',
    `Then run: npm run search:crud:test:prod-controlled -- --yes --prod-controlled --manual-verify-create --run-id ${runId}`
  ]);
};

const runManualVerifyNonActive = async (env, runId) => {
  const state = readManualState(runId);
  const childEnv = childEnvForRun(runId);
  const nonActiveOutput = runDelta('delta', 'after manual frontend Under Review upload', childEnv);

  state.phase = 'non-active-verified';
  for (const item of state.items) {
    await assertNotIndexed(env, item.fileName);
    state.observations.push({
      step: 'non_active_upload_excluded',
      fileName: item.fileName,
      nonActiveDeletedLogged: nonActiveOutput.includes(`"fileName": "${item.fileName}"`) && nonActiveOutput.includes('"action": "non_active_deleted"')
    });
  }

  writeManualState(state);
  printManualNextSteps(state, [
    'Good: uploaded Under Review files are not searchable.',
    'Now approve/change ONLY the prefixed test files to Active through the frontend/admin UI.',
    `Then run: npm run search:crud:test:prod-controlled -- --yes --prod-controlled --manual-verify-create --run-id ${runId}`
  ]);
};

const runManualVerifyCreate = async (env, runId) => {
  const state = readManualState(runId);
  const childEnv = childEnvForRun(runId);
  const createOutput = runDelta('delta', 'after manual Active approval/create', childEnv);

  state.phase = 'created-verified';
  for (const item of state.items) {
    const indexed = await assertIndexed(env, item.fileName);
    item.documentId = indexed.document.id;
    item.listItemId = indexed.document.listItemId;
    item.initialChunkCount = indexed.chunks.length;
    state.observations.push({
      step: 'create',
      fileName: item.fileName,
      listItemId: item.listItemId,
      documentId: item.documentId,
      chunkCount: indexed.chunks.length,
      indexedLogged: createOutput.includes(`"fileName": "${item.fileName}"`) && createOutput.includes('"action": "indexed"')
    });
  }

  writeManualState(state);
  printManualNextSteps(state, [
    'Now make a metadata-only edit through the frontend/admin UI.',
    'Best test: change only Title/Description/metadata, not the file content.',
    `Then run: npm run search:crud:test:prod-controlled -- --yes --prod-controlled --manual-verify-metadata --run-id ${runId}`
  ]);
};

const runManualVerifyMetadata = async (env, runId) => {
  const state = readManualState(runId);
  const childEnv = childEnvForRun(runId);
  const metadataOutput = runDelta('delta', 'after manual metadata-only edit', childEnv);

  state.phase = 'metadata-verified';
  for (const item of state.items) {
    const indexed = await assertIndexed(env, item.fileName);
    state.observations.push({
      step: 'metadata_update',
      fileName: item.fileName,
      chunkCountBefore: item.initialChunkCount,
      chunkCountAfter: indexed.chunks.length,
      metadataUpdatedLogged: metadataOutput.includes(`"fileName": "${item.fileName}"`) && metadataOutput.includes('"action": "metadata_updated"')
    });
  }

  writeManualState(state);
  printManualNextSteps(state, [
    `Optional content-update test: replace ${state.contentTarget || 'one uploaded PDF'} through the frontend/version flow using the matching file in replacementDir.`,
    `Then run: npm run search:crud:test:prod-controlled -- --yes --prod-controlled --manual-verify-content --run-id ${runId}`,
    'If you want to skip content replacement, delete the prefixed test files and run --manual-verify-delete.'
  ]);
};

const runManualVerifyContent = async (env, runId) => {
  const state = readManualState(runId);
  const childEnv = childEnvForRun(runId);
  const contentOutput = runDelta('delta', 'after manual content replacement', childEnv);
  const contentTarget = state.items.find((item) => item.fileName === state.contentTarget) || state.items[0];

  if (!contentTarget) {
    throw new Error('No manual content target found in state.');
  }

  const indexed = await assertIndexed(env, contentTarget.fileName);
  state.phase = 'content-verified';
  state.observations.push({
    step: 'content_update',
    fileName: contentTarget.fileName,
    chunkCountBefore: contentTarget.initialChunkCount,
    chunkCountAfter: indexed.chunks.length,
    indexedLogged: contentOutput.includes(`"fileName": "${contentTarget.fileName}"`) && contentOutput.includes('"action": "indexed"')
  });

  writeManualState(state);
  printManualNextSteps(state, [
    'Now delete ONLY the three prefixed test files through the frontend/admin UI.',
    `Then run: npm run search:crud:test:prod-controlled -- --yes --prod-controlled --manual-verify-delete --run-id ${runId}`
  ]);
};

const runManualVerifyDelete = async (env, runId) => {
  const state = readManualState(runId);
  const childEnv = childEnvForRun(runId);
  const deleteOutput = runDelta('delta', 'after manual frontend delete', childEnv);

  state.phase = 'delete-verified';
  for (const item of state.items) {
    await assertNotIndexed(env, item.fileName);
    state.observations.push({
      step: 'delete',
      fileName: item.fileName,
      deletedLogged: deleteOutput.includes(`"fileName": "${item.fileName}"`) && deleteOutput.includes('"action": "deleted"')
    });
  }

  state.completedAt = new Date().toISOString();
  writeManualState(state);
  printManualNextSteps(state, [
    'Manual CRUD battle test complete.',
    'Review the state JSON and delta logs before packaging.'
  ]);
};

const runManual = async (env, absolutePath) => {
  const runId = argValue('--run-id');
  if (manualPrepare) {
    await runManualPrepare(env, absolutePath);
    return;
  }

  if (!runId) {
    throw new Error('Manual verify mode requires --run-id <runId> from the --manual-prepare output.');
  }

  if (manualVerifyNonActive) {
    await runManualVerifyNonActive(env, runId);
  } else if (manualVerifyCreate) {
    await runManualVerifyCreate(env, runId);
  } else if (manualVerifyMetadata) {
    await runManualVerifyMetadata(env, runId);
  } else if (manualVerifyContent) {
    await runManualVerifyContent(env, runId);
  } else if (manualVerifyDelete) {
    await runManualVerifyDelete(env, runId);
  }
};

const run = async () => {
  if (!confirmed) {
    throw new Error([
      'This script mutates the configured SharePoint library and Azure Search prod indexes using ONLY unique test files.',
      'It refuses to run unless both flags are present:',
      'node tools/search-backend/scripts/crud-prod-controlled-test.js config/search.env --yes --prod-controlled'
    ].join('\n'));
  }

  const missing = testFiles.filter((fileName) => !fs.existsSync(path.resolve(fileName)));
  if (missing.length > 0) {
    throw new Error(`Missing controlled test files: ${missing.join(', ')}`);
  }

  const { env, absolutePath } = loadEnv(envPath);
  const hostname = String(env.SHAREPOINT_HOSTNAME || '').toLowerCase();
  if (!hostname.includes('indegene123.sharepoint.com')) {
    throw new Error(`Refusing controlled prod CRUD test against unexpected SharePoint host: ${env.SHAREPOINT_HOSTNAME || '(missing)'}`);
  }

  if (manualMode) {
    await runManual(env, absolutePath);
    return;
  }

  const token = await getGraphToken(env);
  const runId = new Date().toISOString().replace(/[-:.TZ]/g, '').slice(0, 14);
  const prefix = `iknowledge-crud-prod-test-${runId}-`;
  const inventoryPath = path.resolve('tools', 'search-backend', '.inventory', `crud-prod-controlled-${safeSearchKey(runId)}.json`);
  const childEnv = {
    SEARCH_INVENTORY_PATH: inventoryPath,
    SYNC_DELTA_SOURCE_KEY: `crud-prod-controlled-${safeSearchKey(runId)}`
  };
  const uploaded = [];
  const observations = [];

  console.log(`Using env: ${absolutePath}`);
  console.log(`Using isolated CRUD inventory: ${inventoryPath}`);
  console.log(`Test prefix: ${prefix}`);
  console.log('No existing SharePoint files will be modified by name or id.');

  try {
    runDelta('checkpoint', 'checkpoint before controlled prod test files', childEnv);

    console.log('\n=== Create controlled files ===');
    for (const originalName of testFiles) {
      const fileName = `${prefix}${originalName}`;
      const buffer = fs.readFileSync(path.resolve(originalName));
      const created = await withRetry(
        () => uploadDriveItemContentByPath(env, token, fileName, buffer, mimeTypeFor(originalName)),
        { retries: 4, baseDelayMs: 1500, label: `upload controlled file ${fileName}` }
      );
      const listItem = await withRetry(
        () => fetchListItemByDriveItemId(env, token, created.id),
        { retries: 6, baseDelayMs: 1500, label: `fetch created list item ${fileName}` }
      );
      await updateFieldsBestEffort(env, token, listItem.id, {
        Title: `CRUD Prod Test ${runId} ${path.basename(originalName, path.extname(originalName))}`,
        Status: 'Active'
      }, `set active metadata ${fileName}`);
      uploaded.push({
        originalName,
        fileName,
        driveItemId: created.id,
        listItemId: String(listItem.id)
      });
      console.log(`Created ${fileName} (listItemId=${listItem.id})`);
    }

    const createOutput = runDelta('delta', 'after controlled create', childEnv);
    for (const item of uploaded) {
      const indexed = await assertIndexed(env, item.fileName);
      item.documentId = indexed.document.id;
      item.initialChunkCount = indexed.chunks.length;
      observations.push({
        step: 'create',
        fileName: item.fileName,
        listItemId: item.listItemId,
        chunkCount: indexed.chunks.length
      });
    }

    console.log('\n=== Metadata-only update ===');
    for (const item of uploaded) {
      await updateFieldsBestEffort(env, token, item.listItemId, {
        Title: `CRUD Prod Test ${runId} Metadata Updated ${path.basename(item.originalName, path.extname(item.originalName))}`,
        Status: 'Active'
      }, `metadata-only update ${item.fileName}`);
    }
    const metadataOutput = runDelta('delta', 'after controlled metadata-only update', childEnv);
    for (const item of uploaded) {
      const indexed = await assertIndexed(env, item.fileName, { titleIncludes: 'Metadata Updated' });
      observations.push({
        step: 'metadata_update',
        fileName: item.fileName,
        chunkCountBefore: item.initialChunkCount,
        chunkCountAfter: indexed.chunks.length,
        metadataUpdatedLogged: metadataOutput.includes(`"fileName": "${item.fileName}"`) && metadataOutput.includes('"action": "metadata_updated"')
      });
    }

    console.log('\n=== Content update on first PDF ===');
    const contentTarget = uploaded.find((item) => item.originalName.endsWith('.pdf'));
    const replacementPdf = testFiles.find((fileName) => fileName.endsWith('.pdf') && fileName !== contentTarget.originalName);
    if (contentTarget && replacementPdf) {
      await withRetry(
        () => updateDriveItemContent(env, token, contentTarget.driveItemId, fs.readFileSync(path.resolve(replacementPdf)), mimeTypeFor(replacementPdf)),
        { retries: 4, baseDelayMs: 1500, label: `content update ${contentTarget.fileName}` }
      );
      const contentOutput = runDelta('delta', 'after controlled content update', childEnv);
      const indexed = await assertIndexed(env, contentTarget.fileName);
      observations.push({
        step: 'content_update',
        fileName: contentTarget.fileName,
        replacementSource: replacementPdf,
        chunkCountBefore: contentTarget.initialChunkCount,
        chunkCountAfter: indexed.chunks.length,
        indexedLogged: contentOutput.includes(`"fileName": "${contentTarget.fileName}"`) && contentOutput.includes('"action": "indexed"')
      });
    }

    if (!skipInactive && uploaded.length > 0) {
      console.log('\n=== Active to non-active and reactivation ===');
      const inactiveTarget = uploaded[uploaded.length - 1];
      const inactiveApplied = await updateFieldsBestEffort(env, token, inactiveTarget.listItemId, {
        Status: 'Inactive'
      }, `set inactive ${inactiveTarget.fileName}`);
      if (inactiveApplied) {
        const inactiveOutput = runDelta('delta', 'after controlled inactive status', childEnv);
        await assertNotIndexed(env, inactiveTarget.fileName);
        observations.push({
          step: 'inactive',
          fileName: inactiveTarget.fileName,
          nonActiveDeletedLogged: inactiveOutput.includes(`"fileName": "${inactiveTarget.fileName}"`) && inactiveOutput.includes('"action": "non_active_deleted"')
        });

        await updateFieldsBestEffort(env, token, inactiveTarget.listItemId, {
          Status: 'Active'
        }, `reactivate ${inactiveTarget.fileName}`);
        const reactivateOutput = runDelta('delta', 'after controlled reactivation', childEnv);
        const indexed = await assertIndexed(env, inactiveTarget.fileName);
        observations.push({
          step: 'reactivate',
          fileName: inactiveTarget.fileName,
          chunkCount: indexed.chunks.length,
          indexedLogged: reactivateOutput.includes(`"fileName": "${inactiveTarget.fileName}"`) && reactivateOutput.includes('"action": "indexed"')
        });
      } else {
        observations.push({
          step: 'inactive',
          fileName: inactiveTarget.fileName,
          skipped: true,
          reason: 'SharePoint rejected Inactive status value'
        });
      }
    }

    if (!keepFiles) {
      console.log('\n=== Delete controlled files ===');
      for (const item of uploaded) {
        await withRetry(
          () => deleteDriveItem(env, token, item.driveItemId),
          { retries: 4, baseDelayMs: 1500, label: `delete controlled file ${item.fileName}` }
        );
        item.driveItemId = '';
      }
      const deleteOutput = runDelta('delta', 'after controlled delete', childEnv);
      for (const item of uploaded) {
        await assertNotIndexed(env, item.fileName);
        observations.push({
          step: 'delete',
          fileName: item.fileName,
          deletedLogged: deleteOutput.includes('"action": "deleted"')
        });
      }
    }

    const summary = {
      ok: true,
      env: absolutePath,
      runId,
      prefix,
      uploaded: uploaded.map((item) => ({
        originalName: item.originalName,
        fileName: item.fileName,
        listItemId: item.listItemId,
        documentId: item.documentId,
        initialChunkCount: item.initialChunkCount
      })),
      keepFiles,
      inventoryPath,
      observations
    };
    const reportPath = path.resolve('C:/tmp', `iknowledge-crud-prod-controlled-${runId}.json`);
    fs.writeFileSync(reportPath, JSON.stringify(summary, null, 2));
    console.log(JSON.stringify({ ...summary, reportPath }, null, 2));
  } finally {
    if (!keepFiles) {
      await cleanupFiles(env, token, uploaded, childEnv);
    }
  }
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
