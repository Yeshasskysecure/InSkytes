const { loadEnv, toNumber } = require('./lib/env');
const { chunkText } = require('./lib/chunker');
const { deleteDocuments, searchIndex, uploadDocuments } = require('./lib/azureSearchClient');
const { extractDocumentText } = require('./lib/extractors');
const { downloadDriveItem, fetchDriveDeltaPage, fetchListDeltaPage, fetchListItemByDriveItemId, getGraphToken } = require('./lib/graphClient');
const { findSourceItemByDriveItemId, forgetSourceItemByDriveItemId, loadInventory, rememberSourceItem, saveInventory, summarizeInventory } = require('./lib/inventory');
const { mapDocumentMetadataRecord, mapDocumentRecord } = require('./lib/mappers');
const { embedTextsInBatches } = require('./lib/openAiClient');
const { withRetry } = require('./lib/retry');

const envPath = process.argv[2] || 'config/search.env';
const mode = process.argv[3] || 'delta';
const maxChangesArg = process.argv[4] || '25';
const maxChanges = String(maxChangesArg).toLowerCase() === 'all'
  ? Number.MAX_SAFE_INTEGER
  : toNumber(maxChangesArg, 25);

const logProgress = (event, details = {}) => {
  console.log(JSON.stringify({
    event,
    at: new Date().toISOString(),
    ...details
  }));
};

const isActive = (status) => String(status || '').toLowerCase() === 'active';
const sourceSignature = (item) => {
  const fields = item.fields || {};
  const driveItem = item.driveItem || {};
  const edited = fields.Edited || item.lastModifiedDateTime || fields.Modified || '';
  const editor = fields.Editor0 || fields.Editor || '';
  return [
    item.id || '',
    driveItem.id || '',
    driveItem.eTag || '',
    driveItem.cTag || '',
    edited,
    editor,
    fields.Status || ''
  ].join('|');
};
const maxChunksPerDocument = (env) => toNumber(env.SYNC_MAX_CHUNKS_PER_DOCUMENT, 0);
const contentSignature = (item) => {
  const driveItem = item.driveItem || {};
  return [
    item.id || '',
    driveItem.id || '',
    driveItem.cTag || driveItem.eTag || '',
    driveItem.size || 0
  ].join('|');
};
const inventoryIdentity = (record, item) => {
  const driveItem = (item && item.driveItem) || {};
  return {
    id: record.id,
    driveItemId: driveItem.id || '',
    listItemId: String((item && item.id) || record.listItemId || ''),
    eTag: driveItem.eTag || '',
    cTag: driveItem.cTag || '',
    size: driveItem.size || 0
  };
};

const hasSameContent = (existing, item) => {
  if (!existing || existing.indexStatus === 'failed') return false;
  if (!['indexed', 'metadata_only'].includes(existing.indexStatus)) return false;
  const driveItem = item.driveItem || {};
  const signature = contentSignature(item);
  if (existing.contentSignature && existing.contentSignature === signature) return true;
  return Boolean(
    existing.driveItemId
      && existing.driveItemId === driveItem.id
      && existing.cTag
      && driveItem.cTag
      && existing.cTag === driveItem.cTag
      && Number(existing.size || 0) === Number(driveItem.size || 0)
  );
};

const documentMetadataPatch = (metadataRecord) => {
  const { contentPreview, ...patch } = metadataRecord;
  patch.indexedAt = new Date().toISOString();
  return patch;
};

const chunkMetadataPatch = (id, metadataRecord) => ({
  id,
  title: metadataRecord.title,
  fileName: metadataRecord.fileName,
  webUrl: metadataRecord.webUrl,
  status: metadataRecord.status,
  bu: metadataRecord.bu,
  buFilterValues: metadataRecord.buFilterValues,
  department: metadataRecord.department,
  departmentFilterValues: metadataRecord.departmentFilterValues,
  departmentFacetValues: metadataRecord.departmentFacetValues,
  diseaseArea: metadataRecord.diseaseArea,
  diseaseAreaFilterValues: metadataRecord.diseaseAreaFilterValues,
  therapyArea: metadataRecord.therapyArea,
  therapyAreaFilterValues: metadataRecord.therapyAreaFilterValues,
  client: metadataRecord.client,
  clientFilterValues: metadataRecord.clientFilterValues,
  region: metadataRecord.region,
  regionFilterValues: metadataRecord.regionFilterValues,
  documentType: metadataRecord.documentType,
  documentTypeFilterValues: metadataRecord.documentTypeFilterValues,
  description: metadataRecord.description,
  authors: metadataRecord.authors,
  lastModifiedDateTime: metadataRecord.lastModifiedDateTime,
  indexedAt: new Date().toISOString()
});

const findChunkIdsForDocument = async (env, documentId) => {
  const results = await withRetry(() => searchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
    search: '*',
    filter: `documentId eq '${String(documentId).replace(/'/g, "''")}'`,
    select: 'id',
    top: 1000
  }), { retries: 5, baseDelayMs: 1500, label: `lookup chunks for metadata update ${documentId}` });
  return (results.value || []).map((item) => item.id).filter(Boolean);
};

const deleteChunksForDocument = async (env, documentId) => {
  const results = await withRetry(() => searchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
    search: '*',
    filter: `documentId eq '${String(documentId).replace(/'/g, "''")}'`,
    select: 'id',
    top: 1000
  }), { retries: 5, baseDelayMs: 1500, label: `lookup chunks for ${documentId}` });
  const ids = (results.value || []).map((item) => item.id).filter(Boolean);
  await withRetry(() => deleteDocuments(env, env.AZURE_SEARCH_CHUNKS_INDEX, 'id', ids), { retries: 5, baseDelayMs: 1500, label: `delete chunks for ${documentId}` });
  return ids.length;
};

const updateMetadataForUnchangedContent = async (env, inventory, existing, metadataRecord, listItem) => {
  const chunkIds = existing.indexStatus === 'indexed'
    ? await findChunkIdsForDocument(env, metadataRecord.id)
    : [];

  await withRetry(() => uploadDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, [documentMetadataPatch(metadataRecord)]), {
    retries: 5,
    baseDelayMs: 1500,
    label: `update document metadata ${metadataRecord.fileName}`
  });

  if (chunkIds.length > 0) {
    await withRetry(() => uploadDocuments(env, env.AZURE_SEARCH_CHUNKS_INDEX, chunkIds.map((id) => chunkMetadataPatch(id, metadataRecord))), {
      retries: 5,
      baseDelayMs: 1500,
      label: `update chunk metadata ${metadataRecord.fileName}`
    });
  }

  const signature = sourceSignature(listItem);
  rememberSourceItem(inventory, {
    ...inventoryIdentity(metadataRecord, listItem),
    sourceSignature: signature,
    metadataSignature: signature,
    contentSignature: contentSignature(listItem),
    driveItemId: (listItem.driveItem && listItem.driveItem.id) || existing.driveItemId || '',
    fileName: metadataRecord.fileName,
    status: metadataRecord.status,
    indexStatus: existing.indexStatus,
    chunkCount: existing.chunkCount || chunkIds.length,
    lastMetadataSyncedAt: new Date().toISOString()
  });

  return { fileName: metadataRecord.fileName, action: 'metadata_updated', chunksUpdated: chunkIds.length };
};

const deleteDocumentArtifacts = async (env, documentId) => {
  const deletedChunks = await deleteChunksForDocument(env, documentId);
  await withRetry(() => deleteDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, 'id', [String(documentId)]), { retries: 5, baseDelayMs: 1500, label: `delete document ${documentId}` });
  return { deletedChunks, deletedDocuments: 1 };
};

const findDocumentIdsForListItem = async (env, listItemId) => {
  const results = await withRetry(() => searchIndex(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, {
    search: '*',
    filter: `listItemId eq '${String(listItemId).replace(/'/g, "''")}'`,
    select: 'id',
    top: 1000
  }), { retries: 5, baseDelayMs: 1500, label: `lookup documents for list item ${listItemId}` });
  return (results.value || []).map((item) => item.id).filter(Boolean);
};

const deleteDocumentArtifactsByListItemId = async (env, listItemId) => {
  const documentIds = await findDocumentIdsForListItem(env, listItemId);
  let deletedChunks = 0;
  for (const documentId of documentIds) {
    const deletion = await deleteDocumentArtifacts(env, documentId);
    deletedChunks += deletion.deletedChunks;
  }
  return { deletedChunks, deletedDocuments: documentIds.length };
};

const buildChunkRecords = async (env, documentRecord, extractedText) => {
  const chunks = chunkText(extractedText, {
    targetTokens: toNumber(env.EXTRACTION_CHUNK_TARGET_TOKENS, 750),
    overlapTokens: toNumber(env.EXTRACTION_CHUNK_OVERLAP_TOKENS, 120),
    maxChars: toNumber(env.EXTRACTION_CHUNK_MAX_CHARS, 5000)
  });
  const chunkLimit = maxChunksPerDocument(env);
  if (chunkLimit > 0 && chunks.length > chunkLimit) {
    return {
      tooManyChunks: true,
      chunks: [],
      chunkCount: chunks.length,
      chunkLimit
    };
  }
  const inputs = chunks.map((chunk, index) => [
    `Title: ${documentRecord.title}`,
    documentRecord.description ? `Description: ${String(documentRecord.description).slice(0, 500)}` : '',
    documentRecord.documentType ? `Document Type: ${documentRecord.documentType}` : '',
    documentRecord.bu ? `BU: ${documentRecord.bu}` : '',
    documentRecord.department ? `Department: ${documentRecord.department}` : '',
    `Chunk: ${index + 1}`,
    chunk
  ].filter(Boolean).join('\n'));
  const embeddings = await withRetry(
    () => embedTextsInBatches(env, inputs),
    { retries: 8, baseDelayMs: 5000, label: `embed chunks for ${documentRecord.fileName}` }
  );
  const vectorField = env.AZURE_SEARCH_VECTOR_FIELD || 'chunkVector';
  const now = new Date().toISOString();

  return chunks.map((chunk, index) => ({
    id: `${documentRecord.id}-chunk-${index}`,
    documentId: documentRecord.id,
    siteId: documentRecord.siteId,
    driveId: documentRecord.driveId,
    listId: documentRecord.listId,
    listItemId: documentRecord.listItemId,
    title: documentRecord.title,
    fileName: documentRecord.fileName,
    webUrl: documentRecord.webUrl,
    status: documentRecord.status,
    bu: documentRecord.bu,
    buFilterValues: documentRecord.buFilterValues,
    department: documentRecord.department,
    departmentFilterValues: documentRecord.departmentFilterValues,
    departmentFacetValues: documentRecord.departmentFacetValues,
    diseaseArea: documentRecord.diseaseArea,
    diseaseAreaFilterValues: documentRecord.diseaseAreaFilterValues,
    therapyArea: documentRecord.therapyArea,
    therapyAreaFilterValues: documentRecord.therapyAreaFilterValues,
    client: documentRecord.client,
    clientFilterValues: documentRecord.clientFilterValues,
    region: documentRecord.region,
    regionFilterValues: documentRecord.regionFilterValues,
    documentType: documentRecord.documentType,
    documentTypeFilterValues: documentRecord.documentTypeFilterValues,
    description: documentRecord.description,
    authors: documentRecord.authors,
    chunkOrdinal: index,
    chunkText: chunk,
    [vectorField]: embeddings[index],
    sourceKind: 'document',
    lastModifiedDateTime: documentRecord.lastModifiedDateTime,
    indexedAt: now
  }));
};

const collectDeltaChanges = async (env, token, inventory, sourceKey) => {
  const source = inventory.syncSources[sourceKey] || {};
  const startLink = mode === 'init' || mode === 'checkpoint'
    ? undefined
    : source.pendingDeltaNextLink || source.deltaLink;
  const changedItems = [];
  let deltaLink = '';
  let pendingDeltaNextLink = '';
  let nextLink = startLink;
  let pageNumber = 0;

  logProgress('delta-drive-collection-started', {
    sourceKey,
    hasCheckpoint: Boolean(startLink),
    maxChanges: maxChanges === Number.MAX_SAFE_INTEGER ? 'all' : maxChanges
  });

  while (changedItems.length < maxChanges) {
    pageNumber += 1;
    logProgress('delta-drive-page-requested', {
      sourceKey,
      pageNumber,
      collectedItems: changedItems.length,
      hasContinuationLink: Boolean(nextLink)
    });
    const page = await withRetry(() => fetchDriveDeltaPage(env, token, nextLink), { retries: 6, baseDelayMs: 1500, label: 'fetch drive delta page' });
    const pageItems = page.value || [];
    changedItems.push(...pageItems);
    logProgress('delta-drive-page-completed', {
      sourceKey,
      pageNumber,
      pageItems: pageItems.length,
      collectedItems: changedItems.length,
      hasNextLink: Boolean(page['@odata.nextLink']),
      hasDeltaLink: Boolean(page['@odata.deltaLink'])
    });
    if (page['@odata.nextLink']) {
      nextLink = page['@odata.nextLink'];
    } else {
      deltaLink = page['@odata.deltaLink'] || '';
      break;
    }
  }

  if (changedItems.length >= maxChanges && nextLink && !deltaLink) {
    pendingDeltaNextLink = nextLink;
  }

  logProgress('delta-drive-collection-completed', {
    sourceKey,
    pages: pageNumber,
    changedItems: Math.min(changedItems.length, maxChanges),
    hasNewDeltaLink: Boolean(deltaLink),
    hasPendingDeltaNextLink: Boolean(pendingDeltaNextLink)
  });

  return { changedItems: changedItems.slice(0, maxChanges), deltaLink, pendingDeltaNextLink };
};

const collectListDeltaChanges = async (env, token, inventory, sourceKey) => {
  const source = inventory.syncSources[sourceKey] || {};
  const bootstrapLatest = String(env.SYNC_LIST_DELTA_BOOTSTRAP_LATEST || process.env.SYNC_LIST_DELTA_BOOTSTRAP_LATEST || 'true').toLowerCase() !== 'false';

  if ((mode === 'checkpoint' || (mode === 'delta' && bootstrapLatest && !source.deltaLink && !source.pendingDeltaNextLink))) {
    logProgress('delta-list-bootstrap-latest-started', { sourceKey });
    const page = await withRetry(() => fetchListDeltaPage(env, token, 'latest'), { retries: 6, baseDelayMs: 1500, label: 'bootstrap list item delta link' });
    logProgress('delta-list-bootstrap-latest-completed', {
      sourceKey,
      hasDeltaLink: Boolean(page['@odata.deltaLink'])
    });
    return {
      changedItems: [],
      deltaLink: page['@odata.deltaLink'] || '',
      pendingDeltaNextLink: '',
      bootstrapped: true
    };
  }

  const startLink = mode === 'init' || mode === 'checkpoint'
    ? undefined
    : source.pendingDeltaNextLink || source.deltaLink;
  const changedItems = [];
  let deltaLink = '';
  let pendingDeltaNextLink = '';
  let nextLink = startLink;
  let pageNumber = 0;

  logProgress('delta-list-collection-started', {
    sourceKey,
    hasCheckpoint: Boolean(startLink),
    maxChanges: maxChanges === Number.MAX_SAFE_INTEGER ? 'all' : maxChanges
  });

  while (changedItems.length < maxChanges) {
    pageNumber += 1;
    logProgress('delta-list-page-requested', {
      sourceKey,
      pageNumber,
      collectedItems: changedItems.length,
      hasContinuationLink: Boolean(nextLink)
    });
    const page = await withRetry(() => fetchListDeltaPage(env, token, nextLink), { retries: 6, baseDelayMs: 1500, label: 'fetch list item delta page' });
    const pageItems = page.value || [];
    changedItems.push(...pageItems);
    logProgress('delta-list-page-completed', {
      sourceKey,
      pageNumber,
      pageItems: pageItems.length,
      collectedItems: changedItems.length,
      hasNextLink: Boolean(page['@odata.nextLink']),
      hasDeltaLink: Boolean(page['@odata.deltaLink'])
    });
    if (page['@odata.nextLink']) {
      nextLink = page['@odata.nextLink'];
    } else {
      deltaLink = page['@odata.deltaLink'] || '';
      break;
    }
  }

  if (changedItems.length >= maxChanges && nextLink && !deltaLink) {
    pendingDeltaNextLink = nextLink;
  }

  logProgress('delta-list-collection-completed', {
    sourceKey,
    pages: pageNumber,
    changedItems: Math.min(changedItems.length, maxChanges),
    hasNewDeltaLink: Boolean(deltaLink),
    hasPendingDeltaNextLink: Boolean(pendingDeltaNextLink)
  });

  return { changedItems: changedItems.slice(0, maxChanges), deltaLink, pendingDeltaNextLink, bootstrapped: false };
};

const run = async () => {
  const { env, absolutePath } = loadEnv(envPath);
  const sourceKey = process.env.SYNC_DELTA_SOURCE_KEY || env.SYNC_DELTA_SOURCE_KEY || `drive-delta-${env.SHAREPOINT_LIBRARY_DRIVE_ID || 'library'}`;
  const listSourceKey = process.env.SYNC_LIST_DELTA_SOURCE_KEY || env.SYNC_LIST_DELTA_SOURCE_KEY || `list-delta-${env.SHAREPOINT_LIBRARY_LIST_ID || 'library'}`;
  if (!process.env.SEARCH_INVENTORY_PATH && (env.SEARCH_INVENTORY_PATH || env.SYNC_INVENTORY_PATH)) {
    process.env.SEARCH_INVENTORY_PATH = env.SEARCH_INVENTORY_PATH || env.SYNC_INVENTORY_PATH;
  }
  const { inventory, absolutePath: inventoryPath } = loadInventory();
  const token = await getGraphToken(env);
  console.log(`Using env: ${absolutePath}`);
  console.log(`Using inventory: ${inventoryPath}`);
  console.log(`Mode: ${mode}; max changes: ${maxChanges === Number.MAX_SAFE_INTEGER ? 'all' : maxChanges}; source: ${sourceKey}`);

  const { changedItems: driveChangedItems, deltaLink, pendingDeltaNextLink } = await collectDeltaChanges(env, token, inventory, sourceKey);
  const listDeltaEnabled = String(env.SYNC_LIST_DELTA_ENABLED || process.env.SYNC_LIST_DELTA_ENABLED || 'true').toLowerCase() !== 'false';
  const listDelta = listDeltaEnabled
    ? await collectListDeltaChanges(env, token, inventory, listSourceKey)
    : { changedItems: [], deltaLink: '', pendingDeltaNextLink: '', bootstrapped: false };

  if (mode === 'checkpoint') {
    if (!deltaLink) {
      throw new Error('Delta checkpoint did not reach a deltaLink. Run checkpoint with max changes "all".');
    }
    inventory.syncSources[sourceKey] = {
      ...(inventory.syncSources[sourceKey] || {}),
      deltaLink,
      pendingDeltaNextLink: '',
      lastCheckpointAt: new Date().toISOString(),
      lastSyncStatus: 'checkpoint'
    };
    if (listDeltaEnabled && listDelta.deltaLink) {
      inventory.syncSources[listSourceKey] = {
        ...(inventory.syncSources[listSourceKey] || {}),
        deltaLink: listDelta.deltaLink,
        pendingDeltaNextLink: '',
        lastCheckpointAt: new Date().toISOString(),
        lastSyncStatus: 'checkpoint'
      };
    }
    saveInventory(inventory);
    console.log(JSON.stringify({
      checkpointed: true,
      enumeratedItems: driveChangedItems.length,
      hasNewDeltaLink: true,
      inventory: summarizeInventory(inventory)
    }, null, 2));
    return;
  }

  const documentsToUpload = [];
  const chunksToUpload = [];
  const report = [];
  const listDeletedItems = listDelta.changedItems.filter((item) => item.deleted && item.id);
  for (const deletedItem of listDeletedItems) {
    logProgress('delta-list-deleted-item-processing', {
      listItemId: deletedItem.id
    });
    const deletion = await deleteDocumentArtifactsByListItemId(env, deletedItem.id);
    report.push({ listItemId: deletedItem.id, action: 'list_item_deleted', deletedChunks: deletion.deletedChunks, deletedDocuments: deletion.deletedDocuments });
  }

  const driveItemsById = new Map();
  driveChangedItems.forEach((item) => {
    if (item.id) driveItemsById.set(item.id, item);
  });
  listDelta.changedItems.forEach((item) => {
    const driveItem = item.driveItem;
    if (item.deleted || !driveItem || !driveItem.id || driveItemsById.has(driveItem.id)) return;
    driveItemsById.set(driveItem.id, driveItem);
  });
  const changedItems = Array.from(driveItemsById.values());
  logProgress('delta-changed-items-ready', {
    changedItems: changedItems.length,
    driveChangedItems: driveChangedItems.length,
    listChangedItems: listDelta.changedItems.length,
    listDeletedItems: listDeletedItems.length,
    maxChanges: maxChanges === Number.MAX_SAFE_INTEGER ? 'all' : maxChanges
  });

  for (let itemIndex = 0; itemIndex < changedItems.length; itemIndex += 1) {
    const driveDeltaItem = changedItems[itemIndex];
    const driveItemId = driveDeltaItem.id;
    if (!driveItemId) continue;
    const itemProgress = {
      itemIndex: itemIndex + 1,
      totalItems: changedItems.length,
      driveItemId,
      name: driveDeltaItem.name || ''
    };
    logProgress('delta-item-processing-started', itemProgress);

    if (driveDeltaItem.deleted) {
      logProgress('delta-item-delete-started', itemProgress);
      const inventoryItem = findSourceItemByDriveItemId(inventory, driveItemId);
      const documentId = inventoryItem && inventoryItem.id ? inventoryItem.id : driveItemId;
      const deletedChunks = await deleteChunksForDocument(env, documentId);
      await withRetry(() => deleteDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, 'id', [documentId]), { retries: 5, baseDelayMs: 1500, label: `delete delta document ${documentId}` });
      forgetSourceItemByDriveItemId(inventory, driveItemId);
      rememberSourceItem(inventory, {
        id: documentId,
        driveItemId,
        indexStatus: 'deleted',
        chunkCount: 0,
        deletedChunks,
        lastIndexedAt: new Date().toISOString()
      });
      report.push({ id: documentId, action: 'deleted', deletedChunks });
      logProgress('delta-item-delete-completed', {
        ...itemProgress,
        documentId,
        deletedChunks
      });
      continue;
    }

    if (!driveDeltaItem.file) {
      report.push({ id: driveItemId, action: 'ignored', reason: 'not_file' });
      logProgress('delta-item-ignored', { ...itemProgress, reason: 'not_file' });
      continue;
    }

    logProgress('delta-item-list-metadata-started', itemProgress);
    const listItem = await withRetry(() => fetchListItemByDriveItemId(env, token, driveItemId), { retries: 6, baseDelayMs: 1500, label: `fetch list item ${driveItemId}` });
    if (!listItem) {
      report.push({ id: driveItemId, action: 'ignored', reason: 'no_list_item' });
      logProgress('delta-item-ignored', { ...itemProgress, reason: 'no_list_item' });
      continue;
    }

    const metadataRecord = mapDocumentMetadataRecord(env, listItem);
    metadataRecord.id = metadataRecord.id || driveItemId;

    if (!isActive(metadataRecord.status)) {
      logProgress('delta-item-non-active-delete-started', {
        ...itemProgress,
        fileName: metadataRecord.fileName,
        status: metadataRecord.status
      });
      const signature = sourceSignature(listItem);
      const deletion = await deleteDocumentArtifacts(env, metadataRecord.id);
      rememberSourceItem(inventory, {
        ...inventoryIdentity(metadataRecord, listItem),
        sourceSignature: signature,
        metadataSignature: signature,
        contentSignature: contentSignature(listItem),
        driveItemId,
        fileName: metadataRecord.fileName,
        status: metadataRecord.status,
        indexStatus: 'non_active_deleted',
        skipReason: 'non_active_status',
        chunkCount: 0,
        deletedChunks: deletion.deletedChunks,
        deletedDocuments: deletion.deletedDocuments,
        lastIndexedAt: new Date().toISOString()
      });
      report.push({ fileName: metadataRecord.fileName, action: 'non_active_deleted', deletedChunks: deletion.deletedChunks, deletedDocuments: deletion.deletedDocuments });
      logProgress('delta-item-non-active-delete-completed', {
        ...itemProgress,
        fileName: metadataRecord.fileName,
        status: metadataRecord.status,
        deletedChunks: deletion.deletedChunks,
        deletedDocuments: deletion.deletedDocuments
      });
      continue;
    }

    const existing = inventory.sourceItems[metadataRecord.id];
    if (hasSameContent(existing, listItem)) {
      report.push(await updateMetadataForUnchangedContent(env, inventory, existing, metadataRecord, listItem));
      logProgress('delta-item-metadata-only-update-completed', {
        ...itemProgress,
        fileName: metadataRecord.fileName
      });
      continue;
    }

    documentsToUpload.push(metadataRecord);

    try {
      logProgress('delta-item-download-started', {
        ...itemProgress,
        fileName: metadataRecord.fileName,
        mimeType: driveDeltaItem.file.mimeType || ''
      });
      const buffer = await withRetry(() => downloadDriveItem(env, token, driveItemId), { retries: 6, baseDelayMs: 1500, label: `download ${metadataRecord.fileName}` });
      logProgress('delta-item-download-completed', {
        ...itemProgress,
        fileName: metadataRecord.fileName,
        bytes: buffer.length
      });
      logProgress('delta-item-extraction-started', {
        ...itemProgress,
        fileName: metadataRecord.fileName
      });
      const extraction = await extractDocumentText(env, metadataRecord.fileName, driveDeltaItem.file.mimeType, buffer);
      logProgress('delta-item-extraction-completed', {
        ...itemProgress,
        fileName: metadataRecord.fileName,
        method: extraction.method,
        textChars: extraction.text ? extraction.text.length : 0,
        reason: extraction.reason || ''
      });
      if (!extraction.text || extraction.text.trim().length < 50) {
        const signature = sourceSignature(listItem);
        const deletedChunks = await deleteChunksForDocument(env, metadataRecord.id);
        rememberSourceItem(inventory, {
          ...inventoryIdentity(metadataRecord, listItem),
          sourceSignature: signature,
          metadataSignature: signature,
          contentSignature: contentSignature(listItem),
          driveItemId,
          fileName: metadataRecord.fileName,
          status: metadataRecord.status,
          indexStatus: 'metadata_only',
          skipReason: extraction.reason || extraction.method,
          chunkCount: 0,
          deletedChunks,
          lastIndexedAt: new Date().toISOString()
        });
        report.push({ fileName: metadataRecord.fileName, action: 'metadata_only', reason: extraction.reason || extraction.method, deletedChunks });
        logProgress('delta-item-metadata-only-completed', {
          ...itemProgress,
          fileName: metadataRecord.fileName,
          reason: extraction.reason || extraction.method,
          deletedChunks
        });
        continue;
      }

      logProgress('delta-item-chunking-started', {
        ...itemProgress,
        fileName: metadataRecord.fileName
      });
      const documentRecord = mapDocumentRecord(env, listItem, extraction.text);
      const chunkRecords = await buildChunkRecords(env, documentRecord, extraction.text);
      if (chunkRecords.tooManyChunks) {
        const signature = sourceSignature(listItem);
        const deletedChunks = await deleteChunksForDocument(env, metadataRecord.id);
        metadataOnlyDocuments.push(metadataRecord);
        rememberSourceItem(inventory, {
          ...inventoryIdentity(metadataRecord, listItem),
          sourceSignature: signature,
          metadataSignature: signature,
          contentSignature: contentSignature(listItem),
          driveItemId,
          fileName: metadataRecord.fileName,
          status: metadataRecord.status,
          indexStatus: 'metadata_only',
          skipReason: 'too_many_chunks',
          chunkCount: 0,
          deletedChunks,
          lastIndexedAt: new Date().toISOString()
        });
        report.push({
          fileName: metadataRecord.fileName,
          action: 'metadata_only',
          reason: 'too_many_chunks',
          chunkCount: chunkRecords.chunkCount,
          chunkLimit: chunkRecords.chunkLimit,
          deletedChunks
        });
        logProgress('delta-item-metadata-only-completed', {
          ...itemProgress,
          fileName: metadataRecord.fileName,
          reason: 'too_many_chunks',
          chunkCount: chunkRecords.chunkCount,
          chunkLimit: chunkRecords.chunkLimit,
          deletedChunks
        });
        continue;
      }
      const deletedChunks = await deleteChunksForDocument(env, documentRecord.id);
      documentsToUpload[documentsToUpload.length - 1] = documentRecord;
      chunksToUpload.push(...chunkRecords);
      const signature = sourceSignature(listItem);
      rememberSourceItem(inventory, {
        ...inventoryIdentity(documentRecord, listItem),
        sourceSignature: signature,
        metadataSignature: signature,
        contentSignature: contentSignature(listItem),
        driveItemId,
        fileName: documentRecord.fileName,
        status: documentRecord.status,
        indexStatus: 'indexed',
        extractionMethod: extraction.method,
        chunkCount: chunkRecords.length,
        deletedChunks,
        lastIndexedAt: new Date().toISOString()
      });
      report.push({ fileName: documentRecord.fileName, action: 'indexed', chunks: chunkRecords.length, method: extraction.method, deletedChunks });
      logProgress('delta-item-index-records-prepared', {
        ...itemProgress,
        fileName: documentRecord.fileName,
        chunks: chunkRecords.length,
        method: extraction.method,
        deletedChunks
      });
    } catch (error) {
      const signature = sourceSignature(listItem);
      const deletedChunks = await deleteChunksForDocument(env, metadataRecord.id);
      rememberSourceItem(inventory, {
        ...inventoryIdentity(metadataRecord, listItem),
        sourceSignature: signature,
        metadataSignature: signature,
        contentSignature: contentSignature(listItem),
        driveItemId,
        fileName: metadataRecord.fileName,
        status: metadataRecord.status,
        indexStatus: 'failed',
        lastError: error.message,
        chunkCount: 0,
        deletedChunks,
        lastIndexedAt: new Date().toISOString()
      });
      report.push({ fileName: metadataRecord.fileName, action: 'failed', reason: error.message, deletedChunks });
      logProgress('delta-item-failed', {
        ...itemProgress,
        fileName: metadataRecord.fileName,
        reason: error.message
      });
    }
  }

  if (documentsToUpload.length > 0) {
    logProgress('delta-upload-documents-started', { documents: documentsToUpload.length });
    await withRetry(() => uploadDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, documentsToUpload), { retries: 5, baseDelayMs: 1500, label: 'upload delta documents' });
    logProgress('delta-upload-documents-completed', { documents: documentsToUpload.length });
  }
  if (chunksToUpload.length > 0) {
    logProgress('delta-upload-chunks-started', { chunks: chunksToUpload.length });
    await withRetry(() => uploadDocuments(env, env.AZURE_SEARCH_CHUNKS_INDEX, chunksToUpload), { retries: 5, baseDelayMs: 1500, label: 'upload delta chunks' });
    logProgress('delta-upload-chunks-completed', { chunks: chunksToUpload.length });
  }
  if (deltaLink) {
    inventory.syncSources[sourceKey] = {
      ...(inventory.syncSources[sourceKey] || {}),
      deltaLink,
      pendingDeltaNextLink: '',
      lastSyncCompletedAt: new Date().toISOString(),
      lastSyncStatus: 'success'
    };
  } else if (pendingDeltaNextLink) {
    inventory.syncSources[sourceKey] = {
      ...(inventory.syncSources[sourceKey] || {}),
      pendingDeltaNextLink,
      lastSyncCompletedAt: new Date().toISOString(),
      lastSyncStatus: 'partial'
    };
  }
  if (listDeltaEnabled && listDelta.deltaLink) {
    inventory.syncSources[listSourceKey] = {
      ...(inventory.syncSources[listSourceKey] || {}),
      deltaLink: listDelta.deltaLink,
      pendingDeltaNextLink: '',
      lastSyncCompletedAt: new Date().toISOString(),
      lastSyncStatus: listDelta.bootstrapped ? 'bootstrapped' : 'success'
    };
  } else if (listDeltaEnabled && listDelta.pendingDeltaNextLink) {
    inventory.syncSources[listSourceKey] = {
      ...(inventory.syncSources[listSourceKey] || {}),
      pendingDeltaNextLink: listDelta.pendingDeltaNextLink,
      lastSyncCompletedAt: new Date().toISOString(),
      lastSyncStatus: 'partial'
    };
  }
  saveInventory(inventory);

  console.log(JSON.stringify({
    changedItems: changedItems.length,
    driveChangedItems: driveChangedItems.length,
    listChangedItems: listDelta.changedItems.length,
    listDeltaBootstrapped: Boolean(listDelta.bootstrapped),
    uploaded: {
      documents: documentsToUpload.length,
      chunks: chunksToUpload.length
    },
    hasNewDeltaLink: Boolean(deltaLink),
    hasPendingDeltaNextLink: Boolean(pendingDeltaNextLink),
    inventory: summarizeInventory(inventory),
    report
  }, null, 2));
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
