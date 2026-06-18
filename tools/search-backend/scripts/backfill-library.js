const fs = require('fs');
const path = require('path');
const { loadEnv, toNumber } = require('../dist/runtime/env');
const { chunkText } = require('../dist/runtime/chunker');
const { deleteDocuments, searchIndex, uploadDocuments } = require('../dist/runtime/azureSearchClient');
const { extractDocumentText } = require('../dist/runtime/extractors');
const { downloadDriveItem, fetchLibraryItemsPage, fetchPeopleItems, getGraphToken } = require('../dist/runtime/graphClient');
const { loadInventory, rememberSourceItem, saveInventory, summarizeInventory } = require('./lib/inventory');
const { mapDocumentMetadataRecord, mapDocumentRecord, mapPeopleRecords } = require('../dist/runtime/mappers');
const { embedTexts } = require('../dist/runtime/openAiClient');
const { withRetry } = require('../dist/runtime/retry');

const envPath = process.argv[2] || 'config/search.env';
const maxItemsArg = process.argv[3] || 'all';
const maxItems = String(maxItemsArg).toLowerCase() === 'all'
  ? Number.MAX_SAFE_INTEGER
  : toNumber(maxItemsArg, 25);
const force = process.argv.includes('--force');
let nextEmbeddingRequestAt = 0;
let stopRequested = false;
const GRAPH_TOKEN_REFRESH_AFTER_MS = 45 * 60 * 1000;
const DEFAULT_METADATA_ONLY_EXTENSIONS = [
  'msg',
  'eml',
  'zip',
  'sketch',
  'idml',
  'indd',
  'ai',
  'ase',
  'psd',
  'eps',
  'ttf',
  'otf',
  'woff',
  'woff2',
  'lst',
  'thmx'
];

const isActive = (status) => String(status || '').toLowerCase() === 'active';
const resolveControlPath = (value, fallbackName) => path.resolve(value || path.join('tools', 'search-backend', '.inventory', fallbackName));
const setupStopHandlers = () => {
  const requestStop = (signal) => {
    stopRequested = true;
    console.log(`\n${signal} received. Finishing in-flight item(s), saving inventory, then stopping safely...`);
  };
  process.once('SIGINT', () => requestStop('SIGINT'));
  process.once('SIGTERM', () => requestStop('SIGTERM'));
};
const extensionFromName = (fileName) => {
  const match = String(fileName || '').match(/\.([^.]+)$/);
  return match ? match[1].toLowerCase() : '';
};
const skippedMediaExtensions = (env) => String(env.EXTRACTION_SKIP_MEDIA_EXTENSIONS || 'mp4,m4v,mov,avi,mkv,wmv,mp3,wav,m4a')
  .split(',')
  .map((item) => item.trim().toLowerCase())
  .filter(Boolean);
const metadataOnlyExtensions = (env) => String(env.EXTRACTION_METADATA_ONLY_EXTENSIONS || DEFAULT_METADATA_ONLY_EXTENSIONS.join(','))
  .split(',')
  .map((item) => item.trim().toLowerCase())
  .filter(Boolean);
const isSkippedMediaFile = (env, fileName) => skippedMediaExtensions(env).includes(extensionFromName(fileName));
const isMetadataOnlyFile = (env, fileName) => {
  const extension = extensionFromName(fileName);
  return Boolean(extension && metadataOnlyExtensions(env).includes(extension));
};
const maxExtractionBytes = (env) => Number(env.EXTRACTION_MAX_FILE_MB || 100) * 1024 * 1024;
const maxChunksPerDocument = (env) => toNumber(env.SYNC_MAX_CHUNKS_PER_DOCUMENT, 0);
const isExpiredGraphTokenError = (error) => {
  const message = error instanceof Error ? error.message : String(error);
  return message.includes('Graph GET failed 401')
    && (message.includes('InvalidAuthenticationToken') || message.includes('token is expired') || message.includes('Lifetime validation failed'));
};
const createGraphTokenProvider = (env, logStream) => {
  let token = '';
  let tokenAcquiredAt = 0;
  let refreshPromise = null;

  const refresh = async (reason) => {
    if (logStream) {
      writeLog(logStream, { action: 'graph_token_refresh', reason });
    }
    token = await getGraphToken(env);
    tokenAcquiredAt = Date.now();
    return token;
  };

  return async (options = {}) => {
    const forceRefresh = Boolean(options.force);
    const reason = options.reason || (forceRefresh ? 'forced' : 'expired_or_missing');
    const isFresh = token && (Date.now() - tokenAcquiredAt) < GRAPH_TOKEN_REFRESH_AFTER_MS;
    if (!forceRefresh && isFresh) {
      return token;
    }

    if (!refreshPromise) {
      refreshPromise = refresh(reason).finally(() => {
        refreshPromise = null;
      });
    }

    return refreshPromise;
  };
};
const withFreshGraphToken = async (getGraphTokenForRun, operation) => {
  try {
    return await operation(await getGraphTokenForRun());
  } catch (error) {
    if (!isExpiredGraphTokenError(error)) {
      throw error;
    }
    return operation(await getGraphTokenForRun({ force: true, reason: 'graph_401_expired_token' }));
  }
};
const shouldKeepMetadataOnlyAfterExtractionFailure = (reason) => {
  const message = String(reason || '').toLowerCase();
  return message.includes('unsupportedcontent')
    || message.includes('unsupported media type')
    || message.includes('unsupportedmediatype')
    || message.includes('invalidcontent')
    || message.includes('corrupt')
    || message.includes('password')
    || message.includes('could not find')
    || message.includes('not a zip file')
    || message.includes('polling timed out');
};
const waitForEmbeddingSlot = async (env) => {
  const minDelayMs = Number(env.SYNC_EMBEDDING_MIN_DELAY_MS || 15000);
  if (!Number.isFinite(minDelayMs) || minDelayMs <= 0) return;

  const now = Date.now();
  const waitMs = Math.max(0, nextEmbeddingRequestAt - now);
  nextEmbeddingRequestAt = Math.max(now, nextEmbeddingRequestAt) + minDelayMs;
  if (waitMs > 0) {
    await new Promise((resolve) => setTimeout(resolve, waitMs));
  }
};
const graphRetryOptions = (env, label) => ({
  label,
  retries: toNumber(env.SYNC_GRAPH_RETRIES, 10),
  baseDelayMs: toNumber(env.SYNC_GRAPH_RETRY_BASE_DELAY_MS, 3000),
  maxDelayMs: toNumber(env.SYNC_GRAPH_RETRY_MAX_DELAY_MS, 180000)
});
const embedInputsForBackfill = async (env, fileName, inputs) => {
  const batchSize = Math.max(1, Math.min(toNumber(env.SYNC_EMBEDDING_BATCH_SIZE, 5), inputs.length || 1));
  const embeddings = [];

  for (let index = 0; index < inputs.length; index += batchSize) {
    const batch = inputs.slice(index, index + batchSize);
    const batchEmbeddings = await withRetry(
      async () => {
        await waitForEmbeddingSlot(env);
        return embedTexts(env, batch);
      },
      { label: `embed chunks for ${fileName} batch ${Math.floor(index / batchSize) + 1}`, retries: 8, baseDelayMs: 10000 }
    );
    embeddings.push(...batchEmbeddings);
  }

  return embeddings;
};

const ensureLogStream = () => {
  const logDir = path.resolve('tools', 'search-backend', '.inventory', 'runs');
  fs.mkdirSync(logDir, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const logPath = path.join(logDir, `backfill-${stamp}.ndjson`);
  return { logPath, stream: fs.createWriteStream(logPath, { flags: 'a' }) };
};

const writeLog = (stream, item) => {
  stream.write(`${JSON.stringify({ at: new Date().toISOString(), ...item })}\n`);
};

const formatDurationMs = (durationMs) => {
  if (!Number.isFinite(durationMs)) return 'n/a';
  if (durationMs < 1000) return `${durationMs}ms`;
  return `${(durationMs / 1000).toFixed(1)}s`;
};

const logItemCompletion = (stats, maxItems, result, fallbackName) => {
  const displayMax = maxItems === Number.MAX_SAFE_INTEGER ? 'all' : maxItems;
  const fileName = result.fileName || fallbackName || result.id || '(unknown item)';
  const parts = [`action=${result.action}`];

  if (result.action === 'indexed') {
    parts.push(`chunks=${result.chunks || 0}`);
    parts.push(`method=${result.method || 'n/a'}`);
  }

  if (result.action === 'metadata_only') {
    parts.push(`chunks=0`);
    parts.push(`reason=${result.reason || 'n/a'}`);
  }

  if (result.action === 'failed' || result.action === 'failed_unhandled') {
    parts.push(`reason=${result.reason || 'n/a'}`);
  }

  if (Number.isFinite(result.durationMs)) {
    parts.push(`duration=${formatDurationMs(result.durationMs)}`);
  }

  console.log(`Completed ${stats.processed}/${displayMax}: ${fileName} -> ${parts.join('; ')}`);
  if (result.webUrl) {
    console.log(`  Link: ${result.webUrl}`);
  }
};

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
    fileUniqueId: record.fileUniqueId || record.uniqueId || '',
    fileRef: record.fileRef || record.serverRelativeUrl || '',
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
  const results = await searchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
    search: '*',
    filter: `documentId eq '${String(documentId).replace(/'/g, "''")}'`,
    select: 'id',
    top: 1000
  });

  return (results.value || []).map((item) => item.id).filter(Boolean);
};

const deleteChunksForDocument = async (env, documentId) => {
  const results = await searchIndex(env, env.AZURE_SEARCH_CHUNKS_INDEX, {
    search: '*',
    filter: `documentId eq '${String(documentId).replace(/'/g, "''")}'`,
    select: 'id',
    top: 1000
  });

  const ids = (results.value || []).map((item) => item.id).filter(Boolean);
  await deleteDocuments(env, env.AZURE_SEARCH_CHUNKS_INDEX, 'id', ids);
  return ids.length;
};

const updateMetadataForUnchangedContent = async (env, inventory, existing, metadataRecord, item, signature, logStream, startedAt) => {
  const chunkIds = existing.indexStatus === 'indexed'
    ? await withRetry(
      () => findChunkIdsForDocument(env, metadataRecord.id),
      { label: `lookup chunks for metadata update ${metadataRecord.fileName}` }
    )
    : [];

  await withRetry(
    () => uploadDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, [documentMetadataPatch(metadataRecord)]),
    { label: `update document metadata ${metadataRecord.fileName}` }
  );

  if (chunkIds.length > 0) {
    await withRetry(
      () => uploadDocuments(env, env.AZURE_SEARCH_CHUNKS_INDEX, chunkIds.map((id) => chunkMetadataPatch(id, metadataRecord))),
      { label: `update chunk metadata ${metadataRecord.fileName}` }
    );
  }

  rememberSourceItem(inventory, {
    ...inventoryIdentity(metadataRecord, item),
    sourceSignature: signature,
    metadataSignature: signature,
    contentSignature: contentSignature(item),
    fileName: metadataRecord.fileName,
    webUrl: metadataRecord.webUrl,
    status: metadataRecord.status,
    indexStatus: existing.indexStatus,
    chunkCount: existing.chunkCount || chunkIds.length,
    lastMetadataSyncedAt: new Date().toISOString(),
    lastIdentitySyncedAt: new Date().toISOString()
  });

  writeLog(logStream, {
    action: 'metadata_updated',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    chunksUpdated: chunkIds.length,
    durationMs: Date.now() - startedAt
  });

  return {
    action: 'metadata_updated',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    webUrl: metadataRecord.webUrl,
    chunksUpdated: chunkIds.length,
    chunks: 0,
    durationMs: Date.now() - startedAt
  };
};

const updateIdentityForUnchangedContent = async (env, inventory, existing, metadataRecord, item, signature, logStream, startedAt) => {
  await withRetry(
    () => uploadDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, [documentMetadataPatch(metadataRecord)]),
    { label: `update document identity metadata ${metadataRecord.fileName}` }
  );

  rememberSourceItem(inventory, {
    ...inventoryIdentity(metadataRecord, item),
    sourceSignature: signature,
    metadataSignature: existing?.metadataSignature || signature,
    contentSignature: existing?.contentSignature || contentSignature(item),
    fileName: metadataRecord.fileName,
    webUrl: metadataRecord.webUrl,
    status: metadataRecord.status,
    indexStatus: existing?.indexStatus || 'indexed',
    extractionMethod: existing?.extractionMethod,
    chunkCount: existing?.chunkCount || 0,
    deletedChunks: existing?.deletedChunks || 0,
    lastIndexedAt: existing?.lastIndexedAt,
    lastMetadataSyncedAt: existing?.lastMetadataSyncedAt,
    lastIdentitySyncedAt: new Date().toISOString()
  });

  writeLog(logStream, {
    action: 'identity_metadata_updated',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    durationMs: Date.now() - startedAt
  });

  return {
    action: 'identity_metadata_updated',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    webUrl: metadataRecord.webUrl,
    chunks: 0,
    durationMs: Date.now() - startedAt
  };
};

const deleteDocumentArtifacts = async (env, documentId) => {
  const deletedChunks = await deleteChunksForDocument(env, documentId);
  await deleteDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, 'id', [String(documentId)]);
  return { deletedChunks, deletedDocuments: 1 };
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
  const embeddings = await embedInputsForBackfill(env, documentRecord.fileName, inputs);
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

const shouldSkipUnchanged = (inventory, documentId, signature, status) => {
  const existing = inventory.sourceItems[documentId];
  if (!isActive(status) && existing && existing.indexStatus !== 'non_active_deleted') {
    return false;
  }
  return Boolean(existing && existing.sourceSignature === signature && existing.indexStatus !== 'failed');
};

const needsMetadataRefresh = (existing, signature) =>
  Boolean(
    existing
      && ['indexed', 'metadata_only'].includes(existing.indexStatus)
      && (!existing.lastMetadataSyncedAt || existing.metadataSignature !== signature)
  );

const needsIdentityRefresh = (existing) =>
  Boolean(
    existing
      && ['indexed', 'metadata_only'].includes(existing.indexStatus)
      && (!existing.lastIdentitySyncedAt || !existing.fileUniqueId || !existing.fileRef)
  );

const uploadMetadataOnlyDocument = async (env, inventory, metadataRecord, signature, reason, logStream, startedAt, extra = {}) => {
  const deletedChunks = await withRetry(
    () => deleteChunksForDocument(env, metadataRecord.id),
    { label: `delete chunks for metadata-only ${metadataRecord.fileName}` }
  );
  await withRetry(
    () => uploadDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, [metadataRecord]),
    { label: `upload metadata-only ${metadataRecord.fileName}` }
  );
  rememberSourceItem(inventory, {
    ...inventoryIdentity(metadataRecord, extra.item),
    sourceSignature: signature,
    metadataSignature: signature,
    contentSignature: extra.item ? contentSignature(extra.item) : undefined,
    fileName: metadataRecord.fileName,
    webUrl: metadataRecord.webUrl,
    status: metadataRecord.status,
    indexStatus: 'metadata_only',
    skipReason: reason,
    chunkCount: 0,
    deletedChunks,
    lastIndexedAt: new Date().toISOString(),
    lastIdentitySyncedAt: new Date().toISOString(),
    ...extra.inventory
  });
  writeLog(logStream, {
    action: 'metadata_only',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    reason,
    deletedChunks,
    durationMs: Date.now() - startedAt,
    ...extra.log
  });
  return {
    action: 'metadata_only',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    webUrl: metadataRecord.webUrl,
    reason,
    deletedChunks,
    chunks: 0,
    durationMs: Date.now() - startedAt
  };
};

const processLibraryItem = async (env, getGraphTokenForRun, inventory, item, logStream) => {
  const startedAt = Date.now();
  const signature = sourceSignature(item);
  const metadataRecord = mapDocumentMetadataRecord(env, item);
  const driveItem = item.driveItem || {};
  const existing = inventory.sourceItems[metadataRecord.id];
  writeLog(logStream, {
    action: 'start_item',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    status: metadataRecord.status,
    size: driveItem.size || 0
  });

  if (!force && shouldSkipUnchanged(inventory, metadataRecord.id, signature, metadataRecord.status)) {
    if (needsMetadataRefresh(existing, signature)) {
      return updateMetadataForUnchangedContent(env, inventory, existing, metadataRecord, item, signature, logStream, startedAt);
    }
    if (needsIdentityRefresh(existing)) {
      return updateIdentityForUnchangedContent(env, inventory, existing, metadataRecord, item, signature, logStream, startedAt);
    }

    rememberSourceItem(inventory, {
      ...inventoryIdentity(metadataRecord, item),
      sourceSignature: signature,
      metadataSignature: existing?.metadataSignature || signature,
      contentSignature: existing?.contentSignature || contentSignature(item),
      fileName: metadataRecord.fileName,
      webUrl: metadataRecord.webUrl,
      status: metadataRecord.status,
      indexStatus: existing?.indexStatus || 'indexed',
      extractionMethod: existing?.extractionMethod,
      chunkCount: existing?.chunkCount || 0,
      deletedChunks: existing?.deletedChunks || 0,
      lastIndexedAt: existing?.lastIndexedAt,
      lastMetadataSyncedAt: existing?.lastMetadataSyncedAt,
      lastIdentitySyncedAt: existing?.lastIdentitySyncedAt
    });
    writeLog(logStream, {
      action: 'skipped_unchanged',
      id: metadataRecord.id,
      fileName: metadataRecord.fileName,
      durationMs: Date.now() - startedAt
    });
    return {
      action: 'skipped_unchanged',
      id: metadataRecord.id,
      fileName: metadataRecord.fileName,
      webUrl: metadataRecord.webUrl,
      chunks: 0,
      durationMs: Date.now() - startedAt
    };
  }

  if (!isActive(metadataRecord.status)) {
    const deletion = await withRetry(
      () => deleteDocumentArtifacts(env, metadataRecord.id),
      { label: `delete non-active artifacts for ${metadataRecord.fileName}` }
    );
    rememberSourceItem(inventory, {
      ...inventoryIdentity(metadataRecord, item),
      sourceSignature: signature,
      metadataSignature: signature,
      contentSignature: contentSignature(item),
      fileName: metadataRecord.fileName,
      webUrl: metadataRecord.webUrl,
      status: metadataRecord.status,
      indexStatus: 'non_active_deleted',
      skipReason: 'non_active_status',
      chunkCount: 0,
      deletedChunks: deletion.deletedChunks,
      deletedDocuments: deletion.deletedDocuments,
      lastIndexedAt: new Date().toISOString(),
      lastIdentitySyncedAt: new Date().toISOString()
    });
    writeLog(logStream, {
      action: 'non_active_deleted',
      id: metadataRecord.id,
      fileName: metadataRecord.fileName,
      deletedChunks: deletion.deletedChunks,
      deletedDocuments: deletion.deletedDocuments,
      durationMs: Date.now() - startedAt
    });
    return {
      action: 'non_active_deleted',
      id: metadataRecord.id,
      fileName: metadataRecord.fileName,
      webUrl: metadataRecord.webUrl,
      deletedChunks: deletion.deletedChunks,
      chunks: 0,
      durationMs: Date.now() - startedAt
    };
  }

  if (!driveItem.id || !driveItem.file) {
    return uploadMetadataOnlyDocument(env, inventory, metadataRecord, signature, 'not_a_file_drive_item', logStream, startedAt, { item });
  }

  if (!force && hasSameContent(existing, item)) {
    return updateMetadataForUnchangedContent(env, inventory, existing, metadataRecord, item, signature, logStream, startedAt);
  }

  if (isSkippedMediaFile(env, metadataRecord.fileName) || isMetadataOnlyFile(env, metadataRecord.fileName)) {
    const skipReason = isSkippedMediaFile(env, metadataRecord.fileName)
      ? 'media_pipeline_not_enabled'
      : 'metadata_only_extension';
    return uploadMetadataOnlyDocument(env, inventory, metadataRecord, signature, skipReason, logStream, startedAt, {
      item,
      log: {
        skippedBeforeDownload: true,
        extension: extensionFromName(metadataRecord.fileName) || '(none)'
      }
    });
  }

  if (Number(driveItem.size || 0) > maxExtractionBytes(env)) {
    return uploadMetadataOnlyDocument(env, inventory, metadataRecord, signature, 'exceeds_size_limit', logStream, startedAt, {
      item,
      log: {
        skippedBeforeDownload: true,
        size: driveItem.size || 0
      }
    });
  }

  writeLog(logStream, {
    action: 'download_start',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    size: driveItem.size || 0
  });
  const buffer = await withRetry(
    () => withFreshGraphToken(getGraphTokenForRun, (graphToken) => downloadDriveItem(env, graphToken, driveItem.id)),
    graphRetryOptions(env, `download ${metadataRecord.fileName}`)
  );
  writeLog(logStream, {
    action: 'download_complete',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    bytes: buffer.byteLength
  });
  writeLog(logStream, {
    action: 'extraction_start',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName
  });
  const extraction = await extractDocumentText(env, metadataRecord.fileName, driveItem.file.mimeType, buffer);
  writeLog(logStream, {
    action: 'extraction_complete',
    id: metadataRecord.id,
    fileName: metadataRecord.fileName,
    method: extraction.method,
    textLength: extraction.text ? extraction.text.length : 0,
    reason: extraction.reason
  });

  if (extraction.method === 'failed') {
    if (shouldKeepMetadataOnlyAfterExtractionFailure(extraction.reason)) {
      return uploadMetadataOnlyDocument(env, inventory, metadataRecord, signature, extraction.reason || extraction.method, logStream, startedAt, {
        item,
        log: {
          extractionMethod: extraction.method,
          extractionFailureHandledAsMetadataOnly: true
        }
      });
    }

    rememberSourceItem(inventory, {
      ...inventoryIdentity(metadataRecord, item),
      sourceSignature: signature,
      metadataSignature: signature,
      contentSignature: contentSignature(item),
      fileName: metadataRecord.fileName,
      webUrl: metadataRecord.webUrl,
      status: metadataRecord.status,
      indexStatus: 'failed',
      lastError: extraction.reason || extraction.method,
      lastIndexedAt: new Date().toISOString()
    });
    writeLog(logStream, {
      action: 'failed',
      id: metadataRecord.id,
      fileName: metadataRecord.fileName,
      reason: extraction.reason || extraction.method,
      durationMs: Date.now() - startedAt
    });
    return {
      action: 'failed',
      id: metadataRecord.id,
      fileName: metadataRecord.fileName,
      webUrl: metadataRecord.webUrl,
      reason: extraction.reason || extraction.method,
      chunks: 0,
      durationMs: Date.now() - startedAt
    };
  }

  if (!extraction.text || extraction.text.trim().length < 50) {
    return uploadMetadataOnlyDocument(env, inventory, metadataRecord, signature, extraction.reason || extraction.method, logStream, startedAt, { item });
  }

  const documentRecord = mapDocumentRecord(env, item, extraction.text);
  writeLog(logStream, {
    action: 'embedding_start',
    id: documentRecord.id,
    fileName: documentRecord.fileName
  });
  const chunkRecords = await buildChunkRecords(env, documentRecord, extraction.text);
  if (chunkRecords.tooManyChunks) {
    return uploadMetadataOnlyDocument(env, inventory, metadataRecord, signature, 'too_many_chunks', logStream, startedAt, {
      item,
      log: {
        extractionMethod: extraction.method,
        chunkCount: chunkRecords.chunkCount,
        chunkLimit: chunkRecords.chunkLimit
      }
    });
  }
  writeLog(logStream, {
    action: 'embedding_complete',
    id: documentRecord.id,
    fileName: documentRecord.fileName,
    chunks: chunkRecords.length
  });
  const deletedChunks = await withRetry(
    () => deleteChunksForDocument(env, documentRecord.id),
    { label: `delete old chunks for ${documentRecord.fileName}` }
  );
  await withRetry(
    () => uploadDocuments(env, env.AZURE_SEARCH_DOCUMENTS_INDEX, [documentRecord]),
    { label: `upload document ${documentRecord.fileName}` }
  );
  await withRetry(
    () => uploadDocuments(env, env.AZURE_SEARCH_CHUNKS_INDEX, chunkRecords),
    { label: `upload chunks for ${documentRecord.fileName}` }
  );

  rememberSourceItem(inventory, {
    ...inventoryIdentity(documentRecord, item),
    sourceSignature: signature,
    metadataSignature: signature,
    contentSignature: contentSignature(item),
    fileName: documentRecord.fileName,
    webUrl: documentRecord.webUrl,
    status: documentRecord.status,
    indexStatus: 'indexed',
    extractionMethod: extraction.method,
    chunkCount: chunkRecords.length,
    deletedChunks,
    lastIndexedAt: new Date().toISOString(),
    lastIdentitySyncedAt: new Date().toISOString()
  });
  writeLog(logStream, {
    action: 'indexed',
    id: documentRecord.id,
    fileName: documentRecord.fileName,
    method: extraction.method,
    chunks: chunkRecords.length,
    deletedChunks,
    durationMs: Date.now() - startedAt
  });
  return {
    action: 'indexed',
    id: documentRecord.id,
    fileName: documentRecord.fileName,
    webUrl: documentRecord.webUrl,
    method: extraction.method,
    chunks: chunkRecords.length,
    durationMs: Date.now() - startedAt
  };
};

const uploadPeople = async (env, getGraphTokenForRun, logStream) => {
  const peopleItems = await withRetry(
    () => withFreshGraphToken(getGraphTokenForRun, (graphToken) => fetchPeopleItems(env, graphToken, toNumber(env.SYNC_PEOPLE_TOP, 500))),
    { label: 'fetch people list' }
  );
  const peopleRecords = mapPeopleRecords(peopleItems);
  await withRetry(
    () => uploadDocuments(env, env.AZURE_SEARCH_PEOPLE_INDEX, peopleRecords),
    { label: 'upload people records' }
  );
  writeLog(logStream, { action: 'people_uploaded', count: peopleRecords.length });
  return peopleRecords.length;
};

const waitForControlResume = async (pauseFile, stopFile) => {
  while (!stopRequested && fs.existsSync(pauseFile)) {
    console.log(`Pause file detected: ${pauseFile}. Waiting 30s. Delete it to resume.`);
    await new Promise((resolve) => setTimeout(resolve, 30000));
  }

  if (fs.existsSync(stopFile)) {
    stopRequested = true;
    console.log(`Stop file detected: ${stopFile}. Finishing in-flight item(s), saving inventory, then stopping safely...`);
  }
};

const processPageItems = async (env, getGraphTokenForRun, inventory, items, stats, maxItems, concurrency, checkpointEvery, pauseMs, stream, pauseFile, stopFile) => {
  let cursor = 0;

  const worker = async () => {
    while (!stopRequested && cursor < items.length && stats.processed < maxItems) {
      await waitForControlResume(pauseFile, stopFile);
      if (stopRequested) break;
      if (cursor >= items.length || stats.processed >= maxItems) break;
      const item = items[cursor];
      cursor += 1;

      try {
        const itemName = (item.fields && item.fields.FileLeafRef) || (item.driveItem && item.driveItem.name) || item.id;
        stats.started += 1;
        console.log(`Processing ${stats.started}/${maxItems === Number.MAX_SAFE_INTEGER ? 'all' : maxItems}: ${itemName}`);
        const result = await processLibraryItem(env, getGraphTokenForRun, inventory, item, stream);
        stats.processed += 1;
        stats[result.action] = (stats[result.action] || 0) + 1;
        stats.chunks += result.chunks || 0;
        logItemCompletion(stats, maxItems, result, itemName);
      } catch (error) {
        const metadataRecord = mapDocumentMetadataRecord(env, item);
        const signature = sourceSignature(item);
        rememberSourceItem(inventory, {
          ...inventoryIdentity(metadataRecord, item),
          sourceSignature: signature,
          metadataSignature: signature,
          contentSignature: contentSignature(item),
          fileName: metadataRecord.fileName,
          webUrl: metadataRecord.webUrl,
          status: metadataRecord.status,
          indexStatus: 'failed',
          lastError: error.message,
          lastIndexedAt: new Date().toISOString()
        });
        stats.processed += 1;
        stats.failed += 1;
        writeLog(stream, {
          action: 'failed_unhandled',
          id: metadataRecord.id,
          fileName: metadataRecord.fileName,
          reason: error.message
        });
        logItemCompletion(stats, maxItems, {
          action: 'failed_unhandled',
          id: metadataRecord.id,
          fileName: metadataRecord.fileName,
          webUrl: metadataRecord.webUrl,
          reason: error.message
        });
      }

      if (stats.processed % checkpointEvery === 0) {
        saveInventory(inventory);
        console.log(`Progress: ${JSON.stringify(stats)}`);
      }

      if (pauseMs > 0) {
        await new Promise((resolve) => setTimeout(resolve, pauseMs));
      }
    }
  };

  const workerCount = Math.min(Math.max(concurrency, 1), items.length);
  await Promise.all(Array.from({ length: workerCount }, worker));
};

const run = async () => {
  setupStopHandlers();
  const { env, absolutePath } = loadEnv(envPath);
  if (env.SEARCH_INVENTORY_PATH || env.SYNC_INVENTORY_PATH) {
    process.env.SEARCH_INVENTORY_PATH = env.SEARCH_INVENTORY_PATH || env.SYNC_INVENTORY_PATH;
  }
  const { inventory, absolutePath: inventoryPath } = loadInventory();
  const pageSize = Math.min(toNumber(env.SYNC_PAGE_SIZE, 50), 200);
  const concurrency = Math.min(toNumber(env.SYNC_CONCURRENCY, 3), 10);
  const checkpointEvery = toNumber(env.SYNC_CHECKPOINT_EVERY, 1);
  const pauseMs = Number(env.SYNC_PAUSE_BETWEEN_ITEMS_MS || 0);
  const pauseFile = resolveControlPath(env.SYNC_PAUSE_FILE, 'pause-backfill.flag');
  const stopFile = resolveControlPath(env.SYNC_STOP_FILE, 'stop-backfill.flag');
  const { logPath, stream } = ensureLogStream();
  const stats = {
    processed: 0,
    indexed: 0,
    metadata_only: 0,
    skipped_unchanged: 0,
    failed: 0,
    metadata_updated: 0,
    started: 0,
    chunks: 0
  };

  console.log(`Using env: ${absolutePath}`);
  console.log(`Using inventory: ${inventoryPath}`);
  console.log(`Run log: ${logPath}`);
  console.log(`Max library items: ${maxItems === Number.MAX_SAFE_INTEGER ? 'all' : maxItems}; page size: ${pageSize}; concurrency: ${concurrency}; force: ${force}`);
  console.log(`Pause by creating: ${pauseFile}`);
  console.log(`Stop safely by creating: ${stopFile} or pressing Ctrl+C once`);

  try {
    const getGraphTokenForRun = createGraphTokenProvider(env, stream);
    await getGraphTokenForRun({ force: true, reason: 'startup' });
    let nextLink = '';
    let pageNumber = 0;

    while (!stopRequested && stats.processed < maxItems) {
      await waitForControlResume(pauseFile, stopFile);
      if (stopRequested) break;
      pageNumber += 1;
      const remaining = maxItems - stats.processed;
      const page = await withRetry(
        () => withFreshGraphToken(
          getGraphTokenForRun,
          (graphToken) => fetchLibraryItemsPage(env, graphToken, Math.min(pageSize, remaining), nextLink || undefined)
        ),
        graphRetryOptions(env, `fetch library page ${pageNumber}`)
      );
      const items = page.value || [];
      if (items.length === 0) break;

      await processPageItems(env, getGraphTokenForRun, inventory, items, stats, maxItems, concurrency, checkpointEvery, pauseMs, stream, pauseFile, stopFile);

      nextLink = page['@odata.nextLink'] || '';
      if (!nextLink) break;
    }

    const peopleCount = stopRequested ? 0 : await uploadPeople(env, getGraphTokenForRun, stream);
    saveInventory(inventory);
    console.log(JSON.stringify({
      stats,
      people: peopleCount,
      stoppedEarly: stopRequested,
      inventory: summarizeInventory(inventory),
      logPath
    }, null, 2));
  } finally {
    saveInventory(inventory);
    stream.end();
  }
};

run().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
