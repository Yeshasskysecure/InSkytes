# Search Backend Readiness

This repo is currently an SPFx frontend plus a separate search/backend tooling
folder under `tools/search-backend`.

## Current State

- Graph app-only auth works.
- Azure AI Search indexes exist:
  - `km-documents-index`
  - `km-chunks-index`
  - `km-people-index`
- Azure OpenAI embeddings work.
- Document Intelligence extraction works for PDFs.
- The current library has 577 items: 576 Active, 1 Under Review.
- Sample sync indexed active document chunks and kept media metadata-only.
- Full-library local backfill command exists. It uses the compiled TypeScript
  runtime modules, streams SharePoint pages, checkpoints local inventory, skips
  unchanged files, writes a per-run log, and retries transient network calls.
- Delta sync foundation exists and stores progress in local inventory.
- Frontend search/OpenAI hardcoded secrets have been removed from
  `SearchConfig.ts` and the known page-level OpenAI config constants.

## Important Security Reality

SPFx packages are client-side JavaScript. Installing the `.sppkg` through an
admin-controlled app catalog does not make bundled secrets private. Any user who
can load the page can inspect downloaded JavaScript and network calls.

Never put these in SPFx:

- Graph client secret
- Azure Search admin key
- Azure OpenAI key
- Document Intelligence key

The safe pattern is:

```text
SPFx UI -> backend/search API -> Azure services
```

If a backend deployment is not approved yet, the ingestion/sync scripts can run
from a scheduled runner/admin machine, but the chatbot and semantic search API
still need a safe server-side runtime before production.

The frontend now has a safe API client contract in
`src/webparts/migration/services/KnowledgeSearchApiClient.ts`. It only calls a
configured backend URL and does not carry Azure keys.

## Commands

```powershell
npm run search:env:check
npm run search:inspect
npm run search:analyze
npm run search:indexes:reset
npm run search:sync:sample
npm run search:backfill
npm run search:evaluate
npm run search:delta:init
npm run search:delta
```

Frontend-equivalent query checks:

```powershell
npm run search:documents -- "*" 5
npm run search:semantic -- "CMMI Appraisal Results Record Indegene Limited" 5
npm run search:people -- "Omnichannel Activation" 5
npm run search:chat -- "Who leads Omnichannel Activation?"
```

## CRUD Freshness

For SharePoint CRUD correctness:

- Initial sync indexes current content.
- Graph delta sync detects changes/deletions.
- Changed active files are re-downloaded, extracted, chunked, embedded, and uploaded.
- Changed non-active files keep metadata but have chunks deleted.
- Deleted files are removed from documents and chunks indexes.

Webhooks are optional. They are useful for near-real-time wake-up, but delta sync
is still the source of truth.

## Chatbot Retrieval

The chatbot should not call OpenAI directly from SPFx. Use a server-side API that:

1. Classifies the question as document, people, or mixed.
2. Retrieves active document chunks from `km-chunks-index`.
3. Retrieves people/team data from `km-people-index` when needed.
4. Builds numbered grounding sources.
5. Calls Azure OpenAI with strict "answer only from sources" instructions.
6. Returns the answer plus citations to SPFx.

`npm run search:chat -- "<question>"` exercises this flow locally.

## Full Library Chunking

Use this sequence for a full local backfill:

```powershell
npm run search:env:check
npm run search:indexes:reset
npm run search:backfill
npm run search:evaluate
```

Current extraction behavior:

- Active documents get metadata plus chunks.
- Non-active documents get metadata only and any existing chunks are deleted.
- PDFs, Office files, text, HTML, CSV, and Excel use local extraction where
  possible, with Document Intelligence fallback.
- PPTX is intentionally supported through Document Intelligence Read/Layout
  APIs. Microsoft documents Office format support for DOCX, XLSX, PPTX, and
  HTML in Document Intelligence v4.0.
- Media extensions from `EXTRACTION_SKIP_MEDIA_EXTENSIONS` are metadata-only
  until a transcript pipeline is approved. The default is
  `mp4,m4v,mov,avi,mkv,wmv,mp3,wav,m4a`.
- Extensions from `EXTRACTION_METADATA_ONLY_EXTENSIONS` are intentionally
  metadata-only. The default is `msg,zip,sketch` because Document Intelligence
  rejected `.msg` with unsupported media type in live testing, and zip/sketch
  do not have a reliable text extractor in the current pipeline.
- Media and oversized files are skipped before binary download so large videos
  do not stall the backfill just to become metadata-only records.
- Backfill uses controlled item-level concurrency via `SYNC_CONCURRENCY`
  instead of a fully serial loop. Keep the default conservative and increase
  only after checking Graph, Document Intelligence, OpenAI, and Search throttling
  metrics.
- Embedding calls are globally throttled with `SYNC_EMBEDDING_MIN_DELAY_MS`.
  Live 200-item testing hit Azure OpenAI S0 embedding 429s, so the throttle is
  intentional. Increase Azure OpenAI quota before raising this too aggressively.
- Backfill uses `SYNC_EMBEDDING_BATCH_SIZE` separately from the general OpenAI
  batch size so large PDFs do not send too many chunk inputs in one embedding
  request under low S0 quota.
- Chunk text is embedded with the configured Azure OpenAI embedding deployment
  and stored in `km-chunks-index`.

This local full sync is good enough for controlled backfills, but it will be
slow because correctness is prioritized over throughput. For 15k-60k production
scale, the same logic should run from a scheduled backend/worker with retry,
throttling, central inventory, and monitoring.

For a long PowerShell run, prefer:

```powershell
npm run search:backfill *> .\tools\search-backend\.inventory\backfill-console.log
```

The script also writes structured per-item logs under:

```text
tools/search-backend/.inventory/runs/*.ndjson
```
