# Search Backend Foundation

This folder is intentionally outside `src/` so none of the Graph, Azure AI
Search, Azure OpenAI, or Document Intelligence secrets can be bundled into the
SPFx frontend.

## Phase 1 Scope

The first coding phase is backend foundation only:

- validate search environment configuration
- define the three Azure AI Search indexes
- define safe query contracts and filter inputs
- keep SharePoint/SPFx frontend behavior unchanged

## Indexes

- `km-documents-index`: one row per SharePoint file for metadata search, filters,
  facets, and result cards.
- `km-chunks-index`: one row per extracted chunk for hybrid vector + keyword
  retrieval and chatbot grounding.
- `km-people-index`: one row per Who's Who list item for people and org answers.

## Search Strategy

Document semantic search uses Azure AI Search hybrid search:

1. Embed the query with the same embedding deployment used during ingestion.
2. Send a text query plus `vectorQueries` against `chunkVector`.
3. Keep `k=50` when semantic ranking is enabled so the semantic ranker has the
   recommended number of candidates.
4. Always apply `status eq 'Active'` to document/chunk queries.
5. Deduplicate chunk hits by `documentId` before returning document cards.

## Extraction Strategy

Use `EXTRACTION_STRATEGY=local-first` initially:

- local libraries for born-digital Office/text files where reliable
- Document Intelligence fallback for scanned/OCR/layout-heavy files
- media skipped in phase 1 unless a transcript pipeline is added

## Local Validation

Use:

```powershell
node tools/search-backend/scripts/validate-env.js config/search.env
```

The script reports missing values without printing secrets.

For the maintained list of active scripts, deployment contents, and archived
legacy scripts, see `SCRIPT_INVENTORY.md`.
