# Search, Chatbot, and Upload AI Test Runbook

Last updated: May 26, 2026

## Current Gate Status

The backend search/chat code is in a testable state, but do not deploy to Azure production until the frontend test matrix passes and the hosted App Service settings are confirmed.

Already passed locally:

- Search backend build.
- Search API smoke/evaluation.
- Broad corpus red-team run against live indexed content.
- Upload AI document/media backend smoke.
- Frontend large MP4 path reached backend and returned `200` for transcription, transcript analysis, and metadata extraction.

Still required before Azure deployment:

- Client/front-end test matrix pass with screenshots and backend correlation IDs for failures.
- App Service app settings configured with secrets only in Azure, not SPFx.
- Easy Auth enabled and verified.
- Application Insights enabled.
- Search/chat release gate run after final merge.

## Local Backend

Run this in PowerShell from the repo root:

```powershell
Get-NetTCPConnection -LocalPort 7072,7073 -ErrorAction SilentlyContinue

$env:SEARCH_API_ALLOWED_ORIGINS="https://indegene123.sharepoint.com"
$env:SYNC_WORKER_ENABLED="false"
$env:SYNC_TIMER_ENABLED="false"
$env:SEARCH_CHAT_MEMORY_ENABLED="false"
$env:SEARCH_CHAT_RATE_LIMIT_ENABLED="false"

npm run search:local-api
```

Expected backend line:

```text
iKnowledge search API listening on port 7072
```

Verify from another PowerShell:

```powershell
Invoke-RestMethod http://localhost:7072/healthz
```

## Local Frontend

Run this in a second PowerShell from the repo root:

```powershell
npx gulp serve
```

Open the SharePoint workbench URL that gulp prints. It should contain:

```text
debugManifestsFile=https://localhost:4321/temp/build/manifests.js
```

When that debug manifest is present, the frontend automatically uses:

```text
http://localhost:7072
```

No devtools override is required.

Manual override options, only if needed:

```text
?ikSearchApi=http://localhost:7072
```

or browser console:

```javascript
localStorage.setItem('IKNOWLEDGE_SEARCH_API_BASE_URL', 'http://localhost:7072');
```

Clear override:

```javascript
localStorage.removeItem('IKNOWLEDGE_SEARCH_API_BASE_URL');
sessionStorage.removeItem('IKNOWLEDGE_SEARCH_API_BASE_URL');
```

## What To Watch

Backend terminal should show `200` requests for:

```text
/api/search/documents
/api/search/semantic
/api/search/people
/api/chat
/api/upload/extract-metadata
/api/upload/analyze-transcript
/api/upload/transcribe-media
```

For failures, copy:

- Query or action.
- Screenshot.
- Backend log line with `correlationId`.
- Actual response behavior.

## Green Flag Criteria

Give the backend deployment green flag only when all are true:

- `npm run search:build` passes.
- `npm run search:release:gate` passes.
- Frontend search tests pass for exact, semantic, metadata, file type, author, and no-result scenarios.
- Chatbot tests pass for document Q&A, Who's Who, follow-ups, off-topic handling, unsafe action refusal, prompt injection, and no inline `[1]` citations.
- Upload AI tests pass for document metadata extraction and MP4/MP3 transcription/analysis.
- Known limitations are accepted: old MP4 files without transcript chunks can be found by metadata but cannot be deeply summarized from transcript evidence.
- Azure App Service has secrets in app settings, not frontend bundles.
- Easy Auth and CORS are configured for the SharePoint origin.

## Azure Deployment Prep

Before deploying backend ZIP:

```powershell
npm run search:env:check
npm run search:build
npm run search:release:gate
```

Recommended hosted settings:

```text
SEARCH_API_ALLOWED_ORIGINS=https://indegene123.sharepoint.com
SEARCH_API_REQUIRE_EASY_AUTH=true
SYNC_WORKER_ENABLED=false
SYNC_TIMER_ENABLED=false
SEARCH_CHAT_MEMORY_ENABLED=true
SEARCH_CHAT_RATE_LIMIT_ENABLED=true
```

Keep all Azure OpenAI, Whisper, Graph, and Azure Search keys in App Service configuration or Key Vault-backed settings. Do not put them in SPFx code, package assets, or browser runtime config.

