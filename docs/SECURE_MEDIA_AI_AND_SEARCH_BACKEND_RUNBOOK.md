  # Secure Media AI and Search Backend Runbook

This runbook is for the iKnowledgeNext search/chat backend and the media upload AI feature that uses Whisper/OpenAI.

## 0. Immediate Secret Safety

The Whisper key that was pasted in chat must be treated as exposed.

Do this before production testing:

1. Open Azure Portal.
2. Go to the Azure OpenAI resource that contains the `whisper` deployment.
3. Open **Keys and Endpoint**.
4. Regenerate the key that was pasted.
5. Update App Service or Key Vault with the new value.
6. Do not paste the new key into chat, SPFx code, localStorage, browser devtools snippets, or committed files.

Why: SPFx runs in the browser. Any key placed in frontend code, runtime config, localStorage, or bundled JavaScript can be read by users.

## 1. Current State

Search and chatbot now use the backend API:

- `tools/search-backend/scripts/local-api-server.js`
- `tools/search-backend/src/api/knowledgeSearchApi.ts`
- frontend runtime base URL from `src/webparts/migration/services/SearchConfig.ts`

The media upload/video analysis path still uses the old client-side OpenAI service:

- `src/webparts/migration/components/upload/FileUpload/FileUpload.tsx`
- `src/webparts/migration/components/upload/VideoAnalysis/VideoAnalysis.tsx`
- `src/webparts/migration/services/AzureOpenAIService.ts`

That old path calls Azure OpenAI/Whisper directly from the browser with `api-key` headers. It worked only when keys were available to the frontend. We should not restore that for production.

## 2. Correct Target Architecture

Use this shape:

```text
SPFx upload UI
  -> iKnowledge backend App Service
      -> Azure OpenAI Whisper
      -> Azure OpenAI chat model
```

Frontend:

- Sends media/transcript requests only to the iKnowledge backend.
- Does not know Azure OpenAI keys.
- Does not call `*.cognitiveservices.azure.com` directly.

Backend:

- Reads secrets from App Service settings or Key Vault references.
- Calls Whisper and GPT server-side.
- Logs correlation IDs and safe metadata only.

## 3. Whisper Endpoint Choice

Your current Whisper deployment is named:

```text
whisper
```

The Azure portal URL you showed used:

```text
/audio/translations
```

Use the endpoint intentionally:

- `audio/transcriptions`: speech to text in the spoken language.
- `audio/translations`: speech to English translation.

For upload transcript generation, prefer `audio/transcriptions` unless the business explicitly wants translation to English.

Example endpoint shape:

```text
https://<resource-name>.cognitiveservices.azure.com/openai/deployments/whisper/audio/transcriptions?api-version=2024-06-01
```

Do not include the key in the URL.

## 4. App Service Settings

Run these in Azure Cloud Shell PowerShell after rotating the key.

Set variables:

```powershell
$rg = "Knowledge_Management_Enterprise"
$app = "iknowledge-search-api-prod"
```

Set non-secret Whisper settings:

```powershell
az webapp config appsettings set `
  --resource-group $rg `
  --name $app `
  --settings `
    AZURE_OPENAI_WHISPER_ENDPOINT="https://<your-whisper-resource>.cognitiveservices.azure.com" `
    AZURE_OPENAI_WHISPER_DEPLOYMENT="whisper" `
    AZURE_OPENAI_WHISPER_API_VERSION="2024-06-01" `
    AZURE_OPENAI_WHISPER_MODE="transcriptions"
```

For the key, prefer Key Vault. If Key Vault is not ready yet, use App Service app settings as a temporary step:

```powershell
az webapp config appsettings set `
  --resource-group $rg `
  --name $app `
  --settings AZURE_OPENAI_WHISPER_API_KEY="<rotated-whisper-key>"
```

Do not commit this value anywhere.

## 5. Preferred Key Vault Setup

Run in Azure Cloud Shell PowerShell.

Set variables:

```powershell
$rg = "Knowledge_Management_Enterprise"
$app = "iknowledge-search-api-prod"
$vault = "<your-key-vault-name>"
$secretName = "azure-openai-whisper-api-key"
```

Enable managed identity on the App Service:

```powershell
az webapp identity assign `
  --resource-group $rg `
  --name $app
```

Get the principal ID:

```powershell
$principalId = az webapp identity show `
  --resource-group $rg `
  --name $app `
  --query principalId `
  -o tsv
```

Store the rotated Whisper key in Key Vault:

```powershell
az keyvault secret set `
  --vault-name $vault `
  --name $secretName `
  --value "<rotated-whisper-key>"
```

Grant the App Service identity permission to read secrets. If the vault uses Azure RBAC:

```powershell
$vaultId = az keyvault show `
  --resource-group $rg `
  --name $vault `
  --query id `
  -o tsv

az role assignment create `
  --assignee $principalId `
  --role "Key Vault Secrets User" `
  --scope $vaultId
```

Set the App Service setting as a Key Vault reference:

```powershell
az webapp config appsettings set `
  --resource-group $rg `
  --name $app `
  --settings AZURE_OPENAI_WHISPER_API_KEY="@Microsoft.KeyVault(VaultName=$vault;SecretName=$secretName)"
```

Restart after setting secrets:

```powershell
az webapp restart `
  --resource-group $rg `
  --name $app
```

## 6. Backend Media AI Endpoints

The backend now includes upload AI proxy endpoints:

```text
POST /api/upload/transcribe-media
POST /api/upload/analyze-transcript
POST /api/upload/extract-metadata
```

Expected behavior:

- `/api/upload/transcribe-media`
  - Accepts multipart media upload.
  - Enforces file size and content type limits.
  - Calls Azure OpenAI Whisper from the backend using `AZURE_OPENAI_WHISPER_*`.
  - Returns transcript entries and transcript text.

- `/api/upload/analyze-transcript`
  - Accepts transcript text and optional taxonomy/document-type context.
  - Calls the backend GPT deployment using existing `AZURE_OPENAI_*` settings.
  - Returns title, description/abstract, BU, department, document type, and other upload metadata.

- `/api/upload/extract-metadata`
  - Accepts parsed document text from the existing upload flow.
  - Calls the backend GPT deployment using existing `AZURE_OPENAI_*` settings.
  - Returns metadata in the same shape the upload form expects.

The frontend still calls `AzureOpenAIService` method names, but those methods now prefer the backend API when `knowledgeSearchApiBaseUrl` is configured.

Updated frontend touchpoints:

- `src/webparts/migration/components/upload/FileUpload/FileUpload.tsx`
- `src/webparts/migration/components/upload/VideoAnalysis/VideoAnalysis.tsx`
- `src/webparts/migration/services/AzureOpenAIService.ts`
- `src/webparts/migration/services/KnowledgeSearchApiClient.ts`

The old direct browser-side OpenAI path is still present only as a temporary fallback. Do not provide browser-side keys in production.

## 7. Local Test Flow

Use two terminals.

Terminal 1, backend API on port `7072`:

```powershell
$env:SEARCH_API_ALLOWED_ORIGINS="https://indegene123.sharepoint.com"
$env:SYNC_WORKER_ENABLED="false"
$env:SYNC_TIMER_ENABLED="false"
$env:SEARCH_CHAT_MEMORY_ENABLED="false"
$env:SEARCH_CHAT_RATE_LIMIT_ENABLED="false"

npm run search:local-api
```

Terminal 2, SPFx local workbench bundle:

```powershell
npx gulp serve
```

In the SharePoint workbench browser console, force the frontend to call the local backend:

```javascript
localStorage.setItem("IKNOWLEDGE_SEARCH_API_BASE_URL", "http://localhost:7072");
location.reload();
```

Do not set OpenAI, Whisper, or Azure Search keys in browser localStorage.

Expected network behavior:

- Good: browser calls `http://localhost:7072/api/...` or hosted App Service `/api/...`.
- Bad: browser calls `https://*.cognitiveservices.azure.com/openai/...` directly.

For a no-upload backend smoke test:

```powershell
npm run search:upload-ai:smoke
```

Report:

```text
C:\tmp\iknowledge-upload-ai-smoke-report.json
```

For the large MP4 frontend test:

1. Keep Terminal 1 backend running on `7072`.
2. Keep Terminal 2 `gulp serve` running.
3. Open the SharePoint workbench.
4. Confirm the browser has `IKNOWLEDGE_SEARCH_API_BASE_URL=http://localhost:7072`.
5. Select one large MP4 in the upload UI.
6. Wait only until transcript/metadata fills.
7. Cancel before final SharePoint upload.
8. Watch the backend terminal for:

```text
POST /api/upload/transcribe-media
POST /api/upload/analyze-transcript
```

The large video path still uses browser-side audio decode/chunking first. The backend receives the already chunked WAV/audio pieces and sends them to Whisper with the server-side key.

## 8. Backend Build and Deployment

Run in repo PowerShell:

```powershell
npm run search:build
```

Create the backend production ZIP:

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\search-backend\scripts\create-prod-api-package.ps1
```

Deploy the ZIP from Azure Cloud Shell or local Azure CLI:

```powershell
$rg = "Knowledge_Management_Enterprise"
$app = "iknowledge-search-api-prod"
$zip = "C:\tmp\<created-zip-name>.zip"

az webapp deployment source config-zip `
  --resource-group $rg `
  --name $app `
  --src $zip
```

Restart:

```powershell
az webapp restart `
  --resource-group $rg `
  --name $app
```

Smoke test:

```powershell
Invoke-RestMethod "https://iknowledge-search-api-prod.azurewebsites.net/healthz"
```

## 8.1 Local Whisper Env

For local backend testing, add these to ignored `config/search.env` after rotating the Whisper key:

```text
AZURE_OPENAI_WHISPER_ENDPOINT=https://<your-whisper-resource>.cognitiveservices.azure.com
AZURE_OPENAI_WHISPER_API_KEY=<rotated-whisper-key>
AZURE_OPENAI_WHISPER_API_VERSION=2024-06-01
AZURE_OPENAI_WHISPER_DEPLOYMENT=whisper
AZURE_OPENAI_WHISPER_MODE=transcriptions
SEARCH_UPLOAD_MAX_MEDIA_BYTES=25165824
SEARCH_UPLOAD_METADATA_MAX_CHARS=120000
SEARCH_UPLOAD_TRANSCRIPT_MAX_CHARS=120000
```

If you paste the full portal target URL instead of the base endpoint, the backend accepts it only when it already contains `/audio/transcriptions` or `/audio/translations`.

## 9. Frontend Runtime Config

The SPFx frontend should know only the backend API base URL and, if Easy Auth/AAD is enabled, the AAD resource.

Example runtime config:

```javascript
window.__IKNOWLEDGE_SEARCH_CONFIG__ = {
  knowledgeSearchApiBaseUrl: "https://iknowledge-search-api-prod.azurewebsites.net",
  knowledgeSearchApiAadResource: "<app-registration-resource-or-client-id>"
};
```

Do not include:

```text
azureOpenAiApiKey
azureOpenAiWhisperApiKey
azureSearchKey
```

## 10. Cleanup Plan

Only after backend media endpoints are implemented and tested:

1. Remove browser-side OpenAI/Whisper calls from upload flows.
2. Remove runtime secret fields from `SearchConfig.ts`.
3. Remove direct Azure Search/OpenAI API-key paths that are no longer referenced.
4. Keep search/chat backend scripts needed for build, deploy, backfill, delta sync, smoke tests, and production diagnosis.
5. Do not delete `.archive/search-legacy` until the team confirms it is not needed for audit/reference.

## 11. Verification Checklist

Before saying production is ready:

- Whisper key regenerated after exposure.
- No OpenAI/Search keys in SPFx bundle.
- No OpenAI/Search keys in SharePoint page scripts.
- `config/*.env` remains ignored by git.
- Backend App Service has all required app settings.
- App Service `/healthz` returns healthy.
- Search API smoke test passes.
- Chat smoke test passes.
- Media upload transcript test passes through backend only.
- Browser network tab shows no direct `cognitiveservices.azure.com` calls from SPFx.
- App Service logs show successful media endpoint calls without logging secrets.

## 12. Official References

- Azure OpenAI Whisper quickstart: https://learn.microsoft.com/en-us/azure/ai-services/openai/whisper-quickstart
- Azure OpenAI REST API reference: https://learn.microsoft.com/en-us/azure/cognitive-services/openai/reference
- Azure App Service Key Vault references: https://learn.microsoft.com/en-us/azure/app-service/app-service-key-vault-references
- Azure App Service app settings: https://learn.microsoft.com/en-us/azure/app-service/configure-common
- Azure App Service managed identities: https://learn.microsoft.com/en-us/azure/app-service/overview-managed-identity
