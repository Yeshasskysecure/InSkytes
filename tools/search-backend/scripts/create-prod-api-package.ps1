param(
  [string]$Version = (Get-Date -Format "yyyyMMdd-HHmmss"),
  [string]$StagePath = "C:\tmp\iknowledge-search-api-prod-stage",
  [string]$ZipPath = "",
  [switch]$SkipNpmInstall
)

$ErrorActionPreference = "Stop"

$repo = Resolve-Path (Join-Path $PSScriptRoot "..\..\..")
$repoPath = $repo.Path
if (-not $ZipPath) {
  $ZipPath = "C:\tmp\iknowledge-search-api-prod-$Version.zip"
}

$rootPackage = Get-Content (Join-Path $repoPath "package.json") -Raw | ConvertFrom-Json
$mammothVersion = $rootPackage.dependencies.mammoth
$xlsxVersion = $rootPackage.dependencies.xlsx
if (-not $mammothVersion -or -not $xlsxVersion) {
  throw "Root package.json must define mammoth and xlsx dependencies."
}

Write-Host "Building backend TypeScript..."
Push-Location $repoPath
npm run search:build
Pop-Location

Write-Host "Preparing clean stage: $StagePath"
Remove-Item $StagePath -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $ZipPath -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory $StagePath | Out-Null
New-Item -ItemType Directory (Join-Path $StagePath "tools\search-backend\config") -Force | Out-Null
New-Item -ItemType Directory (Join-Path $StagePath "tools\search-backend\scripts") -Force | Out-Null
New-Item -ItemType Directory (Join-Path $StagePath "tools\search-backend\scripts\lib") -Force | Out-Null

@(
  "[config]",
  "SCM_DO_BUILD_DURING_DEPLOYMENT=false"
) | Set-Content -Path (Join-Path $StagePath ".deployment") -Encoding ASCII

$packageJson = [ordered]@{
  name = "iknowledge-search-api-prod"
  version = "1.0.0"
  private = $true
  scripts = [ordered]@{
    "search:api:start" = "node tools/search-backend/scripts/local-api-server.js process"
  }
  engines = [ordered]@{
    node = ">=22.0.0"
  }
  dependencies = [ordered]@{
    mammoth = $mammothVersion
    xlsx = $xlsxVersion
  }
}
$packageJson | ConvertTo-Json -Depth 8 | Set-Content -Path (Join-Path $StagePath "package.json") -Encoding ASCII

Copy-Item (Join-Path $repoPath "tools\search-backend\dist") (Join-Path $StagePath "tools\search-backend\dist") -Recurse
Copy-Item (Join-Path $repoPath "tools\search-backend\config\search-domain-normalizations.json") (Join-Path $StagePath "tools\search-backend\config\search-domain-normalizations.json")
Copy-Item (Join-Path $repoPath "tools\search-backend\scripts\local-api-server.js") (Join-Path $StagePath "tools\search-backend\scripts\local-api-server.js")
Copy-Item (Join-Path $repoPath "tools\search-backend\scripts\delta-sync.js") (Join-Path $StagePath "tools\search-backend\scripts\delta-sync.js")
$runtimeLibFiles = @(
  "azureSearchClient.js",
  "chunker.js",
  "departmentTaxonomy.js",
  "env.js",
  "extractors.js",
  "graphClient.js",
  "inventory.js",
  "mappers.js",
  "openAiClient.js",
  "retry.js"
)
foreach ($fileName in $runtimeLibFiles) {
  Copy-Item (Join-Path $repoPath "tools\search-backend\scripts\lib\$fileName") (Join-Path $StagePath "tools\search-backend\scripts\lib\$fileName")
}

@(
  "Included:",
  "- .deployment disabling App Service build automation for this ready-to-run package",
  "- package.json with only backend runtime dependencies",
  "- node_modules generated from the minimal package.json",
  "- tools/search-backend/dist",
  "- tools/search-backend/config/search-domain-normalizations.json",
  "- tools/search-backend/scripts/local-api-server.js",
  "- tools/search-backend/scripts/delta-sync.js",
  "- tools/search-backend/scripts/lib runtime helpers used by delta-sync.js",
  "",
  "Excluded:",
  "- config/search.env and config/search.dev.env",
  "- tools/search-backend/.inventory",
  "- tools/search-backend/src",
  "- tools/search-backend/evals and tools/search-backend/evaluation",
  "- local indexing logs and ndjson run logs",
  "- reset-indexes.js, backfill-library.js, purge scripts, smoke/eval/admin scripts",
  "- scripts/lib/indexDefinitions.js local reset helper",
  "- SPFx frontend source"
) | Set-Content -Path (Join-Path $StagePath "DEPLOYMENT-MANIFEST.txt") -Encoding ASCII

if ($SkipNpmInstall) {
  Write-Host "Skipping npm install. Use this mode when Linux/Cloud Shell will install runtime dependencies before deployment."
} else {
  Write-Host "Installing production runtime dependencies in stage..."
  Push-Location $StagePath
  npm install --omit=dev --no-audit --no-fund
  Pop-Location
}

Write-Host "Creating zip: $ZipPath"
$tar = Get-Command tar.exe -ErrorAction SilentlyContinue
if ($env:OS -eq "Windows_NT" -and $tar) {
  # Compress-Archive writes Windows path separators into entry names. Kudu on Linux
  # can reject those paths, so use bsdtar's zip mode to emit forward slashes.
  & tar.exe -a -cf $ZipPath -C $StagePath .
} else {
  Compress-Archive -Path (Join-Path $StagePath "*") -DestinationPath $ZipPath -Force
}

$zip = Get-Item $ZipPath
Write-Host "Created package:"
$zip | Select-Object FullName, Length, LastWriteTime | Format-List
Write-Host "Inspect staged files at: $StagePath"
