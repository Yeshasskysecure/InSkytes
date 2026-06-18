param(
  [string]$TenantId = "6d787ab7-f295-424f-b8cc-e3116a0f8520",
  [string]$TenantName = "indegene123",
  [string]$ClientId = "3baf2dbc-7365-4f9a-9826-3405a80fb4cf",
  [ValidateSet("PnPInteractive", "OSLogin", "DeviceLogin", "Interactive", "AppSecret")]
  [string]$AuthMode = "OSLogin",
  [string]$SitePath = "/sites/iKnowledgeNext",
  [string]$PageName = "Migration.aspx",
  [string]$PageTitle = "Migration",
  [string]$PackagePath = ".\sharepoint\solution\sharepoint-migration-site.sppkg",
  [string]$WebPartComponentId = "a1b2c3d4-e5f6-7890-abcd-ef1234567892",
  [switch]$SkipPackageDeployment
)

$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
  Install-Module PnP.PowerShell -Scope CurrentUser -Force
}

Import-Module PnP.PowerShell

$siteUrl = "https://$TenantName.sharepoint.com$SitePath"
$adminUrl = "https://$TenantName-admin.sharepoint.com"
$resolvedPackagePath = Resolve-Path $PackagePath

function Connect-IknPnP {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Url,
    [string]$ClientSecret
  )

  if ($AuthMode -eq "AppSecret") {
    Connect-PnPOnline -Url $Url -ClientId $ClientId -ClientSecret $ClientSecret
    return
  }

  if ($AuthMode -eq "DeviceLogin") {
    Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -DeviceLogin
    return
  }

  if ($AuthMode -eq "OSLogin") {
    Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -OSLogin
    return
  }

  if ($AuthMode -eq "PnPInteractive") {
    Connect-PnPOnline -Url $Url -ClientId "31359c7f-bd7e-475c-86db-fdb8c937548e" -Tenant $TenantId -Interactive
    return
  }

  Connect-PnPOnline -Url $Url -ClientId $ClientId -Tenant $TenantId -Interactive
}

Write-Host "This script will create/update a single-part app page at:" -ForegroundColor Cyan
Write-Host "  $siteUrl/SitePages/$PageName"
Write-Host ""

$clientSecret = $null
if ($AuthMode -eq "AppSecret") {
  $clientSecretSecure = Read-Host "Enter client secret" -AsSecureString
  $clientSecret = [System.Net.NetworkCredential]::new("", $clientSecretSecure).Password
}

if (-not $SkipPackageDeployment) {
  Write-Host "Connecting to tenant app catalog..." -ForegroundColor Cyan
  Connect-IknPnP -Url $adminUrl -ClientSecret $clientSecret

  Write-Host "Uploading and deploying SPFx package..." -ForegroundColor Cyan
  Add-PnPApp `
    -Path $resolvedPackagePath `
    -Scope Tenant `
    -Overwrite `
    -Publish `
    -SkipFeatureDeployment
} else {
  Write-Host "Skipping app package upload/deployment." -ForegroundColor Yellow
}

Write-Host "Connecting to target site..." -ForegroundColor Cyan
Connect-IknPnP -Url $siteUrl -ClientSecret $clientSecret

$existingPage = Get-PnPPage -Identity $PageName -ErrorAction SilentlyContinue

if ($existingPage) {
  Write-Host "Page exists. Setting layout to SingleWebPartAppPage..." -ForegroundColor Cyan
  Set-PnPPage -Identity $PageName -LayoutType SingleWebPartAppPage
} else {
  Write-Host "Creating single-part app page..." -ForegroundColor Cyan
  Add-PnPPage -Name $PageName -LayoutType SingleWebPartAppPage | Out-Null
}

$component = Get-PnPPageComponent -Page $PageName -ListAvailable |
  Where-Object { $_.Id -eq $WebPartComponentId -or $_.Name -eq $PageTitle } |
  Select-Object -First 1

if (-not $component) {
  throw "Could not find web part component '$WebPartComponentId'. Confirm the app package is deployed and the web part supports SharePointFullPage."
}

$page = Get-PnPPage -Identity $PageName
$existingControls = $page.Controls | Where-Object {
  $_.WebPartId -eq $WebPartComponentId -or $_.Title -eq $PageTitle
}

if (-not $existingControls) {
  Write-Host "Adding web part to page..." -ForegroundColor Cyan
  Add-PnPPageWebPart -Page $PageName -Component $component
} else {
  Write-Host "Web part is already on the page." -ForegroundColor Yellow
}

Write-Host "Publishing page..." -ForegroundColor Cyan
Set-PnPPage -Identity $PageName -Publish

Write-Host ""
Write-Host "Done:" -ForegroundColor Green
Write-Host "$siteUrl/SitePages/$PageName"
