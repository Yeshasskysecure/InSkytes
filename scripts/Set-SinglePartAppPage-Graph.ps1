param(
  [string]$TenantId = "547b64a7-e66e-48df-a146-3e898cbcb60f",
  [string]$ClientId = "f9403a61-d04e-46cf-a5d1-cbc68a66f740",
  [string]$Scope = "https://graph.microsoft.com/.default",
  [string]$SiteId = "skysecuretech.sharepoint.com,9d3cf7ea-0b61-4ad9-8780-433625ec6a6f,900862b0-5891-47d9-9ac2-83afc691e3c8",
  [string]$PageName = "Migration.aspx",
  [string]$LayoutType = "SingleWebPartAppPage"
)

$ErrorActionPreference = "Stop"

function Invoke-GraphRequest {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Method,
    [Parameter(Mandatory = $true)]
    [string]$Uri,
    [object]$Body
  )

  $headers = @{
    Authorization = "Bearer $accessToken"
    "Content-Type" = "application/json"
  }

  if ($null -ne $Body) {
    $jsonBody = $Body | ConvertTo-Json -Depth 20
    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $jsonBody
  }

  return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
}

Write-Host "This script sets a SharePoint page layout through Microsoft Graph:" -ForegroundColor Cyan
Write-Host "  Site: $SiteId"
Write-Host "  Page: $PageName"
Write-Host "  Layout: $LayoutType"
Write-Host ""

$clientSecretSecure = Read-Host "Enter Graph client secret" -AsSecureString
$clientSecret = [System.Net.NetworkCredential]::new("", $clientSecretSecure).Password

Write-Host "Getting Graph access token..." -ForegroundColor Cyan
$tokenResponse = Invoke-RestMethod `
  -Method Post `
  -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
  -ContentType "application/x-www-form-urlencoded" `
  -Body @{
    client_id = $ClientId
    client_secret = $clientSecret
    scope = $Scope
    grant_type = "client_credentials"
  }

$accessToken = $tokenResponse.access_token
$encodedSiteId = [System.Uri]::EscapeDataString($SiteId)

Write-Host "Finding Site Pages library..." -ForegroundColor Cyan
$lists = Invoke-GraphRequest `
  -Method Get `
  -Uri "https://graph.microsoft.com/v1.0/sites/$encodedSiteId/lists?`$select=id,displayName,name,webUrl"

$sitePagesList = $lists.value |
  Where-Object {
    $_.displayName -eq "Site Pages" -or
    $_.name -eq "SitePages" -or
    $_.name -eq "sitepages" -or
    $_.webUrl -like "*/SitePages"
  } |
  Select-Object -First 1

if ($sitePagesList) {
  Write-Host "Finding page item '$PageName'..." -ForegroundColor Cyan
  $encodedListId = [System.Uri]::EscapeDataString($sitePagesList.id)
  $items = Invoke-GraphRequest `
    -Method Get `
    -Uri "https://graph.microsoft.com/v1.0/sites/$encodedSiteId/lists/$encodedListId/items?`$expand=fields(`$select=FileLeafRef,Title,PageLayoutType)&`$top=200"

  $pageItem = $items.value |
    Where-Object { $_.fields.FileLeafRef -eq $PageName } |
    Select-Object -First 1

  if (-not $pageItem) {
    throw "Could not find '$PageName' in Site Pages. Create the page first, then rerun this script."
  }

  Write-Host "Setting PageLayoutType to '$LayoutType'..." -ForegroundColor Cyan
  $encodedItemId = [System.Uri]::EscapeDataString($pageItem.id)
  Invoke-GraphRequest `
    -Method Patch `
    -Uri "https://graph.microsoft.com/v1.0/sites/$encodedSiteId/lists/$encodedListId/items/$encodedItemId/fields" `
    -Body @{
      PageLayoutType = $LayoutType
    } | Out-Null
} else {
  Write-Host "Site Pages library was not listed. Trying drive lookup fallback..." -ForegroundColor Yellow
  Write-Host "Lists returned by Graph:" -ForegroundColor Yellow
  $lists.value |
    Select-Object displayName, name, webUrl |
    Format-Table -AutoSize

  $drives = Invoke-GraphRequest `
    -Method Get `
    -Uri "https://graph.microsoft.com/v1.0/sites/$encodedSiteId/drives?`$select=id,name,webUrl"

  Write-Host "Drives returned by Graph:" -ForegroundColor Yellow
  $drives.value |
    Select-Object name, webUrl |
    Format-Table -AutoSize

  $sitePagesDrive = $drives.value |
    Where-Object {
      $_.name -eq "Site Pages" -or
      $_.name -eq "SitePages" -or
      $_.webUrl -like "*/SitePages"
    } |
    Select-Object -First 1

  if (-not $sitePagesDrive) {
    throw "Could not find the 'Site Pages' library or drive through Graph."
  }

  $encodedDriveId = [System.Uri]::EscapeDataString($sitePagesDrive.id)
  $encodedPageName = [System.Uri]::EscapeDataString($PageName)

  Write-Host "Finding page drive item '$PageName'..." -ForegroundColor Cyan
  $driveItem = Invoke-GraphRequest `
    -Method Get `
    -Uri "https://graph.microsoft.com/v1.0/drives/$encodedDriveId/root:/$encodedPageName"

  Write-Host "Setting PageLayoutType to '$LayoutType'..." -ForegroundColor Cyan
  $encodedDriveItemId = [System.Uri]::EscapeDataString($driveItem.id)
  Invoke-GraphRequest `
    -Method Patch `
    -Uri "https://graph.microsoft.com/v1.0/drives/$encodedDriveId/items/$encodedDriveItemId/listItem/fields" `
    -Body @{
      PageLayoutType = $LayoutType
    } | Out-Null
}

Write-Host ""
Write-Host "Done. Refresh the page:" -ForegroundColor Green
Write-Host "https://skysecuretech.sharepoint.com/sites/InSkytes/SitePages/$PageName"
