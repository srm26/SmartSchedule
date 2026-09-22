# GES SmartSchedule - Deploy to Azure App Service
# Usage: .\deploy.ps1
# Zips the app, uploads to blob storage, restarts the App Service.

$APP_NAME   = "app-wus2-elecsch-dv-01"
$RG         = "rg-wus2-elecsch-dv-compute"
$STORAGE    = "stwus2elecschdv"
$CONTAINER  = "smartschedule-deploy"
$BLOB       = "app.zip"
$STORAGE_RG = "rg-wus2-elecsch-dv-storage"

Write-Host "==> Setting subscription..." -ForegroundColor Cyan
az account set --subscription "67791e0d-e90b-433b-832b-d393f8d4fc96"

# Create blob container if it doesn't exist
Write-Host "==> Ensuring blob container exists..." -ForegroundColor Cyan
az storage container create --name $CONTAINER --account-name $STORAGE --auth-mode login 2>$null

# Create zip (exclude secrets, state files, Excel/docs)
Write-Host "==> Creating zip..." -ForegroundColor Cyan
$zip = "$env:TEMP\smartschedule-app.zip"
if (Test-Path $zip) { Remove-Item $zip }
$files = Get-ChildItem -Path $PSScriptRoot -File | Where-Object {
    $_.Name -notmatch "\.(xlsx|docx|webp)$" -and
    $_.Name -ne "deploy.ps1"
}
Compress-Archive -Path $files.FullName -DestinationPath $zip -Force

# Upload
Write-Host "==> Uploading to blob storage..." -ForegroundColor Cyan
az storage blob upload `
    --account-name $STORAGE `
    --container-name $CONTAINER `
    --name $BLOB `
    --file $zip `
    --overwrite `
    --auth-mode login

# Generate SAS URL (valid 1 year)
Write-Host "==> Generating SAS URL..." -ForegroundColor Cyan
$expiry = (Get-Date).AddYears(1).ToString("yyyy-MM-dd")
$sasUrl = az storage blob generate-sas `
    --account-name $STORAGE `
    --container-name $CONTAINER `
    --name $BLOB `
    --permissions r `
    --expiry $expiry `
    --https-only `
    --full-uri `
    --auth-mode login `
    --as-user `
    -o tsv

Write-Host "SAS URL: $sasUrl"

# Set WEBSITE_RUN_FROM_PACKAGE
Write-Host "==> Setting WEBSITE_RUN_FROM_PACKAGE..." -ForegroundColor Cyan
az webapp config appsettings set `
    --name $APP_NAME `
    --resource-group $RG `
    --settings "WEBSITE_RUN_FROM_PACKAGE=$sasUrl" `
    --output none

# Restart
Write-Host "==> Restarting App Service..." -ForegroundColor Cyan
az webapp restart --name $APP_NAME --resource-group $RG

Write-Host "==> Done! App Service is restarting with the new package." -ForegroundColor Green
Write-Host "    URL: https://app-wus2-elecsch-dv-01.ase-wus2-elecsch-dv.appserviceenvironment.net"
