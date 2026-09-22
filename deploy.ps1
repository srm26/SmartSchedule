# GES SmartSchedule - Deploy to Azure App Service
# Usage: .\deploy.ps1
# Push to GitHub first, then run this to deploy to Azure.

$APP_NAME  = "app-wus2-elecsch-dv-01"
$RG        = "rg-wus2-elecsch-dv-compute"
$STORAGE   = "stwus2elecschdv"
$STOR_RG   = "rg-wus2-elecsch-dv-storage"
$CONTAINER = "smartschedule-deploy"
$BLOB      = "app.zip"
$SUB       = "67791e0d-e90b-433b-832b-d393f8d4fc96"
$ZIP       = "$env:TEMP\smartschedule-app.zip"

Set-Location $PSScriptRoot

Write-Host "==> Setting subscription..." -ForegroundColor Cyan
az account set --subscription $SUB

# Get storage key
Write-Host "==> Getting storage key..." -ForegroundColor Cyan
$KEY = az storage account keys list --account-name $STORAGE --resource-group $STOR_RG --query "[0].value" -o tsv

# Create zip (exclude secrets, data files, deploy script itself)
Write-Host "==> Creating deployment zip..." -ForegroundColor Cyan
if (Test-Path $ZIP) { Remove-Item $ZIP -Force }
$include = @("app.py","requirements.txt","startup.sh",".gitignore","Event_via_SharedMailbox.txt")
$paths = $include | ForEach-Object { Join-Path $PSScriptRoot $_ } | Where-Object { Test-Path $_ }
Compress-Archive -LiteralPath $paths -DestinationPath $ZIP -Force
$size = [math]::Round(([System.IO.FileInfo]$ZIP).Length / 1KB)
Write-Host "  Zip: $size KB" -ForegroundColor Gray

# Upload
Write-Host "==> Uploading to blob storage..." -ForegroundColor Cyan
az storage blob upload `
    --account-name $STORAGE --account-key $KEY `
    --container-name $CONTAINER --name $BLOB `
    --file $ZIP --overwrite --output none

# Generate SAS URL (1 year expiry)
Write-Host "==> Generating SAS URL..." -ForegroundColor Cyan
$EXPIRY = (Get-Date).AddYears(1).ToString("yyyy-MM-dd")
$SAS_URL = az storage blob generate-sas `
    --account-name $STORAGE --account-key $KEY `
    --container-name $CONTAINER --name $BLOB `
    --permissions r --expiry $EXPIRY `
    --https-only --full-uri -o tsv

# Set WEBSITE_RUN_FROM_PACKAGE
Write-Host "==> Setting WEBSITE_RUN_FROM_PACKAGE..." -ForegroundColor Cyan
az webapp config appsettings set `
    --name $APP_NAME --resource-group $RG `
    --settings "WEBSITE_RUN_FROM_PACKAGE=$SAS_URL" --output none

# Restart
Write-Host "==> Restarting App Service..." -ForegroundColor Cyan
az webapp restart --name $APP_NAME --resource-group $RG
Write-Host ""
Write-Host "==> Deployment complete!" -ForegroundColor Green
Write-Host "    App: https://$APP_NAME.ase-wus2-elecsch-dv.appserviceenvironment.net" -ForegroundColor Green
