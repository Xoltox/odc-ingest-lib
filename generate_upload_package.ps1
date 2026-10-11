# Builds dist/RFPIngestLib.zip for upload to ODC (zip root holds the DLLs, not a folder).
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$proj = Join-Path $here 'RFPIngestLib.csproj'
$publish = Join-Path $here 'bin/Release/net10.0/linux-x64/publish'
$dist = Join-Path $here 'dist'
$zip = Join-Path $dist 'RFPIngestLib.zip'

if (Test-Path $publish) { Remove-Item $publish -Recurse -Force }
dotnet publish $proj -c Release -r linux-x64 --self-contained false
if ($LASTEXITCODE -ne 0) { throw 'dotnet publish failed' }

New-Item -ItemType Directory -Force $dist | Out-Null
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $publish '*') -DestinationPath $zip
Write-Host "Created $zip"
