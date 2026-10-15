<#
.SYNOPSIS
    Builds dist/PrinterPackager.exe: one file with the app, logos and the PSADT v4 template inside.
.DESCRIPTION
    Runs on Windows or Linux with PowerShell 7 and the .NET SDK (8 or newer).
    PSADT is downloaded from its official GitHub release and checked against a pinned SHA256.
    IntuneWinAppUtil.exe is NOT bundled (its license forbids redistribution); the app fetches it from Microsoft on first use.
.EXAMPLE
    pwsh build/Build.ps1 -Version 2.0.0
#>
[CmdletBinding()]
param
(
    [string]$Version = '2.0.0',
    [string]$Dotnet = $(if ($env:DOTNET) { $env:DOTNET } else { 'dotnet' })
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$work = Join-Path $root 'build/.work'
$payload = Join-Path $work 'payload'
$dist = Join-Path $root 'dist'

$psadtVersion = '4.1.8'
$psadtUrl = "https://github.com/PSAppDeployToolkit/PSAppDeployToolkit/releases/download/$psadtVersion/PSAppDeployToolkit_Template_v4.zip"
$psadtSha256 = '50CB8D32973FC7648060A48CAD63912ECB5CACA5A70754F37E83AA06BD380283'

New-Item -ItemType Directory -Force -Path $work | Out-Null
if (Test-Path $payload) { Remove-Item $payload -Recurse -Force }

$psadtZip = Join-Path $work "PSADT-$psadtVersion.zip"
if (-not (Test-Path $psadtZip) -or (Get-FileHash $psadtZip -Algorithm SHA256).Hash -ne $psadtSha256)
{
    Write-Host "Downloading PSADT $psadtVersion..."
    Invoke-WebRequest -Uri $psadtUrl -OutFile $psadtZip -UseBasicParsing
}
$hash = (Get-FileHash $psadtZip -Algorithm SHA256).Hash
if ($hash -ne $psadtSha256) { throw "PSADT download hash mismatch: $hash" }

Write-Host 'Staging payload...'
New-Item -ItemType Directory -Force -Path (Join-Path $payload 'assets') | Out-Null
Copy-Item (Join-Path $root 'Start-PrinterPackager.ps1') $payload
Copy-Item (Join-Path $root 'assets/logos') (Join-Path $payload 'assets/logos') -Recurse
Copy-Item (Join-Path $root 'assets/branding') (Join-Path $payload 'assets/branding') -Recurse
Expand-Archive -Path $psadtZip -DestinationPath (Join-Path $payload 'PSAppDeployToolkit_Template_v4')

$payloadZip = Join-Path $work 'payload.zip'
if (Test-Path $payloadZip) { Remove-Item $payloadZip -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.Compression.ZipFile]::CreateFromDirectory($payload, $payloadZip, [IO.Compression.CompressionLevel]::Optimal, $false)
Write-Host ("Payload: {0:N1} MB" -f ((Get-Item $payloadZip).Length / 1MB))

Write-Host 'Building launcher...'
$numericVersion = ($Version -replace '[^0-9.].*$', '')
& $Dotnet publish (Join-Path $root 'src/Launcher/PrinterPackager.csproj') -c Release -o $dist `
    "-p:PayloadZip=$payloadZip" "-p:Version=$numericVersion" "-p:InformationalVersion=$Version" --nologo
if ($LASTEXITCODE) { throw "dotnet publish failed ($LASTEXITCODE)" }

Get-ChildItem $dist | Where-Object { $_.Name -ne 'PrinterPackager.exe' } | Remove-Item -Recurse -Force
$exe = Get-Item (Join-Path $dist 'PrinterPackager.exe')
Write-Host ("Built {0} ({1:N1} MB)" -f $exe.FullName, ($exe.Length / 1MB))
