<#
.SYNOPSIS
    End-to-end test on a Windows machine with admin rights (used by CI):
    create a test TCP/IP printer -> package it with PrinterPackager.exe -> remove it ->
    install the generated PSADT package -> Intune detection -> uninstall -> GUI screenshots.
#>
param
(
    [Parameter(Mandatory)][string]$Exe,
    [string]$WorkDir = 'C:\e2e'
)

$ErrorActionPreference = 'Stop'
$printerName = 'E2E HP LaserJet 4th floor'
$host1 = '10.123.45.67'
$portName = "IP_$host1"
$out = Join-Path $WorkDir 'out'
$shots = Join-Path $WorkDir 'shots'
New-Item -ItemType Directory -Force -Path $out, $shots | Out-Null

function Step($t) { Write-Host "`n=== $t" -ForegroundColor Cyan }
function Assert($cond, $msg) { if (-not $cond) { throw "ASSERT FAILED: $msg" } ; Write-Host "  ok: $msg" -ForegroundColor Green }
function Get-TestPrinter { Get-Printer -ErrorAction SilentlyContinue | Where-Object Name -eq $printerName }
function Invoke-Detection($path)
{
    $p = Start-Process -FilePath powershell.exe -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$path`"" -Wait -PassThru -NoNewWindow
    return $p.ExitCode
}
function Show-PsadtLogs { Get-ChildItem 'C:\Windows\Logs\Software' -Filter '*.log' -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "--- $($_.FullName)"; Get-Content $_.FullName -Tail 80 } }

try
{
    Step 'Spooler'
    Set-Service -Name Spooler -StartupType Automatic
    Start-Service -Name Spooler

    Step 'Create test printer'
    $driver = $null
    foreach ($d in 'Microsoft PS Class Driver', 'Microsoft PCL6 Class Driver', 'Generic / Text Only', 'Microsoft Print To PDF')
    {
        if (-not (Get-PrinterDriver -Name $d -ErrorAction SilentlyContinue)) { try { Add-PrinterDriver -Name $d -ErrorAction Stop } catch { continue } }
        if (Get-PrinterDriver -Name $d -ErrorAction SilentlyContinue) { $driver = $d; break }
    }
    Assert $driver "driver available ($driver)"
    if (Get-TestPrinter) { Remove-Printer -Name $printerName }
    if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $portName -PrinterHostAddress $host1 }
    Add-Printer -Name $printerName -DriverName $driver -PortName $portName -Location 'CI rack' -Comment 'End-to-end test'
    Assert (Get-TestPrinter) 'test printer created'

    Step 'Package with PrinterPackager.exe (headless)'
    $p = Start-Process -FilePath $Exe -ArgumentList '-PrinterName', "`"$printerName`"", '-OutputDir', "`"$out`"", '-AppVendor', 'Connectify', '-AppVersion', '1.2.0', '-Zip' -Wait -PassThru
    Get-Content (Join-Path $env:TEMP 'IntunePrinterPackager\launcher.log') -Tail 60 -ErrorAction SilentlyContinue
    Assert ($p.ExitCode -eq 0) "packager exit code 0 (got $($p.ExitCode))"
    $folder = Get-ChildItem $out -Directory | Select-Object -First 1
    Assert $folder "output folder $($folder.Name)"
    $intunewin = Get-ChildItem $folder.FullName -Filter '*.intunewin' | Select-Object -First 1
    Assert ($intunewin -and $intunewin.Length -gt 100KB) ".intunewin built ($($intunewin.Name), $([int]($intunewin.Length/1KB)) KB)"
    Assert (Get-ChildItem $folder.FullName -Filter '*.zip') '.zip built'
    foreach ($f in 'Logo.png', 'Detection.ps1', 'Intune-Upload-Info.txt', 'Package\Invoke-AppDeployToolkit.exe', 'Package\Files\printer_config.json', 'Package\Install-Printer.cmd') { Assert (Test-Path (Join-Path $folder.FullName $f)) "$f exists" }
    $logoHash = (Get-FileHash (Join-Path $folder.FullName 'Logo.png')).Hash
    $hpHash = (Get-FileHash (Join-Path $PSScriptRoot '..\assets\logos\hp.png')).Hash
    Assert ($logoHash -eq $hpHash) 'HP logo picked for the HP printer'
    Assert ((Get-FileHash (Join-Path $folder.FullName 'Package\Assets\AppIcon.png')).Hash -eq $hpHash) 'PSADT dialog icon is the HP logo'
    $errs = $null
    [void][Management.Automation.Language.Parser]::ParseFile((Join-Path $folder.FullName 'Package\Invoke-AppDeployToolkit.ps1'), [ref]$null, [ref]$errs)
    Assert (-not $errs) 'generated Invoke-AppDeployToolkit.ps1 parses'
    Get-Content (Join-Path $folder.FullName 'Intune-Upload-Info.txt')
    Get-Content (Join-Path $folder.FullName 'Package\Files\printer_config.json')
    $detection = Join-Path $folder.FullName 'Detection.ps1'

    Step 'Detection before install (printer exists but no package marker)'
    Assert ((Invoke-Detection $detection) -eq 1) 'not detected yet'

    Step 'Remove test printer and port'
    Remove-Printer -Name $printerName
    Remove-PrinterPort -Name $portName
    Assert (-not (Get-TestPrinter)) 'printer removed'

    Step 'Install via PSADT (silent, like Intune)'
    $adt = Join-Path $folder.FullName 'Package\Invoke-AppDeployToolkit.exe'
    $p = Start-Process -FilePath $adt -ArgumentList '-DeploymentType', 'Install', '-DeployMode', 'Silent' -Wait -PassThru
    Assert ($p.ExitCode -eq 0) "install exit code 0 (got $($p.ExitCode))"
    $installed = Get-TestPrinter
    Assert $installed 'printer installed'
    Assert ($installed.PortName -eq $portName) "port $portName"
    Assert ((Get-PrinterPort -Name $portName).PrinterHostAddress -eq $host1) "port points to $host1"
    Assert ($installed.Location -eq 'CI rack') 'location restored'
    Assert ((Invoke-Detection $detection) -eq 0) 'detected after install'

    Step 'Install again (idempotent)'
    $p = Start-Process -FilePath $adt -ArgumentList '-DeploymentType', 'Install', '-DeployMode', 'Silent' -Wait -PassThru
    Assert ($p.ExitCode -eq 0) "second install exit code 0 (got $($p.ExitCode))"

    Step 'Uninstall via PSADT'
    $p = Start-Process -FilePath $adt -ArgumentList '-DeploymentType', 'Uninstall', '-DeployMode', 'Silent' -Wait -PassThru
    Assert ($p.ExitCode -eq 0) "uninstall exit code 0 (got $($p.ExitCode))"
    Assert (-not (Get-TestPrinter)) 'printer removed by uninstall'
    Assert ((Invoke-Detection $detection) -eq 1) 'not detected after uninstall'

    Step 'GUI screenshots'
    if (-not (Get-PrinterDriver -Name $driver -ErrorAction SilentlyContinue)) { Add-PrinterDriver -Name $driver }
    Add-PrinterPort -Name $portName -PrinterHostAddress $host1 -ErrorAction SilentlyContinue
    Add-Printer -Name $printerName -DriverName $driver -PortName $portName -Location 'CI rack'
    $p = Start-Process -FilePath $Exe -ArgumentList '-Screenshot', "`"$shots\gui`"" -PassThru
    if (-not $p.WaitForExit(120000)) { $p.Kill(); throw 'GUI did not close within 2 minutes' }
    Assert (Test-Path "$shots\gui-ready.png") 'ready screenshot saved'
    Write-Host "`nALL E2E CHECKS PASSED" -ForegroundColor Green
}
catch
{
    Write-Host $_ -ForegroundColor Red
    Show-PsadtLogs
    Get-Content (Join-Path $env:TEMP 'IntunePrinterPackager\launcher.log') -ErrorAction SilentlyContinue
    Get-ChildItem "$env:ProgramData\IntunePrinterPackager\logs" -ErrorAction SilentlyContinue | Get-Content -Tail 80
    exit 1
}
