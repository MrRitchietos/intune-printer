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
function Remove-TestPrinter
{
    # A deleted queue can stay pending while the spooler holds a handle; a spooler restart finishes it.
    Get-TestPrinter | Remove-Printer -ErrorAction SilentlyContinue
    Restart-Service -Name Spooler -Force
    Start-Sleep -Seconds 3
    Get-TestPrinter | Remove-Printer -ErrorAction SilentlyContinue
    if (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue) { Remove-PrinterPort -Name $portName }
}
function Invoke-AsSystem
{
    # Runs a program as SYSTEM through a scheduled task, the way Intune runs installs, and returns its exit code.
    param ([string]$FilePath, [string]$Arguments)
    $name = 'PrinterPackagerE2E'
    Unregister-ScheduledTask -TaskName $name -Confirm:$false -ErrorAction SilentlyContinue
    $action = New-ScheduledTaskAction -Execute $FilePath -Argument $Arguments -WorkingDirectory (Split-Path -Parent $FilePath)
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
    Register-ScheduledTask -TaskName $name -Action $action -Principal $principal -Settings $settings | Out-Null
    $started = Get-Date
    Start-ScheduledTask -TaskName $name
    $deadline = $started.AddMinutes(25)
    do
    {
        Start-Sleep -Seconds 3
        $task = Get-ScheduledTask -TaskName $name
        $info = Get-ScheduledTaskInfo -TaskName $name
    }
    while (($task.State -eq 'Running' -or $info.LastRunTime -lt $started.AddSeconds(-5)) -and (Get-Date) -lt $deadline)
    $code = $info.LastTaskResult
    Unregister-ScheduledTask -TaskName $name -Confirm:$false
    Write-Host "  SYSTEM run took $([int]((Get-Date) - $started).TotalSeconds) s"
    return $code
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
    if (Get-TestPrinter) { Remove-TestPrinter }
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
    Remove-TestPrinter
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

    Step 'Library: save a printer with its driver files'
    $libName = 'E2E Zebra Label Printer'
    $libHost = '10.123.45.68'
    $libPort = "IP_$libHost"
    $libDriver = 'Microsoft Print To PDF'
    $lib = Join-Path $WorkDir 'library'
    Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.DriverName -eq $libDriver } | Remove-Printer -ErrorAction SilentlyContinue
    if (-not (Get-PrinterDriver -Name $libDriver -ErrorAction SilentlyContinue)) { Add-PrinterDriver -Name $libDriver }
    if (-not (Get-PrinterPort -Name $libPort -ErrorAction SilentlyContinue)) { Add-PrinterPort -Name $libPort -PrinterHostAddress $libHost }
    Add-Printer -Name $libName -DriverName $libDriver -PortName $libPort -Location 'Warehouse'
    # Package the built-in driver like a vendor driver, so the driver install path really runs.
    $env:PRINTERPACKAGER_PACKAGE_INBOX = '1'
    $p = Start-Process -FilePath $Exe -ArgumentList '-PrinterName', "`"$libName`"", '-SaveToLibrary', '-LibraryDir', "`"$lib`"" -Wait -PassThru
    Assert ($p.ExitCode -eq 0) "save to library exit code 0 (got $($p.ExitCode))"
    $entry = Join-Path $lib $libName
    foreach ($f in 'printer.json', 'capture.json', 'Logo.png') { Assert (Test-Path (Join-Path $entry $f)) "library has $f" }
    Assert (Get-ChildItem (Join-Path $entry 'Driver') -Filter '*.inf' -Recurse) 'library has the driver files'
    Assert ((Get-FileHash (Join-Path $entry 'Logo.png')).Hash -eq (Get-FileHash (Join-Path $PSScriptRoot '..\assets\logos\zebra.png')).Hash) 'Zebra logo picked'

    Step 'Library: advanced settings (location + port 9101)'
    $json = Get-Content (Join-Path $entry 'printer.json') -Raw | ConvertFrom-Json
    $json.Settings.Advanced = [pscustomobject]@{ Location = 'Floor 2'; PortNumber = 9101; Duplex = 'TwoSidedLongEdge' }
    [IO.File]::WriteAllText((Join-Path $entry 'printer.json'), ($json | ConvertTo-Json -Depth 6))

    Step 'Remove the printer, its port and its driver'
    Get-Printer | Where-Object Name -eq $libName | Remove-Printer
    Restart-Service -Name Spooler -Force; Start-Sleep -Seconds 3
    Remove-PrinterPort -Name $libPort -ErrorAction SilentlyContinue
    Remove-PrinterDriver -Name $libDriver
    Assert (-not (Get-PrinterDriver -Name $libDriver -ErrorAction SilentlyContinue)) 'driver removed'

    Step 'Build the package from the library (printer not installed)'
    $out2 = Join-Path $WorkDir 'out-library'
    $p = Start-Process -FilePath $Exe -ArgumentList '-PrinterName', "`"$libName`"", '-FromLibrary', '-LibraryDir', "`"$lib`"", '-OutputDir', "`"$out2`"", '-AppVendor', 'Connectify', '-AppVersion', '1.0.0' -Wait -PassThru
    Remove-Item Env:\PRINTERPACKAGER_PACKAGE_INBOX
    Assert ($p.ExitCode -eq 0) "package from library exit code 0 (got $($p.ExitCode))"
    $folder2 = Get-ChildItem $out2 -Directory | Select-Object -First 1
    Assert (Get-ChildItem $folder2.FullName -Filter '*.intunewin') '.intunewin built from the library'
    $cfg = Get-Content (Join-Path $folder2.FullName 'Package\Files\printer_config.json') -Raw | ConvertFrom-Json
    Assert ($cfg.PortName -eq "IP_${libHost}_9101" -and $cfg.PortNumber -eq 9101) "advanced port used ($($cfg.PortName))"
    Assert ($cfg.Location -eq 'Floor 2') 'advanced location used'
    Assert ($cfg.DriverFolder -eq 'Driver') 'driver files are in the package'
    $adt2 = Join-Path $folder2.FullName 'Package\Invoke-AppDeployToolkit.exe'
    $det2 = Join-Path $folder2.FullName 'Detection.ps1'

    Step 'Install as SYSTEM (like Intune), driver must be installed again'
    $code = Invoke-AsSystem -FilePath $adt2 -Arguments '-DeploymentType Install -DeployMode Silent'
    Assert ($code -eq 0) "SYSTEM install exit code 0 (got $code)"
    Assert (Get-PrinterDriver -Name $libDriver -ErrorAction SilentlyContinue) 'driver installed again by the package'
    $lp = Get-Printer | Where-Object Name -eq $libName
    Assert $lp 'printer installed as SYSTEM'
    Assert ($lp.PortName -eq "IP_${libHost}_9101" -and (Get-PrinterPort -Name $lp.PortName).PortNumber -eq 9101) 'port 9101 created'
    Assert ($lp.Location -eq 'Floor 2') 'location Floor 2'
    Assert ((Invoke-Detection $det2) -eq 0) 'detected after SYSTEM install'

    Step 'Install again as SYSTEM (driver already there, must be skipped)'
    $code = Invoke-AsSystem -FilePath $adt2 -Arguments '-DeploymentType Install -DeployMode Silent'
    Assert ($code -eq 0) "second SYSTEM install exit code 0 (got $code)"
    $log = Get-ChildItem 'C:\Windows\Logs\Software' -Filter '*Zebra*Install.log' | Sort-Object LastWriteTime | Select-Object -Last 1
    Assert ($log -and (Select-String -Path $log.FullName -Pattern 'already installed; using it' -SimpleMatch -Quiet)) 'driver step skipped the second time'

    Step 'Uninstall as SYSTEM'
    $code = Invoke-AsSystem -FilePath $adt2 -Arguments '-DeploymentType Uninstall -DeployMode Silent'
    Assert ($code -eq 0) "SYSTEM uninstall exit code 0 (got $code)"
    Assert (-not (Get-Printer | Where-Object Name -eq $libName)) 'printer removed'
    Assert ((Invoke-Detection $det2) -eq 1) 'not detected after uninstall'
    if (-not (Get-PrinterDriver -Name $libDriver -ErrorAction SilentlyContinue)) { Add-PrinterDriver -Name $libDriver -ErrorAction SilentlyContinue }

    Step 'GUI screenshots'
    if (-not (Get-PrinterDriver -Name $driver -ErrorAction SilentlyContinue)) { Add-PrinterDriver -Name $driver }
    Add-PrinterPort -Name $portName -PrinterHostAddress $host1 -ErrorAction SilentlyContinue
    Add-Printer -Name $printerName -DriverName $driver -PortName $portName -Location 'CI rack'
    $p = Start-Process -FilePath $Exe -ArgumentList '-Screenshot', "`"$shots\gui`"" -PassThru
    if (-not $p.WaitForExit(120000)) { $p.Kill(); throw 'GUI did not close within 2 minutes' }
    Assert (Test-Path "$shots\gui-ready.png") 'ready screenshot saved'
    Assert (Test-Path "$shots\gui-advanced.png") 'advanced settings window opened and saved'
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
