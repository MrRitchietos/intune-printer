<#
.SYNOPSIS
    Intune Printer Packager for PSADT v4 (WPF GUI).

.DESCRIPTION
    Discovers the printers installed on this computer and turns one of them into a ready-to-upload
    Intune Win32 app (.intunewin) built on PSAppDeployToolkit v4, plus an optional .zip for manual installs.
    Driver files, the TCP/IP port, printer preferences (PrintBrm + PrintUI) and the manufacturer logo
    are packaged automatically.

    Normally started through PrinterPackager.exe, which carries the PSADT template and logos and passes
    -AppRoot. The script also runs on its own: missing tools are downloaded (pinned versions, SHA256 checked).

.PARAMETER PrinterName
    Headless mode: package this printer without showing the window (for automation and tests).

.EXAMPLE
    .\Start-PrinterPackager.ps1

.EXAMPLE
    .\Start-PrinterPackager.ps1 -PrinterName 'Office HP' -OutputDir C:\Packages -AppVendor Contoso -Zip
#>
[CmdletBinding()]
param
(
    # Folder holding assets\ and (optionally) PSAppDeployToolkit_Template_v4\. Set by PrinterPackager.exe.
    [string]$AppRoot,

    # Folder PrinterPackager.exe was started from (e.g. a network share). Searched for IntuneWinAppUtil.exe.
    [string]$LauncherDir,

    # Callback from PrinterPackager.exe that closes its splash screen once the window is visible.
    [object]$OnReady,

    # Set by PrinterPackager.exe: its own version and full path (for the update button).
    [string]$ToolVersion,
    [string]$LauncherPath,

    # Headless mode.
    [string]$PrinterName,
    [string]$OutputDir,
    [string]$AppVendor,
    [string]$AppVersion = '1.0.0',
    [string]$PrinterHostAddress,
    [switch]$NoDriver,
    [switch]$NoPreferences,
    [switch]$NoIntuneWin,
    [switch]$Zip,
    [switch]$SetDefault,

    # Printer library (headless): -SaveToLibrary stores the printer, -FromLibrary packages the saved one.
    [string]$LibraryDir,
    [switch]$SaveToLibrary,
    [switch]$FromLibrary,

    # GUI self-test: saves <path>-loading.png and <path>-ready.png, then closes the window.
    [string]$Screenshot
)

$IsHeadless = -not [string]::IsNullOrWhiteSpace($PrinterName)

#region Bootstrap (only when started as a .ps1 file)
if ($PSCommandPath)
{
    $principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
    {
        if ($IsHeadless) { throw 'Administrator rights are required. Run this from an elevated PowerShell.' }
        $shell = (Get-Process -Id $PID).Path
        Start-Process -FilePath $shell -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', "`"$PSCommandPath`"") -Verb RunAs
        return
    }
    if (-not $IsHeadless -and [Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA')
    {
        & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -STA -File $PSCommandPath
        return
    }
    if (-not $AppRoot) { $AppRoot = $PSScriptRoot }
}
if (-not $AppRoot) { throw 'AppRoot was not provided.' }
#endregion

#region Shared context
$DataRoot = Join-Path -Path $env:ProgramData -ChildPath 'IntunePrinterPackager'
$Ctx = @{
    AppRoot        = $AppRoot
    LauncherDir    = $LauncherDir
    LogoDir        = Join-Path -Path $AppRoot -ChildPath 'assets\logos'
    BrandingDir    = Join-Path -Path $AppRoot -ChildPath 'assets\branding'
    ToolsCache     = Join-Path -Path $DataRoot -ChildPath 'tools'
    LogFile        = Join-Path -Path $DataRoot -ChildPath ("logs\PrinterPackager_{0:yyyyMMdd}.log" -f (Get-Date))
    PsadtVersion   = '4.1.8'
    PsadtUrl       = 'https://github.com/PSAppDeployToolkit/PSAppDeployToolkit/releases/download/4.1.8/PSAppDeployToolkit_Template_v4.zip'
    PsadtSha256    = '50CB8D32973FC7648060A48CAD63912ECB5CACA5A70754F37E83AA06BD380283'
    IntuneWinVersion = '1.8.7'
    IntuneWinUrl   = 'https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/v1.8.7/IntuneWinAppUtil.exe'
    IntuneWinSha256 = 'C1BA45B5CB939E84AF064BB7FF4B38FB3DFE33C8DC1078FD9B157672EAE671F6'
    UpdateRepo     = 'MrRitchietos/intune-printer'
    ToolVersion    = $ToolVersion
    LauncherPath   = $LauncherPath
}
try { New-Item -Path (Split-Path -Parent $Ctx.LogFile) -ItemType Directory -Force | Out-Null } catch { $Ctx.LogFile = $null }
$Sync = $null
#endregion

#region Worker code (runs in a background runspace so the window never freezes)
$WorkerCode = {

$Brands = @(
    [pscustomobject]@{ Name = 'Oce';            File = 'oce.png';            Pattern = '\bOc[e\u00e9]\b' }
    [pscustomobject]@{ Name = 'Fujifilm';       File = 'fujifilm.png';       Pattern = 'FUJIFILM|Fuji ?Xerox|\bApeos|DocuCentre|DocuPrint' }
    [pscustomobject]@{ Name = 'Xerox';          File = 'xerox.png';          Pattern = 'Xerox|WorkCentre|VersaLink|AltaLink|Phaser|PrimeLink' }
    [pscustomobject]@{ Name = 'HP';             File = 'hp.png';             Pattern = '\bHP\b|Hewlett|LaserJet|OfficeJet|DeskJet|DesignJet|PageWide|\bHP[A-Z0-9]' }
    [pscustomobject]@{ Name = 'Canon';          File = 'canon.png';          Pattern = 'Canon|imageRUNNER|imageCLASS|imagePROGRAF|i-SENSYS|\bUFR ?II|\biR-ADV' }
    [pscustomobject]@{ Name = 'Konica Minolta'; File = 'konica-minolta.png'; Pattern = 'Konica|Minolta|bizhub|\bKM[_ ]' }
    [pscustomobject]@{ Name = 'Kyocera';        File = 'kyocera.png';        Pattern = 'Kyocera|TASKalfa|ECOSYS' }
    [pscustomobject]@{ Name = 'Ricoh';          File = 'ricoh.png';          Pattern = 'Ricoh|Aficio' }
    [pscustomobject]@{ Name = 'Brother';        File = 'brother.png';        Pattern = 'Brother' }
    [pscustomobject]@{ Name = 'Epson';          File = 'epson.png';          Pattern = 'Epson|EcoTank|WorkForce' }
    [pscustomobject]@{ Name = 'Lexmark';        File = 'lexmark.png';        Pattern = 'Lexmark' }
    [pscustomobject]@{ Name = 'Sharp';          File = 'sharp.png';          Pattern = 'Sharp|\bMX-[0-9A-Z]' }
    [pscustomobject]@{ Name = 'Toshiba';        File = 'toshiba.png';        Pattern = 'Toshiba|e-STUDIO' }
    [pscustomobject]@{ Name = 'OKI';            File = 'oki.png';            Pattern = '\bOKI\b|OKIDATA' }
    [pscustomobject]@{ Name = 'Samsung';        File = 'samsung.png';        Pattern = 'Samsung' }
    [pscustomobject]@{ Name = 'Dell';           File = 'dell.png';           Pattern = '\bDell\b' }
    [pscustomobject]@{ Name = 'Zebra';          File = 'zebra.png';          Pattern = 'Zebra|ZDesigner|\bZ[DTQ]\d{3}|\bG[KXC]4\d\d' }
    [pscustomobject]@{ Name = 'Intermec';       File = 'intermec.png';       Pattern = 'Intermec' }
    [pscustomobject]@{ Name = 'Honeywell';      File = 'honeywell.png';      Pattern = 'Honeywell' }
)
$GenericBrand = [pscustomobject]@{ Name = 'Generic printer'; File = 'generic.png'; Pattern = $null }

function Write-Log
{
    param ([string]$Message, [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO')
    $line = '{0:HH:mm:ss}  {1,-5}  {2}' -f (Get-Date), $Level, $Message
    if ($Sync) { $Sync.Log.Enqueue($line) } else { Write-Host $line }
    if ($Ctx.LogFile) { try { [IO.File]::AppendAllText($Ctx.LogFile, "$line`r`n") } catch { } }
}

function Set-Step
{
    param ([string]$Text, [int]$Percent = -1)
    if ($script:ProgressRange -and $Percent -ge 0) { $Percent = $script:ProgressRange[0] + [int](($script:ProgressRange[1] - $script:ProgressRange[0]) * $Percent / 100) }
    if ($Sync) { $Sync.Step = $Text; $Sync.Percent = $Percent }
    Write-Log $Text
}

function Get-PrinterBrand
{
    param ([string[]]$Text)
    foreach ($t in $Text)
    {
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        foreach ($b in $Brands) { if ($t -match $b.Pattern) { return $b } }
    }
    return $GenericBrand
}

function Get-LogoList
{
    $items = foreach ($b in (@($Brands | Sort-Object Name) + $GenericBrand))
    {
        $path = Join-Path -Path $Ctx.LogoDir -ChildPath $b.File
        if (Test-Path -LiteralPath $path) { [pscustomobject]@{ Name = $b.Name; Path = $path } }
    }
    return @($items)
}

function ConvertTo-SafeFileName
{
    param ([string]$Name, [int]$MaxLength = 80)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $chars = foreach ($c in $Name.ToCharArray()) { if ($invalid -contains $c) { '_' } else { $c } }
    $safe = (-join $chars).Trim().TrimEnd('.')
    if ($safe.Length -gt $MaxLength) { $safe = $safe.Substring(0, $MaxLength).Trim() }
    return $safe
}

# ASCII-only PowerShell literal, so generated scripts survive any file encoding.
function ConvertTo-PSLiteral
{
    param ([string]$Value)
    if ($null -eq $Value) { return "''" }
    if ($Value -match '^[\x20-\x7E]*$') { return "'" + $Value.Replace("'", "''") + "'" }
    $parts = New-Object System.Collections.Generic.List[string]
    $run = ''
    foreach ($c in $Value.ToCharArray())
    {
        if ([int]$c -ge 0x20 -and [int]$c -le 0x7E) { $run += $c; continue }
        if ($run) { $parts.Add("'" + $run.Replace("'", "''") + "'"); $run = '' }
        $parts.Add('[string][char]0x{0:X4}' -f [int]$c)
    }
    if ($run) { $parts.Add("'" + $run.Replace("'", "''") + "'") }
    return '(' + ($parts -join ' + ') + ')'
}

function Invoke-Native
{
    param ([string]$FilePath, [string]$Arguments, [int]$TimeoutSeconds = 600)
    $psi = New-Object System.Diagnostics.ProcessStartInfo -ArgumentList $FilePath, $Arguments
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.CreateNoWindow = $true
    $proc = [Diagnostics.Process]::Start($psi)
    $stdout = $proc.StandardOutput.ReadToEndAsync()
    $stderr = $proc.StandardError.ReadToEndAsync()
    if (-not $proc.WaitForExit($TimeoutSeconds * 1000))
    {
        try { $proc.Kill() } catch { }
        throw "$([IO.Path]::GetFileName($FilePath)) did not finish within $TimeoutSeconds seconds."
    }
    $proc.WaitForExit()
    return [pscustomobject]@{ ExitCode = $proc.ExitCode; Output = ($stdout.Result + $stderr.Result).Trim() }
}

function Invoke-WithTimeout
{
    param ([string]$Script, [hashtable]$Parameters = @{}, [int]$TimeoutSeconds = 20)
    $ps = [powershell]::Create()
    [void]$ps.AddScript($Script)
    foreach ($k in $Parameters.Keys) { [void]$ps.AddParameter($k, $Parameters[$k]) }
    $handle = $ps.BeginInvoke()
    if ($handle.AsyncWaitHandle.WaitOne($TimeoutSeconds * 1000))
    {
        try { return $ps.EndInvoke($handle) } catch { return $null } finally { $ps.Dispose() }
    }
    [void]$ps.BeginStop($null, $null)
    return $null
}

function Save-Download
{
    param ([string]$Url, [string]$Destination, [string]$Sha256)
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $tmp = "$Destination.download"
    $wc = New-Object Net.WebClient
    $wc.Headers['User-Agent'] = 'IntunePrinterPackager'
    $wc.Proxy = [Net.WebRequest]::GetSystemWebProxy()
    $wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
    try { $wc.DownloadFile($Url, $tmp) } finally { $wc.Dispose() }
    $hash = (Get-FileHash -LiteralPath $tmp -Algorithm SHA256).Hash
    if ($hash -ne $Sha256)
    {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        throw "Download of $Url failed the SHA256 check (got $hash)."
    }
    Move-Item -LiteralPath $tmp -Destination $Destination -Force
}

function Test-PsadtTemplate
{
    param ([string]$Path)
    return ($Path -and (Test-Path -LiteralPath (Join-Path $Path 'Invoke-AppDeployToolkit.ps1')) -and (Test-Path -LiteralPath (Join-Path $Path 'PSAppDeployToolkit\PSAppDeployToolkit.psd1')))
}

function Resolve-PsadtTemplate
{
    $candidates = @(
        (Join-Path $Ctx.AppRoot 'PSAppDeployToolkit_Template_v4')
        (Join-Path (Split-Path -Parent $Ctx.AppRoot) 'PSAppDeployToolkit_Template_v4')
        (Join-Path $Ctx.ToolsCache "PSADT-$($Ctx.PsadtVersion)")
    )
    foreach ($c in $candidates) { if (Test-PsadtTemplate $c) { return $c } }

    $target = Join-Path $Ctx.ToolsCache "PSADT-$($Ctx.PsadtVersion)"
    Set-Step "Downloading PSAppDeployToolkit $($Ctx.PsadtVersion) (one time)..."
    New-Item -Path $Ctx.ToolsCache -ItemType Directory -Force | Out-Null
    $zipFile = Join-Path $Ctx.ToolsCache "PSADT-$($Ctx.PsadtVersion).zip"
    Save-Download -Url $Ctx.PsadtUrl -Destination $zipFile -Sha256 $Ctx.PsadtSha256
    if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::ExtractToDirectory($zipFile, $target)
    Remove-Item -LiteralPath $zipFile -Force
    if (-not (Test-PsadtTemplate $target)) { throw 'The downloaded PSADT template is incomplete.' }
    return $target
}

function Resolve-IntuneWinAppUtil
{
    # A copy placed next to PrinterPackager.exe (or the script) always wins, so offline sites can drop it on the share.
    $candidates = @()
    if ($Ctx.LauncherDir) { $candidates += Join-Path $Ctx.LauncherDir 'IntuneWinAppUtil.exe' }
    $candidates += Join-Path $Ctx.AppRoot 'IntuneWinAppUtil.exe'
    $candidates += Join-Path (Split-Path -Parent $Ctx.AppRoot) 'IntuneWinAppUtil.exe'
    $onPath = Get-Command -Name 'IntuneWinAppUtil.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($onPath) { $candidates += $onPath.Source }
    foreach ($c in $candidates) { if ($c -and (Test-Path -LiteralPath $c)) { return $c } }

    $cached = Join-Path $Ctx.ToolsCache "IntuneWinAppUtil-$($Ctx.IntuneWinVersion)\IntuneWinAppUtil.exe"
    if ((Test-Path -LiteralPath $cached) -and (Get-FileHash -LiteralPath $cached -Algorithm SHA256).Hash -eq $Ctx.IntuneWinSha256) { return $cached }

    Set-Step "Downloading IntuneWinAppUtil $($Ctx.IntuneWinVersion) from Microsoft (one time)..."
    New-Item -Path (Split-Path -Parent $cached) -ItemType Directory -Force | Out-Null
    try
    {
        Save-Download -Url $Ctx.IntuneWinUrl -Destination $cached -Sha256 $Ctx.IntuneWinSha256
    }
    catch
    {
        throw "IntuneWinAppUtil.exe could not be downloaded ($($_.Exception.Message)). Put IntuneWinAppUtil.exe next to PrinterPackager.exe, or untick the Intune package option."
    }
    return $cached
}

function Get-PrinterInventory
{
    Set-Step 'Reading installed printers...'
    $printers = @(Get-Printer -ErrorAction Stop)
    $ports = @{}
    foreach ($p in @(Get-PrinterPort -ErrorAction SilentlyContinue)) { $ports[$p.Name] = $p }
    $drivers = @{}
    foreach ($d in @(Get-PrinterDriver -ErrorAction SilentlyContinue))
    {
        if (-not $drivers.ContainsKey($d.Name) -or $d.PrinterEnvironment -eq 'Windows x64') { $drivers[$d.Name] = $d }
    }

    $skippedShared = 0
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($p in $printers)
    {
        if ("$($p.Type)" -eq 'Connection') { $skippedShared++; continue }
        if ($p.Name -match '^(Microsoft (XPS|Print to PDF)|Fax|OneNote|Send To OneNote|Adobe PDF)' -or
            $p.PortName -match '^(PORTPROMPT:|nul:|FILE:|SHRFAX:|XPSPort:|Ne\d\d:)') { continue }

        $port = $ports[$p.PortName]
        $drv = $drivers[$p.DriverName]
        $hostAddress = if ($port -and $port.PrinterHostAddress) { "$($port.PrinterHostAddress)" } else { '' }
        $portKind = if ($hostAddress) { 'TCP/IP' }
            elseif ($p.PortName -like 'WSD*') { 'WSD' }
            elseif ($p.PortName -like 'USB*' -or $p.PortName -like 'DOT4*') { 'USB' }
            elseif ($p.PortName -match '^https?://') { 'IPP' }
            else { 'Other' }
        $manufacturer = if ($drv) { "$($drv.Manufacturer)" } else { '' }
        $brand = Get-PrinterBrand -Text @($manufacturer, $p.DriverName, $p.Name)

        $list.Add([pscustomobject]@{
            Key            = "pc:$($p.Name)"
            Name           = $p.Name
            DriverName     = $p.DriverName
            Manufacturer   = $manufacturer
            PortName       = $p.PortName
            PortKind       = $portKind
            HostAddress    = $hostAddress
            Address        = if ($hostAddress) { $hostAddress } else { $p.PortName }
            Protocol       = if ($port -and $port.Protocol -eq 2) { 2 } else { 1 }
            PortNumber     = if ($port -and $port.PortNumber) { [int]$port.PortNumber } else { 9100 }
            LprQueueName   = if ($port -and $port.LprQueueName) { "$($port.LprQueueName)" } else { '' }
            SnmpEnabled    = [bool]($port -and $port.SNMPEnabled)
            SnmpCommunity  = if ($port -and $port.SNMPCommunity) { "$($port.SNMPCommunity)" } else { 'public' }
            SnmpIndex      = if ($port -and $port.SNMPIndex) { [int]$port.SNMPIndex } else { 1 }
            Location       = "$($p.Location)"
            Comment        = "$($p.Comment)"
            InfPath        = if ($drv) { "$($drv.InfPath)" } else { '' }
            Brand          = $brand.Name
            LogoPath       = Join-Path $Ctx.LogoDir $brand.File
        })
    }
    $sorted = @($list | Sort-Object Name)
    $msg = "Found $($sorted.Count) printer(s)."
    if ($skippedShared) { $msg += " Skipped $skippedShared printer(s) connected through a print server." }
    Write-Log $msg 'OK'
    return [pscustomobject]@{ __Result = $true; Printers = $sorted }
}

function Export-PrinterDriverFiles
{
    param ([string]$DriverName, [string]$InfPath, [string]$DestinationDir)

    $repo = Join-Path $env:SystemRoot 'System32\DriverStore\FileRepository'
    New-Item -Path $DestinationDir -ItemType Directory -Force | Out-Null

    # Built-in Windows drivers (ntprint / prnms) exist on every device and are huge; don't package them.
    if ($InfPath -and (Split-Path -Leaf $InfPath) -match '^(ntprint4?|prnms\d+)\.inf$' -and $env:PRINTERPACKAGER_PACKAGE_INBOX -ne '1')
    {
        Write-Log "'$DriverName' is a built-in Windows driver; devices already have it."
        return 'INBOX'
    }

    # 1. Normal case: the driver package lives in the DriverStore.
    if ($InfPath -and (Test-Path -LiteralPath $InfPath) -and $InfPath -like "$repo\*")
    {
        $folder = Split-Path -Parent $InfPath
        Write-Log "Copying driver package from $folder"
        Copy-Item -Path (Join-Path $folder '*') -Destination $DestinationDir -Recurse -Force
        return (Split-Path -Leaf $InfPath)
    }

    # 2. Published oemNN.inf: export just that package.
    if ($InfPath -and (Test-Path -LiteralPath $InfPath) -and $InfPath -like '*.inf')
    {
        $leaf = Split-Path -Leaf $InfPath
        $tmp = Join-Path $env:SystemRoot ("Temp\PrinterPackager_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
        try
        {
            New-Item -Path $tmp -ItemType Directory -Force | Out-Null
            $r = Invoke-Native -FilePath (Join-Path $env:SystemRoot 'System32\pnputil.exe') -Arguments "/export-driver $leaf `"$tmp`""
            $inf = Get-ChildItem -LiteralPath $tmp -Filter '*.inf' -Recurse | Select-Object -First 1
            if ($inf)
            {
                Copy-Item -Path (Join-Path $inf.DirectoryName '*') -Destination $DestinationDir -Recurse -Force
                Write-Log "Exported driver package $leaf with pnputil."
                return $inf.Name
            }
            Write-Log "pnputil could not export $leaf (exit $($r.ExitCode))." 'WARN'
        }
        finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    }

    # 3. Last resort: find the INF in the DriverStore that defines this driver name.
    Write-Log 'Searching the DriverStore for the driver package...'
    $match = Get-ChildItem -LiteralPath $repo -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object { Get-ChildItem -LiteralPath $_.FullName -Filter '*.inf' -File -ErrorAction SilentlyContinue } |
        Where-Object { Select-String -LiteralPath $_.FullName -Pattern $DriverName -SimpleMatch -Quiet -ErrorAction SilentlyContinue } |
        Select-Object -First 1
    if ($match)
    {
        Copy-Item -Path (Join-Path $match.DirectoryName '*') -Destination $DestinationDir -Recurse -Force
        Write-Log "Copied driver package from $($match.DirectoryName)"
        return $match.Name
    }
    return $null
}

function Export-PrinterPreferences
{
    # PrintBrm: full -NOBIN backup -> unpack -> keep only this queue/port/driver -> repack.
    param ([string]$PrinterName, [string]$PortName, [bool]$KeepPort, [string]$DriverName, [string]$DestinationFile)

    $printBrm = Join-Path $env:SystemRoot 'System32\spool\tools\PrintBrm.exe'
    if (-not (Test-Path -LiteralPath $printBrm)) { Write-Log 'PrintBrm.exe is not available on this computer; skipping.' 'WARN'; return $false }

    # PrintBrm takes quotes literally, so it gets space-free paths under C:\Windows\Temp without quotes.
    $work = Join-Path $env:SystemRoot ("Temp\PrinterPackager_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $fullCab = Join-Path $work 'full.printerExport'
    $unpacked = Join-Path $work 'unpacked'
    $outCab = Join-Path $work 'printer.printerExport'
    try
    {
        New-Item -Path $work -ItemType Directory -Force | Out-Null
        $r = Invoke-Native -FilePath $printBrm -Arguments "-B -NOBIN -F $fullCab"
        if (-not (Test-Path -LiteralPath $fullCab)) { Write-Log "PrintBrm backup failed (exit $($r.ExitCode)). $($r.Output)" 'WARN'; return $false }

        $r = Invoke-Native -FilePath $printBrm -Arguments "-R -F $fullCab -D $unpacked"
        if (-not (Test-Path -LiteralPath $unpacked)) { Write-Log "PrintBrm unpack failed (exit $($r.ExitCode))." 'WARN'; return $false }

        $keepFiles = New-Object System.Collections.Generic.List[string]
        $xmlPath = Join-Path $unpacked 'BrmPrinters.xml'
        if (Test-Path -LiteralPath $xmlPath)
        {
            $xml = New-Object xml
            $xml.Load($xmlPath)
            foreach ($q in @($xml.SelectNodes('/PRINTERS/PRINTQUEUE')))
            {
                if ($q.PrinterName -eq $PrinterName) { if ($q.FileName) { $keepFiles.Add($q.FileName) } }
                else { [void]$q.ParentNode.RemoveChild($q) }
            }
            $xml.Save($xmlPath)
        }

        $xmlPath = Join-Path $unpacked 'BrmPorts.xml'
        if (Test-Path -LiteralPath $xmlPath)
        {
            $xml = New-Object xml
            $xml.Load($xmlPath)
            foreach ($pt in @($xml.SelectNodes('/PRINTERPORTS/SPM')))
            {
                if (-not $KeepPort -or $pt.PortName -ne $PortName) { [void]$pt.ParentNode.RemoveChild($pt) }
            }
            $xml.Save($xmlPath)
        }

        $xmlPath = Join-Path $unpacked 'BrmDrivers.xml'
        if (Test-Path -LiteralPath $xmlPath)
        {
            $xml = New-Object xml
            $xml.Load($xmlPath)
            foreach ($dr in @($xml.SelectNodes('/PRINTERDRIVERS/DRIVER')))
            {
                if ($dr.DriverName -ne $DriverName) { [void]$dr.ParentNode.RemoveChild($dr) }
            }
            $xml.Save($xmlPath)
        }

        $printersDir = Join-Path $unpacked 'Printers'
        if ((Test-Path -LiteralPath $printersDir) -and $keepFiles.Count -gt 0)
        {
            Get-ChildItem -LiteralPath $printersDir -Filter '*.xml' | Where-Object { -not $keepFiles.Contains($_.Name) } | Remove-Item -Force
        }

        $r = Invoke-Native -FilePath $printBrm -Arguments "-B -NOBIN -F $outCab -D $unpacked"
        if (-not (Test-Path -LiteralPath $outCab)) { Write-Log "PrintBrm repack failed (exit $($r.ExitCode))." 'WARN'; return $false }
        Copy-Item -LiteralPath $outCab -Destination $DestinationFile -Force
        Write-Log ("Printer settings captured with PrintBrm ({0:N0} KB)." -f ((Get-Item -LiteralPath $DestinationFile).Length / 1KB)) 'OK'
        return $true
    }
    catch
    {
        Write-Log "PrintBrm export failed: $($_.Exception.Message)" 'WARN'
        return $false
    }
    finally { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}

function Export-PrinterSettingsDat
{
    # PrintUI /Ss: vendor DevMode (trays, duplex, colour, finishing). /q keeps it from showing a hidden, blocking error box.
    param ([string]$PrinterName, [string]$DestinationFile)
    $tmp = Join-Path $env:SystemRoot ("Temp\PrinterPackager_" + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.dat')
    try
    {
        $rundll = Join-Path $env:SystemRoot 'System32\rundll32.exe'
        foreach ($flags in 'g d', 'g')
        {
            $r = Invoke-Native -FilePath $rundll -Arguments "printui.dll,PrintUIEntry /Ss /n `"$PrinterName`" /a $tmp $flags /q" -TimeoutSeconds 120
            if ((Test-Path -LiteralPath $tmp) -and (Get-Item -LiteralPath $tmp).Length -gt 0)
            {
                Copy-Item -LiteralPath $tmp -Destination $DestinationFile -Force
                Write-Log ("Printer preferences captured with PrintUI ({0:N0} KB)." -f ((Get-Item -LiteralPath $DestinationFile).Length / 1KB)) 'OK'
                return $true
            }
        }
        Write-Log "PrintUI did not produce a settings file (exit $($r.ExitCode)); skipping." 'WARN'
    }
    catch { Write-Log "PrintUI export failed: $($_.Exception.Message)" 'WARN' }
    finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    return $false
}

function Convert-LogoToIcon
{
    # Square 256x256 PNG on white, so Intune and the PSADT dialogs never stretch it and dark logos stay visible.
    param ([string]$Source, [string]$Destination)
    Add-Type -AssemblyName System.Drawing
    $img = [Drawing.Image]::FromFile($Source)
    try
    {
        if ($img.Width -eq 256 -and $img.Height -eq 256 -and $Source -like "$($Ctx.LogoDir)*")
        {
            Copy-Item -LiteralPath $Source -Destination $Destination -Force
            return
        }
        $size = 256; $pad = 28
        $scale = [Math]::Min(($size - 2 * $pad) / $img.Width, ($size - 2 * $pad) / $img.Height)
        $w = [int][Math]::Max(1, [Math]::Round($img.Width * $scale))
        $h = [int][Math]::Max(1, [Math]::Round($img.Height * $scale))
        $bmp = New-Object Drawing.Bitmap $size, $size
        $g = [Drawing.Graphics]::FromImage($bmp)
        try
        {
            $g.SmoothingMode = 'AntiAlias'
            $g.InterpolationMode = 'HighQualityBicubic'
            $g.Clear([Drawing.Color]::White)
            $g.DrawImage($img, [int](($size - $w) / 2), [int](($size - $h) / 2), $w, $h)
        }
        finally { $g.Dispose() }
        $bmp.Save($Destination, [Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
    }
    finally { $img.Dispose() }
}

function Set-TemplateValue
{
    param ([string]$Code, [string]$Name, [string]$Value)
    $m = [regex]::Match($Code, "(?m)^(?<pre>[ \t]*$Name[ \t]*=[ \t]*)'[^'\r\n]*'")
    if (-not $m.Success) { throw "The PSADT template has no '$Name' entry." }
    return $Code.Substring(0, $m.Index) + $m.Groups['pre'].Value + "'" + $Value.Replace("'", "''") + "'" + $Code.Substring($m.Index + $m.Length)
}

function Get-DeploymentFunctions
{
    # Inserted into Invoke-AppDeployToolkit.ps1. Single-quoted: nothing here is expanded at packaging time.
    return @'
function Get-PackagedPrinterConfig
{
    $configFile = Join-Path -Path $adtSession.DirFiles -ChildPath 'printer_config.json'
    if (!(Test-Path -LiteralPath $configFile))
    {
        throw "printer_config.json was not found in [$($adtSession.DirFiles)]."
    }
    return [System.IO.File]::ReadAllText($configFile) | ConvertFrom-Json
}

function Get-PackagedPrinter
{
    param ([System.String]$Name)
    return Get-Printer -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
}

function Get-NativeSystem32
{
    if ([System.Environment]::Is64BitOperatingSystem -and ![System.Environment]::Is64BitProcess)
    {
        return Join-Path -Path $env:SystemRoot -ChildPath 'Sysnative'
    }
    return Join-Path -Path $env:SystemRoot -ChildPath 'System32'
}

function Get-PackagedPrinterRegistryKey
{
    param ([System.String]$Name, [System.Boolean]$Writable)
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    $path = "SOFTWARE\IntunePrinterPackager\Printers\$Name"
    if ($Writable) { return $base.CreateSubKey($path) }
    return $base.OpenSubKey($path)
}

function Install-PackagedPrinter
{
    $config = Get-PackagedPrinterConfig
    $system32 = Get-NativeSystem32
    $staging = Join-Path -Path $env:SystemRoot -ChildPath ("Temp\PrinterPackager_" + [System.Guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -Path $staging -ItemType Directory -Force | Out-Null

    try
    {
        ## 1. Driver: skip completely when this driver is already installed; otherwise stage it, then add it to the spooler.
        $driverInstalled = [System.Boolean](Get-PrinterDriver -Name $config.DriverName -ErrorAction SilentlyContinue)
        if ($driverInstalled)
        {
            Write-ADTLogEntry -Message "Printer driver [$($config.DriverName)] is already installed; using it."
        }
        elseif ($config.DriverFolder)
        {
            # Windows protected print mode only allows Windows' own drivers; say so instead of failing with a vague error.
            $wpp = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\WPP' -Name 'WindowsProtectedPrintMode' -ErrorAction SilentlyContinue
            if ($wpp -and $wpp.WindowsProtectedPrintMode -eq 1)
            {
                throw "Windows protected print mode is turned on for this device, which blocks the printer driver [$($config.DriverName)]. Turn it off by policy for this device or use a printer with a built-in Windows driver."
            }
            $driverPath = Join-Path -Path $adtSession.DirFiles -ChildPath $config.DriverFolder
            if (Test-Path -LiteralPath $driverPath)
            {
                $stagedDriver = Join-Path -Path $staging -ChildPath 'Driver'
                Copy-Item -LiteralPath $driverPath -Destination $stagedDriver -Recurse -Force
                foreach ($inf in (Get-ChildItem -LiteralPath $stagedDriver -Filter '*.inf' -Recurse))
                {
                    Write-ADTLogEntry -Message "Adding driver package [$($inf.Name)] to the DriverStore."
                    $result = Start-ADTProcess -FilePath (Join-Path -Path $system32 -ChildPath 'pnputil.exe') -ArgumentList '/add-driver', $inf.FullName -CreateNoWindow -IgnoreExitCodes '*' -PassThru -Timeout (New-TimeSpan -Minutes 10) -TimeoutAction Continue
                    Write-ADTLogEntry -Message "pnputil exit code [$($result.ExitCode)]."
                }
            }
        }

        if (!$driverInstalled)
        {
            Write-ADTLogEntry -Message "Adding printer driver [$($config.DriverName)] to the spooler."
            try
            {
                Add-PrinterDriver -Name $config.DriverName -ErrorAction Stop
            }
            catch
            {
                $storeInf = $null
                if ($config.DriverInf)
                {
                    $pattern = [System.IO.Path]::GetFileNameWithoutExtension($config.DriverInf) + '.inf_*'
                    $storeInf = Get-ChildItem -Path (Join-Path -Path $system32 -ChildPath 'DriverStore\FileRepository') -Directory -Filter $pattern -ErrorAction SilentlyContinue |
                        Sort-Object -Property LastWriteTime -Descending |
                        ForEach-Object { Join-Path -Path $_.FullName -ChildPath $config.DriverInf } |
                        Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
                }
                if (!$storeInf) { throw }
                Write-ADTLogEntry -Message "Retrying with INF [$storeInf]."
                Add-PrinterDriver -Name $config.DriverName -InfPath $storeInf -ErrorAction Stop
            }
        }

        ## 2. TCP/IP port.
        $port = Get-PrinterPort -Name $config.PortName -ErrorAction SilentlyContinue
        if (!$port)
        {
            Write-ADTLogEntry -Message "Creating printer port [$($config.PortName)] for [$($config.PrinterHostAddress)]."
            if ($config.Protocol -eq 2)
            {
                $queue = if ($config.LprQueueName) { $config.LprQueueName } else { 'lp' }
                Add-PrinterPort -Name $config.PortName -LprHostAddress $config.PrinterHostAddress -LprQueueName $queue -ErrorAction Stop
            }
            elseif ($config.SnmpEnabled)
            {
                Add-PrinterPort -Name $config.PortName -PrinterHostAddress $config.PrinterHostAddress -PortNumber $config.PortNumber -SNMP $config.SnmpIndex -SNMPCommunity $config.SnmpCommunity -ErrorAction Stop
            }
            else
            {
                Add-PrinterPort -Name $config.PortName -PrinterHostAddress $config.PrinterHostAddress -PortNumber $config.PortNumber -ErrorAction Stop
            }
        }
        elseif ($port.PrinterHostAddress -and $port.PrinterHostAddress -ne $config.PrinterHostAddress)
        {
            Write-ADTLogEntry -Message "Port [$($config.PortName)] already exists but points to [$($port.PrinterHostAddress)] instead of [$($config.PrinterHostAddress)]." -Severity 2
        }

        ## 3. Queue settings and DevMode via PrintBrm.
        $exportFile = Join-Path -Path $adtSession.DirFiles -ChildPath 'printer.printerExport'
        $printBrm = Join-Path -Path $system32 -ChildPath 'spool\tools\PrintBrm.exe'
        if ($config.UsePrintBrm -and (Test-Path -LiteralPath $exportFile) -and (Test-Path -LiteralPath $printBrm))
        {
            Write-ADTLogEntry -Message 'Restoring printer settings with PrintBrm.'
            $stagedExport = Join-Path -Path $staging -ChildPath 'printer.printerExport'
            Copy-Item -LiteralPath $exportFile -Destination $stagedExport -Force
            $result = Start-ADTProcess -FilePath $printBrm -ArgumentList '-R', '-F', $stagedExport, '-O', 'FORCE' -CreateNoWindow -IgnoreExitCodes '*' -PassThru -Timeout (New-TimeSpan -Minutes 10) -TimeoutAction Continue
            Write-ADTLogEntry -Message "PrintBrm exit code [$($result.ExitCode)]."

            if ($config.SourcePrinterName -ne $config.PrinterName -and !(Get-PackagedPrinter -Name $config.PrinterName))
            {
                $restored = Get-PackagedPrinter -Name $config.SourcePrinterName
                if ($restored)
                {
                    Write-ADTLogEntry -Message "Renaming [$($config.SourcePrinterName)] to [$($config.PrinterName)]."
                    Rename-Printer -InputObject $restored -NewName $config.PrinterName -ErrorAction Stop
                }
            }
        }

        ## 4. Make sure the queue exists and uses the packaged driver and port.
        $printer = Get-PackagedPrinter -Name $config.PrinterName
        if (!$printer)
        {
            Write-ADTLogEntry -Message "Creating printer [$($config.PrinterName)]."
            $params = @{ Name = $config.PrinterName; DriverName = $config.DriverName; PortName = $config.PortName }
            if ($config.Comment) { $params.Comment = $config.Comment }
            if ($config.Location) { $params.Location = $config.Location }
            Add-Printer @params -ErrorAction Stop
            # Windows 11 Ready Print may prefer its own IPP driver for new printers; make sure the packaged driver is used.
            $printer = Get-PackagedPrinter -Name $config.PrinterName
            if ($printer -and $printer.DriverName -ne $config.DriverName)
            {
                Write-ADTLogEntry -Message "Windows chose driver [$($printer.DriverName)]; switching to [$($config.DriverName)]." -Severity 2
                Set-Printer -InputObject $printer -DriverName $config.DriverName -ErrorAction Stop
            }
        }
        elseif ($printer.PortName -ne $config.PortName -or $printer.DriverName -ne $config.DriverName)
        {
            Write-ADTLogEntry -Message "Updating printer [$($config.PrinterName)] to use port [$($config.PortName)] and driver [$($config.DriverName)]."
            Set-Printer -InputObject $printer -PortName $config.PortName -DriverName $config.DriverName -ErrorAction Stop
        }

        ## 5. Vendor preferences via PrintUI.
        $settingsDat = Join-Path -Path $adtSession.DirFiles -ChildPath 'printer_settings.dat'
        if ($config.UseSettingsDat -and (Test-Path -LiteralPath $settingsDat))
        {
            Write-ADTLogEntry -Message 'Restoring printer preferences with PrintUI.'
            $stagedDat = Join-Path -Path $staging -ChildPath 'printer_settings.dat'
            Copy-Item -LiteralPath $settingsDat -Destination $stagedDat -Force
            Start-ADTProcess -FilePath (Join-Path -Path $system32 -ChildPath 'rundll32.exe') -ArgumentList 'printui.dll,PrintUIEntry', '/Sr', '/n', $config.PrinterName, '/a', $stagedDat, 'g', 'd', 'r', '/q' -CreateNoWindow -IgnoreExitCodes '*' -Timeout (New-TimeSpan -Minutes 2) -TimeoutAction Continue
        }

        ## 6. Basic print configuration.
        $printConfig = @{}
        if ($config.DuplexingMode) { $printConfig.DuplexingMode = $config.DuplexingMode }
        if ($config.PaperSize) { $printConfig.PaperSize = $config.PaperSize }
        if ($null -ne $config.Color) { $printConfig.Color = [System.Boolean]$config.Color }
        if ($printConfig.Count)
        {
            # Set-PrintConfiguration can wait on an unreachable printer, so it gets its own process and a time limit.
            $job = Start-Job -ScriptBlock { param ($Name, $Settings) Set-PrintConfiguration -PrinterName $Name @Settings -ErrorAction Stop } -ArgumentList $config.PrinterName, $printConfig
            if (Wait-Job -Job $job -Timeout 60)
            {
                try { Receive-Job -Job $job -ErrorAction Stop | Out-Null; Write-ADTLogEntry -Message 'Print configuration applied.' }
                catch { Write-ADTLogEntry -Message "Print configuration not applied: $($_.Exception.Message)" -Severity 2 }
            }
            else
            {
                Stop-Job -Job $job
                Write-ADTLogEntry -Message 'Print configuration skipped: no answer within 60 seconds (printer unreachable?).' -Severity 2
            }
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }

        ## 7. Default printer (a per-user setting, so it is applied for the signed-in user).
        if ($config.IsDefault)
        {
            Invoke-ADTAllUsersRegistryAction -ScriptBlock {
                Set-ADTRegistryKey -LiteralPath 'HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows' -Name 'LegacyDefaultPrinterMode' -Value 1 -Type DWord -SID $_.SID
            }
            $printUi = 'printui.dll,PrintUIEntry', '/y', '/q', '/n', $config.PrinterName
            try
            {
                if ([System.Security.Principal.WindowsIdentity]::GetCurrent().IsSystem)
                {
                    Start-ADTProcessAsUser -FilePath (Join-Path -Path $system32 -ChildPath 'rundll32.exe') -ArgumentList $printUi -IgnoreExitCodes '*' -Timeout (New-TimeSpan -Minutes 1) -TimeoutAction Continue
                }
                else
                {
                    Start-ADTProcess -FilePath (Join-Path -Path $system32 -ChildPath 'rundll32.exe') -ArgumentList $printUi -IgnoreExitCodes '*' -Timeout (New-TimeSpan -Minutes 1) -TimeoutAction Continue
                }
                Write-ADTLogEntry -Message "[$($config.PrinterName)] set as the default printer."
            }
            catch
            {
                Write-ADTLogEntry -Message "Default printer not set (no signed-in user?): $($_.Exception.Message)" -Severity 2
            }
        }

        ## 8. Marker for the Intune detection script.
        $key = Get-PackagedPrinterRegistryKey -Name $config.PrinterName -Writable $true
        try
        {
            $key.SetValue('Version', [System.String]$config.AppVersion)
            $key.SetValue('PortName', [System.String]$config.PortName)
            $key.SetValue('InstalledOn', (Get-Date).ToString('s'))
        }
        finally { $key.Close() }

        if (!(Get-PackagedPrinter -Name $config.PrinterName))
        {
            throw "Printer [$($config.PrinterName)] is not present after the installation."
        }
        Write-ADTLogEntry -Message "Printer [$($config.PrinterName)] installed."
    }
    finally
    {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Uninstall-PackagedPrinter
{
    $config = Get-PackagedPrinterConfig

    $printer = Get-PackagedPrinter -Name $config.PrinterName
    if ($printer)
    {
        Write-ADTLogEntry -Message "Removing printer [$($config.PrinterName)]."
        Remove-Printer -InputObject $printer -ErrorAction Stop
    }

    $others = @(Get-Printer -ErrorAction SilentlyContinue)
    if (!($others | Where-Object { $_.PortName -eq $config.PortName }) -and (Get-PrinterPort -Name $config.PortName -ErrorAction SilentlyContinue))
    {
        Write-ADTLogEntry -Message "Removing printer port [$($config.PortName)]."
        try { Remove-PrinterPort -Name $config.PortName -ErrorAction Stop }
        catch { Write-ADTLogEntry -Message "Port not removed: $($_.Exception.Message)" -Severity 2 }
    }

    # Only remove a driver this package brought along; built-in Windows drivers stay.
    if ($config.DriverFolder -and !($others | Where-Object { $_.DriverName -eq $config.DriverName }) -and (Get-PrinterDriver -Name $config.DriverName -ErrorAction SilentlyContinue))
    {
        Write-ADTLogEntry -Message "Removing printer driver [$($config.DriverName)]."
        try { Remove-PrinterDriver -Name $config.DriverName -ErrorAction Stop }
        catch { Write-ADTLogEntry -Message "Driver not removed (it may still be in use until the spooler restarts): $($_.Exception.Message)" -Severity 2 }
    }

    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try { $base.DeleteSubKeyTree("SOFTWARE\IntunePrinterPackager\Printers\$($config.PrinterName)", $false) } catch { }
    Write-ADTLogEntry -Message "Printer [$($config.PrinterName)] removed."
}

function Install-ADTDeployment
{
    [CmdletBinding()]
    param
    (
    )

    ##================================================
    ## MARK: Pre-Install
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"
    Show-ADTInstallationProgress -StatusMessage "Installing printer $($adtSession.AppName)..."

    ##================================================
    ## MARK: Install
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType
    Install-PackagedPrinter

    ##================================================
    ## MARK: Post-Install
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"
    Show-ADTInstallationPrompt -Message "Printer $($adtSession.AppName) is installed." -ButtonRightText 'OK' -NoWait
}

function Uninstall-ADTDeployment
{
    [CmdletBinding()]
    param
    (
    )

    ##================================================
    ## MARK: Pre-Uninstall
    ##================================================
    $adtSession.InstallPhase = "Pre-$($adtSession.DeploymentType)"
    Show-ADTInstallationProgress -StatusMessage "Removing printer $($adtSession.AppName)..."

    ##================================================
    ## MARK: Uninstall
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType
    Uninstall-PackagedPrinter

    ##================================================
    ## MARK: Post-Uninstallation
    ##================================================
    $adtSession.InstallPhase = "Post-$($adtSession.DeploymentType)"
}

function Repair-ADTDeployment
{
    [CmdletBinding()]
    param
    (
    )

    ##================================================
    ## MARK: Repair
    ##================================================
    $adtSession.InstallPhase = $adtSession.DeploymentType
    Show-ADTInstallationProgress -StatusMessage "Repairing printer $($adtSession.AppName)..."
    Install-PackagedPrinter
}
'@
}

function New-DetectionScript
{
    param ([string]$PrinterName, [string]$PortName, [string]$Version)
    return @"
# Intune detection script: printer is installed, uses the packaged port, and package version $Version (or newer) was applied.
`$printerName = $(ConvertTo-PSLiteral $PrinterName)
`$portName = $(ConvertTo-PSLiteral $PortName)
`$version = $(ConvertTo-PSLiteral $Version)

`$printer = Get-Printer -ErrorAction SilentlyContinue | Where-Object { `$_.Name -eq `$printerName } | Select-Object -First 1
if (-not `$printer -or `$printer.PortName -ne `$portName) { exit 1 }

`$base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
`$key = `$base.OpenSubKey("SOFTWARE\IntunePrinterPackager\Printers\`$printerName")
if (-not `$key) { exit 1 }
`$installed = [string]`$key.GetValue('Version')
`$key.Close()

`$a = `$null; `$b = `$null
if ([version]::TryParse(`$installed, [ref]`$a) -and [version]::TryParse(`$version, [ref]`$b)) { `$ok = `$a -ge `$b } else { `$ok = `$installed -eq `$version }
if (-not `$ok) { exit 1 }

Write-Output "Printer '`$printerName' is installed (package `$installed)."
exit 0
"@
}

function Write-TextFile
{
    param ([string]$Path, [string]$Content, [switch]$Bom, [switch]$Ascii)
    $Content = $Content -replace "`r?`n", "`r`n"
    $encoding = if ($Ascii) { New-Object Text.ASCIIEncoding } else { New-Object Text.UTF8Encoding($Bom.IsPresent) }
    [IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Save-PrinterCapture
{
    # Captures everything needed to rebuild the printer elsewhere: driver files, PrintBrm export,
    # PrintUI settings and the basic print configuration. Used for packages and for the printer library.
    param ($Printer, [string]$Destination, [bool]$IncludeDriver = $true, [bool]$IncludePreferences = $true)

    $p = $Printer
    New-Item -Path $Destination -ItemType Directory -Force | Out-Null
    $driverFolder = ''
    $driverInf = ''
    $builtInDriver = $false
    if ($IncludeDriver)
    {
        Set-Step "Collecting driver files for '$($p.DriverName)'..." 15
        $driverDir = Join-Path $Destination 'Driver'
        $driverInf = Export-PrinterDriverFiles -DriverName $p.DriverName -InfPath $p.InfPath -DestinationDir $driverDir
        if ($driverInf -eq 'INBOX')
        {
            Remove-Item -LiteralPath $driverDir -Recurse -Force -ErrorAction SilentlyContinue
            $driverInf = ''
            $builtInDriver = $true
        }
        elseif ($driverInf)
        {
            $driverFolder = 'Driver'
            $mb = (Get-ChildItem -LiteralPath $driverDir -Recurse -File | Measure-Object -Property Length -Sum).Sum / 1MB
            Write-Log ("Driver package collected ({0:N1} MB)." -f $mb) 'OK'
        }
        else
        {
            Remove-Item -LiteralPath $driverDir -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "No driver package found for '$($p.DriverName)'. Devices must already have this driver." 'WARN'
        }
    }

    $usePrintBrm = $false
    $useSettingsDat = $false
    $duplex = ''; $paper = ''; $color = $null
    if ($IncludePreferences)
    {
        Set-Step 'Capturing printer settings (PrintBrm)...' 35
        $usePrintBrm = Export-PrinterPreferences -PrinterName $p.Name -PortName $p.PortName -KeepPort ($p.PortKind -eq 'TCP/IP') -DriverName $p.DriverName -DestinationFile (Join-Path $Destination 'printer.printerExport')

        Set-Step 'Capturing printer preferences (PrintUI)...' 50
        $useSettingsDat = Export-PrinterSettingsDat -PrinterName $p.Name -DestinationFile (Join-Path $Destination 'printer_settings.dat')

        Set-Step 'Reading print configuration...' 55
        $cfg = Invoke-WithTimeout -Script 'param($Name) Get-PrintConfiguration -PrinterName $Name -ErrorAction Stop' -Parameters @{ Name = $p.Name } -TimeoutSeconds 20
        $cfg = @($cfg) | Select-Object -First 1
        if ($cfg)
        {
            $duplex = "$($cfg.DuplexingMode)"
            $paper = "$($cfg.PaperSize)"
            if ($null -ne $cfg.Color) { $color = [bool]$cfg.Color }
        }
        else { Write-Log 'Print configuration could not be read (printer offline?); skipped.' 'WARN' }
    }

    $capture = [ordered]@{
        CapturedOn     = (Get-Date).ToString('s')
        CapturedFrom   = $env:COMPUTERNAME
        DriverFolder   = $driverFolder
        DriverInf      = "$driverInf"
        BuiltInDriver  = $builtInDriver
        UsePrintBrm    = $usePrintBrm
        UseSettingsDat = $useSettingsDat
        DuplexingMode  = $duplex
        PaperSize      = $paper
        Color          = $color
    }
    Write-TextFile -Path (Join-Path $Destination 'capture.json') -Content ($capture | ConvertTo-Json -Depth 4)
}

function Save-PrinterToLibrary
{
    # One folder per printer: capture.json + captured files, printer.json (printer info + settings), Logo.png.
    # Entries that already come from the library (printer not on this PC) only get settings and logo updated.
    param ([object[]]$Requests, [string]$LibraryDir)
    $ErrorActionPreference = 'Stop'
    if (-not $LibraryDir) { throw 'Choose a library folder first.' }
    New-Item -Path $LibraryDir -ItemType Directory -Force | Out-Null
    $results = New-Object System.Collections.Generic.List[object]
    $n = $Requests.Count
    for ($i = 0; $i -lt $n; $i++)
    {
        $req = $Requests[$i]
        $p = $req.Printer
        $name = "$($req.TargetName)".Trim(); if (-not $name) { $name = $p.Name }
        $script:ProgressRange = @([int](100 * $i / $n), [int](100 * ($i + 1) / $n))
        try
        {
            if ($name -match '[\\,!]') { throw 'The printer name cannot contain \ , or !.' }
            $target = Join-Path $LibraryDir (ConvertTo-SafeFileName $name)
            Set-Step "Saving $name to the library ($($i + 1) of $n)..." 0
            if ($p.CaptureDir)
            {
                if ($p.CaptureDir -ne $target)
                {
                    if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
                    # Renamed inside the same library: move, so no duplicate stays behind. From another library: copy.
                    if ((Split-Path -Parent $p.CaptureDir) -eq $LibraryDir.TrimEnd('\')) { Move-Item -LiteralPath $p.CaptureDir -Destination $target }
                    else { Copy-Item -LiteralPath $p.CaptureDir -Destination $target -Recurse -Force }
                }
            }
            else
            {
                $tmp = "$target.saving"
                if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force }
                Save-PrinterCapture -Printer $p -Destination $tmp -IncludeDriver ([bool]$req.IncludeDriver) -IncludePreferences ([bool]$req.IncludePreferences)
                if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Recurse -Force }
                Move-Item -LiteralPath $tmp -Destination $target
            }
            $logoNew = Join-Path $target 'Logo.new.png'
            Convert-LogoToIcon -Source $req.LogoPath -Destination $logoNew
            Move-Item -LiteralPath $logoNew -Destination (Join-Path $target 'Logo.png') -Force

            $printerInfo = [ordered]@{}
            foreach ($prop in 'Name', 'DriverName', 'Manufacturer', 'PortName', 'PortKind', 'HostAddress', 'Protocol', 'PortNumber', 'LprQueueName', 'SnmpEnabled', 'SnmpCommunity', 'SnmpIndex', 'Location', 'Comment', 'Brand')
            {
                $printerInfo[$prop] = $p.$prop
            }
            if ($p.SourceName) { $printerInfo.Name = $p.SourceName }
            $entry = [ordered]@{
                SchemaVersion = 1
                SavedOn       = (Get-Date).ToString('s')
                SavedBy       = "$env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME"
                Printer       = $printerInfo
                Settings      = [ordered]@{
                    TargetName  = $name
                    HostAddress = "$($req.HostAddress)".Trim()
                    AppName     = "$($req.AppName)"
                    Advanced    = $req.Advanced
                }
            }
            Write-TextFile -Path (Join-Path $target 'printer.json') -Content ($entry | ConvertTo-Json -Depth 6)
            Write-Log "Saved '$name' to the library." 'OK'
            $results.Add([pscustomobject]@{ Printer = $name; Status = 'Library'; Version = ''; Url = ''; Logo = (Join-Path $target 'Logo.png'); Folder = $target; Error = '' })
        }
        catch
        {
            Write-Log "$name failed: $($_.Exception.Message)" 'ERROR'
            $results.Add([pscustomobject]@{ Printer = $name; Status = 'Failed'; Version = ''; Url = ''; Logo = $req.LogoPath; Folder = ''; Error = $_.Exception.Message })
        }
        finally { $script:ProgressRange = $null }
    }
    Set-Step 'Done.' 100
    return [pscustomobject]@{ __Result = $true; Results = $results.ToArray() }
}

function Get-PrinterLibrary
{
    param ([string]$LibraryDir)
    Set-Step 'Reading the printer library...'
    $items = New-Object System.Collections.Generic.List[object]
    if ($LibraryDir -and (Test-Path -LiteralPath $LibraryDir))
    {
        foreach ($dir in (Get-ChildItem -LiteralPath $LibraryDir -Directory | Where-Object { $_.Name -notlike '*.saving' }))
        {
            $file = Join-Path $dir.FullName 'printer.json'
            if (-not (Test-Path -LiteralPath $file) -or -not (Test-Path -LiteralPath (Join-Path $dir.FullName 'capture.json'))) { continue }
            try
            {
                $entry = [IO.File]::ReadAllText($file) | ConvertFrom-Json
                $info = $entry.Printer
                $logo = Join-Path $dir.FullName 'Logo.png'
                if (-not (Test-Path -LiteralPath $logo)) { $logo = Join-Path $Ctx.LogoDir (Get-PrinterBrand -Text @($info.Manufacturer, $info.DriverName, $info.Name)).File }
                $address = if ($entry.Settings.HostAddress) { "$($entry.Settings.HostAddress)" } else { "$($info.HostAddress)" }
                $items.Add([pscustomobject]@{
                    Key           = "lib:$($dir.Name)"
                    Name          = "$($entry.Settings.TargetName)"
                    DriverName    = "$($info.DriverName)"
                    Manufacturer  = "$($info.Manufacturer)"
                    PortName      = "$($info.PortName)"
                    PortKind      = "$($info.PortKind)"
                    HostAddress   = "$($info.HostAddress)"
                    Address       = $address
                    Protocol      = [int]$info.Protocol
                    PortNumber    = [int]$info.PortNumber
                    LprQueueName  = "$($info.LprQueueName)"
                    SnmpEnabled   = [bool]$info.SnmpEnabled
                    SnmpCommunity = "$($info.SnmpCommunity)"
                    SnmpIndex     = [int]$info.SnmpIndex
                    Location      = "$($info.Location)"
                    Comment       = "$($info.Comment)"
                    InfPath       = ''
                    Brand         = "$($info.Brand)"
                    LogoPath      = $logo
                    CaptureDir    = $dir.FullName
                    SourceName    = "$($info.Name)"
                    SavedSettings = $entry.Settings
                    SavedOn       = "$($entry.SavedOn)"
                })
            }
            catch { Write-Log "Library entry '$($dir.Name)' could not be read: $($_.Exception.Message)" 'WARN' }
        }
    }
    $sorted = @($items | Sort-Object Name)
    Write-Log "Library: $($sorted.Count) printer(s) in $LibraryDir" 'OK'
    return [pscustomobject]@{ __Result = $true; Printers = $sorted }
}

function New-PrinterPackage
{
    param ([hashtable]$Request)

    $ErrorActionPreference = 'Stop'
    $p = $Request.Printer
    $targetName = "$($Request.TargetName)".Trim()
    if (-not $targetName) { $targetName = $p.Name }
    if ($targetName -match '[\\,!]') { throw 'The printer name on devices cannot contain \ , or !.' }
    $appVendor = "$($Request.AppVendor)".Trim(); if (-not $appVendor) { $appVendor = 'Company' }
    $appName = "$($Request.AppName)".Trim(); if (-not $appName) { $appName = "Printer - $targetName" }
    $appVersion = "$($Request.AppVersion)".Trim(); if (-not $appVersion) { $appVersion = '1.0.0' }
    if ($appVersion -notmatch '^[0-9A-Za-z.\-_]+$') { throw "Version '$appVersion' can only contain letters, digits, dots and dashes." }
    $outDir = "$($Request.OutputDir)".Trim()
    if (-not $outDir) { throw 'Choose an output folder.' }

    # Port: keep the existing TCP/IP port, or build a standard one from the IP/hostname that was entered.
    $hostAddress = "$($Request.HostAddress)".Trim()
    if (-not $hostAddress) { throw "Enter the IP address or hostname of '$($p.Name)'. It uses a $($p.PortKind) port, which can't be recreated on other devices." }
    if ($hostAddress -notmatch '^[A-Za-z0-9.\-:_\[\]]+$') { throw "'$hostAddress' is not a valid IP address or hostname." }
    $adv = if ($Request.Advanced) { $Request.Advanced } else { @{} }
    $keepPort = ($p.PortKind -eq 'TCP/IP' -and $hostAddress -eq $p.HostAddress)
    $protocol = if ($keepPort) { [int]$p.Protocol } else { 1 }
    $portNumber = if ($keepPort) { [int]$p.PortNumber } else { 9100 }
    $lprQueue = if ($keepPort) { "$($p.LprQueueName)" } else { '' }
    $snmpEnabled = if ($keepPort) { [bool]$p.SnmpEnabled } else { $false }
    $snmpCommunity = if ($p.SnmpCommunity) { "$($p.SnmpCommunity)" } else { 'public' }
    if ($adv.Protocol) { $protocol = [int]$adv.Protocol }
    if ($adv.PortNumber) { $portNumber = [int]$adv.PortNumber }
    if ($adv.LprQueueName) { $lprQueue = "$($adv.LprQueueName)" }
    if ($null -ne $adv.SnmpEnabled) { $snmpEnabled = [bool]$adv.SnmpEnabled }
    if ($adv.SnmpCommunity) { $snmpCommunity = "$($adv.SnmpCommunity)" }
    $portChanged = $keepPort -and ($protocol -ne [int]$p.Protocol -or $portNumber -ne [int]$p.PortNumber -or $lprQueue -ne "$($p.LprQueueName)")
    $portName = if ($keepPort -and -not $portChanged) { $p.PortName }
        elseif ($protocol -eq 2) { "IP_${hostAddress}_LPR" }
        elseif ($portNumber -ne 9100) { "IP_${hostAddress}_$portNumber" }
        else { "IP_$hostAddress" }
    $location = if ($null -ne $adv.Location) { "$($adv.Location)" } else { "$($p.Location)" }
    $comment = if ($null -ne $adv.Comment) { "$($adv.Comment)" } else { "$($p.Comment)" }

    $displayName = "$appVendor - $appName"
    $folderName = ConvertTo-SafeFileName "$appVendor - $appName - $appVersion"
    $root = Join-Path $outDir $folderName
    $packageDir = Join-Path $root 'Package'
    $filesDir = Join-Path $packageDir 'Files'

    Set-Step 'Preparing...' 2
    $template = Resolve-PsadtTemplate
    $intuneWinUtil = $null
    if ($Request.CreateIntuneWin)
    {
        try { $intuneWinUtil = Resolve-IntuneWinAppUtil }
        catch
        {
            # Offline "save files": keep going without the .intunewin, but always make the .zip.
            if (-not $Request.IntuneWinOptional) { throw }
            Write-Log "$($_.Exception.Message) Saving without .intunewin." 'WARN'
            $Request.CreateIntuneWin = $false
            $Request.CreateZip = $true
        }
    }

    if (Test-Path -LiteralPath $root)
    {
        $ours = (Test-Path -LiteralPath (Join-Path $root 'Package\Invoke-AppDeployToolkit.ps1')) -or (Test-Path -LiteralPath (Join-Path $root 'Intune-Upload-Info.txt'))
        if (-not $ours -and (Get-ChildItem -LiteralPath $root -Force | Select-Object -First 1))
        {
            throw "The folder '$root' already exists and was not made by Printer Packager. Choose another output folder or version."
        }
        Write-Log "Replacing the previous build in $root" 'WARN'
        Remove-Item -LiteralPath $root -Recurse -Force
    }
    New-Item -Path $root -ItemType Directory -Force | Out-Null

    Set-Step "Copying PSAppDeployToolkit template..." 8
    Copy-Item -LiteralPath $template -Destination $packageDir -Recurse -Force
    Get-ChildItem -LiteralPath $packageDir -Recurse -Filter 'Add * Here.txt' | Remove-Item -Force
    New-Item -Path $filesDir -ItemType Directory -Force | Out-Null

    # Use the library capture, or capture the printer on this PC now.
    $tempCapture = $null
    if ($Request.CaptureDir)
    {
        Set-Step 'Using the saved printer from the library...' 15
        $captureDir = $Request.CaptureDir
    }
    else
    {
        $tempCapture = Join-Path $env:SystemRoot ("Temp\PrinterPackager_cap_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $captureDir = $tempCapture
        [void](Save-PrinterCapture -Printer $p -Destination $captureDir -IncludeDriver ([bool]$Request.IncludeDriver) -IncludePreferences ([bool]$Request.IncludePreferences))
    }
    try
    {
        $cap = [IO.File]::ReadAllText((Join-Path $captureDir 'capture.json')) | ConvertFrom-Json
        foreach ($item in 'Driver', 'printer.printerExport', 'printer_settings.dat')
        {
            $src = Join-Path $captureDir $item
            if (Test-Path -LiteralPath $src) { Copy-Item -LiteralPath $src -Destination $filesDir -Recurse -Force }
        }
    }
    finally
    {
        if ($tempCapture) { Remove-Item -LiteralPath $tempCapture -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $driverFolder = "$($cap.DriverFolder)"
    $driverInf = "$($cap.DriverInf)"
    $builtInDriver = [bool]$cap.BuiltInDriver
    $usePrintBrm = [bool]$cap.UsePrintBrm -and (Test-Path -LiteralPath (Join-Path $filesDir 'printer.printerExport'))
    $useSettingsDat = [bool]$cap.UseSettingsDat -and (Test-Path -LiteralPath (Join-Path $filesDir 'printer_settings.dat'))
    $duplex = "$($cap.DuplexingMode)"
    $paper = "$($cap.PaperSize)"
    $color = $cap.Color
    if ($adv.Duplex) { $duplex = "$($adv.Duplex)" }
    if ($adv.PaperSize) { $paper = "$($adv.PaperSize)" }
    if ($adv.Color -eq 'Color') { $color = $true } elseif ($adv.Color -eq 'Mono') { $color = $false }

    Set-Step 'Writing deployment script...' 60
    $config = [ordered]@{
        AppName            = $appName
        AppVendor          = $appVendor
        AppVersion         = $appVersion
        PrinterName        = $targetName
        SourcePrinterName  = if ($p.SourceName) { $p.SourceName } else { $p.Name }
        DriverName         = $p.DriverName
        DriverFolder       = $driverFolder
        DriverInf          = "$driverInf"
        PortName           = $portName
        PrinterHostAddress = $hostAddress
        Protocol           = $protocol
        PortNumber         = $portNumber
        LprQueueName       = $lprQueue
        SnmpEnabled        = $snmpEnabled
        SnmpCommunity      = $snmpCommunity
        SnmpIndex          = if ($p.SnmpIndex) { [int]$p.SnmpIndex } else { 1 }
        Comment            = $comment
        Location           = $location
        IsDefault          = [bool]$Request.SetDefault
        UsePrintBrm        = $usePrintBrm
        UseSettingsDat     = $useSettingsDat
        DuplexingMode      = $duplex
        PaperSize          = $paper
        Color              = $color
    }
    Write-TextFile -Path (Join-Path $filesDir 'printer_config.json') -Content ($config | ConvertTo-Json -Depth 4)

    $adtScript = Join-Path $packageDir 'Invoke-AppDeployToolkit.ps1'
    $code = [IO.File]::ReadAllText($adtScript)
    foreach ($pair in @(@('AppVendor', $appVendor), @('AppName', $appName), @('AppVersion', $appVersion), @('AppScriptAuthor', 'Intune Printer Packager'), @('AppScriptDate', (Get-Date -Format 'yyyy-MM-dd'))))
    {
        $code = Set-TemplateValue -Code $code -Name $pair[0] -Value $pair[1]
    }
    $start = $code.IndexOf('function Install-ADTDeployment')
    $end = [regex]::Match($code, '(?m)^##=+\r?\n## MARK: Initialization')
    if ($start -lt 0 -or -not $end.Success -or $end.Index -le $start)
    {
        throw 'This PSADT template is not supported (Install-ADTDeployment / MARK: Initialization not found).'
    }
    $code = $code.Substring(0, $start) + (Get-DeploymentFunctions) + "`r`n`r`n`r`n" + $code.Substring($end.Index)
    Write-TextFile -Path $adtScript -Content $code -Bom

    $parseErrors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($adtScript, [ref]$null, [ref]$parseErrors)
    if ($parseErrors) { throw "The generated Invoke-AppDeployToolkit.ps1 has syntax errors: $($parseErrors[0].Message)" }
    Write-Log 'Invoke-AppDeployToolkit.ps1 written.' 'OK'

    Set-Step 'Adding logo...' 65
    $logo = Join-Path $root 'Logo.png'
    Convert-LogoToIcon -Source $Request.LogoPath -Destination $logo
    Copy-Item -LiteralPath $logo -Destination (Join-Path $packageDir 'Assets\AppIcon.png') -Force
    Write-Log "Logo: $(Split-Path -Leaf $Request.LogoPath)" 'OK'

    $detection = New-DetectionScript -PrinterName $targetName -PortName $portName -Version $appVersion
    Write-TextFile -Path (Join-Path $root 'Detection.ps1') -Content $detection -Ascii
    Write-TextFile -Path (Join-Path $packageDir 'Detection.ps1') -Content $detection -Ascii

    foreach ($action in 'Install', 'Uninstall')
    {
        $cmd = @"
@echo off
rem $action this printer on this computer. Windows asks for administrator rights.
set "PKG=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "`$p = Start-Process -FilePath (Join-Path `$env:PKG 'Invoke-AppDeployToolkit.exe') -ArgumentList '-DeploymentType $action -DeployMode Interactive' -Verb RunAs -Wait -PassThru; exit `$p.ExitCode"
if errorlevel 1 (echo. & echo $action failed with exit code %errorlevel%. See C:\Windows\Logs\Software for the log. & pause)
"@
        Write-TextFile -Path (Join-Path $packageDir "$action-Printer.cmd") -Content $cmd -Ascii
    }

    $installCmd = 'Invoke-AppDeployToolkit.exe -DeploymentType Install -DeployMode Silent'
    $uninstallCmd = 'Invoke-AppDeployToolkit.exe -DeploymentType Uninstall -DeployMode Silent'

    $intuneFile = $null
    if ($Request.CreateIntuneWin)
    {
        Set-Step 'Building .intunewin package...' 75
        $r = Invoke-Native -FilePath $intuneWinUtil -Arguments "-c `"$packageDir`" -s Invoke-AppDeployToolkit.exe -o `"$root`" -q" -TimeoutSeconds 1800
        $built = Join-Path $root 'Invoke-AppDeployToolkit.intunewin'
        if (-not (Test-Path -LiteralPath $built)) { throw "IntuneWinAppUtil failed (exit $($r.ExitCode)). $($r.Output)" }
        $intuneFile = Join-Path $root ((ConvertTo-SafeFileName $displayName) + '.intunewin')
        Move-Item -LiteralPath $built -Destination $intuneFile -Force
        Write-Log ("Intune package: {0} ({1:N1} MB)" -f (Split-Path -Leaf $intuneFile), ((Get-Item -LiteralPath $intuneFile).Length / 1MB)) 'OK'
    }

    $zipFile = $null
    if ($Request.CreateZip)
    {
        Set-Step 'Creating .zip for manual installs...' 90
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zipFile = Join-Path $root ((ConvertTo-SafeFileName $displayName) + '.zip')
        [IO.Compression.ZipFile]::CreateFromDirectory($packageDir, $zipFile, [IO.Compression.CompressionLevel]::Optimal, $false)
        Write-Log ("Zip: {0} ({1:N1} MB)" -f (Split-Path -Leaf $zipFile), ((Get-Item -LiteralPath $zipFile).Length / 1MB)) 'OK'
    }

    $info = @"
Intune Win32 app - $displayName
Created $(Get-Date -Format 'yyyy-MM-dd HH:mm') with Intune Printer Packager (PSADT $($Ctx.PsadtVersion))

App package file   : $(if ($intuneFile) { Split-Path -Leaf $intuneFile } else { '(not built)' })
Name               : $displayName
Publisher          : $appVendor
App version        : $appVersion
Logo               : Logo.png

Install command    : $installCmd
Uninstall command  : $uninstallCmd
Install behavior   : System
Device restart     : No specific action

Requirements       : 64-bit, Windows 10 1809 or later
Detection rules    : Use a custom detection script -> Detection.ps1
                     Run script as 32-bit process on 64-bit clients: No
                     Enforce script signature check: No

Printer            : $targetName
Driver             : $($p.DriverName)$(if ($builtInDriver) { '  (built into Windows)' } elseif (-not $driverFolder) { '  (NOT included - must already be on the devices)' })
Port               : $portName -> $hostAddress
Default printer    : $(if ($Request.SetDefault) { 'Yes' } else { 'No' })

Manual install     : open the Package folder (or the .zip) and run Install-Printer.cmd
Logs on devices    : C:\Windows\Logs\Software
"@
    Write-TextFile -Path (Join-Path $root 'Intune-Upload-Info.txt') -Content $info -Bom

    Set-Step 'Done.' 100
    Write-Log "Package ready: $root" 'OK'
    return [pscustomobject]@{
        __Result         = $true
        Folder           = $root
        PackageDir       = $packageDir
        IntuneWin        = $intuneFile
        Zip              = $zipFile
        Logo             = $logo
        Detection        = Join-Path $root 'Detection.ps1'
        DisplayName      = $displayName
        Publisher        = $appVendor
        Version          = $appVersion
        InstallCommand   = $installCmd
        UninstallCommand = $uninstallCmd
        DriverIncluded   = [bool]$driverFolder -or $builtInDriver
    }
}


#region Intune (Microsoft Graph)
# Sign-in uses the Microsoft Graph Command Line Tools public client (same as Connect-MgGraph).
# Tokens live in memory only and are gone when the app closes.
# An own app registration (public client, redirect http://localhost) can be set as "ClientId" in settings.json.
$GraphClientId = if ($Ctx.ClientId) { $Ctx.ClientId } else { '14d82eec-204b-4c2f-b7e8-296a70dab67e' }
$GraphScopes = 'https://graph.microsoft.com/DeviceManagementApps.ReadWrite.All https://graph.microsoft.com/User.Read offline_access openid profile'
$GraphBase = 'https://graph.microsoft.com/beta'

function ConvertFrom-JwtPayload
{
    param ([string]$Token)
    $part = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    switch ($part.Length % 4) { 2 { $part += '==' } 3 { $part += '=' } }
    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($part)) | ConvertFrom-Json
}

function New-AuthObject
{
    param ($TokenResponse)
    $claims = ConvertFrom-JwtPayload $TokenResponse.access_token
    $auth = @{
        AccessToken  = $TokenResponse.access_token
        RefreshToken = $TokenResponse.refresh_token
        ExpiresOn    = (Get-Date).AddSeconds([int]$TokenResponse.expires_in - 120)
        TenantId     = $claims.tid
        User         = if ($claims.upn) { $claims.upn } else { $claims.unique_name }
        TenantName   = ''
    }
    try
    {
        $org = Invoke-RestMethod -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName' -Headers @{ Authorization = "Bearer $($auth.AccessToken)" }
        $auth.TenantName = @($org.value)[0].displayName
    }
    catch { }
    return $auth
}

function Invoke-TokenRequest
{
    param ([string]$Tenant, [hashtable]$Body)
    $Body.client_id = $GraphClientId
    return Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token" -Body $Body -ContentType 'application/x-www-form-urlencoded'
}

function Connect-IntuneBrowser
{
    # Authorization code + PKCE with a loopback redirect: the normal Microsoft sign-in page (MFA included) in the browser.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $tcp = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
    $tcp.Start(); $port = $tcp.LocalEndpoint.Port; $tcp.Stop()
    $redirect = "http://localhost:$port/"
    $bytes = New-Object byte[] 48
    [Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($bytes)
    $verifier = [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $sha = [Security.Cryptography.SHA256]::Create()
    $challenge = [Convert]::ToBase64String($sha.ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier))).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    $state = [guid]::NewGuid().ToString('N')

    $listener = New-Object Net.HttpListener
    $listener.Prefixes.Add($redirect)
    $listener.Start()
    try
    {
        $url = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?client_id=' + $GraphClientId +
            '&response_type=code&response_mode=query&prompt=select_account' +
            '&redirect_uri=' + [Uri]::EscapeDataString($redirect) +
            '&scope=' + [Uri]::EscapeDataString($GraphScopes) +
            '&code_challenge=' + $challenge + '&code_challenge_method=S256&state=' + $state
        Set-Step 'Sign in in the browser window that just opened...'
        Start-Process $url
        $pending = $listener.GetContextAsync()
        if (-not $pending.Wait([TimeSpan]::FromMinutes(5))) { throw 'Sign-in timed out.' }
        $context = $pending.Result
        $query = $context.Request.QueryString
        $ok = $query['code'] -and $query['state'] -eq $state
        $html = if ($ok) { 'Signed in. You can close this tab and go back to Printer Packager.' } else { "Sign-in failed: $($query['error_description'])" }
        $page = [Text.Encoding]::UTF8.GetBytes("<html><body style=`"font-family:Segoe UI,sans-serif;padding:40px`"><h3>$([Net.WebUtility]::HtmlEncode($html))</h3></body></html>")
        $context.Response.ContentType = 'text/html; charset=utf-8'
        $context.Response.OutputStream.Write($page, 0, $page.Length)
        $context.Response.Close()
        if (-not $ok) { throw "Sign-in failed: $($query['error_description'])" }
        $code = $query['code']
    }
    finally { $listener.Stop(); $listener.Close() }

    Set-Step 'Finishing sign-in...'
    $token = Invoke-TokenRequest -Tenant 'organizations' -Body @{ grant_type = 'authorization_code'; code = $code; redirect_uri = $redirect; code_verifier = $verifier; scope = $GraphScopes }
    $auth = New-AuthObject $token
    Write-Log "Signed in as $($auth.User) ($($auth.TenantName))." 'OK'
    return [pscustomobject]@{ __Result = $true; Auth = $auth }
}

function Connect-IntuneDeviceCode
{
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $dc = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode' -Body @{ client_id = $GraphClientId; scope = $GraphScopes } -ContentType 'application/x-www-form-urlencoded'
    Set-Step "Go to $($dc.verification_uri) and enter code $($dc.user_code)"
    if ($Sync) { $Sync.DeviceCode = $dc.user_code }
    $deadline = (Get-Date).AddSeconds([int]$dc.expires_in)
    while ((Get-Date) -lt $deadline)
    {
        Start-Sleep -Seconds ([Math]::Max(3, [int]$dc.interval))
        try
        {
            $token = Invoke-TokenRequest -Tenant 'organizations' -Body @{ grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; device_code = $dc.device_code }
            $auth = New-AuthObject $token
            Write-Log "Signed in as $($auth.User) ($($auth.TenantName))." 'OK'
            return [pscustomobject]@{ __Result = $true; Auth = $auth }
        }
        catch
        {
            $body = "$($_.ErrorDetails.Message)"
            if ($body -notmatch 'authorization_pending|slow_down') { throw "Sign-in failed: $body" }
        }
    }
    throw 'Sign-in timed out.'
}

function Update-GraphToken
{
    param ([hashtable]$Auth)
    if ((Get-Date) -lt $Auth.ExpiresOn) { return }
    $token = Invoke-TokenRequest -Tenant $Auth.TenantId -Body @{ grant_type = 'refresh_token'; refresh_token = $Auth.RefreshToken; scope = $GraphScopes }
    $Auth.AccessToken = $token.access_token
    if ($token.refresh_token) { $Auth.RefreshToken = $token.refresh_token }
    $Auth.ExpiresOn = (Get-Date).AddSeconds([int]$token.expires_in - 120)
}

function Invoke-Graph
{
    param ([hashtable]$Auth, [string]$Method = 'GET', [string]$Path, $Body)
    $uri = if ($Path -like 'https://*') { $Path } else { "$GraphBase/$Path" }
    $json = if ($null -ne $Body) { $Body | ConvertTo-Json -Depth 10 -Compress } else { $null }
    for ($attempt = 1; ; $attempt++)
    {
        Update-GraphToken $Auth
        try
        {
            $params = @{ Method = $Method; Uri = $uri; Headers = @{ Authorization = "Bearer $($Auth.AccessToken)" }; ContentType = 'application/json; charset=utf-8' }
            if ($json) { $params.Body = [Text.Encoding]::UTF8.GetBytes($json) }
            return Invoke-RestMethod @params
        }
        catch
        {
            $status = 0
            try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            if ($attempt -lt 5 -and ($status -eq 429 -or $status -ge 500)) { Start-Sleep -Seconds (3 * $attempt); continue }
            $detail = "$($_.ErrorDetails.Message)"
            if (-not $detail)
            {
                # Windows PowerShell 5.1 leaves ErrorDetails empty; read the response body ourselves.
                try
                {
                    $reader = New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $detail = $reader.ReadToEnd()
                    $reader.Dispose()
                }
                catch { }
            }
            try { $detail = ($detail | ConvertFrom-Json).error.message } catch { }
            if ($status -eq 403) { $detail = "Access denied. Your account needs the Intune Administrator or Application Manager role, and the tenant must allow 'DeviceManagementApps.ReadWrite.All'. ($detail)" }
            throw "Graph $Method $Path failed ($status): $detail"
        }
    }
}

function Get-PrinterAppTag
{
    param ([string]$PrinterName)
    return "PrinterPackager:$PrinterName"
}

function Get-PrinterApps
{
    # All Win32 apps that Printer Packager published, keyed by printer name.
    param ([hashtable]$Auth)
    $map = @{}
    # Same query as the IntuneWin32App module; Intune rejects $select together with this isof() filter.
    $next = "deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')"
    while ($next)
    {
        $page = Invoke-Graph -Auth $Auth -Path $next
        foreach ($app in @($page.value))
        {
            if ("$($app.notes)" -match '(?m)^PrinterPackager:(.+)$') { $map[$Matches[1].Trim()] = $app }
        }
        $next = $page.'@odata.nextLink'
    }
    return $map
}

function Get-IntuneWinInfo
{
    # Reads Detection.xml and extracts the encrypted payload from a .intunewin file.
    param ([string]$Path, [string]$ExtractTo)
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path)
    try
    {
        $meta = $zip.Entries | Where-Object { $_.FullName -match '(^|/)Metadata/Detection\.xml$' } | Select-Object -First 1
        if (-not $meta) { throw "$Path is not a valid .intunewin file." }
        $reader = New-Object IO.StreamReader($meta.Open())
        try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
        $info = $xml.ApplicationInfo
        $content = $zip.Entries | Where-Object { $_.FullName -match "(^|/)Contents/$([regex]::Escape($info.FileName))$" } | Select-Object -First 1
        $target = Join-Path $ExtractTo $info.FileName
        [IO.Compression.ZipFileExtensions]::ExtractToFile($content, $target, $true)
        return [pscustomobject]@{ Info = $info; EncryptedFile = $target }
    }
    finally { $zip.Dispose() }
}

function Wait-IntuneFileState
{
    param ([hashtable]$Auth, [string]$FileUri, [string]$Stage)
    for ($i = 0; $i -lt 120; $i++)
    {
        $file = Invoke-Graph -Auth $Auth -Path $FileUri
        switch -Wildcard ($file.uploadState)
        {
            "${Stage}Success" { return $file }
            "${Stage}Failed" { throw "Intune reported $Stage failed." }
            "${Stage}TimedOut" { throw "Intune reported $Stage timed out." }
        }
        Start-Sleep -Seconds 3
    }
    throw "Intune did not finish $Stage in time."
}

function Send-IntuneContent
{
    # contentVersion -> file -> Azure blob (6 MB blocks) -> commit -> committedContentVersion
    param ([hashtable]$Auth, [string]$AppId, [string]$IntuneWinPath, [int]$StartPercent, [int]$EndPercent, [string]$Label)
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("PrinterPackager_up_" + [guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -Path $tmp -ItemType Directory -Force | Out-Null
    try
    {
        $pkg = Get-IntuneWinInfo -Path $IntuneWinPath -ExtractTo $tmp
        $base = "deviceAppManagement/mobileApps/$AppId/microsoft.graph.win32LobApp/contentVersions"
        $version = Invoke-Graph -Auth $Auth -Method POST -Path $base -Body @{}
        $fileBody = [ordered]@{
            '@odata.type' = '#microsoft.graph.mobileAppContentFile'
            name          = $pkg.Info.FileName
            size          = [int64]$pkg.Info.UnencryptedContentSize
            sizeEncrypted = [int64](Get-Item -LiteralPath $pkg.EncryptedFile).Length
            manifest      = $null
            isDependency  = $false
        }
        $file = Invoke-Graph -Auth $Auth -Method POST -Path "$base/$($version.id)/files" -Body $fileBody
        $fileUri = "$base/$($version.id)/files/$($file.id)"
        $file = Wait-IntuneFileState -Auth $Auth -FileUri $fileUri -Stage 'azureStorageUriRequest'

        $chunkSize = 6MB
        $stream = [IO.File]::OpenRead($pkg.EncryptedFile)
        $ids = New-Object System.Collections.Generic.List[string]
        try
        {
            $count = [Math]::Max(1, [Math]::Ceiling($stream.Length / $chunkSize))
            $buffer = New-Object byte[] $chunkSize
            for ($i = 0; $i -lt $count; $i++)
            {
                $read = $stream.Read($buffer, 0, $chunkSize)
                $data = New-Object byte[] $read
                [Array]::Copy($buffer, $data, $read)
                $id = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($i.ToString('0000')))
                $ids.Add($id)
                for ($try = 1; ; $try++)
                {
                    try
                    {
                        $wc = New-Object Net.WebClient
                        $wc.Headers['x-ms-blob-type'] = 'BlockBlob'
                        [void]$wc.UploadData("$($file.azureStorageUri)&comp=block&blockid=$([Uri]::EscapeDataString($id))", 'PUT', $data)
                        $wc.Dispose()
                        break
                    }
                    catch { if ($try -ge 4) { throw "Upload to Azure storage failed: $($_.Exception.Message)" }; Start-Sleep -Seconds (2 * $try) }
                }
                $pct = $StartPercent + [int](($EndPercent - $StartPercent) * ($i + 1) / $count)
                Set-Step ("{0}: uploading {1:N0} of {2:N0} MB" -f $Label, ([Math]::Min($stream.Length, ($i + 1) * $chunkSize) / 1MB), ($stream.Length / 1MB)) $pct
            }
        }
        finally { $stream.Dispose() }

        $blockList = '<?xml version="1.0" encoding="utf-8"?><BlockList>' + (($ids | ForEach-Object { "<Latest>$_</Latest>" }) -join '') + '</BlockList>'
        $wc = New-Object Net.WebClient
        $wc.Headers['Content-Type'] = 'application/xml'
        [void]$wc.UploadString("$($file.azureStorageUri)&comp=blocklist", 'PUT', $blockList)
        $wc.Dispose()

        $enc = $pkg.Info.EncryptionInfo
        $commit = @{ fileEncryptionInfo = [ordered]@{
            encryptionKey        = $enc.EncryptionKey
            macKey               = $enc.MacKey
            initializationVector = $enc.InitializationVector
            mac                  = $enc.Mac
            profileIdentifier    = 'ProfileVersion1'
            fileDigest           = $enc.FileDigest
            fileDigestAlgorithm  = $enc.FileDigestAlgorithm
        } }
        Set-Step "${Label}: Intune is processing the upload..."
        Invoke-Graph -Auth $Auth -Method POST -Path "$fileUri/commit" -Body $commit | Out-Null
        [void](Wait-IntuneFileState -Auth $Auth -FileUri $fileUri -Stage 'commitFile')
        Invoke-Graph -Auth $Auth -Method PATCH -Path "deviceAppManagement/mobileApps/$AppId" -Body @{ '@odata.type' = '#microsoft.graph.win32LobApp'; committedContentVersion = $version.id } | Out-Null
    }
    finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

function Get-NextVersion
{
    # Keeps the requested version unless Intune already has the same or a newer one; then bumps the last number.
    param ([string]$Requested, [string]$Existing)
    $r = $null; $e = $null
    if (-not $Existing -or -not [version]::TryParse($Existing, [ref]$e)) { return $Requested }
    if ([version]::TryParse($Requested, [ref]$r) -and $r -gt $e) { return $Requested }
    $parts = @($e.Major, [Math]::Max(0, $e.Minor), [Math]::Max(0, $e.Build))
    $parts[2]++
    return ($parts -join '.')
}

function Publish-PrinterApp
{
    param ([hashtable]$Auth, $Package, [string]$PrinterName, $Existing, [string]$Description, [int]$StartPercent, [int]$EndPercent)
    $label = $PrinterName
    $detection = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Package.Detection))
    $icon = [Convert]::ToBase64String([IO.File]::ReadAllBytes($Package.Logo))
    $body = [ordered]@{
        '@odata.type'                  = '#microsoft.graph.win32LobApp'
        displayName                    = $Package.DisplayName
        description                    = $Description
        publisher                      = $Package.Publisher
        displayVersion                 = $Package.Version
        developer                      = 'Printer Packager'
        notes                          = Get-PrinterAppTag $PrinterName
        fileName                       = Split-Path -Leaf $Package.IntuneWin
        setupFilePath                  = 'Invoke-AppDeployToolkit.exe'
        installCommandLine             = $Package.InstallCommand
        uninstallCommandLine           = $Package.UninstallCommand
        applicableArchitectures        = 'x64'
        minimumSupportedWindowsRelease = '1809'
        installExperience              = [ordered]@{ runAsAccount = 'system'; deviceRestartBehavior = 'suppress'; maxRunTimeInMinutes = 30 }
        returnCodes                    = @(
            @{ returnCode = 0; type = 'success' }, @{ returnCode = 1707; type = 'success' }, @{ returnCode = 3010; type = 'softReboot' },
            @{ returnCode = 1641; type = 'hardReboot' }, @{ returnCode = 1618; type = 'retry' })
        detectionRules                 = @([ordered]@{
            '@odata.type'         = '#microsoft.graph.win32LobAppPowerShellScriptDetection'
            enforceSignatureCheck = $false
            runAs32Bit            = $false
            scriptContent         = $detection
        })
        largeIcon                      = @{ type = 'image/png'; value = $icon }
        isFeatured                     = $false
    }

    if ($Existing)
    {
        $appId = $Existing.id
        Set-Step "${label}: updating the existing Intune app..." $StartPercent
        Invoke-Graph -Auth $Auth -Method PATCH -Path "deviceAppManagement/mobileApps/$appId" -Body $body | Out-Null
        $action = 'Updated'
    }
    else
    {
        Set-Step "${label}: creating the Intune app..." $StartPercent
        $app = Invoke-Graph -Auth $Auth -Method POST -Path 'deviceAppManagement/mobileApps' -Body $body
        $appId = $app.id
        $action = 'Created'
    }

    Send-IntuneContent -Auth $Auth -AppId $appId -IntuneWinPath $Package.IntuneWin -StartPercent $StartPercent -EndPercent ($EndPercent - 2) -Label $label

    if (-not $Existing)
    {
        Set-Step "${label}: making it available in Company Portal..."
        $assign = @{ mobileAppAssignments = @(@{
            '@odata.type' = '#microsoft.graph.mobileAppAssignment'
            intent        = 'available'
            target        = @{ '@odata.type' = '#microsoft.graph.allLicensedUsersAssignmentTarget' }
            settings      = @{ '@odata.type' = '#microsoft.graph.win32LobAppAssignmentSettings'; notifications = 'showAll'; deliveryOptimizationPriority = 'notConfigured' }
        }) }
        Invoke-Graph -Auth $Auth -Method POST -Path "deviceAppManagement/mobileApps/$appId/assign" -Body $assign | Out-Null
    }
    Write-Log "$action Intune app '$($Package.DisplayName)' version $($Package.Version)." 'OK'
    return [pscustomobject]@{ AppId = $appId; Action = $action; Url = "https://intune.microsoft.com/#view/Microsoft_Intune_Apps/SettingsMenu/~/0/appId/$appId" }
}

function Publish-Printers
{
    # Bulk: one Intune app per printer. A failure on one printer doesn't stop the others.
    param ([hashtable]$Auth, [object[]]$Requests)
    $ErrorActionPreference = 'Stop'
    Set-Step 'Looking up existing printer apps in Intune...' 1
    $existing = Get-PrinterApps -Auth $Auth
    $results = New-Object System.Collections.Generic.List[object]
    $n = $Requests.Count
    for ($i = 0; $i -lt $n; $i++)
    {
        $req = $Requests[$i]
        $name = "$($req.TargetName)".Trim(); if (-not $name) { $name = $req.Printer.Name }
        $from = [int](100 * $i / $n); $to = [int](100 * ($i + 1) / $n)
        try
        {
            $app = $existing[$name]
            $req.AppVersion = Get-NextVersion -Requested $req.AppVersion -Existing $(if ($app) { $app.displayVersion })
            $req.CreateIntuneWin = $true
            Set-Step "Packaging $name ($($i + 1) of $n)..." $from
            $script:ProgressRange = @($from, [int]($from + ($to - $from) * 0.5))
            try { $pkg = New-PrinterPackage -Request $req } finally { $script:ProgressRange = $null }
            $desc = "Installs the printer $name ($($req.HostAddress)).`nDriver: $($req.Printer.DriverName)"
            if ($req.Printer.Location) { $desc += "`nLocation: $($req.Printer.Location)" }
            $pub = Publish-PrinterApp -Auth $Auth -Package $pkg -PrinterName $name -Existing $app -Description $desc -StartPercent ([int]($from + ($to - $from) * 0.5)) -EndPercent $to
            $results.Add([pscustomobject]@{ Printer = $name; Status = $pub.Action; Version = $pkg.Version; Url = $pub.Url; Logo = $pkg.Logo; Folder = $pkg.Folder; Error = '' })
        }
        catch
        {
            Write-Log "$name failed: $($_.Exception.Message)" 'ERROR'
            $results.Add([pscustomobject]@{ Printer = $name; Status = 'Failed'; Version = ''; Url = ''; Logo = $req.LogoPath; Folder = ''; Error = $_.Exception.Message })
        }
    }
    Set-Step 'Done.' 100
    return [pscustomobject]@{ __Result = $true; Results = $results.ToArray() }
}

function Save-Printers
{
    # Bulk "save files": one package folder per printer.
    param ([object[]]$Requests)
    $results = New-Object System.Collections.Generic.List[object]
    $n = $Requests.Count
    for ($i = 0; $i -lt $n; $i++)
    {
        $req = $Requests[$i]
        $name = "$($req.TargetName)".Trim(); if (-not $name) { $name = $req.Printer.Name }
        $script:ProgressRange = @([int](100 * $i / $n), [int](100 * ($i + 1) / $n))
        try
        {
            $pkg = New-PrinterPackage -Request $req
            $status = if ($pkg.IntuneWin) { 'Saved' } else { 'SavedNoIntuneWin' }
            $results.Add([pscustomobject]@{ Printer = $name; Status = $status; Version = $pkg.Version; Url = ''; Logo = $pkg.Logo; Folder = $pkg.Folder; Error = '' })
        }
        catch
        {
            Write-Log "$name failed: $($_.Exception.Message)" 'ERROR'
            $results.Add([pscustomobject]@{ Printer = $name; Status = 'Failed'; Version = ''; Url = ''; Logo = $req.LogoPath; Folder = ''; Error = $_.Exception.Message })
        }
        finally { $script:ProgressRange = $null }
    }
    return [pscustomobject]@{ __Result = $true; Results = $results.ToArray() }
}
#endregion

#region App update (GitHub releases)
function ConvertTo-ToolVersion
{
    param ([string]$Text)
    $v = $null
    if ([version]::TryParse(($Text -replace '^v', '' -replace '[-+].*$', ''), [ref]$v)) { return $v }
    return $null
}

function Get-LatestRelease
{
    # Newest non-pre-release on GitHub, compared with the running version.
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$($Ctx.UpdateRepo)/releases/latest" -Headers @{ 'User-Agent' = 'IntunePrinterPackager' } -TimeoutSec 15
    $asset = @($release.assets) | Where-Object { $_.name -eq 'PrinterPackager.exe' } | Select-Object -First 1
    $latest = ConvertTo-ToolVersion $release.tag_name
    $current = ConvertTo-ToolVersion $Ctx.ToolVersion
    # Builds with a suffix (2.2.0-ci.5, 2.2.0-beta.1) count as older than the same final version.
    $isPreview = "$($Ctx.ToolVersion)" -match '-'
    $newer = $latest -and $asset -and (-not $current -or $latest -gt $current -or ($latest -eq $current -and $isPreview))
    return [pscustomobject]@{
        __Result = $true
        Tag      = "$($release.tag_name)"
        Url      = if ($asset) { "$($asset.browser_download_url)" } else { '' }
        Newer    = [bool]$newer
        Current  = "$($Ctx.ToolVersion)"
    }
}

function Install-AppUpdate
{
    # Downloads the new exe and swaps it in next to the running one (a running exe can be renamed, not overwritten).
    # When that folder isn't writable (e.g. a read-only share) the new exe goes to Downloads instead.
    param ([string]$Url, [string]$Tag)
    $ErrorActionPreference = 'Stop'
    if (-not $Ctx.LauncherPath) { throw 'Updating only works when started from PrinterPackager.exe.' }
    $exe = $Ctx.LauncherPath
    $dir = Split-Path -Parent $exe
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("PrinterPackager-$Tag.exe")
    Set-Step "Downloading $Tag..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $wc = New-Object Net.WebClient
    $wc.Headers['User-Agent'] = 'IntunePrinterPackager'
    $wc.Proxy = [Net.WebRequest]::GetSystemWebProxy()
    $wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
    try { $wc.DownloadFile($Url, $tmp) } finally { $wc.Dispose() }
    if ((Get-Item -LiteralPath $tmp).Length -lt 1MB) { throw 'The download is incomplete.' }

    Set-Step 'Installing the update...'
    $old = "$exe.old"
    try
    {
        if (Test-Path -LiteralPath $old) { Remove-Item -LiteralPath $old -Force }
        Rename-Item -LiteralPath $exe -NewName (Split-Path -Leaf $old)
        try { Copy-Item -LiteralPath $tmp -Destination $exe -Force }
        catch { Rename-Item -LiteralPath $old -NewName (Split-Path -Leaf $exe); throw }
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        Write-Log "Updated to $Tag." 'OK'
        return [pscustomobject]@{ __Result = $true; Restart = $exe; SavedTo = '' }
    }
    catch
    {
        $downloads = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'
        New-Item -Path $downloads -ItemType Directory -Force | Out-Null
        $saved = Join-Path $downloads 'PrinterPackager.exe'
        Move-Item -LiteralPath $tmp -Destination $saved -Force
        Write-Log "Could not replace $exe ($($_.Exception.Message)). The new version was saved to $saved." 'WARN'
        return [pscustomobject]@{ __Result = $true; Restart = ''; SavedTo = $saved }
    }
}
#endregion

#region New printer (without an installed printer)
function Get-InfPrinterModels
{
    # Reads the printer model names from a driver .inf ([Manufacturer] -> model sections, %strings% resolved).
    # Only 64-bit (NTamd64) model sections count when the INF has them.
    param ([string]$InfPath)
    $sections = @{}
    $current = $null
    foreach ($raw in [IO.File]::ReadAllLines($InfPath))
    {
        $line = ($raw -replace '^\s+|\s+$', '')
        if (-not $line -or $line.StartsWith(';')) { continue }
        if ($line -match '^\[(.+)\]$') { $current = $Matches[1].Trim().ToLowerInvariant(); if (-not $sections.ContainsKey($current)) { $sections[$current] = New-Object System.Collections.Generic.List[string] }; continue }
        if ($current) { $sections[$current].Add($line) }
    }
    $strings = @{}
    foreach ($key in @($sections.Keys | Where-Object { $_ -eq 'strings' -or $_ -like 'strings.*' }))
    {
        foreach ($l in $sections[$key])
        {
            if ($l -match '^([^=]+?)\s*=\s*(.*)$' -and -not $strings.ContainsKey($Matches[1].Trim().ToLowerInvariant()))
            {
                $strings[$Matches[1].Trim().ToLowerInvariant()] = ($Matches[2] -replace '\s*;.*$', '').Trim().Trim('"')
            }
        }
    }
    $resolve = {
        param ([string]$Text)
        $t = ($Text -replace '\s*;.*$', '').Trim()
        if ($t -match '^%(.+)%$' -and $strings.ContainsKey($Matches[1].ToLowerInvariant())) { return $strings[$Matches[1].ToLowerInvariant()] }
        return $t.Trim('"')
    }
    $class = $sections['version'] | Where-Object { $_ -match '^Class\s*=' } | Select-Object -First 1
    $isPrinter = "$class" -match '=\s*"?Printer"?'

    $models = New-Object System.Collections.Generic.List[object]
    $is64 = $false
    foreach ($l in @($sections['manufacturer']))
    {
        if ($l -notmatch '^([^=]+?)\s*=\s*(.+)$') { continue }
        $maker = & $resolve $Matches[1]
        $parts = @($Matches[2].Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $base = $parts[0].ToLowerInvariant()
        $decorations = @($parts | Select-Object -Skip 1 | ForEach-Object { $_.ToLowerInvariant() })
        $amd64 = @($decorations | Where-Object { $_ -like 'ntamd64*' })
        $names = if ($amd64) { $is64 = $true; @($amd64 | ForEach-Object { "$base.$_" }) } elseif ($decorations) { @() } else { @($base) }
        foreach ($sectionName in $names)
        {
            foreach ($ml in @($sections[$sectionName]))
            {
                if ($ml -match '^([^=]+?)\s*=') { $models.Add([pscustomobject]@{ Manufacturer = $maker; Name = (& $resolve $Matches[1]) }) }
            }
        }
    }
    $unique = @($models | Where-Object { $_.Name } | Sort-Object Name -Unique)
    return [pscustomobject]@{ IsPrinterInf = $isPrinter; Has64Bit = $is64; Models = $unique }
}

function New-LibraryPrinter
{
    # Builds a library entry for a printer that is not installed anywhere: driver files + settings, no captured preferences.
    param ([hashtable]$Definition, [string]$LibraryDir)
    $ErrorActionPreference = 'Stop'
    $d = $Definition
    $name = "$($d.Name)".Trim()
    if (-not $name) { throw 'Enter a printer name.' }
    if ($name -match '[\\,!]') { throw 'The printer name cannot contain \ , or !.' }
    $hostAddress = "$($d.HostAddress)".Trim()
    if ($hostAddress -notmatch '^[A-Za-z0-9.\-:_\[\]]+$') { throw 'Enter a valid IP address or hostname.' }
    if (-not "$($d.DriverName)".Trim()) { throw 'Choose a driver.' }
    if (-not $LibraryDir) { throw 'Choose a library folder first.' }
    New-Item -Path $LibraryDir -ItemType Directory -Force | Out-Null

    $target = Join-Path $LibraryDir (ConvertTo-SafeFileName $name)
    if (Test-Path -LiteralPath $target) { throw "The library already has a printer called '$name'." }
    $tmp = "$target.saving"
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force }
    New-Item -Path $tmp -ItemType Directory -Force | Out-Null
    try
    {
        $driverFolder = ''
        $driverInf = ''
        $builtIn = $false
        $manufacturer = "$($d.Manufacturer)"
        if ($d.InfPath)
        {
            Set-Step 'Copying the driver folder...' 20
            $source = Split-Path -Parent $d.InfPath
            Copy-Item -Path (Join-Path $source '*') -Destination (New-Item -Path (Join-Path $tmp 'Driver') -ItemType Directory -Force).FullName -Recurse -Force
            if (-not (Get-ChildItem -LiteralPath (Join-Path $tmp 'Driver') -Filter '*.cat' -Recurse | Select-Object -First 1))
            {
                Write-Log 'This driver folder has no .cat file (not signed). Windows will probably refuse to install it.' 'WARN'
            }
            $driverFolder = 'Driver'
            $driverInf = Split-Path -Leaf $d.InfPath
        }
        else
        {
            Set-Step "Collecting driver files for '$($d.DriverName)'..." 20
            $drv = Get-PrinterDriver -Name $d.DriverName -ErrorAction Stop | Sort-Object { $_.PrinterEnvironment -ne 'Windows x64' } | Select-Object -First 1
            if (-not $manufacturer) { $manufacturer = "$($drv.Manufacturer)" }
            $inf = Export-PrinterDriverFiles -DriverName $d.DriverName -InfPath $drv.InfPath -DestinationDir (Join-Path $tmp 'Driver')
            if ($inf -eq 'INBOX') { $builtIn = $true; Remove-Item -LiteralPath (Join-Path $tmp 'Driver') -Recurse -Force -ErrorAction SilentlyContinue }
            elseif ($inf) { $driverFolder = 'Driver'; $driverInf = $inf }
            else { throw "The driver files for '$($d.DriverName)' could not be found on this PC." }
        }

        $capture = [ordered]@{
            CapturedOn     = (Get-Date).ToString('s')
            CapturedFrom   = "$env:COMPUTERNAME (new printer)"
            DriverFolder   = $driverFolder
            DriverInf      = $driverInf
            BuiltInDriver  = $builtIn
            UsePrintBrm    = $false
            UseSettingsDat = $false
            DuplexingMode  = ''
            PaperSize      = ''
            Color          = $null
        }
        Write-TextFile -Path (Join-Path $tmp 'capture.json') -Content ($capture | ConvertTo-Json -Depth 4)

        $brand = Get-PrinterBrand -Text @($manufacturer, $d.DriverName, $name)
        $logo = if ($d.LogoPath) { $d.LogoPath } else { Join-Path $Ctx.LogoDir $brand.File }
        Convert-LogoToIcon -Source $logo -Destination (Join-Path $tmp 'Logo.png')
        $protocol = if ($d.Protocol) { [int]$d.Protocol } else { 1 }
        $portNumber = if ($d.PortNumber) { [int]$d.PortNumber } else { 9100 }
        $portName = if ($protocol -eq 2) { "IP_${hostAddress}_LPR" } elseif ($portNumber -ne 9100) { "IP_${hostAddress}_$portNumber" } else { "IP_$hostAddress" }
        $advanced = [ordered]@{}
        foreach ($k in 'Duplex', 'Color', 'PaperSize') { if ($d[$k]) { $advanced[$k] = $d[$k] } }
        $entry = [ordered]@{
            SchemaVersion = 1
            SavedOn       = (Get-Date).ToString('s')
            SavedBy       = "$env:USERDOMAIN\$env:USERNAME on $env:COMPUTERNAME"
            Printer       = [ordered]@{
                Name          = $name
                DriverName    = "$($d.DriverName)".Trim()
                Manufacturer  = $manufacturer
                PortName      = $portName
                PortKind      = 'TCP/IP'
                HostAddress   = $hostAddress
                Protocol      = $protocol
                PortNumber    = $portNumber
                LprQueueName  = if ($protocol -eq 2) { if ($d.LprQueueName) { "$($d.LprQueueName)" } else { 'lp' } } else { '' }
                SnmpEnabled   = [bool]$d.SnmpEnabled
                SnmpCommunity = if ($d.SnmpCommunity) { "$($d.SnmpCommunity)" } else { 'public' }
                SnmpIndex     = 1
                Location      = "$($d.Location)"
                Comment       = "$($d.Comment)"
                Brand         = $brand.Name
            }
            Settings      = [ordered]@{
                TargetName  = $name
                HostAddress = $hostAddress
                AppName     = "Printer - $name"
                Advanced    = $advanced
            }
        }
        Write-TextFile -Path (Join-Path $tmp 'printer.json') -Content ($entry | ConvertTo-Json -Depth 6)
        Move-Item -LiteralPath $tmp -Destination $target
        Set-Step 'Done.' 100
        Write-Log "New printer '$name' saved to the library." 'OK'
        return [pscustomobject]@{ __Result = $true; Name = $name; Folder = $target }
    }
    catch
    {
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
        throw
    }
}
#endregion

#region Client projects
# The library holds one folder per client site; each client folder holds client.json and one folder per printer.
function Export-ClientProject
{
    param ([string]$ClientDir, [string]$ZipPath)
    $ErrorActionPreference = 'Stop'
    Set-Step "Exporting $(Split-Path -Leaf $ClientDir)..."
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $ZipPath) { Remove-Item -LiteralPath $ZipPath -Force }
    [IO.Compression.ZipFile]::CreateFromDirectory($ClientDir, $ZipPath, [IO.Compression.CompressionLevel]::Optimal, $true)
    Write-Log ("Exported client to {0} ({1:N1} MB)." -f $ZipPath, ((Get-Item -LiteralPath $ZipPath).Length / 1MB)) 'OK'
    return [pscustomobject]@{ __Result = $true; Zip = $ZipPath }
}

function Import-ClientProject
{
    param ([string]$ZipPath, [string]$LibraryRoot, [bool]$Overwrite)
    $ErrorActionPreference = 'Stop'
    Set-Step 'Importing client...'
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try
    {
        # Windows PowerShell writes zip entries with backslashes, newer .NET with slashes.
        $names = @($zip.Entries | ForEach-Object { $_.FullName -replace '\\', '/' })
        $first = @($names | ForEach-Object { $_.Split('/')[0] } | Select-Object -Unique)
        if ($first.Count -ne 1 -or -not ($names -match '^[^/]+/client\.json$')) { throw 'This zip is not a Printer Packager client export.' }
        $name = $first[0]
    }
    finally { $zip.Dispose() }
    $target = Join-Path $LibraryRoot $name
    if (Test-Path -LiteralPath $target)
    {
        if (-not $Overwrite) { return [pscustomobject]@{ __Result = $true; Exists = $true; Name = $name } }
        Remove-Item -LiteralPath $target -Recurse -Force
    }
    New-Item -Path $LibraryRoot -ItemType Directory -Force | Out-Null
    [IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $LibraryRoot)
    Write-Log "Imported client '$name'." 'OK'
    return [pscustomobject]@{ __Result = $true; Exists = $false; Name = $name }
}
#endregion

}
. $WorkerCode
$WorkerCodeText = $WorkerCode.ToString()
#endregion

#region Headless mode
if ($IsHeadless)
{
    $ErrorActionPreference = 'Stop'
    if ($FromLibrary)
    {
        $printer = (Get-PrinterLibrary -LibraryDir $LibraryDir).Printers | Where-Object { $_.Name -eq $PrinterName } | Select-Object -First 1
        if (-not $printer) { throw "Printer '$PrinterName' is not in the library $LibraryDir." }
    }
    else
    {
        $inventory = Get-PrinterInventory
        $printer = $inventory.Printers | Where-Object { $_.Name -eq $PrinterName } | Select-Object -First 1
        if (-not $printer) { throw "Printer '$PrinterName' was not found. Available: $(($inventory.Printers | ForEach-Object Name) -join ', ')" }
    }
    if (-not $OutputDir) { $OutputDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Printer Packages' }
    $request = @{
        Printer            = $printer
        TargetName         = $printer.Name
        HostAddress        = if ($PrinterHostAddress) { $PrinterHostAddress } else { $printer.Address }
        AppVendor          = if ($AppVendor) { $AppVendor } else { 'Company' }
        AppName            = "Printer - $($printer.Name)"
        AppVersion         = $AppVersion
        LogoPath           = $printer.LogoPath
        IncludeDriver      = -not $NoDriver
        IncludePreferences = -not $NoPreferences
        SetDefault         = [bool]$SetDefault
        CreateIntuneWin    = -not $NoIntuneWin
        CreateZip          = [bool]$Zip
        OutputDir          = $OutputDir
        CaptureDir         = $printer.CaptureDir
        Advanced           = @{}
    }
    if ($printer.SavedSettings -and $printer.SavedSettings.Advanced)
    {
        foreach ($prop in $printer.SavedSettings.Advanced.PSObject.Properties) { $request.Advanced[$prop.Name] = $prop.Value }
    }
    if ($SaveToLibrary)
    {
        $result = Save-PrinterToLibrary -Requests @($request) -LibraryDir $LibraryDir
        $failed = @($result.Results | Where-Object Status -eq 'Failed')
        if ($failed) { throw $failed[0].Error }
    }
    else { $result = New-PrinterPackage -Request $request }
    $result | Select-Object -Property * -ExcludeProperty __Result | Format-List | Out-String | Write-Host
    return
}
#endregion

#region Window
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Printer Packager" Width="1240" Height="800" MinWidth="1040" MinHeight="660"
        WindowStartupLocation="CenterScreen" Background="#F3F3F3" Foreground="#1B1B1B"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13" UseLayoutRounding="True" SnapsToDevicePixels="True">
    <Window.Resources>
        <SolidColorBrush x:Key="Accent" Color="#C80032"/>
        <SolidColorBrush x:Key="Muted" Color="#616161"/>
        <SolidColorBrush x:Key="Line" Color="#E5E5E5"/>

        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="18"/>
        </Style>
        <Style x:Key="SectionTitle" TargetType="TextBlock">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
        </Style>
        <Style x:Key="Label" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource Muted}"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,0,0,4"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#D1D1D1"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border x:Name="Bd" Background="White" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsKeyboardFocused" Value="True">
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#C80032"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Background" Value="#F7F7F7"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#1B1B1B"/>
            <Setter Property="BorderBrush" Value="#D1D1D1"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#F5F5F5"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Background" Value="#EBEBEB"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Foreground" Value="#A0A0A0"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="AccentButton" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="18,9"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="#C80032" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Bd" Property="Background" Value="#A8002A"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="Bd" Property="Background" Value="#8C0023"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Bd" Property="Background" Value="#E6A3B4"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="LinkButton" TargetType="Button">
            <Setter Property="Foreground" Value="#C80032"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <TextBlock x:Name="T" Text="{Binding Content, RelativeSource={RelativeSource TemplatedParent}}" Foreground="{TemplateBinding Foreground}"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="T" Property="TextDecorations" Value="Underline"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="T" Property="Foreground" Value="#A0A0A0"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SegmentButton" TargetType="RadioButton">
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="Bd" Background="White" BorderBrush="#D1D1D1" BorderThickness="1" CornerRadius="14" Padding="14,4">
                            <ContentPresenter x:Name="Cp" HorizontalAlignment="Center" VerticalAlignment="Center" TextElement.Foreground="#1B1B1B"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="#C80032"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#C80032"/>
                                <Setter TargetName="Cp" Property="TextElement.Foreground" Value="White"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="BorderBrush" Value="#C80032"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="{StaticResource Muted}"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
            <Setter Property="BorderThickness" Value="0,0,0,1"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="White"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#F7F7F7"/></Trigger>
                <Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="#FBE5EB"/></Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FocusVisualStyle" Value="{x:Null}"/>
            <Setter Property="Foreground" Value="#1B1B1B"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="DataGridCell">
                        <Border Background="Transparent" Padding="10,0">
                            <ContentPresenter VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Background="#F3F3F3">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="White" BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1" Padding="24,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Image Name="ImgBrand" Width="36" Height="36" VerticalAlignment="Center" RenderOptions.BitmapScalingMode="HighQuality"/>
                <Border Grid.Column="1" BorderBrush="{StaticResource Line}" BorderThickness="1,0,0,0" Margin="16,2,0,2" Padding="16,0,0,0">
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Printer Packager" FontSize="18" FontWeight="SemiBold"/>
                        <TextBlock Text="Turn an installed printer into an Intune app (PSAppDeployToolkit v4)" Foreground="{StaticResource Muted}"/>
                    </StackPanel>
                </Border>
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <StackPanel Name="PanelAccount" VerticalAlignment="Center" Margin="0,0,12,0" Visibility="Collapsed">
                        <TextBlock Name="TxtAccountUser" FontWeight="SemiBold" HorizontalAlignment="Right"/>
                        <TextBlock Name="TxtAccountTenant" Foreground="{StaticResource Muted}" FontSize="12" HorizontalAlignment="Right"/>
                    </StackPanel>
                    <Button Name="BtnUpdate" Margin="0,0,10,0" ToolTip="Check for a new version">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE896;" FontFamily="Segoe MDL2 Assets" Margin="0,1,8,0" VerticalAlignment="Center"/>
                            <TextBlock Name="TxtUpdate" Text="Update"/>
                        </StackPanel>
                    </Button>
                    <Button Name="BtnSignIn" Margin="0,0,10,0">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE77B;" FontFamily="Segoe MDL2 Assets" Margin="0,1,8,0" VerticalAlignment="Center"/>
                            <TextBlock Name="TxtSignIn" Text="Sign in to Intune"/>
                        </StackPanel>
                    </Button>
                    <Button Name="BtnRefresh">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE72C;" FontFamily="Segoe MDL2 Assets" Margin="0,1,8,0" VerticalAlignment="Center"/>
                            <TextBlock Text="Refresh"/>
                        </StackPanel>
                    </Button>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Body -->
        <Grid Grid.Row="1" Margin="24,20,24,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="400"/>
            </Grid.ColumnDefinitions>

            <!-- Printer list -->
            <Border Style="{StaticResource Card}" Padding="0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid Margin="18,16,18,12">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="240"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel VerticalAlignment="Center">
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                <RadioButton Name="RbSourcePc" Content="This PC" GroupName="Source" IsChecked="True" Style="{StaticResource SegmentButton}"/>
                                <RadioButton Name="RbSourceLibrary" Content="Library" GroupName="Source" Style="{StaticResource SegmentButton}" Margin="6,0,0,0"/>
                                <Button Name="BtnNewPrinter" Content="+ New printer" Style="{StaticResource LinkButton}" Margin="14,0,0,0" VerticalAlignment="Center"/>
                            </StackPanel>
                            <TextBlock Text="1. Choose a printer" FontSize="14" FontWeight="SemiBold"/>
                            <TextBlock Name="TxtCount" Text="Loading..." Foreground="{StaticResource Muted}" FontSize="12"/>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Tick one or more printers." Foreground="{StaticResource Muted}" FontSize="12"/>
                                <Button Name="BtnSelectAll" Content="Select all" Style="{StaticResource LinkButton}" FontSize="12" Margin="8,0,0,0"/>
                            </StackPanel>
                        </StackPanel>
                        <Grid Grid.Column="1" VerticalAlignment="Top">
                            <TextBox Name="TxtFilter" Padding="30,6,8,6"/>
                            <TextBlock Text="&#xE721;" FontFamily="Segoe MDL2 Assets" Foreground="{StaticResource Muted}" Margin="10,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                            <TextBlock Name="TxtFilterHint" Text="Search" Foreground="#9E9E9E" Margin="31,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                        </Grid>
                    <StackPanel Grid.Row="1" Grid.ColumnSpan="2" Margin="0,10,0,0">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Client" FontWeight="SemiBold" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <ComboBox Name="CmbClient" Width="260" DisplayMemberPath="Display"/>
                            <Button Name="BtnNewClient" Content="New" Style="{StaticResource LinkButton}" Margin="12,0,0,0" VerticalAlignment="Center"/>
                            <Button Name="BtnExportClient" Content="Export" Style="{StaticResource LinkButton}" Margin="12,0,0,0" VerticalAlignment="Center"/>
                            <Button Name="BtnImportClient" Content="Import" Style="{StaticResource LinkButton}" Margin="12,0,0,0" VerticalAlignment="Center"/>
                        </StackPanel>
                        <StackPanel Name="PanelLibraryPath" Orientation="Horizontal" Visibility="Collapsed" Margin="0,6,0,0">
                            <TextBlock Text="Library folder:" Foreground="{StaticResource Muted}" FontSize="12"/>
                            <TextBlock Name="TxtLibraryPath" FontSize="12" Margin="4,0,0,0" MaxWidth="360" TextTrimming="CharacterEllipsis"/>
                            <Button Name="BtnLibraryBrowse" Content="Change..." Style="{StaticResource LinkButton}" FontSize="12" Margin="8,0,0,0"/>
                        </StackPanel>
                    </StackPanel>
                    </Grid>
                    <DataGrid Name="GridPrinters" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended"
                              HeadersVisibility="Column" GridLinesVisibility="None" BorderThickness="0" Background="White"
                              RowHeight="48" CanUserResizeRows="False" CanUserReorderColumns="False" HorizontalScrollBarVisibility="Disabled" Margin="8,0,8,8">
                        <DataGrid.Columns>
                            <DataGridTemplateColumn Width="40" CanUserSort="False">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <CheckBox Margin="0" Focusable="False" VerticalAlignment="Center" IsChecked="{Binding IsSelected, Mode=TwoWay, RelativeSource={RelativeSource AncestorType=DataGridRow}}"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Width="48" CanUserSort="False">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <Image Source="{Binding LogoImage}" Width="30" Height="30" RenderOptions.BitmapScalingMode="HighQuality"/>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Printer" Width="2.2*" SortMemberPath="Name">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <StackPanel>
                                            <TextBlock Text="{Binding Name}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                                            <TextBlock Text="{Binding DriverName}" Foreground="#616161" FontSize="12" TextTrimming="CharacterEllipsis"/>
                                        </StackPanel>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTextColumn Header="Address" Binding="{Binding Address}" Width="1.2*"/>
                            <DataGridTextColumn Header="Port" Binding="{Binding PortKind}" Width="70"/>
                            <DataGridTextColumn Header="Location" Binding="{Binding Location}" Width="1*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <TextBlock Name="TxtEmpty" Grid.Row="1" Text="No printers found on this computer." Foreground="{StaticResource Muted}" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/>
                </Grid>
            </Border>

            <!-- Settings -->
            <Border Grid.Column="2" Style="{StaticResource Card}" Padding="0">
                <Grid>
                    <ScrollViewer Name="PanelSettings" VerticalScrollBarVisibility="Auto" Visibility="Collapsed">
                        <StackPanel Margin="18,16,18,18">
                            <TextBlock Name="TxtPanelTitle" Text="2. Check the package" Style="{StaticResource SectionTitle}"/>
                            <StackPanel Name="PanelMulti" Visibility="Collapsed">
                                <ItemsControl Name="ListMulti">
                                    <ItemsControl.ItemTemplate>
                                        <DataTemplate>
                                            <Grid Margin="0,0,0,8">
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="Auto"/>
                                                    <ColumnDefinition Width="*"/>
                                                </Grid.ColumnDefinitions>
                                                <Image Source="{Binding Logo}" Width="30" Height="30" RenderOptions.BitmapScalingMode="HighQuality" VerticalAlignment="Top"/>
                                                <StackPanel Grid.Column="1" Margin="10,0,0,0">
                                                    <TextBlock Text="{Binding Name}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                                                    <TextBlock Text="{Binding Detail}" Foreground="{Binding DetailColor}" FontSize="12" TextTrimming="CharacterEllipsis"/>
                                                </StackPanel>
                                            </Grid>
                                        </DataTemplate>
                                    </ItemsControl.ItemTemplate>
                                </ItemsControl>
                                <TextBlock Text="Click a single printer to change its name, IP or logo." Foreground="{StaticResource Muted}" FontSize="12" TextWrapping="Wrap" Margin="0,2,0,0"/>
                            </StackPanel>
                            <StackPanel Name="PanelSingle">
                            <Grid Margin="0,0,0,14">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Border BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="10" Width="76" Height="76" VerticalAlignment="Top">
                                    <Image Name="ImgLogo" Margin="4" RenderOptions.BitmapScalingMode="HighQuality"/>
                                </Border>
                                <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
                                    <TextBlock Name="TxtSelName" FontSize="15" FontWeight="SemiBold" TextWrapping="Wrap"/>
                                    <TextBlock Name="TxtSelDriver" Foreground="{StaticResource Muted}" FontSize="12" TextWrapping="Wrap" Margin="0,2,0,6"/>
                                    <StackPanel Orientation="Horizontal">
                                        <ComboBox Name="CmbLogo" Width="150" DisplayMemberPath="Name"/>
                                        <Button Name="BtnLogoBrowse" Content="Own logo..." Style="{StaticResource LinkButton}" Margin="10,0,0,0" VerticalAlignment="Center"/>
                                    </StackPanel>
                                </StackPanel>
                            </Grid>

                            <TextBlock Text="Printer name on devices" Style="{StaticResource Label}"/>
                            <TextBox Name="TxtTargetName" Margin="0,0,0,10"/>
                            <TextBlock Text="IP address or hostname" Style="{StaticResource Label}"/>
                            <TextBox Name="TxtHost"/>
                            <TextBlock Name="TxtHostHint" Foreground="#B54700" FontSize="12" TextWrapping="Wrap" Margin="0,4,0,0" Visibility="Collapsed"/>
                            <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                                <Button Name="BtnAdvanced" Content="Advanced settings..." Style="{StaticResource LinkButton}"/>
                                <TextBlock Name="TxtAdvancedSummary" Foreground="{StaticResource Muted}" FontSize="12" Margin="10,0,0,0" TextTrimming="CharacterEllipsis" MaxWidth="220"/>
                            </StackPanel>
                            </StackPanel>

                            <Border BorderBrush="{StaticResource Line}" BorderThickness="0,1,0,0" Margin="0,16,0,14"/>
                            <TextBlock Text="Intune app" Style="{StaticResource SectionTitle}"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="12"/>
                                    <ColumnDefinition Width="100"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel>
                                    <TextBlock Text="Publisher" Style="{StaticResource Label}"/>
                                    <TextBox Name="TxtVendor"/>
                                </StackPanel>
                                <StackPanel Grid.Column="2">
                                    <TextBlock Text="Version" Style="{StaticResource Label}"/>
                                    <TextBox Name="TxtVersion" Text="1.0.0"/>
                                </StackPanel>
                            </Grid>
                            <StackPanel Name="PanelAppName">
                                <TextBlock Text="App name" Style="{StaticResource Label}" Margin="0,10,0,4"/>
                                <TextBox Name="TxtAppName"/>
                            </StackPanel>

                            <Border BorderBrush="{StaticResource Line}" BorderThickness="0,1,0,0" Margin="0,16,0,14"/>
                            <TextBlock Text="Options" Style="{StaticResource SectionTitle}"/>
                            <CheckBox Name="ChkDriver" Content="Include the driver" IsChecked="True"/>
                            <CheckBox Name="ChkPrefs" Content="Copy printer preferences (trays, duplex, colour)" IsChecked="True"/>
                            <CheckBox Name="ChkDefault" Content="Make it the default printer"/>
                            <CheckBox Name="ChkZip" Content="Also make a .zip for manual installs" Margin="0"/>
                        </StackPanel>
                    </ScrollViewer>
                    <StackPanel Name="PanelNoSelection" VerticalAlignment="Center" HorizontalAlignment="Center" Margin="30">
                        <TextBlock Text="&#xE749;" FontFamily="Segoe MDL2 Assets" FontSize="40" Foreground="#C8C8C8" HorizontalAlignment="Center"/>
                        <TextBlock Text="Tick a printer on the left" FontSize="14" FontWeight="SemiBold" HorizontalAlignment="Center" Margin="0,12,0,4"/>
                        <TextBlock Text="Its driver, port, preferences and logo are packaged for you." Foreground="{StaticResource Muted}" TextWrapping="Wrap" TextAlignment="Center"/>
                    </StackPanel>
                </Grid>
            </Border>
        </Grid>

        <!-- Action bar -->
        <Border Grid.Row="2" Style="{StaticResource Card}" Margin="24,20,24,0" Padding="18,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Save to" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,12,0"/>
                <Grid Grid.Column="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBox Name="TxtOutput"/>
                    <Button Name="BtnBrowseOutput" Grid.Column="1" Content="Browse..." Margin="8,0,0,0"/>
                    <Button Name="BtnOpenOutput" Grid.Column="2" Content="Open" Margin="8,0,0,0"/>
                </Grid>
                <Button Name="BtnSaveLibrary" Grid.Column="2" Content="Save to library" Margin="24,0,10,0" Padding="16,9"/>
                <Button Name="BtnPackageLocal" Grid.Column="3" Content="Save files only" Margin="0,0,10,0" Padding="16,9"/>
                <Button Name="BtnPackage" Grid.Column="4" Style="{StaticResource AccentButton}" Content="Publish to Company Portal"/>
            </Grid>
        </Border>

        <!-- Log -->
        <Expander Name="ExpLog" Grid.Row="3" Margin="24,12,24,16">
            <Expander.Header>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Activity log" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtStatus" Text="" Foreground="{StaticResource Muted}" Margin="12,0,0,0"/>
                </StackPanel>
            </Expander.Header>
            <TextBox Name="TxtLog" Height="150" Margin="0,8,0,0" IsReadOnly="True" FontFamily="Cascadia Mono, Consolas" FontSize="12"
                     VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" TextWrapping="NoWrap" VerticalContentAlignment="Top"/>
        </Expander>

        <!-- Busy overlay -->
        <Grid Name="Overlay" Grid.RowSpan="4" Background="#B3F3F3F3" Visibility="Collapsed">
            <Border Name="BusyCard" Style="{StaticResource Card}" Width="440" HorizontalAlignment="Center" VerticalAlignment="Center" Padding="28,26">
                <Border.Effect><DropShadowEffect BlurRadius="24" ShadowDepth="4" Opacity="0.15"/></Border.Effect>
                <StackPanel>
                    <Grid Width="44" Height="44" Margin="0,0,0,16" RenderTransformOrigin="0.5,0.5">
                        <Grid.RenderTransform><RotateTransform Angle="0"/></Grid.RenderTransform>
                        <Ellipse Stroke="#F4D3DC" StrokeThickness="4"/>
                        <Path Stroke="#C80032" StrokeThickness="4" StrokeStartLineCap="Round" StrokeEndLineCap="Round" Data="M 22,2 A 20,20 0 0 1 42,22"/>
                        <Grid.Triggers>
                            <EventTrigger RoutedEvent="FrameworkElement.Loaded">
                                <BeginStoryboard>
                                    <Storyboard RepeatBehavior="Forever">
                                        <DoubleAnimation Storyboard.TargetProperty="(UIElement.RenderTransform).(RotateTransform.Angle)" From="0" To="360" Duration="0:0:0.9"/>
                                    </Storyboard>
                                </BeginStoryboard>
                            </EventTrigger>
                        </Grid.Triggers>
                    </Grid>
                    <TextBlock Name="TxtBusyTitle" FontSize="16" FontWeight="SemiBold" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap"/>
                    <TextBlock Name="TxtBusyStep" Foreground="{StaticResource Muted}" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Margin="0,6,0,16"/>
                    <ProgressBar Name="PrgBusy" Height="4" Foreground="#C80032" Background="#F4D3DC" BorderThickness="0" IsIndeterminate="True"/>
                </StackPanel>
            </Border>

            <!-- Result list card (bulk / publish) -->
            <Border Name="ResultListCard" Style="{StaticResource Card}" Width="640" HorizontalAlignment="Center" VerticalAlignment="Center" Padding="28,24" Visibility="Collapsed">
                <Border.Effect><DropShadowEffect BlurRadius="24" ShadowDepth="4" Opacity="0.15"/></Border.Effect>
                <StackPanel>
                    <TextBlock Name="TxtResultListTitle" FontSize="18" FontWeight="SemiBold"/>
                    <TextBlock Name="TxtResultListSub" Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,4,0,16"/>
                    <ScrollViewer MaxHeight="380" VerticalScrollBarVisibility="Auto">
                        <ItemsControl Name="ListResults">
                            <ItemsControl.ItemTemplate>
                                <DataTemplate>
                                    <Border BorderBrush="#EDEDED" BorderThickness="0,0,0,1" Padding="0,8">
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="Auto"/>
                                                <ColumnDefinition Width="*"/>
                                                <ColumnDefinition Width="Auto"/>
                                            </Grid.ColumnDefinitions>
                                            <Image Source="{Binding Logo}" Width="34" Height="34" RenderOptions.BitmapScalingMode="HighQuality" VerticalAlignment="Top"/>
                                            <StackPanel Grid.Column="1" Margin="12,0,12,0">
                                                <TextBlock Text="{Binding Printer}" FontWeight="SemiBold" TextTrimming="CharacterEllipsis"/>
                                                <TextBlock Text="{Binding Detail}" Foreground="{Binding DetailColor}" FontSize="12" TextWrapping="Wrap"/>
                                            </StackPanel>
                                            <Button Grid.Column="2" Content="{Binding LinkText}" Tag="{Binding Url}" Visibility="{Binding LinkVisibility}" Style="{StaticResource LinkButton}" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                </DataTemplate>
                            </ItemsControl.ItemTemplate>
                        </ItemsControl>
                    </ScrollViewer>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,18,0,0">
                        <Button Name="BtnResultListOpen" Content="Open folder" Margin="0,0,10,0"/>
                        <Button Name="BtnResultListClose" Style="{StaticResource AccentButton}" Content="Done"/>
                    </StackPanel>
                </StackPanel>
            </Border>

            <!-- Result card -->
            <Border Name="ResultCard" Style="{StaticResource Card}" Width="620" HorizontalAlignment="Center" VerticalAlignment="Center" Padding="28,24" Visibility="Collapsed">
                <Border.Effect><DropShadowEffect BlurRadius="24" ShadowDepth="4" Opacity="0.15"/></Border.Effect>
                <StackPanel>
                    <Grid Margin="0,0,0,18">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Border BorderBrush="{StaticResource Line}" BorderThickness="1" CornerRadius="10" Width="64" Height="64">
                            <Image Name="ImgResultLogo" Margin="3" RenderOptions.BitmapScalingMode="HighQuality"/>
                        </Border>
                        <StackPanel Grid.Column="1" Margin="16,0,0,0" VerticalAlignment="Center">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE73E;" FontFamily="Segoe MDL2 Assets" Foreground="#107C10" FontSize="16" Margin="0,2,8,0"/>
                                <TextBlock Text="Package ready" FontSize="18" FontWeight="SemiBold"/>
                            </StackPanel>
                            <TextBlock Name="TxtResultName" Foreground="{StaticResource Muted}" TextWrapping="Wrap" Margin="0,2,0,0"/>
                        </StackPanel>
                    </Grid>
                    <TextBlock Text="Upload to Intune (Apps > Windows > Add > Windows app (Win32))" Style="{StaticResource Label}"/>
                    <TextBox Name="TxtResultFile" IsReadOnly="True" Margin="0,0,0,10"/>
                    <TextBlock Text="Install command" Style="{StaticResource Label}"/>
                    <Grid Margin="0,0,0,10">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                        <TextBox Name="TxtResultInstall" IsReadOnly="True" FontFamily="Cascadia Mono, Consolas" FontSize="12"/>
                        <Button Name="BtnCopyInstall" Grid.Column="1" Content="Copy" Margin="8,0,0,0"/>
                    </Grid>
                    <TextBlock Text="Uninstall command" Style="{StaticResource Label}"/>
                    <Grid Margin="0,0,0,10">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                        <TextBox Name="TxtResultUninstall" IsReadOnly="True" FontFamily="Cascadia Mono, Consolas" FontSize="12"/>
                        <Button Name="BtnCopyUninstall" Grid.Column="1" Content="Copy" Margin="8,0,0,0"/>
                    </Grid>
                    <TextBlock Name="TxtResultNotes" TextWrapping="Wrap" Foreground="{StaticResource Muted}" Margin="0,4,0,20"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button Name="BtnResultOpen" Content="Open folder" Margin="0,0,10,0"/>
                        <Button Name="BtnResultClose" Style="{StaticResource AccentButton}" Content="Done"/>
                    </StackPanel>
                </StackPanel>
            </Border>
        </Grid>
    </Grid>
</Window>
'@

$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
$ui = @{}
foreach ($node in $xaml.SelectNodes('//*[@Name]')) { $ui[$node.Name] = $window.FindName($node.Name) }

# Raw .NET values for data binding (PowerShell's PSObject wrapper breaks image/brush bindings).
function ConvertTo-Brush
{
    param ([string]$Color)
    $brush = [Windows.Media.Brush](New-Object Windows.Media.BrushConverter).ConvertFromString($Color)
    $brush.Freeze()
    return , $brush
}

function New-Bitmap
{
    param ([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $null }
    $bmp = New-Object Windows.Media.Imaging.BitmapImage
    $bmp.BeginInit()
    $bmp.CacheOption = 'OnLoad'
    $bmp.UriSource = New-Object Uri $Path
    $bmp.EndInit()
    $bmp.Freeze()
    return $bmp
}

$ui.ImgBrand.Source = New-Bitmap (Join-Path $Ctx.BrandingDir 'connectify-icon.png')
$window.Icon = New-Bitmap (Join-Path $Ctx.BrandingDir 'connectify-icon.png')
if ($ToolVersion) { $window.Title = "Printer Packager $ToolVersion" }
#endregion

#region Settings
$SettingsFile = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'IntunePrinterPackager\settings.json'
$Settings = [pscustomobject]@{
    Vendor             = 'Company'
    OutputDir          = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Printer Packages'
    IncludeDriver      = $true
    IncludePreferences = $true
    CreateZip          = $false
    LibraryDir         = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Printer Library'
    ClientId           = ''
    Client             = ''
}
try
{
    if (Test-Path -LiteralPath $SettingsFile)
    {
        $saved = [IO.File]::ReadAllText($SettingsFile) | ConvertFrom-Json
        foreach ($prop in $saved.PSObject.Properties) { if ($Settings.PSObject.Properties[$prop.Name]) { $Settings.($prop.Name) = $prop.Value } }
    }
}
catch { }
if ($OutputDir) { $Settings.OutputDir = $OutputDir }
if ($AppVendor) { $Settings.Vendor = $AppVendor }
if ($LibraryDir) { $Settings.LibraryDir = $LibraryDir }
if ($Settings.ClientId) { $Ctx.ClientId = "$($Settings.ClientId)" }

$ui.TxtVendor.Text = $Settings.Vendor
$ui.TxtOutput.Text = $Settings.OutputDir
$ui.ChkDriver.IsChecked = [bool]$Settings.IncludeDriver
$ui.ChkPrefs.IsChecked = [bool]$Settings.IncludePreferences
$ui.ChkZip.IsChecked = [bool]$Settings.CreateZip
$ui.TxtLibraryPath.Text = $Settings.LibraryDir

function Save-Settings
{
    try
    {
        $Settings.Vendor = $ui.TxtVendor.Text.Trim()
        $Settings.OutputDir = $ui.TxtOutput.Text.Trim()
        $Settings.IncludeDriver = [bool]$ui.ChkDriver.IsChecked
        $Settings.IncludePreferences = [bool]$ui.ChkPrefs.IsChecked
        $Settings.CreateZip = [bool]$ui.ChkZip.IsChecked
        $Settings.LibraryDir = $ui.TxtLibraryPath.Text.Trim()
        New-Item -Path (Split-Path -Parent $SettingsFile) -ItemType Directory -Force | Out-Null
        [IO.File]::WriteAllText($SettingsFile, ($Settings | ConvertTo-Json))
    }
    catch { }
}
#endregion

#region Background jobs
$Sync = [hashtable]::Synchronized(@{
    Log     = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    Step       = ''
    Percent    = -1
    DeviceCode = $null
})
$Job = $null
$WorkerBootstrap = @'
param ($Code, $Command, $Arguments, $Sync, $Ctx)
$ErrorActionPreference = 'Stop'
. ([scriptblock]::Create($Code))
& $Command @Arguments
'@

function Show-Error
{
    param ($ErrorObject, [string]$Title = 'Something went wrong')
    $msg = if ($ErrorObject -is [Management.Automation.ErrorRecord]) { $ErrorObject.Exception.Message } elseif ($ErrorObject -is [Exception]) { $ErrorObject.Message } else { "$ErrorObject" }
    Write-Log $msg 'ERROR'
    [void][Windows.MessageBox]::Show($window, $msg, $Title, 'OK', 'Error')
}

function Start-Work
{
    param ([string]$Title, [string]$Command, [hashtable]$Arguments = @{}, [scriptblock]$OnDone)
    if ($script:Job) { return }
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($WorkerBootstrap).AddArgument($WorkerCodeText).AddArgument($Command).AddArgument($Arguments).AddArgument($Sync).AddArgument($Ctx)
    $Sync.Step = ''
    $Sync.Percent = -1
    $ui.TxtBusyTitle.Text = $Title
    $ui.TxtBusyStep.Text = ''
    $ui.PrgBusy.IsIndeterminate = $true
    $ui.ResultCard.Visibility = 'Collapsed'
    $ui.BusyCard.Visibility = 'Visible'
    $ui.Overlay.Visibility = 'Visible'
    $ui.TxtStatus.Text = $Title
    $script:Job = @{ PS = $ps; RS = $rs; Handle = $ps.BeginInvoke(); OnDone = $OnDone }
}

function Complete-Work
{
    $job = $script:Job
    $script:Job = $null
    $value = $null
    $failure = $null
    try
    {
        $output = $job.PS.EndInvoke($job.Handle)
        $value = @($output | Where-Object { $_ -and $_.PSObject.Properties['__Result'] }) | Select-Object -Last 1
        if (-not $value) { $failure = 'The task finished without a result. See the activity log.' }
    }
    catch
    {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        $failure = $ex.Message
    }
    foreach ($e in $job.PS.Streams.Error) { Write-Log "$e" 'WARN' }
    $job.PS.Dispose()
    $job.RS.Dispose()
    $ui.BusyCard.Visibility = 'Collapsed'
    $ui.Overlay.Visibility = 'Collapsed'
    $ui.TxtStatus.Text = ''
    & $job.OnDone $value $failure
}

$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(120)
$timer.Add_Tick({
    try
    {
        $line = $null
        $sb = New-Object Text.StringBuilder
        while ($Sync.Log.TryDequeue([ref]$line)) { [void]$sb.AppendLine($line) }
        if ($sb.Length)
        {
            $ui.TxtLog.AppendText($sb.ToString())
            $ui.TxtLog.ScrollToEnd()
        }
        if ($script:UpdateCheck -and $script:UpdateCheck.Handle.IsCompleted) { Complete-UpdateCheck }
        if ($script:Job)
        {
            $ui.TxtBusyStep.Text = $Sync.Step
            if ($Sync.DeviceCode)
            {
                [Windows.Clipboard]::SetText($Sync.DeviceCode)
                Start-Process 'https://microsoft.com/devicelogin'
                $Sync.DeviceCode = $null
                $ui.TxtBusyStep.Text = $Sync.Step + '  (code copied, paste it in the browser)'
            }
            $pct = [int]$Sync.Percent
            if ($pct -ge 0) { $ui.PrgBusy.IsIndeterminate = $false; $ui.PrgBusy.Value = $pct } else { $ui.PrgBusy.IsIndeterminate = $true }
            if ($script:Job.Handle.IsCompleted) { Complete-Work }
        }
    }
    catch { Show-Error $_ }
})
#endregion

#region Printer list
$AllPrinters = @()
$LogoItems = @(Get-LogoList)
$ui.CmbLogo.ItemsSource = $LogoItems
$Overrides = @{}
$CurrentPrinter = $null
$Loading = $false
$Auth = $null
$Source = 'pc'

function Update-PrinterList
{
    param ([object[]]$Printers)
    # Function output is wrapped in a PSObject, which WPF can't show as an image; store the raw bitmap.
    foreach ($p in $Printers)
    {
        $img = [Windows.Media.ImageSource](New-Bitmap $p.LogoPath)
        $p.PSObject.Properties.Remove('LogoImage')
        $p.PSObject.Properties.Add([Management.Automation.PSNoteProperty]::new('LogoImage', $img))
    }
    $script:AllPrinters = @($Printers)
    $view = [Windows.Data.CollectionViewSource]::GetDefaultView($script:AllPrinters)
    $view.Filter = [Predicate[object]]{
        param ($item)
        $f = $ui.TxtFilter.Text.Trim()
        if (-not $f) { return $true }
        return ("$($item.Name) $($item.DriverName) $($item.Address) $($item.Location) $($item.Brand)" -like "*$f*")
    }
    $ui.GridPrinters.ItemsSource = $view
    $where = if ($script:Source -eq 'library') { 'in the library' } else { 'on this computer' }
    $ui.TxtCount.Text = if ($script:AllPrinters.Count -eq 1) { "1 printer $where" } else { "$($script:AllPrinters.Count) printers $where" }
    $ui.TxtEmpty.Text = if ($script:Source -ne 'library') { 'No printers found on this computer.' }
        elseif (-not $ui.CmbClient.SelectedItem) { 'Create a client first (Client > New).' }
        else { 'No printers for this client yet. Save them from This PC or use + New printer.' }
    $ui.TxtEmpty.Visibility = if ($script:AllPrinters.Count) { 'Collapsed' } else { 'Visible' }
}

function Start-PrinterScan
{
    $script:ReselectNames = @($ui.GridPrinters.SelectedItems | ForEach-Object { $_.Name })
    if ($script:Source -eq 'library')
    {
        $command = 'Get-PrinterLibrary'
        $arguments = @{ LibraryDir = (Get-ClientDir) }
        $title = 'Loading the printer library...'
    }
    else
    {
        $command = 'Get-PrinterInventory'
        $arguments = @{}
        $title = 'Loading printers...'
    }
    Start-Work -Title $title -Command $command -Arguments $arguments -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Could not read the printers'; return }
        Update-PrinterList $result.Printers
        if ($script:ReselectAfterLibrary) { $script:ReselectNames = @($script:ReselectAfterLibrary); $script:ReselectAfterLibrary = $null }
        foreach ($p in $script:AllPrinters) { if ($script:ReselectNames -contains $p.Name) { [void]$ui.GridPrinters.SelectedItems.Add($p) } }
        if ($Screenshot) { Invoke-ScreenshotReady }
    }
}

# Per-printer edits (name, IP, logo, app name) are kept while switching between printers.
function Get-PrinterSettings
{
    param ($Printer)
    $key = "$($Printer.Key)"
    if (-not $Overrides.ContainsKey($key))
    {
        $o = @{
            TargetName  = $Printer.Name
            HostAddress = $Printer.HostAddress
            AppName     = "Printer - $($Printer.Name)"
            LogoPath    = $Printer.LogoPath
            Advanced    = @{}
        }
        # Library entries start from what was saved with them.
        $saved = $Printer.SavedSettings
        if ($saved)
        {
            if ($saved.HostAddress) { $o.HostAddress = "$($saved.HostAddress)" }
            if ($saved.AppName) { $o.AppName = "$($saved.AppName)" }
            if ($saved.Advanced)
            {
                foreach ($prop in $saved.Advanced.PSObject.Properties) { $o.Advanced[$prop.Name] = $prop.Value }
            }
        }
        $Overrides[$key] = $o
    }
    return $Overrides[$key]
}

function Save-CurrentPrinterSettings
{
    if ($script:Loading -or -not $script:CurrentPrinter) { return }
    $o = Get-PrinterSettings $script:CurrentPrinter
    $o.TargetName = $ui.TxtTargetName.Text
    $o.HostAddress = $ui.TxtHost.Text
    $o.AppName = $ui.TxtAppName.Text
    if ($ui.CmbLogo.SelectedItem) { $o.LogoPath = $ui.CmbLogo.SelectedItem.Path }
}

function Update-SelectedPrinter
{
    $selected = @($ui.GridPrinters.SelectedItems)
    $script:CurrentPrinter = $null
    if ($selected.Count -eq 0)
    {
        $ui.PanelSettings.Visibility = 'Collapsed'
        $ui.PanelNoSelection.Visibility = 'Visible'
        return
    }
    $ui.PanelNoSelection.Visibility = 'Collapsed'
    $ui.PanelSettings.Visibility = 'Visible'

    if ($selected.Count -gt 1)
    {
        $ui.TxtPanelTitle.Text = "2. $($selected.Count) printers selected"
        $ui.PanelSingle.Visibility = 'Collapsed'
        $ui.PanelAppName.Visibility = 'Collapsed'
        $ui.PanelMulti.Visibility = 'Visible'
        $ui.ListMulti.ItemsSource = @(foreach ($p in $selected)
        {
            $o = Get-PrinterSettings $p
            $ip = "$($o.HostAddress)".Trim()
            [pscustomobject]@{
                Name        = $o.TargetName
                Logo        = [Windows.Media.ImageSource](New-Bitmap $o.LogoPath)
                Detail      = [string]$(if ($ip) { "$ip  -  $($p.DriverName)" } else { 'No IP address: click it to enter one' })
                DetailColor = [Windows.Media.Brush](ConvertTo-Brush $(if ($ip) { '#616161' } else { '#B54700' }))
            }
        })
        return
    }

    $p = $selected[0]
    $o = Get-PrinterSettings $p
    $script:Loading = $true
    try
    {
        $ui.TxtPanelTitle.Text = '2. Check the package'
        $ui.PanelMulti.Visibility = 'Collapsed'
        $ui.PanelSingle.Visibility = 'Visible'
        $ui.PanelAppName.Visibility = 'Visible'
        $ui.TxtSelName.Text = $p.Name
        $ui.TxtSelDriver.Text = $p.DriverName
        $ui.TxtTargetName.Text = $o.TargetName
        $ui.TxtHost.Text = $o.HostAddress
        $ui.TxtAppName.Text = $o.AppName
        $items = @($LogoItems)
        $match = $items | Where-Object { $_.Path -eq $o.LogoPath } | Select-Object -First 1
        if (-not $match)
        {
            $match = [pscustomobject]@{ Name = "Own: $([IO.Path]::GetFileNameWithoutExtension($o.LogoPath))"; Path = $o.LogoPath }
            $items += $match
        }
        $ui.CmbLogo.ItemsSource = $items
        $ui.CmbLogo.SelectedItem = $match
        $ui.ImgLogo.Source = New-Bitmap $o.LogoPath
        if ($p.PortKind -eq 'TCP/IP') { $ui.TxtHostHint.Visibility = 'Collapsed' }
        else
        {
            $ui.TxtHostHint.Text = "This printer uses a $($p.PortKind) port. Enter its IP address or hostname so other devices can reach it."
            $ui.TxtHostHint.Visibility = 'Visible'
        }
        Update-AdvancedSummary $o
    }
    finally { $script:Loading = $false }
    $script:CurrentPrinter = $p
}

$ui.GridPrinters.Add_SelectionChanged({ try { Update-SelectedPrinter } catch { Show-Error $_ } })
$ui.TxtFilter.Add_TextChanged({
    try
    {
        $ui.TxtFilterHint.Visibility = if ($ui.TxtFilter.Text) { 'Collapsed' } else { 'Visible' }
        $view = $ui.GridPrinters.ItemsSource
        if ($view) { $view.Refresh() }
    }
    catch { Show-Error $_ }
})
$ui.TxtTargetName.Add_TextChanged({
    if ($script:Loading -or -not $script:CurrentPrinter) { return }
    $o = Get-PrinterSettings $script:CurrentPrinter
    if ($ui.TxtAppName.Text -eq "Printer - $($o.TargetName)") { $ui.TxtAppName.Text = "Printer - $($ui.TxtTargetName.Text.Trim())" }
    Save-CurrentPrinterSettings
})
$ui.TxtHost.Add_TextChanged({ Save-CurrentPrinterSettings })
$ui.TxtAppName.Add_TextChanged({ Save-CurrentPrinterSettings })
$ui.CmbLogo.Add_SelectionChanged({
    $item = $ui.CmbLogo.SelectedItem
    $ui.ImgLogo.Source = if ($item) { New-Bitmap $item.Path } else { $null }
    Save-CurrentPrinterSettings
})
$ui.BtnRefresh.Add_Click({ try { Start-PrinterScan } catch { Show-Error $_ } })
#endregion

#region Intune sign-in
function Update-AccountUi
{
    if ($script:Auth)
    {
        $ui.TxtAccountUser.Text = $script:Auth.User
        $ui.TxtAccountTenant.Text = $script:Auth.TenantName
        $ui.PanelAccount.Visibility = 'Visible'
        $ui.TxtSignIn.Text = 'Sign out'
    }
    else
    {
        $ui.PanelAccount.Visibility = 'Collapsed'
        $ui.TxtSignIn.Text = 'Sign in to Intune'
    }
}

$SignInDone = {
    param ($result, $failure)
    if ($failure) { Show-Error $failure 'Sign-in failed'; $script:AfterSignIn = $null; return }
    $script:Auth = $result.Auth
    Update-AccountUi
    if ($script:AfterSignIn) { $next = $script:AfterSignIn; $script:AfterSignIn = $null; & $next }
}

function Start-SignIn
{
    param ([scriptblock]$Then)
    $script:AfterSignIn = $Then
    $Sync.DeviceCode = $null
    Start-Work -Title 'Signing in to Intune...' -Command 'Connect-IntuneBrowser' -OnDone {
        param ($result, $failure)
        if ($failure)
        {
            $answer = [Windows.MessageBox]::Show($window, "Browser sign-in did not work:`n$failure`n`nSign in with a code instead?", 'Printer Packager', 'YesNo', 'Question')
            if ("$answer" -eq 'Yes')
            {
                Start-Work -Title 'Signing in with a code...' -Command 'Connect-IntuneDeviceCode' -OnDone $SignInDone
            }
            else { $script:AfterSignIn = $null }
            return
        }
        & $SignInDone $result $null
    }
}

$ui.BtnSignIn.Add_Click({
    try
    {
        if ($script:Auth) { $script:Auth = $null; Update-AccountUi; Write-Log 'Signed out.' }
        else { Start-SignIn }
    }
    catch { Show-Error $_ }
})
#endregion

#region Actions
$ui.BtnLogoBrowse.Add_Click({
    try
    {
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Title = 'Choose a logo'
        $dlg.Filter = 'Images (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp'
        if ($dlg.ShowDialog($window))
        {
            $item = [pscustomobject]@{ Name = "Own: $([IO.Path]::GetFileNameWithoutExtension($dlg.FileName))"; Path = $dlg.FileName }
            $ui.CmbLogo.ItemsSource = @($LogoItems) + $item
            $ui.CmbLogo.SelectedItem = $item
        }
    }
    catch { Show-Error $_ }
})

$ui.BtnBrowseOutput.Add_Click({
    try
    {
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Where should the printer packages be saved?'
        $dlg.ShowNewFolderButton = $true
        if (Test-Path -LiteralPath $ui.TxtOutput.Text) { $dlg.SelectedPath = $ui.TxtOutput.Text }
        if ($dlg.ShowDialog() -eq 'OK') { $ui.TxtOutput.Text = $dlg.SelectedPath; Save-Settings }
    }
    catch { Show-Error $_ }
})

function Open-Folder
{
    param ([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
    Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$Path`""
}
$ui.BtnOpenOutput.Add_Click({ try { Open-Folder $ui.TxtOutput.Text.Trim() } catch { Show-Error $_ } })

$LastResult = $null
function Show-Result
{
    param ($Result)
    $script:LastResult = $Result
    $ui.ImgResultLogo.Source = New-Bitmap $Result.Logo
    $ui.TxtResultName.Text = "$($Result.DisplayName)  -  version $($Result.Version)"
    $ui.TxtResultFile.Text = if ($Result.IntuneWin) { $Result.IntuneWin } elseif ($Result.Zip) { $Result.Zip } else { $Result.Folder }
    $ui.TxtResultInstall.Text = $Result.InstallCommand
    $ui.TxtResultUninstall.Text = $Result.UninstallCommand
    $notes = 'Detection: custom script Detection.ps1. Logo: Logo.png. Everything is in the folder, see Intune-Upload-Info.txt.'
    if (-not $Result.DriverIncluded) { $notes += ' The driver is NOT included: devices need it already.' }
    if (-not $Result.IntuneWin) { $notes += ' No .intunewin was made: IntuneWinAppUtil.exe is not available offline. Put it next to PrinterPackager.exe.' }
    $ui.TxtResultNotes.Text = $notes
    $ui.BusyCard.Visibility = 'Collapsed'
    $ui.ResultListCard.Visibility = 'Collapsed'
    $ui.ResultCard.Visibility = 'Visible'
    $ui.Overlay.Visibility = 'Visible'
}
$ui.BtnCopyInstall.Add_Click({ [Windows.Clipboard]::SetText($ui.TxtResultInstall.Text) })
$ui.BtnCopyUninstall.Add_Click({ [Windows.Clipboard]::SetText($ui.TxtResultUninstall.Text) })
$ui.BtnResultOpen.Add_Click({ try { Open-Folder $script:LastResult.Folder } catch { Show-Error $_ } })
$ui.BtnResultClose.Add_Click({ $ui.ResultCard.Visibility = 'Collapsed'; $ui.Overlay.Visibility = 'Collapsed' })

function Show-ResultList
{
    param ([object[]]$Results, [string]$Mode)
    $ok = @($Results | Where-Object { $_.Status -ne 'Failed' }).Count
    $failed = $Results.Count - $ok
    switch ($Mode)
    {
        'Published'
        {
            $ui.TxtResultListTitle.Text = "Published $ok of $($Results.Count) printer(s)"
            $ui.TxtResultListSub.Text = 'Users find them in Company Portal (available for all users). It can take a few minutes before they show up.'
        }
        'Library'
        {
            $ui.TxtResultListTitle.Text = "Saved $ok of $($Results.Count) printer(s) to $($ui.CmbClient.SelectedItem.Name)"
            $ui.TxtResultListSub.Text = "Find them under Library. Use Export to take this client to another PC."
        }
        default
        {
            $ui.TxtResultListTitle.Text = "Saved $ok of $($Results.Count) printer(s)"
            $ui.TxtResultListSub.Text = "Files are in $($ui.TxtOutput.Text)."
        }
    }
    if ($failed) { $ui.TxtResultListSub.Text += " $failed failed, see below." }
    $ui.ListResults.ItemsSource = @(foreach ($r in $Results)
    {
        $detail = switch ($r.Status)
        {
            'Failed' { $r.Error }
            'Created' { "New app, version $($r.Version)" }
            'Updated' { "Existing app updated to version $($r.Version)" }
            'Library' { 'Saved to the library' }
            'SavedNoIntuneWin' { "Saved without .intunewin (IntuneWinAppUtil.exe not available offline), version $($r.Version)" }
            default { "Saved, version $($r.Version)" }
        }
        [pscustomobject]@{
            Printer        = $r.Printer
            Logo           = [Windows.Media.ImageSource](New-Bitmap $r.Logo)
            Detail         = [string]$detail
            DetailColor    = [Windows.Media.Brush](ConvertTo-Brush $(if ($r.Status -eq 'Failed') { '#C42B1C' } else { '#616161' }))
            Url            = [string]$r.Url
            LinkText       = 'Open in Intune'
            LinkVisibility = [Windows.Visibility]$(if ($r.Url) { 'Visible' } else { 'Collapsed' })
        }
    })
    $ui.BusyCard.Visibility = 'Collapsed'
    $ui.ResultCard.Visibility = 'Collapsed'
    $ui.ResultListCard.Visibility = 'Visible'
    $ui.Overlay.Visibility = 'Visible'
}
$ui.ListResults.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler]{
    param ($s, $e)
    $url = $e.OriginalSource.Tag
    if ($url) { Start-Process $url }
})
$ui.BtnResultListOpen.Add_Click({ try { Open-Folder $ui.TxtOutput.Text.Trim() } catch { Show-Error $_ } })
$ui.BtnResultListClose.Add_Click({ $ui.ResultListCard.Visibility = 'Collapsed'; $ui.Overlay.Visibility = 'Collapsed' })

function Get-PackageRequests
{
    # One request per ticked printer; returns $null (after telling the user) when something is missing.
    $selected = @($ui.GridPrinters.SelectedItems)
    if (-not $selected.Count) { [void][Windows.MessageBox]::Show($window, 'Tick at least one printer first.', 'Printer Packager', 'OK', 'Information'); return $null }
    $missing = @($selected | Where-Object { -not "$((Get-PrinterSettings $_).HostAddress)".Trim() } | ForEach-Object { $_.Name })
    if ($missing.Count)
    {
        [void][Windows.MessageBox]::Show($window, "These printers have no IP address yet. Click each one and fill it in:`n`n$($missing -join "`n")", 'Printer Packager', 'OK', 'Warning')
        return $null
    }
    Save-Settings
    return @(foreach ($p in $selected)
    {
        $o = Get-PrinterSettings $p
        @{
            Printer            = $p
            TargetName         = $o.TargetName
            HostAddress        = $o.HostAddress
            AppVendor          = $ui.TxtVendor.Text
            AppName            = $o.AppName
            AppVersion         = $ui.TxtVersion.Text
            LogoPath           = $o.LogoPath
            Advanced           = $o.Advanced
            CaptureDir         = $p.CaptureDir
            IncludeDriver      = [bool]$ui.ChkDriver.IsChecked
            IncludePreferences = [bool]$ui.ChkPrefs.IsChecked
            SetDefault         = [bool]$ui.ChkDefault.IsChecked
            CreateIntuneWin    = $true
            CreateZip          = [bool]$ui.ChkZip.IsChecked
            OutputDir          = $ui.TxtOutput.Text
        }
    })
}

function Start-SaveFiles
{
    $requests = Get-PackageRequests
    if (-not $requests) { return }
    foreach ($r in $requests) { $r.IntuneWinOptional = $true }
    $ui.ExpLog.IsExpanded = $true
    if ($requests.Count -eq 1)
    {
        Start-Work -Title "Packaging $($requests[0].TargetName)..." -Command 'New-PrinterPackage' -Arguments @{ Request = $requests[0] } -OnDone {
            param ($result, $failure)
            if ($failure) { Show-Error $failure 'Packaging failed'; return }
            Show-Result $result
        }
        return
    }
    Start-Work -Title "Packaging $($requests.Count) printers..." -Command 'Save-Printers' -Arguments @{ Requests = $requests } -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Packaging failed'; return }
        Show-ResultList $result.Results 'Saved'
    }
}

function Start-Publish
{
    $requests = Get-PackageRequests
    if (-not $requests) { return }
    if (-not $script:Auth) { Start-SignIn -Then { Start-Publish }; return }
    $ui.ExpLog.IsExpanded = $true
    $title = if ($requests.Count -eq 1) { "Publishing $($requests[0].TargetName)..." } else { "Publishing $($requests.Count) printers..." }
    Start-Work -Title $title -Command 'Publish-Printers' -Arguments @{ Auth = $script:Auth; Requests = $requests } -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Publishing failed'; return }
        Show-ResultList $result.Results 'Published'
    }
}
$ui.BtnPackage.Add_Click({ try { Start-Publish } catch { Show-Error $_ } })
$ui.BtnPackageLocal.Add_Click({ try { Start-SaveFiles } catch { Show-Error $_ } })
#endregion

#region Advanced settings
function Update-AdvancedSummary
{
    param ($Settings)
    $a = $Settings.Advanced
    $parts = @()
    if ($a.Duplex) { $parts += @{ OneSided = 'one-sided'; TwoSidedLongEdge = 'duplex long edge'; TwoSidedShortEdge = 'duplex short edge' }[$a.Duplex] }
    if ($a.Color) { $parts += @{ Color = 'colour'; Mono = 'black and white' }[$a.Color] }
    if ($a.PaperSize) { $parts += $a.PaperSize }
    if ($a.Protocol -or $a.PortNumber -or $a.LprQueueName -or $null -ne $a.SnmpEnabled) { $parts += 'port changed' }
    if ($null -ne $a.Location -or $null -ne $a.Comment) { $parts += 'location/comment' }
    $ui.TxtAdvancedSummary.Text = if ($parts) { 'Changed: ' + ($parts -join ', ') } else { '' }
}

function Show-AdvancedDialog
{
    param ($Printer, $Settings, [string]$SnapPath)
    [xml]$advXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Advanced settings" Width="520" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" Background="White" Foreground="#1B1B1B"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13" UseLayoutRounding="True">
    <Window.Resources>
        <Style TargetType="TextBlock" x:Key="Label">
            <Setter Property="Foreground" Value="#616161"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,0,0,4"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="6,5"/>
            <Setter Property="BorderBrush" Value="#D1D1D1"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="6,5"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="16,7"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#D1D1D1"/>
        </Style>
    </Window.Resources>
    <Border Padding="22,18,22,18" Background="White">
    <StackPanel>
        <TextBlock Name="TxtTitle" FontSize="16" FontWeight="SemiBold"/>
        <TextBlock Text="These settings are used on the devices. Leave on 'Keep' to use what was captured." Foreground="#616161" TextWrapping="Wrap" Margin="0,4,0,16"/>

        <TextBlock Text="Location" Style="{StaticResource Label}"/>
        <TextBox Name="TxtLocation" Margin="0,0,0,10"/>
        <TextBlock Text="Comment" Style="{StaticResource Label}"/>
        <TextBox Name="TxtComment" Margin="0,0,0,16"/>

        <TextBlock Text="Print defaults" FontWeight="SemiBold" Margin="0,0,0,8"/>
        <Grid Margin="0,0,0,16">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text="Duplex" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbDuplex">
                    <ComboBoxItem Content="Keep" Tag="" IsSelected="True"/>
                    <ComboBoxItem Content="One-sided" Tag="OneSided"/>
                    <ComboBoxItem Content="Long edge" Tag="TwoSidedLongEdge"/>
                    <ComboBoxItem Content="Short edge" Tag="TwoSidedShortEdge"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="2">
                <TextBlock Text="Colour" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbColor">
                    <ComboBoxItem Content="Keep" Tag="" IsSelected="True"/>
                    <ComboBoxItem Content="Colour" Tag="Color"/>
                    <ComboBoxItem Content="Black and white" Tag="Mono"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="4">
                <TextBlock Text="Paper size" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbPaper">
                    <ComboBoxItem Content="Keep" Tag="" IsSelected="True"/>
                    <ComboBoxItem Content="A4" Tag="A4"/>
                    <ComboBoxItem Content="A3" Tag="A3"/>
                    <ComboBoxItem Content="A5" Tag="A5"/>
                    <ComboBoxItem Content="Letter" Tag="Letter"/>
                    <ComboBoxItem Content="Legal" Tag="Legal"/>
                </ComboBox>
            </StackPanel>
        </Grid>

        <TextBlock Text="Port" FontWeight="SemiBold" Margin="0,0,0,8"/>
        <Grid Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text="Protocol" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbProtocol">
                    <ComboBoxItem Content="RAW" Tag="1"/>
                    <ComboBoxItem Content="LPR" Tag="2"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="2">
                <TextBlock Name="LblPortValue" Text="Port number" Style="{StaticResource Label}"/>
                <TextBox Name="TxtPortValue"/>
            </StackPanel>
        </Grid>
        <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
            <CheckBox Name="ChkSnmp" Content="SNMP status" VerticalAlignment="Center"/>
            <TextBlock Text="Community" Foreground="#616161" Margin="16,0,6,0" VerticalAlignment="Center"/>
            <TextBox Name="TxtCommunity" Width="120"/>
        </StackPanel>

        <Border Name="PanelVendor" Background="#F7F7F7" CornerRadius="6" Padding="12" Margin="0,0,0,16">
            <StackPanel>
                <TextBlock Text="Trays, finishing and other driver settings" FontWeight="SemiBold"/>
                <TextBlock Text="Open the printer's own settings, change them under Advanced &gt; Printing Defaults, click OK. They are captured when you package or save to the library." TextWrapping="Wrap" Foreground="#616161" Margin="0,4,0,8"/>
                <Button Name="BtnVendor" Content="Open printer settings..." HorizontalAlignment="Left"/>
            </StackPanel>
        </Border>

        <Grid>
            <Button Name="BtnReset" Content="Reset" HorizontalAlignment="Left"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="BtnCancel" Content="Cancel" Margin="0,0,10,0" IsCancel="True"/>
                <Button Name="BtnOk" Content="OK" IsDefault="True" Background="#C80032" Foreground="White" BorderBrush="#C80032"/>
            </StackPanel>
        </Grid>
    </StackPanel>
    </Border>
</Window>
'@
    $dlg = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $advXaml))
    $dlg.Owner = $window
    $d = @{}
    foreach ($node in $advXaml.SelectNodes('//*[@Name]')) { $d[$node.Name] = $dlg.FindName($node.Name) }
    $a = $Settings.Advanced

    function Select-Tag ($Combo, [string]$Tag)
    {
        foreach ($item in $Combo.Items) { if ("$($item.Tag)" -eq $Tag) { $Combo.SelectedItem = $item; return } }
    }
    $fill = {
        param ($adv)
        $d.TxtTitle.Text = "Advanced settings - $($Settings.TargetName)"
        $d.TxtLocation.Text = if ($null -ne $adv.Location) { $adv.Location } else { $Printer.Location }
        $d.TxtComment.Text = if ($null -ne $adv.Comment) { $adv.Comment } else { $Printer.Comment }
        Select-Tag $d.CmbDuplex "$($adv.Duplex)"
        Select-Tag $d.CmbColor "$($adv.Color)"
        Select-Tag $d.CmbPaper "$($adv.PaperSize)"
        $proto = if ($adv.Protocol) { [int]$adv.Protocol } elseif ($Printer.PortKind -eq 'TCP/IP' -and $Printer.Protocol) { [int]$Printer.Protocol } else { 1 }
        Select-Tag $d.CmbProtocol "$proto"
        $script:AdvPortNumber = if ($adv.PortNumber) { "$($adv.PortNumber)" } elseif ($Printer.PortNumber) { "$($Printer.PortNumber)" } else { '9100' }
        $script:AdvLprQueue = if ($adv.LprQueueName) { "$($adv.LprQueueName)" } elseif ($Printer.LprQueueName) { "$($Printer.LprQueueName)" } else { 'lp' }
        $d.TxtPortValue.Text = if ($proto -eq 2) { $script:AdvLprQueue } else { $script:AdvPortNumber }
        $d.LblPortValue.Text = if ($proto -eq 2) { 'Queue name' } else { 'Port number' }
        $d.ChkSnmp.IsChecked = if ($null -ne $adv.SnmpEnabled) { [bool]$adv.SnmpEnabled } else { [bool]$Printer.SnmpEnabled }
        $d.TxtCommunity.Text = if ($adv.SnmpCommunity) { $adv.SnmpCommunity } elseif ($Printer.SnmpCommunity) { $Printer.SnmpCommunity } else { 'public' }
    }
    & $fill $a

    $d.CmbProtocol.Add_SelectionChanged({
        $isLpr = "$($d.CmbProtocol.SelectedItem.Tag)" -eq '2'
        if ($isLpr) { $script:AdvPortNumber = $d.TxtPortValue.Text; $d.TxtPortValue.Text = $script:AdvLprQueue; $d.LblPortValue.Text = 'Queue name' }
        else { $script:AdvLprQueue = $d.TxtPortValue.Text; $d.TxtPortValue.Text = $script:AdvPortNumber; $d.LblPortValue.Text = 'Port number' }
    })
    if ($Printer.CaptureDir)
    {
        $d.PanelVendor.Visibility = 'Collapsed'
    }
    else
    {
        $d.BtnVendor.Add_Click({
            try { Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\rundll32.exe') -ArgumentList "printui.dll,PrintUIEntry /p /n `"$($Printer.Name)`"" }
            catch { Show-Error $_ }
        })
    }
    $d.BtnReset.Add_Click({ & $fill @{} })
    $d.BtnOk.Add_Click({
        $isLpr = "$($d.CmbProtocol.SelectedItem.Tag)" -eq '2'
        $value = $d.TxtPortValue.Text.Trim()
        $num = 0
        if (-not $isLpr -and (-not [int]::TryParse($value, [ref]$num) -or $num -lt 1 -or $num -gt 65535))
        {
            [void][Windows.MessageBox]::Show($dlg, 'The port number must be between 1 and 65535.', 'Advanced settings', 'OK', 'Warning')
            return
        }
        $new = @{}
        if ($d.TxtLocation.Text -ne "$($Printer.Location)") { $new.Location = $d.TxtLocation.Text }
        if ($d.TxtComment.Text -ne "$($Printer.Comment)") { $new.Comment = $d.TxtComment.Text }
        foreach ($pair in @(@('Duplex', $d.CmbDuplex), @('Color', $d.CmbColor), @('PaperSize', $d.CmbPaper)))
        {
            $tag = "$($pair[1].SelectedItem.Tag)"
            if ($tag) { $new[$pair[0]] = $tag }
        }
        $origProto = if ($Printer.PortKind -eq 'TCP/IP' -and $Printer.Protocol) { [int]$Printer.Protocol } else { 1 }
        $proto = [int]"$($d.CmbProtocol.SelectedItem.Tag)"
        if ($proto -ne $origProto) { $new.Protocol = $proto }
        if ($isLpr) { if ($value -and $value -ne "$($Printer.LprQueueName)") { $new.LprQueueName = $value } }
        elseif ($num -ne [int]$Printer.PortNumber -and -not ($num -eq 9100 -and -not $Printer.PortNumber)) { $new.PortNumber = $num }
        if ([bool]$d.ChkSnmp.IsChecked -ne [bool]$Printer.SnmpEnabled) { $new.SnmpEnabled = [bool]$d.ChkSnmp.IsChecked }
        if ([bool]$d.ChkSnmp.IsChecked -and $d.TxtCommunity.Text.Trim() -and $d.TxtCommunity.Text.Trim() -ne "$($Printer.SnmpCommunity)") { $new.SnmpCommunity = $d.TxtCommunity.Text.Trim() }
        $Settings.Advanced = $new
        $dlg.DialogResult = $true
    })
    if ($SnapPath)
    {
        # Self-test: render the dialog once, save it, close it.
        $dlg.Add_ContentRendered({ try { Save-WindowImage -Path $SnapPath -Target $dlg } finally { $dlg.Close() } })
    }
    [void]$dlg.ShowDialog()
    Update-AdvancedSummary $Settings
}

$ui.BtnAdvanced.Add_Click({
    try
    {
        if (-not $script:CurrentPrinter) { return }
        Show-AdvancedDialog -Printer $script:CurrentPrinter -Settings (Get-PrinterSettings $script:CurrentPrinter)
    }
    catch { Show-Error $_ }
})
#endregion

#region Library
function Set-Source
{
    param ([string]$Name)
    if ($script:Source -eq $Name) { return }
    $script:Source = $Name
    $isLibrary = $Name -eq 'library'
    $ui.PanelLibraryPath.Visibility = if ($isLibrary) { 'Visible' } else { 'Collapsed' }
    # Library entries carry their own captured driver and settings.
    $ui.ChkDriver.IsEnabled = -not $isLibrary
    $ui.ChkPrefs.IsEnabled = -not $isLibrary
    $ui.GridPrinters.UnselectAll()
    Start-PrinterScan
}
$ui.RbSourcePc.Add_Checked({ try { Set-Source 'pc' } catch { Show-Error $_ } })
$ui.RbSourceLibrary.Add_Checked({ try { Set-Source 'library' } catch { Show-Error $_ } })
$ui.BtnSelectAll.Add_Click({ $ui.GridPrinters.SelectAll() })

$ui.BtnLibraryBrowse.Add_Click({
    try
    {
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Choose the printer library folder (a share works fine).'
        $dlg.ShowNewFolderButton = $true
        if (Test-Path -LiteralPath $ui.TxtLibraryPath.Text) { $dlg.SelectedPath = $ui.TxtLibraryPath.Text }
        if ($dlg.ShowDialog() -eq 'OK')
        {
            $ui.TxtLibraryPath.Text = $dlg.SelectedPath
            Save-Settings
            Update-ClientList -Select $Settings.Client
            if ($script:Source -eq 'library') { $script:Overrides.Clear(); Start-PrinterScan }
        }
    }
    catch { Show-Error $_ }
})

function Start-SaveLibrary
{
    $requests = Get-PackageRequests
    if (-not $requests) { return }
    $clientDir = Get-ClientDirOrAsk
    if (-not $clientDir) { return }
    $ui.ExpLog.IsExpanded = $true
    $title = if ($requests.Count -eq 1) { "Saving $($requests[0].TargetName) to the library..." } else { "Saving $($requests.Count) printers to the library..." }
    Start-Work -Title $title -Command 'Save-PrinterToLibrary' -Arguments @{ Requests = $requests; LibraryDir = $clientDir } -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Saving to the library failed'; return }
        Update-ClientList -Select $Settings.Client
        if ($script:ExportAfterSave)
        {
            if (@($result.Results | Where-Object Status -eq 'Failed').Count) { $script:ExportAfterSave = $false; Show-ResultList $result.Results 'Library'; return }
            Start-ClientExport
            return
        }
        Show-ResultList $result.Results 'Library'
        if ($script:Source -eq 'library') { $script:Overrides.Clear(); Start-PrinterScan }
    }
}
$ui.BtnSaveLibrary.Add_Click({ try { Start-SaveLibrary } catch { Show-Error $_ } })
#endregion

#region App update
$UpdateInfo = $null
$UpdateCheck = $null

function Start-UpdateCheck
{
    # Quiet check in the background at start; only the button text changes.
    if (-not $ToolVersion) { return }
    $ps = [powershell]::Create()
    [void]$ps.AddScript($WorkerBootstrap).AddArgument($WorkerCodeText).AddArgument('Get-LatestRelease').AddArgument(@{}).AddArgument($null).AddArgument($Ctx)
    $script:UpdateCheck = @{ PS = $ps; Handle = $ps.BeginInvoke() }
}

function Complete-UpdateCheck
{
    $check = $script:UpdateCheck
    $script:UpdateCheck = $null
    try
    {
        $info = @($check.PS.EndInvoke($check.Handle)) | Where-Object { $_ -and $_.PSObject.Properties['__Result'] } | Select-Object -Last 1
        if ($info)
        {
            $script:UpdateInfo = $info
            if ($info.Newer)
            {
                $ui.TxtUpdate.Text = "Update to $($info.Tag)"
                $ui.BtnUpdate.BorderBrush = [Windows.Media.Brush](ConvertTo-Brush '#C80032')
                Write-Log "A new version is available: $($info.Tag) (this is $ToolVersion)."
            }
        }
    }
    catch { }
    finally { $check.PS.Dispose() }
}

function Start-AppUpdate
{
    param ($Info)
    $answer = [Windows.MessageBox]::Show($window, "Version $($Info.Tag) is available (you have $($Info.Current)).`n`nUpdate now? Printer Packager restarts afterwards.", 'Printer Packager', 'YesNo', 'Question')
    if ("$answer" -ne 'Yes') { return }
    Start-Work -Title "Updating to $($Info.Tag)..." -Command 'Install-AppUpdate' -Arguments @{ Url = $Info.Url; Tag = $Info.Tag } -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Update failed'; return }
        if ($result.Restart)
        {
            Start-Process -FilePath $result.Restart
            $window.Close()
            return
        }
        [void][Windows.MessageBox]::Show($window, "The new version could not replace this one here (no write access?).`n`nIt was saved to:`n$($result.SavedTo)`n`nCopy it over the old PrinterPackager.exe.", 'Printer Packager', 'OK', 'Information')
    }
}

$ui.BtnUpdate.Add_Click({
    try
    {
        if (-not $ToolVersion)
        {
            [void][Windows.MessageBox]::Show($window, 'Updating works from PrinterPackager.exe. This copy runs as a script.', 'Printer Packager', 'OK', 'Information')
            return
        }
        if ($script:UpdateInfo -and $script:UpdateInfo.Newer) { Start-AppUpdate $script:UpdateInfo; return }
        Start-Work -Title 'Checking for updates...' -Command 'Get-LatestRelease' -OnDone {
            param ($result, $failure)
            if ($failure) { Show-Error $failure 'Could not check for updates'; return }
            $script:UpdateInfo = $result
            if ($result.Newer) { Start-AppUpdate $result }
            else { [void][Windows.MessageBox]::Show($window, "You have the latest version ($ToolVersion).", 'Printer Packager', 'OK', 'Information') }
        }
    }
    catch { Show-Error $_ }
})
#endregion

#region New printer
function Show-NewPrinterDialog
{
    [xml]$npXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="New printer" Width="560" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" Background="White" Foreground="#1B1B1B"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13" UseLayoutRounding="True">
    <Window.Resources>
        <Style TargetType="TextBlock" x:Key="Label">
            <Setter Property="Foreground" Value="#616161"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Margin" Value="0,0,0,4"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="6,5"/>
            <Setter Property="BorderBrush" Value="#D1D1D1"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="6,5"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="16,7"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#D1D1D1"/>
        </Style>
    </Window.Resources>
    <Border Padding="22,18,22,18" Background="White">
    <StackPanel>
        <Grid Margin="0,0,0,16">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text="New printer" FontSize="16" FontWeight="SemiBold"/>
                <TextBlock Text="Set up a printer without having it installed. It is saved in the selected client, publish it from there." Foreground="#616161" TextWrapping="Wrap" Margin="0,4,0,0"/>
            </StackPanel>
            <Border Grid.Column="1" BorderBrush="#E5E5E5" BorderThickness="1" CornerRadius="8" Width="64" Height="64" Margin="12,0,0,0">
                <Image Name="ImgLogo" Margin="3" RenderOptions.BitmapScalingMode="HighQuality"/>
            </Border>
        </Grid>

        <Grid Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text="Printer name" Style="{StaticResource Label}"/>
                <TextBox Name="TxtName"/>
            </StackPanel>
            <StackPanel Grid.Column="2">
                <TextBlock Text="IP address or hostname" Style="{StaticResource Label}"/>
                <TextBox Name="TxtHost"/>
            </StackPanel>
        </Grid>

        <TextBlock Text="Driver" FontWeight="SemiBold" Margin="0,6,0,6"/>
        <RadioButton Name="RbInstalled" Content="A driver that is on this PC" IsChecked="True" Margin="0,0,0,6"/>
        <ComboBox Name="CmbInstalled" Margin="22,0,0,10" IsEditable="False"/>
        <RadioButton Name="RbFolder" Content="A driver from a folder (.inf)" Margin="0,0,0,6"/>
        <Grid Name="PanelFolder" Margin="22,0,0,4" IsEnabled="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Name="TxtInf" IsReadOnly="True"/>
            <Button Name="BtnInf" Grid.Column="1" Content="Browse..." Margin="8,0,0,0"/>
        </Grid>
        <ComboBox Name="CmbModel" Margin="22,4,0,0" IsEnabled="False"/>
        <TextBlock Name="TxtInfHint" Foreground="#B54700" FontSize="12" TextWrapping="Wrap" Margin="22,4,0,0" Visibility="Collapsed"/>

        <Grid Margin="0,14,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text="Location" Style="{StaticResource Label}"/>
                <TextBox Name="TxtLocation"/>
            </StackPanel>
            <StackPanel Grid.Column="2">
                <TextBlock Text="Comment" Style="{StaticResource Label}"/>
                <TextBox Name="TxtComment"/>
            </StackPanel>
        </Grid>

        <Grid Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text="Duplex" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbDuplex">
                    <ComboBoxItem Content="Driver default" Tag="" IsSelected="True"/>
                    <ComboBoxItem Content="One-sided" Tag="OneSided"/>
                    <ComboBoxItem Content="Long edge" Tag="TwoSidedLongEdge"/>
                    <ComboBoxItem Content="Short edge" Tag="TwoSidedShortEdge"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="2">
                <TextBlock Text="Colour" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbColor">
                    <ComboBoxItem Content="Driver default" Tag="" IsSelected="True"/>
                    <ComboBoxItem Content="Colour" Tag="Color"/>
                    <ComboBoxItem Content="Black and white" Tag="Mono"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="4">
                <TextBlock Text="Paper size" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbPaper">
                    <ComboBoxItem Content="Driver default" Tag="" IsSelected="True"/>
                    <ComboBoxItem Content="A4" Tag="A4"/>
                    <ComboBoxItem Content="A3" Tag="A3"/>
                    <ComboBoxItem Content="A5" Tag="A5"/>
                    <ComboBoxItem Content="Letter" Tag="Letter"/>
                    <ComboBoxItem Content="Legal" Tag="Legal"/>
                </ComboBox>
            </StackPanel>
        </Grid>

        <Grid Margin="0,0,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
                <TextBlock Text="Port protocol" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbProtocol">
                    <ComboBoxItem Content="RAW" Tag="1" IsSelected="True"/>
                    <ComboBoxItem Content="LPR" Tag="2"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Grid.Column="2">
                <TextBlock Name="LblPortValue" Text="Port number" Style="{StaticResource Label}"/>
                <TextBox Name="TxtPortValue" Text="9100"/>
            </StackPanel>
            <StackPanel Grid.Column="4">
                <TextBlock Text="Logo" Style="{StaticResource Label}"/>
                <ComboBox Name="CmbLogo" DisplayMemberPath="Name"/>
            </StackPanel>
        </Grid>
        <CheckBox Name="ChkSnmp" Content="SNMP status (community public)" Margin="0,0,0,16"/>

        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
            <Button Name="BtnCancel" Content="Cancel" Margin="0,0,10,0" IsCancel="True"/>
            <Button Name="BtnOk" Content="Save to client" IsDefault="True" Background="#C80032" Foreground="White" BorderBrush="#C80032"/>
        </StackPanel>
    </StackPanel>
    </Border>
</Window>
'@
    $dlg = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $npXaml))
    $dlg.Owner = $window
    $n = @{}
    foreach ($node in $npXaml.SelectNodes('//*[@Name]')) { $n[$node.Name] = $dlg.FindName($node.Name) }

    $drivers = @(Get-PrinterDriver -ErrorAction SilentlyContinue | Where-Object { $_.PrinterEnvironment -eq 'Windows x64' } | Sort-Object Name)
    $n.CmbInstalled.ItemsSource = @($drivers | ForEach-Object { $_.Name })
    if ($drivers.Count) { $n.CmbInstalled.SelectedIndex = 0 }
    $n.CmbLogo.ItemsSource = $LogoItems
    $script:NpLogoAuto = $true
    $script:NpInf = $null

    $updateLogo = {
        if (-not $script:NpLogoAuto) { return }
        $driver = if ($n.RbFolder.IsChecked) { "$($n.CmbModel.SelectedItem)" } else { "$($n.CmbInstalled.SelectedItem)" }
        $brand = Get-PrinterBrand -Text @($driver, $n.TxtName.Text)
        $match = $LogoItems | Where-Object { $_.Name -eq $brand.Name } | Select-Object -First 1
        if ($match) { $script:NpSettingLogo = $true; $n.CmbLogo.SelectedItem = $match; $script:NpSettingLogo = $false }
    }
    $n.CmbLogo.Add_SelectionChanged({
        if (-not $script:NpSettingLogo) { $script:NpLogoAuto = $false }
        $item = $n.CmbLogo.SelectedItem
        $n.ImgLogo.Source = if ($item) { New-Bitmap $item.Path } else { $null }
    })
    $n.CmbInstalled.Add_SelectionChanged({ & $updateLogo })
    $n.CmbModel.Add_SelectionChanged({ & $updateLogo })
    $n.TxtName.Add_TextChanged({ & $updateLogo })
    $n.RbInstalled.Add_Checked({ $n.CmbInstalled.IsEnabled = $true; $n.PanelFolder.IsEnabled = $false; $n.CmbModel.IsEnabled = $false; & $updateLogo })
    $n.RbFolder.Add_Checked({ $n.CmbInstalled.IsEnabled = $false; $n.PanelFolder.IsEnabled = $true; $n.CmbModel.IsEnabled = [bool]$script:NpInf; & $updateLogo })
    $n.CmbProtocol.Add_SelectionChanged({
        $isLpr = "$($n.CmbProtocol.SelectedItem.Tag)" -eq '2'
        $n.LblPortValue.Text = if ($isLpr) { 'Queue name' } else { 'Port number' }
        $n.TxtPortValue.Text = if ($isLpr) { 'lp' } else { '9100' }
    })
    $n.BtnInf.Add_Click({
        try
        {
            $ofd = New-Object Microsoft.Win32.OpenFileDialog
            $ofd.Title = 'Choose the printer driver .inf file (64-bit)'
            $ofd.Filter = 'Driver information (*.inf)|*.inf'
            if (-not $ofd.ShowDialog($dlg)) { return }
            $info = Get-InfPrinterModels -InfPath $ofd.FileName
            $hints = @()
            if (-not $info.IsPrinterInf) { $hints += 'This .inf is not a printer driver.' }
            if (-not $info.Has64Bit) { $hints += 'No 64-bit models found in this .inf.' }
            if (-not (Get-ChildItem -LiteralPath (Split-Path -Parent $ofd.FileName) -Filter '*.cat' -ErrorAction SilentlyContinue | Select-Object -First 1)) { $hints += 'No .cat file next to it: the driver is probably not signed and Windows will refuse it.' }
            $n.TxtInfHint.Text = $hints -join ' '
            $n.TxtInfHint.Visibility = if ($hints) { 'Visible' } else { 'Collapsed' }
            $script:NpInf = $ofd.FileName
            $n.TxtInf.Text = $ofd.FileName
            $n.CmbModel.ItemsSource = @($info.Models | ForEach-Object { $_.Name })
            $n.CmbModel.IsEnabled = $true
            if ($info.Models.Count) { $n.CmbModel.SelectedIndex = 0 }
            $script:NpManufacturer = if ($info.Models.Count) { $info.Models[0].Manufacturer } else { '' }
        }
        catch { Show-Error $_ 'Could not read the .inf' }
    })
    & $updateLogo

    $script:NpResult = $null
    $n.BtnOk.Add_Click({
        $isLpr = "$($n.CmbProtocol.SelectedItem.Tag)" -eq '2'
        $value = $n.TxtPortValue.Text.Trim()
        $num = 0
        $problem = $null
        if (-not $n.TxtName.Text.Trim()) { $problem = 'Enter a printer name.' }
        elseif ($n.TxtName.Text -match '[\\,!]') { $problem = 'The printer name cannot contain \ , or !.' }
        elseif ($n.TxtHost.Text.Trim() -notmatch '^[A-Za-z0-9.\-:_\[\]]+$') { $problem = 'Enter a valid IP address or hostname.' }
        elseif ($n.RbFolder.IsChecked -and (-not $script:NpInf -or -not $n.CmbModel.SelectedItem)) { $problem = 'Choose a driver .inf and a model.' }
        elseif (-not $n.RbFolder.IsChecked -and -not $n.CmbInstalled.SelectedItem) { $problem = 'Choose a driver.' }
        elseif (-not $isLpr -and (-not [int]::TryParse($value, [ref]$num) -or $num -lt 1 -or $num -gt 65535)) { $problem = 'The port number must be between 1 and 65535.' }
        if ($problem) { [void][Windows.MessageBox]::Show($dlg, $problem, 'New printer', 'OK', 'Warning'); return }

        $script:NpResult = @{
            Name          = $n.TxtName.Text.Trim()
            HostAddress   = $n.TxtHost.Text.Trim()
            DriverName    = if ($n.RbFolder.IsChecked) { "$($n.CmbModel.SelectedItem)" } else { "$($n.CmbInstalled.SelectedItem)" }
            InfPath       = if ($n.RbFolder.IsChecked) { $script:NpInf } else { '' }
            Manufacturer  = if ($n.RbFolder.IsChecked) { "$script:NpManufacturer" } else { '' }
            Location      = $n.TxtLocation.Text
            Comment       = $n.TxtComment.Text
            Protocol      = if ($isLpr) { 2 } else { 1 }
            PortNumber    = if ($isLpr) { 0 } else { $num }
            LprQueueName  = if ($isLpr) { $value } else { '' }
            SnmpEnabled   = [bool]$n.ChkSnmp.IsChecked
            Duplex        = "$($n.CmbDuplex.SelectedItem.Tag)"
            Color         = "$($n.CmbColor.SelectedItem.Tag)"
            PaperSize     = "$($n.CmbPaper.SelectedItem.Tag)"
            LogoPath      = if ($n.CmbLogo.SelectedItem) { $n.CmbLogo.SelectedItem.Path } else { '' }
        }
        $dlg.DialogResult = $true
    })
    if ($Screenshot -and -not $script:NpShotDone)
    {
        $dlg.Add_ContentRendered({ try { Save-WindowImage -Path "$Screenshot-newprinter.png" -Target $dlg } finally { $dlg.Close() } })
        $script:NpShotDone = $true
    }
    [void]$dlg.ShowDialog()
    return $script:NpResult
}

$ui.BtnNewPrinter.Add_Click({
    try
    {
        $clientDir = Get-ClientDirOrAsk
        if (-not $clientDir) { return }
        $definition = Show-NewPrinterDialog
        if (-not $definition) { return }
        Save-Settings
        Start-Work -Title "Saving $($definition.Name) to $($ui.CmbClient.SelectedItem.Name)..." -Command 'New-LibraryPrinter' -Arguments @{ Definition = $definition; LibraryDir = $clientDir } -OnDone {
            param ($result, $failure)
            if ($failure) { Show-Error $failure 'New printer failed'; return }
            $script:ReselectAfterLibrary = $result.Name
            Update-ClientList -Select $Settings.Client
            if ($script:Source -ne 'library') { $ui.RbSourceLibrary.IsChecked = $true }
            else { Start-PrinterScan }
        }
    }
    catch { Show-Error $_ }
})
#endregion

#region Clients
function Get-ClientDir
{
    $root = $ui.TxtLibraryPath.Text.Trim()
    $client = $ui.CmbClient.SelectedItem
    if (-not $root -or -not $client) { return $null }
    return Join-Path $root $client.Folder
}

function Update-ClientList
{
    param ([string]$Select)
    $root = $ui.TxtLibraryPath.Text.Trim()
    $clients = @()
    if ($root -and (Test-Path -LiteralPath $root))
    {
        $clients = @(Get-ChildItem -LiteralPath $root -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'client.json') } | ForEach-Object {
            $info = $null
            try { $info = [IO.File]::ReadAllText((Join-Path $_.FullName 'client.json')) | ConvertFrom-Json } catch { }
            $count = @(Get-ChildItem -LiteralPath $_.FullName -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'printer.json') }).Count
            $label = if ($info -and $info.Name) { "$($info.Name)" } else { $_.Name }
            [pscustomobject]@{
                Name      = $label
                Display   = if ($count -eq 1) { "$label  (1 printer)" } else { "$label  ($count printers)" }
                Count     = $count
                Folder    = $_.Name
                Publisher = if ($info -and $info.Publisher) { "$($info.Publisher)" } else { '' }
            }
        } | Sort-Object Name)
    }
    $script:LoadingClients = $true
    try
    {
        $ui.CmbClient.ItemsSource = $clients
        $pick = $clients | Where-Object { $_.Folder -eq $Select -or $_.Name -eq $Select } | Select-Object -First 1
        if (-not $pick -and $clients.Count) { $pick = $clients[0] }
        $ui.CmbClient.SelectedItem = $pick
    }
    finally { $script:LoadingClients = $false }
    Update-ClientSelection
}

function Update-ClientSelection
{
    $client = $ui.CmbClient.SelectedItem
    $ui.BtnExportClient.IsEnabled = [bool]$client
    if ($client)
    {
        $Settings.Client = $client.Folder
        if ($client.Publisher) { $ui.TxtVendor.Text = $client.Publisher }
    }
}

$ui.CmbClient.Add_SelectionChanged({
    try
    {
        if ($script:LoadingClients) { return }
        Update-ClientSelection
        Save-Settings
        if ($script:Source -eq 'library') { $script:Overrides.Clear(); Start-PrinterScan }
    }
    catch { Show-Error $_ }
})

function Show-NewClientDialog
{
    [xml]$ncXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="New client" Width="420" SizeToContent="Height" ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner" Background="White" Foreground="#1B1B1B"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="13" UseLayoutRounding="True">
    <Border Padding="22,18,22,18">
        <StackPanel>
            <TextBlock Text="New client" FontSize="16" FontWeight="SemiBold"/>
            <TextBlock Text="A project for one client site. Save its printers here and publish them later from anywhere." Foreground="#616161" TextWrapping="Wrap" Margin="0,4,0,14"/>
            <TextBlock Text="Client name" Foreground="#616161" FontSize="12" Margin="0,0,0,4"/>
            <TextBox Name="TxtName" Padding="6,5" Margin="0,0,0,10"/>
            <TextBlock Text="Publisher shown in Company Portal" Foreground="#616161" FontSize="12" Margin="0,0,0,4"/>
            <TextBox Name="TxtPublisher" Padding="6,5" Margin="0,0,0,16"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button Name="BtnCancel" Content="Cancel" Padding="16,7" Margin="0,0,10,0" IsCancel="True" Background="White"/>
                <Button Name="BtnOk" Content="Create" Padding="16,7" IsDefault="True" Background="#C80032" Foreground="White" BorderBrush="#C80032"/>
            </StackPanel>
        </StackPanel>
    </Border>
</Window>
'@
    $dlg = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $ncXaml))
    $dlg.Owner = $window
    $txtName = $dlg.FindName('TxtName')
    $txtPublisher = $dlg.FindName('TxtPublisher')
    $script:NcPublisherTouched = $false
    $txtName.Add_TextChanged({ if (-not $script:NcPublisherTouched) { $script:NcSetting = $true; $txtPublisher.Text = $txtName.Text; $script:NcSetting = $false } })
    $txtPublisher.Add_TextChanged({ if (-not $script:NcSetting) { $script:NcPublisherTouched = $true } })
    $script:NcResult = $null
    $dlg.FindName('BtnOk').Add_Click({
        $name = $txtName.Text.Trim()
        if (-not $name) { [void][Windows.MessageBox]::Show($dlg, 'Enter a client name.', 'New client', 'OK', 'Warning'); return }
        $script:NcResult = @{ Name = $name; Publisher = $(if ($txtPublisher.Text.Trim()) { $txtPublisher.Text.Trim() } else { $name }) }
        $dlg.DialogResult = $true
    })
    [void]$txtName.Focus()
    [void]$dlg.ShowDialog()
    return $script:NcResult
}

function New-Client
{
    # Returns the new client's folder name, or $null when cancelled.
    $c = Show-NewClientDialog
    if (-not $c) { return $null }
    $root = $ui.TxtLibraryPath.Text.Trim()
    if (-not $root) { throw 'Choose a library folder first.' }
    $folder = ConvertTo-SafeFileName $c.Name
    $dir = Join-Path $root $folder
    if (Test-Path -LiteralPath (Join-Path $dir 'client.json')) { throw "Client '$($c.Name)' already exists." }
    New-Item -Path $dir -ItemType Directory -Force | Out-Null
    $info = [ordered]@{ Name = $c.Name; Publisher = $c.Publisher; Created = (Get-Date).ToString('s'); CreatedBy = "$env:USERDOMAIN\$env:USERNAME" }
    Write-TextFile -Path (Join-Path $dir 'client.json') -Content ($info | ConvertTo-Json)
    Write-Log "Created client '$($c.Name)'." 'OK'
    Update-ClientList -Select $folder
    if ($script:Source -eq 'library') { $script:Overrides.Clear(); Start-PrinterScan }
    return $folder
}

function Get-ClientDirOrAsk
{
    # The selected client's folder; asks to create one when there is none yet.
    $dir = Get-ClientDir
    if ($dir) { return $dir }
    $answer = [Windows.MessageBox]::Show($window, 'Printers are saved per client. Create a client now?', 'Printer Packager', 'YesNo', 'Question')
    if ("$answer" -ne 'Yes') { return $null }
    if (-not (New-Client)) { return $null }
    return Get-ClientDir
}

$ui.BtnNewClient.Add_Click({ try { [void](New-Client) } catch { Show-Error $_ } })

function Start-ClientExport
{
    $dir = Get-ClientDir
    if (-not $dir) { return }
    $client = $ui.CmbClient.SelectedItem
    $saved = @(Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'printer.json') }).Count
    $ticked = @($ui.GridPrinters.SelectedItems).Count
    if ($script:Source -eq 'pc' -and $ticked -and -not $script:ExportAfterSave)
    {
        # Ticked printers on this PC are not in the client yet: offer to save them first, then export.
        $answer = [Windows.MessageBox]::Show($window, "Save the $ticked ticked printer(s) to $($client.Name) first?`n`nExport only takes printers that are saved to the client.", 'Printer Packager', 'YesNoCancel', 'Question')
        if ("$answer" -eq 'Cancel') { return }
        if ("$answer" -eq 'Yes') { $script:ExportAfterSave = $true; Start-SaveLibrary; return }
    }
    $script:ExportAfterSave = $false
    if (-not $saved)
    {
        [void][Windows.MessageBox]::Show($window, "$($client.Name) has no printers yet.`n`nTick printers under This PC and click Save to library (or use + New printer), then export.", 'Printer Packager', 'OK', 'Information')
        return
    }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Title = 'Export client'
    $dlg.Filter = 'Printer Packager client (*.zip)|*.zip'
    $dlg.FileName = "$($client.Folder) printers.zip"
    if (-not $dlg.ShowDialog($window)) { return }
    $script:ExportCount = $saved
    Start-Work -Title "Exporting $($client.Name)..." -Command 'Export-ClientProject' -Arguments @{ ClientDir = $dir; ZipPath = $dlg.FileName } -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Export failed'; return }
        [void][Windows.MessageBox]::Show($window, "Exported $($script:ExportCount) printer(s) to:`n$($result.Zip)`n`nOpen it on another PC with Import.", 'Printer Packager', 'OK', 'Information')
    }
}
$ui.BtnExportClient.Add_Click({ try { Start-ClientExport } catch { Show-Error $_ } })

function Start-ClientImport
{
    param ([string]$Zip, [bool]$Overwrite)
    Start-Work -Title 'Importing client...' -Command 'Import-ClientProject' -Arguments @{ ZipPath = $Zip; LibraryRoot = $ui.TxtLibraryPath.Text.Trim(); Overwrite = $Overwrite } -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Import failed'; return }
        if ($result.Exists)
        {
            $answer = [Windows.MessageBox]::Show($window, "Client '$($result.Name)' already exists in your library. Replace it with the imported one?", 'Printer Packager', 'YesNo', 'Question')
            if ("$answer" -eq 'Yes') { Start-ClientImport -Zip $script:ImportZip -Overwrite $true }
            return
        }
        Update-ClientList -Select $result.Name
        if ($script:Source -ne 'library') { $ui.RbSourceLibrary.IsChecked = $true } else { $script:Overrides.Clear(); Start-PrinterScan }
    }
}

$ui.BtnImportClient.Add_Click({
    try
    {
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.Title = 'Import client'
        $dlg.Filter = 'Printer Packager client (*.zip)|*.zip'
        if (-not $dlg.ShowDialog($window)) { return }
        $script:ImportZip = $dlg.FileName
        Start-ClientImport -Zip $dlg.FileName -Overwrite $false
    }
    catch { Show-Error $_ }
})
#endregion

Update-ClientList -Select $Settings.Client

#region Screenshot self-test
function Save-WindowImage
{
    param ([string]$Path, $Target = $window)
    $Target.UpdateLayout()
    $source = [Windows.PresentationSource]::FromVisual($Target)
    $scale = if ($source) { $source.CompositionTarget.TransformToDevice.M11 } else { 1 }
    $content = $Target.Content
    $rtb = [Windows.Media.Imaging.RenderTargetBitmap]::new([int]($content.ActualWidth * $scale), [int]($content.ActualHeight * $scale), [double](96 * $scale), [double](96 * $scale), [Windows.Media.PixelFormats]::Pbgra32)
    $rtb.Render($content)
    $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $fs = [IO.File]::Create($Path)
    try { $enc.Save($fs) } finally { $fs.Dispose() }
}

function Invoke-ScreenshotReady
{
    if ($script:AllPrinters.Count -and -not $ui.GridPrinters.SelectedItem)
    {
        $pick = $script:AllPrinters | Where-Object { $_.Brand -ne 'Generic printer' } | Select-Object -First 1
        if (-not $pick) { $pick = $script:AllPrinters[0] }
        $ui.GridPrinters.SelectedItem = $pick
    }
    $t = New-Object Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(800)
    $t.Add_Tick({
        $this.Stop()
        try { Save-WindowImage "$Screenshot-ready.png" } catch { Write-Log "Screenshot failed: $_" 'ERROR' }
        try
        {
            if ($script:CurrentPrinter) { Show-AdvancedDialog -Printer $script:CurrentPrinter -Settings (Get-PrinterSettings $script:CurrentPrinter) -SnapPath "$Screenshot-advanced.png" }
        }
        catch { Write-Log "Advanced dialog failed: $_" 'ERROR' }
        try { [void](Show-NewPrinterDialog) } catch { Write-Log "New printer dialog failed: $_" 'ERROR' }
        $window.Close()
    })
    $t.Start()
}
#endregion

$window.Add_ContentRendered({
    try
    {
        if ($OnReady) { try { $OnReady.Invoke() } catch { } }
        $window.Activate() | Out-Null
        $timer.Start()
        Start-PrinterScan
        Start-UpdateCheck
        if ($Screenshot)
        {
            $window.Dispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::ApplicationIdle, [Action]{ try { Save-WindowImage "$Screenshot-loading.png" } catch { } }) | Out-Null
        }
    }
    catch { Show-Error $_ }
})
$window.Add_Closing({ Save-Settings })

Write-Log "Printer Packager started (PSADT $($Ctx.PsadtVersion)). Running as $([Security.Principal.WindowsIdentity]::GetCurrent().Name)."
[void]$window.ShowDialog()
$timer.Stop()
