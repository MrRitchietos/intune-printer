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
    if ($InfPath -and (Split-Path -Leaf $InfPath) -match '^(ntprint4?|prnms\d+)\.inf$')
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
        ## 1. Driver: add the packaged driver to the DriverStore, then to the spooler.
        if ($config.DriverFolder)
        {
            $driverPath = Join-Path -Path $adtSession.DirFiles -ChildPath $config.DriverFolder
            if (Test-Path -LiteralPath $driverPath)
            {
                $stagedDriver = Join-Path -Path $staging -ChildPath 'Driver'
                Copy-Item -LiteralPath $driverPath -Destination $stagedDriver -Recurse -Force
                foreach ($inf in (Get-ChildItem -LiteralPath $stagedDriver -Filter '*.inf' -Recurse))
                {
                    Write-ADTLogEntry -Message "Adding driver package [$($inf.Name)] to the DriverStore."
                    $result = Start-ADTProcess -FilePath (Join-Path -Path $system32 -ChildPath 'pnputil.exe') -ArgumentList '/add-driver', $inf.FullName -CreateNoWindow -IgnoreExitCodes '*' -PassThru
                    Write-ADTLogEntry -Message "pnputil exit code [$($result.ExitCode)]."
                }
            }
        }

        if (!(Get-PrinterDriver -Name $config.DriverName -ErrorAction SilentlyContinue))
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
            $result = Start-ADTProcess -FilePath $printBrm -ArgumentList '-R', '-F', $stagedExport, '-O', 'FORCE' -CreateNoWindow -IgnoreExitCodes '*' -PassThru
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
            Start-ADTProcess -FilePath (Join-Path -Path $system32 -ChildPath 'rundll32.exe') -ArgumentList 'printui.dll,PrintUIEntry', '/Sr', '/n', $config.PrinterName, '/a', $stagedDat, 'g', 'd', 'r', '/q' -CreateNoWindow -IgnoreExitCodes '*'
        }

        ## 6. Basic print configuration.
        $printConfig = @{}
        if ($config.DuplexingMode) { $printConfig.DuplexingMode = $config.DuplexingMode }
        if ($config.PaperSize) { $printConfig.PaperSize = $config.PaperSize }
        if ($null -ne $config.Color) { $printConfig.Color = [System.Boolean]$config.Color }
        if ($printConfig.Count)
        {
            try { Set-PrintConfiguration -PrinterName $config.PrinterName @printConfig -ErrorAction Stop }
            catch { Write-ADTLogEntry -Message "Print configuration not applied: $($_.Exception.Message)" -Severity 2 }
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
                    Start-ADTProcessAsUser -FilePath (Join-Path -Path $system32 -ChildPath 'rundll32.exe') -ArgumentList $printUi -IgnoreExitCodes '*'
                }
                else
                {
                    Start-ADTProcess -FilePath (Join-Path -Path $system32 -ChildPath 'rundll32.exe') -ArgumentList $printUi -IgnoreExitCodes '*'
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
    $keepPort = ($p.PortKind -eq 'TCP/IP' -and $hostAddress -eq $p.HostAddress)
    $portName = if ($keepPort) { $p.PortName } else { "IP_$hostAddress" }

    $displayName = "$appVendor - $appName"
    $folderName = ConvertTo-SafeFileName "$appVendor - $appName - $appVersion"
    $root = Join-Path $outDir $folderName
    $packageDir = Join-Path $root 'Package'
    $filesDir = Join-Path $packageDir 'Files'

    Set-Step 'Preparing...' 2
    $template = Resolve-PsadtTemplate
    $intuneWinUtil = if ($Request.CreateIntuneWin) { Resolve-IntuneWinAppUtil } else { $null }

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

    $driverFolder = ''
    $driverInf = ''
    $builtInDriver = $false
    if ($Request.IncludeDriver)
    {
        Set-Step "Collecting driver files for '$($p.DriverName)'..." 15
        $driverDir = Join-Path $filesDir 'Driver'
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
    if ($Request.IncludePreferences)
    {
        Set-Step 'Capturing printer settings (PrintBrm)...' 35
        $usePrintBrm = Export-PrinterPreferences -PrinterName $p.Name -PortName $p.PortName -KeepPort $keepPort -DriverName $p.DriverName -DestinationFile (Join-Path $filesDir 'printer.printerExport')

        Set-Step 'Capturing printer preferences (PrintUI)...' 50
        $useSettingsDat = Export-PrinterSettingsDat -PrinterName $p.Name -DestinationFile (Join-Path $filesDir 'printer_settings.dat')

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

    Set-Step 'Writing deployment script...' 60
    $config = [ordered]@{
        AppName            = $appName
        AppVendor          = $appVendor
        AppVersion         = $appVersion
        PrinterName        = $targetName
        SourcePrinterName  = $p.Name
        DriverName         = $p.DriverName
        DriverFolder       = $driverFolder
        DriverInf          = "$driverInf"
        PortName           = $portName
        PrinterHostAddress = $hostAddress
        Protocol           = if ($keepPort) { $p.Protocol } else { 1 }
        PortNumber         = if ($keepPort) { $p.PortNumber } else { 9100 }
        LprQueueName       = if ($keepPort) { $p.LprQueueName } else { '' }
        SnmpEnabled        = if ($keepPort) { $p.SnmpEnabled } else { $false }
        SnmpCommunity      = $p.SnmpCommunity
        SnmpIndex          = $p.SnmpIndex
        Comment            = $p.Comment
        Location           = $p.Location
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
$GraphClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
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
    $next = "deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')&`$select=id,displayName,displayVersion,notes"
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
#endregion

}
. $WorkerCode
$WorkerCodeText = $WorkerCode.ToString()
#endregion

#region Headless mode
if ($IsHeadless)
{
    $ErrorActionPreference = 'Stop'
    $inventory = Get-PrinterInventory
    $printer = $inventory.Printers | Where-Object { $_.Name -eq $PrinterName } | Select-Object -First 1
    if (-not $printer) { throw "Printer '$PrinterName' was not found. Available: $(($inventory.Printers | ForEach-Object Name) -join ', ')" }
    if (-not $OutputDir) { $OutputDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Printer Packages' }
    $result = New-PrinterPackage -Request @{
        Printer            = $printer
        TargetName         = $printer.Name
        HostAddress        = if ($PrinterHostAddress) { $PrinterHostAddress } else { $printer.HostAddress }
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
    }
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
                <Image Name="ImgBrand" Height="34" VerticalAlignment="Center" RenderOptions.BitmapScalingMode="HighQuality"/>
                <Border Grid.Column="1" BorderBrush="{StaticResource Line}" BorderThickness="1,0,0,0" Margin="20,2,0,2" Padding="20,0,0,0">
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock Text="Printer Packager" FontSize="18" FontWeight="SemiBold"/>
                        <TextBlock Text="Turn an installed printer into an Intune app (PSAppDeployToolkit v4)" Foreground="{StaticResource Muted}"/>
                    </StackPanel>
                </Border>
                <Button Name="BtnRefresh" Grid.Column="2" VerticalAlignment="Center">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="&#xE72C;" FontFamily="Segoe MDL2 Assets" Margin="0,1,8,0" VerticalAlignment="Center"/>
                        <TextBlock Text="Refresh printers"/>
                    </StackPanel>
                </Button>
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
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="240"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel VerticalAlignment="Center">
                            <TextBlock Text="1. Choose a printer" FontSize="14" FontWeight="SemiBold"/>
                            <TextBlock Name="TxtCount" Text="Loading..." Foreground="{StaticResource Muted}" FontSize="12"/>
                        </StackPanel>
                        <Grid Grid.Column="1">
                            <TextBox Name="TxtFilter" Padding="30,6,8,6"/>
                            <TextBlock Text="&#xE721;" FontFamily="Segoe MDL2 Assets" Foreground="{StaticResource Muted}" Margin="10,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                            <TextBlock Name="TxtFilterHint" Text="Search" Foreground="#9E9E9E" Margin="31,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"/>
                        </Grid>
                    </Grid>
                    <DataGrid Name="GridPrinters" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single"
                              HeadersVisibility="Column" GridLinesVisibility="None" BorderThickness="0" Background="White"
                              RowHeight="48" CanUserResizeRows="False" CanUserReorderColumns="False" HorizontalScrollBarVisibility="Disabled" Margin="8,0,8,8">
                        <DataGrid.Columns>
                            <DataGridTemplateColumn Width="52" CanUserSort="False">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <Image Source="{Binding LogoPath}" Width="30" Height="30" RenderOptions.BitmapScalingMode="HighQuality"/>
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
                            <TextBlock Text="2. Check the package" Style="{StaticResource SectionTitle}"/>
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
                            <TextBlock Text="App name" Style="{StaticResource Label}" Margin="0,10,0,4"/>
                            <TextBox Name="TxtAppName"/>

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
                        <TextBlock Text="Select a printer on the left" FontSize="14" FontWeight="SemiBold" HorizontalAlignment="Center" Margin="0,12,0,4"/>
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
                <Button Name="BtnPackageLocal" Grid.Column="2" Content="Create .zip only" Margin="24,0,10,0" Padding="16,9"/>
                <Button Name="BtnPackage" Grid.Column="3" Style="{StaticResource AccentButton}" Content="Create Intune package"/>
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

$ui.ImgBrand.Source = New-Bitmap (Join-Path $Ctx.BrandingDir 'connectify-logo.png')
$window.Icon = New-Bitmap (Join-Path $Ctx.BrandingDir 'connectify-icon.png')
#endregion

#region Settings
$SettingsFile = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'IntunePrinterPackager\settings.json'
$Settings = [pscustomobject]@{
    Vendor             = 'Company'
    OutputDir          = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Printer Packages'
    IncludeDriver      = $true
    IncludePreferences = $true
    CreateZip          = $false
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

$ui.TxtVendor.Text = $Settings.Vendor
$ui.TxtOutput.Text = $Settings.OutputDir
$ui.ChkDriver.IsChecked = [bool]$Settings.IncludeDriver
$ui.ChkPrefs.IsChecked = [bool]$Settings.IncludePreferences
$ui.ChkZip.IsChecked = [bool]$Settings.CreateZip

function Save-Settings
{
    try
    {
        $Settings.Vendor = $ui.TxtVendor.Text.Trim()
        $Settings.OutputDir = $ui.TxtOutput.Text.Trim()
        $Settings.IncludeDriver = [bool]$ui.ChkDriver.IsChecked
        $Settings.IncludePreferences = [bool]$ui.ChkPrefs.IsChecked
        $Settings.CreateZip = [bool]$ui.ChkZip.IsChecked
        New-Item -Path (Split-Path -Parent $SettingsFile) -ItemType Directory -Force | Out-Null
        [IO.File]::WriteAllText($SettingsFile, ($Settings | ConvertTo-Json))
    }
    catch { }
}
#endregion

#region Background jobs
$Sync = [hashtable]::Synchronized(@{
    Log     = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
    Step    = ''
    Percent = -1
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
        if ($script:Job)
        {
            $ui.TxtBusyStep.Text = $Sync.Step
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
$CustomLogo = $null

function Update-PrinterList
{
    param ([object[]]$Printers)
    $script:AllPrinters = @($Printers)
    $view = [Windows.Data.CollectionViewSource]::GetDefaultView($script:AllPrinters)
    $view.Filter = [Predicate[object]]{
        param ($item)
        $f = $ui.TxtFilter.Text.Trim()
        if (-not $f) { return $true }
        return ("$($item.Name) $($item.DriverName) $($item.Address) $($item.Location) $($item.Brand)" -like "*$f*")
    }
    $ui.GridPrinters.ItemsSource = $view
    $ui.TxtCount.Text = if ($script:AllPrinters.Count -eq 1) { '1 printer on this computer' } else { "$($script:AllPrinters.Count) printers on this computer" }
    $ui.TxtEmpty.Visibility = if ($script:AllPrinters.Count) { 'Collapsed' } else { 'Visible' }
}

function Start-PrinterScan
{
    $script:ReselectName = if ($ui.GridPrinters.SelectedItem) { $ui.GridPrinters.SelectedItem.Name } else { $null }
    Start-Work -Title 'Loading printers...' -Command 'Get-PrinterInventory' -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Could not read the printers'; return }
        Update-PrinterList $result.Printers
        if ($script:ReselectName)
        {
            $match = $script:AllPrinters | Where-Object { $_.Name -eq $script:ReselectName } | Select-Object -First 1
            if ($match) { $ui.GridPrinters.SelectedItem = $match; $ui.GridPrinters.ScrollIntoView($match) }
        }
        if ($Screenshot) { Invoke-ScreenshotReady }
    }
}

$LastAutoAppName = ''
function Update-SelectedPrinter
{
    $p = $ui.GridPrinters.SelectedItem
    if (-not $p)
    {
        $ui.PanelSettings.Visibility = 'Collapsed'
        $ui.PanelNoSelection.Visibility = 'Visible'
        return
    }
    $ui.PanelNoSelection.Visibility = 'Collapsed'
    $ui.PanelSettings.Visibility = 'Visible'
    $ui.TxtSelName.Text = $p.Name
    $ui.TxtSelDriver.Text = $p.DriverName
    $ui.TxtTargetName.Text = $p.Name
    $ui.TxtHost.Text = $p.HostAddress
    $script:LastAutoAppName = "Printer - $($p.Name)"
    $ui.TxtAppName.Text = $script:LastAutoAppName
    $script:CustomLogo = $null
    $ui.CmbLogo.ItemsSource = $LogoItems
    $ui.CmbLogo.SelectedItem = $LogoItems | Where-Object { $_.Path -eq $p.LogoPath } | Select-Object -First 1
    if ($p.PortKind -eq 'TCP/IP')
    {
        $ui.TxtHostHint.Visibility = 'Collapsed'
    }
    else
    {
        $ui.TxtHostHint.Text = "This printer uses a $($p.PortKind) port. Enter its IP address or hostname so other devices can reach it."
        $ui.TxtHostHint.Visibility = 'Visible'
    }
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
    if ($ui.TxtAppName.Text -eq $script:LastAutoAppName)
    {
        $script:LastAutoAppName = "Printer - $($ui.TxtTargetName.Text.Trim())"
        $ui.TxtAppName.Text = $script:LastAutoAppName
    }
})
$ui.CmbLogo.Add_SelectionChanged({
    $item = $ui.CmbLogo.SelectedItem
    $ui.ImgLogo.Source = if ($item) { New-Bitmap $item.Path } else { $null }
})
$ui.BtnRefresh.Add_Click({ try { Start-PrinterScan } catch { Show-Error $_ } })
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
    $ui.TxtResultNotes.Text = $notes
    $ui.BusyCard.Visibility = 'Collapsed'
    $ui.ResultCard.Visibility = 'Visible'
    $ui.Overlay.Visibility = 'Visible'
}
$ui.BtnCopyInstall.Add_Click({ [Windows.Clipboard]::SetText($ui.TxtResultInstall.Text) })
$ui.BtnCopyUninstall.Add_Click({ [Windows.Clipboard]::SetText($ui.TxtResultUninstall.Text) })
$ui.BtnResultOpen.Add_Click({ try { Open-Folder $script:LastResult.Folder } catch { Show-Error $_ } })
$ui.BtnResultClose.Add_Click({ $ui.ResultCard.Visibility = 'Collapsed'; $ui.Overlay.Visibility = 'Collapsed' })

function Start-Package
{
    param ([bool]$Intune)
    $p = $ui.GridPrinters.SelectedItem
    if (-not $p) { [void][Windows.MessageBox]::Show($window, 'Select a printer first.', 'Printer Packager', 'OK', 'Information'); return }
    $logo = $ui.CmbLogo.SelectedItem
    Save-Settings
    $request = @{
        Printer            = $p
        TargetName         = $ui.TxtTargetName.Text
        HostAddress        = $ui.TxtHost.Text
        AppVendor          = $ui.TxtVendor.Text
        AppName            = $ui.TxtAppName.Text
        AppVersion         = $ui.TxtVersion.Text
        LogoPath           = if ($logo) { $logo.Path } else { $p.LogoPath }
        IncludeDriver      = [bool]$ui.ChkDriver.IsChecked
        IncludePreferences = [bool]$ui.ChkPrefs.IsChecked
        SetDefault         = [bool]$ui.ChkDefault.IsChecked
        CreateIntuneWin    = $Intune
        CreateZip          = (-not $Intune) -or [bool]$ui.ChkZip.IsChecked
        OutputDir          = $ui.TxtOutput.Text
    }
    $ui.ExpLog.IsExpanded = $true
    Start-Work -Title "Packaging $($ui.TxtTargetName.Text.Trim())..." -Command 'New-PrinterPackage' -Arguments @{ Request = $request } -OnDone {
        param ($result, $failure)
        if ($failure) { Show-Error $failure 'Packaging failed'; return }
        Show-Result $result
    }
}
$ui.BtnPackage.Add_Click({ try { Start-Package -Intune $true } catch { Show-Error $_ } })
$ui.BtnPackageLocal.Add_Click({ try { Start-Package -Intune $false } catch { Show-Error $_ } })
#endregion

#region Screenshot self-test
function Save-WindowImage
{
    param ([string]$Path)
    $window.UpdateLayout()
    $source = [Windows.PresentationSource]::FromVisual($window)
    $scale = if ($source) { $source.CompositionTarget.TransformToDevice.M11 } else { 1 }
    $content = $window.Content
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
