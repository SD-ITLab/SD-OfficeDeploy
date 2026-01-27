[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ProductID,

    [Parameter(Mandatory = $false)]
    [string]$Language = "de-de",

    [Parameter(Mandatory = $false)]
    [ValidateSet("32", "64", "ARM64")]
    [string]$Arch = "64",

    [Parameter(Mandatory = $false)]
    [string]$Channel = "Current",

    [string[]]$ExcludeApps,
    [string[]]$ShortcutApps,

    # Debug: Temp-Ordner nach der Installation NICHT löschen
    [switch]$KeepTemp
)

# ------------------------------------------------------------
# Globales Verhalten / Defaults
# ------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'
$PSDefaultParameterValues['Invoke-WebRequest:UseBasicParsing'] = $true
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Write-Host ""
Write-Host "=== CLS Office Installer – Retail ==="
Write-Host "Produkt-ID : $ProductID"
Write-Host "Sprache    : $Language"
Write-Host "Architektur: $Arch"
Write-Host "Channel    : $Channel"
if ($ExcludeApps)  { Write-Host "ExcludeApps : $($ExcludeApps -join ', ')" }
if ($ShortcutApps) { Write-Host "Shortcuts   : $($ShortcutApps -join ', ')" }
Write-Host ""

# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------

function Fail-And-Exit {
    param(
        [string]$Message,
        [int]$Code = 1
    )
    Write-Host ""
    Write-Host "FEHLER: $Message" -ForegroundColor Red
    exit $Code
}

function Get-ExistingOffice {
    $keys = @(
        'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration'
    )

    foreach ($k in $keys) {
        if (Test-Path $k) {
            try {
                return Get-ItemProperty -Path $k -ErrorAction Stop
            } catch {}
        }
    }
    return $null
}

function Get-ODTUrl {
    [CmdletBinding()]
    param()

    $downloadPage = "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
    Write-Host "Lade ODT-Downloadseite: $downloadPage"

    try {
        $resp = Invoke-WebRequest -Uri $downloadPage -ErrorAction Stop
    }
    catch {
        Write-Host "Konnte ODT-Downloadseite nicht laden: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }

    $pattern = 'https://download\.microsoft\.com/[^\s''"]*officedeploymenttool_[^''"]*\.exe'
    $matches = [regex]::Matches($resp.Content, $pattern)

    if ($matches.Count -eq 0) {
        Write-Host "Keine passende ODT-URL auf der Downloadseite gefunden." -ForegroundColor Yellow
        return $null
    }

    $url = $matches[0].Value
    Write-Host "Gefundene ODT-URL: $url"
    return $url
}

function Ensure-OfficeDeploymentTool {
    param(
        [string]$SetupPath,   # z.B. C:\Users\...\AppData\Local\Temp\...\setup.exe
        [string]$TmpPath,     # z.B. C:\Users\...\AppData\Local\Temp\...\officedeploymenttool.exe
        [string]$FallbackUrl,
        [string]$ExtractDir   # z.B. C:\Users\...\AppData\Local\Temp\...\ (WorkDir)
    )

    if (Test-Path $SetupPath) {
        Write-Host "setup.exe bereits vorhanden: $SetupPath"
        return
    }

    # 1. Versuche dynamische URL
    $odtUrl = Get-ODTUrl
    if (-not $odtUrl) {
        Write-Host "Nutze Fallback-URL für ODT." -ForegroundColor Yellow
        $odtUrl = $FallbackUrl
    }

    Write-Host "Lade Office Deployment Tool von:"
    Write-Host "  $odtUrl"

    try {
        Invoke-WebRequest -Uri $odtUrl -OutFile $TmpPath -ErrorAction Stop
    }
    catch {
        Fail-And-Exit "Download des Office Deployment Tools fehlgeschlagen: $($_.Exception.Message)"
    }

    if (-not (Test-Path $ExtractDir)) {
        New-Item -Path $ExtractDir -ItemType Directory | Out-Null
    }

    Write-Host "Entpacke Office Deployment Tool nach: $ExtractDir"
    & $TmpPath /quiet /extract:$ExtractDir | Out-Null

    # setup.exe rekursiv suchen (ODT legt gern Unterordner an)
    $foundSetup = Get-ChildItem -Path $ExtractDir -Recurse -Filter setup.exe -ErrorAction SilentlyContinue |
                  Select-Object -First 1

    if (-not $foundSetup) {
        Fail-And-Exit "setup.exe wurde nach dem Entpacken nicht gefunden."
    }

    # Nur kopieren, wenn Source und Ziel unterschiedlich sind
    if ($foundSetup.FullName -ne $SetupPath) {
        Copy-Item -Path $foundSetup.FullName -Destination $SetupPath -Force
    }

    if (-not (Test-Path $SetupPath)) {
        Fail-And-Exit "setup.exe konnte nicht an Zielpfad kopiert werden: $SetupPath"
    }

    Write-Host "setup.exe erfolgreich bereitgestellt: $SetupPath"
}

function Get-OfficeClientEdition {
    param([string]$ArchParam)

    switch ($ArchParam) {
        "32"   { return "32" }
        "64"   { return "64" }
        "ARM64" {
            Write-Host "ARM64 gewählt – verwende OfficeClientEdition=64." -ForegroundColor Yellow
            return "64"
        }
        default {
            Write-Host "Unbekannte Architektur '$ArchParam', verwende 64-Bit." -ForegroundColor Yellow
            return "64"
        }
    }
}

function New-DesktopShortcutsFromStartMenu {
    param(
        [string[]]$Apps
    )

    # Hier liegen die Startmenü-Verknüpfungen
    $startMenuRoot = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    # aktueller Benutzer-Desktop
    $desktopPath   = [Environment]::GetFolderPath('Desktop')

    if (-not (Test-Path $desktopPath)) {
        Write-Host "Benutzer-Desktop nicht gefunden – Shortcuts werden übersprungen." -ForegroundColor Yellow
        return
    }

    foreach ($app in $Apps) {
        Write-Host "Suche Startmenü-Shortcut für '$app' ..."

        # „Word“, „Word 2024“, „Microsoft Word“ etc. werden durch das * mit erwischt
        $shortcut = Get-ChildItem -Path $startMenuRoot -Recurse -Filter "$app*.lnk" -ErrorAction SilentlyContinue |
                    Select-Object -First 1

        if (-not $shortcut) {
            Write-Host "  -> Kein Startmenü-Shortcut für '$app' gefunden, überspringe." -ForegroundColor Yellow
            continue
        }

        $dest = Join-Path $desktopPath $shortcut.Name

        try {
            Copy-Item -Path $shortcut.FullName -Destination $dest -Force
            Write-Host "  -> Desktop-Shortcut erstellt: $dest"
        }
        catch {
            Write-Host "  -> Fehler beim Kopieren von '$($shortcut.Name)': $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# ------------------------------------------------------------
# Array-Parameter normalisieren (falls als "A,B,C" übergeben)
# ------------------------------------------------------------
if ($ExcludeApps) {
    if ($ExcludeApps.Count -eq 1) {
        $ExcludeApps = $ExcludeApps[0] -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
}

if ($ShortcutApps) {
    if ($ShortcutApps.Count -eq 1) {
        $ShortcutApps = $ShortcutApps[0] -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
}

# ------------------------------------------------------------
# Blockade bei vorhandener Office-Installation
# ------------------------------------------------------------
$existing = Get-ExistingOffice
if ($existing) {
    Write-Host "Es ist bereits eine Office-Installation vorhanden!" -ForegroundColor Red
    Write-Host "ProductReleaseIds: $($existing.ProductReleaseIds)" -ForegroundColor Yellow
    Write-Host "Bitte zuerst deinstallieren." -ForegroundColor Yellow
    exit 2001
}

# ------------------------------------------------------------
# Basis-Pfade – immer Temp (autarker Installer)
# ------------------------------------------------------------

$UseTemp = $true

$WorkDir = Join-Path $env:TEMP ([guid]::NewGuid())
New-Item -Path $WorkDir -ItemType Directory | Out-Null

$SetupExe   = Join-Path $WorkDir "setup.exe"
$ConfigPath = Join-Path $WorkDir "office_config.xml"
$InstallLog = Join-Path $WorkDir "office_install.log"
$OdtTmp     = Join-Path $WorkDir "officedeploymenttool.exe"

$OdtFallbackUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19628-20046.exe"

# ------------------------------------------------------------
# 1/3 – Konfiguration erstellen
# ------------------------------------------------------------

Write-Host "[1/3] Erstelle Office-Konfiguration ..." -ForegroundColor Cyan

$edition = Get-OfficeClientEdition -ArchParam $Arch

$excludeXml = ""
if ($ExcludeApps -and $ExcludeApps.Count -gt 0) {
    $excludeLines = $ExcludeApps | ForEach-Object {
        '      <ExcludeApp ID="{0}" />' -f $_
    }
    $excludeXml = ($excludeLines -join "`r`n")
}

$configXml = @"
<Configuration>
  <Add OfficeClientEdition="$edition" Channel="$Channel">
    <Product ID="$ProductID">
      <Language ID="$Language" />
$excludeXml
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Updates Enabled="TRUE" Channel="$Channel" />
  <RemoveMSI All="True" />
</Configuration>
"@

$configXml | Out-File -FilePath $ConfigPath -Encoding UTF8 -Force
Write-Host "Konfiguration geschrieben: $ConfigPath"

# ------------------------------------------------------------
# ODT-setup.exe bereitstellen (immer Temp)
# ------------------------------------------------------------

Ensure-OfficeDeploymentTool `
    -SetupPath $SetupExe `
    -TmpPath $OdtTmp `
    -FallbackUrl $OdtFallbackUrl `
    -ExtractDir $WorkDir

# ------------------------------------------------------------
# 2/3 – Office-Installation
# ------------------------------------------------------------

Write-Host ""
Write-Host "[2/3] Starte Office-Installation (dies kann einige Minuten dauern)..." -ForegroundColor Cyan

if (Test-Path $InstallLog) {
    Remove-Item $InstallLog -Force -ErrorAction SilentlyContinue
}

$arguments = "/configure `"$ConfigPath`""

Write-Host "Rufe setup.exe auf:"
Write-Host "  `"$SetupExe`" $arguments"

$process = Start-Process -FilePath $SetupExe `
                         -ArgumentList $arguments `
                         -Wait `
                         -PassThru `
                         -NoNewWindow

$exitCode = $process.ExitCode

if ($exitCode -ne 0) {
    Write-Host "Fehler bei der Installation, ExitCode: $exitCode" -ForegroundColor Red
    Write-Host "Siehe Log (falls vorhanden): $InstallLog"
    exit $exitCode
}

Write-Host "Office-Installation erfolgreich." -ForegroundColor Green

# ------------------------------------------------------------
# 3/3 – Shortcuts anlegen (Startmenü → Desktop)
# ------------------------------------------------------------

if ($ShortcutApps -and $ShortcutApps.Count -gt 0) {
    Write-Host ""
    Write-Host "[3/3] Erstelle Desktop-Shortcuts ..." -ForegroundColor Cyan
    New-DesktopShortcutsFromStartMenu -Apps $ShortcutApps
}
else {
    Write-Host "Keine ShortcutApps angegeben – überspringe Shortcuts."
}

# ------------------------------------------------------------
# Temp aufräumen
# ------------------------------------------------------------

if ($UseTemp -and -not $KeepTemp -and (Test-Path $WorkDir)) {
    Remove-Item -Path $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "Fertig ✅ Office Installation abgeschlossen." -ForegroundColor Green
exit 0
