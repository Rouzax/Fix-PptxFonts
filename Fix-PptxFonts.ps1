<#
.SYNOPSIS
  Replace hard-coded fonts in an unpacked PPTX with theme tokens.

.DESCRIPTION
  - Arial Nova Light -> +mj-lt (Headings / Major Latin)
  - Arial            -> +mn-lt (Body / Minor Latin)

  Optionally map common Arial variants to +mn-lt via -IncludeVariants.
  Only files under the \ppt\ subtree are processed (*.xml, *.rels).
  Writes UTF-8 without BOM.

.PARAMETER Root
  Root folder of the unpacked PPTX (the folder containing 'ppt').

.PARAMETER IncludeTheme
  Also process theme files under \ppt\theme.

.PARAMETER DryRun
  Report changes without writing files.

.PARAMETER NoBackup
  Do not create .bak backups.

.PARAMETER IncludeVariants
  Also map common Arial variants to +mn-lt.

.EXAMPLE
  .\Fix-PptxFonts.ps1 -Root "D:\work\deck_unpacked"

.EXAMPLE
  .\Fix-PptxFonts.ps1 -Root "D:\work\deck_unpacked" -DryRun

.EXAMPLE
  .\Fix-PptxFonts.ps1 -Root . -IncludeVariants
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Root,

    [switch]$IncludeTheme,
    [switch]$DryRun,
    [switch]$NoBackup,
    [switch]$IncludeVariants
)

# Resolve root and verify ppt subtree exists
try {
    $Root = (Resolve-Path -LiteralPath $Root).Path
} catch {
    Write-Host "Root not found: $Root" -ForegroundColor Yellow
    exit 1
}
$pptPath = Join-Path $Root 'ppt'
if (-not (Test-Path -LiteralPath $pptPath)) {
    Write-Host "This folder does not contain a 'ppt' subfolder: $Root" -ForegroundColor Yellow
    exit 1
}

# Gather target files (exclude ppt\theme by default)
$files = Get-ChildItem -Path $pptPath -Recurse -File | Where-Object {
  $_.Extension -in @('.xml', '.rels') -and
  ( $IncludeTheme -or ($_.FullName -notmatch '[\\\/]ppt[\\\/]theme[\\\/]') )
}


if (-not $files) {
    Write-Host "No XML/RELS files found under: $pptPath" -ForegroundColor Yellow
    exit 0
}

# Regex patterns (case-insensitive), safe for any typeface attribute
# 1) Arial Nova Light (several naming variants) -> +mj-lt (Headings)
$reMajor = '(?i)(typeface\s*=\s*["''])(?:Arial\s*Nova\s*Light|ArialNova[-\s]?Light|Arial\s*Nova\s*Lt|ArialNovaLt)(["''])'
$reMajorRepl = '$1+mj-lt$2'

# 2) Arial -> +mn-lt (Body). Match exactly "Arial" (avoid "Arial Narrow", etc.)
$reMinor = '(?i)(typeface\s*=\s*["''])Arial(\s*)(["''])'
$reMinorRepl = '$1+mn-lt$3'

# 3) Optional Arial variants -> +mn-lt (Body)
$reMinorVariants = '(?i)(typeface\s*=\s*["''])(?:ArialMT|ArialPSMT|Arial\s+Unicode\s+MS|Arial\s+Narrow|Arial\s+Black)(["''])'
$reMinorVariantsRepl = '$1+mn-lt$2'

# Always write UTF-8 without BOM (PS5/PS7 compatible)
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$totalFiles = 0
$totalMajor = 0
$totalMinor = 0
$totalMinorV = 0

foreach ($f in $files) {
    try {
        $text = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8
    } catch {
        Write-Host "SKIP (read error): $($f.FullName)" -ForegroundColor Yellow
        continue
    }

    $orig = $text

    # Count and replace for Major (Arial Nova Light -> +mj-lt)
    $mMajor = [regex]::Matches($text, $reMajor).Count
    if ($mMajor -gt 0) {
        $text = [regex]::Replace($text, $reMajor, $reMajorRepl)
    }

    # Count and replace for Minor (Arial -> +mn-lt)
    $mMinor = [regex]::Matches($text, $reMinor).Count
    if ($mMinor -gt 0) {
        $text = [regex]::Replace($text, $reMinor, $reMinorRepl)
    }

    # Optional variants
    $mMinorVar = 0
    if ($IncludeVariants) {
        $mMinorVar = [regex]::Matches($text, $reMinorVariants).Count
        if ($mMinorVar -gt 0) {
            $text = [regex]::Replace($text, $reMinorVariants, $reMinorVariantsRepl)
        }
    }

    if ($text -ne $orig) {
        $totalFiles++
        $totalMajor += $mMajor
        $totalMinor += $mMinor
        $totalMinorV += $mMinorVar

        if (-not $DryRun) {
            if (-not $NoBackup) {
                $bak = "$($f.FullName).bak"
                if (-not (Test-Path -LiteralPath $bak)) {
                    try {
                        Copy-Item -LiteralPath $f.FullName -Destination $bak -Force
                    } catch {
                        Write-Host "WARN: Could not create backup: $bak" -ForegroundColor Yellow
                    }
                    end
                }
            }
            try {
                [System.IO.File]::WriteAllText($f.FullName, $text, $Utf8NoBom)
            } catch {
                Write-Host "ERROR writing file (kept original): $($f.FullName)" -ForegroundColor Red
                continue
            }
        }

        Write-Host $f.FullName
        Write-Host ("  Headings (Arial Nova Light -> +mj-lt): {0}" -f $mMajor)
        Write-Host ("  Body (Arial -> +mn-lt): {0}" -f $mMinor)
        if ($IncludeVariants) {
            Write-Host ("  Body (Arial variants -> +mn-lt): {0}" -f $mMinorVar)
        }
    }
}

Write-Host ("DONE. Files changed: {0} | Headings replacements: {1} | Body replacements: {2}" -f $totalFiles, $totalMajor, $totalMinor)
if ($IncludeVariants) {
    Write-Host ("      Body variant replacements: {0}" -f $totalMinorV)
}
if ($DryRun) {
    Write-Host "Dry run only - no files modified." -ForegroundColor Yellow
}
