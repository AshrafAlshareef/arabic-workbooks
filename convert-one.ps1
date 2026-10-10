#Requires -Version 5.1
<#
    convert-one.ps1
    Converts a single .docx to .pdf using Microsoft Word COM automation.

    Runs as its own short-lived process: Word's ExportAsFixedFormat can deadlock
    when driven from a long-lived PowerShell session, so docx-to-pdf.ps1 spawns
    one of these per workbook.

    Exit codes:  0 = ok   2 = Word not installed   3 = conversion failed
#>

param(
    [Parameter(Mandatory = $true)] [string] $Docx,
    [Parameter(Mandatory = $true)] [string] $Pdf
)

$ErrorActionPreference = 'Stop'

# wdExportFormatPDF
$wdExportFormatPDF = 17

if (-not (Test-Path -LiteralPath $Docx)) {
    Write-Error "Source not found: $Docx"
    exit 3
}

$pdfDir = Split-Path -Parent $Pdf
if ($pdfDir -and -not (Test-Path -LiteralPath $pdfDir)) {
    New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null
}
if (Test-Path -LiteralPath $Pdf) { Remove-Item -LiteralPath $Pdf -Force }

$word = $null
$doc  = $null

try {
    $word = New-Object -ComObject Word.Application
}
catch {
    Write-Error "Microsoft Word is not available for automation."
    exit 2
}

try {
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.ScreenUpdating = $false

    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ...)
    $doc = $word.Documents.Open($Docx, $false, $true, $false)

    # (OutputFileName, ExportFormat)
    $doc.ExportAsFixedFormat($Pdf, $wdExportFormatPDF)

    $doc.Close(0)
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    $doc = $null
}
catch {
    Write-Error $_.Exception.Message
    if ($doc) {
        try { $doc.Close(0) } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch { }
        $doc = $null
    }
    if ($word) {
        try { $word.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    exit 3
}

try { $word.Quit() } catch { }
try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

if (-not (Test-Path -LiteralPath $Pdf)) {
    Write-Error "Word reported success but no PDF was written."
    exit 3
}

exit 0
