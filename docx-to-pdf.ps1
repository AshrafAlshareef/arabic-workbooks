#Requires -Version 5.1
<#
    docx-to-pdf.ps1
    Converts each workbook's generated .docx into a .pdf sitting next to it.

    Each conversion is spawned as its own hidden powershell.exe process with a
    hard timeout, so a Word deadlock can never block the whole build.

    Exit codes:  0 = all ok   2 = some books failed
#>

$ErrorActionPreference = 'Stop'

$root      = $PSScriptRoot
$converter = Join-Path $root 'convert-one.ps1'
$timeoutSec = 300

$books = @(
    [pscustomobject]@{ Name = 'Arabic_Alphabet_Workbook';      Docx = 'Arabic_Alphabet_Workbook.docx' }
    [pscustomobject]@{ Name = 'Arabic_Harakat_Workbook';      Docx = 'Arabic_Harakat_Workbook.docx' }
    [pscustomobject]@{ Name = 'Arabic_Syllable_Analysis_Workbook'; Docx = 'Arabic_Syllable_Analysis_Workbook.docx' }
    [pscustomobject]@{ Name = 'Arabic_Word_Analysis_Workbook'; Docx = 'Arabic_Word_Analysis_Workbook.docx' }
    [pscustomobject]@{ Name = 'Arabic_Letter_Shapes_Workbook'; Docx = 'Arabic_Letter_Shapes_Workbook.docx' }
    [pscustomobject]@{ Name = 'Arabic_Word_Composition_Workbook'; Docx = 'Arabic_Word_Composition_Workbook.docx' }
    [pscustomobject]@{ Name = 'Arabic_Stories_Workbook';      Docx = 'Arabic_Stories_Workbook.docx' }
)

function Stop-HeadlessWord {
    # Only kill invisible Word instances left behind by a previous failed run.
    # A Word window the user can see is never touched.
    $stale = Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
             Where-Object { $_.MainWindowHandle -eq 0 }
    foreach ($p in $stale) {
        Write-Host ("  [CLEAN] stopping stale background Word process (PID {0})" -f $p.Id)
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    }

    $visible = Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
               Where-Object { $_.MainWindowHandle -ne 0 }
    if ($visible) {
        Write-Host '  [NOTE] Microsoft Word is open. Close it for the most reliable PDF export.' -ForegroundColor Yellow
    }
}

function Remove-StaleLock {
    $locks = Get-ChildItem -LiteralPath $root -Recurse -Filter '~$*.docx' -Force -ErrorAction SilentlyContinue
    foreach ($l in $locks) {
        Remove-Item -LiteralPath $l.FullName -Force -ErrorAction SilentlyContinue
    }
}

if (-not (Test-Path -LiteralPath $converter)) {
    Write-Host "  [ERROR] convert-one.ps1 not found next to this script." -ForegroundColor Red
    exit 1
}

$outLog = Join-Path $env:TEMP ('w2pdf-' + [Guid]::NewGuid().ToString('N') + '.log')
$errLog = "$outLog.err"

try {
    Stop-HeadlessWord
    Remove-StaleLock

    $ok = 0
    $failed = @()

    foreach ($book in $books) {
        $dir  = Join-Path $root $book.Name
        $docx = Join-Path $dir  $book.Docx
        $pdf  = Join-Path $dir  ([System.IO.Path]::GetFileNameWithoutExtension($book.Docx) + '.pdf')

        if (-not (Test-Path -LiteralPath $docx)) {
            Write-Host ("  [SKIP] {0}  ({1} not found)" -f $book.Name, $book.Docx) -ForegroundColor Yellow
            $failed += $book.Name
            continue
        }

        $sw = [Diagnostics.Stopwatch]::StartNew()
        $proc = $null
        $finished = $false
        $startedAt = Get-Date
        try {
            $proc = Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -PassThru `
                -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$converter`"",
                              '-Docx', "`"$docx`"", '-Pdf', "`"$pdf`"" `
                -RedirectStandardOutput $outLog -RedirectStandardError $errLog
            $finished = $proc.WaitForExit($timeoutSec * 1000)
        }
        catch {
            Write-Host ("  [FAIL] {0}  (cannot launch Word: {1})" -f $book.Name, $_.Exception.Message) -ForegroundColor Red
            $failed += $book.Name
            continue
        }

        if (-not $finished) {
            Write-Host ("  [FAIL] {0}  (timed out after {1}s)" -f $book.Name, $timeoutSec) -ForegroundColor Red
            try { $proc.Kill() } catch { }
            Start-Sleep -Milliseconds 500
            Stop-HeadlessWord
            $failed += $book.Name
            continue
        }

        $secs = [math]::Round($sw.Elapsed.TotalSeconds, 1)

        # The parameterless WaitForExit is what makes ExitCode readable.
        $rc = $null
        try { $proc.WaitForExit() } catch { }
        try { $proc.Refresh() } catch { }
        try { $rc = $proc.ExitCode } catch { }

        $err  = ''
        if (Test-Path -LiteralPath $errLog) {
            $raw = Get-Content -LiteralPath $errLog -Raw -ErrorAction SilentlyContinue
            if ($raw) { $err = ([string]$raw).Trim() }
            Remove-Item -LiteralPath $errLog -Force -ErrorAction SilentlyContinue
        }

        # Judge success by the artefact itself, not just the exit code.
        $fresh = $false
        if (Test-Path -LiteralPath $pdf) {
            $fresh = (Get-Item -LiteralPath $pdf).LastWriteTime -ge $startedAt.AddSeconds(-2)
        }

        if ($fresh) {
            $mb = [math]::Round((Get-Item -LiteralPath $pdf).Length / 1MB, 2)
            $note = ''
            if ($null -ne $rc -and $rc -ne 0) { $note = "  (Word reported exit code $rc)" }
            Write-Host ("  [OK]   {0}  ->  {1}  ({2} MB, {3}s){4}" -f $book.Name, [System.IO.Path]::GetFileName($pdf), $mb, $secs, $note)
            $ok++
        }
        else {
            if (-not $err) { $err = "exit code $rc, no PDF produced" }
            Write-Host ("  [FAIL] {0}  ({1})" -f $book.Name, $err) -ForegroundColor Red
            $failed += $book.Name
        }

        # let Word fully exit before the next conversion starts
        Start-Sleep -Milliseconds 800
    }

    Stop-HeadlessWord
    Remove-StaleLock

    Write-Host ''
    if ($failed.Count -eq 0) {
        Write-Host ("  Converted {0}/{1} workbooks." -f $ok, $books.Count)
        exit 0
    }

    Write-Host ("  Converted {0}/{1} workbooks. Failed: {2}" -f $ok, $books.Count, ($failed -join ', ')) -ForegroundColor Yellow
    Write-Host '  The merge step will not run: every book needs a fresh PDF.' -ForegroundColor Yellow
    exit 2
}
finally {
    if (Test-Path -LiteralPath $outLog) { Remove-Item -LiteralPath $outLog -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $errLog) { Remove-Item -LiteralPath $errLog -Force -ErrorAction SilentlyContinue }
}
