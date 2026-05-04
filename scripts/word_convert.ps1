<#
.SYNOPSIS
    Converts DOCX→PDF or PDF→DOCX using Microsoft Word COM automation.

.DESCRIPTION
    Opens a single Word.Application instance in the background, processes
    every job in the manifest sequentially, then quits Word cleanly.
    Progress is reported to stdout so the Node.js parent can stream updates.

    Stdout protocol:
        STARTING:<index>
        PROGRESS:<index>:done
        PROGRESS:<index>:error:<message>

.PARAMETER ManifestPath
    Path to a UTF-8 JSON file: [{input, output, name}, ...]

.PARAMETER Direction
    'docx-to-pdf' or 'pdf-to-docx'
#>
param(
    [Parameter(Mandatory)][string]$ManifestPath,
    [Parameter(Mandatory)][string]$Direction
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ── Load manifest ──────────────────────────────────────────────────────────
$jobs = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 |
        ConvertFrom-Json

if (-not $jobs -or $jobs.Count -eq 0) { exit 0 }

# ── Open Word (single instance for the whole batch) ────────────────────────
$word = $null
try {
    $word = New-Object -ComObject Word.Application

    $word.Visible           = $false  # headless — no window shown
    $word.DisplayAlerts     = 0       # wdAlertsNone     — suppress all dialogs
    $word.AutomationSecurity = 3      # msoAutomationSecurityForceDisable — no macros

    # ── Process each job sequentially ──────────────────────────────────────
    for ($i = 0; $i -lt $jobs.Count; $i++) {
        $job     = $jobs[$i]
        $inPath  = [string]$job.input
        $outPath = [string]$job.output

        Write-Output "STARTING:${i}"
        [Console]::Out.Flush()

        $doc = $null
        try {
            if ($Direction -eq 'docx-to-pdf') {

                # Open document: ConfirmConversions=$false, ReadOnly=$true
                $doc = $word.Documents.Open($inPath, $false, $true)

                # ExportAsFixedFormat — full parameter list for clarity
                $doc.ExportAsFixedFormat(
                    $outPath,   # OutputFileName
                    17,         # ExportFormat       — wdExportFormatPDF
                    $false,     # OpenAfterExport
                    0,          # OptimizeFor        — wdExportOptimizeForPrint (highest quality, eCTD)
                    0,          # Range              — wdExportAllDocument
                    1,          # From (ignored when Range=AllDocument)
                    1,          # To   (ignored when Range=AllDocument)
                    0,          # Item               — wdExportDocumentContent
                    $true,      # IncludeDocProps
                    $true,      # KeepIRM
                    0,          # CreateBookmarks    — wdExportCreateNoBookmarks
                    $true,      # DocStructureTags
                    $true,      # BitmapMissingFonts
                    $false      # UseISO19005_1      — set to $true for strict PDF/A-1b
                )

                $doc.Close(0)   # 0 = wdDoNotSaveChanges
                $doc = $null

            } elseif ($Direction -eq 'pdf-to-docx') {

                # Word 2013+ opens PDFs natively via its built-in PDF import engine.
                # ConfirmConversions=$false suppresses the format-detection dialog.
                # ReadOnly=$false is required so Word can perform the internal conversion.
                $doc = $word.Documents.Open($inPath, $false, $false)

                # SaveAs2 — wdFormatDocumentDefault = 16 (.docx, Office Open XML)
                $doc.SaveAs2($outPath, 16)
                $doc.Close(0)
                $doc = $null
            }

            Write-Output "PROGRESS:${i}:done"

        } catch {
            # Close any document that was left open before the exception
            if ($null -ne $doc) {
                try { $doc.Close(0) } catch { }
                $doc = $null
            }
            # Flatten message to single line so the Node parser receives it intact
            $msg = $_.Exception.Message -replace '[\r\n\t]+', ' ' -replace '\s{2,}', ' '
            Write-Output "PROGRESS:${i}:error:${msg}"
        }

        [Console]::Out.Flush()
    }

} catch {
    # Fatal — Word could not be instantiated or the script itself crashed
    $msg = $_.Exception.Message -replace '[\r\n\t]+', ' '
    Write-Error "FATAL: ${msg}"
    exit 1

} finally {
    # Always release the COM object — prevents orphaned WINWORD.EXE processes
    if ($null -ne $word) {
        try {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        } catch { }
        $word = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}
