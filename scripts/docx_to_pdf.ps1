<#
.SYNOPSIS
    Export a .docx to .pdf using the local Microsoft Word engine.

.DESCRIPTION
    Formula layout in the course notes must be rendered by Word's own engine to be
    authoritative. Third-party online previews (Tencent Docs / WPS online) have
    incomplete support for OMML m:nor (normal text inside a formula) and render
    body text inside formulas as italic -- that is a renderer limitation, not a
    document defect. Delivering a PDF sidesteps the discrepancy entirely.

    Note: this script drives Word via COM automation and requires Microsoft Word
    to be installed locally.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File docx_to_pdf.ps1 -Src "note.docx" -Dst "note.pdf"

.NOTES
    IMPORTANT (encoding): this file MUST stay pure ASCII. Windows PowerShell 5.1
    reads a BOM-less file using the ANSI code page (GBK on this machine), so any
    UTF-8 CJK comment gets mis-decoded; GBK trail bytes overlap 0x7B/0x7D, which
    injects stray braces and breaks try/catch parsing ("catch: CommandNotFound").

    The local execution policy is Restricted, so a bare `& script.ps1` is
    SILENTLY IGNORED (no error, no execution). Always invoke with
    `powershell -NoProfile -ExecutionPolicy Bypass -File ...`.
#>
param(
    [Parameter(Mandatory = $true)][string]$Src,
    [string]$Dst
)

$ErrorActionPreference = 'Stop'

# Word COM treats a forward slash as a URL separator
# ("F:/AI/..." -> "F:\//AI/..." and then fails to find the file).
# Normalize every path to backslashes first.
$Src = $Src -replace '/', '\'
if ($Dst) { $Dst = $Dst -replace '/', '\' }

if (-not (Test-Path -LiteralPath $Src)) {
    Write-Error "source not found: $Src"
    exit 1
}
if (-not $Dst) {
    $Dst = [System.IO.Path]::ChangeExtension($Src, '.pdf')
}

$srcFull = (Resolve-Path -LiteralPath $Src).Path
$dstFull = [System.IO.Path]::GetFullPath($Dst)
if (Test-Path -LiteralPath $dstFull) {
    # Some volumes (the F: workspace is one) apply an ACL that denies Remove-Item
    # on existing files. That must not abort the run -- Word's SaveAs can still
    # overwrite the target.
    try { Remove-Item -LiteralPath $dstFull -Force -ErrorAction Stop }
    catch { Write-Output "warn: cannot remove existing target, will overwrite: $dstFull" }
}

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    # ReadOnly=$true, AddToRecentFiles=$false
    $doc = $word.Documents.Open($srcFull, $false, $true)
    # 17 = wdFormatPDF
    $doc.SaveAs([ref]$dstFull, [ref]17)
    Write-Output "PDF written: $dstFull"
}
finally {
    if ($doc) { $doc.Close($false) }
    if ($word) { $word.Quit() }
    foreach ($o in @($doc, $word)) {
        if ($o) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    [GC]::Collect()
}
