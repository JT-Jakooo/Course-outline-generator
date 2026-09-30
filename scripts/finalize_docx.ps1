<#
.SYNOPSIS
    Update all fields (TOC, PAGE, NUMPAGES) in a .docx with Word, then optionally export a PDF.

.DESCRIPTION
    build_note.py inserts a NATIVE Word TOC field and a footer page-number field.
    A field carries no cached result until Word builds it, so a freshly generated
    .docx shows a placeholder line where the TOC should be, and the PDF would show
    that placeholder too. This script opens the file in Word, updates every field
    (main story + tables of contents + header/footer fields), repaginates, saves the
    .docx back so the cached TOC keeps REAL page numbers, and -- in the same pass --
    exports the PDF if -Pdf is given.

    Run this AFTER build_note.py whenever the note has a table of contents or page
    numbers. For a note without either, plain docx_to_pdf.ps1 is enough.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File finalize_docx.ps1 -Src "note.docx" -Pdf "note.pdf"

.EXAMPLE
    # write the finalized docx somewhere else (the F: workspace denies overwriting
    # pre-existing files, so stage it and move it into place afterwards)
    powershell -NoProfile -ExecutionPolicy Bypass -File finalize_docx.ps1 -Src "note.docx" -OutDocx "F:\tmp\note.final.docx" -Pdf "note.pdf"

.NOTES
    IMPORTANT (encoding): keep this file pure ASCII. Windows PowerShell 5.1 reads a
    BOM-less file using the ANSI code page (GBK on this machine), so any UTF-8 CJK
    comment gets mis-decoded; GBK trail bytes overlap 0x7B/0x7D and inject stray
    braces, which breaks try/catch parsing.

    The local execution policy is Restricted, so a bare `& script.ps1` is SILENTLY
    IGNORED. Always invoke with
    `powershell -NoProfile -ExecutionPolicy Bypass -File ...`.

    Word COM treats a forward slash as a URL separator, so every path is normalized
    to backslashes before use.
#>
param(
    [Parameter(Mandatory = $true)][string]$Src,
    [string]$OutDocx,
    [string]$Pdf,
    [switch]$NoSave
)

$ErrorActionPreference = 'Stop'

$Src = $Src -replace '/', '\'
if ($OutDocx) { $OutDocx = $OutDocx -replace '/', '\' }
if ($Pdf) { $Pdf = $Pdf -replace '/', '\' }

if (-not (Test-Path -LiteralPath $Src)) {
    Write-Error "source not found: $Src"
    exit 1
}
$srcFull = (Resolve-Path -LiteralPath $Src).Path
$outFull = $null
if ($OutDocx) { $outFull = [System.IO.Path]::GetFullPath($OutDocx) }
$pdfFull = $null
if ($Pdf) { $pdfFull = [System.IO.Path]::GetFullPath($Pdf) }

function Remove-Target([string]$path) {
    # The F: workspace applies an ACL that denies Remove-Item on existing files.
    # That must not abort the run: Word's SaveAs can still overwrite.
    if ($path -and (Test-Path -LiteralPath $path)) {
        try { Remove-Item -LiteralPath $path -Force -ErrorAction Stop }
        catch { Write-Output "warn: cannot remove existing target, will overwrite: $path" }
    }
}

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    # ReadOnly=$false (we save), AddToRecentFiles=$false
    $doc = $word.Documents.Open($srcFull, $false, $false)

    # 1) main-story fields -- the TOC field lives here
    [void]$doc.Fields.Update()

    # 2) tables of contents, in case Word exposes them separately
    for ($i = 1; $i -le $doc.TablesOfContents.Count; $i++) {
        [void]$doc.TablesOfContents.Item($i).Update()
    }

    # 3) repaginate, then sweep EVERY story (main text, headers, footers, text
    #    boxes). $doc.Fields.Update() only covers the main text story, so the
    #    footer PAGE / NUMPAGES fields must be reached through StoryRanges.
    [void]$doc.Repaginate()
    foreach ($story in $doc.StoryRanges) {
        [void]$story.Fields.Update()
        $next = $story.NextStoryRange
        while ($null -ne $next) {
            [void]$next.Fields.Update()
            $next = $next.NextStoryRange
        }
    }
    [void]$doc.Repaginate()

    # diagnostics: page count + what each footer now reads
    Write-Output ("pages=" + $doc.ComputeStatistics(2))
    foreach ($sec in $doc.Sections) {
        for ($i = 1; $i -le 3; $i++) {
            try {
                $f = $sec.Footers.Item($i)
                if ($f.Exists) {
                    Write-Output ("footer[$i]=" + ($f.Range.Text -replace "[\r\n\a]", ""))
                }
            } catch { }
        }
    }

    # 4) write the .docx back so the cached TOC keeps real page numbers
    if (-not $NoSave) {
        if ($outFull) {
            Remove-Target $outFull
            # 16 = wdFormatDocumentDefault (.docx)
            $doc.SaveAs2([ref]$outFull, [ref]16)
            Write-Output "docx updated: $outFull"
        }
        else {
            $doc.Save()
            Write-Output "docx updated in place: $srcFull"
        }
    }

    # 5) optional PDF export (17 = wdFormatPDF)
    if ($pdfFull) {
        Remove-Target $pdfFull
        $doc.SaveAs([ref]$pdfFull, [ref]17)
        Write-Output "PDF written: $pdfFull"
    }
}
finally {
    if ($doc) { $doc.Close($false) }
    if ($word) { $word.Quit() }
    foreach ($o in @($doc, $word)) {
        if ($o) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
    }
    [GC]::Collect()
}
