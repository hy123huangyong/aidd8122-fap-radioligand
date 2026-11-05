# Build the final report PDF from the python-docx source.
#
# Pipeline (run from the project root):
#   python src/09_build_report.py      # docx, no TOC field populated
#   powershell -File src/build_report_pdf.ps1   # TOC + PDF + image restore
#
# Two things this script exists to handle:
#   * the TOC has to be a real Word field, so it can be updated by the reader;
#   * Word downsamples embedded pictures to ~1364 px on its own save, so the
#     full-resolution figures are restored afterwards by 17_restore_images.py.

$ErrorActionPreference = 'Stop'

$root = 'E:\AIDD\AIDD8122_Final_Project'
$name = 'AIDD8122_Final_Report_HuangYong_P2600222_v2'
$docx = Join-Path $root "report\$name.docx"
$pdf  = Join-Path $root "report\$name.pdf"

if (-not (Test-Path $docx)) { throw "missing $docx - run python src/09_build_report.py first" }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($docx, $false, $false, $false)

    # Remove any pre-existing TOC objects (a rebuilt docx has none)
    while ($doc.TablesOfContents.Count -gt 0) { $doc.TablesOfContents.Item(1).Delete() }

    # Locate the placeholder paragraph carrying the TOC field
    $found = $null
    foreach ($p in $doc.Paragraphs) {
        if ($p.Range.Text -match 'build the table of contents') { $found = $p; break }
    }
    if ($null -eq $found) { throw 'TOC placeholder paragraph not found' }

    $start = $found.Range.Start
    $toc = $doc.TablesOfContents.Add($doc.Range($start, $start), $false, 1, 3,
                                     $false, "", $true, $true, "", $true, $true, $true)
    $toc.Update()

    $rng = $doc.Content
    $rng.Find.ClearFormatting()
    $null = $rng.Find.Execute('Right-click and choose Update Field to build the table of contents.')
    if ($rng.Find.Found) { $rng.Text = '' }

    $doc.Fields.Update() | Out-Null

    # PAGE / NUMPAGES live in header/footer stories; the main-text
    # Fields.Update() does not reach them, so update each story explicitly.
    foreach ($sec in $doc.Sections) {
        foreach ($i in 1..3) {
            try { $null = $sec.Footers.Item($i).Range.Fields.Update() } catch {}
            try { $null = $sec.Headers.Item($i).Range.Fields.Update() } catch {}
        }
    }
    $doc.Repaginate()
    Write-Output ("TOC entries: " + $toc.Range.Paragraphs.Count + " ; pages: " + $doc.ComputeStatistics(2))

    $doc.Save()
    $doc.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 0, 0, 0, $true, $true, 0, $true, $true, $false)
    $doc.Close($false)
}
finally {
    try { $word.Quit() } catch {}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

# Put the 2414-3075 px source figures back into the .docx
$py = (Get-Command python.exe -ErrorAction SilentlyContinue |
       Where-Object { $_.Source -notlike '*WindowsApps*' } |
       Select-Object -First 1).Source
if (-not $py) {
    $py = Join-Path $env:USERPROFILE '.workbuddy\binaries\python\versions\3.13.12\python.exe'
}
if (-not (Test-Path $py)) { throw "python interpreter not found for the image-restore step" }
& $py (Join-Path $root 'src\17_restore_images.py') $docx

Get-Item $pdf | Select-Object Name, Length, LastWriteTime | Format-List
