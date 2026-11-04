$ErrorActionPreference = 'Stop'

$docx = 'E:\AIDD\AIDD8122_Final_Project\report\AIDD8122_Final_Report_HuangYong.docx'
$pdf  = 'E:\AIDD\AIDD8122_Final_Project\report\AIDD8122_Final_Report_HuangYong.pdf'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$doc = $word.Documents.Open($docx, $false, $false, $false)

# Remove any pre-existing TOC objects (a rebuilt docx has none)
while ($doc.TablesOfContents.Count -gt 0) {
    $doc.TablesOfContents.Item(1).Delete()
}

# Locate the placeholder paragraph carrying the TOC field
$found = $null
foreach ($p in $doc.Paragraphs) {
    if ($p.Range.Text -match 'Update Field|build the table of contents') {
        $found = $p
        break
    }
}
if ($null -eq $found) {
    foreach ($f in $doc.Fields) {
        if ($f.Code.Text -match 'TOC') {
            $found = $doc.Paragraphs.Item($doc.Range(0, $f.Result.Start).Paragraphs.Count)
            break
        }
    }
}
if ($null -eq $found) { throw 'TOC placeholder paragraph not found' }

$start = $found.Range.Start
$end   = $found.Range.End
$anchor = $doc.Range($start, $start)

$toc = $doc.TablesOfContents.Add($anchor, $false, 1, 2, $false, "", $true, $true, "", $true, $true, $true)
$toc.Update()

# Drop the original placeholder paragraph if it survived below the TOC
$doc.Content.Find.ClearFormatting()
$rng = $doc.Content
$null = $rng.Find.Execute('Right-click and choose Update Field to build the table of contents.')
if ($rng.Find.Found) { $rng.Text = '' }

$doc.Fields.Update() | Out-Null
$doc.Repaginate()
Write-Output ("TOC entries: " + $toc.Range.Paragraphs.Count + " ; pages: " + $doc.ComputeStatistics(2))

$doc.Save()
$doc.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 0, 0, 0, $true, $true, 0, $true, $true, $false)
$doc.Close($false)
$word.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Get-Item $pdf | Select-Object Name, Length, LastWriteTime | Format-List
