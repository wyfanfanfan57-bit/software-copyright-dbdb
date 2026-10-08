param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)][string]$Out)
# Read outline-level 1/2 headings and their adjusted page number (section 2 restarts at 1),
# write heading_pages.json for patch_toc.py. Works with WPS or MS Word. No PDF export.
$ErrorActionPreference = 'Stop'

function New-OfficeApp {
  try { return New-Object -ComObject KWPS.Application } catch { return New-Object -ComObject Word.Application }
}
$app = New-OfficeApp
$app.Visible = $false
try { $app.DisplayAlerts = 0 } catch {}
$doc = $app.Documents.Open($Path, $false, $true)
$doc.Repaginate()
$results = New-Object System.Collections.ArrayList
$n = $doc.Paragraphs.Count
for ($i = 1; $i -le $n; $i++) {
  $p = $doc.Paragraphs.Item($i)
  $ol = $p.OutlineLevel
  if ($ol -eq 1 -or $ol -eq 2) {
    $txt = $p.Range.Text
    $txt = $txt -replace "[`r`n`a`t`f`v]", ""
    $pgAdj = $p.Range.Information(1)
    $pgPhys = $p.Range.Information(3)
    [void]$results.Add([pscustomobject]@{ level = [int]$ol; text = $txt; pageAdj = [int]$pgAdj; pagePhys = [int]$pgPhys })
  }
}
$totalPhys = $doc.ComputeStatistics(2)
$doc.Close($false)
$app.Quit()
[pscustomobject]@{ totalPhys = [int]$totalPhys; headings = $results } | ConvertTo-Json -Depth 6 -Compress | Out-File -Encoding utf8 $Out
Write-Output "OK $($results.Count) headings, $totalPhys pages"
