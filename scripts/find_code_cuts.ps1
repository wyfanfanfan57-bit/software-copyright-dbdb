param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)][string]$Out, [int]$FrontPages = 30, [int]$BackPages = 30)
# Binary-search cut points in the full source-code docx (WPS or MS Word), no PDF:
#   front_end  = number of paragraphs that exactly fill the first FrontPages pages
#   back_start = 0-based index of the first paragraph on page (N - BackPages + 1)
# Writes cuts.json for build_code.py cut. Requires total pages > FrontPages.
$ErrorActionPreference = 'Stop'

function New-OfficeApp {
  try { return New-Object -ComObject KWPS.Application } catch { return New-Object -ComObject Word.Application }
}
$app = New-OfficeApp
$app.Visible = $false
try { $app.DisplayAlerts = 0 } catch {}
$doc = $app.Documents.Open($Path, $false, $true)
$doc.Repaginate()
$N = $doc.ComputeStatistics(2)
$M = $doc.Paragraphs.Count
Write-Output "pages=$N paragraphs=$M"
if ($N -le $FrontPages) {
  $doc.Close($false); $app.Quit()
  throw "full doc has only $N pages (<= $FrontPages); submit the full document, no cut needed"
}

function Get-Page([int]$k) { return [int]$doc.Paragraphs.Item($k).Range.Information(3) }
function First-On-Page([int]$target) {
  $lo = 1; $hi = $M
  if ((Get-Page $hi) -lt $target) { return -1 }
  while ($lo -lt $hi) {
    $mid = [int][math]::Floor(($lo + $hi) / 2)
    if ((Get-Page $mid) -ge $target) { $hi = $mid } else { $lo = $mid + 1 }
  }
  return $lo
}
$k1 = First-On-Page ($FrontPages + 1)
if ($k1 -lt 0) { throw 'front boundary not found' }
$frontEnd = $k1 - 1
$backFirstPage = $N - $BackPages + 1
$k2 = First-On-Page $backFirstPage
if ($k2 -lt 0) { throw 'back start not found' }
$backStart = $k2 - 1
Write-Output "frontEndCount=$frontEnd (ends page $(Get-Page $frontEnd)); backStart0=$backStart (page $(Get-Page $k2)); last page $(Get-Page $M)"
$obj = [pscustomobject]@{ total_pages=[int]$N; paragraphs=[int]$M; front_end=[int]$frontEnd; back_start=[int]$backStart }
$obj | ConvertTo-Json -Compress | Out-File -Encoding ascii $Out
$doc.Close($false)
$app.Quit()
Write-Output "CUT JSON SAVED"
