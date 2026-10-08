param([Parameter(Mandatory=$true)][string]$Path)
# Verify the cut code docx WITHOUT exporting PDF (WPS or MS Word):
#  - total physical pages must be FrontPages+BackPages (default check: exactly 60 expected by caller logic)
#  - every page except the last must carry >= 50 displayed line numbers
# Uses wdFirstCharacterLineNumber (=10) per paragraph. Exit 1 on failure.
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
$pageFirst = @{}; $pageLast = @{}; $pageParaCount = @{}
$lineSupported = $true
for ($i = 1; $i -le $M; $i++) {
  $rng = $doc.Paragraphs.Item($i).Range
  $pg = [int]$rng.Information(3)
  if (-not $pageParaCount.ContainsKey($pg)) { $pageParaCount[$pg] = 0 }
  $pageParaCount[$pg] += 1
  $ln = 0
  try { $ln = [int]$rng.Information(10) } catch { $lineSupported = $false }
  if ($ln -lt 0) { $lineSupported = $false }
  if ($lineSupported) {
    if (-not $pageFirst.ContainsKey($pg) -or $ln -lt $pageFirst[$pg]) { $pageFirst[$pg] = $ln }
    if (-not $pageLast.ContainsKey($pg)  -or $ln -gt $pageLast[$pg])  { $pageLast[$pg]  = $ln }
  }
}
Write-Output ("lineNumberInfoSupported=" + $lineSupported)
$fail = 0
for ($p = 1; $p -le $N; $p++) {
  $pc = $pageParaCount[$p]
  if ($lineSupported) {
    $f = $pageFirst[$p]; $l = $pageLast[$p]; $cnt = $l - $f + 1
    Write-Output ("page {0,2} lines {1,4}-{2,4} count={3,3} paras={4}" -f $p,$f,$l,$cnt,$pc)
    if ($cnt -lt 50 -and $p -ne $N) { $fail = 1 }
  } else {
    Write-Output ("page {0,2} paragraphsStarting={1}" -f $p,$pc)
    if ($pc -lt 45 -and $p -ne $N) { $fail = 1 }
  }
}
$doc.Close($false)
$app.Quit()
if ($fail -eq 1) { throw 'a non-last page has fewer than 50 lines' }
Write-Output 'COM VERIFY OK'
