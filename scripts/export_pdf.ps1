param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)][string]$Pdf)
# Open a docx in WPS (KWPS) or MS Word (Word.Application), repaginate, export PDF.
# Prints the physical page count. ASCII-only script (Windows PowerShell 5.1 safe).
$ErrorActionPreference = 'Stop'

function New-OfficeApp {
  try {
    return New-Object -ComObject KWPS.Application
  } catch {
    return New-Object -ComObject Word.Application
  }
}
$app = New-OfficeApp
$app.Visible = $false
try { $app.DisplayAlerts = 0 } catch {}
$doc = $app.Documents.Open($Path, $false, $true)
$doc.Repaginate()
$pages = $doc.ComputeStatistics(2)
try {
  $doc.ExportAsFixedFormat($Pdf, 17)
} catch {
  $doc.SaveAs([ref]$Pdf, [ref]17)
}
$doc.Close($false)
$app.Quit()
Write-Output "PDF pages: $pages"
