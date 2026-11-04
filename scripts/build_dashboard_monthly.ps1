# Builds the monthly INR books dashboard from the outputs of build_inr_books.ps1.
param(
  [string]$Config = "$PSScriptRoot\..\data\fy2026-27\config.json",
  [string]$Template = "$PSScriptRoot\..\reports\templates\inr_books_monthly.template.html"
)
$Root = (Resolve-Path "$PSScriptRoot\..").Path
$cfg = Get-Content $Config -Raw -Encoding utf8 | ConvertFrom-Json
function P($rel) { Join-Path $Root $rel }
$o = P $cfg.outFolder

$tbRows = Import-Csv "$o\tb_movements_by_month.csv"
$months = @($tbRows[0].PSObject.Properties.Name | Where-Object { $_ -match '^\d{4}-\d{2}$' })
function Vals($r) { , @($months | ForEach-Object { [double]$r.$_ }) }

$tb = $tbRows | ForEach-Object { [ordered]@{ a = $_.account_no; n = $_.account_name; t = $_.type; o = [double]$_.opening; v = (Vals $_) } }
$pl = Import-Csv "$o\pnl_by_month.csv" | ForEach-Object { [ordered]@{ g = $_.group; a = $_.account_no; n = $_.account_name; v = (Vals $_) } }
$bs = Import-Csv "$o\bs_by_month.csv" | ForEach-Object { [ordered]@{ caption = $_.caption; section = $_.section; group = $_.group; note = $_.note; o = [double]$_.opening; v = (Vals $_) } }
$checks = Import-Csv "$o\checks_by_month.csv" | ForEach-Object { [ordered]@{ m = $_.month; c = $_.check; s = $_.status; v = $_.value; d = $_.detail } }
$review = Import-Csv (P $cfg.reviewItems) | ForEach-Object { [ordered]@{ id = [int]$_.id; m = $_.month; area = $_.area; item = $_.item; amt = [double]$_.amount_inr; s = $_.status; t = $_.treatment } }
$lines = New-Object System.Collections.ArrayList
foreach ($l in Import-Csv "$o\inr_gl_lines.csv") {
  [void]$lines.Add(@($l.month, $l.posting_date, $l.account_no, $l.journal, $l.party, $l.description, [math]::Round([double]$l.debit_inr - [double]$l.credit_inr, 2), $l.treatment))
}
$exRows = Import-Csv "$o\gl_lines_excluded.csv"
$excluded = $exRows | Group-Object month, reason | ForEach-Object { [ordered]@{ m = $_.Group[0].month; r = $_.Group[0].reason; n = $_.Count } }
$usdOnly = Import-Csv "$o\usd_only_items.csv" | ForEach-Object { [ordered]@{ m = $_.month; a = $_.account_no; n = $_.account_name; usd = [double]$_.usd } }
$sources = Import-Csv "$o\source_files.csv" | ForEach-Object { [ordered]@{ f = $_.file; rows = [int]$_.rows; months = $_.months; used = $_.used_for; mod = $_.modified } }
# Line-by-line rec shown under the year-end rec: the opening rec (FY27) or the year's own (FY26).
$recFile = if ($cfg.detailRec) { $cfg.detailRec } else { $cfg.openingRec }
$openRec = if ($recFile) { @(Import-Csv (P $recFile) | ForEach-Object { [ordered]@{ c = $_.fs_caption; g = [double]$_.gl_inr; a = [double]$_.audited_inr; d = [double]$_.difference; s = $_.status; e = $_.explanation } }) } else { $null }
$yeRec = if (Test-Path "$o\year_end_rec.csv") { @(Import-Csv "$o\year_end_rec.csv" | ForEach-Object { [ordered]@{ c = $_.caption; b = [double]$_.books_inr; a = [double]$_.audited_inr; d = [double]$_.difference; s = $_.status; e = $_.explanation } }) } else { $null }

$data = [ordered]@{
  months = $months; tb = @($tb); pl = @($pl); bs = @($bs); checks = @($checks); review = @($review); lines = $lines
  excluded = @($excluded); usdOnly = @($usdOnly); sources = @($sources); openRec = $openRec; yeRec = $yeRec
  meta = [ordered]@{ fy = $cfg.fy; entity = $cfg.entity; yeFile = $cfg.yearEndAudited; reviewFile = $cfg.reviewItems; text = $cfg.text; inrLines = $lines.Count; exLines = $exRows.Count }
}
$json = $data | ConvertTo-Json -Depth 6 -Compress
$outFile = P "reports\inr_books_fy$($cfg.fy).html"
$html = (Get-Content $Template -Raw -Encoding utf8).Replace('/*__DATA__*/null', $json)
$html = $html.Replace('<!--__FY_SWITCH__-->', (& "$PSScriptRoot\fy_switch.ps1" -Current (Split-Path $outFile -Leaf)))
[System.IO.File]::WriteAllText($outFile, $html, (New-Object System.Text.UTF8Encoding $false))
"Wrote $outFile ($([math]::Round($html.Length / 1KB)) KB; months $($months -join ', '))"
