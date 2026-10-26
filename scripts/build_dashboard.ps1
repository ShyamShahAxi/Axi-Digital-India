# Builds reports/inr_recon_31mar2026.html from the template and the CSVs in data/.
# Run build_inr_tb.ps1 and reconcile_to_audited.ps1 first.
param(
  [string]$Root = "$PSScriptRoot\..",
  [string]$Template = "$PSScriptRoot\..\reports\templates\inr_recon.template.html",
  [string]$Out = "$PSScriptRoot\..\reports\inr_recon_31mar2026.html"
)
$d = "$Root\data\fy2025-26"

$recon = Import-Csv "$d\recon_gl_vs_audited_31mar2026.csv" | ForEach-Object {
  [ordered]@{ c = $_.fs_caption; g = [double]$_.gl_inr; a = [double]$_.audited_inr; d = [double]$_.difference; s = $_.status; e = $_.explanation }
}
$tb = Import-Csv "$d\inr_tb_from_gl_31mar2026.csv" | ForEach-Object {
  [ordered]@{ a = $_.account_no; n = $_.account_name; b = [double]$_.closing_inr }
}
$sch3 = Import-Csv "$Root\data\mappings\schedule3.csv" | ForEach-Object {
  [ordered]@{ caption = $_.caption; section = $_.section; group = $_.group; note = $_.note; audited = [double]$_.audited_inr
    accounts = @($_.accounts -split ';' | Where-Object { $_ -and $_ -ne 'PL' }) }
}

# Audited statement of profit and loss (signed FS, in rupees; loss shown negative).
$pnl = @(
  [ordered]@{ key = 'rev'; label = 'Revenue from operations'; note = ''; audited = 0 }
  [ordered]@{ key = 'oi'; label = 'Other income'; note = ''; audited = 0 }
  [ordered]@{ key = 'emp'; label = 'Employee benefit expenses'; note = '14'; audited = 2281566 }
  [ordered]@{ key = 'dep'; label = 'Depreciation and amortisation'; note = ''; audited = 0 }
  [ordered]@{ key = 'oth'; label = 'Other expenses'; note = '15'; audited = 8404303 }
  [ordered]@{ key = 'texp'; label = 'Total expenses'; note = ''; audited = 10685869; style = 'sub' }
  [ordered]@{ key = 'pbt'; label = 'Loss before tax'; note = ''; audited = -10685869; style = 'sub' }
  [ordered]@{ key = 'dt'; label = 'Deferred tax credit'; note = '16'; audited = 2064719 }
  [ordered]@{ key = 'loss'; label = 'Loss for the year'; note = ''; audited = -8621150; style = 'tot' }
)
$plNotes = Import-Csv "$d\pnl_fy2025-26.csv" | Where-Object { $_.note -in '14', '15' } | ForEach-Object {
  [ordered]@{ account = $_.account; note = $_.note; amount = [double]$_.amount_inr }
}

$lineRows = Import-Csv "$d\inr_gl_lines_fy2025-26.csv"
$lines = New-Object System.Collections.ArrayList
foreach ($l in $lineRows) {
  [void]$lines.Add(@($l.posting_date, $l.account_no, $l.journal, $l.document, $l.party, $l.description, [double]$l.debit_inr, [double]$l.credit_inr, $l.record_no))
}
$dates = $lineRows | ForEach-Object { [datetime]::ParseExact($_.posting_date, 'dd/MM/yyyy', $null) } | Sort-Object
$excluded = @(Import-Csv "$d\gl_non_inr_lines_excluded.csv").Count

$data = [ordered]@{
  recon = @($recon); tb = @($tb); sch3 = @($sch3); pnl = $pnl; plNotes = @($plNotes); lines = $lines
  meta = [ordered]@{ eps = '0.94'; glFrom = $dates[0].ToString('d MMM yyyy'); glTo = $dates[-1].ToString('d MMM yyyy'); excluded = $excluded }
}
$json = $data | ConvertTo-Json -Depth 6 -Compress
$html = (Get-Content $Template -Raw -Encoding utf8).Replace('/*__DATA__*/null', $json)
[System.IO.File]::WriteAllText((Resolve-Path (Split-Path $Out)).Path + '\' + (Split-Path $Out -Leaf), $html, (New-Object System.Text.UTF8Encoding $false))
"Wrote $Out ($([math]::Round($html.Length / 1KB)) KB, $($lines.Count) GL lines)"
