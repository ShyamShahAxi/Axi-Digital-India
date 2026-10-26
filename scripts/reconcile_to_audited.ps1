# Compares the INR TB rebuilt from the GL with the audited 31 Mar 2026 figures,
# by financial-statement caption. P&L accounts (5xx/6xx) roll into net loss.
# Differences are kept as unreconciled items with status/explanation from recon_notes.csv.
param(
  [string]$Tb = "$PSScriptRoot\..\data\fy2025-26\inr_tb_from_gl_31mar2026.csv",
  [string]$Map = "$PSScriptRoot\..\data\mappings\gl_to_fs_caption.csv",
  [string]$Notes = "$PSScriptRoot\..\data\mappings\recon_notes.csv",
  [string]$Out = "$PSScriptRoot\..\data\fy2025-26\recon_gl_vs_audited_31mar2026.csv"
)

$auditedLossBeforeTax = 10685869
$captionOf = @{}; $audited = @{}
foreach ($m in Import-Csv $Map) {
  if ($m.account_no -ne '') { $captionOf[$m.account_no] = $m.fs_caption }
  if ($m.audited_inr -ne '') { $audited[$m.fs_caption] = [decimal]$m.audited_inr }
}
$audited['P&L - net loss before tax'] = $auditedLossBeforeTax

$note = @{}
foreach ($n in Import-Csv $Notes) { $note[$n.fs_caption] = $n }

$lines = foreach ($r in Import-Csv $Tb) {
  $cap = if ($r.account_no -match '^[5-6]') { 'P&L - net loss before tax' } else { $captionOf[$r.account_no] }
  if (-not $cap) { $cap = "UNMAPPED $($r.account_no)" }
  [pscustomobject]@{ Caption = $cap; Amt = [decimal]$r.closing_inr }
}
# Captions that exist only in the audited accounts (no GL account behind them).
foreach ($cap in $audited.Keys) {
  if (-not ($lines | Where-Object Caption -eq $cap)) { $lines += [pscustomobject]@{ Caption = $cap; Amt = [decimal]0 } }
}

$recon = $lines | Group-Object Caption | ForEach-Object {
  $gl = ($_.Group | Measure-Object Amt -Sum).Sum
  $aud = if ($audited.ContainsKey($_.Name)) { $audited[$_.Name] } else { 0 }
  $diff = [math]::Round($gl - $aud, 2)
  $n = $note[$_.Name]
  [pscustomobject]@{
    fs_caption = $_.Name
    gl_inr = [math]::Round($gl, 2)
    audited_inr = $aud
    difference = $diff
    status = if ($n) { $n.status } elseif ($diff -eq 0) { 'Reconciled' } else { 'Open' }
    explanation = if ($n) { $n.explanation } else { '' }
  }
} | Sort-Object { [math]::Abs($_.difference) } -Descending
$recon | Export-Csv $Out -NoTypeInformation -Encoding utf8

"Net of all differences (should be 0): " + [math]::Round(($recon | Measure-Object difference -Sum).Sum, 2)
$recon | Select-Object fs_caption, gl_inr, audited_inr, difference, status
