# Rebuilds an INR trial balance from the USD-base Intacct GL export.
# INR-currency lines are taken at their transaction amount. USD/AUD lines are
# translation artefacts of the USD base (revaluations, realised FX, USD
# conversions) and are dropped, except where noted below.
param(
  [string]$Gl = "$PSScriptRoot\..\data\fy2025-26\usd_gl_fy2025-26.csv",
  [string]$Out = "$PSScriptRoot\..\data\fy2025-26\inr_tb_from_gl_31mar2026.csv",
  [string]$Lines = "$PSScriptRoot\..\data\fy2025-26\inr_gl_lines_fy2025-26.csv",
  [string]$Excluded = "$PSScriptRoot\..\data\fy2025-26\gl_non_inr_lines_excluded.csv"
)

$rows = Import-Csv $Gl

# INR side of share capital receipts sits in 410300 (Currency conversion).
$capitalPattern = 'INW 151125I049903038|INW 031025I049907391|ADDITIONAL CAPITAL|Less money received'

# Record 179535333 is a USD FX difference (USD 11.06) keyed with an INR txn
# amount of 1,009; it unbalances the INR ledger, so it is treated as FX.
$fxKeyedAsInr = @('179535333')

$inr = foreach ($r in $rows | Where-Object { $_.'Txn currency' -eq 'INR' -and $_.'Record number' -notin $fxKeyedAsInr }) {
  $acct = $r.'Account no'; $name = $r.'Account title'
  if ($acct -eq '410300') {
    if ($r.'Journal entry line description' -match $capitalPattern -or $r.'Document description' -match $capitalPattern) {
      $acct = '300100-INR'; $name = 'Share capital received (INR)'
    } elseif ($r.'Journal entry line description' -match 'Convert intercompany') {
      $acct = '410300-IC'; $name = 'IC conversion entries (net nil)'
    } else {
      $acct = '410300-OTH'; $name = 'INR lines in Currency conversion (to reclassify)'
    }
  }
  $amt = [decimal]$r.'Net Trx Amount'
  [pscustomobject]@{
    Acct = $acct; Name = $name; Amt = $amt
    Date = $r.'Batch posting date'; Journal = $r.Journal; RecordNo = $r.'Record number'
    Document = (@($r.'Document number', $r.'Journal entry reference no') | Where-Object { $_ } | Select-Object -First 1)
    Description = (@($r.'Journal entry line description', $r.'Document entry memo', $r.'Document description') | Where-Object { $_ } | Select-Object -First 1)
    Party = $r.Supplier
  }
}

$inr | Where-Object { $_.Amt -ne 0 } | ForEach-Object {
  [pscustomobject]@{
    posting_date = $_.Date; account_no = $_.Acct; account_name = $_.Name; journal = $_.Journal
    document = $_.Document; party = $_.Party; description = $_.Description
    debit_inr = if ($_.Amt -gt 0) { $_.Amt } else { 0 }; credit_inr = if ($_.Amt -lt 0) { -$_.Amt } else { 0 }
    record_no = $_.RecordNo
  }
} | Export-Csv $Lines -NoTypeInformation -Encoding utf8

$tb = $inr | Group-Object Acct | ForEach-Object {
  [pscustomobject]@{
    account_no = $_.Name
    account_name = $_.Group[0].Name
    inr_lines = $_.Count
    closing_inr = [math]::Round(($_.Group | Measure-Object Amt -Sum).Sum, 2)
  }
} | Sort-Object account_no
$tb | Export-Csv $Out -NoTypeInformation -Encoding utf8

$rows | Where-Object { $_.'Txn currency' -ne 'INR' } |
  Select-Object 'Batch posting date','Account no','Account title','Txn currency','Net Trx Amount','Net Base Amount','Journal','Record type','Journal entry line description','Document description' |
  Export-Csv $Excluded -NoTypeInformation -Encoding utf8

"INR TB net (should be 0): " + [math]::Round(($tb | Measure-Object closing_inr -Sum).Sum, 2)
