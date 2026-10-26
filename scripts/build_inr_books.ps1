# Builds the INR books month by month from the Intacct GL exports in the year's gl folder.
# Method (agreed with the auditors for FY 2025-26): INR-currency lines at transaction amount;
# USD/AUD lines dropped (revaluation, realised FX, USD side of payments, AUD side of capital);
# INR side of capital receipts reclassified to share capital. Treatment rules live in config.json.
#
# Monthly routine: put the new month's GL export in the gl folder (any file name; if two exports
# cover the same month the newer file wins), run this script, then build_dashboard_monthly.ps1.
param([string]$Config = "$PSScriptRoot\..\data\fy2026-27\config.json")

$Root = (Resolve-Path "$PSScriptRoot\..").Path
$cfg = Get-Content $Config -Raw -Encoding utf8 | ConvertFrom-Json
function P($rel) { Join-Path $Root $rel }
$out = P $cfg.outFolder; New-Item -ItemType Directory -Force $out | Out-Null
$enc = 'utf8'

# --- source files: newest export wins for any month it covers ----------------------
$byMonth = @{}; $sources = @()
foreach ($f in Get-ChildItem (P $cfg.glFolder) -Filter *.csv | Sort-Object LastWriteTime) {
  $rows = @(Import-Csv $f.FullName)
  $grp = $rows | Group-Object { [datetime]::ParseExact($_.'Batch posting date', 'd/MM/yyyy', $null).ToString('yyyy-MM') }
  foreach ($g in $grp) { $byMonth[$g.Name] = [pscustomobject]@{ file = $f.Name; rows = $g.Group } }
  $sources += [pscustomobject]@{ file = $f.Name; rows = $rows.Count; months = (($grp | ForEach-Object Name | Sort-Object) -join ' '); modified = $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm') }
}
$start = [datetime]::ParseExact($cfg.startMonth, 'yyyy-MM', $null)
$lastKey = ($byMonth.Keys | Sort-Object)[-1]
$months = @(); for ($m = $start; $m.ToString('yyyy-MM') -le $lastKey; $m = $m.AddMonths(1)) { $months += $m.ToString('yyyy-MM') }
foreach ($s in $sources) { $s | Add-Member used_for (($months | Where-Object { $byMonth[$_].file -eq $s.file }) -join ' ') }
$sources | Export-Csv "$out\source_files.csv" -NoTypeInformation -Encoding $enc

# --- classify every line -------------------------------------------------------------
$excludeRec = @{}; foreach ($e in $cfg.excludeRecords) { $excludeRec[$e.record] = $e.reason }
$supplierAcct = @{}; foreach ($s in $cfg.usdSuppliers) { $supplierAcct[$s.supplier] = $s.account }
$names = @{}
$lines = New-Object System.Collections.ArrayList
$excluded = New-Object System.Collections.ArrayList
$usdMv = @{}   # account|month -> USD base movement (all currencies), for the bank check
foreach ($month in $months) {
  if (-not $byMonth.ContainsKey($month)) { continue }
  foreach ($r in $byMonth[$month].rows) {
    $acct = $r.'Account no'; $names[$acct] = $r.'Account title'
    $usdMv["$acct|$month"] = $usdMv["$acct|$month"] + [decimal]$r.'Net Base Amount'
    $desc = (@($r.'Journal entry line description', $r.'Document entry memo', $r.'Document description') | Where-Object { $_ } | Select-Object -First 1)
    $amt = [decimal]$r.'Net Trx Amount'
    if ($r.'Txn currency' -ne 'INR' -or $excludeRec.ContainsKey($r.'Record number')) {
      $reason = if ($excludeRec.ContainsKey($r.'Record number')) { $excludeRec[$r.'Record number'] }
        elseif ($r.Journal -eq 'UFX') { 'Month-end revaluation' }
        elseif ($r.'Record type' -eq 'AP Realized Multi-currency Gain/Loss' -or $desc -match 'Creditors Control Correction|Exchange G/L') { 'Realised FX on AP' }
        elseif ($r.'Txn currency' -eq 'AUD') { 'AUD side of capital receipt' }
        elseif ($acct -eq '410300') { 'USD side of INR payment' }
        else { 'USD-denominated entry' }
      [void]$excluded.Add([pscustomobject]@{ month = $month; posting_date = $r.'Batch posting date'; account_no = $acct; account_name = $r.'Account title'
        currency = $r.'Txn currency'; txn_amount = $amt; usd_amount = [decimal]$r.'Net Base Amount'; journal = $r.Journal
        record_type = $r.'Record type'; supplier = $r.Supplier; description = $desc; record_no = $r.'Record number'; reason = $reason })
      continue
    }
    $treat = ''
    if ($acct -in '410300', '300100') {
      if ($desc -match $cfg.capitalPattern -and $amt -lt 0) { $acct = '300100-INR'; $treat = 'Capital receipt to share capital' }
      elseif ($r.'Record type' -eq 'AP Realized Multi-currency Gain/Loss' -or $desc -match 'Creditors Control Correction') { $acct = '410300-RND'; $treat = 'Rounding' }
      elseif ($supplierAcct.ContainsKey($r.Supplier)) { $acct = $supplierAcct[$r.Supplier]; $treat = 'USD supplier bill paid in INR' }
      else {
        $rule = $cfg.descriptionRules | Where-Object { $desc -match [regex]::Escape($_.match) } | Select-Object -First 1
        if ($rule) { $acct = $rule.account; $treat = $rule.treatment }
        elseif ($acct -eq '410300') { $acct = '410300-OTH'; $treat = 'To review' }
      }
    } elseif ($acct -eq '115100' -and $desc -match $cfg.fdPattern) { $acct = '115100-FD'; $treat = 'HDFC fixed deposit / sweep' }
    if ($amt -eq 0) { continue }
    [void]$lines.Add([pscustomobject]@{ month = $month; posting_date = $r.'Batch posting date'; account_no = $acct; gl_account = $r.'Account no'
      journal = $r.Journal; record_type = $r.'Record type'
      document = (@($r.'Document number', $r.'Journal entry reference no') | Where-Object { $_ } | Select-Object -First 1)
      party = $r.Supplier; description = $desc; amount = $amt; treatment = $treat; record_no = $r.'Record number' })
  }
  # INR lines inside mixed-currency entries do not always net to zero; plug the residue.
  $res = [math]::Round(($lines | Where-Object month -eq $month | Measure-Object amount -Sum).Sum, 2)
  if ($res -ne 0) {
    [void]$lines.Add([pscustomobject]@{ month = $month; posting_date = [datetime]::ParseExact($month, 'yyyy-MM', $null).AddMonths(1).AddDays(-1).ToString('d/MM/yyyy')
      account_no = '410300-RND'; gl_account = ''; journal = 'ADJ'; record_type = ''; document = ''; party = ''
      description = 'INR rounding residue in mixed-currency entries'; amount = -$res; treatment = 'Rounding'; record_no = '' })
  }
}
$names['300100-INR'] = 'Share capital received (INR)'; $names['410300-OTH'] = 'INR lines in Currency conversion (to review)'
$names['410300-RND'] = 'Rounding on INR lines in mixed-currency entries'; $names['115100-FD'] = 'HDFC fixed deposits (sweep)'
$names['310600-GL'] = 'Reserves and surplus (GL basis, b/f)'

$lines | Select-Object month, posting_date, account_no, gl_account, journal, record_type, document, party, description,
  @{ n = 'debit_inr'; e = { if ($_.amount -gt 0) { $_.amount } else { 0 } } }, @{ n = 'credit_inr'; e = { if ($_.amount -lt 0) { -$_.amount } else { 0 } } }, treatment, record_no |
  Export-Csv "$out\inr_gl_lines.csv" -NoTypeInformation -Encoding $enc
$excluded | Export-Csv "$out\gl_lines_excluded.csv" -NoTypeInformation -Encoding $enc

# --- trial balance by month ------------------------------------------------------------
$isPL = { param($a) $a -match '^(4|5|6)' -and $a -ne '410300-IC' }
$open = @{}
foreach ($r in Import-Csv (P $cfg.openingTb)) {
  if (-not $names.ContainsKey($r.account_no)) { $names[$r.account_no] = $r.account_name }
  if (-not (& $isPL $r.account_no)) { $open[$r.account_no] = [decimal]$r.closing_inr }
}
$open['310600-GL'] = -(($open.Values | Measure-Object -Sum).Sum)
$mv = @{}; foreach ($l in $lines) { $k = "$($l.account_no)|$($l.month)"; $mv[$k] = $mv[$k] + $l.amount }
$accts = @($open.Keys) + @($lines | ForEach-Object account_no) | Sort-Object -Unique
$tb = foreach ($a in $accts) {
  $row = [ordered]@{ account_no = $a; account_name = $names[$a]; type = $(if (& $isPL $a) { 'P&L' } else { 'BS' }); opening = $(if (& $isPL $a) { 0 } else { [math]::Round([decimal]$open[$a], 2) }) }
  $run = [decimal]$row.opening; $any = $run -ne 0
  foreach ($m in $months) { $v = [math]::Round([decimal]$mv["$a|$m"], 2); $row[$m] = $v; $run += $v; if ($v -ne 0) { $any = $true } }
  $row['closing'] = [math]::Round($run, 2)
  if ($any) { [pscustomobject]$row }
}
$tb | Export-Csv "$out\tb_movements_by_month.csv" -NoTypeInformation -Encoding $enc

# --- P&L by month ----------------------------------------------------------------------
$groupOf = {
  param($a)
  switch -Regex ($a) {
    '^(420100|440100)$' { 'Other income'; break }
    '^600565$' { 'Personnel costs capitalised'; break }
    '^600160$' { 'Contractors'; break }
    '^60(01|02|03)|^6005(45|50|55)$' { 'Employee benefit expenses'; break }
    '^670' { 'Depreciation and amortisation'; break }
    '^61' { 'Travel and entertainment'; break }
    '^63' { 'Technology and AI tools'; break }
    '^650' { 'Professional and compliance fees'; break }
    '^410300' { 'FX and rounding'; break }
    default { 'Other expenses' }
  }
}
$pl = $tb | Where-Object type -eq 'P&L' | ForEach-Object {
  $row = [ordered]@{ group = (& $groupOf $_.account_no); account_no = $_.account_no; account_name = $_.account_name }
  foreach ($m in $months) { $row[$m] = $_.$m }
  $row['ytd'] = $_.closing
  [pscustomobject]$row
} | Sort-Object group, account_no
$pl | Export-Csv "$out\pnl_by_month.csv" -NoTypeInformation -Encoding $enc
$plMonth = @{}; foreach ($m in $months) { $plMonth[$m] = [decimal](($pl | Measure-Object $m -Sum).Sum) }

# --- balance sheet by month (Schedule III, from the audited opening) ----------------------
$sch = Import-Csv (P $cfg.schedule3)
$capOf = @{}; foreach ($s in $sch) { foreach ($a in ($s.accounts -split ';' | Where-Object { $_ })) { $capOf[$a] = $s.caption } }
$bs = foreach ($s in $sch) {
  $accs = @($s.accounts -split ';' | Where-Object { $_ })
  $sign = if ($s.section -eq 'Assets') { 1 } else { -1 }
  $row = [ordered]@{ caption = $s.caption; section = $s.section; group = $s.group; note = $s.note; opening = [decimal]$s.audited_inr }
  $run = [decimal]$s.audited_inr
  foreach ($m in $months) {
    $v = if ($s.accounts -eq 'PL') { $plMonth[$m] } else { [decimal](($tb | Where-Object { $_.account_no -in $accs } | Measure-Object $m -Sum).Sum) }
    $run += $sign * $v; $row[$m] = [math]::Round($run, 2)
  }
  [pscustomobject]$row
}
$bs | Export-Csv "$out\bs_by_month.csv" -NoTypeInformation -Encoding $enc

# --- USD-only entries with substance ----------------------------------------------------------
$usdOnly = $excluded | Where-Object { $_.reason -eq 'USD-denominated entry' } | Group-Object month, account_no | ForEach-Object {
  [pscustomobject]@{ month = $_.Group[0].month; account_no = $_.Group[0].account_no; account_name = $_.Group[0].account_name; lines = $_.Count
    usd = [math]::Round((($_.Group | Measure-Object usd_amount -Sum).Sum), 2) }
} | Where-Object { [math]::Abs($_.usd) -ge 1 } | Sort-Object month, account_no
$usdOnly | Export-Csv "$out\usd_only_items.csv" -NoTypeInformation -Encoding $enc

# --- monthly checks -----------------------------------------------------------------------------
$usdOpen = @{}; foreach ($r in Import-Csv (P $cfg.openingUsd.file)) { $usdOpen[$r.account_no] = [decimal]$r.($cfg.openingUsd.column) }
$checks = New-Object System.Collections.ArrayList
function Add-Check($m, $name, $status, $value, $detail) { [void]$checks.Add([pscustomobject]@{ month = $m; check = $name; status = $status; value = $value; detail = $detail }) }
$usdRun = @{}; $known = @{}; foreach ($a in $open.Keys) { $known[$a] = $true }
foreach ($m in $months) {
  if (-not $byMonth.ContainsKey($m)) { Add-Check $m 'GL export present' 'bad' 0 'No GL lines for this month'; continue }
  Add-Check $m 'GL export present' 'ok' $byMonth[$m].rows.Count "$($byMonth[$m].rows.Count) lines from $($byMonth[$m].file)"
  $cum = 0; foreach ($t in $tb) { $cum += $t.opening; foreach ($mm in $months) { if ($mm -le $m) { $cum += $t.$mm } } }
  Add-Check $m 'Trial balance nets to nil' $(if ([math]::Abs($cum) -lt 0.01) { 'ok' } else { 'bad' }) ([math]::Round($cum, 2)) 'Sum of all INR closing balances'
  $as = ($bs | Where-Object section -eq 'Assets' | Measure-Object $m -Sum).Sum; $el = ($bs | Where-Object section -ne 'Assets' | Measure-Object $m -Sum).Sum
  Add-Check $m 'Balance sheet balances' $(if ([math]::Abs($as - $el) -lt 1) { 'ok' } else { 'bad' }) ([math]::Round($as - $el, 2)) 'Assets less equity and liabilities'
  $oth = $lines | Where-Object { $_.month -eq $m -and $_.account_no -eq '410300-OTH' }
  Add-Check $m 'Currency-conversion lines to review' $(if ($oth) { 'bad' } else { 'ok' }) ([math]::Round((($oth | Measure-Object amount -Sum).Sum), 2)) "$(@($oth).Count) INR lines in 410300 not matched by a rule"
  $mvAccts = $tb | Where-Object { $_.type -eq 'BS' -and $_.account_no -ne '310600-GL' -and $_.$m -ne 0 -and -not $capOf.ContainsKey($_.account_no) }
  Add-Check $m 'Balance sheet accounts mapped' $(if ($mvAccts) { 'bad' } else { 'ok' }) @($mvAccts).Count $(if ($mvAccts) { 'Unmapped: ' + (($mvAccts | ForEach-Object account_no) -join ', ') } else { 'Every account with movement has a Schedule III caption' })
  $new = @($lines | Where-Object month -eq $m | ForEach-Object account_no | Sort-Object -Unique | Where-Object { -not $known.ContainsKey($_) })
  foreach ($a in $new) { $known[$a] = $true }
  if ($new) { Add-Check $m 'New accounts this month' 'info' $new.Count (($new | ForEach-Object { "$_ $($names[$_])" }) -join '; ') }
  $capL = $lines | Where-Object { $_.month -eq $m -and $_.account_no -eq '300100-INR' }
  if ($capL) { Add-Check $m 'Capital receipts' 'info' ([math]::Round(-(($capL | Measure-Object amount -Sum).Sum), 2)) "$(@($capL).Count) receipt(s) moved to share capital" }
  $u = $usdOnly | Where-Object month -eq $m
  if ($u) { Add-Check $m 'USD-only entries not in INR books' 'warn' @($u).Count (($u | ForEach-Object { "$($_.account_no) USD $('{0:N2}' -f $_.usd)" }) -join '; ') }
  $rnd = [math]::Abs([decimal](($lines | Where-Object { $_.month -eq $m -and $_.account_no -eq '410300-RND' } | Measure-Object amount -Sum).Sum))
  Add-Check $m 'Rounding residue' $(if ($rnd -lt 10) { 'ok' } else { 'warn' }) ([math]::Round($rnd, 2)) 'Plugged to 410300-RND'
  # Bank balances should translate to the USD ledger at one month-end rate.
  $rates = foreach ($a in $cfg.bankAccounts) {
    $usdRun[$a] = [decimal]$usdRun[$a] + [decimal]$usdMv["$a|$m"]          # cumulative USD movement
    $usd = [decimal]$usdOpen[$a] + $usdRun[$a]
    $inr = [decimal]0; foreach ($t in ($tb | Where-Object account_no -eq $a)) { $inr = $t.opening; foreach ($mm in $months) { if ($mm -le $m) { $inr += $t.$mm } } }
    if ($usd -ne 0) { [pscustomobject]@{ a = $a; rate = $inr / $usd; usd = $usd } }
  }
  # Reference rate = the account with the largest USD balance; others are compared in rupees.
  # Small differences arise from INR lines posted after the month-end revaluation ran.
  if ($rates) {
    $ref = ($rates | Sort-Object { [math]::Abs($_.usd) } -Descending | Select-Object -First 1).rate
    $diffs = $rates | ForEach-Object { [pscustomobject]@{ a = $_.a; d = [math]::Round($_.rate * $_.usd - $ref * $_.usd, 0) } }
    $worst = ($diffs | ForEach-Object { [math]::Abs($_.d) } | Measure-Object -Maximum).Maximum
    Add-Check $m 'Bank balances agree with USD ledger' $(if ($worst -lt 1000) { 'ok' } elseif ($worst -lt 10000) { 'warn' } else { 'bad' }) $worst `
      ("Month-end rate $('{0:N2}' -f $ref); differences in INR: " + (($diffs | ForEach-Object { "$($_.a) $('{0:N0}' -f $_.d)" }) -join '; '))
  }
  $neg = foreach ($a in $cfg.bankAccounts) {
    $c = [decimal]0; foreach ($t in ($tb | Where-Object account_no -eq $a)) { $c = $t.opening; foreach ($mm in $months) { if ($mm -le $m) { $c += $t.$mm } } }
    if ($c -lt 0) { "$a $('{0:N2}' -f $c)" }
  }
  Add-Check $m 'No bank account overdrawn in the books' $(if ($neg) { 'warn' } else { 'ok' }) @($neg).Count $(if ($neg) { 'Negative month-end balance (timing of transfers or sweeps vs bank statement): ' + ($neg -join '; ') } else { 'All bank balances positive' })
}
$checks | Export-Csv "$out\checks_by_month.csv" -NoTypeInformation -Encoding $enc

# --- year-end reconciliation to the audited accounts (once they are available) ----------------
$ye = P $cfg.yearEndAudited
if (Test-Path $ye) {
  $aud = @{}; foreach ($r in Import-Csv $ye) { $aud[$r.caption] = [decimal]$r.audited_inr }
  $last = $months[-1]
  $bs | ForEach-Object { [pscustomobject]@{ caption = $_.caption; books_inr = $_.$last; audited_inr = $aud[$_.caption]; difference = [math]::Round($_.$last - [decimal]$aud[$_.caption], 2) } } |
    Export-Csv "$out\year_end_rec.csv" -NoTypeInformation -Encoding $enc
}

# --- summary ------------------------------------------------------------------------------------------
"Months: $($months -join ' ')  INR lines: $($lines.Count)  excluded: $($excluded.Count)"
$checks | Where-Object { $_.status -in 'bad', 'warn' } | ForEach-Object { "  [$($_.status)] $($_.month) $($_.check): $($_.detail)" }
"Result by month (Dr = loss): " + (($months | ForEach-Object { "$_ $('{0:N0}' -f $plMonth[$_])" }) -join ' | ')
