# Axi Digital India – INR accounts

INR books for Axi Digital India Private Limited (Intacct entity AUS055), rebuilt from the
USD-base group ledger and reconciled to the audited financial statements.

## data/fy2025-26
| File | What it is |
|---|---|
| `usd_tb_fy2025-26.csv` | USD trial balance, FY 2025-26 (from Intacct) |
| `usd_gl_fy2025-26.csv` | USD GL detail, FY 2025-26 (from Intacct, "055 March 26.csv") |
| `opening_tb_01apr2026.csv` | Audited balance sheet at 31 Mar 2026 = INR opening TB for FY 2026-27 |
| `pnl_fy2025-26.csv` | Audited P&L FY 2025-26 by note |
| `inr_tb_from_gl_31mar2026.csv` | INR TB rebuilt from the GL (output of `scripts/build_inr_tb.ps1`) |
| `gl_non_inr_lines_excluded.csv` | USD/AUD GL lines dropped as translation artefacts |
| `recon_gl_vs_audited_31mar2026.csv` | GL INR vs audited, by FS caption (output of `scripts/reconcile_to_audited.ps1`) |

Audited figures are in ₹00s in the signed accounts; the CSVs carry full rupees.

## Method
INR-currency GL lines are taken at their transaction amount. USD/AUD lines (month-end
revaluations, realised FX on AP payments, USD conversions of intercompany balances, the AUD
side of capital receipts) exist only because the ledger is kept in USD, so they are dropped.
The INR side of share capital receipts, booked to 410300, is reclassified to share capital.

Confirmed with the statutory auditors (Kalantri Surana & Associates): revaluation (400300) and
currency conversion (410300) have no place in the INR TB because, with INR as the functional
currency, there is nothing to revalue. Capital contributions carry no FX; share capital is
the INR actually received.

## Rebuild
```powershell
.\scripts\build_inr_tb.ps1          # INR TB + INR GL lines
.\scripts\reconcile_to_audited.ps1  # differences vs audited, with statuses from recon_notes.csv
.\scripts\build_dashboard.ps1       # reports/inr_recon_31mar2026.html from reports/templates/
```
Balance sheet lines and audited Schedule III figures are in `data/mappings/schedule3.csv`.
