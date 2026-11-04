# Axi Digital India – INR accounts

INR books for Axi Digital India Private Limited (Intacct entity AUS055), rebuilt from the
USD-base group ledger and reconciled to the audited financial statements.

Live dashboard: https://shyamshahaxi.github.io/Axi-Digital-India/ (opens `reports/inr_books_fy2026-27.html`).

## data/fy2025-26

| File | What it is |
|---|---|
| `usd_tb_fy2025-26.csv` | USD trial balance, FY 2025-26 (from Intacct) |
| `gl/usd_gl_fy2025-26.csv` | USD GL detail, FY 2025-26 (from Intacct, "055 March 26.csv") |
| `config.json` | Treatment rules for the monthly build (same layout as FY 2026-27; starts from nil in Oct 2025) |
| `audited_31mar2026.csv` | Signed balance sheet by caption, for the year-end rec |
| `out/` | Monthly build outputs, as for FY 2026-27 |
| `opening_tb_01apr2026.csv` | Audited balance sheet at 31 Mar 2026 = INR opening TB for FY 2026-27 |
| `pnl_fy2025-26.csv` | Audited P&L FY 2025-26 by note |
| `inr_tb_from_gl_31mar2026.csv` | INR TB rebuilt from the GL (output of `scripts/build_inr_tb.ps1`) |
| `gl_non_inr_lines_excluded.csv` | USD/AUD GL lines dropped as translation artefacts |
| `recon_gl_vs_audited_31mar2026.csv` | GL INR vs audited, by FS caption (output of `scripts/reconcile_to_audited.ps1`) |

Audited figures are in ₹00s in the signed accounts; the CSVs carry full rupees.

FY 2025-26 is also built month by month with the same scripts and page as FY 2026-27
(`reports/inr_books_fy2025-26.html`):
`.\scripts\build_inr_books.ps1 -Config data\fy2025-26\config.json`, then
`.\scripts\build_dashboard_monthly.ps1 -Config data\fy2025-26\config.json`. Its Annual rec tab compares
31 March 2026 with the signed accounts by caption (reasons in `data/mappings/year_end_notes_fy2025-26.csv`),
then line by line. Review items are in `data/mappings/review_items_fy2025-26.csv`. The years shown in the
page's financial year switch are listed in `data/mappings/report_years.json`.

## FY 2026-27: monthly INR books
Folder `data/fy2026-27/`:

| Path | What it is |
|---|---|
| `gl/` | Intacct GL detail exports (USD base, AUS055). Any file name, any period. If two exports cover the same month, the newer file is used for that month. |
| `config.json` | Treatment rules: capital-receipt pattern, USD-billed suppliers, description rules, records to leave out, FD pattern, bank accounts |
| `usd_tb_apr-sep2026.csv` | USD TB; its opening column feeds the bank check |
| `audited_31mar2027.csv` | Add at year end (`caption,audited_inr`) to switch on the year-end rec |
| `out/` | Generated: `inr_gl_lines.csv`, `gl_lines_excluded.csv`, `tb_movements_by_month.csv`, `pnl_by_month.csv`, `bs_by_month.csv`, `checks_by_month.csv`, `usd_only_items.csv`, `source_files.csv`, `year_end_rec.csv` |

Mappings: `data/mappings/schedule3_fy2026-27.csv` (captions, audited opening, accounts) and
`data/mappings/review_items_fy2026-27.csv` (judgements and open points, with the month raised).

### Month-end routine
1. Export the month's GL detail from Intacct (same columns as before) into `data/fy2026-27/gl/`.
2. Run `.\scripts\build_inr_books.ps1`. It prints any failing or "look at" checks.
3. Deal with what the checks raise:
   - **Currency-conversion lines to review**: an INR line in 410300 that no rule matches. Add the supplier to `usdSuppliers`, a `descriptionRules` entry, or extend `capitalPattern` in `config.json`.
   - **Balance sheet accounts mapped**: a new balance sheet account. Add it to `schedule3_fy2026-27.csv`.
   - **USD-only entries**: real entries kept only in USD (accruals, amortisation, disposals). Decide whether an INR entry is needed and log it in the review items.
   - **Bank balances / overdrawn**: compare with the bank statements.
4. Update `review_items_fy2026-27.csv`, then run `.\scripts\build_dashboard_monthly.ps1` to produce `reports/inr_books_fy2026-27.html`.
5. Commit.

### Year end
Save the signed balance sheet as `data/fy2026-27/audited_31mar2027.csv` and rebuild. The Annual rec
tab then compares the 31 Mar 2027 books with the audited figures by caption. The opening position
is already reconciled to the FY 2025-26 accounts (`data/fy2025-26/recon_gl_vs_audited_31mar2026.csv`).
For FY 2027-28, copy `config.json` to the new year folder and point `openingTb`/`schedule3` at the
FY 2026-27 closing.

FY 2026-27 treatments so far:
- Capital receipts (via 410300, or a GJ straight to 300100) go to share capital at INR received.
- The 30 Sep GJ moving the FY26 MCA fee (7,39,000) from equity to expense is left out: it is in the audited opening.
- USD-billed AI tools (Anthropic, Lovable, Kimi) and Fireblocks paid in INR: the INR side, booked to 410300, is moved to 630180 / 660100.
- HDFC fixed-deposit and sweep lines in 115100 are split out as `115100-FD` and shown in cash.
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

# FY 2026-27, monthly (needs the FY 2025-26 INR TB above)
.\scripts\build_inr_books.ps1           # INR lines, TB, BS and P&L by month, checks
.\scripts\build_dashboard_monthly.ps1   # reports/inr_books_fy2026-27.html
```
Balance sheet lines and audited Schedule III figures are in `data/mappings/schedule3.csv`.
