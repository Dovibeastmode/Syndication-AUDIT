# PMF Reserve vs Fund Sheet - Available Balance Reconciliation (v2)

Date: 2026-09-16. Inputs: `PMF_Syndication_Fund_6.xlsx` (live sheet export, 9/16),
`Ari Berger - Reserve - PMF Reserve Account Tracking.xlsx` (PMF reserve tracker, refreshed
9/15), `advances_index_20260916190021_11883.xls` (MCA Track export, 9/16 19:00, with PMF's
per-slice "My ..." columns), plus every `Invoice - <contract>` email from MCA Track and
every `Yam ...` notice from PMF's systems accountant in the connected mailbox.

Rerun: `python3 audit/bridge_sheet_to_pmf.py <fund.xlsx> <reserve.xlsx> <mca_export.xls>`.
Deal-by-deal comparison: `audit/deal_by_deal_sheet_vs_mcatrack.csv`.

## Bottom line

| | Amount |
|---|---|
| Sheet Available Cash (4 investors) | 32,848.78 |
| PMF reserve Available Balance (9/15) | 28,703.89 |
| Gap | 4,144.89 |
| Explained by sheet errors and definitional differences (section 1) | 3,293.71 |
| Remaining, all of it in PMF's payback total (section 3) | 851.18 |

PMF's **invoice total ties to the penny** (replica 219,311.80 vs PMF 219,311.81). PMF's
**deposit/credit total ties to the penny** against the 46 Yam emails. The only open number
is 851.16 in PMF's payback total, and the MCA export is one day newer than PMF's refresh.

## 1. The bridge (every line sums; nothing plugged)

| Line | Amount | What it is |
|---|---|---|
| Sheet Available Cash | 32,848.78 | Snapshot!H, sum of 4 investors |
| A. Fee-out charged on reimbursements | +119.40 | Rebuild adds Override J into Paid Back then takes 4% of it. PMF credits reimbursements at 100%. |
| B. Missing reimbursement | +3.25 | ERL 5674176567 clawback credited 3/9, never entered in J |
| C. Three deals with wrong split or commission | -57.69 | See section 2 |
| D. 3%-era deals: fee-in still at 4% | +124.48 | 23 deals funded 8/9 onward. MCA confirms upfront fee 3% on all of them. |
| E. Broker % on 5674190804 | -2.40 | Invoice 326.40 = 300 x 1.088; K says 8% |
| F. Pending vs cleared payments | -3,496.97 | The sheet ingests MCA Track's "Total paid pending" on all 287 deals. PMF credits the reserve only for "Total paid cleared". This is 85% of the gap. |
| G. 3%-era deals: fee-out still at 4% | +10.09 | Same 23 deals |
| H. Cents rounding | +6.13 | Payments_View is written to cents per investor row |
| Corrected sheet | 29,555.07 | |
| PMF replica from the MCA export | 29,555.07 | 83,510.16 - 219,311.80 + 165,356.70 |
| PMF reserve tracker (9/15) | 28,703.89 | |
| Residual | 851.18 | entirely in PMF payback: replica 165,356.70 vs PMF 164,505.54 |

## 2. What to fix in the sheet (each verified against PMF's own numbers)

1. **Ingest cleared, not pending.** `Copy of Raw_Deals` J (Total Paid) currently equals MCA
   Track's "Total paid pending" on every deal. PMF's reserve moves on "Total paid cleared".
   The advances export has both columns; the Advances Report page shows pending. Switch the
   ingest to cleared (or add a pending column and keep both). Worth 3,496.97 today.
2. **Stop charging fee-out on reimbursements.** In `REBUILD_PaymentsView_VALUES_`, apply the
   out-fee to paid x split only, then add Override J. Worth 119.40.
3. **Deploy the 3% fee era.** All 23 deals funded 8/9 onward compute at 4% in and out. PMF
   charges 3% (upfront fee percentage and management fee percentage both 3 in the export).
   Worth 134.57. Check for a duplicate `REBUILD_PaymentsView_VALUES_` loading last.
4. **5674181506 (Open)**: split is 2% at PMF (My funded 800 on 40,000), sheet has 1%. Invoice
   email 4/14 was 840.00 = 800 x 1.05, so broker K is 5%, not 12%. No clawback is pending.
5. **5674183553 (Collections)**: split is 2% at PMF (My funded 160 on 8,000), sheet has 1%. The
   16.00 clawback in J is right (it is 2% x 8,000 x 10%).
6. **5674190804 (Open)**: set K to 0.088 (invoice 326.40).
7. **5674176567 (Collections)**: enter J = 3.25.
8. **5674173171 Pipe (Closed)**: J = 66 is the 2/27 clawback. PMF re-debited the 66 on 7/16
   ("reissuing the commission", deal nearly paid in full) and MCA shows full 12% commission.
   Economically J should be 0. PMF's tracker still shows the 66 as a credit and its invoice
   total does not contain the re-debit, so ask PMF where (or whether) it was booked.
9. **5674182079 (Cancelled)**: split 0 in sheet is right. PMF invoiced 504 + 18 upfront and
   refunded 522 on 4/23 (booked as a deposit). Add the missing Deals_Meta row.
10. **5674174202 Joes's Shop**: cancelled 12/23/2025, invoice 228.90 refunded. Not in the sheet
    or the MCA export. Netted inside PMF's invoice tab (the invoice total ties without it).

## 3. The 851.16 that is left

PMF payback (9/15) 164,505.54 vs cleared payback from the 9/16 export 165,356.70.
Two candidates, both testable:

- **Timing.** PMF refreshed 9/15; the export is 9/16 19:00. 80 deals had activity on 9/15 or
  9/16. Test: refresh the reserve tracker and export MCA Track within the same hour, rerun the
  script. If the residual drops to ~0 it was timing.
- **One deal missing from PMF's balance tab.** Deal 5674179142 (Lowered Payments, 2% split,
  cleared 44,350.00) computes to 851.52 net payback; the residual is 851.16 (0.36 apart, and
  per-deal cents rounding explains only 0.03). If the same-time test still shows ~851, ask PMF
  to confirm 5674179142's payback line in `syndication balance`.
- If PMF booked the 7/16 Pipe re-debit inside payback rather than invoices, the residual is
  785.18 instead and neither candidate fits cleanly; the same-time test settles it.

## 4. What was verified and is NOT a problem

- 287 deals in the sheet, 287 in the MCA export, same set. Funded, payback, status agree on all.
- All 46 Yam credit notices (clawbacks, EPA commission reimbursements, EPA management fee
  refunds, cancellations) reconcile to PMF's deposit lumps to the penny, and to Override J
  except items 7 and 8 above. The Hondo 5674179578 clawback (96.00) was withdrawn by PMF the
  next day and is correctly absent from J.
- MCA Track's "Total Commission" is the post-clawback commission. Override K holds the
  original invoiced rate, which is what PMF charged. Keep K; clawbacks come back through J.
- PMF's per-deal invoice = My funded x (1 + broker %) plus a separate upfront platform fee
  (4%, or 3% from 8/9). Out-fee = management fee percentage x cleared paid. Reimbursements are
  deposits. The reserve balance is deposits - invoices + net cleared payback.
- The 3.5% internal management fee nets to zero fund-wide and never touches the PMF comparison.
