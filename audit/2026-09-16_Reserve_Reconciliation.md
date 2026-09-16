# PMF Reserve vs Fund Sheet - Available Balance Reconciliation

Date: 2026-09-16. Inputs: `PMF_Syndication_Fund_6.xlsx` (export of the live sheet, verified
identical to the live Google Sheet on 9/16) and `Ari Berger - Reserve - PMF Reserve Account
Tracking.xlsx` (PMF's reserve tracker, refreshed 9/15). Rerun anytime with
`python3 audit/reconcile_reserve.py <fund.xlsx> <reserve.xlsx>`.

## Bottom line

| | PMF reserve (9/15) | Fund sheet (9/16) |
|---|---|---|
| Deposits / Contributions | 83,510.16 (80,000 cash + 3,510.16 credits) | 80,000.00 |
| Invoices / Cash Out (Total) | 219,311.81 | 218,413.63 |
| Payback / Cash Received | 164,505.54 | 171,262.41 (includes reimbursements) |
| Available | **28,703.89** | **32,848.78** |
| Gap | | **4,144.89** sheet higher |

The bridge, every line penny-tied:

| Bridge item | Amount | Direction |
|---|---|---|
| Payback on deals (ex reimbursements): sheet 168,396.90 vs PMF 164,505.54 | 3,891.36 | sheet higher, NOT yet explained |
| Invoices: sheet 218,413.63 vs PMF ex-cancelled 218,789.81 | 376.18 | sheet lower, NOT yet explained |
| Reimbursements net: PMF 2,988.16 vs sheet 2,865.51 (J 2,984.91 minus 119.40 fee-out wrongly charged on them) | 122.65 | sheet lower (offsets) |
| **Total** | **4,144.89** | |

The three "definitional" pieces are fully explained. The two operational pieces (3,891.36 of
payback and 376.18 of invoices) need PMF's per-deal ledger to close; see section 4.

## 1. Confirmed sheet bugs (fix these)

1. **Fee-out is charged on reimbursements.** `REBUILD_PaymentsView_VALUES_` adds Override J to
   Paid Back to Investor and then applies the 4% out-fee to the sum. PMF credits reimbursements
   at 100%. Overcharge: **119.40** across 58 deals. Fix: apply fee-out to paid x split only, add J
   after.
2. **The 3% fee era is not live in the deployed sheet.** All 23 deals funded 8/9 onward still
   compute fee-in and fee-out at 4%. Overcharge: fee-in **124.56**, fee-out **15.51**. Either
   `PATCH_REBUILD_PaymentsView_FeeSchedule.gs` was never pasted over the old function, or a
   duplicate `REBUILD_PaymentsView_VALUES_` is loading last (see README hygiene rule). Confirm
   with PMF that the 3% rate is actually what they invoice (the PMF invoice gap of 376.18 is in
   the direction of PMF charging MORE, so do not assume 3% until the per-deal invoice tab says so).
3. **Missing reimbursement**: Override J is blank for `5674176567 (1)` (Collections). PMF credited
   a **3.25** clawback on 2026-03-09. Enter 3.25.
4. **Cancelled deal `5674182079 (1)`** is split 0 in the sheet (correct). PMF invoiced it at
   45,000 x 1% x 1.16 = **522.00** on 4/20 and refunded the same 522.00 inside the 4/23 deposit.
   Net zero on the balance, but it is why PMF's Deposits and Invoices are both ~522 higher than
   the sheet. Not a bug; document it.
5. **Deals_Meta is missing a row** for `5674182079 (1)` (287 deals in Override, 286 in
   Deals_Meta). Harmless today (split 0) but the Attention formula will never flag it.

## 2. Confirmed data points that are NOT bugs

- Every one of the 58 Override J entries matches PMF's credit list: 44 are exactly
  `funded x split x (original broker % - MCA Track's current commission %)` (clawbacks), the
  other 14 are the EPA / management-fee refund deals PMF names in its deposit notes. Nothing in
  J that PMF did not pay; nothing PMF paid that is not in J except the 3.25 above.
- The 58 "commission mismatches" between MCA Track's Total Commission column and Override K
  are post-clawback commissions, not entry errors. Override K holds the ORIGINAL rate, which is
  what PMF invoiced; the clawback comes back through J. Keep K as is.
- Copy of Raw_Deals and Override A:F agree on all 287 deals. The sync is clean.
- The 3.5% management fee nets to zero fund-wide and does not touch the reserve comparison.

## 3. Open per-deal flags

| Deal | Status | Flag | $ |
|---|---|---|---|
| 5674181506 (1) | Open 6 | MCA Track commission dropped 12% -> 5% but PMF has sent no clawback credit | 28.00 owed to fund? ask PMF |
| 5674190804 (1) | Open 2 | MCA Track commission 8.8%, Override K = 8% | sheet under-invoiced 2.40, set K = 0.088 |
| 48 deals at split 2% | mixed | Cannot verify PMF actually allocated 2% without PMF's per-deal invoices. If any are 1% at PMF both invoices and payback move by thousands (2%-deal totals: invoices 23,752.72, payback 18,811.60 at the 1% level) | unknown |

## 4. What is needed to close the remaining 3,891.36 + 376.18

The reserve tracker is an IMPORTRANGE from PMF's internal sheet
(`1vDGFdwb-N0OdyWE0CzODE14EFBEU2xzUDi_fWW2aEDQ`, tabs `syndication invoice`, `syndication
balance`, `deposits`, `withdrawals`, `totals`). Ari Sasson declined to share it with a personal
email (9/10 thread). Needed, in order of usefulness:

1. **`syndication invoice` and `syndication balance` tabs** (a CSV export or a copy shared to
   ari.berger@aspirefundingplatform.com): per-deal invoice amount and per-deal payback. This
   closes everything in one pass: split % per deal, fee rates PMF really applies, and which
   deals' paybacks differ.
2. **MCA Track Syndication report per contract** (portal, "Syndication" or "Investors" tab on
   each deal): confirms the split % PMF booked for each of the 48 2%-split deals.
3. **`Syndication & Reserve Guide.pdf`** (Ari Sasson email, 8/17): PMF's own rules for fee-in,
   fee-out, timing of payback credits (settled vs debited), and reimbursements. Could not be
   downloaded through the mail connector; upload it.

Likely explanations for the 3,891.36 payback gap, to be tested against item 1:
- **Timing**: PMF refreshed 9/15; the sheet's 9/16 ingest is one day newer. Liquid_Allocation's
  9/15 row totals 32,456.86, so at most ~392 of the gap is the extra day.
- **Settled vs debited**: MCA Track's Total Paid counts a debit on its originate date; PMF's
  reserve may credit only settled (cleared) payments. The 8/10 ACH-return email lists fund deals
  `5674178065` and `5674181223` among bounced debits.
- **Split %** differences on 2% deals (above).
- **Fee-out base**: PMF may take its out-fee on gross payback including something the sheet does
  not (or vice versa). The guide PDF answers this.

## 5. Browser access to MCA Track

This session runs in a remote container and cannot use the Chrome extension. To do the
deal-by-deal MCA Track walk-through, either start the session from Claude Code desktop with the
browser extension attached, or export the Advances Report plus each deal's Syndication tab and
upload them here. Everything above was done from the two workbooks, the connected Gmail, and
Drive.
