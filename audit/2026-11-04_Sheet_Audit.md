# Fund Sheet Audit vs MCA Track and the PMF Reserve (2026-11-04)

Inputs: `PMF_Syndication_Fund_12.xlsx` (sheet, 11/04 5 AM ingest), PMF reserve tracker
(refreshed 10/29), and six MCA Track reports pulled 11/04: advances index, syndicator
invoices, syndicator earnings (per payment), balance report, periodic income, syndicator
payments (empty, as expected for a reserve account). Plus every PMF credit notice in the
mailbox.

Rerun:

```
python3 -I audit/audit_2026-11-04.py SHEET.xlsx RESERVE.xlsx ADVANCES.xls INVOICES.xlsx EARNINGS.xlsx BALANCE.xls audit
```

Deal-by-deal output: `audit/deal_by_deal_2026-11-04.csv` (one row per deal, sorted by size of
difference).

Assumption: "remove my MGMT fees" means the 3.5% internal fee Ozzie and Ezzie pay you. PMF's own
platform fees (4% or 3% in, same out) are real charges against the reserve and stay in.

## 1. Bottom line

| | Amount |
|---|---|
| PMF reserve, last refresh 10/29 | 31,384.66 |
| Collections cleared since that refresh (net of PMF out-fee) | +2,516.76 |
| New invoices, credits or withdrawals since 10/29 | 0.00 |
| Expected reserve, 11/04 evening | 33,901.42 |
| Sheet Available Cash, 11/04 (4 investors) | 34,585.85 |
| Gap, sheet higher | 684.43 |

Every dollar of the 684.43 is explained in section 3. 2.78 is cents rounding.

## 2. PMF side reconciles to the penny

| PMF reserve line | Tracker | Rebuilt from MCA Track | Difference |
|---|---|---|---|
| Deposits | 84,499.76 | 80,000.00 cash + 4,499.76 credits, each matched to a PMF notice | 0.00 |
| Withdrawals | -22,000.00 | 12,000 + 10,000 on 10/7 (Time_Stamps) | 0.00 |
| Invoices | 220,936.51 | Syndicator invoice report grand total, 289 deals | 0.00 |
| Payback | 189,821.41 (10/29) | 192,338.17 cleared through 11/04 | +2,516.76 since refresh |

- The last deal was funded 10/22, so no invoices were added after the 10/29 refresh.
- The 10/28 credit of 989.60 equals the 14 notices dated 10/27 and 10/28 exactly.
- The earnings report lists cleared payments only. On 10/29, about 1,269 of payments dated on or
  before 10/29 were still in flight; they cleared after the refresh, which is why the tracker
  sits below the ledger total for that date. The 851.16 left over in the September audit was the
  same effect (one day of clearings), not a missing deal.

## 3. Bridge: sheet to reserve, deal by deal

| Line | Sheet to PMF | Cause |
|---|---|---|
| Sheet Available Cash | 34,585.85 | |
| Collections in flight | -1,508.81 | Sheet stores MCA's pending total; PMF credits only cleared payments |
| 5674178801 paid off on 11/04 after your 5 AM ingest | +461.59 | Timing: next ingest shows it Closed and paid in full |
| Fee-out charged on reimbursements | +156.45 | Old Payments_View build takes 4% of Override J; PMF credits J at 100% |
| Fee-in at 4% on the 23 deals PMF charged 3% | +124.60 | Old build ignores the 3% era (8/10 to 9/14) |
| Fee-out at 4% on the same 23 deals | +86.26 | Same cause |
| 5674173171 Pipe reimbursement | +66.00 | Sheet J is blank; PMF's books still carry the 66.00 clawback credit |
| 5674193206 broker % blank | -56.70 | New deal (10/22) never set up: invoice commission is 8.1% |
| 5674181370 LRS broker % | -17.50 | K changed from 0.12 to 0.05 since September; invoice says 12% |
| 5674190251 Broyles upfront fee | +0.90 | PMF invoiced 1.80 (2%), not 3%; PMF-side quirk, leave it |
| Cents rounding | +2.78 | Per-row rounding in Payments_View vs per-payment rounding at MCA |
| Expected reserve, 11/04 | 33,901.42 | |

Timing items (in flight, the 11/04 payoff) are not errors. Everything else is.

## 4. Deal-by-deal field check, all 289 deals

| Check against MCA Track | Result |
|---|---|
| Same 289 contract IDs in sheet, advances, invoices, balance report | Pass |
| Funded and payback (whole deal) | Pass, 289 of 289 |
| Funded date (Deal_Locks vs invoice start date) | Pass, 289 of 289 |
| Split % vs PMF syndicator % | Pass, except 5674193206 blank (sheet silently uses 1%; PMF is 1%) |
| Broker % vs PMF's original invoice commission | 2 fail: 5674181370 (0.05, should be 0.12), 5674193206 (blank, should be 0.081) |
| Status | Pass, except 5674178801 (paid off after the 5 AM ingest) |
| Total paid vs MCA pending | Pass, except 5674178801 (same reason) |
| Reimbursements (J) vs PMF credits | Pass, except 5674173171 Pipe (see questions) |
| Platform fee rate | Live code is the OLD build: flat 4% everywhere |
| Investor shares per deal sum to 100% | Pass, 289 of 289 |
| Snapshot H recomputed from Payments_View and Time_Stamps | Pass, each investor to the cent |
| Withdrawals subtract from Available Cash | Pass |

The MCA invoice and balance reports carry the ORIGINAL commission; the advances export carries
the commission after clawbacks. Override K must match the invoice report, and clawbacks come back
through J. All 287 other deals follow that rule.

## 5. Platform fee era is not one window

| Funded | Deals | PMF fee in / out |
|---|---|---|
| through 8/4 | 266 | 4% / 4% |
| 8/10 to 9/14 | 23 | 3% / 3% |
| 9/22 and 10/22 | 2 | 4% / 4% |

The 2026-09-17 Payments_View file has 3% from 8/9 with no end date. Before deploying it, add a
4% row starting any date from 9/15 to 9/22 (all give identical results for existing deals),
otherwise it undercharges 5674191346 and 5674193206.

## 6. Management fee (internal 3.5%) removed

| Investor | Available Cash now | Remove 3.5% transfer | Without mgmt fee | After all fixes (= PMF view) |
|---|---|---|---|---|
| You | 4,236.62 | -2,230.11 | 2,006.51 | 1,799.95 |
| Dad | 14,291.09 | 0.00 | 14,291.09 | 14,078.02 |
| Ozzie | 9,348.44 | +1,858.60 | 11,207.04 | 11,096.53 |
| Ezzie | 6,709.70 | +371.51 | 7,081.21 | 6,926.92 |
| Total | 34,585.85 | 0.00 | 34,585.85 | 33,901.42 |

- The transfer nets to zero: Ozzie and Ezzie paid 2,230.11, you received 2,230.11.
- The last column spreads each deal's difference to investors by that deal's locked share.
- Investors!G2 (MGMT FEES Earned) shows 2,766.48. It ignores the waiver checkbox, so it
  overstates what you actually received by 536.37.
- The waiver checkbox (Override M) is ticked on all 71 deals that have a reimbursement. It
  switches off the 3.5% on the whole deal, not just on the reimbursement: 404.99 from Ozzie and
  131.44 from Ezzie not collected.

## 7. Fix list

Data (Override, no code):
1. 5674181370: K back to 0.12.
2. 5674193206: I = 0.01, K = 0.081.
3. 5674173171: J = 66 to match PMF's books (pending your decision, question 1).

Code (Apps Script, then Deploy > New version):
4. Delete the duplicate `REBUILD_PaymentsView_VALUES_`, paste the 2026-09-17 build with the extra
   4% row from section 5. Fixes 367.31 (the reimbursement fee-out line and both 3% lines).

Basis:
5. Ingest the MCA export (cleared basis) instead of the Advances Report page, or accept that the
   sheet runs ahead of the reserve by whatever is in flight (about 1,500 today).

Housekeeping:
6. Liquid_Allocation rows 10/7 and 10/8 still show cash before the withdrawals. No deal used
   them, but fix them for history (ID 1 minus 12,000, ID 3 minus 10,000 on both rows).
7. Check off the Attention item for cancelled deal 5674182079.

After 1 to 4 and a same-day refresh, the sheet should match PMF to within cents plus whatever is
in flight that day.
