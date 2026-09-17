#!/usr/bin/env python3
"""Penny bridge: fund sheet Available Cash -> PMF reserve Available Balance.

Usage:
  python3 audit/bridge_sheet_to_pmf.py <PMF_Syndication_Fund.xlsx> <Reserve_Tracking.xlsx> <mca_advances_export.xls>

The MCA Track "advances index" export (xls) carries PMF's own per-slice columns
(My funded amount, My commission amount, Upfront fees, Total paid cleared/pending,
Management fee percentage). Those are the truth for what PMF invoices and credits.
Requires openpyxl and xlrd.
"""
import datetime
import sys
from collections import defaultdict

import openpyxl
import xlrd

FEE_SWITCH = datetime.datetime(2026, 8, 9)

# Manual facts established from the Yam emails (asasson@pmfus.com) and invoice emails
# (noreply@mcatrack.com). Update when new notices arrive.
K_FIX = {"5674190804": 0.088,   # invoice 326.40 = 300 x 1.088
         "5674181506": 0.05,    # invoice 840.00 = 800 x 1.05 (2% split)
         "5674182079": 0.12}    # cancelled; invoiced 504 + 18 upfront, refunded 522
CREDIT_FIX = {"5674176567": 3.25,   # Yam clawback 3/9, never entered in Override J
              "5674182079": 522.0}  # cancelled-deal refund, booked by PMF as a deposit


def load_sheet(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


def load_mca(path):
    s = xlrd.open_workbook(path).sheet_by_index(0)
    hdr = s.row_values(0)
    rows = [dict(zip(hdr, s.row_values(i))) for i in range(1, s.nrows)]
    return {str(r["Contract id"]).strip(): r for r in rows if r["Contract id"]}


def main(fund_path, reserve_path, mca_path):
    fund = load_sheet(fund_path)
    res = load_sheet(reserve_path)["Reserve Breakdown"]
    mca = load_mca(mca_path)
    pmf_dep, pmf_inv, pmf_pay, pmf_avail = res[2][1], res[2][6], res[2][7], res[2][8]

    ov = {r[1][:10]: r for r in fund["Override"][1:] if r[1]}
    pv = defaultdict(lambda: defaultdict(float))
    for r in fund["Payments_View"][1:]:
        if r[0]:
            a = pv[r[0][:10]]
            a["contrib"] += r[4] or 0; a["paid"] += r[5] or 0
            a["feein"] += r[12] or 0; a["feeout"] += r[13] or 0
    snap = [r for r in fund["Snapshot"][1:] if r[0]]
    sheet_avail = sum(r[7] for r in snap)

    missing = sorted(set(mca) ^ set(ov))
    if missing:
        print("WARNING deals not in both sources:", missing)

    def sp(m):
        return m["My funded amount"] / m["Advance amount"] if m["Advance amount"] else 0

    def korig(c):
        return K_FIX.get(c, ov[c][10])

    credit = {c: (o[9] or 0) for c, o in ov.items()}
    credit.update(CREDIT_FIX)

    cat = defaultdict(float)
    for c, m in mca.items():
        o, a = ov[c], pv[c]
        s = o[8] or 0; S = sp(m); pct = m["Management fee percentage"] / 100; J = o[9] or 0
        plat = 0.03 if datetime.datetime.strptime(m["Start date"], "%m/%d/%Y") >= FEE_SWITCH else 0.04
        sheet_net = -(a["contrib"] + a["feein"]) + (a["paid"] - a["feeout"])
        pmf_net = (-(m["My funded amount"] * (1 + korig(c)) + m["Upfront fees"])
                   + m["Total paid cleared"] * S * (1 - pct) + credit[c])
        diff = pmf_net - sheet_net
        parts = {"A. sheet charges 4% fee-out on reimbursements (remove)": 0.04 * J,
                 "B. reimbursement credits missing from Override J": CREDIT_FIX.get(c, 0) if c != "5674182079" else 0}
        if c in ("5674181506", "5674183553", "5674182079"):
            parts["C. deals with wrong split/commission in sheet (5674181506, 5674183553, 5674182079)"] = diff - sum(parts.values())
        else:
            parts["D. 3%-era deals: sheet charges 4% fee-in"] = a["feein"] - o[2] * s * (plat + (o[10] or 0))
            parts["E. broker % differs from invoice (5674190804)"] = -(o[2] * s * (K_FIX[c] - o[10])) if c in K_FIX else 0
            parts["F. sheet counts PENDING paid; PMF credits CLEARED only"] = -(m["Total paid pending"] - m["Total paid cleared"]) * s * (1 - plat)
            parts["G. 3%-era deals: sheet charges 4% fee-out"] = m["Total paid cleared"] * s * (0.04 - plat)
        parts["H. per-deal cents rounding in Payments_View"] = diff - sum(parts.values())
        for k, v in parts.items():
            cat[k] += v

    corrected = sheet_avail + sum(cat.values())
    inv = sum(m["My funded amount"] * (1 + korig(c)) + m["Upfront fees"] for c, m in mca.items())
    pay = sum(m["Total paid cleared"] * sp(m) * (1 - m["Management fee percentage"] / 100) for m in mca.values())
    print("Sheet Available Cash                       %12.2f" % sheet_avail)
    for k, v in sorted(cat.items()):
        print("  %10.2f  %s" % (v, k))
    print("Corrected sheet (what PMF's ledger implies) %12.2f" % corrected)
    print("PMF replica: deposits %.2f - invoices %.2f + cleared payback %.2f = %.2f"
          % (sum(credit.values()) + 80000, inv, pay, sum(credit.values()) + 80000 - inv + pay))
    print("PMF reserve tracker: deposits %.2f invoices %.2f payback %.2f avail %.2f (as of %s)"
          % (pmf_dep, pmf_inv, pmf_pay, pmf_avail, res[1][11]))
    print("  invoices: replica - PMF = %.2f" % (inv - pmf_inv))
    print("  payback : replica - PMF = %.2f" % (pay - pmf_pay))
    print("Residual (corrected sheet - PMF avail):    %12.2f" % (corrected - pmf_avail))


if __name__ == "__main__":
    main(*sys.argv[1:4])
