#!/usr/bin/env python3
"""Reconcile the fund sheet's Available Cash against PMF's reserve Available Balance.

Usage:
  python3 audit/reconcile_reserve.py <PMF_Syndication_Fund.xlsx> <PMF_Reserve_Account_Tracking.xlsx>

Prints the bridge between the two numbers and every per-deal input mismatch it can
detect from the workbooks alone. Requires openpyxl.
"""
import datetime
import re
import sys
from collections import defaultdict

import openpyxl

FEE_SWITCH = datetime.datetime(2026, 8, 9)   # first 3% deal funded 8/9 (5674188893)


def platform_rate(funded):
    return 0.03 if (funded and funded >= FEE_SWITCH) else 0.04


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


def main(fund_path, reserve_path):
    fund = load(fund_path)
    res = load(reserve_path)["Reserve Breakdown"]

    # ---- PMF side -------------------------------------------------------------
    pmf_dep, pmf_wd, pmf_inv, pmf_pay, pmf_avail = res[2][1], res[2][4], res[2][6], res[2][7], res[2][8]
    deposits = [(r[0], r[1], r[2]) for r in res[4:] if r[1]]
    cash_in = sum(a for _, a, n in deposits if n and n.strip().lower() == "goach")
    credits = [(d, a, n) for d, a, n in deposits if not (n and n.strip().lower() == "goach")]
    credit_total = sum(a for _, a, _ in credits)
    note_ids = set(re.findall(r"56741\d{5}", " ".join(n for _, _, n in credits if n)))

    # ---- sheet side -----------------------------------------------------------
    ov = {r[1]: r for r in fund["Override"][1:] if r[1]}
    raw = {r[1]: r for r in fund["Copy of Raw_Deals"][1:] if r[1]}
    snap = [r for r in fund["Snapshot"][1:] if r[0]]
    contrib = sum(r[1] for r in snap)
    cash_out = sum(r[3] for r in snap)
    withdrawals = sum(r[4] for r in snap)
    cash_recv = sum(r[5] for r in snap)
    sheet_avail = sum(r[7] for r in snap)

    agg = defaultdict(lambda: defaultdict(float))
    for r in fund["Payments_View"][1:]:
        if not r[0]:
            continue
        a = agg[r[0]]
        a["contrib"] += r[4] or 0
        a["paid"] += r[5] or 0
        a["feein"] += r[12] or 0
        a["feeout"] += r[13] or 0
        a["date"] = r[18]

    reimb_J = sum((o[9] or 0) for o in ov.values())
    reimb_feeout = era_feeout = era_feein = 0.0
    for cid, a in agg.items():
        o = ov[cid]
        s = o[8] or 0
        paid = (o[4] or 0) * s
        if a["date"] and a["date"] >= FEE_SWITCH:
            era_feeout += a["feeout"] - paid * 0.03
            era_feein += a["feein"] - (o[2] or 0) * s * (0.03 + (o[10] or 0))
        else:
            reimb_feeout += a["feeout"] - paid * 0.04

    cancelled = [(cid, o) for cid, o in ov.items() if str(o[0]).lower().startswith("cancel")]
    cancelled_invoice = 0.0
    for cid, o in cancelled:
        r = raw[cid]
        cancelled_invoice += r[6] * 0.01 * (1 + 0.04 + r[7] / r[6])   # PMF invoiced at 1% then refunded

    missing_J = []
    pending_clawback = []
    for cid, o in ov.items():
        r = raw[cid]
        if r[7] is None or not r[6] or o[10] is None or not o[8]:
            continue
        implied = round(r[6] * o[8] * (o[10] - r[7] / r[6]), 2)
        if implied > 0.01 and not o[9]:
            (missing_J if cid[:10] in note_ids else pending_clawback).append((cid, o[0], implied))

    print("PMF reserve (as of %s)" % res[1][11].date())
    print("  deposits %.2f = cash in %.2f + credits %.2f ; invoices %.2f ; payback %.2f ; avail %.2f"
          % (pmf_dep, cash_in, credit_total, pmf_inv, pmf_pay, pmf_avail))
    print("Fund sheet (Snapshot totals)")
    print("  contributions %.2f ; cash out %.2f ; withdrawals %.2f ; cash received %.2f ; avail %.2f"
          % (contrib, cash_out, withdrawals, cash_recv, sheet_avail))
    print("Gap (sheet - PMF): %.2f" % (sheet_avail - pmf_avail))
    print()
    print("Bridge items detectable from the workbooks:")
    print("  cancelled-deal invoice+refund in PMF only (net 0 on balance): %.2f" % cancelled_invoice)
    print("  PMF credits ex cancelled refund: %.2f ; Override J total: %.2f ; missing in J: %.2f"
          % (credit_total - cancelled_invoice, reimb_J, credit_total - cancelled_invoice - reimb_J))
    print("  sheet charges fee-out on reimbursements (4%% era): %.2f" % reimb_feeout)
    print("  sheet still charges 4%% fee-in on 3%%-era deals: %.2f" % era_feein)
    print("  sheet still charges 4%% fee-out on 3%%-era deals: %.2f" % era_feeout)
    sheet_net_reimb = reimb_J - reimb_feeout
    pay_ex_reimb = cash_recv - sheet_net_reimb
    inv_gap = cash_out - (pmf_inv - cancelled_invoice)
    pay_gap = pay_ex_reimb - pmf_pay
    print("  invoices: sheet %.2f vs PMF ex cancelled %.2f -> sheet lower by %.2f"
          % (cash_out, pmf_inv - cancelled_invoice, -inv_gap))
    print("  payback ex reimbursements: sheet %.2f vs PMF %.2f -> sheet higher by %.2f"
          % (pay_ex_reimb, pmf_pay, pay_gap))
    print("  check: %.2f + %.2f - %.2f = %.2f (gap)"
          % (pay_gap, -inv_gap, (credit_total - cancelled_invoice) - sheet_net_reimb,
             pay_gap - inv_gap - ((credit_total - cancelled_invoice) - sheet_net_reimb)))
    print()
    print("Per-deal flags:")
    for cid, st, amt in missing_J:
        print("  reimbursement in PMF notes but Override J blank: %s %s implied %.2f" % (cid, st, amt))
    for cid, st, amt in pending_clawback:
        print("  MCA Track commission dropped but no PMF credit received: %s %s implied %.2f" % (cid, st, amt))
    for cid, o in cancelled:
        print("  cancelled deal (split 0 in sheet, invoiced+refunded at PMF): %s" % cid)
    for cid, o in ov.items():
        r = raw[cid]
        if r[7] and r[6] and o[10] is not None and r[7] / r[6] > o[10] + 0.0005:
            print("  MCA Track commission ABOVE Override K (under-invoiced in sheet): %s raw %.3f K %.3f diff %.2f"
                  % (cid, r[7] / r[6], o[10], r[6] * (o[8] or 0) * (r[7] / r[6] - o[10])))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
