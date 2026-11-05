#!/usr/bin/env python3
"""Deal-by-deal audit: fund sheet vs MCA Track vs PMF reserve (2026-11-04 data).

Usage (all arguments are file paths):
  python3 -I audit/audit_2026-11-04.py SHEET.xlsx RESERVE.xlsx ADVANCES.xls INVOICES.xlsx \
      EARNINGS.xlsx BALANCE.xls OUT_DIR

What it does
  1. Rebuilds PMF's reserve balance from MCA Track's own reports:
       deposits (cash + credits) - withdrawals - invoices + cleared net payback.
     Invoices come from the syndicator invoice report (ties to the reserve tracker),
     payback from the per-payment earnings ledger (cleared payments, net of PMF's out-fee).
  2. Compares every deal in the sheet against MCA Track field by field
     (status, funded, payback, split, broker %, platform fee rates, paid basis,
     reimbursements, funded date).
  3. Bridges the sheet's Available Cash to the reserve, deal by deal, with each
     dollar of difference attributed to a cause.
  4. Shows the 3.5% internal management transfer separately: it nets to zero
     fund-wide, so the fund-level comparison is unaffected by it.

Writes OUT_DIR/deal_by_deal_2026-11-04.csv and prints the summary.
Requires openpyxl and xlrd.
"""
import collections
import csv
import datetime
import os
import sys

import openpyxl
import xlrd

D = datetime.datetime
R2 = lambda x: round(x + 0.0, 2)

# ---------------------------------------------------------------------------
# PMF credits per deal (reimbursements), from every "Yam ..." notice sent by
# PMF's systems accountant, cross-checked against the reserve tracker's deposit
# lumps (sum of credits + 80,000 cash = 84,499.76). The cancelled-deal refund
# (5674182079, 522.00) is listed separately because the sheet models a
# cancellation as split 0, not as a reimbursement.
# ---------------------------------------------------------------------------
CREDITS = {
    # through 2026-08-10 (verified in the 2026-09-16 audit)
    "5674173375": 8.5, "5674174559": 73.72, "5674174009": 12.0, "5674178352": 10.0,
    "5674177642": 24.0, "5674172505": 144.0, "5674174307": 42.0, "5674173025": 75.11,
    "5674174432": 40.8, "5674177881": 132.0, "5674173051": 7.0, "5674173493": 82.6,
    "5674178448": 99.0, "5674176512": 9.71, "5674177762": 12.0, "5674172686": 99.0,
    "5674178065": 24.0, "5674173171": 66.0, "5674173980": 48.0, "5674176486": 21.0,
    "5674174279": 6.0, "5674178798": 11.69, "5674179004": 40.0, "5674179003": 60.0,
    "5674179831": 5.0, "5674179828": 5.0, "5674180215": 35.0, "5674180601": 12.36,
    "5674180841": 180.0, "5674181127": 16.25, "5674181223": 348.0, "5674181151": 9.0,
    "5674181370": 30.0, "5674181674": 22.5, "5674181952": 18.25, "5674182772": 175.0,
    "5674182967": 9.0, "5674183553": 16.0, "5674183456": 20.0, "5674183607": 52.0,
    "5674183661": 60.0, "5674184096": 50.0, "5674184420": 175.0, "5674184556": 6.0,
    "5674184531": 16.8, "5674184570": 4.2, "5674185287": 25.0, "5674185249": 8.0,
    "5674185559": 10.5, "5674185610": 12.8, "5674185848": 105.0, "5674185804": 9.0,
    "5674185922": 36.0, "5674185921": 60.0, "5674186082": 6.0, "5674186104": 89.12,
    "5674186245": 45.0, "5674186369": 15.0, "5674186355": 150.0,
    "5674176567": 3.25,                      # ERL clawback, 2026-03-09
    # 2026-10-28 deposit of 989.60 (14 notices dated 10/27-10/28)
    "5674186083": 50.0, "5674186485": 25.2, "5674190223": 10.5, "5674188393": 7.5,
    "5674187666": 20.0, "5674190918": 312.0, "5674190251": 3.0, "5674188469": 72.0,
    "5674187438": 50.76, "5674189436": 324.14, "5674190031": 9.9, "5674188350": 104.6,
}
CANCEL_REFUND = {"5674182079": 522.0}        # SECURITY SUBCONTRACTING, 2026-04-23

STATUS_OPEN = ("Open", "Lowered")
STATUS_CLOSED = ("Closed", "EPA", "LOC", "Overpaid")


def bucket(s):
    s = str(s or "")
    if s.startswith(STATUS_OPEN):
        return "open"
    if s.startswith(STATUS_CLOSED):
        return "closed"
    return "default"


def cid10(v):
    s = str(v or "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s[:10] if len(s) >= 10 and s[:10].isdigit() else ""


def xlsx(path, sheet=None):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    return [list(r) for r in ws.iter_rows(values_only=True)]


def table(rows):
    hdr = rows[0]
    return [dict(zip(hdr, r)) for r in rows[1:] if any(v not in (None, "") for v in r)]


def xls(path):
    s = xlrd.open_workbook(path).sheet_by_index(0)
    hdr = s.row_values(0)
    return [dict(zip(hdr, s.row_values(i))) for i in range(1, s.nrows)]


def num(v):
    if v in (None, ""):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace("$", "").replace(",", "").replace("%", "").strip())
    except ValueError:
        return 0.0


def main(sheet_p, reserve_p, adv_p, inv_p, earn_p, bal_p, out_dir):
    # ---------------- PMF / MCA Track side ----------------
    res = xlsx(reserve_p)
    pmf = {"deposits": num(res[2][1]), "withdrawals": num(res[2][4]),
           "invoices": num(res[2][6]), "payback": num(res[2][7]),
           "available": num(res[2][8]), "asof": res[1][11]}

    adv = {cid10(r["Contract id"]): r for r in xls(adv_p) if cid10(r["Contract id"])}
    inv = {cid10(r["Contract ID"]): r for r in table(xlsx(inv_p)) if cid10(r["Contract ID"])}
    bal = {cid10(r["Contract ID"]): r for r in xls(bal_p) if cid10(r["Contract ID"])}
    led = collections.defaultdict(lambda: {"gross": 0.0, "fee": 0.0, "net": 0.0, "n": 0,
                                           "pct": set(), "last": None})
    led_by_day = collections.defaultdict(float)
    for r in table(xlsx(earn_p)):
        c = cid10(r["Contract ID"])
        if not c:
            continue
        L = led[c]
        L["gross"] += num(r["Gross Payment"]); L["fee"] += num(r["Management Fee"])
        L["net"] += num(r["Net Payment"]); L["n"] += 1
        L["pct"].add(num(r["Management Fee Percentage"]))
        d = r["Merchant Payment Date"]
        L["last"] = d if L["last"] is None or d > L["last"] else L["last"]
        led_by_day[d.date()] += num(r["Net Payment"])

    inv_total = sum(num(r["Total Amount"]) for r in inv.values())
    led_net = sum(L["net"] for L in led.values())
    credits_total = sum(CREDITS.values()) + sum(CANCEL_REFUND.values())
    cash_in = pmf["deposits"] - credits_total

    # ---------------- sheet side ----------------
    wb = openpyxl.load_workbook(sheet_p, data_only=True, read_only=True)
    rows = lambda name: [list(r) for r in wb[name].iter_rows(values_only=True)]
    ov = {cid10(r[1]): r for r in rows("Override")[1:] if cid10(r[1])}
    raw = {cid10(r[1]): r for r in rows("Copy of Raw_Deals")[1:] if cid10(r[1])}
    meta = {cid10(r[1]): r for r in rows("Deals_Meta")[1:] if cid10(r[1])}
    locks = {cid10(r[0]): r for r in rows("Deal_Locks")[1:] if cid10(r[0])}
    snap = [r for r in rows("Snapshot")[1:] if r[0]]
    invs = [r for r in rows("Investors")[1:] if r[0] and str(r[0]).strip().startswith("ID")]
    ts = [r for r in rows("Time_Stamps")[1:] if r[0]]
    pvr = [r for r in rows("Payments_View")[1:] if cid10(r[0])]

    pv = collections.defaultdict(lambda: {"E": 0.0, "F": 0.0, "M": 0.0, "N": 0.0, "date": None,
                                          "inv": {}})
    for r in pvr:
        c = cid10(r[0]); P = pv[c]
        P["E"] += num(r[4]); P["F"] += num(r[5]); P["M"] += num(r[12]); P["N"] += num(r[13])
        P["date"] = r[18]
        P["inv"][int(num(r[2]))] = {"pct": num(r[3]), "E": num(r[4]), "F": num(r[5]),
                                    "M": num(r[12]), "N": num(r[13])}

    sheet_avail = sum(num(r[7]) for r in snap)
    contrib = sum(num(r[1]) for r in ts); wd = sum(num(r[2]) for r in ts)

    # ---------------- deal sets ----------------
    S_sheet, S_adv, S_inv, S_bal = set(ov), set(adv), set(inv), set(bal)
    print("== DEAL SETS")
    print(f"  sheet {len(S_sheet)}  advances {len(S_adv)}  invoices {len(S_inv)}  balance {len(S_bal)}  ledger {len(led)}")
    for a, b, n in [(S_sheet, S_adv, "sheet-only vs advances"), (S_adv, S_sheet, "advances-only vs sheet"),
                    (S_adv, S_inv, "advances not invoiced"), (S_inv, S_adv, "invoiced not in advances")]:
        if a - b:
            print(f"  {n}: {sorted(a - b)}")

    # ---------------- PMF rebuild ----------------
    pmf_1104 = cash_in + credits_total + pmf["withdrawals"] - inv_total + led_net
    print("\n== PMF RESERVE REBUILT FROM MCA TRACK")
    print(f"  tracker (refreshed {pmf['asof']:%m/%d}): deposits {pmf['deposits']:,.2f}  withdrawals {pmf['withdrawals']:,.2f}"
          f"  invoices {pmf['invoices']:,.2f}  payback {pmf['payback']:,.2f}  available {pmf['available']:,.2f}")
    print(f"  deposits = cash {cash_in:,.2f} + credits {credits_total:,.2f} (emails)")
    print(f"  invoice report total {inv_total:,.2f}  (tracker {pmf['invoices']:,.2f}, diff {inv_total - pmf['invoices']:+.2f})")
    print(f"  cleared net payback, ledger through 11/04: {led_net:,.2f}  (tracker 10/29 {pmf['payback']:,.2f},"
          f" cleared since refresh {led_net - pmf['payback']:+,.2f})")
    print(f"  expected reserve at 11/04 = {pmf_1104:,.2f}")

    # ---------------- per deal ----------------
    fee_in_live = []      # detect which Payments_View build is running
    out = []
    cat = collections.defaultdict(float)
    flags = collections.defaultdict(list)
    for c in sorted(S_sheet | S_adv):
        o, a, i, b, L, P = ov.get(c), adv.get(c), inv.get(c), bal.get(c), led.get(c), pv.get(c)
        rw, lk = raw.get(c), locks.get(c)
        if not (o and a and i):
            continue
        split_blank = o[8] in (None, "")
        split_s = 0.01 if split_blank else num(o[8])     # Payments_View treats a blank split as 1%
        split_m = num(b["Syndicator Percentage"]) / 100 if b else None
        if split_blank:
            flags["split blank (sheet silently uses 1%)"].append((c, o[0], split_m))
        k_s = None if o[10] in (None, "") else num(o[10])
        k_m = num(i["Commission"]) / num(i["Funded"]) if num(i["Funded"]) else None
        up_pct = num(i["Mgmt Fee"]) / 100
        out_pct = sorted(L["pct"])[0] / 100 if L and L["pct"] else num(a["Management fee percentage"]) / 100
        J_s = num(o[9]); J_t = CREDITS.get(c, 0.0)
        whole_paid_s = num(o[4])
        pend, clr = num(a["Total paid pending"]), num(a["Total paid cleared"])
        sd_m = D.strptime(i["Start Date"], "%m/%d/%Y")
        sd_s = lk[1] if lk and isinstance(lk[1], D) else (rw[4] if rw and isinstance(rw[4], D) else None)

        # field checks
        if bucket(o[0]) != bucket(a["Status"]) or not str(o[0]).startswith(str(a["Status"]).split()[0]):
            flags["status"].append((c, o[0], a["Status"]))
        if abs(num(o[2]) - num(a["Advance amount"])) > 0.005 or abs(num(o[3]) - num(a["Payback"])) > 0.005:
            flags["funded/payback"].append((c, num(o[2]), num(a["Advance amount"]), num(o[3]), num(a["Payback"])))
        if split_m is not None and abs(split_s - split_m) > 1e-9 and c not in CANCEL_REFUND:
            flags["split"].append((c, o[0], split_s, split_m))
        if k_m is not None and (k_s is None or abs(k_s - k_m) > 0.00005) and c not in CANCEL_REFUND:
            flags["broker %"].append((c, o[0], k_s, round(k_m, 4)))
        if abs(J_s - J_t) > 0.005:
            flags["reimbursement J"].append((c, o[0], J_s, J_t))
        if sd_s and abs((sd_s.date() - sd_m.date()).days) > 0:
            flags["funded date"].append((c, sd_s.date(), sd_m.date()))
        if abs(whole_paid_s - pend) > 0.005:
            flags["paid vs pending (whole deal)"].append((c, o[0], whole_paid_s, pend, clr))

        # money: sheet fund-level vs PMF
        sE, sF, sM, sN = (P["E"], P["F"], P["M"], P["N"]) if P else (0, 0, 0, 0)
        sheet_net = -sE - sM + sF - sN
        pmf_net = (-num(i["Total Amount"]) + (L["net"] if L else 0.0) + J_t
                   + CANCEL_REFUND.get(c, 0.0))
        diff = sheet_net - pmf_net

        # attribution, fund level (the 3.5% mgmt transfer nets to zero inside each deal)
        #   diff = -(principal gap) - (fee-in gap) + (paid gap) + (J gap) - (fee-out gap)
        inv_F, inv_C, inv_U = num(i["Funded"]), num(i["Commission"]), num(i["U M Fee"])
        led_g = L["gross"] if L else 0.0
        led_fee = L["fee"] if L else 0.0
        paid_exJ = sF - J_s
        parts = {"G reimbursement J differs from PMF credit": J_s - J_t}
        split_bad = (split_m is not None and abs(split_s - split_m) > 1e-9) or c in CANCEL_REFUND
        if split_bad:
            parts["C split / cancelled deal set differently from PMF"] = diff - parts["G reimbursement J differs from PMF credit"]
        else:
            r_out_s = (sN / sF) if sF else 0.0                     # platform fee-out rate on paid incl. J (old build)
            sheet_broker = sE * (k_s or 0)
            parts["D fee-in rate: sheet vs PMF invoice"] = -((sM - sheet_broker) - inv_U)
            parts["E broker % differs from PMF invoice"] = -(sheet_broker - inv_C)
            dpaid = paid_exJ - led_g
            parts["F paid basis: sheet PENDING (5 AM) vs PMF CLEARED (evening)"] = dpaid * (1 - out_pct)
            parts["H fee-out charged on reimbursements"] = -round(r_out_s, 4) * J_s if J_s else 0.0
            parts["I fee-out rate: sheet vs PMF"] = -(round(r_out_s, 4) - out_pct) * paid_exJ
            parts["J cents rounding"] = diff - sum(parts.values())
        for k, v in parts.items():
            cat[k] += v

        # which PV build is live: fee-in on a 3%-era deal
        if sd_m >= D(2026, 8, 9) and up_pct == 0.03 and sE > 0:
            fee_in_live.append(sM / sE - (k_s or 0))

        out.append({
            "contract": c, "status_sheet": o[0], "status_mca": a["Status"],
            "funded_date_sheet": sd_s.date() if sd_s else "", "funded_date_mca": sd_m.date(),
            "split_sheet": split_s, "split_mca": split_m, "broker_sheet": k_s,
            "broker_mca_invoice": round(k_m, 4) if k_m is not None else "",
            "platform_fee_in_mca": up_pct, "platform_fee_out_mca": out_pct,
            "whole_paid_sheet": whole_paid_s, "whole_paid_pending_mca": pend, "whole_paid_cleared_mca": clr,
            "J_sheet": J_s, "credit_pmf": J_t + CANCEL_REFUND.get(c, 0.0),
            "principal_sheet": R2(sE), "principal_pmf": num(i["Funded"]),
            "fee_in_sheet": R2(sM), "fee_in_pmf": R2(num(i["Commission"]) + num(i["U M Fee"])),
            "paid_sheet_exJ": R2(sF - J_s), "paid_cleared_pmf": R2(led_g),
            "fee_out_sheet": R2(sN), "fee_out_pmf": R2(led_fee),
            "net_sheet": R2(sheet_net), "net_pmf": R2(pmf_net), "diff": R2(diff),
            "waiver_M": o[12] is True,
        })

    total_diff = sum(r["diff"] for r in out)
    print("\n== FIELD CHECKS, SHEET vs MCA TRACK")
    for k in ["status", "funded/payback", "split", "split blank (sheet silently uses 1%)", "broker %",
              "reimbursement J", "funded date",
              "paid vs pending (whole deal)"]:
        v = flags.get(k, [])
        print(f"  {k}: {len(v)} deal(s)")
        for x in v[:12]:
            print("     ", x)
        if len(v) > 12:
            print(f"      ... {len(v) - 12} more")
    if fee_in_live:
        avg = sum(fee_in_live) / len(fee_in_live)
        print(f"\n  live Payments_View build: platform fee-in on Aug-Sep 3% deals computes at {avg:.2%}"
              f" -> {'OLD build (flat 4%)' if avg > 0.035 else 'NEW build'}")

    print("\n== AVAILABLE CASH BRIDGE (fund level, internal 3.5% mgmt transfer excluded: nets to 0)")
    print(f"  sheet Available Cash (Snapshot H, 4 investors)   {sheet_avail:12,.2f}")
    print(f"  = contributions {contrib:,.2f} - withdrawals {wd:,.2f} + deal cash {sheet_avail - contrib + wd:,.2f}")
    for k in sorted(cat):
        print(f"    {-cat[k]:12,.2f}  {k}")
    print(f"  expected reserve at 11/04 (from MCA Track)          {pmf_1104:12,.2f}")
    print(f"  check: sheet - sum(deal diffs) = {sheet_avail - total_diff:,.2f}"
          f"   (contrib-wd+PMF deals = {contrib - wd + sum(r['net_pmf'] for r in out):,.2f})")
    print(f"  reserve tracker 10/29 {pmf['available']:,.2f} + cleared since refresh {led_net - pmf['payback']:,.2f}"
          f" = {pmf['available'] + led_net - pmf['payback']:,.2f}")

    os.makedirs(out_dir, exist_ok=True)
    p = os.path.join(out_dir, "deal_by_deal_2026-11-04.csv")
    with open(p, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys()))
        w.writeheader()
        for r in sorted(out, key=lambda r: -abs(r["diff"])):
            w.writerow(r)
    print("\n  wrote", p)
    # ---------------- per investor, 3.5% management transfer removed ----------------
    # The live (old) Payments_View build charges a flat 4% platform out-fee on paid incl.
    # reimbursements; Ozzie/Ezzie's fee-out above that is the 3.5% transfer, and You (ID 1)
    # receive exactly the sum as negative fee-out. Removing it moves money between
    # investors only; the fund total does not change.
    names = {int(num(str(r[0]).replace("ID", ""))): str(r[1]) for r in invs}
    diff_by = {r["contract"]: r["diff"] for r in out}
    agg = collections.defaultdict(lambda: collections.defaultdict(float))
    for r in pvr:
        c = cid10(r[0]); i = int(num(r[2]))
        F, N, pct = num(r[5]), num(r[13]), num(r[3])
        plat = round(F * 0.04 + 1e-9, 2)
        agg[i]["mgmt"] += (N - plat) if i != 1 else -(plat - N)
        agg[i]["pmf_diff"] += diff_by.get(c, 0.0) * pct
    print("\n== PER INVESTOR (3.5% mgmt transfer removed; fund total unchanged)")
    print(f"  {'':8s}{'Snapshot H':>12s}{'mgmt adj':>11s}{'no mgmt':>11s}{'= PMF':>11s}")
    tot = collections.defaultdict(float)
    for r in snap:
        i = next(k for k, v in names.items() if v == r[0])
        H = num(r[7]); adj = agg[i]["mgmt"]; nm = H + adj; pm = nm - agg[i]["pmf_diff"]
        for k, v in (("H", H), ("adj", adj), ("nm", nm), ("pm", pm)):
            tot[k] += v
        print(f"  {r[0]:8s}{H:12,.2f}{adj:11,.2f}{nm:11,.2f}{pm:11,.2f}")
    print(f"  {'TOTAL':8s}{tot['H']:12,.2f}{tot['adj']:11,.2f}{tot['nm']:11,.2f}{tot['pm']:11,.2f}")

    big = sorted(out, key=lambda r: -abs(r["diff"]))[:15]
    print("\n== LARGEST PER-DEAL DIFFERENCES (sheet net - PMF net)")
    for r in big:
        print(f"  {r['contract']} {str(r['status_sheet'])[:22]:22s} diff {r['diff']:9.2f}  paid sheet/clr "
              f"{r['paid_sheet_exJ']:9.2f}/{r['paid_cleared_pmf']:9.2f}  feein {r['fee_in_sheet']:7.2f}/{r['fee_in_pmf']:7.2f}"
              f"  J {r['J_sheet']}/{r['credit_pmf']}")
    return out, cat, pmf, pmf_1104, sheet_avail, led_by_day


if __name__ == "__main__":
    main(*sys.argv[1:8])
