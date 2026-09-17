/***********************
 * Payments View (VALUES) - header-driven layout
 * FULL FILE replacement - includes the main function AND all four helpers
 * (buildRow_Reordered_, ensurePVHeaders_, writePV_byHeaders_, clearPV_).
 * Select all in the old file and paste this over everything.
 *
 * BUILD 2026-09-17 - two changes vs the 2026-08-11 version, both verified
 * against PMF's own per-slice numbers in the MCA Track advances export:
 *
 *   1. NO PLATFORM FEE-OUT ON REIMBURSEMENTS. PMF credits clawbacks, EPA
 *      commission reimbursements and EPA management fee refunds to the
 *      reserve at 100%. The old code added Override J into "paid back" and
 *      then took the platform out-fee on the sum (119.40 overcharged fund-wide
 *      on 9/16). The platform out-fee now applies to collections only; J is
 *      still added to Paid Back to Investor after that.
 *      The 3.5% internal mgmt transfer is UNCHANGED (still computed on
 *      collections + reimbursements) - whether Ozzie/Ezzie owe mgmt fee on a
 *      commission refund is Dovi's call, not PMF's. To exclude reimbursements
 *      from the mgmt transfer too, change the one line marked MGMT-BASE below.
 *
 *   2. BUILD STAMP. The function writes PV_BUILD to Script Properties on every
 *      run and the console's pipeline log prints it ("payments_view OK :: 41s
 *      :: build 2026-09-17"). If the log shows an older build or none, a
 *      duplicate REBUILD_PaymentsView_VALUES_ in another .gs file is loading
 *      last and this file is NOT the one running. Ctrl+Shift+F the project
 *      for "function REBUILD_PaymentsView_VALUES_" and delete the extra copy.
 *      (This is exactly why the 3% fee era below was not live on 9/16 even
 *      though it is in this file.)
 *
 * Unchanged from the 2026-08-11 version:
 *   PLATFORM FEE SCHEDULE, applied per deal by funded date (the date comes
 *   from Deal_Locks column B, falling back to raw Start Date):
 *     funded on/after 2026-08-09  ->  3% in / 3% out  (first deal: 5674188893)
 *     funded before that          ->  4% in / 4% out
 *     no funded date on record    ->  4% (old era default)
 *   Per-deal MANAGER MGMT-FEE SKIP FLAG in Override column M.
 *
 * Fee In includes: platform fee (per schedule) + mgmt_in + broker_fee (Override!K)
 * Platform Fee Out (per schedule) is a real cost (leaves the fund).
 * MGMT Fee Out is an internal TRANSFER: non-managers -> Manager (ID 1).
 *
 * NOTE ON "TOTAL PAID": this function uses Override column E as collections.
 * PMF credits the reserve on CLEARED payments only. The Advances Report page
 * shows PENDING (cleared + in-flight debits). The console v2.6 ingest fills
 * column E from "Total paid cleared" whenever an MCA Track export is pasted.
 * Nothing in this file needs to know which basis was ingested.
 ***********************/

var PV_BUILD = "2026-09-17";

function REBUILD_PaymentsView_VALUES_() {
  try { PropertiesService.getScriptProperties().setProperty("PV_BUILD", PV_BUILD); } catch (e) {}

  const ss = SpreadsheetApp.getActive();

  const shOverride  = ss.getSheetByName("Override");
  const shInvestors = ss.getSheetByName("Investors");
  const shTS        = ss.getSheetByName("Time_Stamps");
  const shLocks     = ss.getSheetByName("Deal_Locks");
  const shRaw       = ss.getSheetByName("Copy of Raw_Deals");
  const shMeta      = ss.getSheetByName("Deals_Meta");
  const shPV        = ss.getSheetByName("Payments_View");

  if (!shOverride || !shInvestors || !shTS || !shLocks || !shRaw || !shMeta || !shPV) {
    throw new Error("Missing one of: Override, Investors, Time_Stamps, Deal_Locks, Copy of Raw_Deals, Deals_Meta, Payments_View");
  }

  // -----------------------------
  // Helpers
  const DAY_MS = 24 * 60 * 60 * 1000;
  const round2 = n => Math.round((Number(n) + Number.EPSILON) * 100) / 100;

  const normDeal = v => {
    const m = String(v ?? "").match(/\d+/);
    return m ? m[0] : "";
  };
  const normInv = v => {
    const m = String(v ?? "").match(/\d+/);
    return m ? Number(m[0]) : null;
  };
  const num = v => {
    if (v === "" || v == null) return 0;
    if (typeof v === "number") return v;
    const n = Number(String(v).replace(/[$,%\s,]/g, ""));
    return isNaN(n) ? 0 : n;
  };
  const pct = v => {
    if (v === "" || v == null) return null;
    if (typeof v === "number") return v;
    const s = String(v).trim();
    if (!s) return null;
    if (s.includes("%")) {
      const n = Number(s.replace(/[%\s,]/g, ""));
      return isNaN(n) ? null : n / 100;
    }
    const n = Number(s.replace(/[^\d.\-]/g, ""));
    if (isNaN(n)) return null;
    return n > 1.5 ? n / 100 : n;
  };
  const dateOnly = d => {
    if (!d) return null;
    const x = new Date(d);
    if (isNaN(x.getTime())) return null;
    x.setHours(0, 0, 0, 0);
    return x;
  };

  // Parse the per-deal mgmt-skip flag (Override column M)
  const parseSkip = v => {
    if (v === true) return true;
    const s = String(v ?? "").trim().toLowerCase();
    return s === "true" || s === "yes" || s === "y" || s === "x" || s === "1";
  };

  const MANAGER_ID = 1;

  // PMF platform fee schedule, applied PER DEAL by funded date. Newest era
  // first; the {from: null} row is the fallback for everything older and for
  // deals with no funded date on record. To change the rate again, add a new
  // row at the top - nothing else in this function needs touching.
  const FEE_SCHEDULE = [
    // months are 0-indexed: (2026, 7, 9) = August 9, 2026. MCA Track shows
    // 5674188893 funded 8/9 with upfront fee 3%, so the cutoff is 8/9.
    { from: new Date(2026, 7, 9), feeIn: 0.03, feeOut: 0.03 },
    { from: null,                 feeIn: 0.04, feeOut: 0.04 }, // the original 4% era
  ];
  const feeFor = d => {
    for (const s of FEE_SCHEDULE) {
      if (!s.from || (d && d.getTime() >= s.from.getTime())) return s;
    }
    return FEE_SCHEDULE[FEE_SCHEDULE.length - 1];
  };

  // -----------------------------
  const HEADERS = [
    "Contract ID",
    "Status",
    "Investor ID",
    "Investor %",
    "Investor Contribution",
    "Paid Back to Investor",
    "Total Expected Payback (Net of Fees)",
    "Current Profit(including fees)",
    "Remaining still owed",
    "Reached Breakeven?",
    "% Paid",
    "Performance",
    "Fee In",
    "Fee Out",
    "Remaining Balance",
    "Remaining Duration",
    "Expected Duration",
    "weekly/daily",
    "Date Funded",
    "PMF Total Funded",
    "PMF Total Payback",
    "PMF Total Paid",
    "PMF Total Outstanding Balance"
  ];

  ensurePVHeaders_(shPV, HEADERS);

  // -----------------------------
  // 1) Investors
  const invLast = shInvestors.getLastRow();
  const invRows = Math.max(0, invLast - 1);
  const invData = invRows ? shInvestors.getRange(2, 1, invRows, 6).getValues() : [];

  const investors = invData
    .map(r => ({
      id_raw: r[0],
      id: normInv(r[0]),
      joined: dateOnly(r[3]),
      mgmt_in: num(r[4]),
      mgmt_out: num(r[5]),
    }))
    .filter(x => x.id);

  if (!investors.length) {
    clearPV_(shPV);
    return;
  }

  // -----------------------------
  // 2) Time_Stamps
  const tsLast = shTS.getLastRow();
  const tsRows = Math.max(0, tsLast - 1);
  const tsData = tsRows ? shTS.getRange(2, 1, tsRows, 4).getValues() : [];

  const ts = tsData
    .map(r => ({
      id_raw: r[0],
      id: normInv(r[0]),
      contrib: num(r[1]),
      withdraw: num(r[2]),
      d: dateOnly(r[3])
    }))
    .filter(x => x.d);

  function capAsOf(invIdRaw, asOfDate) {
    let c = 0;
    const t = asOfDate.getTime();
    for (const row of ts) {
      if (row.id_raw === invIdRaw && row.d.getTime() <= t) {
        c += row.contrib - row.withdraw;
      }
    }
    return c;
  }
  function fundCapAsOf(asOfDate) {
    let c = 0;
    const t = asOfDate.getTime();
    for (const row of ts) {
      if (row.d.getTime() <= t) {
        c += row.contrib - row.withdraw;
      }
    }
    return c;
  }

  // -----------------------------
  // 3) Copy of Raw Deals
  const rawLast = shRaw.getLastRow();
  const rawRows = Math.max(0, rawLast - 1);
  const rawLastCol = shRaw.getLastColumn();

  const fundedByContractRaw = new Map();
  const rawRichByDigits = new Map();

  if (rawRows) {
    const rawVals = shRaw.getRange(2, 1, rawRows, rawLastCol).getValues();
    const contractRich = shRaw.getRange(2, 2, rawRows, 1).getRichTextValues();

    for (let i = 0; i < rawRows; i++) {
      const r = rawVals[i];
      const contractRaw = r[1];
      if (!contractRaw) continue;

      const digits = normDeal(contractRaw);
      if (digits && contractRich[i][0]) rawRichByDigits.set(digits, contractRich[i][0]);

      if (!fundedByContractRaw.has(contractRaw)) {
        fundedByContractRaw.set(contractRaw, dateOnly(r[4]));
      }
    }
  }

  // -----------------------------
  // 4) Deals_Meta
  const metaLast = shMeta.getLastRow();
  const metaRows = Math.max(0, metaLast - 1);
  const metaData = metaRows ? shMeta.getRange(2, 2, metaRows, 3).getValues() : [];
  const metaMap = new Map();
  for (const r of metaData) {
    if (!r[0]) continue;
    metaMap.set(r[0], { dur: num(r[1]) || r[1] || "", freq: r[2] || "" });
  }

  // -----------------------------
  // 5) Deal_Locks
  const locksLast = shLocks.getLastRow();
  const locksRows = Math.max(0, locksLast - 1);

  const lockHeader = shLocks.getRange(1, 1, 1, shLocks.getLastColumn()).getValues()[0].map(String);
  const lockColsByInv = new Map();
  for (let i = 0; i < lockHeader.length; i++) {
    const m = lockHeader[i].match(/^ID\s*(\d+)\s*%$/i);
    if (m) lockColsByInv.set(Number(m[1]), i);
  }

  const locksData = locksRows ? shLocks.getRange(2, 1, locksRows, shLocks.getLastColumn()).getValues() : [];
  const locksMap = new Map();
  const fundedByDealDigits = new Map();

  for (const r of locksData) {
    const deal = normDeal(r[0]);
    if (!deal) continue;

    const fd = dateOnly(r[1]);
    if (fd) fundedByDealDigits.set(deal, fd);

    const m = new Map();
    for (const inv of investors) {
      const col = lockColsByInv.get(inv.id);
      if (col == null) continue;
      const p = pct(r[col]);
      if (p != null) m.set(inv.id, p);
    }
    locksMap.set(deal, m);
  }

  // -----------------------------
  // 6) Override deals (A..M)  <-- reads through column M (skip flag)
  const ovLast = shOverride.getLastRow();
  const ovRows = Math.max(0, ovLast - 1);
  if (!ovRows) {
    clearPV_(shPV);
    return;
  }
  const ov = shOverride.getRange(2, 1, ovRows, 13).getValues();

  // -----------------------------
  // 7) Build rows
  const out = [];
  const outRichA = [];
  const today = dateOnly(new Date());

  for (const d of ov) {
    const status = d[0];
    const cid_raw = d[1];
    if (!status || !cid_raw) continue;

    const cid_digits = normDeal(cid_raw);

    const tfunded  = round2(num(d[2]));
    const tpayback = round2(num(d[3]));
    const tpaid    = round2(num(d[4]));
    const tbalance = round2(num(d[5]));
    const ppaid    = d[6];
    const perf     = d[7];

    const dfunded = fundedByDealDigits.get(cid_digits) || fundedByContractRaw.get(cid_raw) || null;

    // THIS DEAL's platform fee era, by funded date (no date = old 4% era)
    const sched = feeFor(dfunded);
    const PLATFORM_FEE_IN  = sched.feeIn;
    const PLATFORM_FEE_OUT = sched.feeOut;

    const fund_pct = (d[8] === "" || d[8] == null) ? 0.01 : (pct(d[8]) ?? num(d[8]) ?? 0.01);

    const reimb_total = round2(num(d[9]));
    const broker_fee_pct = (d[10] === "" || d[10] == null) ? 0 : (pct(d[10]) ?? 0);

    // Per-deal manager mgmt-fee skip flag (Override column M = index 12)
    const skipMgmt = parseSkip(d[12]);

    const cleared = tpaid;
    const pending = round2(Math.max(0, tpayback - tpaid));

    const fund_contrib_total  = round2(tfunded  * fund_pct);
    const fund_returned_total = round2(cleared  * fund_pct);
    const fund_pending_total  = round2(pending  * fund_pct);
    const fund_expected_return_total = round2(tpayback * fund_pct);

    const meta = metaMap.get(cid_raw) || { dur: "", freq: "" };
    const expDur = meta.dur;

    const richCID = rawRichByDigits.get(cid_digits) ||
      SpreadsheetApp.newRichTextValue().setText(String(cid_raw)).build();

    // ---- First pass: base splits per investor (before mgmt transfer)
    const rowsForDeal = [];

    for (const inv of investors) {

      if (dfunded && inv.joined && inv.joined.getTime() > dfunded.getTime()) {
        rowsForDeal.push({
          eligible: false,
          inv,
          ipctVal: 0,
          contrib: 0,
          returned_plus_reimb_gross: 0,
          pend_gross: 0,
          expected_gross: 0,
          fee_in: 0,
          fee_out_platform: 0,
          fee_out_mgmt: 0,
          effMgmtOut: 0,
          remaining_net_nonmgr: 0,
          expected_net_nonmgr: 0
        });
        continue;
      }

      let ipctVal = 0;
      const lock = locksMap.get(cid_digits);
      if (lock && lock.has(inv.id)) {
        ipctVal = lock.get(inv.id);
      } else if (dfunded) {
        const inv_cap = capAsOf(inv.id_raw, dfunded);
        const fund_cap = fundCapAsOf(dfunded);
        ipctVal = (fund_cap === 0) ? 0 : inv_cap / fund_cap;
      }

      const contrib        = round2(fund_contrib_total  * ipctVal);
      const returned_gross = round2(fund_returned_total * ipctVal);
      const pend_gross     = round2(fund_pending_total  * ipctVal);
      const reimb_alloc    = round2(reimb_total * ipctVal);

      const returned_plus_reimb_gross = round2(returned_gross + reimb_alloc);
      const expected_gross = round2(fund_expected_return_total * ipctVal);

      // Fee In (real cost)
      const fee_in_pct = PLATFORM_FEE_IN + inv.mgmt_in + broker_fee_pct;
      const fee_in = round2(contrib * fee_in_pct);

      // Effective mgmt-out for THIS deal (0 when the deal is flagged)
      const effMgmtOut = skipMgmt ? 0 : (inv.mgmt_out || 0);

      // Fee Out: platform (real) on COLLECTIONS ONLY - PMF never takes its
      // out-fee on a reimbursement (build 2026-09-17). Mgmt (internal
      // transfer) base is unchanged - see MGMT-BASE in the header note.
      const fee_out_platform = round2(returned_gross * PLATFORM_FEE_OUT);
      const fee_out_mgmt     = round2(returned_plus_reimb_gross * effMgmtOut);   // MGMT-BASE

      const remaining_net_nonmgr = round2(pend_gross * (1 - PLATFORM_FEE_OUT - effMgmtOut));
      const expected_net_nonmgr  = round2(expected_gross * (1 - PLATFORM_FEE_OUT - effMgmtOut) - fee_in);

      rowsForDeal.push({
        eligible: true,
        inv,
        ipctVal,
        contrib,
        returned_plus_reimb_gross,
        pend_gross,
        expected_gross,
        fee_in,
        fee_out_platform,
        fee_out_mgmt,
        effMgmtOut,
        remaining_net_nonmgr,
        expected_net_nonmgr
      });
    }

    // ---- Second pass: total mgmt transfer collected from non-managers
    let mgmt_cash_transfer = 0;
    let mgmt_pending_transfer = 0;
    let mgmt_expected_transfer = 0;

    for (const row of rowsForDeal) {
      const invId = row.inv.id;
      if (!row.eligible) continue;
      if (invId === MANAGER_ID) continue;

      mgmt_cash_transfer     = round2(mgmt_cash_transfer     + row.fee_out_mgmt);
      mgmt_pending_transfer  = round2(mgmt_pending_transfer  + round2(row.pend_gross     * (row.effMgmtOut || 0)));
      mgmt_expected_transfer = round2(mgmt_expected_transfer + round2(row.expected_gross * (row.effMgmtOut || 0)));
    }

    // ---- Third pass: write final rows (manager receives the transfers)
    for (const row of rowsForDeal) {
      const inv = row.inv;

      if (!row.eligible) {
        outRichA.push([richCID]);
        out.push(buildRow_Reordered_({
          cid_raw, status, iid: inv.id, ipctVal: 0, contrib: 0,
          paid_back_to_inv: 0, expected_payback_net_fees: 0, current_profit: 0,
          remaining_still_owed: 0, clearedFlag: "No", ppaid, perf,
          fee_in: 0, fee_out: 0, remaining_balance: 0, remaining_duration: "",
          exp_dur: expDur, freq: meta.freq, dfunded, tfunded, tpayback, tpaid, tbalance
        }));
        continue;
      }

      let remaining_duration = "";
      if (dfunded && expDur !== "" && expDur != null && !isNaN(Number(expDur))) {
        const durN = Number(expDur);
        const elapsedDays = Math.floor((today.getTime() - dateOnly(dfunded).getTime()) / DAY_MS);
        if (String(meta.freq || "").toLowerCase().includes("week")) {
          const elapsedWeeks = Math.floor(elapsedDays / 7);
          remaining_duration = Math.max(0, Math.ceil(durN - elapsedWeeks));
        } else {
          remaining_duration = Math.max(0, Math.ceil(durN - elapsedDays));
        }
      }

      let paid_back_to_inv = row.returned_plus_reimb_gross;
      let fee_out = round2(row.fee_out_platform + row.fee_out_mgmt);
      let remaining_balance = row.remaining_net_nonmgr;
      let expected_payback_net_fees = row.expected_net_nonmgr;

      if (inv.id === MANAGER_ID) {
        paid_back_to_inv = row.returned_plus_reimb_gross;

        // Manager nets the mgmt transfer out of Fee Out (0 when deal is flagged)
        fee_out = round2(row.fee_out_platform - mgmt_cash_transfer);

        const manager_base_remaining = round2(row.pend_gross * (1 - PLATFORM_FEE_OUT));
        remaining_balance = round2(manager_base_remaining + mgmt_pending_transfer);

        const manager_base_expected = round2(row.expected_gross * (1 - PLATFORM_FEE_OUT) - row.fee_in);
        expected_payback_net_fees = round2(manager_base_expected + mgmt_expected_transfer);
      }

      const current_profit = round2(paid_back_to_inv - row.contrib - row.fee_in - fee_out);
      const clearedFlag = (paid_back_to_inv >= row.contrib ? "Yes" : "No");
      const remaining_still_owed = row.pend_gross;

      outRichA.push([richCID]);
      out.push(buildRow_Reordered_({
        cid_raw, status, iid: inv.id, ipctVal: row.ipctVal, contrib: row.contrib,
        paid_back_to_inv, expected_payback_net_fees, current_profit,
        remaining_still_owed, clearedFlag, ppaid, perf,
        fee_in: row.fee_in, fee_out, remaining_balance, remaining_duration,
        exp_dur: expDur, freq: meta.freq, dfunded, tfunded, tpayback, tpaid, tbalance
      }));
    }
  }

  // -----------------------------
  // 8) Write to Payments_View
  writePV_byHeaders_(shPV, HEADERS, out, outRichA);
}

/** Build row in the header order */
function buildRow_Reordered_(x) {
  return [
    x.cid_raw,
    x.status,
    x.iid,
    x.ipctVal,
    x.contrib,
    x.paid_back_to_inv,
    x.expected_payback_net_fees,
    x.current_profit,
    x.remaining_still_owed,
    x.clearedFlag,
    x.ppaid,
    x.perf,
    x.fee_in,
    x.fee_out,
    x.remaining_balance,
    x.remaining_duration,
    x.exp_dur,
    x.freq,
    x.dfunded || "",
    x.tfunded,
    x.tpayback,
    x.tpaid,
    x.tbalance
  ];
}

function ensurePVHeaders_(shPV, headers) {
  if (shPV.getLastColumn() < headers.length) {
    shPV.insertColumnsAfter(shPV.getLastColumn(), headers.length - shPV.getLastColumn());
  }
  shPV.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function writePV_byHeaders_(shPV, headers, rows, richA) {
  const last = shPV.getLastRow();
  if (last > 1) shPV.getRange(2, 1, last - 1, headers.length).clearContent();
  if (!rows.length) return;

  shPV.getRange(2, 1, rows.length, headers.length).setValues(rows);

  if (richA && richA.length) {
    shPV.getRange(2, 1, richA.length, 1).setRichTextValues(richA);
  }
}

function clearPV_(shPV) {
  const last = shPV.getLastRow();
  if (last > 1) shPV.getRange(2, 1, last - 1, shPV.getLastColumn()).clearContent();
}
