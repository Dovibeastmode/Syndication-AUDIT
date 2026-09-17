/*********************************************************************
 * FUND CONSOLE - Web App server (Code side) - v2.6
 *
 * Replaces WebApp_Code.gs v2.5 in the SAME Apps Script project that
 * contains the pipeline (works with both SYN_ and PMF_ prefixed
 * projects - it auto-detects which functions exist).
 *
 * Deploy: Deploy > Manage deployments > edit > New version
 *   (keeps the same URL). First deployment: Execute as Me,
 *   access "Only myself".
 *********************************************************************/

var WA_VERSION = "2.6";
/* v2.6 (2026-09-17, after the PMF reserve reconciliation):
 *  - CLEARED vs PENDING. PMF credits the reserve only for payments that have
 *    CLEARED. The Advances Report page shows PENDING (cleared + in-flight
 *    debits), which is what every ingest to date has stored in Total Paid.
 *    The ingest now recognizes an MCA TRACK EXPORT (the advances xls opened
 *    in Sheets or Excel and copied, or a CSV) by its header row and fills the
 *    canonical Total Paid column from "Total paid cleared". Total Balance is
 *    then Payback minus cleared. Nothing downstream changes.
 *    Both preview and commit report which basis was ingested, and the
 *    pipeline log records it ("paid basis :: cleared (MCA Track export)").
 *  - Optional MCA_EXPORT_URL Script Property: if set, the daily auto-ingest
 *    fetches it first (with the login cookies). If the response is a text
 *    export (CSV, TSV or an HTML table with a "Total paid cleared" column)
 *    it is used; a binary .xls cannot be parsed server-side, in which case
 *    the run logs that and falls back to the Advances Report page (pending
 *    basis). Until a text export URL is wired, the 5 AM run will put the
 *    sheet back on the pending basis - see the deploy note.
 *  - The pipeline log now prints the Payments_View BUILD stamp
 *    ("payments_view OK :: 41s :: build 2026-09-17"). If the stamp is missing
 *    or older than the file you pasted, a duplicate REBUILD_PaymentsView_VALUES_
 *    in another .gs file is the one actually running.
 *  - Export pastes lose the missed-payment suffix on statuses ("Open 3"
 *    becomes "Open"). Bucketing and attention rules are prefix-based, so
 *    nothing breaks; only the suffix display is lost.
 */
/* v2.5 (daily auto-ingest):
 *  - WA_autoIngestTimer logs into MCA Track server-side (UrlFetchApp), pulls
 *    every Advances Report page, and commits through the exact same guarded
 *    path as a manual ingest. Runs on a Google time-driven trigger (5-6 AM),
 *    no PC or browser needed.
 *  - It NEVER force-commits: if known deals are missing from the fetch it
 *    aborts and writes the reason to the LAST RUN log.
 *  - Setup (done once, by Dovi, in the Apps Script editor under Project
 *    Settings > Script Properties):
 *      MCA_LOGIN_URL    the login page URL
 *      MCA_ADVANCES_URL the Advances Report URL with {PAGE} where the page
 *                       number goes; set page size 100 in it
 *      MCA_USER         the MCA Track username
 *      MCA_PASS         the MCA Track password
 *      MCA_EXPORT_URL   (optional, v2.6) a text export URL, see above
 *    Credentials stay in Script Properties - the console never displays them.
 */
/* v2.4: single-pass pipeline, every step timed, full log persisted.
 * v2.3: commit canonicalizes the paste by contract ID position; Override
 *       sync lives here; fail-fast pipeline; same-day liquidity overwrite.
 * v2.1: strict closed whitelist (Closed, EPA, LOC, Overpaid); attention
 *       dismissals persisted.
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index.html")
    .setTitle("Syndication Fund Console")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/* ---------- helpers ---------- */
function WA_tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function WA_d_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v))
    return Utilities.formatDate(v, WA_tz_(), "yyyy-MM-dd");
  return String(v);
}
function WA_num_(v) {
  if (v === "" || v == null) return 0;
  if (typeof v === "number") return v;
  const n = Number(String(v).replace(/[$,%\s,]/g, ""));
  return isNaN(n) ? 0 : n;
}
function WA_bool_(v) {
  if (v === true) return true;
  const s = String(v == null ? "" : v).trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "x" || s === "1";
}
const WA_CID_RE = /^\d{7,}\s*\(\d+\)$/;

/* Which investor is looking? 0 = manager/owner. Investors col H may hold
   an email per investor for read-only portal access. */
function WA_lens_() {
  let email = "";
  try { email = (Session.getActiveUser().getEmail() || "").toLowerCase(); } catch (e) {}
  let owner = "";
  try { owner = (Session.getEffectiveUser().getEmail() || "").toLowerCase(); } catch (e) {}
  if (!email || email === owner) return { lens: 0, name: "Manager" };
  const inv = SpreadsheetApp.getActive().getSheetByName("Investors");
  const last = inv.getLastRow();
  const vals = inv.getRange(2, 1, Math.max(0, last - 1), 8).getValues();
  for (const r of vals) {
    const m = String(r[0] || "").match(/\d+/);
    const invEmail = String(r[7] || "").trim().toLowerCase();
    if (m && invEmail && invEmail === email) return { lens: Number(m[0]), name: String(r[1] || "") };
  }
  throw new Error("No access: your Google account (" + email + ") is not registered. Ask the fund manager to add your email to the Investors sheet.");
}
function WA_mgrOnly_() {
  const who = WA_lens_();
  if (who.lens > 0) throw new Error("Manager only.");
  return who;
}

/* ---------- pipeline resolver (SYN_ or PMF_ project) ---------- */
function WA_call_(names) {
  const g = globalThis;
  for (const n of names) if (typeof g[n] === "function") { g[n](); return n; }
  return null;
}
/* Pipeline order per Dovi's rulings: Override and Payments_View update
   FIRST, then the day's liquidity row snapshots the post-update book.
   Fail-fast: the first error aborts the rest. Every step is timed and the
   whole log is persisted so an error can never be lost. */
function WA_runPipeline() {
  const log = [];
  const t0 = Date.now();
  const call = names => { if (!WA_call_(names)) throw new Error("script function not found"); };
  const steps = [
    ["1 clean raw", () => call(["CLEANUP_CopyOfRawDeals_"])],
    ["2 sync override", () => { const r = WA_syncOverride_(); return r.updated + " updated, " + r.added + " new"; }],
    ["3 deal locks", () => call(["REBUILD_DealLocks_FromLiquidAllocation"])],
    ["4 payments_view", () => {
      try { PropertiesService.getScriptProperties().deleteProperty("PV_BUILD"); } catch (e) {}
      call(["REBUILD_PaymentsView_VALUES_"]);
      let b = "";
      try { b = PropertiesService.getScriptProperties().getProperty("PV_BUILD") || ""; } catch (e) {}
      return b ? "build " + b : "build UNKNOWN - an old copy of REBUILD_PaymentsView_VALUES_ is running";
    }],
    ["5 liquidity snapshot", () => WA_appendLiquidToday_()],
  ];
  for (const [label, fn] of steps) {
    const s0 = Date.now();
    try {
      const r = fn();
      log.push(label + " OK :: " + Math.round((Date.now() - s0) / 1000) + "s" + (typeof r === "string" && r ? " :: " + r : ""));
    } catch (e) {
      log.push(label + " ERROR after " + Math.round((Date.now() - s0) / 1000) + "s: " + e.message);
      log.push("ABORTED - later steps skipped so bad data cannot flow downstream");
      break;
    }
  }
  SpreadsheetApp.flush();
  WA_saveLastRun_(log, Date.now() - t0);
  return log;
}
function WA_saveLastRun_(log, ms) {
  try {
    PropertiesService.getScriptProperties().setProperty("WA_LAST_RUN", JSON.stringify({
      when: Utilities.formatDate(new Date(), WA_tz_(), "yyyy-MM-dd HH:mm"), ms, log: log.slice(0, 40)
    }));
  } catch (e) {}
}
function WA_lastRun_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("WA_LAST_RUN") || "null"); }
  catch (e) { return null; }
}
/* Which "Total Paid" basis the last commit stored: cleared or pending. */
function WA_savePaidBasis_(basis) {
  try { PropertiesService.getScriptProperties().setProperty("WA_PAID_BASIS", basis); } catch (e) {}
}
function WA_paidBasis_() {
  try { return PropertiesService.getScriptProperties().getProperty("WA_PAID_BASIS") || "pending (Advances Report page)"; }
  catch (e) { return "unknown"; }
}
/* After any single edit: rebuild PV, then refresh TODAY's liquidity row so it
   converges to its final form as attention items get resolved. */
function WA_rebuildPV_() {
  try { WA_call_(["REBUILD_PaymentsView_VALUES_"]); SpreadsheetApp.flush(); } catch (e) {}
  try { WA_appendLiquidToday_(); SpreadsheetApp.flush(); } catch (e) {}
}

/* ---------- Override sync (owned by the console) ----------
   Header-driven and batched. */
function WA_syncOverride_() {
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName("Copy of Raw_Deals");
  const dst = ss.getSheetByName("Override");
  if (!src || !dst) throw new Error("missing Copy of Raw_Deals or Override");
  const srcLastRow = src.getLastRow(), srcLastCol = src.getLastColumn();
  if (srcLastRow < 2) throw new Error("Copy of Raw_Deals has no data rows");
  const srcVals = src.getRange(1, 1, srcLastRow, srcLastCol).getValues();
  const norm = s => String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
  const hdr = srcVals[0].map(norm);
  const col = n => hdr.indexOf(norm(n));
  const cC = col("Contract ID"), cS = col("Status/Last Valid"), cF = col("Total Funded"),
    cP = col("Total Payback"), cPd = col("Total Paid"), cB = col("Total Balance"),
    cPct = col("% Paid"), cPerf = col("Performance Standard");
  if ([cC, cS, cF, cP, cPd, cB, cPct, cPerf].some(i => i === -1))
    throw new Error("Copy of Raw_Deals row 1 is not the cleaned header row");
  const fresh = new Map();
  for (let r = 1; r < srcVals.length; r++) {
    const row = srcVals[r];
    const cid = String(row[cC] || "").trim();
    if (!cid || fresh.has(cid)) continue;
    fresh.set(cid, [row[cS], cid, row[cF], row[cP], row[cPd], row[cB], row[cPct], row[cPerf]]);
  }
  if (fresh.size < 5) throw new Error("only " + fresh.size + " deals in Copy of Raw_Deals - refusing to sync");

  dst.getRange(1, 1, 1, 11).setValues([[
    "Status/Last Valid", "Contract ID", "Total Funded", "Total Payback", "Total Paid",
    "Total Balance", "% Paid", "Performance Standard", "Split %", "Reimbursments", "Brokers Comission"
  ]]);
  // only touch the real deal rows: find the last row with a contract ID in B,
  // not getLastRow() (checkboxes and the column L formula extend far below)
  const dstLast = dst.getLastRow();
  const colB = dstLast > 1 ? dst.getRange(2, 2, dstLast - 1, 1).getValues() : [];
  let lastFilled = -1;
  for (let i = 0; i < colB.length; i++) if (String(colB[i][0] || "").trim()) lastFilled = i;
  const nRows = lastFilled + 1;
  const dstVals = nRows > 0 ? dst.getRange(2, 1, nRows, 8).getValues() : [];
  let updated = 0;
  const outAH = dstVals.map(r => {
    const cid = String(r[1] || "").trim();
    const m = cid ? fresh.get(cid) : null;
    if (m) { updated++; fresh.delete(cid); return m; }
    return [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7]];
  });
  // one batched write updates A:H only - I/J/K (manual) and L/M (formula +
  // waiver checkbox) are never touched
  if (outAH.length) dst.getRange(2, 1, outAH.length, 8).setValues(outAH);
  const news = [...fresh.values()].map(m => m.concat(["", "", ""]));
  if (news.length) dst.getRange(nRows + 2, 1, news.length, 11).setValues(news);
  return { updated, added: news.length };
}

/* ---------- today's liquidity row (owned by the console) ----------
   Same-day OVERWRITE, never a duplicate row, headers reconciled against the
   investor list on every run. Reads Snapshot AFTER a flush so the row
   reflects the freshly rebuilt book. */
function WA_appendLiquidToday_() {
  SpreadsheetApp.flush();
  const ss = SpreadsheetApp.getActive();
  const snap = ss.getSheetByName("Snapshot");
  const inv = ss.getSheetByName("Investors");
  const liquid = ss.getSheetByName("Liquid_Allocation");
  if (!snap || !inv || !liquid) throw new Error("missing Snapshot, Investors, or Liquid_Allocation");
  const invLast = inv.getLastRow();
  const invList = (invLast > 1 ? inv.getRange(2, 1, invLast - 1, 2).getValues() : [])
    .map(r => ({ id: (String(r[0] || "").match(/\d+/) || [null])[0], name: String(r[1] || "").trim() }))
    .filter(x => x.id && x.name);
  if (!invList.length) throw new Error("no investors found");
  const snapLast = snap.getLastRow();
  const cashByName = new Map();
  (snapLast > 1 ? snap.getRange(2, 1, snapLast - 1, 8).getValues() : []).forEach(r => {
    const nm = String(r[0] || "").trim();
    if (nm) cashByName.set(nm, WA_num_(r[7]));
  });
  WA_reconcileLiquidHeaders_();
  const lastCol = Math.max(2, liquid.getLastColumn());
  const hdr = liquid.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const colById = new Map();
  for (let c = 2; c <= lastCol; c++) {
    const m = hdr[c - 1].match(/^ID\s*(\d+)$/i);
    if (m) colById.set(m[1], c);
  }
  const tz = WA_tz_();
  const todayKey = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const p = todayKey.split("-");
  const todayNoon = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0, 0);
  const liLast = liquid.getLastRow();
  let targetRow = liLast + 1;
  if (liLast > 1) {
    const tail = Math.min(10, liLast - 1);
    const dates = liquid.getRange(liLast - tail + 1, 1, tail, 1).getValues();
    for (let i = dates.length - 1; i >= 0; i--) {
      const v = dates[i][0];
      if (!v) continue;
      let key = "";
      try { key = Utilities.formatDate(new Date(v), tz, "yyyy-MM-dd"); } catch (e) {}
      if (key === todayKey) targetRow = liLast - tail + 1 + i;
      break;
    }
  }
  const rowVals = new Array(lastCol).fill("");
  rowVals[0] = todayNoon;
  for (const x of invList) {
    const c = colById.get(x.id);
    if (c) rowVals[c - 1] = cashByName.has(x.name) ? cashByName.get(x.name) : 0;
  }
  liquid.getRange(targetRow, 1, 1, lastCol).setValues([rowVals]);
  liquid.getRange(targetRow, 1).setNumberFormat("m/d/yyyy");
  return (targetRow > liLast ? "added " : "updated ") + todayKey;
}

/* ---------- main state read ---------- */
function WA_getState() {
  const who = WA_lens_();
  const ss = SpreadsheetApp.getActive();
  const state = { lens: who.lens, lensName: who.name, version: WA_VERSION,
    asOf: Utilities.formatDate(new Date(), WA_tz_(), "yyyy-MM-dd HH:mm") };

  // Investors (8 cols: through email in H)
  const inv = ss.getSheetByName("Investors");
  const invLast = inv.getLastRow();
  state.investors = inv.getRange(2, 1, Math.max(0, invLast - 1), 8).getValues()
    .filter(r => r[0] !== "")
    .map(r => ({ id: Number((String(r[0]).match(/\d+/) || [0])[0]), name: String(r[1] || ""),
      join: WA_d_(r[3]), mgmtIn: WA_num_(r[4]), mgmtOut: WA_num_(r[5]),
      mgmtEarned: WA_num_(r[6]), email: String(r[7] || "").trim() }));

  // Snapshot (already computed by the sheet's own formulas)
  const snap = ss.getSheetByName("Snapshot");
  const nInv = state.investors.length;
  state.snapshot = snap.getRange(2, 1, nInv, 15).getValues().map(r => ({
    name: String(r[0] || ""), contrib: WA_num_(r[1]), cashOutOpen: WA_num_(r[2]),
    cashOutTotal: WA_num_(r[3]), wd: WA_num_(r[4]), cashRec: WA_num_(r[5]),
    nav: WA_num_(r[6]), avail: WA_num_(r[7]), deployed: WA_num_(r[8]),
    expProfitOpen: WA_num_(r[9]), projLife: WA_num_(r[10]), expectedPayback: WA_num_(r[11]),
    outstanding: WA_num_(r[12]), fees: WA_num_(r[13]), projReturn: WA_num_(r[14]),
  }));

  // Analytics (selected columns)
  const ana = ss.getSheetByName("Analytics");
  state.analytics = ana.getRange(2, 1, nInv, 21).getValues().map(r => ({
    name: String(r[0] || ""), deals: WA_num_(r[1]), open: WA_num_(r[2]),
    closed: WA_num_(r[3]), defaults: WA_num_(r[4]), defaultRate: WA_num_(r[5]),
    recovery: WA_num_(r[7]), lost: WA_num_(r[8]), avgToDefault: WA_num_(r[9]),
    retClosed: WA_num_(r[10]), retAfterFees: WA_num_(r[11]), profClosed: WA_num_(r[12]),
    avgToClose: WA_num_(r[13]), avgDealSize: WA_num_(r[14]),
    avgDur: WA_num_(r[16]), projAnnual: WA_num_(r[17]),
  }));

  // Deals_Meta: durations, dates
  const dm = ss.getSheetByName("Deals_Meta");
  const dmLast = dm.getLastRow();
  const dmVals = dm.getRange(2, 1, Math.max(0, dmLast - 1), 6).getValues();
  const meta = {};
  for (const r of dmVals) {
    const cid = String(r[1] || "").trim();
    if (!cid) continue;
    meta[cid] = { dur: r[2] === "" ? null : WA_num_(r[2]), freq: String(r[3] || ""),
      dc: WA_d_(r[4]), dd: WA_d_(r[5]) };
  }

  // Copy of Raw_Deals: funded dates
  const raw = ss.getSheetByName("Copy of Raw_Deals");
  const rawLast = raw.getLastRow();
  const rawVals = rawLast > 1 ? raw.getRange(2, 2, rawLast - 1, 4).getValues() : [];
  const fundedBy = {};
  for (const r of rawVals) {
    const cid = String(r[0] || "").trim();
    if (WA_CID_RE.test(cid)) fundedBy[cid] = WA_d_(r[3]);
  }

  // Override: the deal list + manual columns, through M (waiver flag)
  const ov = ss.getSheetByName("Override");
  const ovLast = ov.getLastRow();
  const ovVals = ovLast > 1 ? ov.getRange(2, 1, ovLast - 1, 13).getValues() : [];
  state.deals = [];
  for (const r of ovVals) {
    const cid = String(r[1] || "").trim();
    if (!cid) continue;
    const m = meta[cid] || {};
    state.deals.push({ cid, status: String(r[0] || ""), gross: WA_num_(r[2]),
      payback: WA_num_(r[3]), paid: WA_num_(r[4]), balance: WA_num_(r[5]),
      split: r[8] === "" ? null : WA_num_(r[8]), reimb: WA_num_(r[9]),
      broker: r[10] === "" ? null : WA_num_(r[10]),
      waived: WA_bool_(r[12]),
      funded: fundedBy[cid] || null, dur: m.dur != null ? m.dur : null, freq: m.freq || "",
      dc: m.dc || null, dd: m.dd || null });
  }

  // Payments_View: per-contract fund aggregates + fund-wide closed sums
  const pv = ss.getSheetByName("Payments_View");
  const pvLast = pv.getLastRow();
  const pvVals = pvLast > 1 ? pv.getRange(2, 1, pvLast - 1, 14).getValues() : [];
  const agg = {}; // cid -> {e: contributions, h: profit}
  let cE = 0, cF = 0, cM = 0, cN = 0;
  for (const r of pvVals) {
    const cid = String(r[0] || "").trim();
    if (!cid) continue;
    const st = String(r[1] || "");
    const E = WA_num_(r[4]), F = WA_num_(r[5]), H = WA_num_(r[7]), M = WA_num_(r[12]), N = WA_num_(r[13]);
    const a = agg[cid] || (agg[cid] = { e: 0, h: 0 });
    a.e += E; a.h += H;
    if (/^(Closed|EPA|LOC|Overpaid)/.test(st)) { cE += E; cF += F; cM += M; cN += N; }
  }
  state.pvAgg = agg;
  state.fundClosed = { cE, cF, cM, cN };

  // Liquid_Allocation: full history (capped at the most recent 400 rows)
  const la = ss.getSheetByName("Liquid_Allocation");
  const laLast = la.getLastRow();
  if (laLast > 1) {
    const laCols = la.getLastColumn();
    const hdr = la.getRange(1, 1, 1, laCols).getValues()[0].map(String);
    const start = Math.max(2, laLast - 399);
    const rows = la.getRange(start, 1, laLast - start + 1, laCols).getValues();
    state.liquid = { cols: hdr.slice(1),
      rows: rows.filter(r => r[0]).map(r => [WA_d_(r[0])].concat(r.slice(1).map(WA_num_))) };
  } else state.liquid = null;

  // Time_Stamps ledger (most recent 300)
  const tsh = ss.getSheetByName("Time_Stamps");
  const tsLast = tsh.getLastRow();
  const tsVals = tsLast > 1 ? tsh.getRange(2, 1, tsLast - 1, 4).getValues() : [];
  state.timeStamps = tsVals
    .filter(r => r[0] !== "" && r[0] != null)
    .map(r => ({ idRaw: String(r[0]), id: Number((String(r[0]).match(/\d+/) || [0])[0]),
      contrib: WA_num_(r[1]), wd: WA_num_(r[2]), date: WA_d_(r[3]) }))
    .slice(-300);

  // Attention dismissals + last pipeline run + auto-ingest status (manager only)
  state.dismissed = who.lens === 0 ? WA_dismissed_() : [];
  state.lastRun = who.lens === 0 ? WA_lastRun_() : null;
  state.auto = who.lens === 0 ? WA_autoStatus_() : null;
  state.paidBasis = who.lens === 0 ? WA_paidBasis_() : null;

  // Investor lens: strip manager-only data
  if (who.lens > 0) {
    const idx = state.investors.findIndex(v => v.id === who.lens);
    state.snapshot = idx >= 0 ? [state.snapshot[idx]] : [];
    state.analytics = idx >= 0 ? [state.analytics[idx]] : [];
    state.deals = []; state.pvAgg = {}; state.liquid = null;
    state.timeStamps = state.timeStamps.filter(t => t.id === who.lens);
    state.investors = state.investors.map(v => ({ id: v.id, name: v.name }));
    state.myDeals = WA_getPV(who.lens, "", 600).rows;
  }
  return state;
}

/* ---------- Payments_View rows (filtered, capped) ---------- */
function WA_getPV(invId, cidQuery, cap) {
  const who = WA_lens_();
  if (who.lens > 0) invId = who.lens; // investors only ever see their own rows
  cap = cap || 300;
  const pv = SpreadsheetApp.getActive().getSheetByName("Payments_View");
  const last = pv.getLastRow();
  if (last < 2) return { rows: [], total: 0 };
  const vals = pv.getRange(2, 1, last - 1, 14).getValues();
  const rows = [];
  let total = 0;
  for (const r of vals) {
    const cid = String(r[0] || "").trim();
    if (!cid) continue;
    const iid = Number(r[2]);
    if (invId && iid !== Number(invId)) continue;
    if (cidQuery && cid.indexOf(cidQuery) === -1) continue;
    total++;
    if (rows.length < cap) rows.push({ cid, status: String(r[1] || ""), id: iid,
      ip: WA_num_(r[3]), contrib: WA_num_(r[4]), paid: WA_num_(r[5]),
      expn: WA_num_(r[6]), profit: WA_num_(r[7]), feeIn: WA_num_(r[12]), feeOut: WA_num_(r[13]) });
  }
  return { rows, total };
}

/* ---------- Deal_Locks lookup ---------- */
function WA_getLock(cidQuery) {
  WA_mgrOnly_();
  const dl = SpreadsheetApp.getActive().getSheetByName("Deal_Locks");
  const last = dl.getLastRow();
  const lastCol = dl.getLastColumn();
  const hdr = dl.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const vals = dl.getRange(2, 1, Math.max(0, last - 1), lastCol).getValues();
  for (const r of vals) {
    const cid = String(r[0] || "");
    if (cid.indexOf(cidQuery) !== -1) {
      return { cid, funded: WA_d_(r[1]), cols: hdr.slice(3), vals: r.slice(3).map(WA_num_) };
    }
  }
  return null;
}

/* ---------- ingest: export detection (v2.6) ----------
   An MCA Track EXPORT (the advances xls opened in Sheets/Excel and copied,
   or a CSV) has a header row and carries BOTH "Total paid cleared" and
   "Total paid pending". PMF credits the reserve on CLEARED only, so when an
   export is detected the paste is rewritten into the same 13-column line
   shape as an Advances Report page copy, with Total Paid = cleared and
   Total Balance = Payback - cleared. Every parser below then works unchanged.
   Contract IDs in the export are bare digits with the advance number in its
   own column; they are re-joined as "5674190918 (1)" to match Override. */
function WA_splitLine_(line) {
  if (line.indexOf("\t") !== -1) return line.split("\t");
  const out = []; let cur = "", q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line.charAt(i);
    if (ch === '"') { if (q && line.charAt(i + 1) === '"') { cur += '"'; i++; } else q = !q; }
    else if (ch === "," && !q) { out.push(cur); cur = ""; }
    else cur += ch;
  }
  out.push(cur);
  return out;
}
function WA_normalizeIngest_(text) {
  const src = String(text || "");
  const lines = src.split(/\r?\n/);
  const norm = s => String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
  let hdrIdx = -1, hdr = null;
  for (let i = 0; i < Math.min(lines.length, 40); i++) {
    const f = WA_splitLine_(lines[i]).map(norm);
    if (f.indexOf("contract id") !== -1 && f.indexOf("total paid cleared") !== -1) { hdrIdx = i; hdr = f; break; }
  }
  if (hdrIdx === -1) return { text: src, basis: "pending (Advances Report page)", isExport: false, rows: 0 };
  const col = n => hdr.indexOf(norm(n));
  const c = { cid: col("Contract id"), adv: col("Advance number"), status: col("Status"),
    start: col("Start date"), factor: col("Factor rate"), funded: col("Advance amount"),
    comm: col("Commission amount"), payback: col("Payback"), cleared: col("Total paid cleared"),
    pct: col("Percentage paid"), perf: col("Performance") };
  if ([c.cid, c.status, c.funded, c.payback, c.cleared].some(i => i === -1))
    throw new Error("export header found but a required column is missing (Contract id, Status, Advance amount, Payback, Total paid cleared)");
  const out = []; const seen = new Set();
  for (let i = hdrIdx + 1; i < lines.length; i++) {
    const f = WA_splitLine_(lines[i]).map(s => String(s).trim());
    const digits = String(f[c.cid] || "").replace(/\D/g, "");
    if (digits.length < 7) continue;
    const advN = c.adv !== -1 ? (Math.round(WA_num_(f[c.adv])) || 1) : 1;
    const cid = digits + " (" + advN + ")";
    if (seen.has(cid)) continue;
    seen.add(cid);
    const g = k => (k !== -1 && f[k] !== undefined) ? f[k] : "";
    const payback = WA_num_(f[c.payback]);
    const paid = WA_num_(f[c.cleared]);
    const balance = Math.round((payback - paid) * 100) / 100;
    out.push(["", cid, "", g(c.status), g(c.start), g(c.factor), String(WA_num_(f[c.funded])),
      String(WA_num_(g(c.comm))), String(payback), String(paid), String(balance), g(c.pct), g(c.perf)].join("\t"));
  }
  return { text: out.join("\n"), basis: "cleared (MCA Track export)", isExport: true, rows: out.length };
}

/* ---------- ingest: preview then commit ---------- */
function WA_parseRaw_(text) {
  const out = []; const seen = new Set();
  for (const line of String(text).split(/\r?\n/)) {
    const f = line.split("\t").map(s => String(s).trim());
    let i = -1;
    for (let k = 0; k < f.length; k++) if (WA_CID_RE.test(f[k])) { i = k; break; }
    if (i === -1) continue;
    const cid = f[i];
    if (seen.has(cid)) continue;
    seen.add(cid);
    out.push({ cid, status: f[i + 2] || "", gross: WA_num_(f[i + 5]),
      payback: WA_num_(f[i + 7]), paid: WA_num_(f[i + 8]) });
  }
  return out;
}
/* Bucket rule (locked by Dovi 2026-07-14): closed is a strict whitelist.
   Everything else that is not Open/Lowered is a default. */
function WA_cat_(s) {
  if (!s) return "default";
  if (/^(Open|Lowered)/.test(s)) return "open";
  if (/^(Closed|EPA|LOC|Overpaid)/.test(s)) return "closed";
  return "default";
}
/* Which Override deals are absent from this paste? Committing while deals
   are missing rebuilds Deal_Locks without them (allocations lost), so the
   client must show this list before commit. */
function WA_missingFromPaste_(parsedRows) {
  const inPaste = new Set(parsedRows.map(r => r.cid));
  const ov = SpreadsheetApp.getActive().getSheetByName("Override");
  const last = ov.getLastRow();
  const missing = [];
  if (last > 1) {
    const vals = ov.getRange(2, 1, last - 1, 2).getValues();
    for (const r of vals) {
      const c = String(r[1] || "").trim();
      if (c && !inPaste.has(c)) missing.push({ cid: c, status: String(r[0] || "") });
    }
  }
  // open-category deals first: those are the ones that still move money
  missing.sort((a, b) => (WA_cat_(a.status) === "open" ? 0 : 1) - (WA_cat_(b.status) === "open" ? 0 : 1));
  return missing;
}
function WA_previewIngest(text) {
  WA_mgrOnly_();
  const n = WA_normalizeIngest_(text);
  const rows = WA_parseRaw_(n.text);
  if (!rows.length) return { error: "No contract rows detected in the paste." };
  const ov = SpreadsheetApp.getActive().getSheetByName("Override");
  const last = ov.getLastRow();
  const cur = new Map();
  if (last > 1) {
    const vals = ov.getRange(2, 1, last - 1, 2).getValues();
    for (const r of vals) { const c = String(r[1] || "").trim(); if (c) cur.set(c, String(r[0] || "")); }
  }
  const res = { parsed: rows.length, updated: 0, drastic: [], newDeals: [], overrideCount: cur.size,
    paidBasis: n.basis, isExport: n.isExport };
  for (const r of rows) {
    if (!cur.has(r.cid)) { res.newDeals.push({ cid: r.cid, status: r.status, gross: r.gross }); continue; }
    res.updated++;
    const oc = WA_cat_(cur.get(r.cid)), nc = WA_cat_(r.status);
    if (oc !== nc) res.drastic.push({ cid: r.cid, from: cur.get(r.cid), to: r.status, toCat: nc });
  }
  const missing = WA_missingFromPaste_(rows);
  res.missingCount = missing.length;
  res.missing = missing.slice(0, 50);
  return res;
}
/* Canonical 13-column layout for Copy of Raw_Deals - the layout every
   downstream step is built against. */
var WA_RAW_HEADER = ["", "Contract ID", "", "Status/Last Valid", "Start Date", "Factor Rate",
  "Total Funded", "Total Commission", "Total Payback", "Total Paid", "Total Balance",
  "% Paid", "Performance Standard"];
/* Locate each deal line by its contract ID in ANY column, then take the
   fields at fixed offsets from it. This is what makes the commit immune to
   MCA Track's copy landing shifted by a column. */
function WA_parseRawFull_(text) {
  const out = []; const seen = new Set();
  for (const line of String(text).split(/\r?\n/)) {
    const f = line.split("\t").map(s => String(s).trim());
    let i = -1;
    for (let k = 0; k < f.length; k++) if (WA_CID_RE.test(f[k])) { i = k; break; }
    if (i === -1) continue;
    const cid = f[i];
    if (seen.has(cid)) continue;
    seen.add(cid);
    const g = k => f[i + k] !== undefined ? f[i + k] : "";
    out.push(["", cid, "", g(2), g(3), g(4), g(5), g(6), g(7), g(8), g(9), g(10), g(11)]);
  }
  return out;
}
function WA_commitIngest(text, force) {
  WA_mgrOnly_();
  const n = WA_normalizeIngest_(text);
  const parsed = WA_parseRaw_(n.text);
  if (parsed.length < 5) {
    return { ok: false, error: "Only " + parsed.length + " deal rows detected - refusing to commit. Paste all pages first." };
  }
  const missing = WA_missingFromPaste_(parsed);
  if (missing.length > 0 && force !== true) {
    return { ok: false, needsForce: true, missingCount: missing.length, missing: missing.slice(0, 20),
      parsed: parsed.length };
  }
  // write the canonicalized grid, never the raw paste
  const grid = [WA_RAW_HEADER].concat(WA_parseRawFull_(n.text));
  const raw = SpreadsheetApp.getActive().getSheetByName("Copy of Raw_Deals");
  raw.clearContents();
  raw.getRange(1, 1, grid.length, 13).setValues(grid);
  SpreadsheetApp.flush();
  WA_savePaidBasis_(n.basis);
  const plog = WA_runPipeline();
  const log = ["paid basis :: " + n.basis].concat(plog);
  WA_saveLastRun_(log, 0);
  return { ok: true, log, parsed: parsed.length, missingCount: missing.length, paidBasis: n.basis };
}

/* ---------- writes: Override + Deals_Meta ---------- */
function WA_findRow_(sheet, col, cid) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const vals = sheet.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++)
    if (String(vals[i][0] || "").trim() === cid) return i + 2;
  return -1;
}
function WA_updateOverride(cid, patch) {
  WA_mgrOnly_();
  const ov = SpreadsheetApp.getActive().getSheetByName("Override");
  const row = WA_findRow_(ov, 2, cid);
  if (row === -1) throw new Error("Contract not found in Override: " + cid);
  if (patch.split !== undefined) ov.getRange(row, 9).setValue(patch.split);
  if (patch.reimb !== undefined) ov.getRange(row, 10).setValue(patch.reimb);
  if (patch.broker !== undefined) ov.getRange(row, 11).setValue(patch.broker);
  // column 13 = M, the per-deal mgmt fee waiver checkbox; column 12 = L formula, never touched
  if (patch.waived !== undefined) ov.getRange(row, 13).setValue(patch.waived === true);
  WA_rebuildPV_();
  return { ok: true };
}
function WA_updateMeta(cid, patch) {
  WA_mgrOnly_();
  const dm = SpreadsheetApp.getActive().getSheetByName("Deals_Meta");
  let row = WA_findRow_(dm, 2, cid);
  if (row === -1) {
    // new deal not yet in Deals_Meta: append its row
    row = dm.getLastRow() + 1;
    dm.getRange(row, 2).setValue(cid);
  }
  if (patch.dur !== undefined) dm.getRange(row, 3).setValue(patch.dur === null ? "" : patch.dur);
  if (patch.freq !== undefined) dm.getRange(row, 4).setValue(patch.freq);
  if (patch.dc !== undefined) dm.getRange(row, 5).setValue(patch.dc ? new Date(patch.dc + "T12:00:00") : "");
  if (patch.dd !== undefined) dm.getRange(row, 6).setValue(patch.dd ? new Date(patch.dd + "T12:00:00") : "");
  WA_rebuildPV_();
  return { ok: true };
}

/* ---------- daily auto-ingest (MCA Track fetched server-side) ---------- */
function WA_autoCfg_() {
  const p = PropertiesService.getScriptProperties();
  return { loginUrl: p.getProperty("MCA_LOGIN_URL") || "", advUrl: p.getProperty("MCA_ADVANCES_URL") || "",
    user: p.getProperty("MCA_USER") || "", pass: p.getProperty("MCA_PASS") || "",
    exportUrl: p.getProperty("MCA_EXPORT_URL") || "" };
}
function WA_autoStatus_() {
  const c = WA_autoCfg_();
  let trig = false;
  try { trig = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "WA_autoIngestTimer"); } catch (e) {}
  return { loginUrl: !!c.loginUrl, advUrl: !!c.advUrl, advUrlHasPage: c.advUrl.indexOf("{PAGE}") !== -1,
    user: !!c.user, pass: !!c.pass, trigger: trig, exportUrl: !!c.exportUrl };
}
function WA_cookieJar_(jar, resp) {
  const h = resp.getAllHeaders();
  let sc = h["Set-Cookie"] || h["set-cookie"];
  if (!sc) return jar;
  if (!Array.isArray(sc)) sc = [sc];
  for (const c of sc) {
    const kv = String(c).split(";")[0];
    const eq = kv.indexOf("=");
    if (eq > 0) jar[kv.slice(0, eq).trim()] = kv.slice(eq + 1).trim();
  }
  return jar;
}
function WA_cookieHdr_(jar) {
  return Object.keys(jar).map(k => k + "=" + jar[k]).join("; ");
}
function WA_absUrl_(base, href) {
  if (!href) return base;
  if (/^https?:/i.test(href)) return href;
  const m = base.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
  if (!m) return href;
  if (href.charAt(0) === "/") return m[1] + href;
  const dir = (m[2] || "/").replace(/[^\/]*$/, "");
  return m[1] + dir + href;
}
/* Log into MCA Track: read the login form (hidden fields included), post the
   credentials, keep the session cookies. */
function WA_mcaLogin_(cfg, log) {
  const jar = {};
  const r0 = UrlFetchApp.fetch(cfg.loginUrl, { muteHttpExceptions: true, followRedirects: true });
  WA_cookieJar_(jar, r0);
  const html = r0.getContentText();
  const formM = html.match(/<form[^>]*>[\s\S]*?<\/form>/i);
  if (!formM) throw new Error("no <form> found on the login page (" + r0.getResponseCode() + ")");
  const form = formM[0];
  const actionM = form.match(/action\s*=\s*["']([^"']*)["']/i);
  const action = WA_absUrl_(cfg.loginUrl, actionM ? actionM[1] : "");
  const payload = {};
  let userField = "", passField = "";
  const inputs = form.match(/<input[^>]*>/gi) || [];
  for (const inp of inputs) {
    const name = (inp.match(/name\s*=\s*["']([^"']*)["']/i) || [])[1];
    if (!name) continue;
    const type = ((inp.match(/type\s*=\s*["']([^"']*)["']/i) || [])[1] || "text").toLowerCase();
    const value = (inp.match(/value\s*=\s*["']([^"']*)["']/i) || [])[1] || "";
    if (type === "password") { passField = name; continue; }
    if (type === "hidden") { payload[name] = value; continue; }
    if (!userField && /user|email|login|name/i.test(name)) userField = name;
  }
  if (!userField) userField = "username";
  if (!passField) passField = "password";
  payload[userField] = cfg.user;
  payload[passField] = cfg.pass;
  log.push("login form :: action found, user field '" + userField + "'");
  let resp = UrlFetchApp.fetch(action, { method: "post", payload, muteHttpExceptions: true,
    followRedirects: false, headers: { Cookie: WA_cookieHdr_(jar) } });
  WA_cookieJar_(jar, resp);
  // follow up to 3 redirects by hand so every Set-Cookie is captured
  for (let i = 0; i < 3; i++) {
    const code = resp.getResponseCode();
    if (code < 300 || code >= 400) break;
    const loc = resp.getAllHeaders()["Location"] || resp.getAllHeaders()["location"];
    if (!loc) break;
    resp = UrlFetchApp.fetch(WA_absUrl_(action, String(loc)), { muteHttpExceptions: true,
      followRedirects: false, headers: { Cookie: WA_cookieHdr_(jar) } });
    WA_cookieJar_(jar, resp);
  }
  log.push("login :: HTTP " + resp.getResponseCode() + ", " + Object.keys(jar).length + " cookies");
  return jar;
}
/* Turn an HTML table's rows into tab-separated lines. keepAll=true keeps
   every row (header included) for export tables; otherwise only rows that
   carry a "5674183695 (1)" style contract ID, as on the Advances page. */
function WA_htmlToLines_(html, keepAll) {
  const unesc = s => s.replace(/<[^>]+>/g, "\t").replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, '"').replace(/&#0?39;|&apos;/gi, "'").replace(/&lt;/gi, "<").replace(/&gt;/gi, ">");
  const lines = [];
  const trs = html.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  for (const tr of trs) {
    const cells = (tr.match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) || [])
      .map(c => unesc(c).replace(/\s*\t\s*/g, " ").replace(/\s+/g, " ").trim());
    if (!cells.length) continue;
    if (keepAll || cells.some(c => WA_CID_RE.test(c))) lines.push(cells.join("\t"));
  }
  return lines;
}
/* Fetch every Advances page until a page adds nothing new. */
function WA_mcaFetchAll_(cfg, jar, log) {
  const seen = new Set();
  const lines = [];
  for (let page = 1; page <= 25; page++) {
    const url = cfg.advUrl.replace("{PAGE}", String(page));
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true,
      headers: { Cookie: WA_cookieHdr_(jar) } });
    if (r.getResponseCode() !== 200) { log.push("page " + page + " :: HTTP " + r.getResponseCode() + " - stopping"); break; }
    const pageLines = WA_htmlToLines_(r.getContentText(), false);
    let fresh = 0;
    for (const ln of pageLines) {
      const cid = (ln.split("\t").find(c => WA_CID_RE.test(c)) || "").trim();
      if (cid && !seen.has(cid)) { seen.add(cid); lines.push(ln); fresh++; }
    }
    log.push("page " + page + " :: " + pageLines.length + " rows, " + fresh + " new");
    if (fresh === 0) break;
  }
  return lines;
}
/* v2.6: try the text export first (cleared basis). Returns null when the URL
   is not set or the response cannot be parsed, so the page path runs. */
function WA_mcaFetchExport_(cfg, jar, log) {
  if (!cfg.exportUrl) return null;
  let r;
  try {
    r = UrlFetchApp.fetch(cfg.exportUrl, { muteHttpExceptions: true, followRedirects: true,
      headers: { Cookie: WA_cookieHdr_(jar) } });
  } catch (e) { log.push("export :: fetch failed: " + e.message); return null; }
  if (r.getResponseCode() !== 200) { log.push("export :: HTTP " + r.getResponseCode() + " - falling back to page"); return null; }
  const ct = String((r.getAllHeaders()["Content-Type"] || r.getAllHeaders()["content-type"] || "")).toLowerCase();
  const bytes = r.getContent();
  // binary .xls starts with the compound-document magic D0 CF 11 E0
  if (bytes.length > 4 && (bytes[0] & 0xff) === 0xd0 && (bytes[1] & 0xff) === 0xcf) {
    log.push("export :: binary .xls returned - cannot parse server-side, falling back to page (pending basis)");
    return null;
  }
  let body = r.getContentText();
  if (/<table/i.test(body)) body = WA_htmlToLines_(body, true).join("\n");
  let n;
  try { n = WA_normalizeIngest_(body); } catch (e) { log.push("export :: " + e.message); return null; }
  if (!n.isExport) { log.push("export :: no 'Total paid cleared' header in the response (" + ct + ") - falling back to page"); return null; }
  log.push("export :: " + n.rows + " deals, cleared basis");
  return n;
}
/* The shared auto-ingest body. commit=false is a safe dry run. */
function WA_autoIngest_(commit) {
  const t0 = Date.now();
  const log = ["AUTO-INGEST " + (commit ? "RUN" : "DRY RUN")];
  try {
    const cfg = WA_autoCfg_();
    const miss = ["MCA_LOGIN_URL", "MCA_ADVANCES_URL", "MCA_USER", "MCA_PASS"]
      .filter(k => !PropertiesService.getScriptProperties().getProperty(k));
    if (miss.length) throw new Error("missing Script Properties: " + miss.join(", "));
    if (cfg.advUrl.indexOf("{PAGE}") === -1) throw new Error("MCA_ADVANCES_URL must contain {PAGE} where the page number goes");
    const jar = WA_mcaLogin_(cfg, log);
    let text, basis;
    const exp = WA_mcaFetchExport_(cfg, jar, log);
    if (exp) { text = exp.text; basis = exp.basis; }
    else {
      const lines = WA_mcaFetchAll_(cfg, jar, log);
      text = lines.join("\n"); basis = "pending (Advances Report page)";
    }
    log.push("paid basis :: " + basis);
    const parsed = WA_parseRaw_(text);
    log.push("parsed :: " + parsed.length + " deals total");
    if (parsed.length < 5) throw new Error("only " + parsed.length + " deals parsed - login or page URL is probably wrong");
    const sample = parsed[0];
    log.push("sample :: " + sample.cid + " :: " + sample.status + " :: funded " + sample.gross + " :: paid " + sample.paid);
    const missing = WA_missingFromPaste_(parsed);
    if (missing.length) log.push("missing from fetch :: " + missing.slice(0, 8).map(x => x.cid).join(", ") + (missing.length > 8 ? " +" + (missing.length - 8) + " more" : ""));
    if (!commit) {
      log.push("dry run complete - nothing written");
      WA_saveLastRun_(log, Date.now() - t0);
      return log;
    }
    if (missing.length) throw new Error(missing.length + " known deals missing from the fetch - auto-ingest never force-commits. Review and ingest manually if this is expected.");
    const grid = [WA_RAW_HEADER].concat(WA_parseRawFull_(text));
    const raw = SpreadsheetApp.getActive().getSheetByName("Copy of Raw_Deals");
    raw.clearContents();
    raw.getRange(1, 1, grid.length, 13).setValues(grid);
    SpreadsheetApp.flush();
    WA_savePaidBasis_(basis);
    const plog = WA_runPipeline();
    const full = log.concat(plog);
    WA_saveLastRun_(full, Date.now() - t0);
    return full;
  } catch (e) {
    log.push("AUTO-INGEST ERROR: " + e.message);
    WA_saveLastRun_(log, Date.now() - t0);
    return log;
  }
}
/* Trigger target - runs unattended, no lens check. */
function WA_autoIngestTimer() { WA_autoIngest_(true); }
/* Console endpoints. */
function WA_autoIngestRun() { WA_mgrOnly_(); return WA_autoIngest_(true); }
function WA_autoIngestDry() { WA_mgrOnly_(); return WA_autoIngest_(false); }
function WA_installAutoTrigger() {
  WA_mgrOnly_();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "WA_autoIngestTimer") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("WA_autoIngestTimer").timeBased().everyDays(1).atHour(5).create();
  return { ok: true };
}
function WA_removeAutoTrigger() {
  WA_mgrOnly_();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "WA_autoIngestTimer") ScriptApp.deleteTrigger(t); });
  return { ok: true };
}

/* ---------- attention dismissals (Script Properties) ---------- */
function WA_dismissed_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("WA_ATTN_DISMISSED") || "[]"); }
  catch (e) { return []; }
}
function WA_saveDismissed_(list) {
  while (list.length > 500) list.shift();
  PropertiesService.getScriptProperties().setProperty("WA_ATTN_DISMISSED", JSON.stringify(list));
}
function WA_dismissAttn(key) {
  WA_mgrOnly_();
  const list = WA_dismissed_();
  if (list.indexOf(key) === -1) list.push(key);
  WA_saveDismissed_(list);
  return { ok: true };
}
function WA_restoreAttn(key) {
  WA_mgrOnly_();
  WA_saveDismissed_(WA_dismissed_().filter(k => k !== key));
  return { ok: true };
}

/* ---------- writes: Time_Stamps ledger ---------- */
function WA_addTimeStamp(invId, contrib, withdraw, dateStr) {
  WA_mgrOnly_();
  const id = Number(invId);
  if (!id || id < 1) throw new Error("Pick an investor.");
  const c = WA_num_(contrib), w = WA_num_(withdraw);
  if (c < 0 || w < 0) throw new Error("Amounts must be positive numbers.");
  if (c === 0 && w === 0) throw new Error("Enter a deposit or a withdrawal amount.");
  if (!dateStr) throw new Error("Pick a date.");
  const ts = SpreadsheetApp.getActive().getSheetByName("Time_Stamps");
  ts.appendRow(["ID " + id, c, w, new Date(dateStr + "T12:00:00")]);
  SpreadsheetApp.flush();
  // cash on hand changed - refresh today's liquidity row
  try { WA_appendLiquidToday_(); } catch (e) {}
  return { ok: true };
}

/* ---------- writes: Investors admin ---------- */
function WA_invRow_(inv, id) {
  const last = inv.getLastRow();
  if (last < 2) return -1;
  const vals = inv.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    const m = String(vals[i][0] || "").match(/\d+/);
    if (m && Number(m[0]) === Number(id)) return i + 2;
  }
  return -1;
}
function WA_updateInvestor(id, patch) {
  WA_mgrOnly_();
  const inv = SpreadsheetApp.getActive().getSheetByName("Investors");
  const row = WA_invRow_(inv, id);
  if (row === -1) throw new Error("Investor ID " + id + " not found.");
  // Only touch the manual cells. C (allocation) and G (fees earned) are formulas.
  if (patch.name !== undefined) inv.getRange(row, 2).setValue(String(patch.name));
  if (patch.join !== undefined && patch.join) inv.getRange(row, 4).setValue(new Date(patch.join + "T12:00:00"));
  if (patch.mgmtIn !== undefined) inv.getRange(row, 5).setValue(WA_num_(patch.mgmtIn));
  if (patch.mgmtOut !== undefined) inv.getRange(row, 6).setValue(WA_num_(patch.mgmtOut));
  if (patch.email !== undefined) inv.getRange(row, 8).setValue(String(patch.email).trim());
  // mgmt rates and join date change the fee math and eligibility
  if (patch.mgmtIn !== undefined || patch.mgmtOut !== undefined || patch.join !== undefined) WA_rebuildPV_();
  else SpreadsheetApp.flush();
  return { ok: true };
}
function WA_addInvestor(data) {
  WA_mgrOnly_();
  if (!data || !String(data.name || "").trim()) throw new Error("Investor needs a name.");
  if (!data.join) throw new Error("Investor needs a join date.");
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName("Investors");
  const last = inv.getLastRow();
  let maxId = 0;
  if (last > 1) {
    for (const r of inv.getRange(2, 1, last - 1, 1).getValues()) {
      const m = String(r[0] || "").match(/\d+/);
      if (m) maxId = Math.max(maxId, Number(m[0]));
    }
  }
  const id = maxId + 1;
  const row = last + 1;
  inv.getRange(row, 1).setValue("ID " + id);
  inv.getRange(row, 2).setValue(String(data.name).trim());
  inv.getRange(row, 4).setValue(new Date(data.join + "T12:00:00"));
  inv.getRange(row, 5).setValue(WA_num_(data.mgmtIn));
  inv.getRange(row, 6).setValue(WA_num_(data.mgmtOut));
  if (data.email) inv.getRange(row, 8).setValue(String(data.email).trim());

  // Keep Liquid_Allocation headers in sync so the new investor's column
  // exists - a missing header silently breaks their allocations.
  WA_reconcileLiquidHeaders_();

  // First deposit, if provided
  if (WA_num_(data.contrib) > 0) {
    ss.getSheetByName("Time_Stamps").appendRow(["ID " + id, WA_num_(data.contrib), 0, new Date(data.join + "T12:00:00")]);
  }
  WA_rebuildPV_();
  return { ok: true, id,
    note: "Check the Snapshot and Analytics tabs: if their formulas do not fill row " + row + " automatically, copy the row above down." };
}
function WA_reconcileLiquidHeaders_() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName("Investors");
  const liquid = ss.getSheetByName("Liquid_Allocation");
  const last = inv.getLastRow();
  if (last < 2) return;
  const ids = [];
  for (const r of inv.getRange(2, 1, last - 1, 1).getValues()) {
    const m = String(r[0] || "").match(/\d+/);
    if (m) ids.push(Number(m[0]));
  }
  const lastCol = Math.max(1, liquid.getLastColumn());
  const hdr = liquid.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const have = new Set();
  for (const h of hdr) { const m = h.match(/^ID\s*(\d+)$/i); if (m) have.add(Number(m[1])); }
  let nextCol = lastCol + 1;
  if (!hdr[0]) { liquid.getRange(1, 1).setValue("Funded Date"); }
  for (const id of ids) {
    if (!have.has(id)) { liquid.getRange(1, nextCol).setValue("ID " + id); nextCol++; }
  }
}
