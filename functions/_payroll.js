// ════════════════════════════════════════════════════════════════════
// PAYROLL DEDUCTION ENGINE — 2026 CRA rates, age-stratified
//
// Pure functions. No IO, no side effects. The pay-run endpoints
// (calculate / run / t4) all share this single source of truth so the
// rates can never drift between preview and commit.
//
// Scope: Postma Contracting Inc. family payroll. Six kids ages 8–20.
// - Under 18: CPP exempt, family-EI exempt, under-18 tax supplement applies.
// - 18+:      CPP 5.95% above the basic exemption, PRORATED PER PAY PERIOD
//             ($3,500 ÷ pays per year — CRA T4127 / PDOC method), family-EI exempt.
//
// All rates verified 2026 against T4032-ON and CRA indexation factor 1.020
// (federal) / 1.019 (Ontario). ON bracket thresholds cross-checked via
// WebSearch against CRA and multiple payroll sources (Apr 23 2026).
//
// DO NOT pull numbers from older spec documents — they are stale.
// ════════════════════════════════════════════════════════════════════

// ─── Named constants, 2026 ──────────────────────────────────────────

export const RATES_2026 = {
  // CPP — both sides 5.95% on pensionable earnings between exemption and YMPE
  cppRate: 0.0595,
  cppBasicExemption: 3500,
  cppYmpe: 74600,
  cppMaxAnnualEe: 4230.45,     // (74600 − 3500) × 5.95%

  // CPP2 — both sides 4% on earnings between YMPE and YAMPE
  cpp2Rate: 0.04,
  cpp2Yampe: 85000,
  cpp2MaxAnnualEe: 416,        // (85000 − 74600) × 4%

  // EI — not applicable for family-EI-exempt employees. Included for
  // completeness; the caller gates on employee.familyEiExempt === true.
  eiRateEe: 0.0163,
  eiRateEr: 0.0228,
  eiEmployerMultiplier: 1.4,   // employer premium = 1.4 × employee premium
  eiMie: 68900,

  // Federal basic personal amount (linear phase-down above ~$181K — for
  // family payroll well under that, we use the full amount).
  fedBpa: 16452,

  // ── Verified Nov 16 2026 against CRA T4032-ON / T4127 (Jan 1 2026) ──
  // CPP base/first-additional split (K2 credit uses only the 4.95% base part)
  cppBaseRate: 0.0495,
  cppFirstAdditionalRate: 0.0100,
  cppBaseMaxEe: 3519.45,       // max base-portion contribution (K2 cap)
  eiMaxEe: 1123.07,            // max EI premium (K2 cap)
  canadaEmploymentAmount: 1501, // CEA → K4 = min(14% × A, 14% × 1,501)
  // Ontario surtax (V1) on basic Ontario tax T4
  onSurtax1Threshold: 5818,  onSurtax1Rate: 0.20,
  onSurtax2Threshold: 7446,  onSurtax2Rate: 0.36,

  // TD1 claim-code chart step (T4008-ON Jan 2026, Chart 1): code 1 = BPA,
  // code 2 = 16,452.01–19,285, … each code $2,833 wide. The formula uses
  // the mid-point of the range: BPA + (code − 1.5) × step.
  fedClaimStep: 2833,

  // There is NO under-18 supplement in CRA's withholding formulas (T4127
  // K1 = 14% × TD1 claim amount, no age term). Kept at 0 so old callers
  // don't break. (Was 844 — removed Nov 2026 in the CRA-formula rewrite.)
  fedUnder18Supplement: 0,
  under18SupplementReductionThreshold: 3533,

  // Federal tax brackets — 2026 (post the 14% reduction in effect full-year).
  // k = CRA's published constant K (T4032-ON Chart 1) — verified Nov 16 2026.
  fedBrackets: [
    { upTo: 58523,  rate: 0.14,  k: 0 },
    { upTo: 117045, rate: 0.205, k: 3804 },
    { upTo: 181440, rate: 0.26,  k: 10241 },
    { upTo: 258482, rate: 0.29,  k: 15685 },
    { upTo: Infinity, rate: 0.33, k: 26024 },
  ],

  // Ontario BPA and under-18 supplement
  onBpa: 12989,
  onUnder18Supplement: 0,      // not in T4127 — see fedUnder18Supplement
  // TD1ON claim-code chart step (T4008-ON Jan 2026, Chart 2): code 2 =
  // 12,989.01–15,787, each code $2,798 wide.
  onClaimStep: 2798,

  // k = CRA's published constant KP (T4032-ON Chart 2) — verified Nov 16 2026.
  // Ontario tax brackets — 2026 (indexed 1.019, except $150K/$220K which
  // are fixed by statute and not indexed)
  onBrackets: [
    { upTo: 53891,  rate: 0.0505, k: 0 },
    { upTo: 107785, rate: 0.0915, k: 2210 },
    { upTo: 150000, rate: 0.1116, k: 4376 },
    { upTo: 220000, rate: 0.1216, k: 5876 },
    { upTo: Infinity, rate: 0.1316, k: 8076 },
  ],

  // Ontario low-income tax reduction — $300 base reduction, phases out
  // linearly as basic ON tax rises from $300 to $600. Effect: a resident
  // with taxable income up to ~$18,930 owes zero ON tax; above ~$24,870
  // the reduction is exhausted.
  onLowIncomeReductionBase: 300,
};

// ─── Pay frequency ──────────────────────────────────────────────────

/** Pay periods per year (CRA "P"). Default monthly. */
export const PAY_FREQUENCIES = { weekly: 52, biweekly: 26, semimonthly: 24, monthly: 12 };

export function normalizePayFrequency(v) {
  const k = String(v || '').toLowerCase().replace(/[^a-z]/g, '');
  return PAY_FREQUENCIES[k] ? k : 'monthly';
}

/**
 * How many regular pay periods a run covers, from its period dates.
 * A normal monthly run (Oct 1–31) = 1. A catch-up run (Apr 1–Sep 30) = 6.
 * Rounded to the nearest whole period, never less than 1. CRA allows one
 * per-period basic exemption for each period the payment covers.
 */
export function periodsCovered(startIso, endIso, periodsPerYear = 12) {
  const a = Date.parse(startIso), b = Date.parse(endIso);
  if (isNaN(a) || isNaN(b) || b < a) return 1;
  const days = Math.round((b - a) / 86400000) + 1;
  return Math.max(1, Math.round(days / (365 / periodsPerYear)));
}

// ─── Age math ───────────────────────────────────────────────────────

/** Compute an integer age on a given pay date from an ISO DOB. */
export function ageOnDate(dobIso, payDateIso) {
  const dob = new Date(dobIso);
  const pay = new Date(payDateIso);
  if (isNaN(dob.getTime()) || isNaN(pay.getTime())) return null;
  let age = pay.getUTCFullYear() - dob.getUTCFullYear();
  const m = pay.getUTCMonth() - dob.getUTCMonth();
  if (m < 0 || (m === 0 && pay.getUTCDate() < dob.getUTCDate())) age--;
  return age;
}

/**
 * Does CPP apply to a pay on this date? CRA: CPP starts the month AFTER
 * the employee turns 18 and stops the month after they turn 70. A pay in
 * the 18th-birthday month itself is still exempt. Unknown DOB → applies.
 */
export function cppApplies(dobIso, payDateIso) {
  const dob = new Date(dobIso), pay = new Date(payDateIso);
  if (isNaN(dob.getTime()) || isNaN(pay.getTime())) return true;
  const monthIndex = d => d.getUTCFullYear() * 12 + d.getUTCMonth();
  const startsMonth = monthIndex(dob) + 18 * 12 + 1;   // month after 18th birthday
  const stopsMonth = monthIndex(dob) + 70 * 12 + 1;    // month after 70th birthday
  const m = monthIndex(pay);
  return m >= startsMonth && m < stopsMonth;
}

// ─── Bracket helpers ────────────────────────────────────────────────

/** Progressive tax on `taxable` against a bracket table. */
function bracketTax(taxable, brackets) {
  if (taxable <= 0) return 0;
  let tax = 0;
  let prevCeiling = 0;
  for (const { upTo, rate } of brackets) {
    if (taxable <= upTo) {
      tax += (taxable - prevCeiling) * rate;
      return tax;
    }
    tax += (upTo - prevCeiling) * rate;
    prevCeiling = upTo;
  }
  return tax;
}

// ─── CPP ────────────────────────────────────────────────────────────

/**
 * CPP base + CPP2 employee portion for a pay run — CRA T4127 / PDOC method.
 *
 *   CPP = 5.95% × (gross − exemption for this pay), capped at the annual max
 *   exemption for this pay = $3,500 ÷ P × periods covered
 *
 * P = pay periods per year (12 for monthly). So a $1,500 monthly pay:
 *   ($1,500 − $291.67) × 5.95% = $71.90   — the number CRA's calculator gives.
 * A catch-up payment covering 6 months gets 6 × $291.67 of exemption.
 *
 * The exemption is NOT "used up" by the first pay of the year — every pay
 * gets its own share. (The app used to front-load the whole $3,500 on the
 * first run, which disagreed with CRA's calculator and under-deducted for
 * a mid-year start. Fixed Nov 2026 after Andrea checked it against PDOC.)
 *
 * Inputs:
 *   grossPay        - this pay run's gross
 *   ytdPensionable  - YTD pensionable earnings BEFORE this run (for CPP2 only)
 *   ytdCppPaid      - YTD CPP (base) already withheld BEFORE this run
 *   ytdCpp2Paid     - YTD CPP2 already withheld BEFORE this run
 *   cppExempt       - under 18 etc.
 *   periodsPerYear  - P (12 monthly)
 *   periodsInRun    - pay periods this run covers (1 normally)
 *
 * Returns { cppBase, cpp2, math } — math is shown to the user.
 */
export function calculateCpp(grossPay, ytdPensionable, ytdCppPaid, ytdCpp2Paid, cppExempt = false, periodsPerYear = 12, periodsInRun = 1) {
  const R = RATES_2026;
  const P = periodsPerYear > 0 ? periodsPerYear : 12;
  const N = periodsInRun > 0 ? periodsInRun : 1;
  const exemption = (R.cppBasicExemption / P) * N;
  const math = { periodsPerYear: P, periodsCovered: N, exemption: round2(exemption), pensionable: 0, rate: R.cppRate, capped: false };
  if (cppExempt || grossPay <= 0) return { cppBase: 0, cpp2: 0, math };

  // CPP base: per-pay exemption, then cap at what's left of the annual max
  const pensionable = Math.max(0, grossPay - exemption);
  math.pensionable = round2(pensionable);
  const uncapped = pensionable * R.cppRate;
  const roomLeft = Math.max(0, R.cppMaxAnnualEe - ytdCppPaid);
  const cppBaseThisRun = Math.min(uncapped, roomLeft);
  math.capped = uncapped > roomLeft;

  // CPP2 (YMPE → YAMPE @ 4%) — cumulative on YTD earnings
  const ytdAfter = ytdPensionable + grossPay;
  const cpp2CappedAfter = Math.min(Math.max(0, ytdAfter - R.cppYmpe), R.cpp2Yampe - R.cppYmpe);
  const cumulativeCpp2 = cpp2CappedAfter * R.cpp2Rate;
  const cpp2ThisRun = Math.max(0, Math.min(cumulativeCpp2, R.cpp2MaxAnnualEe) - ytdCpp2Paid);

  return { cppBase: round2(cppBaseThisRun), cpp2: round2(cpp2ThisRun), math };
}

// ─── TD1 claim codes ────────────────────────────────────────────────

/**
 * Total claim amount implied by a TD1 claim code.
 *   0      → no personal amounts at this employer (another employer has
 *            the basic amount, or income is taxed from dollar one)
 *   1      → basic personal amount only
 *   2–10   → mid-point of that code's range on the CRA chart
 *   >10    → beyond the chart; CRA says calculate manually — we treat the
 *            value as the claim AMOUNT itself if it looks like dollars
 *            (≥ 1000), otherwise cap at code 10.
 * Anything missing/invalid → code 1 (CRA's default when no TD1 is on file).
 */
export function claimAmountForCode(code, bpa, step) {
  const c = Number(code);
  if (!isFinite(c) || c < 0) return bpa;          // no/invalid code → basic
  if (c === 0) return 0;
  if (c === 1) return bpa;
  if (c >= 1000) return c;                         // explicit dollar amount
  const k = Math.min(Math.floor(c), 10);
  return bpa + (k - 1.5) * step;
}

// ─── Income tax — CRA T4127 "Option 1" (the PDOC method) ───────────

/** Bracket rate R and CRA's published constant K for annual income A. */
function rateAndConstant(A, brackets) {
  for (const b of brackets) if (A <= b.upTo) return { R: b.rate, K: b.k };
  const last = brackets[brackets.length - 1];
  return { R: last.rate, K: last.k };
}

/** Ontario Health Premium (V2) on annual taxable income A — T4127. */
export function ontarioHealthPremium(A) {
  if (A <= 20000) return 0;
  if (A <= 36000) return Math.min(300, 0.06 * (A - 20000));
  if (A <= 48000) return Math.min(450, 300 + 0.06 * (A - 36000));
  if (A <= 72000) return Math.min(600, 450 + 0.25 * (A - 48000));
  if (A <= 200000) return Math.min(750, 600 + 0.25 * (A - 72000));
  return Math.min(900, 750 + 0.25 * (A - 200000));
}

/**
 * Federal + Ontario income tax for ONE regular pay period, by CRA's
 * T4127 Option 1 formula — the same arithmetic CRA's online calculator
 * (PDOC) uses. Annualise the pay (A = P × (I − F5A)), work out a year's
 * tax with every credit CRA's formula allows, divide by P.
 *
 *   I    gross pay this period         P   pay periods per year
 *   C    CPP base+first-add this period (for F5A deduction and K2 credit)
 *   EI   EI premium this period         PM  months of CPP contributions
 *   tcFed / tcOn  TD1 / TD1ON total claim amounts
 *
 * Returns { fed, on, math } per period (unrounded→rounded to cents).
 */
export function periodIncomeTax({ I, P, C = 0, C2 = 0, EI = 0, tcFed, tcOn, PM = 12 }) {
  const R = RATES_2026;
  // F5A — deduction for the enhanced (first + second additional) CPP
  const F5 = C * (R.cppFirstAdditionalRate / R.cppRate) + C2;
  const A = Math.max(0, P * (I - F5));
  const math = { annualIncome: round2(A), periodsPerYear: P };
  if (A <= 0) return { fed: 0, on: 0, math };

  // CPP base-portion and EI amounts that earn the K2 / K2P credits
  const cppCreditBase = Math.min(P * C * (R.cppBaseRate / R.cppRate), R.cppBaseMaxEe * (PM / 12));
  const eiCreditBase = Math.min(P * EI, R.eiMaxEe);

  // ── Federal ──
  const lowF = R.fedBrackets[0].rate;
  const { R: Rf, K: Kf } = rateAndConstant(A, R.fedBrackets);
  const K1 = lowF * tcFed;
  const K2 = lowF * cppCreditBase + lowF * eiCreditBase;
  const K4 = Math.min(lowF * A, lowF * R.canadaEmploymentAmount);
  const T1 = Math.max(0, Rf * A - Kf - K1 - K2 - K4);

  // ── Ontario ──
  const lowO = R.onBrackets[0].rate;
  const { R: Vo, K: KPo } = rateAndConstant(A, R.onBrackets);
  const K1P = lowO * tcOn;
  const K2P = lowO * cppCreditBase + lowO * eiCreditBase;
  const T4 = Math.max(0, Vo * A - KPo - K1P - K2P);
  let V1 = 0;
  if (T4 > R.onSurtax1Threshold) V1 += R.onSurtax1Rate * (T4 - R.onSurtax1Threshold);
  if (T4 > R.onSurtax2Threshold) V1 += R.onSurtax2Rate * (T4 - R.onSurtax2Threshold);
  const V2 = ontarioHealthPremium(A);
  const base = T4 + V1;
  const S = Math.max(0, Math.min(base, 2 * R.onLowIncomeReductionBase - base));
  const T2 = Math.max(0, T4 + V1 + V2 - S);

  Object.assign(math, {
    fedAnnual: round2(T1), onAnnual: round2(T2), healthPremium: round2(V2), onReduction: round2(S),
    credits: { basicFed: round2(K1), cppEiFed: round2(K2), employment: round2(K4), basicOn: round2(K1P), cppEiOn: round2(K2P) },
  });
  return { fed: round2(T1 / P), on: round2(T2 / P), math };
}

// ─── Top-level calculator ───────────────────────────────────────────

/**
 * Full pay-run calculation. Returns the deductions to apply THIS run and a
 * breakdown for display/audit. All values rounded to cents.
 *
 * input = {
 *   employee: { dob, relationship?, familyEiExempt, td1FedClaim?, td1OnClaim? },
 *   payDate: 'YYYY-MM-DD',
 *   grossPay: number,
 *   ytd: {
 *     gross: number,       // YTD gross BEFORE this run
 *     cppBase: number,     // YTD base CPP withheld BEFORE this run
 *     cpp2: number,
 *     fedTax: number,
 *     onTax: number,
 *   },
 * }
 *
 * output = { age, gross, cpp, cpp2, ei, fedTax, onTax, netPay,
 *            employerCpp, employerEi, employerShare, totalRemittance, employerCost,
 *            flags, breakdown }
 *
 * EMPLOYER SHARE: the corporation matches every dollar of CPP/CPP2 it
 * withholds and pays 1.4× the employee's EI. Both go to CRA in the SAME
 * remittance as the employee deductions. totalRemittance is the cheque
 * CRA expects; employerCost is what the corp spends on top of gross.
 */
export function calculatePayRun({ employee, payDate, grossPay, ytd, periodStart, periodEnd }) {
  const gross = Math.max(0, Number(grossPay) || 0);
  const ytdGross = Math.max(0, Number(ytd?.gross) || 0);
  const ytdCppBase = Math.max(0, Number(ytd?.cppBase) || 0);
  const ytdCpp2 = Math.max(0, Number(ytd?.cpp2) || 0);
  const ytdFedTax = Math.max(0, Number(ytd?.fedTax) || 0);
  const ytdOnTax = Math.max(0, Number(ytd?.onTax) || 0);

  const age = ageOnDate(employee?.dob, payDate);
  const isUnder18 = age != null && age < 18;
  const cppExempt = !cppApplies(employee?.dob, payDate);  // CRA: from the month after 18th birthday to the month after 70th
  const familyEiExempt = employee?.familyEiExempt !== false;  // default true for family

  const payFrequency = normalizePayFrequency(employee?.payFrequency);
  const P = PAY_FREQUENCIES[payFrequency];
  const N = (periodStart && periodEnd) ? periodsCovered(periodStart, periodEnd, P) : 1;
  const { cppBase, cpp2, math: cppMath } = calculateCpp(gross, ytdGross, ytdCppBase, ytdCpp2, cppExempt, P, N);
  const ei = familyEiExempt ? 0 : round2(Math.min(gross * RATES_2026.eiRateEe, RATES_2026.eiMie * RATES_2026.eiRateEe));
  // TD1 claim codes from the employee record (default 1 = basic amount,
  // which is also CRA's rule when no TD1 is on file). 0 = no claim here.
  const td1Fed = employee?.td1FedClaim ?? 1;
  const td1On = employee?.td1OnClaim ?? 1;
  // Income tax — CRA T4127 Option 1, the method CRA's calculator uses.
  // A catch-up run covering N regular periods is taxed as N regular pays
  // (the same spreading the CPP exemption uses), so a six-month catch-up
  // is not annualised as if it were one month's pay × 12.
  const tcFed = claimAmountForCode(td1Fed, RATES_2026.fedBpa, RATES_2026.fedClaimStep);
  const tcOn = claimAmountForCode(td1On, RATES_2026.onBpa, RATES_2026.onClaimStep);
  const perPeriod = periodIncomeTax({
    I: gross / N, P, C: cppBase / N, C2: cpp2 / N, EI: ei / N, tcFed, tcOn,
  });
  // TD1 / TD1ON box "my total income for the year will be less than my
  // total claim amount" — CRA: the employer then deducts NO income tax.
  // (Typical for a part-time student; CPP still applies if 18+.)
  const incomeBelowClaim = !!employee?.td1IncomeBelowClaim;
  const fedTax = incomeBelowClaim ? 0 : round2(perPeriod.fed * N);
  const onTax = incomeBelowClaim ? 0 : round2(perPeriod.on * N);

  const totalDeductions = round2(cppBase + cpp2 + ei + fedTax + onTax);
  const netPay = round2(gross - totalDeductions);

  // Employer side — CRA T4001: employer CPP = employee CPP (incl. CPP2),
  // employer EI = 1.4 × employee EI (standard rate; reduced-rate employers N/A).
  const employerCpp = round2(cppBase + cpp2);
  const employerEi = round2(ei * RATES_2026.eiEmployerMultiplier);
  const employerShare = round2(employerCpp + employerEi);
  const totalRemittance = round2(totalDeductions + employerShare);
  const employerCost = round2(gross + employerShare);

  return {
    age,
    gross: round2(gross),
    cpp: cppBase,
    cpp2,
    ei,
    fedTax,
    onTax,
    totalDeductions,
    netPay,
    employerCpp,
    employerEi,
    employerShare,
    totalRemittance,
    employerCost,
    flags: { isUnder18, cppExempt, familyEiExempt, td1Fed: Number(td1Fed), td1On: Number(td1On), incomeBelowClaim },
    breakdown: {
      payFrequency,
      cpp: cppMath,
      tax: { ...perPeriod.math, periodsCovered: N, incomeBelowClaim, method: 'CRA T4127 Option 1 (PDOC)' },
      ytdGrossAfterRun: round2(ytdGross + gross),
      appliedFedBpa: tcFed,
      appliedOnBpa: tcOn,
    },
  };
}

// ─── Recovering CPP that was under-deducted earlier in the year ─────

/**
 * CRA ("Make corrections before filing"): the employer remits BOTH shares
 * of any CPP shortfall, and may recover the employee's share from later
 * pays in the same year — but the amount recovered from a pay can be
 * "equal to, but not more than" what should be deducted from that pay.
 * It can't be taken out of income tax, and not after 12 months.
 *
 * Mutates and returns `result` (from calculatePayRun): adds cppRecovery,
 * raises the employee deductions and the CRA total, lowers net pay. The
 * employer CPP stays the normal match — its share was already counted
 * in full on the run that came up short.
 */
export function applyCppRecovery(result, outstanding) {
  const normal = round2((result.cpp || 0) + (result.cpp2 || 0));
  const recovery = round2(Math.max(0, Math.min(Number(outstanding) || 0, normal)));
  result.cppRecovery = recovery;
  result.cppOutstandingBefore = round2(Math.max(0, Number(outstanding) || 0));
  result.cppOutstandingAfter = round2(Math.max(0, result.cppOutstandingBefore - recovery));
  if (recovery > 0) {
    result.totalDeductions = round2(result.totalDeductions + recovery);
    result.netPay = round2(result.netPay - recovery);
    result.totalRemittance = round2(result.totalRemittance + recovery);
  }
  return result;
}

// ─── Remittance due date ────────────────────────────────────────────

/**
 * CRA source-deduction due date for a pay issued on `payDateIso`.
 *
 *   remitter = 'monthly'   (regular remitter, CRA default for new accounts)
 *     → 15th of the month AFTER the pay month.   Feb 23 → Mar 15
 *   remitter = 'quarterly' (CRA-assigned; AMWA < $3,000 + clean history,
 *                           or a new small employer who requested it)
 *     → 15th of the month after the calendar quarter.  Oct 5 → Jan 15
 *
 * The remitter type is a property of the RP account, assigned by CRA —
 * NOT the HST filing frequency. If the 15th is a weekend/holiday CRA
 * accepts the next business day; we return the 15th.
 */
export function remittanceDueDate(payDateIso, remitter = 'monthly') {
  const d = new Date(payDateIso);
  if (isNaN(d.getTime())) return null;
  const y = d.getUTCFullYear();
  const m0 = d.getUTCMonth();                       // 0-indexed pay month
  let dueM0;                                        // 0-indexed month that contains the 15th
  if (remitter === 'quarterly') {
    dueM0 = (Math.floor(m0 / 3) + 1) * 3;           // month after quarter end: 3, 6, 9, 12
  } else {
    dueM0 = m0 + 1;
  }
  const due = new Date(Date.UTC(y, dueM0, 15));     // Date.UTC rolls month 12 → Jan next year
  return due.toISOString().slice(0, 10);
}

export function normalizeRemitter(v) {
  return String(v || '').toLowerCase() === 'quarterly' ? 'quarterly' : 'monthly';
}

// ─── Small helpers ──────────────────────────────────────────────────

function round2(n) {
  const x = Number(n) || 0;
  return Math.round(Number((x * 100).toFixed(6))) / 100;
}
