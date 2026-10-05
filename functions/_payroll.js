// ════════════════════════════════════════════════════════════════════
// PAYROLL DEDUCTION ENGINE — 2026 CRA rates, age-stratified
//
// Pure functions. No IO, no side effects. The pay-run endpoints
// (calculate / run / t4) all share this single source of truth so the
// rates can never drift between preview and commit.
//
// Scope: Postma Contracting Inc. family payroll. Six kids ages 8–20.
// - Under 18: CPP exempt, family-EI exempt, under-18 tax supplement applies.
// - 18+:      CPP 5.95% above $3,500 basic exemption, family-EI exempt.
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

  // TD1 claim-code chart step (T4008-ON Jan 2026, Chart 1): code 1 = BPA,
  // code 2 = 16,452.01–19,285, … each code $2,833 wide. The formula uses
  // the mid-point of the range: BPA + (code − 1.5) × step.
  fedClaimStep: 2833,

  // Under-18 federal supplement (non-refundable credit, reduced by
  // childcare/attendant claims above $3,533 — N/A for family payroll).
  fedUnder18Supplement: 844,
  under18SupplementReductionThreshold: 3533,

  // Federal tax brackets — 2026 (post the 14% reduction in effect full-year).
  fedBrackets: [
    { upTo: 58523,  rate: 0.14 },
    { upTo: 117045, rate: 0.205 },
    { upTo: 181440, rate: 0.26 },
    { upTo: 258482, rate: 0.29 },
    { upTo: Infinity, rate: 0.33 },
  ],

  // Ontario BPA and under-18 supplement
  onBpa: 12989,
  onUnder18Supplement: 482,
  // TD1ON claim-code chart step (T4008-ON Jan 2026, Chart 2): code 2 =
  // 12,989.01–15,787, each code $2,798 wide.
  onClaimStep: 2798,

  // Ontario tax brackets — 2026 (indexed 1.019, except $150K/$220K which
  // are fixed by statute and not indexed)
  onBrackets: [
    { upTo: 53891,  rate: 0.0505 },
    { upTo: 107785, rate: 0.0915 },
    { upTo: 150000, rate: 0.1116 },
    { upTo: 220000, rate: 0.1216 },
    { upTo: Infinity, rate: 0.1316 },
  ],

  // Ontario low-income tax reduction — $300 base reduction, phases out
  // linearly as basic ON tax rises from $300 to $600. Effect: a resident
  // with taxable income up to ~$18,930 owes zero ON tax; above ~$24,870
  // the reduction is exhausted.
  onLowIncomeReductionBase: 300,
};

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
 * CPP base + CPP2 employee portion for a pay run, using annual cumulative method.
 *
 * Inputs:
 *   grossPay        - this pay run's gross, pre-deduction
 *   ytdPensionable  - YTD pensionable earnings BEFORE this pay run (0 if first run of year)
 *   ytdCppPaid      - YTD CPP (base) already withheld BEFORE this pay run
 *   ytdCpp2Paid     - YTD CPP2 already withheld BEFORE this pay run
 *
 * Returns {cppBase, cpp2, note} — amount to withhold THIS run.
 *
 * Under 18 or CPP-exempt: returns zeros.
 */
export function calculateCpp(grossPay, ytdPensionable, ytdCppPaid, ytdCpp2Paid, cppExempt = false) {
  if (cppExempt || grossPay <= 0) return { cppBase: 0, cpp2: 0 };

  const R = RATES_2026;
  const ytdAfter = ytdPensionable + grossPay;

  // CPP base (0 → YMPE, with $3,500 annual exemption allocated in the same
  // cumulative fashion — CRA method requires per-period exemption but for
  // irregular-pay kid runs the simpler annual approach is practically
  // equivalent and under-18s are exempt anyway).
  const cappedBaseAfter = Math.min(ytdAfter, R.cppYmpe);
  const pensionableAfter = Math.max(0, cappedBaseAfter - R.cppBasicExemption);
  const cumulativeCppBase = pensionableAfter * R.cppRate;
  const cppBaseThisRun = Math.max(0, Math.min(cumulativeCppBase, R.cppMaxAnnualEe) - ytdCppPaid);

  // CPP2 (YMPE → YAMPE @ 4%)
  const cpp2CappedAfter = Math.min(Math.max(0, ytdAfter - R.cppYmpe), R.cpp2Yampe - R.cppYmpe);
  const cumulativeCpp2 = cpp2CappedAfter * R.cpp2Rate;
  const cpp2ThisRun = Math.max(0, Math.min(cumulativeCpp2, R.cpp2MaxAnnualEe) - ytdCpp2Paid);

  return { cppBase: round2(cppBaseThisRun), cpp2: round2(cpp2ThisRun) };
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

// ─── Income tax ─────────────────────────────────────────────────────

/**
 * Federal tax owed THIS pay run using cumulative-YTD method.
 *
 * Applies the under-18 supplement to raise the effective BPA for minors.
 * Under-BPA income owes zero — the method handles irregular/lumpy pay well
 * because tax is always `annual-tax-on-YTD-projection − YTD-already-paid`.
 */
export function calculateFederalTax(grossPay, ytdGross, ytdFedTaxPaid, isUnder18, claimCode = 1) {
  const R = RATES_2026;
  const taxable = ytdGross + grossPay;
  if (taxable <= 0) return 0;

  const effectiveBpa = claimAmountForCode(claimCode, R.fedBpa, R.fedClaimStep) + (isUnder18 ? R.fedUnder18Supplement : 0);
  const bpaCredit = Math.min(taxable, effectiveBpa) * R.fedBrackets[0].rate;

  const basicTax = bracketTax(taxable, R.fedBrackets);
  const annualTax = Math.max(0, basicTax - bpaCredit);

  const thisRunTax = Math.max(0, annualTax - ytdFedTaxPaid);
  return round2(thisRunTax);
}

/**
 * Ontario tax THIS pay run, including under-18 supplement and the
 * low-income tax reduction.
 */
export function calculateOntarioTax(grossPay, ytdGross, ytdOnTaxPaid, isUnder18, claimCode = 1) {
  const R = RATES_2026;
  const taxable = ytdGross + grossPay;
  if (taxable <= 0) return 0;

  const effectiveBpa = claimAmountForCode(claimCode, R.onBpa, R.onClaimStep) + (isUnder18 ? R.onUnder18Supplement : 0);
  const bpaCredit = Math.min(taxable, effectiveBpa) * R.onBrackets[0].rate;

  const basicOnTax = bracketTax(taxable, R.onBrackets);
  const onTaxAfterBpa = Math.max(0, basicOnTax - bpaCredit);

  // Low-income reduction: $300 full relief up to $300 of basic tax; linear
  // phase-out to $0 reduction at $600 basic tax.
  let reduction;
  if (onTaxAfterBpa <= R.onLowIncomeReductionBase) {
    reduction = onTaxAfterBpa;
  } else if (onTaxAfterBpa >= R.onLowIncomeReductionBase * 2) {
    reduction = 0;
  } else {
    reduction = R.onLowIncomeReductionBase * 2 - onTaxAfterBpa;
  }
  const annualOnTax = Math.max(0, onTaxAfterBpa - reduction);

  const thisRunTax = Math.max(0, annualOnTax - ytdOnTaxPaid);
  return round2(thisRunTax);
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
export function calculatePayRun({ employee, payDate, grossPay, ytd }) {
  const gross = Math.max(0, Number(grossPay) || 0);
  const ytdGross = Math.max(0, Number(ytd?.gross) || 0);
  const ytdCppBase = Math.max(0, Number(ytd?.cppBase) || 0);
  const ytdCpp2 = Math.max(0, Number(ytd?.cpp2) || 0);
  const ytdFedTax = Math.max(0, Number(ytd?.fedTax) || 0);
  const ytdOnTax = Math.max(0, Number(ytd?.onTax) || 0);

  const age = ageOnDate(employee?.dob, payDate);
  const isUnder18 = age != null && age < 18;
  const cppExempt = isUnder18;  // Under-18 are CPP exempt under CRA rules
  const familyEiExempt = employee?.familyEiExempt !== false;  // default true for family

  const { cppBase, cpp2 } = calculateCpp(gross, ytdGross, ytdCppBase, ytdCpp2, cppExempt);
  const ei = familyEiExempt ? 0 : round2(Math.min(gross * RATES_2026.eiRateEe, RATES_2026.eiMie * RATES_2026.eiRateEe));
  // TD1 claim codes from the employee record (default 1 = basic amount,
  // which is also CRA's rule when no TD1 is on file). 0 = no claim here.
  const td1Fed = employee?.td1FedClaim ?? 1;
  const td1On = employee?.td1OnClaim ?? 1;
  const fedTax = calculateFederalTax(gross, ytdGross, ytdFedTax, isUnder18, td1Fed);
  const onTax = calculateOntarioTax(gross, ytdGross, ytdOnTax, isUnder18, td1On);

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
    flags: { isUnder18, cppExempt, familyEiExempt, td1Fed: Number(td1Fed), td1On: Number(td1On) },
    breakdown: {
      ytdGrossAfterRun: round2(ytdGross + gross),
      appliedFedBpa: claimAmountForCode(td1Fed, RATES_2026.fedBpa, RATES_2026.fedClaimStep) + (isUnder18 ? RATES_2026.fedUnder18Supplement : 0),
      appliedOnBpa: claimAmountForCode(td1On, RATES_2026.onBpa, RATES_2026.onClaimStep) + (isUnder18 ? RATES_2026.onUnder18Supplement : 0),
    },
  };
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
  return Math.round((Number(n) || 0) * 100) / 100;
}
