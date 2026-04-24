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
  eiMie: 68900,

  // Federal basic personal amount (linear phase-down above ~$177K — for
  // Caleb's kids earning well under that, we use the full amount).
  fedBpa: 16452,

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

// ─── Income tax ─────────────────────────────────────────────────────

/**
 * Federal tax owed THIS pay run using cumulative-YTD method.
 *
 * Applies the under-18 supplement to raise the effective BPA for minors.
 * Under-BPA income owes zero — the method handles irregular/lumpy pay well
 * because tax is always `annual-tax-on-YTD-projection − YTD-already-paid`.
 */
export function calculateFederalTax(grossPay, ytdGross, ytdFedTaxPaid, isUnder18) {
  const R = RATES_2026;
  const taxable = ytdGross + grossPay;
  if (taxable <= 0) return 0;

  const effectiveBpa = R.fedBpa + (isUnder18 ? R.fedUnder18Supplement : 0);
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
export function calculateOntarioTax(grossPay, ytdGross, ytdOnTaxPaid, isUnder18) {
  const R = RATES_2026;
  const taxable = ytdGross + grossPay;
  if (taxable <= 0) return 0;

  const effectiveBpa = R.onBpa + (isUnder18 ? R.onUnder18Supplement : 0);
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
 *   employee: { dob, relationship?, familyEiExempt },
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
 * output = { age, gross, cpp, cpp2, ei, fedTax, onTax, netPay, flags, breakdown }
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
  const fedTax = calculateFederalTax(gross, ytdGross, ytdFedTax, isUnder18);
  const onTax = calculateOntarioTax(gross, ytdGross, ytdOnTax, isUnder18);

  const totalDeductions = round2(cppBase + cpp2 + ei + fedTax + onTax);
  const netPay = round2(gross - totalDeductions);

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
    flags: { isUnder18, cppExempt, familyEiExempt },
    breakdown: {
      ytdGrossAfterRun: round2(ytdGross + gross),
      appliedFedBpa: RATES_2026.fedBpa + (isUnder18 ? RATES_2026.fedUnder18Supplement : 0),
      appliedOnBpa: RATES_2026.onBpa + (isUnder18 ? RATES_2026.onUnder18Supplement : 0),
    },
  };
}

// ─── Remittance due date (monthly — 15th of following month) ────────

/**
 * For a pay run on `payDateIso`, return the CRA source deduction
 * remittance due date (ISO YYYY-MM-DD). Monthly remitter default:
 * due the 15th of the month AFTER the month in which pay was issued.
 *
 * E.g., pay on Feb 23 → remit by Mar 15.
 * If the 15th falls on a weekend/holiday, CRA allows the next business
 * day — for simplicity we return the 15th; UI can show the business-day
 * adjustment if needed.
 */
export function remittanceDueDate(payDateIso) {
  const d = new Date(payDateIso);
  if (isNaN(d.getTime())) return null;
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + 1;  // 0-indexed → month after pay month
  const yNext = m === 12 ? y + 1 : y;
  const mNext = m === 12 ? 0 : m;
  const due = new Date(Date.UTC(yNext, mNext, 15));
  return due.toISOString().slice(0, 10);
}

// ─── Small helpers ──────────────────────────────────────────────────

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}
