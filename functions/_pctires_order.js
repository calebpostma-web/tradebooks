// ════════════════════════════════════════════════════════════════════
// pctires.ca "New Order" email → structured order
//
// The site emails orders@pctires.ca → Caleb + Andrea with one email per
// order. It already carries both sides of the transaction:
//   • what the customer was charged (tires, install, HST, total, deposit)
//   • what PC Tires owes TDG (subtotal, fees, HST, SO#)
// This module turns that email into the invoice row, the payments list and
// the expected TDG bill. Pure function; no IO. Works from the HTML body
// (preferred — label/value table cells) or the plain-text fallback.
// ════════════════════════════════════════════════════════════════════

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', mdash: '—', ndash: '–', times: '×', middot: '·', minus: '−', rsquo: '’', lsquo: '‘', hellip: '…' };
function decode(s) {
  return String(s || '')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(+n))
    .replace(/&#x([0-9a-f]+);/gi, (_, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&([a-z]+);/gi, (m, e) => ENT[e.toLowerCase()] ?? m);
}

/** HTML → "label\tvalue" lines; divs/rows become newlines. */
export function htmlToLines(html) {
  let s = String(html || '');
  s = s.replace(/<style[\s\S]*?<\/style>/gi, '').replace(/<!--[\s\S]*?-->/g, '');
  s = s.replace(/<\/td>\s*<td[^>]*>/gi, '\t');          // cell boundary → tab
  s = s.replace(/<\/(tr|div|p|h\d|li)>/gi, '\n').replace(/<br\s*\/?>/gi, '\n');
  s = s.replace(/<[^>]+>/g, '');
  s = decode(s);
  return s.split('\n').map(l => l.replace(/[  ]+/g, ' ').replace(/\s*\t\s*/g, '\t').trim()).filter(Boolean);
}

const money = (v) => {
  if (v == null) return null;
  const m = /-?\$?\s*(-?[\d,]+\.\d{2})/.exec(String(v));
  if (!m) return null;
  const n = parseFloat(m[1].replace(/,/g, ''));
  return /^\s*-|\(.*\)/.test(String(v)) ? -Math.abs(n) : n;
};

const MONTHS = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,sept:9,oct:10,nov:11,dec:12 };
function isoFromLong(s) {
  // "Mon, Oct 19, 2026, 8:11 a.m."  |  "October 21, 2026 at 10:00 AM"
  const m = /([A-Za-z]{3,9})\.?\s+(\d{1,2}),\s*(\d{4})/.exec(String(s || ''));
  if (!m) return '';
  const mo = MONTHS[m[1].slice(0, 3).toLowerCase()]; if (!mo) return '';
  return `${m[3]}-${String(mo).padStart(2, '0')}-${String(m[2]).padStart(2, '0')}`;
}

/** Find the value for a label in label\tvalue lines (first match). */
function pick(lines, label, { startsWith = false } = {}) {
  const L = label.toLowerCase();
  for (const line of lines) {
    const [k, ...rest] = line.split('\t');
    if (rest.length === 0) continue;
    const key = k.trim().toLowerCase();
    if (key === L || (startsWith && key.startsWith(L))) return rest.join('\t').trim();
  }
  return null;
}
function pickLine(lines, re) { for (const l of lines) { const m = re.exec(l); if (m) return m; } return null; }

// Plain-text fallback: the text part glues labels to values ("NameJane
// Smith Emailjane…"). Re-insert structure by splitting on the known labels.
const TEXT_LABELS = [
  'Order date', 'Name', 'Email', 'Phone', 'Vehicle', 'Lug Spec',
  'Subtotal (tires)', 'Installation', 'Subtotal before tax', 'HST (13%)', 'Order Total', 'Total Charged',
  'Deposit Paid (Stripe)', 'BALANCE OWING AT INSTALL', 'UNPAID — COLLECT AT PICKUP',
  'TDG Order #', 'Reference', 'Subtotal', 'Fees', 'HST (TDG paid)', 'Total Cost', 'Pre-tax margin', 'Stripe fee', 'Net Margin',
  'Install', 'Payment', 'Search Method', 'CASL Opt-in',
];
const LOWER_OK = new Set(['Search Method', 'Stripe fee', 'Payment', 'Email']);   // labels whose values may start lowercase
export function textToLines(text) {
  let s = String(text || '').replace(/\r/g, '').replace(/ tel:\d+/g, '');
  // Section headings first, as placeholders, so "Install" can't eat "Install & Source"
  const HEAD = ['PC Tires Cost (TDG)', 'Customer Charged', 'Install & Source', 'Items', 'Customer'];
  HEAD.forEach((h, i) => { s = s.replace(new RegExp(h.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=\\s|$)', 'g'), `\n§H${i}§\n`); });
  // longest labels first so "Subtotal (tires)" wins over "Subtotal", "Installation" over "Install"
  const labels = [...TEXT_LABELS].sort((a, b) => b.length - a.length);
  for (const L of labels) {
    const tail = LOWER_OK.has(L) ? '' : '(?!\\s*[a-z(])';
    const re = new RegExp('(^|[\\s])' + L.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + tail + '\\s*', 'g');
    s = s.replace(re, (m, pre) => `${pre}\n${L}\t`);
  }
  HEAD.forEach((h, i) => { s = s.split(`§H${i}§`).join(h); });
  return s.split('\n').map(l => l.replace(/\s*\t\s*/g, '\t').trim()).filter(Boolean);
}

export function parseOrderEmail({ html, text, subject = '' }) {
  const lines = html ? htmlToLines(html) : textToLines(text);
  const joined = lines.join('\n');

  const invNum = (/PCT-[A-Z0-9]{4,}/.exec(subject) || /New Order\s*[—-]\s*(PCT-[A-Z0-9]+)/i.exec(joined) || [])[0] || null;
  const pctMatch = /PCT-[A-Z0-9]{4,}/.exec(invNum || '') || /PCT-[A-Z0-9]{4,}/.exec(joined);
  if (!pctMatch) return { ok: false, error: 'No PCT order number found' };
  const orderNo = pctMatch[0];

  const orderDateRaw = pick(lines, 'Order date') || (pickLine(lines, /^[A-Z][a-z]{2}, [A-Z][a-z]{2} \d{1,2}, \d{4}/) || [])[0] || '';
  const customer = {
    name: pick(lines, 'Name') || '',
    email: pick(lines, 'Email') || '',
    phone: (pick(lines, 'Phone') || '').replace(/\D/g, ''),
    vehicle: (pick(lines, 'Vehicle') || '').replace(/^-$/, ''),
  };

  // Items: "4× Antares SMT A7 (2656018, P265/60R18 XL)"
  const items = [];
  for (const l of lines) {
    const m = /^(\d+)\s*[×x]\s*(.+?)\s*\(([^,()]+),\s*([^()]+)\)\s*$/.exec(l);
    if (m) items.push({ qty: +m[1], name: m[2].trim(), sku: m[3].trim(), size: m[4].trim() });
  }
  const addons = (pick(lines, 'Add-ons:') || pick(lines, 'Add-ons') || (pickLine(lines, /^Add-ons:\s*(.+)$/) || [])[1] || '').trim();

  // Customer charged
  const subtotalTires = money(pick(lines, 'Subtotal (tires)'));
  const installation = money(pick(lines, 'Installation'));
  const subtotalBeforeTax = money(pick(lines, 'Subtotal before tax'));
  const hst = money(pick(lines, 'HST (13%)') ?? pick(lines, 'HST', { startsWith: true }));
  const orderTotal = money(pick(lines, 'Order Total') ?? pick(lines, 'Total Charged'));
  const depositPaid = money(pick(lines, 'Deposit Paid (Stripe)') ?? pick(lines, 'Deposit Paid', { startsWith: true }));
  const balanceOwing = money(pick(lines, 'BALANCE OWING AT INSTALL') ?? pick(lines, 'Balance owing', { startsWith: true }));
  const unpaid = money(pick(lines, 'UNPAID — COLLECT AT PICKUP') ?? pick(lines, 'UNPAID', { startsWith: true }));
  const paidInFull = pick(lines, 'Total Charged') != null && depositPaid == null && unpaid == null;

  // TDG cost block — labels repeat ("Subtotal", "HST") so slice after the heading
  const tdgStart = lines.findIndex(l => /^PC Tires Cost \(TDG\)/i.test(l));
  const tdgLines = tdgStart >= 0 ? lines.slice(tdgStart) : [];
  const tdg = {
    orderNo: pick(tdgLines, 'TDG Order #') || (/TDG Order placed:\s*([\d-]+)/.exec(joined) || [])[1] || '',
    reference: pick(tdgLines, 'Reference') || '',
    subtotal: money(pick(tdgLines, 'Subtotal')),
    fees: money(pick(tdgLines, 'Fees')) || 0,
    hst: money(pick(tdgLines, 'HST (TDG paid)')),
    totalCost: money(pick(tdgLines, 'Total Cost')),
  };
  const stripeFeeRaw = pick(tdgLines, 'Stripe fee', { startsWith: true }) || '';
  const stripeFee = { amount: Math.abs(money(stripeFeeRaw) || 0), estimated: /EST/.test(pick(tdgLines, 'Stripe fee', { startsWith: true }) || '') || /Stripe fee[^\t\n]*EST/.test(joined), raw: stripeFeeRaw };

  // Install & source
  const installRaw = pick(lines, 'Install') || '';
  const payment = pick(lines, 'Payment') || '';
  const searchMethod = pick(lines, 'Search Method') || '';
  const phoneOrder = /PHONE ORDER/i.test(searchMethod) || /PHONE ORDER/i.test(joined);

  const subtotal = subtotalBeforeTax ?? ((subtotalTires || 0) + (installation || 0));
  const total = orderTotal ?? (subtotal + (hst || 0));

  // Payments the site already knows about
  const payments = [];
  const orderDate = isoFromLong(orderDateRaw);
  if (depositPaid > 0) payments.push({ date: orderDate, method: 'card', amount: depositPaid, kind: 'deposit', pending: true, via: 'stripe', fee: stripeFee.amount, feeEstimated: stripeFee.estimated });
  else if (paidInFull && total > 0) payments.push({ date: orderDate, method: 'card', amount: total, kind: 'final', pending: true, via: /affirm/i.test(payment + ' ' + stripeFee.raw) ? 'affirm' : 'stripe', fee: stripeFee.amount, feeEstimated: stripeFee.estimated });

  // Invoice lines as the wizard stores them
  const lines_ = items.map(i => ({ desc: `${i.name} ${i.size} (${i.sku})`, qty: i.qty, rate: items.length === 1 && subtotalTires != null ? Math.round(subtotalTires / i.qty * 100) / 100 : null }));
  if (lines_.length === 1 && lines_[0].rate == null && subtotalTires != null) lines_[0].rate = Math.round(subtotalTires / lines_[0].qty * 100) / 100;
  if (lines_.length > 1 && subtotalTires != null) {
    // multiple tire lines: we don't know the split; one line for all tires
    lines_.splice(0, lines_.length, { desc: items.map(i => `${i.qty}× ${i.name} ${i.size}`).join('; '), qty: 1, rate: subtotalTires });
  }
  if (installation > 0) lines_.push({ desc: `Installation${installRaw ? ' — ' + installRaw.replace(/^.*?—\s*/, '') : ''}`, qty: 1, rate: installation });
  if (addons && !/^none$/i.test(addons)) lines_.push({ desc: `Add-ons: ${addons}`, qty: 1, rate: 0 });

  return {
    ok: true,
    orderNo, orderDate, orderDateRaw,
    customer, items, addons, installRaw, installDate: isoFromLong(installRaw), payment, searchMethod, phoneOrder,
    charged: { subtotalTires, installation, subtotal, hst, total, depositPaid, balanceOwing, unpaid, paidInFull },
    tdg, stripeFee,
    status: paidInFull ? 'Paid' : depositPaid > 0 ? 'Deposit Received' : 'Unpaid',
    invoice: {
      invNum: orderNo, dateVal: orderDate, client: customer.name,
      desc: items.map(i => `${i.qty}× ${i.name} ${i.size}`).join('; ') + (installation > 0 ? '; Installation' : ''),
      sub: subtotal, hstAmt: hst, total, hst: hst > 0 ? 'Yes' : 'No',
      notes: [customer.vehicle && `Vehicle: ${customer.vehicle}`, tdg.orderNo && `TDG ${tdg.orderNo} / ${tdg.reference}`, installRaw && `Install: ${installRaw}`, customer.phone && `Ph ${customer.phone}`, phoneOrder && 'Phone order'].filter(Boolean).join(' · '),
      lines: lines_,
      payments,
    },
    expectedBill: tdg.totalCost ? {
      party: 'Tire Discounter Group', description: `TDG ${tdg.orderNo} / ${tdg.reference} — ${orderNo} (${customer.name})`,
      amount: -((tdg.subtotal || 0) + (tdg.fees || 0)), hst: tdg.hst || 0, total: -tdg.totalCost, relatedInvoice: orderNo, ref: `TDG-${tdg.reference || tdg.orderNo}`,
    } : null,
  };
}
