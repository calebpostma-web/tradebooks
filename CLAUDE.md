# tradebooks — project notes for Claude

## What this is

A free, personalized bookkeeping web app for **Caleb Postma**, a Chatham-Kent
contractor running multiple businesses (Postma Contracting Inc, an HVAC arm,
and PCTires). His wife handles day-to-day books. They use MNP as their
accountant — fiscal year ends **March 31** (incorporated FY).

**The whole project exists for one reason:** make year-end and quarterly
prep so clean that MNP only needs to *review* the books, not prepare them.
Target: drop accountant fees from $3-5k/yr (full prep) to ~$500-1.5k/yr
(review-only engagement). Year-End Package builder + CRA payments log are
the load-bearing features.

## Stack

- **Front-end:** single-page app at `app/index.html` — vanilla JS, no build step. Hosted on Cloudflare Pages at https://aibookkeeper.ca
- **Back-end:** Cloudflare Pages Functions in `functions/api/...`. Each file is one route.
- **Auth:** Google sign-in (OAuth) → JWT in sessionStorage. Database = Cloudflare D1 (SQLite). Profile + Google tokens live in `profiles` table.
- **Books storage:** Google Sheets. Each user's books live in their own Google Sheet. Tradebooks acts on the sheet via Sheets API + Drive API using their OAuth refresh token.
- **Receipts / statements / CRA receipts:** uploaded to Google Drive folders (`AI Bookkeeper Receipts/`, `AI Bookkeeper CRA Remittances/`, `AI Bookkeeper Year-End/`).

## Account model — IMPORTANT

**Single-user. One Google account = one set of books.** No multi-user / team
support. Caleb and his wife share `postmacontracting@gmail.com` as the
single TradeBooks login. Don't ever suggest creating separate user accounts
for spouse/family — it creates two disconnected sets of books.

The OAuth flow uses **the same Google account** for both sign-in and Sheet
access. Mixing accounts (e.g. signing in with one email, granting Sheets
access from another) creates broken state.

## Sheet structure

A fully-built sheet has these tabs (modern emoji prefixes — older sheets may
have variants, see "Tab name compatibility" below):

- `📊 Dashboard` — top-line summary
- `🗂 Categories` — category list (default + user-custom)
- `📒 Transactions` — main ledger (cash basis), 14 cols B-N, last col is Total formula
- `🧾 Invoices` — invoice log, 17 cols B-Q, deposit columns O/P/Q
- `📋 HST Returns` — quarterly HST return workbook (C3 is FY Start, auto-detected from latest transaction)
- `📅 Year-End` — year-end summary + corp tax estimator + per-category breakdown
- `💼 Payroll` — pay runs, T4 source of truth
- `📝 Work Log` — contemporaneous work entries (CRA audit defence)
- `📑 CRA Remittances` — every payment to CRA (HST, payroll, corp tax)
- `🏦 Account Balances` — bank reconciliation. Per row: account + period + opening + closing. Sheet computes expected closing from Transactions and flags discrepancies.
- `📓 Adjusting Entries` — year-end accruals (AR, AP, prepaids). Bridges cash-basis to accrual-basis for the T2.
- `🛠 Fixed Assets` — CCA tracker. Per asset: class + cost + UCC. Computes annual CCA for Schedule 8.
- `📊 T2 Worksheet` — consolidated T2-prep view. Pulls from every other tab to produce the data MNP reviews.

Every sheet creation goes through `functions/api/google-setup.js`. Schema
upgrades for existing sheets go through `functions/api/setup/migrate.js`
(idempotent).

## Cash-basis convention

Books are CASH BASIS:
- Revenue posts when money hits the bank (not when invoice issued)
- Expenses post when money leaves the bank (not when bill received)
- HST follows the cash: deposit incl HST → HST collected at deposit time

The category `Internal Transfer` is excluded from P&L AND HST math
(formulas use `<>"Internal Transfer"` filters). Use this category for:
- Bank ↔ AMEX bill payments
- Payments to CRA (HST remittance, payroll source deductions, corp tax)
- Owner draws

## Recent feature areas (chronological)

1. **Deposit invoices** — invoice wizard supports a deposit field with
   optional date received. Sheet schema cols O/P/Q track Deposit Amount /
   Deposit Date / Balance Due. Bank-match logic is deposit-aware (status
   transitions Unpaid → Deposit Received → Paid).

2. **CRA Remittance Log** — `📑 CRA Remittances` tab + `/api/remittance/log`
   endpoint. UI under CRA Tax Filing → Payments Log. Each payment writes
   both a Remittances row AND a matching Internal Transfer transaction so
   bank reconciles cleanly. Supports PDF receipt upload to Drive.

3. **Year-End Package** (`📦 Year-End` tab):
   - Live checklist: transactions categorized, statements uploaded, CRA receipts logged, payroll, etc.
   - Statement archive: monthly bank/AMEX PDF drop zone
   - "Build Package" button: creates `Postma_YearEnd_FY{YYYY}_{date}/` folder in Drive with cover letter, books snapshot XLSX, statement shortcuts, CRA receipt copies, expense receipt shortcuts. One shareable link to email MNP.

4. **Sheet migration system** — `functions/api/setup/migrate.js` is
   idempotent. Each migration checks state then applies only what's missing.
   Triggered from Settings ("🔄 Update sheet to latest schema") and
   auto-detected via banner on app load (`?dryRun=1` mode).

5. **Diagnostic endpoint** — `/api/debug/whoami` returns the server-side
   state for the current JWT user (user_id, google_sub linked, refresh_token
   present, sheet_id, plain-language diagnosis). Hit from browser console:
   ```javascript
   fetch('/api/debug/whoami', {headers: {'Authorization': 'Bearer ' + sessionStorage.getItem('tradebooks_session')}}).then(r => r.json()).then(d => console.log(JSON.stringify(d, null, 2)))
   ```

## Review-only roadmap — STATUS

To unlock MNP review-only engagement (the north-star goal), three pieces were
needed. As of 2026-04-25, all three shipped:

1. ✅ **Year-end accrual entries module** — `📓 Adjusting Entries` tab + `/api/accrual/log` endpoint + UI on Year-End Package tab. AR / AP / Prepaids / Accruals.
2. ✅ **CCA tracker** — `🛠 Fixed Assets` tab with pre-loaded CCA classes (8/10/12/50/etc.). Computes UCC + annual CCA via formulas. Half-year rule applied to Additions.
3. ✅ **T2 worksheet** — `📊 T2 Worksheet` tab pulls everything together: Schedule 125 (Income Statement w/ GIFI codes), Schedule 1 (book-to-tax adjustments), Schedule 8 (CCA summary), Schedule 100 (rough Balance Sheet), tax estimate. THIS is the document MNP reviews.

Foundation also shipped:
- ✅ **Bank reconciliation** — `🏦 Account Balances` tab. Catches missed/duplicate/wrong-sign rows by comparing expected vs actual closing balance.
- ✅ **Per-category breakdown** on Year-End tab via dynamic QUERY pivot.

## Conventions + gotchas

### CRLF line endings
Repo is on Windows. Some files have CRLF. Don't fight it — git handles it.

### Tab name compatibility
Older sheets created during the legacy Apps Script era have different emoji
prefixes (e.g. `🧾 Transactions` instead of `📒 Transactions`). Use
`resolveTabName(sheets, 'Transactions')` from `_sheets.js` to match by
suffix regardless of prefix. Don't hard-code emoji-prefixed tab names in
read paths — only in write paths where the schema is fully known.

### OAuth state prefixes
`tradebooks_setup_` = onboarding's "create my sheet" flow.
`tradebooks_login_` = plain "Sign in with Google" from landing page.
**Both** must be checked in `initAuth` callback handler — there was a
real bug where only setup_ was handled and plain sign-in fell through to
the login screen silently. (See `app/index.html` around line 8985.)

### Profile saves and `sheet_id`
The profile UPSERT in `functions/api/profile.js` defensively does NOT
overwrite `sheet_id` or `script_url` when the payload is empty. This
prevents a save-with-stale-form from disconnecting a user's sheet.

### OAuth scopes
`GOOGLE_SCOPES` includes `auth/spreadsheets` (sheet read/write) AND
`auth/drive.file` (file uploads to folders the app created). The latter
is needed for receipt scanner, CRA receipts, statement archive, year-end
package builder. Adding a new scope requires users to re-sign-in to grant.

### Browser compatibility
- iPad / iOS Safari: works fine after the OAuth callback fix. We thought
  Safari ITP was the cause early on; it wasn't.
- Chrome on iOS uses WebKit (Apple policy) — same engine as Safari.
- Desktop browsers: all fine.

## Where Caleb's actual books live

- **Real account:** postmacontracting@gmail.com
- **Real Google Sheet:** lives in postmacontracting's Drive (created via
  manualCreateSheet in tradebooks)
- **calebpostma@gmail.com** is Caleb's personal test account — disposable
  cruft, can be deleted/ignored
- **andreamariepostma@gmail.com** is his wife's personal email — NOT used
  for tradebooks (any prior account under this email is orphaned)

## Deploy

`git push` to main → Cloudflare Pages auto-deploys → live in 1-2 min.
No build step. Static files + Functions deployed together.

## Local testing

Edit `app/index.html` directly in browser DevTools to iterate fast on
front-end changes without redeploying. Real changes must be committed
+ pushed to take effect for users.

## Status as of 2026-04-25 (end of Day 1)

**Shipped today:**
- Deposit invoice + CRA payments + year-end package features
- Sheet migration system (idempotent, auto-detected on app load via banner)
- Auth flow bug fixed (was silently dropping plain Google sign-in callback)
- OAuth scope expanded to include `drive.file`
- Diagnostic endpoint live for debugging account state (`/api/debug/whoami`)
- 4 importer bug fixes (bank tagging, income categorization, AMEX dedup, HST FY)
- Sign architecture refactor — bank statement direction = source of truth
- Per-row + bulk sign-flip UI for AI extraction errors
- Total (incl HST) column N on Transactions
- 🏦 Account Balances tab (bank reconciliation)
- Per-category breakdown on Year-End tab (dynamic QUERY)
- 📓 Adjusting Entries tab + /api/accrual/log + UI form
- 🛠 Fixed Assets tab (CCA tracker)
- 📊 T2 Worksheet tab (consolidated T2-prep view for MNP)

**Migration count:** 9 idempotent migrations in `functions/api/setup/migrate.js`. Existing users hit "Update sheet to latest schema" once and get all of it.

**The review-only roadmap is done at the structural level.** Caleb's wife can now:
1. Import bank/AMEX statements (with sign-flip safety net)
2. Reconcile against statement balances
3. Log CRA payments
4. Log year-end accruals
5. Track fixed assets / CCA
6. Generate the T2 Worksheet that MNP reviews
7. Build a year-end package with everything bundled into a Drive folder

## Next session — start here

When Caleb comes back, the playbook is:

### Step 1 — paste any bug list from his wife
She's been the first real-world user. Whatever she found (UX confusion,
errors, missing fields, "I expected X here") trumps speculative new
features. Fix what hurts before building forward.

### Step 2 — depending on what's broken or what's next

The review-only roadmap is structurally complete. Future work splits into:

**Real-world testing fixes** — whatever Caleb's wife (or future users) finds when actually using these tabs. Bugs trump new features.

**Polish / nice-to-haves:**
- **Per-FY HST tabs** (deferred from Day 1) — currently single HST Returns tab with editable C3. Could split into per-FY tabs for permanent filed-return archive.
- **Receipt scanner → Transaction link** — scanned PDF goes to Drive but isn't tied to its Transaction row.
- **Mobile UX pass** on Invoice Wizard — kitchen-table flow on phone.
- **"What is this?" tooltips** on accruals UI for non-accountants.

**Integration / enhancement:**
- **Auto-detect AR at FYE** — when user opens accruals form, pre-populate from Invoices tab where Status != Paid and Date Issued is in the FY.
- **Pre-populate AP from common bills** (e.g. recent utility patterns).
- **Snapshot HST returns when filing** — preserve filed numbers.

**Productization (if going beyond Caleb's family):**
- **Google OAuth verification** — weeks-long process, lots of paperwork.
- **Multi-user / team support** — currently single-user; would need real architecture.
- **Stripe billing / subscription handling** — `subscription_status` field exists but no payment flow.

### Why the system is now "review-ready"

By March 31, 2027 (end of FY2027), Caleb's wife's workflow:
1. Throughout the year: import statements, log CRA payments, log invoices, scan receipts → all flowing to the books
2. Reconcile each statement period via 🏦 Account Balances → catches errors early
3. At FYE: log year-end accruals (AR/AP/prepaids) → 📓 Adjusting Entries
4. Update 🛠 Fixed Assets opening UCC + add new acquisitions → CCA computes
5. Open 📊 T2 Worksheet → all the numbers MNP needs are there
6. Build year-end package → Drive folder with cover letter + books snapshot + statements + receipts
7. Email link to MNP → they review (not prepare), verify, file

That should drop accountant fees from $3-5k (full prep) to $500-1.5k (review-only).
