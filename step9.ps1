# =============================================================================
# Step 9 — Rewire frontend to Worker endpoints, drop Apps Script references
# =============================================================================
# Run from the repo root: .\step9.ps1
# Modifies: app/index.html
# Does NOT commit or push — you do that manually after reviewing the diff.
#
# This script is idempotent: re-running it on an already-patched file will
# report "already applied" for each edit instead of erroring out.
#
# If ANY edit fails, the script halts WITHOUT writing the file, and tells
# you which edit failed. No partial state.
# =============================================================================

$ErrorActionPreference = 'Stop'

$filePath = 'app/index.html'
if (-not (Test-Path $filePath)) {
    Write-Error "File not found: $filePath. Run this script from the repo root (where the 'app' folder is)."
    exit 1
}

Write-Host ""
Write-Host "Step 9 patch — AI Bookkeeper frontend" -ForegroundColor Cyan
Write-Host "Reading $filePath..." -ForegroundColor Gray

$content = Get-Content $filePath -Raw
$original = $content

# Track edits for summary at end
$applied = @()
$skipped = @()
$failed = @()

# -----------------------------------------------------------------------------
# Helper: apply a single find/replace. If find is not found AND skipIfMissing
# is $true, it's treated as already-applied (idempotency). Otherwise it fails.
# -----------------------------------------------------------------------------
function Apply-Edit {
    param(
        [string]$name,
        [string]$find,
        [string]$replace,
        [bool]$skipIfMissing = $false
    )
    $script:count = ([regex]::Matches($script:content, [regex]::Escape($find))).Count
    if ($script:count -eq 0) {
        if ($skipIfMissing) {
            Write-Host "  [SKIP] $name  (already applied or absent)" -ForegroundColor DarkGray
            $script:skipped += $name
            return
        } else {
            Write-Host "  [FAIL] $name  (find string not found)" -ForegroundColor Red
            $script:failed += $name
            return
        }
    }
    if ($script:count -gt 1) {
        Write-Host "  [FAIL] $name  (find string matched $($script:count) times, expected 1)" -ForegroundColor Red
        $script:failed += $name
        return
    }
    $script:content = $script:content.Replace($find, $replace)
    Write-Host "  [ OK ] $name" -ForegroundColor Green
    $script:applied += $name
}

Write-Host ""
Write-Host "Applying edits..." -ForegroundColor Cyan

# -----------------------------------------------------------------------------
# Edit 1 — Remove scriptUrl from USER_CONFIG defaults
# -----------------------------------------------------------------------------
Apply-Edit -name "Edit 1: USER_CONFIG.scriptUrl default" `
    -find "  scriptUrl: '',`n" `
    -replace "" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 2 — Delete getSheetsUrl() helper entirely
# -----------------------------------------------------------------------------
$find_e2 = @"
function getSheetsUrl() {
  return USER_CONFIG.scriptUrl || '';
}

// Replace all `SHEETS_URL`` references with ``getSheetsUrl()``
// Or: const SHEETS_URL = () => USER_CONFIG.scriptUrl;
"@
# Note: the above uses PowerShell backtick-escape for backticks in the here-string.
# Actual content contains literal backticks. Let's build it programmatically:
$bt = [char]96  # backtick
$find_e2 = "function getSheetsUrl() {`n  return USER_CONFIG.scriptUrl || '';`n}`n`n// Replace all $($bt)SHEETS_URL$($bt) references with $($bt)getSheetsUrl()$($bt)`n// Or: const SHEETS_URL = () => USER_CONFIG.scriptUrl;"
Apply-Edit -name "Edit 2: getSheetsUrl() helper" `
    -find $find_e2 `
    -replace "" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 3 — Delete Settings panel Apps Script URL field + helper text
# -----------------------------------------------------------------------------
$find_e3 = "    `${f('Apps Script URL','set_scriptUrl',C.scriptUrl)}`n    <div style=`"font-size:.78rem;color:var(--ink3);margin-top:-.5rem;margin-bottom:1rem`">The Apps Script URL is the web app URL from Extensions → Apps Script → Deploy. If you used the `"Create my Google Sheet`" button above, both fields fill in automatically.</div>`n"
Apply-Edit -name "Edit 3: Settings Apps Script URL field" `
    -find $find_e3 `
    -replace "" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 4 — Delete scriptUrl from saveSettings payload
# -----------------------------------------------------------------------------
Apply-Edit -name "Edit 4: saveSettings scriptUrl" `
    -find "    scriptUrl: document.getElementById('set_scriptUrl')?.value || '',`n" `
    -replace "" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 5 — Delete scriptUrl from obProfile defaults
# -----------------------------------------------------------------------------
Apply-Edit -name "Edit 5: obProfile defaults" `
    -find "  sheetId:'', scriptUrl:''" `
    -replace "  sheetId:''" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 6 — Delete Apps Script URL field from onboarding step 3
# -----------------------------------------------------------------------------
Apply-Edit -name "Edit 6: Onboarding Apps Script URL field" `
    -find "        `${field('Apps Script URL (optional)','ob_scriptUrl','https://script.google.com/macros/s/...',p.scriptUrl)}`n" `
    -replace "" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 7 — Clean error helper text
# -----------------------------------------------------------------------------
Apply-Edit -name "Edit 7: Error helper text" `
    -find "Try again, or paste the Sheet ID and Apps Script URL manually below." `
    -replace "Try again, or paste the Sheet ID manually below." `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 8 — Delete obProfile.scriptUrl assignment
# -----------------------------------------------------------------------------
Apply-Edit -name "Edit 8: obProfile.scriptUrl assignment" `
    -find "    obProfile.scriptUrl = setupData.scriptUrl || '';`n" `
    -replace "" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 9 — Clean toast after googleAutoSetup (may already be patched by 8d.1)
# -----------------------------------------------------------------------------
$find_e9 = "    toast(setupData.scriptUrl`n      ? '✅ Sheet created and script deployed automatically!'`n      : '✅ Sheet created! You may need to deploy the script manually for full features.');"
Apply-Edit -name "Edit 9: googleAutoSetup toast" `
    -find $find_e9 `
    -replace "    toast('✅ Sheet created! Your books are ready to go.');" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 10 — Clean comment "Step 2: Create the sheet + Apps Script"
# -----------------------------------------------------------------------------
Apply-Edit -name "Edit 10: Step 2 comment" `
    -find "    // Step 2: Create the sheet + Apps Script" `
    -replace "    // Step 2: Create the sheet" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 11 — Delete scriptUrl from partial-setup profile save (around line 5837)
# -----------------------------------------------------------------------------
# This one occurs twice; handle both occurrences.
$find_e11 = "        scriptUrl: setupData.scriptUrl || '',`n"
$matches_e11 = ([regex]::Matches($content, [regex]::Escape($find_e11))).Count
if ($matches_e11 -eq 0) {
    Write-Host "  [SKIP] Edit 11: partial-setup scriptUrl  (already removed)" -ForegroundColor DarkGray
    $skipped += "Edit 11"
} else {
    # Remove all occurrences (typically 2)
    $content = $content.Replace($find_e11, "")
    Write-Host "  [ OK ] Edit 11: scriptUrl in profile save ($matches_e11 occurrence(s))" -ForegroundColor Green
    $applied += "Edit 11 (x$matches_e11)"
}

# -----------------------------------------------------------------------------
# Edit 12 — Clean successMsg ternary after manual create (may already be patched by 8d.1)
# -----------------------------------------------------------------------------
$find_e12 = "    const successMsg = setupData.scriptUrl`n      ? '✅ Sheet created and Apps Script deployed! Reloading...'`n      : '⚠️ Sheet created but Apps Script deployment failed — you may need to deploy it manually. Reloading...';`n    toast(successMsg);"
Apply-Edit -name "Edit 12: successMsg ternary" `
    -find $find_e12 `
    -replace "    toast('✅ Sheet created! Reloading...');" `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 13 — Rewire sendToSheets (Statement Importer)
# -----------------------------------------------------------------------------
# Replace the url-based preamble with JWT-based preamble
$find_e13a = @"
async function sendToSheets(){
  const btn = document.getElementById('sendSheetsBtn');
  const status = document.getElementById('sendSheetsStatus');
  const url = getSheetsUrl();

  if(!url){
    toast('Connect your Google Sheet first — go to Settings or complete onboarding');
    return;
  }
"@
$replace_e13a = @"
async function sendToSheets(){
  const btn = document.getElementById('sendSheetsBtn');
  const status = document.getElementById('sendSheetsStatus');
  const token = sessionStorage.getItem(AUTH_KEY);

  if(!token){
    toast('Please sign in first');
    return;
  }
  if(!USER_CONFIG.sheetId){
    toast('Connect your Google Sheet first — go to Settings or complete onboarding');
    return;
  }
"@
Apply-Edit -name "Edit 13a: sendToSheets preamble" `
    -find $find_e13a `
    -replace $replace_e13a `
    -skipIfMissing $true

# Replace the fetch call
$find_e13b = @"
    const r = await fetch(url, {
      method: 'POST',
      headers: {'Content-Type': 'text/plain'},
      body: JSON.stringify(payload)
    });
"@
# This pattern occurs in multiple places (sendToSheets, invSendToSheets, rcptSendToSheets).
# We need to be surgical about this one. Count occurrences first.
$count_e13b = ([regex]::Matches($content, [regex]::Escape($find_e13b))).Count
if ($count_e13b -eq 0) {
    Write-Host "  [SKIP] Edit 13b: sendToSheets fetch  (already patched)" -ForegroundColor DarkGray
    $skipped += "Edit 13b"
} else {
    # We'll replace the FIRST occurrence (which is sendToSheets) with /api/import/statement
    # The other two occurrences (invSendToSheets, rcptSendToSheets) will be handled by Edits 14 and 15
    # but they also use the same pattern. To disambiguate, we use the fact that sendToSheets is the first.
    # BUT — the old ones point to `url` and we want different target URLs for each.
    # So let's handle all three by searching with unique surrounding context.
    Write-Host "  [INFO] Edit 13b handled together with 14b, 15b below" -ForegroundColor DarkGray
}

# Surgical replacement of the sendToSheets fetch specifically
# Use the surrounding context (the code right before `const r = await fetch(url`) as anchor
$find_e13b_anchored = @"
    }
  }

  try {
    const r = await fetch(url, {
      method: 'POST',
      headers: {'Content-Type': 'text/plain'},
      body: JSON.stringify(payload)
    });
    const d = await r.json();

    if(d.ok){
      window._sentKeys.add(dupKey);
"@
$replace_e13b_anchored = @"
    }
  }

  try {
    const r = await fetch('/api/import/statement', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + token,
      },
      body: JSON.stringify(payload)
    });
    const d = await r.json();

    if(d.ok){
      window._sentKeys.add(dupKey);
"@
Apply-Edit -name "Edit 13b: sendToSheets fetch" `
    -find $find_e13b_anchored `
    -replace $replace_e13b_anchored `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 14 — Rewire invSendToSheets (Invoice Wizard)
# -----------------------------------------------------------------------------
# 14a: guard block — add JWT and sheetId checks
$find_e14a = @"
  const btn = document.getElementById('invSheetsBtn');
  const status = document.getElementById('invSheetsStatus');
  const url = getSheetsUrl();
  const hst = document.getElementById('invHST').checked;
  const client = document.getElementById('invClient').value.trim();
  const dateVal = document.getElementById('invDate').value;
  const dueVal = document.getElementById('invDue').value;

  if(!url){ toast('Connect your Google Sheet first — complete onboarding or check Settings'); return; }
  if(!client){ toast('Please enter a client name first'); return; }
"@
$replace_e14a = @"
  const btn = document.getElementById('invSheetsBtn');
  const status = document.getElementById('invSheetsStatus');
  const token = sessionStorage.getItem(AUTH_KEY);
  const hst = document.getElementById('invHST').checked;
  const client = document.getElementById('invClient').value.trim();
  const dateVal = document.getElementById('invDate').value;
  const dueVal = document.getElementById('invDue').value;

  if(!token){ toast('Please sign in first'); return; }
  if(!USER_CONFIG.sheetId){ toast('Connect your Google Sheet first — complete onboarding or check Settings'); return; }
  if(!client){ toast('Please enter a client name first'); return; }
"@
Apply-Edit -name "Edit 14a: invSendToSheets preamble" `
    -find $find_e14a `
    -replace $replace_e14a `
    -skipIfMissing $true

# 14b: payload — rename taxAmt/taxCharged to hstAmt/hst
$find_e14b = @"
      desc, sub, taxAmt: hstAmt, total,
      taxCharged: hst ? 'Yes' : 'No',
"@
$replace_e14b = @"
      desc, sub, hstAmt, total,
      hst: hst ? 'Yes' : 'No',
"@
Apply-Edit -name "Edit 14b: invoice payload field names" `
    -find $find_e14b `
    -replace $replace_e14b `
    -skipIfMissing $true

# 14c: fetch — change url to /api/invoice/create, add auth header
$find_e14c = @"
  try {
    const r = await fetch(url, {
      method: 'POST',
      headers: {'Content-Type': 'text/plain'},
      body: JSON.stringify(payload)
    });
    const d = await r.json();
    if(d.ok){
      btn.textContent = '✓ Sent!';
      btn.style.background = 'var(--green)';
      const sheetLink = USER_CONFIG?.sheetUrl ? ``<a href="${USER_CONFIG.sheetUrl}" target="_blank" style="color:var(--green)">Open Sheet ↗</a>`` : '';
      status.innerHTML = ``✅ Invoice #${invNum} added to <strong>Invoices</strong> and <strong>Income</strong> tabs. ${sheetLink}``;
"@
# The above uses embedded ``$`` which makes it tricky. Build programmatically:
$find_e14c = "  try {`n    const r = await fetch(url, {`n      method: 'POST',`n      headers: {'Content-Type': 'text/plain'},`n      body: JSON.stringify(payload)`n    });`n    const d = await r.json();`n    if(d.ok){`n      btn.textContent = '✓ Sent!';`n      btn.style.background = 'var(--green)';"
$replace_e14c = "  try {`n    const r = await fetch('/api/invoice/create', {`n      method: 'POST',`n      headers: {`n        'Content-Type': 'application/json',`n        'Authorization': 'Bearer ' + token,`n      },`n      body: JSON.stringify(payload)`n    });`n    const d = await r.json();`n    if(d.ok){`n      btn.textContent = '✓ Sent!';`n      btn.style.background = 'var(--green)';"
Apply-Edit -name "Edit 14c: invSendToSheets fetch" `
    -find $find_e14c `
    -replace $replace_e14c `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 15 — Rewire rcptSendToSheets (Receipt Scanner) to two-call sequence
# -----------------------------------------------------------------------------
# 15a: preamble
$find_e15a = @"
  const btn = document.getElementById('rcptSheetsBtn');
  const status = document.getElementById('rcptSheetsStatus');
  const url = getSheetsUrl();
  const cat = document.getElementById('rcptCatSel')?.value || r.category || 'Other';
 
  if(!url){
    toast('Connect your Google Sheet first — complete onboarding or check Settings');
    return;
  }
"@
$replace_e15a = @"
  const btn = document.getElementById('rcptSheetsBtn');
  const status = document.getElementById('rcptSheetsStatus');
  const token = sessionStorage.getItem(AUTH_KEY);
  const cat = document.getElementById('rcptCatSel')?.value || r.category || 'Other';
 
  if(!token){
    toast('Please sign in first');
    return;
  }
  if(!USER_CONFIG.sheetId){
    toast('Connect your Google Sheet first — complete onboarding or check Settings');
    return;
  }
"@
Apply-Edit -name "Edit 15a: rcptSendToSheets preamble" `
    -find $find_e15a `
    -replace $replace_e15a `
    -skipIfMissing $true

# 15b: payload + fetch block — replace with two-call sequence
$find_e15b = @"
  const payload = {
    bank: USER_CONFIG?.creditCard || 'Receipt',
    rows: [{
      date: r.date || '',
      vendor: r.vendor || '',
      description: r.notes || r.vendor || '',
      amount: -(parseFloat(r.total) || 0),
      category: cat,
      hst: parseFloat(r.hst) || 0,
      net: -(parseFloat(r.net) || 0),
      receiptVerified: true,
      matchedStatement: matchedTxn ? matchedTxn.vendor : null
    }],
    // ── NEW: include receipt image for Google Drive backup ──
    receiptImage: rcptB64 || null,
    receiptMimeType: rcptMediaType || null
  };
 
  try {
    const resp = await fetch(url, {
      method: 'POST',
      headers: {'Content-Type': 'text/plain'},
      body: JSON.stringify(payload)
    });
    const d = await resp.json();
"@
$replace_e15b = @"
  // Two-call sequence: (1) upload image to Drive (best-effort), (2) write ledger row.
  // No clickthrough from ledger to image in Phase 1 — Drive URL lives in console only.
  const statementPayload = {
    bank: 'Receipt',
    source: 'Receipt Scanner',
    rows: [{
      date: r.date || '',
      vendor: r.vendor || '',
      description: r.notes || r.vendor || '',
      amount: -(parseFloat(r.total) || 0),
      category: cat,
      hst: parseFloat(r.hst) || 0,
      net: -(parseFloat(r.net) || 0),
    }],
  };
 
  try {
    // Step 1: upload the receipt image to Drive (best-effort — don't block ledger write)
    let driveUrl = null;
    if (rcptB64) {
      try {
        const upResp = await fetch('/api/receipt/upload', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + token,
          },
          body: JSON.stringify({
            image: rcptB64,
            mimeType: rcptMediaType || 'image/jpeg',
          }),
        });
        const upData = await upResp.json();
        if (upData.ok) driveUrl = upData.driveUrl;
        else console.warn('Receipt image upload failed:', upData.error);
      } catch (e) {
        console.warn('Receipt image upload threw:', e);
      }
    }

    // Step 2: write the expense row to the ledger
    const resp = await fetch('/api/import/statement', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + token,
      },
      body: JSON.stringify(statementPayload),
    });
    const d = await resp.json();
    if (driveUrl) d.receiptDriveUrl = driveUrl; // pass through for existing toast logic
"@
Apply-Edit -name "Edit 15b: rcptSendToSheets payload + fetch" `
    -find $find_e15b `
    -replace $replace_e15b `
    -skipIfMissing $true

# -----------------------------------------------------------------------------
# Edit 16 — Stub taxFetchSummary (body replacement, preserves signature)
# -----------------------------------------------------------------------------
# We find the entire function body and replace with a single toast call.
# The function is contained — use the full signature and closing brace pattern.
$find_e16 = @"
async function taxFetchSummary(){
  const url=getSheetsUrl();
  if(!url){toast('Connect your Google Sheet first (Settings → Apps Script URL)');return}
"@
# We want to match from 'async function taxFetchSummary(){' to the closing '}'
# of the function. Use regex because the body is ~35 lines of variable content.
# Build a regex that matches the whole function:
$regex_e16 = '(?s)async function taxFetchSummary\(\)\{.*?\n\}'
if ($content -match $regex_e16) {
    $newFunction = "async function taxFetchSummary(){`n  toast('Tax Intelligence is being rebuilt — available in the next update');`n}"
    $content = [regex]::Replace($content, $regex_e16, [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        return $newFunction
    })
    Write-Host "  [ OK ] Edit 16: taxFetchSummary stub" -ForegroundColor Green
    $applied += "Edit 16"
} else {
    # Check if already stubbed
    if ($content -match "async function taxFetchSummary\(\)\{\s*toast\('Tax Intelligence") {
        Write-Host "  [SKIP] Edit 16: taxFetchSummary  (already stubbed)" -ForegroundColor DarkGray
        $skipped += "Edit 16"
    } else {
        Write-Host "  [FAIL] Edit 16: taxFetchSummary  (function not found)" -ForegroundColor Red
        $failed += "Edit 16"
    }
}

# -----------------------------------------------------------------------------
# Edit 17 — Stub taxRunHealthCheck
# -----------------------------------------------------------------------------
$regex_e17 = '(?s)async function taxRunHealthCheck\(\)\{.*?\n\}'
if ($content -match $regex_e17) {
    $newFunction = "async function taxRunHealthCheck(){`n  toast('Tax Intelligence is being rebuilt — available in the next update');`n}"
    $content = [regex]::Replace($content, $regex_e17, [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        return $newFunction
    })
    Write-Host "  [ OK ] Edit 17: taxRunHealthCheck stub" -ForegroundColor Green
    $applied += "Edit 17"
} else {
    if ($content -match "async function taxRunHealthCheck\(\)\{\s*toast\('Tax Intelligence") {
        Write-Host "  [SKIP] Edit 17: taxRunHealthCheck  (already stubbed)" -ForegroundColor DarkGray
        $skipped += "Edit 17"
    } else {
        Write-Host "  [FAIL] Edit 17: taxRunHealthCheck  (function not found)" -ForegroundColor Red
        $failed += "Edit 17"
    }
}

# -----------------------------------------------------------------------------
# Summary and write
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "  Applied:  $($applied.Count)" -ForegroundColor Green
Write-Host "  Skipped:  $($skipped.Count) (already patched or absent)" -ForegroundColor DarkGray
Write-Host "  Failed:   $($failed.Count)" -ForegroundColor $(if($failed.Count -eq 0){'Green'}else{'Red'})

if ($failed.Count -gt 0) {
    Write-Host ""
    Write-Host "Edits that failed:" -ForegroundColor Red
    foreach ($f in $failed) { Write-Host "    - $f" -ForegroundColor Red }
    Write-Host ""
    Write-Host "ABORTING — file NOT written." -ForegroundColor Red
    Write-Host "Screenshot the failed edit(s) in VS Code and send to Claude for diagnosis." -ForegroundColor Yellow
    exit 1
}

if ($content -eq $original) {
    Write-Host ""
    Write-Host "No changes to write — file is already fully patched." -ForegroundColor Yellow
    exit 0
}

Write-Host ""
Write-Host "Writing $filePath..." -ForegroundColor Gray
$content | Set-Content $filePath -NoNewline

Write-Host ""
Write-Host "Done. Next steps:" -ForegroundColor Cyan
Write-Host "  1. Run: git diff app/index.html" -ForegroundColor White
Write-Host "  2. Eyeball the diff — should be ~15-25 hunks, all related to Apps Script / scriptUrl / new endpoints" -ForegroundColor White
Write-Host "  3. Run: grep 'scriptUrl\|script_url\|getSheetsUrl' app/index.html" -ForegroundColor White
Write-Host "     (Expect zero hits, or only hits inside comments)" -ForegroundColor DarkGray
Write-Host "  4. If all looks right:" -ForegroundColor White
Write-Host "       git add app/index.html" -ForegroundColor White
Write-Host "       git commit -m 'Phase 1/9: Rewire frontend to Worker endpoints; drop Apps Script references'" -ForegroundColor White
Write-Host "       git push" -ForegroundColor White
Write-Host ""
