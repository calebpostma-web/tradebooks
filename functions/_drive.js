// ════════════════════════════════════════════════════════════════════
// GOOGLE DRIVE OPERATIONS
//
// Shared helpers for the year-end package builder, statement archive, and
// receipt/remittance uploaders. Avoids the same findOrCreateFolder copy
// living in three files.
// ════════════════════════════════════════════════════════════════════

const DRIVE_API = 'https://www.googleapis.com/drive/v3/files';
const UPLOAD_API = 'https://www.googleapis.com/upload/drive/v3/files';
const DOCS_API = 'https://docs.googleapis.com/v1/documents';

// ── Folder operations ──────────────────────────────────────────────

/**
 * Find an existing folder by name (optionally under a parent), or create it.
 * Returns the folder ID, or null on failure.
 */
export async function findOrCreateFolder(accessToken, name, parentId) {
  const escapedName = name.replace(/'/g, "\\'");
  let query = `name='${escapedName}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
  if (parentId) query += ` and '${parentId}' in parents`;

  const searchResp = await fetch(
    `${DRIVE_API}?q=${encodeURIComponent(query)}&fields=files(id,name)`,
    { headers: { 'Authorization': `Bearer ${accessToken}` } }
  );
  const searchData = await searchResp.json();
  if (searchData.files && searchData.files.length) return searchData.files[0].id;

  const metadata = { name, mimeType: 'application/vnd.google-apps.folder' };
  if (parentId) metadata.parents = [parentId];

  const createResp = await fetch(DRIVE_API, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(metadata),
  });
  const createData = await createResp.json();
  return createData.id || null;
}

/**
 * Create a folder unconditionally (even if a folder with the same name
 * already exists at that location). Used by the year-end package builder
 * which intentionally writes a new dated folder per package run.
 */
export async function createFolder(accessToken, name, parentId) {
  const metadata = { name, mimeType: 'application/vnd.google-apps.folder' };
  if (parentId) metadata.parents = [parentId];
  const resp = await fetch(DRIVE_API, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(metadata),
  });
  const data = await resp.json();
  if (data.error) throw new Error(data.error.message || 'Folder create failed');
  return data.id;
}

/**
 * Make a file (or folder) viewable by anyone with the link. Used so MNP can
 * open the year-end package without needing to be added by email.
 */
export async function makeShareable(accessToken, fileId) {
  await fetch(`${DRIVE_API}/${fileId}/permissions`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ role: 'reader', type: 'anyone' }),
  });
}

// ── File upload + listing ──────────────────────────────────────────

/**
 * Upload a binary file to Drive. Returns { id, viewUrl }. Throws on failure.
 */
export async function uploadFile(accessToken, folderId, filename, mimeType, binary) {
  const boundary = '-------drive-boundary-' + Math.random().toString(36).slice(2);
  const metadata = { name: filename, parents: [folderId], mimeType };
  const preamble = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`;
  const epilogue = `\r\n--${boundary}--`;

  const enc = new TextEncoder();
  const part1 = enc.encode(preamble);
  const part3 = enc.encode(epilogue);
  const fullBody = new Uint8Array(part1.length + binary.length + part3.length);
  fullBody.set(part1, 0);
  fullBody.set(binary, part1.length);
  fullBody.set(part3, part1.length + binary.length);

  const resp = await fetch(`${UPLOAD_API}?uploadType=multipart`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body: fullBody,
  });
  const data = await resp.json();
  if (data.error) throw new Error(data.error.message || 'Drive upload failed');
  return { id: data.id, viewUrl: `https://drive.google.com/file/d/${data.id}/view` };
}

/**
 * List files in a folder. Returns array of { id, name, mimeType, createdTime, webViewLink }.
 * `query` (optional) can add additional filters, e.g. "mimeType='application/pdf'".
 * Pass `pageSize` to cap results. Note: name/mimeType/createdTime are always returned.
 */
export async function listFolderFiles(accessToken, folderId, { query = '', pageSize = 200 } = {}) {
  let q = `'${folderId}' in parents and trashed=false`;
  if (query) q += ` and ${query}`;
  const fields = encodeURIComponent('files(id,name,mimeType,createdTime,modifiedTime,webViewLink,size)');
  const resp = await fetch(
    `${DRIVE_API}?q=${encodeURIComponent(q)}&fields=${fields}&pageSize=${pageSize}`,
    { headers: { 'Authorization': `Bearer ${accessToken}` } }
  );
  const data = await resp.json();
  if (data.error) throw new Error(data.error.message || 'Drive list failed');
  return data.files || [];
}

// ── Shortcut + copy operations ─────────────────────────────────────

/**
 * Create a Drive shortcut to an existing file inside a target folder. Cheaper
 * and cleaner than copying — the year-end package uses shortcuts so files
 * don't get duplicated and stay live with the originals.
 */
export async function createShortcut(accessToken, targetFileId, parentFolderId, name) {
  const metadata = {
    name,
    mimeType: 'application/vnd.google-apps.shortcut',
    parents: [parentFolderId],
    shortcutDetails: { targetId: targetFileId },
  };
  const resp = await fetch(DRIVE_API, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(metadata),
  });
  const data = await resp.json();
  if (data.error) throw new Error(data.error.message || 'Shortcut create failed');
  return data.id;
}

// ── Sheet → XLSX export ────────────────────────────────────────────

/**
 * Export a single tab from a Google Sheet as XLSX bytes. Used by the year-end
 * package builder to dump the Transactions / Invoices tabs into the package
 * folder for MNP to open in Excel.
 *
 * Note: there's no native "export one tab" — Sheets API exports whole files.
 * For per-tab export the cleanest path is to copy the sheet to a temp file,
 * delete other tabs, export, delete the temp. That's heavy. Simpler: export
 * the whole spreadsheet as XLSX (gives MNP all tabs in one workbook, which
 * is actually preferable).
 */
export async function exportSpreadsheetAsXlsx(accessToken, spreadsheetId) {
  const resp = await fetch(
    `${DRIVE_API}/${spreadsheetId}/export?mimeType=${encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}`,
    { headers: { 'Authorization': `Bearer ${accessToken}` } }
  );
  if (!resp.ok) throw new Error(`Spreadsheet export failed: ${resp.status} ${resp.statusText}`);
  return new Uint8Array(await resp.arrayBuffer());
}

// ── Google Docs cover letter ───────────────────────────────────────

/**
 * Create a Google Doc inside a folder with the given title and body.
 * Body is plain text with simple "## heading" prefixes recognised. Returns
 * { id, viewUrl } so the caller can shareable-ify it.
 *
 * Implementation note: Docs API requires creating an empty doc first, then
 * batchUpdate-ing inserts. We then move the file into the target folder via
 * Drive API (Docs API doesn't accept a parents field on create).
 */
export async function createCoverLetterDoc(accessToken, folderId, title, body) {
  // 1. Create empty doc
  const createResp = await fetch(DOCS_API, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ title }),
  });
  const created = await createResp.json();
  if (created.error) throw new Error(created.error.message || 'Doc create failed');
  const docId = created.documentId;

  // 2. Move the doc into the target folder. Docs land in My Drive root by default.
  // We need to remove the My Drive parent and add the target folder parent — Drive's
  // PATCH supports addParents+removeParents query params, body is a no-op metadata patch.
  const moveResp = await fetch(`${DRIVE_API}/${docId}?addParents=${folderId}&removeParents=root&fields=id,parents`, {
    method: 'PATCH',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({}),
  });
  const moveData = await moveResp.json();
  if (moveData.error) throw new Error(`Cover letter created but move failed: ${moveData.error.message || 'unknown'}`);

  // 3. Insert body text. We do a single insert at index 1 (after the title cursor).
  // For headings we lean on \n separators and pre-formatted markers — keeping
  // the doc readable without complex paragraph styling round-trips.
  const requests = [{
    insertText: { location: { index: 1 }, text: body },
  }];
  await fetch(`${DOCS_API}/${docId}:batchUpdate`, {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ requests }),
  });

  return { id: docId, viewUrl: `https://docs.google.com/document/d/${docId}/edit` };
}
