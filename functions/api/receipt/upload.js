// ════════════════════════════════════════════════════════════════════
// POST /api/receipt/upload
// Upload a receipt image to the user's Google Drive and return a share URL.
// Replaces Apps Script's handleReceiptUpload().
// ════════════════════════════════════════════════════════════════════

import { getGoogleAccessToken } from '../../_google.js';
import { authenticateRequest, json, options } from '../../_shared.js';

const DRIVE_API = 'https://www.googleapis.com/drive/v3/files';
const UPLOAD_API = 'https://www.googleapis.com/upload/drive/v3/files';
const PARENT_FOLDER = 'AI Bookkeeper Receipts';

export const onRequestOptions = () => options();

export async function onRequestPost({ request, env }) {
    const auth = await authenticateRequest(request, env);
    if (!auth) return json({ ok: false, error: 'Unauthorized' }, 401);
    const userId = auth.userId;

  let body;
    try {
          body = await request.json();
    } catch {
          return json({ ok: false, error: 'Invalid JSON' }, 400);
    }

  if (!body.image) return json({ ok: false, error: 'No image data' }, 400);

  const tok = await getGoogleAccessToken(env, userId);
    if (!tok.ok) return json({ ok: false, error: tok.error, needsReauth: tok.needsReauth }, 401);

  const year = String(new Date().getFullYear());

  // Find or create parent folder, then year subfolder
  const parentId = await findOrCreateFolder(tok.accessToken, PARENT_FOLDER, null);
    if (!parentId) return json({ ok: false, error: 'Could not create receipts folder' }, 500);

  const yearFolderId = await findOrCreateFolder(tok.accessToken, year, parentId);
    if (!yearFolderId) return json({ ok: false, error: 'Could not create year folder' }, 500);

  // Decode base64 image (accept either "data:image/jpeg;base64,..." or plain base64)
  const base64Data = body.image.includes(',') ? body.image.split(',').pop() : body.image;
    const binary = Uint8Array.from(atob(base64Data), c => c.charCodeAt(0));

  const filename = body.filename || `Receipt_${Date.now()}.jpg`;
    const mimeType = body.mimeType || 'image/jpeg';

  // Multipart upload body
  const boundary = '-------receipt-boundary-' + Math.random().toString(36).slice(2);
    const metadata = { name: filename, parents: [yearFolderId], mimeType };
    const preamble = `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${boundary}\r\nContent-Type: ${mimeType}\r\n\r\n`;
    const epilogue = `\r\n--${boundary}--`;

  const enc = new TextEncoder();
    const part1 = enc.encode(preamble);
    const part3 = enc.encode(epilogue);
    const fullBody = new Uint8Array(part1.length + binary.length + part3.length);
    fullBody.set(part1, 0);
    fullBody.set(binary, part1.length);
    fullBody.set(part3, part1.length + binary.length);

  const uploadResp = await fetch(`${UPLOAD_API}?uploadType=multipart`, {
        method: 'POST',
        headers: {
                'Authorization': `Bearer ${tok.accessToken}`,
                'Content-Type': `multipart/related; boundary=${boundary}`,
        },
        body: fullBody,
  });

  const uploadData = await uploadResp.json();
    if (uploadData.error) return json({ ok: false, error: 'Upload failed: ' + uploadData.error.message }, 500);

  // Make the file viewable by anyone with the link
  await fetch(`${DRIVE_API}/${uploadData.id}/permissions`, {
        method: 'POST',
        headers: {
                'Authorization': `Bearer ${tok.accessToken}`,
                'Content-Type': 'application/json',
        },
        body: JSON.stringify({ role: 'reader', type: 'anyone' }),
  });

  const viewUrl = `https://drive.google.com/file/d/${uploadData.id}/view`;
    return json({ ok: true, driveUrl: viewUrl, fileId: uploadData.id });
}

async function findOrCreateFolder(accessToken, name, parentId) {
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
        headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
        },
        body: JSON.stringify(metadata),
  });
    const createData = await createResp.json();
    return createData.id || null;
}
