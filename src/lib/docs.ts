/* Shared SharePoint document helpers (CFM_CollectionDocs Power Automate flow).
 * action = upload | list | delete. URL ported directly from Ref.html — the flow
 * is scoped to this one purpose and shared by any page that lists/uploads/deletes
 * documents against a SharePoint folder (Daily Collection, Dispute Management, …).
 */
export const DOCS_FLOW_URL =
  'https://9ce6fb095b63e9f49185b707b4b342.5e.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/30/workflows/f925d853a7c944ca88f9e886f1290e45/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=XnsQ8q4II-APCA1K7gBM--Zd0E_2kcqxV5fnaVVDlDU';

export interface DocFile {
  name: string;
  url: string;
  path: string;
  size?: number;
}

export function docExt(name: string): string {
  return (name || '').split('.').pop()?.toLowerCase() || '';
}

export function docIconClass(ext: string): string {
  if (ext === 'pdf') return 'pdf';
  if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'heic'].includes(ext)) return 'img';
  if (['doc', 'docx', 'rtf', 'txt'].includes(ext)) return 'doc';
  if (['xls', 'xlsx', 'xlsb', 'csv'].includes(ext)) return 'xls';
  return 'other';
}

export function docIcon(ext: string): string {
  if (ext === 'pdf') return 'fa-file-pdf';
  if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg', 'heic'].includes(ext)) return 'fa-file-image';
  if (['doc', 'docx', 'rtf', 'txt'].includes(ext)) return 'fa-file-word';
  if (['xls', 'xlsx', 'xlsb', 'csv'].includes(ext)) return 'fa-file-excel';
  if (['ppt', 'pptx'].includes(ext)) return 'fa-file-powerpoint';
  if (['zip', 'rar', '7z'].includes(ext)) return 'fa-file-zipper';
  return 'fa-file';
}

export function fmtBytes(b: number | undefined): string {
  const n = Number(b || 0);
  if (!n) return '';
  if (n < 1024) return `${n} B`;
  if (n < 1048576) return `${(n / 1024).toFixed(0)} KB`;
  return `${(n / 1048576).toFixed(1)} MB`;
}

export function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
    reader.onerror = () => reject(new Error(`Could not read ${file.name}`));
    reader.readAsDataURL(file);
  });
}

export async function docsFlow(payload: Record<string, unknown>): Promise<unknown> {
  const r = await fetch(DOCS_FLOW_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!r.ok) {
    const t = await r.text();
    throw new Error(`Flow ${r.status}: ${t.slice(0, 200)}`);
  }
  const txt = await r.text();
  try {
    return txt ? JSON.parse(txt) : null;
  } catch {
    return txt;
  }
}

export async function listDocs(folderName: string): Promise<DocFile[]> {
  const res = await docsFlow({ action: 'list', folderName });
  if (Array.isArray(res)) return res as DocFile[];
  if (res && Array.isArray((res as { value?: unknown }).value)) return (res as { value: DocFile[] }).value;
  if (res && Array.isArray((res as { files?: unknown }).files)) return (res as { files: DocFile[] }).files;
  return [];
}

// The API gateway in front of the flow times out (502 "NoResponse") on a slow
// upload well before the flow itself gives up — the flow keeps running and
// usually finishes the SharePoint upload anyway. Treat that as inconclusive
// rather than failed.
function isGatewayTimeout(message: string): boolean {
  return /\b502\b|no\s*response|did not receive a response|gateway/i.test(message);
}

export async function uploadDoc(payload: {
  folderName: string;
  collectionId: string;
  collectionRef: string;
  fileName: string;
  fileContentBase64: string;
}): Promise<void> {
  try {
    await docsFlow({ action: 'upload', ...payload });
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    if (!isGatewayTimeout(msg)) throw err;
    // Verify by re-listing the folder instead of trusting the gateway timeout
    // as a real failure — if the file actually landed, report success.
    try {
      const files = await listDocs(payload.folderName);
      if (files.some((f) => f.name === payload.fileName)) return;
    } catch {
      // listing failed too — fall through to the original error below
    }
    throw err;
  }
}
