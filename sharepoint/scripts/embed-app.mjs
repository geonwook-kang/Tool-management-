/* Build-time only: include the offline program in the SharePoint package. */
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const index = process.argv.indexOf('--input');
const input = path.resolve(index >= 0 ? process.argv[index + 1] : process.env.TOOL_MANAGER_APP || '/workspace/outputs/Tool-Manager-SharePoint.html');
const bytes = await readFile(input);
const html = bytes.toString('utf8');
if (bytes.byteLength > 32 * 1024 * 1024 || !html.startsWith('<!doctype html>') || !html.includes('ToolSharePointFile') || !html.includes('</html>')) {
  throw new Error('Build a complete SharePoint edition HTML with the offline OCR runtime before packaging.');
}
if (/<(?:script|link|meta|section|style)\b[^>]*\bdata-pwa-shell\b/i.test(html)) throw new Error('Use the SharePoint edition, not a PWA shell.');
const hash = createHash('sha256').update(bytes).digest('hex');
const generated = path.join(root, 'src/generated');
await mkdir(generated, { recursive: true });
await writeFile(path.join(generated, 'appPayload.ts'), [
  '/* Generated from the verified public program. Do not edit or commit. */',
  `export const APP_HTML: string = ${JSON.stringify(html).replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029')};`,
  `export const APP_SHA256: string = ${JSON.stringify(hash)};`,
  `export const APP_BYTES: number = ${bytes.byteLength};`,
  ''
].join('\n'));
await writeFile(path.join(root, 'app-build-info.json'), JSON.stringify({
  application_sha256: hash, application_bytes: bytes.byteLength,
  sharepoint_framework: '1.24.0', includes_offline_ocr: true,
  requires_graph_app_registration: false, private_company_data_embedded: false
}, null, 2) + '\n');
console.log(`Embedded SharePoint program: ${bytes.byteLength.toLocaleString()} bytes; SHA256 ${hash}`);
