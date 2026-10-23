/* XLSX updates stay in the browser. Only selected dates and photo slots change. */
(function (global) {
  'use strict';
  const NS = {
    s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    p: 'http://schemas.openxmlformats.org/package/2006/relationships',
    xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    ct: 'http://schemas.openxmlformats.org/package/2006/content-types'
  };
  const REL = NS.r + '/', EMU = 9525;
  const state = new WeakMap();
  const required = ['Cal Next', 'Cal Last', 'Serial', 'Asset #', 'Label Photo'];
  const allHeaders = required.concat(['Cal Interval', 'Cal Cycle', 'Manufacturer', 'Model', 'Description', 'SMAX Tool Photo']);
  const photoKind = (update) => update.photoKind === undefined ? 'calibration' : update.photoKind;
  const normHeader = (s) => String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
  const normalizeSerial = (s) => String(s == null ? '' : s).normalize('NFKC').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const normalizeName = (s) => String(s == null ? '' : s).normalize('NFC').trim().replace(/\s+/g, ' ');
  const elements = (node, local) => Array.from(node.getElementsByTagNameNS('*', local));
  const child = (node, local) => Array.from(node.children || []).find((n) => n.localName === local) || null;
  const val = (node, local) => { const n = child(node, local); return n ? n.textContent : ''; };
  const serialize = (doc) => new XMLSerializer().serializeToString(doc);
  function xml(text, path) {
    if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('지원하지 않는 XML 문서입니다: ' + path);
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (elements(doc, 'parsererror').length || doc.documentElement.localName === 'parsererror') throw new Error('엑셀 내부 문서를 읽을 수 없습니다: ' + path);
    return doc;
  }
  async function readXml(zip, path, optional) {
    const entry = zip.file(path);
    if (!entry) { if (optional) return null; throw new Error('엑셀 구성 파일이 없습니다: ' + path); }
    return xml(await entry.async('string'), path);
  }
  function resolve(base, target) {
    if (!target || /[?#\\]/.test(target) || /^[a-z]+:/i.test(target)) throw new Error('잘못된 엑셀 내부 연결입니다.');
    const parts = (target.startsWith('/') ? target.slice(1) : base.slice(0, base.lastIndexOf('/') + 1) + target).split('/');
    const out = [];
    for (const part of parts) {
      if (!part || part === '.') continue;
      if (part === '..') { if (!out.length) throw new Error('잘못된 엑셀 내부 경로입니다.'); out.pop(); }
      else out.push(part);
    }
    return out.join('/');
  }
  const relPath = (path) => path.slice(0, path.lastIndexOf('/') + 1) + '_rels/' + path.slice(path.lastIndexOf('/') + 1) + '.rels';
  function relative(base, destination) {
    const from = base.split('/').slice(0, -1), to = destination.split('/');
    while (from.length && to.length && from[0] === to[0]) { from.shift(); to.shift(); }
    return '../'.repeat(from.length) + to.join('/');
  }
  function relationships(doc) {
    const map = new Map();
    if (doc) for (const n of elements(doc, 'Relationship')) map.set(n.getAttribute('Id'), n);
    return map;
  }
  const relationshipId = (node, attr) => node.getAttributeNS(NS.r, attr || 'id') || node.getAttribute('r:' + (attr || 'id'));
  function columnNumber(ref) {
    const text = String(ref).match(/^[A-Z]+/i);
    return text ? Array.from(text[0].toUpperCase()).reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) : 0;
  }
  function columnLetter(n) {
    let out = '';
    while (n > 0) { n--; out = String.fromCharCode(65 + n % 26) + out; n = Math.floor(n / 26); }
    return out;
  }
  function richText(node) {
    return elements(node, 't').filter((n) => !n.parentElement || n.parentElement.localName !== 'rPh').map((n) => n.textContent).join('');
  }
  function cellValue(cell, strings) {
    if (!cell) return '';
    const type = cell.getAttribute('t');
    if (type === 's') return strings[Number(val(cell, 'v'))] || '';
    if (type === 'inlineStr') return richText(cell);
    return val(cell, 'v');
  }
  function isoDate(text) {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(text || ''));
    if (!match) return null;
    const y = Number(match[1]), m = Number(match[2]), d = Number(match[3]);
    if (y < 1900 || y > 9999 || m < 1 || m > 12 || d < 1 || d > 31) return null;
    const date = new Date(Date.UTC(y, m - 1, d));
    return date.getUTCFullYear() === y && date.getUTCMonth() === m - 1 && date.getUTCDate() === d ? date : null;
  }
  function excelDate(value, date1904) {
    const text = String(value || '').trim();
    if (!text) return '';
    if (/^\d+(?:\.\d+)?$/.test(text)) {
      const n = Math.floor(Number(text));
      if (n < 0 || n > 2958465) return text;
      if (!date1904 && n === 60) return '1900-02-29';
      const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
      const adjusted = date1904 || n < 60 ? n : n - 1;
      return new Date(base + adjusted * 86400000).toISOString().slice(0, 10);
    }
    const match = /^(\d{4})[.\/\-년]\s*(\d{1,2})[.\/\-월]\s*(\d{1,2})(?:일|[ T].*)?$/.exec(text);
    if (match) {
      const iso = match[1] + '-' + match[2].padStart(2, '0') + '-' + match[3].padStart(2, '0');
      if (isoDate(iso)) return iso;
    }
    return text;
  }
  function dateSerial(iso, date1904) {
    const date = isoDate(iso);
    if (!date) throw new Error('날짜는 YYYY-MM-DD 형식의 실제 날짜여야 합니다.');
    const base = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
    let n = Math.round((date.getTime() - base) / 86400000);
    if (!date1904 && n >= 60) n++;
    if (n < 0) throw new Error('이 엑셀의 날짜 체계보다 이전 날짜는 저장할 수 없습니다.');
    return n;
  }
  function addMonths(iso, months) {
    const date = isoDate(iso), count = Number(months);
    if (!date || !Number.isInteger(count) || count < 1 || count > 1200) return '';
    const y = date.getUTCFullYear(), m = date.getUTCMonth(), d = date.getUTCDate();
    const result = new Date(Date.UTC(y, m + count + 1, 0));
    result.setUTCDate(Math.min(d, result.getUTCDate()));
    return result.getUTCFullYear() > 9999 ? '' : result.toISOString().slice(0, 10);
  }
  function monthInterval(interval, cycle) {
    const n = Number(interval), c = String(cycle || '').trim().toUpperCase();
    if (!Number.isInteger(n) || n < 1) return null;
    if (['MONTHS', 'MONTH', '개월', '월'].includes(c)) return n;
    if (['YEARS', 'YEAR', '년'].includes(c)) return n * 12;
    return null;
  }
  function anchorLocation(anchor) {
    const from = child(anchor, 'from');
    return from ? { row: Number(val(from, 'row')) + 1, col: Number(val(from, 'col')) + 1 } : null;
  }
  async function load(input) {
    if (!global.JSZip) throw new Error('엑셀 읽기 모듈이 준비되지 않았습니다.');
    const bytes = input instanceof Uint8Array ? input.slice() : new Uint8Array(input).slice();
    if (bytes.byteLength > 100 * 1024 * 1024) throw new Error('엑셀 파일은 100 MB 이하로 올려 주세요.');
    let zip;
    try { zip = await global.JSZip.loadAsync(bytes); }
    catch (_) { throw new Error('엑셀을 열 수 없습니다. 암호가 없는 정상적인 .xlsx 파일인지 확인해 주세요.'); }
    const entries = Object.values(zip.files);
    if (entries.length > 5000 || entries.reduce((n, f) => n + (f._data && f._data.uncompressedSize || 0), 0) > 250 * 1024 * 1024) throw new Error('압축 해제 후 크기가 너무 큰 엑셀입니다. 사진 크기를 줄인 복사본을 사용해 주세요.');
    // Check integrity after the declared uncompressed size has been bounded.
    try { zip = await global.JSZip.loadAsync(bytes, { checkCRC32: true }); }
    catch (_) { throw new Error('엑셀 파일의 무결성 검사에 실패했습니다. 정상적인 원본을 다시 선택해 주세요.'); }
    if (entries.some((e) => /^_xmlsignatures\//i.test(e.name))) throw new Error('전자 서명된 엑셀은 수정할 수 없습니다. 서명 없는 복사본을 사용해 주세요.');
    if (zip.file('xl/vbaProject.bin')) throw new Error('매크로 없는 .xlsx 파일을 사용해 주세요.');
    const workbook = await readXml(zip, 'xl/workbook.xml');
    const wbRels = relationships(await readXml(zip, 'xl/_rels/workbook.xml.rels'));
    const wbPr = elements(workbook, 'workbookPr')[0];
    const date1904 = !!(wbPr && /^(1|true)$/i.test(wbPr.getAttribute('date1904') || ''));
    const stringsRel = Array.from(wbRels.values()).find((r) => r.getAttribute('Type').endsWith('/sharedStrings'));
    const stringDoc = stringsRel ? await readXml(zip, resolve('xl/workbook.xml', stringsRel.getAttribute('Target'))) : null;
    const strings = stringDoc ? elements(stringDoc, 'si').map(richText) : [];
    const sheets = [], tools = [], warnings = [];
    for (const sheet of elements(workbook, 'sheet')) {
      const rel = wbRels.get(relationshipId(sheet));
      if (!rel || rel.getAttribute('TargetMode') === 'External' || !rel.getAttribute('Type').endsWith('/worksheet')) continue;
      const path = resolve('xl/workbook.xml', rel.getAttribute('Target'));
      const doc = await readXml(zip, path);
      const rows = elements(doc, 'sheetData').flatMap((n) => Array.from(n.children).filter((r) => r.localName === 'row'));
      let headers = null, headerRow = 0;
      for (const row of rows.slice(0, 20)) {
        const found = {};
        for (const cell of Array.from(row.children).filter((n) => n.localName === 'c')) {
          const key = allHeaders.find((h) => normHeader(h) === normHeader(cellValue(cell, strings)));
          if (key) {
            if (found[key]) throw new Error(sheet.getAttribute('name') + ' 시트에 중복된 열 이름이 있습니다: ' + key);
            found[key] = columnNumber(cell.getAttribute('r'));
          }
        }
        if (required.every((h) => found[h])) { headers = found; headerRow = Number(row.getAttribute('r')); break; }
      }
      if (!headers) { warnings.push(sheet.getAttribute('name') + ': 툴 관리 열이 없어 읽기만 유지합니다.'); continue; }
      const relDoc = await readXml(zip, relPath(path), true);
      const drawingRecords = [], sheetRels = relationships(relDoc);
      for (const draw of elements(doc, 'drawing')) {
        const drawRel = sheetRels.get(relationshipId(draw));
        if (!drawRel || drawRel.getAttribute('TargetMode') === 'External') continue;
        const drawPath = resolve(path, drawRel.getAttribute('Target'));
        const drawDoc = await readXml(zip, drawPath);
        drawingRecords.push({ path: drawPath, doc: drawDoc, relDoc: await readXml(zip, relPath(drawPath), true) });
      }
      const record = { name: sheet.getAttribute('name'), path, doc, headers, headerRow, rows, drawings: drawingRecords, relDoc };
      sheets.push(record);
      for (const row of rows) {
        const rowNum = Number(row.getAttribute('r'));
        if (rowNum <= headerRow) continue;
        const cells = new Map(Array.from(row.children).filter((n) => n.localName === 'c').map((c) => [columnNumber(c.getAttribute('r')), c]));
        const get = (key) => cellValue(cells.get(headers[key]), strings).trim();
        const serial = get('Serial');
        if (!serial) continue;
        const asset = get('Asset #'), name = normalizeName(asset.split('_').pop());
        const picturesAt = (col) => !col ? [] : drawingRecords.flatMap((d) => Array.from(d.doc.documentElement.children).filter((a) => {
          const loc = anchorLocation(a);
          return loc && loc.row === rowNum && loc.col === col && elements(a, 'pic').length;
        }));
        const interval = get('Cal Interval');
        tools.push(Object.freeze({
          id: path + ':' + rowNum, sheet: record.name, row: rowNum, name, serial, asset,
          last: excelDate(get('Cal Last'), date1904), due: excelDate(get('Cal Next'), date1904),
          interval: interval === '' ? '' : Number.isFinite(Number(interval)) ? Number(interval) : interval,
          cycle: get('Cal Cycle'), months: monthInterval(interval, get('Cal Cycle')),
          manufacturer: get('Manufacturer'), model: get('Model'), description: get('Description'),
          photoCount: picturesAt(headers['Label Photo']).length,
          smaxPhotoCount: picturesAt(headers['SMAX Tool Photo']).length,
          hasSmaxPhotoColumn: !!headers['SMAX Tool Photo']
        }));
      }
    }
    if (!tools.length) throw new Error('툴 목록을 찾지 못했습니다. Cal Next, Cal Last, Serial, Asset #, Label Photo 열을 확인해 주세요.');
    const model = { tools: Object.freeze(tools), sheets: sheets.map((s) => s.name), warnings, date1904 };
    state.set(model, { bytes, zip, sheets, date1904 });
    return model;
  }
  function validate(model, updateOrUpdates) {
    const errors = [];
    if (!state.has(model)) return ['엑셀 파일을 먼저 열어 주세요.'];
    const updates = Array.isArray(updateOrUpdates) ? updateOrUpdates : [updateOrUpdates];
    if (!updates.length) return ['반영할 사진이 없습니다.'];
    const seen = new Set();
    for (const update of updates) {
      if (!update) { errors.push('사진 정보가 비어 있습니다.'); continue; }
      const tool = model.tools.find((t) => t.id === update.toolId);
      if (!tool) { errors.push('선택한 툴 항목이 엑셀에 없습니다.'); continue; }
      const label = tool.sheet + ' ' + tool.row + '행 (' + tool.serial + ')';
      const kind = photoKind(update), slot = tool.id + '\u0000' + kind;
      if (!['calibration', 'smax'].includes(kind)) { errors.push(label + ': 사진 구분은 검교정 또는 SMAX여야 합니다.'); continue; }
      if (seen.has(slot)) errors.push(label + ': 같은 툴의 같은 사진 칸에 여러 사진을 반영할 수 없습니다.');
      seen.add(slot);
      if (!tool.name || normalizeName(update.name) !== tool.name || !normalizeSerial(update.serial) || normalizeSerial(update.serial) !== normalizeSerial(tool.serial)) errors.push(label + ': 사진에서 확인한 이름·시리얼 번호가 엑셀 항목과 다릅니다.');
      if (model.tools.filter((t) => normalizeSerial(t.serial) === normalizeSerial(tool.serial)).length > 1) errors.push(label + ': 같은 시리얼 번호가 엑셀에 여러 번 있어 자동 반영할 수 없습니다.');
      if (/폐기|disposed|retired/i.test(tool.due)) errors.push(label + ': 폐기된 툴은 갱신할 수 없습니다.');
      if (kind === 'calibration') {
        if (tool.photoCount > 1) errors.push(label + ': Label Photo 위치에 사진이 여러 장 있어 반영할 대상을 확정할 수 없습니다.');
        if (!isoDate(update.last) || !isoDate(update.due)) errors.push(label + ': 교정일과 다음 교정일을 올바르게 입력해 주세요.');
        else {
          if (update.due <= update.last) errors.push(label + ': 다음 교정일은 교정일보다 뒤여야 합니다.');
          if (isoDate(tool.last) && update.last < tool.last) errors.push(label + ': 기존 교정일보다 오래된 사진입니다.');
          if (model.date1904 && update.last < '1904-01-01') errors.push(label + ': 1904년 이전 날짜는 저장할 수 없습니다.');
        }
      } else {
        if (!tool.hasSmaxPhotoColumn) errors.push(label + ': SMAX Tool Photo 열이 없어 사진을 반영할 수 없습니다.');
        if (tool.smaxPhotoCount > 1) errors.push(label + ': SMAX Tool Photo 위치에 사진이 여러 장 있어 반영할 대상을 확정할 수 없습니다.');
        if (update.updateDue !== undefined && typeof update.updateDue !== 'boolean') errors.push(label + ': SMAX 다음 교정일 반영 여부를 확인해 주세요.');
        if (update.updateDue === true) {
          if (!isoDate(update.due)) errors.push(label + ': SMAX 다음 교정일을 올바르게 입력해 주세요.');
          else {
            const paired = updates.find((other) => other && other.toolId === tool.id && photoKind(other) === 'calibration');
            const effectiveLast = paired && isoDate(paired.last) ? paired.last : tool.last;
            if (isoDate(effectiveLast) && update.due <= effectiveLast) errors.push(label + ': SMAX 다음 교정일은 적용할 교정일보다 뒤여야 합니다.');
            if (isoDate(tool.due) && update.due < tool.due) errors.push(label + ': SMAX 다음 교정일이 기존 다음 교정일보다 오래되었습니다.');
            if (model.date1904 && update.due < '1904-01-01') errors.push(label + ': 1904년 이전 날짜는 저장할 수 없습니다.');
            if (paired && paired.due !== update.due) errors.push(label + ': 검교정 사진과 SMAX 사진의 다음 교정일이 다릅니다.');
          }
        }
      }
      const bytes = update.photoBytes;
      if (!(bytes instanceof Uint8Array) || !bytes.length) errors.push(label + ': 반영할 사진이 없습니다.');
      else if (bytes.length > 25 * 1024 * 1024) errors.push(label + ': 사진 크기가 25 MB를 초과합니다.');
      else {
        const png = bytes.length >= 8 && [137,80,78,71,13,10,26,10].every((n, i) => bytes[i] === n);
        const jpg = bytes.length >= 3 && bytes[0] === 255 && bytes[1] === 216 && bytes[2] === 255;
        if (!((update.photoExtension === 'png' && png) || (['jpg', 'jpeg'].includes(update.photoExtension) && jpg))) errors.push(label + ': 사진은 정상적인 JPG 또는 PNG 파일이어야 합니다.');
      }
      if (!(Number(update.photoWidth) > 0 && Number(update.photoHeight) > 0) || Number(update.photoWidth) * Number(update.photoHeight) > 50000000) errors.push(label + ': 사진 해상도를 확인해 주세요. 5천만 픽셀 이하만 지원합니다.');
    }
    return Array.from(new Set(errors));
  }
  function append(doc, parent, namespace, tag, value, attrs) {
    const node = doc.createElementNS(namespace, tag);
    if (value != null) node.textContent = String(value);
    if (attrs) for (const [k, v] of Object.entries(attrs)) node.setAttribute(k, String(v));
    parent.appendChild(node);
    return node;
  }
  function freshRelationship(doc, type, target) {
    const ids = new Set(elements(doc, 'Relationship').map((r) => r.getAttribute('Id')));
    let index = 1;
    while (ids.has('rId' + index)) index++;
    const id = 'rId' + index;
    append(doc, doc.documentElement, NS.p, 'Relationship', null, { Id: id, Type: type, Target: target });
    return id;
  }
  function blankRelationships() { return xml('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="' + NS.p + '"/>', 'relationships'); }
  function setDateCell(sheet, rowNumber, col, serial, iso) {
    const row = sheet.rows.find((r) => Number(r.getAttribute('r')) === rowNumber), ref = columnLetter(col) + rowNumber;
    let cell = Array.from(row.children).find((c) => c.localName === 'c' && c.getAttribute('r') === ref);
    if (!cell) {
      cell = sheet.doc.createElementNS(NS.s, 'c'); cell.setAttribute('r', ref);
      const candidate = Array.from(row.children).find((c) => c.localName === 'c' && [sheet.headers['Cal Last'], sheet.headers['Cal Next']].includes(columnNumber(c.getAttribute('r'))) && c.hasAttribute('s'));
      if (candidate) cell.setAttribute('s', candidate.getAttribute('s'));
      const next = Array.from(row.children).find((c) => c.localName === 'c' && columnNumber(c.getAttribute('r')) > col);
      row.insertBefore(cell, next || null);
    }
    const useIso = ['s', 'inlineStr', 'str', 'd'].includes(cell.getAttribute('t')) || !cell.hasAttribute('s') || cell.getAttribute('s') === '0';
    if (useIso) cell.setAttribute('t', 'd'); else cell.removeAttribute('t');
    for (const node of Array.from(cell.children)) if (['v', 'is', 'f'].includes(node.localName)) cell.removeChild(node);
    append(sheet.doc, cell, NS.s, 'v', useIso ? iso : serial);
  }
  function ensureContentType(doc, extension) {
    if (!elements(doc, 'Default').some((n) => n.getAttribute('Extension').toLowerCase() === extension)) append(doc, doc.documentElement, NS.ct, 'Default', null, { Extension: extension, ContentType: extension === 'png' ? 'image/png' : 'image/jpeg' });
  }
  function ensureDrawingType(doc, path) {
    if (!elements(doc, 'Override').some((n) => n.getAttribute('PartName') === '/' + path)) append(doc, doc.documentElement, NS.ct, 'Override', null, { PartName: '/' + path, ContentType: 'application/vnd.openxmlformats-officedocument.drawing+xml' });
  }
  function fitSize(anchor, width, height) {
    let cx = 220 * EMU, cy = 140 * EMU;
    if (anchor) {
      const direct = child(anchor, 'ext'), xfrm = elements(anchor, 'xfrm')[0], ext = direct || xfrm && child(xfrm, 'ext');
      if (ext && Number(ext.getAttribute('cx')) > 0 && Number(ext.getAttribute('cy')) > 0) { cx = Number(ext.getAttribute('cx')); cy = Number(ext.getAttribute('cy')); }
    }
    const ratio = Math.min(cx / width, cy / height);
    return { cx: Math.max(1, Math.round(width * ratio)), cy: Math.max(1, Math.round(height * ratio)) };
  }
  function buildAnchor(doc, tool, col, relId, size, oldAnchor, kind) {
    const anchor = doc.createElementNS(NS.xdr, 'xdr:oneCellAnchor'), oldFrom = oldAnchor && child(oldAnchor, 'from');
    if (oldFrom) anchor.appendChild(oldFrom.cloneNode(true));
    else {
      const from = append(doc, anchor, NS.xdr, 'xdr:from');
      append(doc, from, NS.xdr, 'xdr:col', col - 1); append(doc, from, NS.xdr, 'xdr:colOff', 0);
      append(doc, from, NS.xdr, 'xdr:row', tool.row - 1); append(doc, from, NS.xdr, 'xdr:rowOff', 0);
    }
    append(doc, anchor, NS.xdr, 'xdr:ext', null, size);
    let pic;
    if (oldAnchor && elements(oldAnchor, 'pic')[0]) pic = elements(oldAnchor, 'pic')[0].cloneNode(true);
    else {
      pic = doc.createElementNS(NS.xdr, 'xdr:pic');
      const nv = append(doc, pic, NS.xdr, 'xdr:nvPicPr'), ids = elements(doc, 'cNvPr').map((n) => Number(n.getAttribute('id')) || 0);
      append(doc, nv, NS.xdr, 'xdr:cNvPr', null, { id: Math.max(0, ...ids) + 1, name: (kind === 'smax' ? 'Tool SMAX ' : 'Tool Label ') + tool.row });
      const cnv = append(doc, nv, NS.xdr, 'xdr:cNvPicPr'); append(doc, cnv, NS.a, 'a:picLocks', null, { noChangeAspect: 1 });
      const fill = append(doc, pic, NS.xdr, 'xdr:blipFill'); append(doc, fill, NS.a, 'a:blip');
      const stretch = append(doc, fill, NS.a, 'a:stretch'); append(doc, stretch, NS.a, 'a:fillRect');
      const sp = append(doc, pic, NS.xdr, 'xdr:spPr'); append(doc, sp, NS.a, 'a:xfrm');
      const geom = append(doc, sp, NS.a, 'a:prstGeom', null, { prst: 'rect' }); append(doc, geom, NS.a, 'a:avLst');
    }
    const blip = elements(pic, 'blip')[0];
    if (!blip) throw new Error('기존 사진의 엑셀 구조를 지원하지 않습니다.');
    blip.removeAttributeNS(NS.r, 'link'); blip.setAttributeNS(NS.r, 'r:embed', relId);
    for (const n of elements(pic, 'srcRect').concat(elements(blip, 'extLst'))) n.parentNode.removeChild(n);
    const xfrm = elements(pic, 'xfrm')[0];
    if (xfrm) {
      for (const attr of ['rot', 'flipH', 'flipV']) xfrm.removeAttribute(attr);
      let off = child(xfrm, 'off'); if (!off) off = append(doc, xfrm, NS.a, 'a:off'); off.setAttribute('x', '0'); off.setAttribute('y', '0');
      let ext = child(xfrm, 'ext'); if (!ext) ext = append(doc, xfrm, NS.a, 'a:ext'); ext.setAttribute('cx', String(size.cx)); ext.setAttribute('cy', String(size.cy));
    }
    anchor.appendChild(pic);
    const client = oldAnchor && child(oldAnchor, 'clientData');
    if (client) anchor.appendChild(client.cloneNode(true)); else append(doc, anchor, NS.xdr, 'xdr:clientData');
    return anchor;
  }
  async function apply(model, updates) {
    const problems = validate(model, updates);
    if (problems.length) throw new Error(problems.join('\n'));
    const fresh = await load(state.get(model).bytes), data = state.get(fresh), zip = data.zip;
    const types = await readXml(zip, '[Content_Types].xml'), originalTypes = serialize(types);
    const changedSheets = new Set(), changedDrawings = new Set();
    for (const update of updates) {
      const tool = fresh.tools.find((t) => t.id === update.toolId), sheet = data.sheets.find((s) => s.path + ':' + tool.row === tool.id);
      const kind = photoKind(update), photoColumn = sheet.headers[kind === 'smax' ? 'SMAX Tool Photo' : 'Label Photo'];
      // Calibration retains the legacy two-date update. A SMAX screenshot
      // cannot supply Cal Last, and photo-only SMAX updates touch neither date.
      if (kind === 'calibration') {
        setDateCell(sheet, tool.row, sheet.headers['Cal Last'], dateSerial(update.last, data.date1904), update.last);
        setDateCell(sheet, tool.row, sheet.headers['Cal Next'], dateSerial(update.due, data.date1904), update.due);
        changedSheets.add(sheet);
      } else if (update.updateDue === true) {
        setDateCell(sheet, tool.row, sheet.headers['Cal Next'], dateSerial(update.due, data.date1904), update.due);
        changedSheets.add(sheet);
      }
      let drawing, oldAnchor;
      for (const d of sheet.drawings) {
        const found = Array.from(d.doc.documentElement.children).find((anchor) => {
          const location = anchorLocation(anchor);
          return location && location.row === tool.row && location.col === photoColumn && elements(anchor, 'pic').length;
        });
        if (found) { drawing = d; oldAnchor = found; break; }
      }
      if (!drawing) drawing = sheet.drawings[0];
      if (!drawing) {
        let n = 1; while (zip.file('xl/drawings/drawing' + n + '.xml')) n++;
        const path = 'xl/drawings/drawing' + n + '.xml';
        drawing = { path, doc: xml('<xdr:wsDr xmlns:xdr="' + NS.xdr + '" xmlns:a="' + NS.a + '"/>', path), relDoc: blankRelationships() };
        zip.file(path, serialize(drawing.doc)); sheet.drawings.push(drawing);
        if (!sheet.relDoc) sheet.relDoc = blankRelationships();
        const id = freshRelationship(sheet.relDoc, REL + 'drawing', relative(sheet.path, path)), node = sheet.doc.createElementNS(NS.s, 'drawing');
        node.setAttributeNS(NS.r, 'r:id', id);
        const before = Array.from(sheet.doc.documentElement.children).find((c) => ['legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'].includes(c.localName));
        sheet.doc.documentElement.insertBefore(node, before || null); zip.file(relPath(sheet.path), serialize(sheet.relDoc)); ensureDrawingType(types, path);
        changedSheets.add(sheet);
      }
      if (!drawing.relDoc) drawing.relDoc = blankRelationships();
      const extension = update.photoExtension === 'jpeg' ? 'jpg' : update.photoExtension;
      const mediaPrefix = kind === 'smax' ? 'tool-smax-' : 'tool-label-';
      let imageIndex = 1; while (zip.file('xl/media/' + mediaPrefix + imageIndex + '.' + extension)) imageIndex++;
      const imagePath = 'xl/media/' + mediaPrefix + imageIndex + '.' + extension;
      zip.file(imagePath, update.photoBytes); ensureContentType(types, extension);
      const newId = freshRelationship(drawing.relDoc, REL + 'image', relative(drawing.path, imagePath));
      const size = fitSize(oldAnchor, Number(update.photoWidth), Number(update.photoHeight)), anchor = buildAnchor(drawing.doc, tool, photoColumn, newId, size, oldAnchor, kind);
      if (oldAnchor) oldAnchor.parentNode.replaceChild(anchor, oldAnchor);
      else {
        drawing.doc.documentElement.appendChild(anchor);
        const row = sheet.rows.find((r) => Number(r.getAttribute('r')) === tool.row), oldHeight = Number(row.getAttribute('ht') || 15);
        if (oldHeight < size.cy / EMU * 0.75) { row.setAttribute('ht', String(Math.ceil(size.cy / EMU * 0.75))); row.setAttribute('customHeight', '1'); changedSheets.add(sheet); }
      }
      changedDrawings.add(drawing);
    }
    for (const sheet of changedSheets) zip.file(sheet.path, serialize(sheet.doc));
    for (const drawing of changedDrawings) { zip.file(drawing.path, serialize(drawing.doc)); zip.file(relPath(drawing.path), serialize(drawing.relDoc)); }
    if (serialize(types) !== originalTypes) zip.file('[Content_Types].xml', serialize(types));
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  global.ToolWorkbook = Object.freeze({ load, normalizeSerial, normalizeName, addMonths, monthInterval, validate, apply });
})(window);
