/* Reapply one client's verified photos to a newer shared workbook.
 * Only independent edits merge automatically. A changed target photo or a
 * conflicting date/identity is held for review rather than overwritten.
 */
(function (global) {
  'use strict';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XMLNS = 'http://www.w3.org/2000/xmlns/';
  const headers = ['Cal Next', 'Cal Last', 'Serial', 'Asset #', 'Label Photo', 'SMAX Tool Photo'];
  const required = headers.slice(0, 5);
  const list = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
  const ownChild = (node, name) => Array.from(node.children || []).find(child => child.localName === name);
  const normalizeHeader = value => String(value || '').trim().replace(/\s+/g, ' ').toLowerCase();
  const kindOf = update => update.photoKind === undefined ? 'calibration' : update.photoKind;
  function conflict(message) { const error = new Error(message); error.code = 'shared-rebase-conflict'; return error; }
  function invalid(message) { const error = new Error(message); error.code = 'shared-rebase-invalid'; return error; }
  function bytes(value, label) {
    const result = value instanceof Uint8Array ? value.slice() : value instanceof ArrayBuffer ? new Uint8Array(value.slice(0)) : null;
    if (!result || !result.length || result.length > 100 * 1024 * 1024) throw invalid(label + ' 엑셀의 원본 데이터를 확인할 수 없습니다.');
    return result;
  }
  function parse(text) {
    if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw invalid('지원하지 않는 엑셀 XML입니다.');
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (!doc.documentElement || doc.documentElement.localName === 'parsererror' || list(doc, 'parsererror').length) throw invalid('엑셀 내부 사진 정보를 읽을 수 없습니다.');
    return doc;
  }
  function resolve(base, target) {
    if (!target || /[?#\\]/.test(target) || /^[a-z]+:/i.test(target)) throw invalid('엑셀 내부 사진 연결이 올바르지 않습니다.');
    const path = (target.startsWith('/') ? target.slice(1) : base.slice(0, base.lastIndexOf('/') + 1) + target).split('/'), result = [];
    for (const piece of path) {
      if (!piece || piece === '.') continue;
      if (piece === '..') { if (!result.length) throw invalid('엑셀 내부 사진 경로가 올바르지 않습니다.'); result.pop(); }
      else result.push(piece);
    }
    return result.join('/');
  }
  const relPath = part => part.slice(0, part.lastIndexOf('/') + 1) + '_rels/' + part.slice(part.lastIndexOf('/') + 1) + '.rels';
  const relationId = (node, key) => node.getAttributeNS(REL, key || 'id') || node.getAttribute('r:' + (key || 'id'));
  const column = ref => Array.from((String(ref).match(/^[A-Z]+/i) || [''])[0].toUpperCase()).reduce((value, ch) => value * 26 + ch.charCodeAt(0) - 64, 0);
  async function packageOf(source) {
    const zip = await global.JSZip.loadAsync(source);
    const docs = new Map(), maps = new Map();
    async function document(part, optional) {
      if (docs.has(part)) return docs.get(part);
      const file = zip.file(part);
      if (!file) { if (optional) return null; throw invalid('엑셀 사진 구성 파일이 없습니다: ' + part); }
      const value = parse(await file.async('string')); docs.set(part, value); return value;
    }
    async function relationships(part) {
      if (maps.has(part)) return maps.get(part);
      const doc = await document(relPath(part), true), result = new Map();
      if (doc) for (const node of list(doc, 'Relationship')) result.set(node.getAttribute('Id'), node);
      maps.set(part, result); return result;
    }
    const workbookRels = await relationships('xl/workbook.xml');
    const stringsRel = Array.from(workbookRels.values()).find(node => (node.getAttribute('Type') || '').endsWith('/sharedStrings'));
    const stringsDoc = stringsRel ? await document(resolve('xl/workbook.xml', stringsRel.getAttribute('Target'))) : null;
    const richText = node => list(node, 't').filter(child => child.parentElement?.localName !== 'rPh').map(child => child.textContent).join('');
    const strings = stringsDoc ? list(stringsDoc, 'si').map(richText) : [];
    const cellValue = cell => cell.getAttribute('t') === 's' ? strings[Number(ownChild(cell, 'v')?.textContent)] || '' : cell.getAttribute('t') === 'inlineStr' ? richText(cell) : ownChild(cell, 'v')?.textContent || '';
    const sheetHeaders = new Map();
    async function photoColumn(part, kind) {
      if (!sheetHeaders.has(part)) {
        const doc = await document(part), rows = list(doc, 'sheetData').flatMap(node => Array.from(node.children).filter(row => row.localName === 'row'));
        let found = null;
        for (const row of rows.slice(0, 20)) {
          const candidate = {};
          for (const cell of Array.from(row.children).filter(child => child.localName === 'c')) {
            const heading = headers.find(value => normalizeHeader(value) === normalizeHeader(cellValue(cell)));
            if (heading) { if (candidate[heading]) throw invalid('엑셀 사진 열 이름이 중복됐습니다.'); candidate[heading] = column(cell.getAttribute('r')); }
          }
          if (required.every(heading => candidate[heading])) { found = candidate; break; }
        }
        if (!found) throw invalid('엑셀 사진 열을 찾지 못했습니다.');
        sheetHeaders.set(part, found);
      }
      const result = sheetHeaders.get(part)[kind === 'smax' ? 'SMAX Tool Photo' : 'Label Photo'];
      if (!result) throw conflict('현재 공유 엑셀에 반영할 사진 열이 없습니다.');
      return result;
    }
    async function fingerprint(tool, kind) {
      const marker = tool.id.lastIndexOf(':'), part = tool.id.slice(0, marker), col = await photoColumn(part, kind);
      const doc = await document(part), sheetRels = await relationships(part), matches = [];
      for (const node of list(doc, 'drawing')) {
        const relation = sheetRels.get(relationId(node));
        if (!relation || relation.getAttribute('TargetMode') === 'External' || !(relation.getAttribute('Type') || '').endsWith('/drawing')) continue;
        const drawingPart = resolve(part, relation.getAttribute('Target')), drawing = await document(drawingPart);
        for (const anchor of Array.from(drawing.documentElement.children)) {
          const from = ownChild(anchor, 'from');
          if (!from || Number(ownChild(from, 'row')?.textContent) + 1 !== tool.row || Number(ownChild(from, 'col')?.textContent) + 1 !== col || !list(anchor, 'pic').length) continue;
          matches.push({ anchor, drawingPart });
        }
      }
      if (!matches.length) return 'empty';
      if (matches.length !== 1 || list(matches[0].anchor, 'pic').length !== 1) throw conflict('사진 칸에 여러 사진이 있어 자동 병합할 수 없습니다.');
      const { anchor, drawingPart } = matches[0], blips = list(anchor, 'blip');
      if (blips.length !== 1 || relationId(blips[0], 'link')) throw conflict('연결된 사진의 구조를 자동 병합할 수 없습니다.');
      const relation = (await relationships(drawingPart)).get(relationId(blips[0], 'embed'));
      if (!relation || relation.getAttribute('TargetMode') === 'External' || !(relation.getAttribute('Type') || '').endsWith('/image')) throw conflict('원본 사진의 내부 연결을 확인할 수 없습니다.');
      const file = zip.file(resolve(drawingPart, relation.getAttribute('Target')));
      if (!file) throw invalid('원본 사진 파일을 찾지 못했습니다.');
      const content = await file.async('uint8array');
      const digest = Array.from(new Uint8Array(await global.crypto.subtle.digest('SHA-256', content)), value => value.toString(16).padStart(2, '0')).join('');
      function shape(element) {
        const attributes = Array.from(element.attributes || []).filter(attr => attr.namespaceURI !== XMLNS && !(attr.namespaceURI === REL && ['embed', 'link'].includes(attr.localName))).map(attr => [attr.namespaceURI || '', attr.localName, attr.value]).sort((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
        let text = Array.from(element.childNodes).filter(child => child.nodeType === 3 || child.nodeType === 4).map(child => child.textContent).join('').trim();
        if (['from', 'to'].includes(element.parentElement?.localName) && ['row', 'col'].includes(element.localName)) text = String(Number(text) - (element.localName === 'row' ? tool.row - 1 : col - 1));
        return [element.namespaceURI || '', element.localName, attributes, text, Array.from(element.children || []).map(shape)];
      }
      return digest + ':' + JSON.stringify(shape(anchor));
    }
    return { fingerprint };
  }
  function samePolicy(old, current) {
    if (old.months !== null || current.months !== null) return old.months === current.months;
    return String(old.interval) === String(current.interval) && normalizeHeader(old.cycle) === normalizeHeader(current.cycle);
  }
  function matches(model, original) {
    const serial = global.ToolWorkbook.normalizeSerial(original.serial);
    const candidates = model.tools.filter(tool => global.ToolWorkbook.normalizeSerial(tool.serial) === serial);
    if (candidates.length !== 1 || candidates[0].name !== original.name) throw conflict('공유 엑셀의 담당자 또는 시리얼이 변경됐거나 중복됐습니다. 최신 항목을 확인해 주세요.');
    return candidates[0];
  }
  async function rebase(options) {
    const workbook = global.ToolWorkbook;
    if (!workbook || !global.JSZip || !global.crypto?.subtle) throw invalid('공유 엑셀 병합 모듈을 준비하지 못했습니다.');
    const baselineModel = options?.baselineModel, incoming = options?.updates;
    if (!baselineModel || !Array.isArray(baselineModel.tools) || !Array.isArray(incoming) || !incoming.length) throw invalid('병합할 원본과 사진 정보를 확인해 주세요.');
    const updates = incoming.map(update => ({ ...update, photoBytes: update?.photoBytes instanceof Uint8Array ? update.photoBytes.slice() : update?.photoBytes }));
    const initialProblems = workbook.validate(baselineModel, updates);
    if (initialProblems.length) throw conflict(initialProblems.join('\n'));
    const oldBytes = bytes(options.baselineBytes, '이전'), currentBytes = bytes(options.currentBytes, '최신');
    const [baseline, current] = await Promise.all([workbook.load(oldBytes), workbook.load(currentBytes)]);
    const [oldPackage, newPackage] = await Promise.all([packageOf(oldBytes), packageOf(currentBytes)]);
    const rebound = [];
    for (const input of updates) {
      const original = baselineModel.tools.find(tool => tool.id === input.toolId), old = baseline.tools.find(tool => tool.id === input.toolId);
      if (!original || !old || ['serial', 'name', 'last', 'due', 'months', 'photoCount', 'smaxPhotoCount'].some(field => original[field] !== old[field])) throw invalid('이전 엑셀과 사진 작업의 기준이 다릅니다. 최신 원본을 다시 확인해 주세요.');
      const target = matches(current, old), kind = kindOf(input);
      if (!samePolicy(old, target)) throw conflict('공유 엑셀의 교정 주기가 변경됐습니다. 날짜를 다시 확인해 주세요.');
      if (/폐기|disposed|retired/i.test(target.due)) throw conflict('폐기된 툴에 사진을 병합할 수 없습니다.');
      const fields = kind === 'calibration' ? ['last', 'due'] : input.updateDue === true ? ['due'] : [];
      for (const field of fields) if (target[field] !== old[field] && target[field] !== input[field]) throw conflict('같은 툴의 교정 날짜를 다른 사용자가 변경했습니다. 최신 내용을 확인해 주세요.');
      const [previousPhoto, currentPhoto] = await Promise.all([oldPackage.fingerprint(old, kind), newPackage.fingerprint(target, kind)]);
      if (previousPhoto !== currentPhoto) throw conflict('같은 툴의 사진 칸을 다른 사용자가 변경했습니다. 사진을 덮어쓰지 않았습니다.');
      const update = { ...input, toolId: target.id, name: target.name, serial: target.serial, photoKind: kind, photoBytes: input.photoBytes.slice() };
      if (kind === 'smax') { update.last = target.last; if (input.updateDue !== true) update.due = target.due; }
      rebound.push(update);
    }
    const problems = workbook.validate(current, rebound);
    if (problems.length) throw conflict(problems.join('\n'));
    const result = await workbook.apply(current, rebound), model = await workbook.load(result);
    return { bytes: result, model, updates: rebound };
  }
  global.ToolSharedMerge = Object.freeze({ rebase });
})(window);
