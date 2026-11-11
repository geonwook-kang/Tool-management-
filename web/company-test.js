/* Separate, session-only company workbook test. Never connect or write an original. */
(function (global) {
  'use strict';
  if (!document.querySelector('meta[name="tool-manager-company-test-entry"]')) return;
  const $ = id => document.getElementById(id);
  const clock = () => performance.now();
  const seconds = began => Math.round((clock() - began) / 10) / 100;
  const copy = value => JSON.parse(JSON.stringify(value));
  const progress = { version: '7.2-T1', originalWriteEnabled: false, persistentWorkbookStorage: false,
    companyCloudSyncTested: false, input: null, load: null, conversion: null, lastError: '' };
  let loadedModel = null, originalFile = null, uiReady = false, scheduled = false;
  function denied() { throw new Error('전사 테스트판은 원본 연결·직접 저장을 사용하지 않습니다. 파일을 불러온 뒤 테스트 결과를 다운로드하세요.'); }
  const disconnected = () => ({ connected: false, persistent: false, supported: true, sourceMode: 'device',
    mobile: false, hosted: false, directEntry: true, sharedEntry: false, pcEntry: true, companyTestEntry: true,
    needsPermission: false, needsReconnect: false, backupAvailable: false, name: '', version: '' });
  global.ToolSources = Object.freeze({ mode: () => 'device', inspect: disconnected,
    use: mode => { if (mode !== 'device') denied(); return disconnected(); },
    restore: async () => disconnected(), disconnect: async () => disconnected(), takeBackup: async () => null,
    connect: denied, requestAccess: denied, read: denied, save: denied });

  function summary(model, bytes) {
    const meta = model.companyTest || {}, numbers = new Map();
    for (const tool of model.tools) {
      const serial = global.ToolWorkbook.normalizeSerial(tool.serial);
      if (!serial) continue;
      if (!numbers.has(serial)) numbers.set(serial, []);
      numbers.get(serial).push({ sheet: tool.sheet, row: tool.row });
    }
    const duplicates = [...numbers].filter(([, rows]) => rows.length > 1)
      .map(([serial, rows]) => ({ serial, count: rows.length, rows }));
    return { bytes: Number(meta.inputBytes) || bytes.byteLength, allSheetNames: meta.allSheetNames || [],
      supportedSheetNames: model.sheets.slice(), toolCount: model.tools.length, mediaCount: meta.mediaCount || 0,
      duplicateSerialCount: duplicates.length, duplicates, warnings: model.warnings.slice() };
  }
  const originalWorkbook = global.ToolWorkbook;
  const relationshipPath = path => path.slice(0, path.lastIndexOf('/') + 1) + '_rels/' + path.slice(path.lastIndexOf('/') + 1) + '.rels';
  function resolve(base, target) {
    if (!target || /[?#\\]/.test(target) || /^[a-z]+:/i.test(target)) throw new Error('결과 파일의 내부 연결을 확인할 수 없습니다.');
    const out = [];
    for (const part of (target.startsWith('/') ? target.slice(1) : base.slice(0, base.lastIndexOf('/') + 1) + target).split('/')) {
      if (!part || part === '.') continue;
      if (part === '..') { if (!out.length) throw new Error('결과 파일의 내부 경로를 확인할 수 없습니다.'); out.pop(); }
      else out.push(part);
    }
    return out.join('/');
  }
  function xml(text) {
    if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('결과 파일에 지원하지 않는 XML이 있습니다.');
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagNameNS('*', 'parsererror').length) throw new Error('결과 파일의 XML을 읽지 못했습니다.');
    return doc;
  }
  async function audit(model, updates, bytes) {
    const meta = model.companyTest;
    if (!meta || !meta.parts || !Array.isArray(meta.sheetParts)) throw new Error('테스트 파일의 원본 보존 정보를 준비하지 못했습니다.');
    const zip = await global.JSZip.loadAsync(bytes), allowed = new Set(['[Content_Types].xml']), selectedDrawings = new Set();
    const selectedPaths = new Set(updates.map(update => model.tools.find(tool => tool.id === update.toolId)?.id.split(':')[0]));
    for (const sheet of meta.sheetParts.filter(sheet => selectedPaths.has(sheet.path))) {
      allowed.add(sheet.path); allowed.add(relationshipPath(sheet.path));
      for (const drawing of sheet.drawings) { allowed.add(drawing); allowed.add(relationshipPath(drawing)); selectedDrawings.add(drawing); }
      // A sheet without a previous picture may acquire a drawing. Permit only
      // drawing parts actually referenced by this selected sheet's output XML.
      const sheetEntry = zip.file(sheet.path), relEntry = zip.file(relationshipPath(sheet.path));
      if (!sheetEntry || !relEntry) continue;
      const sheetDoc = xml(await sheetEntry.async('string')), relDoc = xml(await relEntry.async('string'));
      const ids = new Set([...sheetDoc.getElementsByTagNameNS('*', 'drawing')].map(node => node.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || node.getAttribute('r:id')));
      for (const rel of relDoc.getElementsByTagNameNS('*', 'Relationship')) {
        if (!ids.has(rel.getAttribute('Id')) || rel.getAttribute('TargetMode') === 'External' || !/\/drawing$/.test(rel.getAttribute('Type') || '')) continue;
        const path = resolve(sheet.path, rel.getAttribute('Target'));
        if (!Object.prototype.hasOwnProperty.call(meta.parts, path) && !/^xl\/drawings\/drawing\d+\.xml$/.test(path)) throw new Error('새 사진의 도형 연결을 확인할 수 없습니다.');
        allowed.add(path); allowed.add(relationshipPath(path)); selectedDrawings.add(path);
      }
    }
    const changedParts = [], addedParts = [], removedParts = [], unexpectedParts = [];
    let unchangedPartCount = 0, originalPartCount = 0;
    for (const [name, before] of Object.entries(meta.parts)) {
      if (name.endsWith('/')) continue;
      originalPartCount++;
      const entry = zip.file(name);
      if (!entry) { removedParts.push(name); unexpectedParts.push(name); continue; }
      const after = entry._data;
      if (!after || !Number.isFinite(before.crc32) || !Number.isFinite(before.bytes)) throw new Error('파일 구성 요소의 보존 정보를 읽지 못했습니다.');
      if ((after.crc32 >>> 0) === (before.crc32 >>> 0) && Number(after.uncompressedSize) === Number(before.bytes)) unchangedPartCount++;
      else { changedParts.push(name); if (!allowed.has(name)) unexpectedParts.push(name); }
    }
    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir || Object.prototype.hasOwnProperty.call(meta.parts, name)) continue;
      addedParts.push(name);
      if (!allowed.has(name) && !/^xl\/media\/tool-(?:label|smax)-\d+\.(?:jpg|jpeg|png)$/i.test(name)) unexpectedParts.push(name);
    }
    // Inspect all other worksheet relationships, including sheets that have
    // no tool-management columns and therefore are absent from sheetParts.
    // Only small relationship XML is inflated; image bytes are not re-read.
    const changed = new Set(changedParts.concat(addedParts));
    const touchedDrawings = new Set([...selectedDrawings].filter(path => changed.has(path) || changed.has(relationshipPath(path))));
    if (touchedDrawings.size) for (const path of Object.keys(meta.parts).filter(name => /^xl\/worksheets\/[^/]+\.xml$/.test(name) && !selectedPaths.has(name))) {
      const relEntry = zip.file(relationshipPath(path));
      if (!relEntry) continue;
      const relDoc = xml(await relEntry.async('string'));
      for (const rel of relDoc.getElementsByTagNameNS('*', 'Relationship')) {
        if (rel.getAttribute('TargetMode') === 'External' || !/\/drawing$/.test(rel.getAttribute('Type') || '')) continue;
        const drawing = resolve(path, rel.getAttribute('Target'));
        if (touchedDrawings.has(drawing)) unexpectedParts.push(drawing);
      }
    }
    return { outputBytes: bytes.byteLength, originalPartCount, unchangedPartCount, changedParts, addedParts,
      removedParts, unexpectedParts: [...new Set(unexpectedParts)], verified: unexpectedParts.length === 0,
      verificationMethod: 'Original input CRC checked by workbook engine; output central-directory CRC32 and uncompressed sizes compared for every retained part; selected XML relationships inspected.' };
  }
  global.ToolWorkbook = Object.freeze({ ...originalWorkbook,
    load: async input => {
      const began = clock(); progress.lastError = ''; progress.conversion = null;
      const selectedFile = $('workbook')?.files[0]; originalFile = null;
      try {
        const model = await originalWorkbook.load(input);
        loadedModel = model; progress.input = summary(model, input); progress.load = { seconds: seconds(began), completed: true };
        if (selectedFile && selectedFile.size === progress.input.bytes) originalFile = selectedFile;
        schedule(); return model;
      } catch (error) { progress.load = { seconds: seconds(began), completed: false }; progress.lastError = error.message; schedule(); throw error; }
    },
    apply: async (model, updates) => {
      const began = clock(); progress.lastError = ''; progress.conversion = { completed: false, phase: '사진·날짜를 반영하고 있습니다.' }; schedule();
      try {
        const result = await originalWorkbook.apply(model, updates);
        progress.conversion.phase = '변경 대상 외 파일 보존 상태를 검사하고 있습니다.'; schedule();
        const check = await audit(model, updates, result);
        progress.conversion = { ...check, seconds: seconds(began), completed: check.verified, phase: check.verified ? '테스트 결과 다운로드 준비 완료' : '보존 검사 실패' };
        if (!check.verified) throw new Error('변경 대상 외 파일 내용이 달라 테스트 결과를 생성하지 않았습니다. 테스트 결과 보고서를 확인해 주세요.');
        schedule(); return result;
      } catch (error) { progress.lastError = error.message; progress.conversion = { ...progress.conversion, completed: false, seconds: seconds(began) }; schedule(); throw error; }
    }
  });

  function report() {
    const app = global.ToolApp?.state;
    return copy({ ...progress, createdAt: new Date().toISOString(), workbook: app?.workbookName || '',
      testResultReady: !!(app?.preparedResult && progress.conversion?.verified),
      photos: (app?.photos || []).map(photo => ({ filename: photo.file.name, approval: photo.approval || null,
        assessment: photo.assessment || null, error: photo.error || '', photoKind: photo.review?.photoKind || photo.assessment?.photoKind || null })),
      note: '이 보고서는 이 PC에서 직접 선택한 파일의 테스트 결과입니다. 회사 공유 원본은 수정하지 않았고, 회사 서버 동기화나 모든 사진의 인식 성공을 검증한 결과는 아닙니다.' });
  }
  function downloadReport() {
    const blob = new Blob([JSON.stringify(report(), null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob), link = document.createElement('a');
    link.href = url; link.download = 'Tool-Company-Test-Report.json'; document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  function downloadOriginalCopy() {
    if (!originalFile || !global.ToolApp?.state || global.ToolApp.state.model !== loadedModel) return;
    const url = URL.createObjectURL(originalFile), link = document.createElement('a');
    link.href = url; link.download = originalFile.name.replace(/\.xlsx$/i, '') + '_전사테스트.xlsx';
    document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  global.ToolCompanyTest = Object.freeze({ report, downloadReport, audit });
  function setText(element, value) { if (element && element.textContent !== value) element.textContent = value; }
  function setHidden(element, value) { if (element && element.hidden !== value) element.hidden = value; }
  function setDisabled(element, value) { if (element && element.disabled !== value) element.disabled = value; }
  function draw() {
    if (!uiReady) return;
    const app = global.ToolApp?.state, meta = progress.input, busy = !!(app?.busy || app?.preparing || app?.source?.saving);
    setText(document.querySelector('main>p.muted'), 'CALIBRATION WORKSPACE · 전사 파일 테스트판 7.2-T1');
    setText(document.querySelector('h1'), '회사 전체 파일로, 별도 사본 테스트.');
    setText(document.querySelector('h1+p'), '파일과 사진을 이 PC에서 처리하고 테스트 결과를 다운로드합니다. 원본 파일에는 쓰지 않습니다.');
    setText(document.querySelector('header .tag'), 'PC · 원본 수정 없는 테스트');
    setText($('quick-workbook'), app?.model ? '다른 전사 파일 불러오기' : '전사 엑셀 파일 불러오기');
    $('quick-workbook').onclick = () => { if (!busy) { $('workbook').value = ''; $('workbook').click(); } };
    setDisabled($('quick-workbook'), busy);
    setText($('quick-summary'), app?.model ? app.workbookName + ' · 이번 창에서 테스트 중' : '회사 전체 .xlsx 파일을 선택하세요. 최대 250 MiB까지 불러올 수 있습니다.');
    setText($('quick-save-note'), '결과는 별도의 _전사테스트 엑셀로 다운로드합니다. 창을 닫으면 작업 내용이 사라지므로 결과를 먼저 받으세요.');
    setText($('source-help'), '선택한 파일은 이번 창의 임시 사본으로만 처리합니다. 결과 엑셀을 다운로드해 보관하세요.');
    for (const id of ['connection-settings', 'quick-clear-workbook', 'quick-download-workbook', 'connect-onedrive', 'reconnect-onedrive', 'disconnect-onedrive', 'retry-connected-save', 'download-backup', 'm365-panel', 'shared-panel', 'sharepoint-panel', 'source-save-note', 'connection-status']) setHidden($(id), true);
    for (const option of $('source-mode').options) if (option.value !== 'device') { setHidden(option, true); setDisabled(option, true); }
    const memory = $('remember-matches');
    if (memory.checked) memory.checked = false;
    setDisabled(memory, true); setHidden(memory.closest('.panel'), true);
    memory.onchange = () => { memory.checked = false; };
    setText(document.querySelector('.export-bar > div > p'), '다운로드한 테스트 결과를 Excel에서 확인하세요. 회사 원본과 공유 파일은 자동으로 변경되지 않습니다.');
    if (!busy && !app?.preparationError) setText($('export'), '전사 테스트 결과 엑셀 받기 ↓');
    if (app?.preparedResult) setText($('automatic-result-download'), '전사 테스트 결과 엑셀 받기 ↓');
    if (!app?.model && !busy && !$('status').classList.contains('error')) setText($('status'), '전사 엑셀 파일을 불러온 뒤 툴 사진을 추가하세요. 이 창에서는 원본에 저장하지 않습니다.');
    setText(document.querySelector('footer'), 'TOOL MANAGER · 전사 파일 테스트 7.2-T1 · 이 창의 임시 사본만 처리합니다. 결과 엑셀을 다운로드해 보관하세요.');
    setText($('company-test-summary'), meta ?
      `${(meta.bytes / 1048576).toFixed(1)} MiB · 전체 ${meta.allSheetNames.length}개 시트 / 툴 관리 ${meta.supportedSheetNames.length}개 시트 · ${meta.toolCount}개 툴 · 기존 이미지 ${meta.mediaCount}개 · 중복 시리얼 ${meta.duplicateSerialCount}개` :
      '1. 전사 엑셀 불러오기 → 2. 사진 추가 → 3. 결과 엑셀 다운로드 → 4. Excel에서 확인');
    const phase = busy ? progress.conversion?.phase || '파일 또는 사진을 처리 중입니다. 큰 파일은 인식 외에 검사·변환 시간이 추가됩니다.' :
      progress.lastError ? '테스트 오류: ' + progress.lastError :
      app?.preparedResult && progress.conversion?.verified ? `변환·보존 검사 완료 · 변경 대상 외 ${progress.conversion.unchangedPartCount}개 구성 요소 확인 · 결과를 다운로드하세요.` :
      app?.model ? '파일을 불러왔습니다. 사진을 추가하면 인식과 변환을 검사합니다.' : '파일은 회사 서버로 전송하지 않으며 이전 PC판의 파일 연결을 복원하지 않습니다.';
    setText($('company-test-audit'), phase);
    setDisabled($('company-test-report'), busy || !meta && !progress.lastError);
    setHidden($('company-test-copy'), !originalFile || app?.model !== loadedModel);
    setDisabled($('company-test-copy'), busy);
    const warning = $('company-test-warnings');
    const warningText = meta ? meta.warnings.join('\n') + (meta.duplicateSerialCount ? '\n같은 시리얼이 여러 행에 있으면 자동 반영이 보류될 수 있습니다.' : '') : '';
    setText(warning, warningText.trim()); setHidden(warning, !warningText.trim());
  }
  function schedule() {
    if (!uiReady || scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => { scheduled = false; draw(); });
  }
  function install() {
    const panel = document.createElement('section'); panel.className = 'panel'; panel.id = 'company-test-panel';
    const heading = document.createElement('h2'); heading.textContent = '전사 파일 변환 테스트';
    const intro = document.createElement('p'); intro.className = 'muted'; intro.textContent = '현재 창에서만 사본을 처리합니다. 기존 운영판·회사 원본·OneDrive 파일에 직접 저장하지 않습니다.';
    const summary = document.createElement('p'); summary.id = 'company-test-summary'; summary.className = 'summary';
    const audit = document.createElement('p'); audit.id = 'company-test-audit'; audit.className = 'summary'; audit.setAttribute('role', 'status');
    const warnings = document.createElement('p'); warnings.id = 'company-test-warnings'; warnings.className = 'warning'; warnings.style.whiteSpace = 'pre-wrap'; warnings.hidden = true;
    const button = document.createElement('button'); button.id = 'company-test-report'; button.type = 'button'; button.textContent = '테스트 결과 보고서 받기'; button.onclick = downloadReport; button.disabled = true;
    const clone = document.createElement('button'); clone.id = 'company-test-copy'; clone.type = 'button'; clone.textContent = '원본 그대로 테스트 사본 받기'; clone.onclick = downloadOriginalCopy; clone.hidden = true; clone.style.marginRight = '8px';
    panel.append(heading, intro, summary, audit, warnings, clone, button); $('quick-start').before(panel);
    uiReady = true;
    const observer = new MutationObserver(schedule);
    observer.observe(document.querySelector('main'), { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ['disabled', 'hidden'] });
    draw();
  }
  document.addEventListener('DOMContentLoaded', () => setTimeout(install, 0));
})(window);
