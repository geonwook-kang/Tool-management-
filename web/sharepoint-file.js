/* Same-origin SharePoint REST source for the SharePoint-hosted edition.
 * The browser's existing company session is used; no OAuth token is copied.
 * A public/GitHub page cannot use this adapter against another origin.
 */
(function (global) {
  'use strict';
  const DATABASE = 'tool-management-sharepoint-file-v1', LOCK = 'tool-management-sharepoint-workbook';
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', LIMIT = 100 * 1024 * 1024;
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  let selected = null, backup = null, cached = null, databasePromise = null, queue = Promise.resolve(), pendingLock = null;
  const initial = () => ({ mode: 'sharepoint', connected: false, name: '', needsPermission: false, needsReconnect: false,
    requiresSignIn: false, persistent: false, busy: false, uncertain: false, version: '', lastSavedAt: '', statusMessage: '' });
  let status = initial();
  function failure(code, message, cause) { const error = new Error(message); error.code = code; if (cause) error.cause = cause; return error; }
  function origin() { return typeof global.origin === 'string' && global.origin !== 'null' ? global.origin : global.location.origin; }
  function siteValue(raw) {
    let url; try { url = new URL(String(raw || '')); } catch (_) { throw failure('invalid-site', 'SharePoint 사이트 주소를 확인해 주세요.'); }
    const local = url.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
    if (url.origin !== origin() || !(url.protocol === 'https:' || local) || url.username || url.password || url.search || url.hash || /[\u0000-\u001f\u007f]/.test(url.href)) throw failure('wrong-origin', '이 연결은 회사 SharePoint 안의 툴 매니저 페이지에서 사용할 수 있습니다.');
    let path; try { path = decodeURIComponent(url.pathname); } catch (_) { throw failure('invalid-site', 'SharePoint 사이트 주소를 확인해 주세요.'); }
    if (path.split('/').some(part => part === '.' || part === '..') || /[\\\u0000-\u001f\u007f]/.test(path) || /\/_api(?:\/|$)/i.test(path)) throw failure('invalid-site', 'SharePoint 사이트 주소를 확인해 주세요.');
    url.pathname = url.pathname.replace(/\/+$/, ''); return url.href.replace(/\/$/, '');
  }
  function context() {
    const value = global.ToolSharePointContext;
    if (!value || typeof value !== 'object' || !value.siteUrl) throw failure('host-required', '회사 SharePoint에 추가한 툴 매니저에서 이 연결을 열어 주세요.');
    const siteUrl = siteValue(value.siteUrl); return { siteUrl, fileUrl: typeof value.fileUrl === 'string' ? value.fileUrl : '' };
  }
  function supported() { try { context(); return global.isSecureContext !== false && !!(global.fetch && global.crypto?.subtle && global.indexedDB); } catch (_) { return false; } }
  function inspect() { return { ...status, supported: supported(), backupAvailable: !!backup, pendingCheckout: !!pendingLock, target: selected ? { ...selected } : null }; }
  const complete = extra => ({ ...inspect(), busy: false, ...(extra || {}) });
  function serial(action) { const task = queue.then(async () => { status.busy = true; try { return await action(); } finally { status.busy = false; } }); queue = task.catch(() => {}); return task; }
  function openDatabase() {
    if (databasePromise) return databasePromise;
    databasePromise = new Promise((resolve, reject) => {
      if (!global.indexedDB) return reject(failure('storage-unavailable', '저장 전 백업을 보관할 브라우저 저장소가 없습니다.'));
      let request; try { request = global.indexedDB.open(DATABASE, 2); } catch (error) { return reject(error); }
      request.onupgradeneeded = () => { for (const store of ['targets', 'backups', 'locks']) if (!request.result.objectStoreNames.contains(store)) request.result.createObjectStore(store); };
      request.onsuccess = () => { const db = request.result; db.onversionchange = () => { db.close(); databasePromise = null; }; resolve(db); };
      request.onerror = () => reject(failure('storage-unavailable', 'SharePoint 연결·백업 저장소를 열지 못했습니다.', request.error));
      request.onblocked = () => reject(failure('storage-unavailable', '다른 툴 매니저 창을 닫고 다시 시도해 주세요.'));
    }); databasePromise.catch(() => { databasePromise = null; }); return databasePromise;
  }
  async function stored(store, operation, key, value) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      let tx, request, result;
      try {
        try { tx = db.transaction(store, operation === 'get' ? 'readonly' : 'readwrite', operation === 'get' ? undefined : { durability: 'strict' }); }
        catch (error) { if (error instanceof TypeError) tx = db.transaction(store, operation === 'get' ? 'readonly' : 'readwrite'); else throw error; }
        const objectStore = tx.objectStore(store); request = operation === 'get' ? objectStore.get(key) : operation === 'delete' ? objectStore.delete(key) : objectStore.put(value, key);
        request.onsuccess = () => { result = request.result; }; request.onerror = () => {};
        tx.oncomplete = () => resolve(result); tx.onerror = tx.onabort = () => reject(tx.error || request.error || failure('storage-failed', 'SharePoint 저장 전 백업을 보관하지 못했습니다.'));
      } catch (error) { reject(error); }
    });
  }
  function filePath(raw, siteUrl, alreadyDecoded) {
    const value = String(raw || '').trim(); let url, path;
    if (!value || value.length > 12000 || /[\\\u0000-\u001f\u007f]/.test(value)) throw failure('invalid-file', '공유 엑셀의 파일 주소를 확인해 주세요.');
    if (alreadyDecoded) path = value;
    else {
      try { url = new URL(value, siteUrl + '/'); } catch (_) { throw failure('invalid-file', '공유 엑셀의 파일 주소를 확인해 주세요.'); }
      const sharingPath = url.protocol === 'https:' && url.pathname.startsWith('/:x:/r/');
      // A company's vanity sharing domain can differ from the web part's
      // canonical host. Extract only its path; never request the alias host or
      // copy its query proof. REST stays at the active same-origin site.
      if (url.origin !== origin() && !sharingPath || url.username || url.password) throw failure('wrong-origin', '현재 SharePoint 사이트 안의 엑셀 파일을 선택해 주세요.');
      try { path = decodeURIComponent(url.pathname); } catch (_) { throw failure('invalid-file', '공유 엑셀의 파일 주소를 확인해 주세요.'); }
    }
    // A browser's path-based Excel sharing link identifies the same file. Its
    // transient query/share proof is neither sent to REST nor persisted.
    if (path.startsWith('/:x:/r/')) path = path.slice('/:x:/r'.length);
    const sitePath = decodeURIComponent(new URL(siteUrl).pathname).replace(/\/+$/, '');
    if (!path.startsWith('/') || path.startsWith('//') || sitePath && !path.startsWith(sitePath + '/') || path.split('/').some(part => part === '.' || part === '..') || /[\\\u0000-\u001f\u007f]/.test(path) || !/\.xlsx$/i.test(path)) throw failure('invalid-file', '현재 SharePoint 사이트의 .xlsx 파일 주소를 지정해 주세요.');
    return path;
  }
  function targetValue(value) {
    const host = context(), siteUrl = siteValue(value?.siteUrl || host.siteUrl);
    if (siteUrl !== host.siteUrl) throw failure('invalid-site', '현재 SharePoint 웹파트와 같은 사이트의 파일을 선택해 주세요.');
    const serverRelativeUrl = value?.serverRelativeUrl ? filePath(value.serverRelativeUrl, siteUrl, true) : filePath(value?.fileUrl || host.fileUrl, siteUrl);
    const uniqueId = value?.uniqueId ? String(value.uniqueId).toLowerCase() : '';
    if (uniqueId && !GUID.test(uniqueId)) throw failure('invalid-target', 'SharePoint 파일 ID가 올바르지 않습니다.');
    return { siteUrl, serverRelativeUrl, uniqueId, name: serverRelativeUrl.slice(serverRelativeUrl.lastIndexOf('/') + 1) };
  }
  function encode(value) { let binary = ''; for (const byte of new TextEncoder().encode(value)) binary += String.fromCharCode(byte); return global.btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); }
  const targetKey = target => encode(JSON.stringify([target.siteUrl, target.serverRelativeUrl, target.uniqueId]));
  const referenceKey = target => JSON.stringify([target.siteUrl, target.serverRelativeUrl]);
  function endpoint(target, content) {
    const suffix = content === true ? '/$value' : typeof content === 'string' ? '/' + content : '';
    const url = new URL(target.siteUrl + "/_api/web/GetFileByServerRelativePath(decodedUrl=@file)" + suffix);
    url.searchParams.set('@file', "'" + target.serverRelativeUrl.replace(/'/g, "''") + "'");
    if (!content) url.searchParams.set('$select', 'Name,Length,UniqueId,ServerRelativeUrl,ETag,Exists,CheckOutType');
    return url.href;
  }
  function versionTag(value) { if (typeof value !== 'string' || !value || value.length > 4096 || /[\r\n]/.test(value) || value === '*') throw failure('invalid-version', 'SharePoint 파일의 변경 버전을 확인하지 못했습니다.'); return value; }
  function bytesValue(value) { const bytes = value instanceof Uint8Array ? value.slice() : value instanceof ArrayBuffer ? new Uint8Array(value.slice(0)) : null; if (!bytes || bytes.length < 4 || bytes.length > LIMIT || bytes[0] !== 80 || bytes[1] !== 75 || bytes[2] !== 3 || bytes[3] !== 4) throw failure('invalid-file', '100 MB 이하의 정상적인 .xlsx 엑셀을 사용해 주세요.'); return bytes; }
  async function digest(bytes) { return Array.from(new Uint8Array(await global.crypto.subtle.digest('SHA-256', bytes)), n => n.toString(16).padStart(2, '0')).join(''); }
  function httpError(code) {
    const choices = { 401: ['auth-required', '회사 SharePoint 로그인이 필요합니다. SharePoint 페이지에서 다시 로그인해 주세요.'],
      403: ['permission-denied', '현재 회사 계정으로 이 파일을 수정할 수 없습니다. 파일의 편집 권한을 확인해 주세요.'],
      404: ['not-found', '지정한 SharePoint 엑셀을 찾지 못했습니다. 파일 위치를 확인해 주세요.'],
      412: ['conflict', '다른 사용자나 Excel이 파일을 변경했습니다. 덮어쓰지 않았습니다. 최신 파일을 다시 읽어 주세요.'],
      409: ['file-locked', 'SharePoint 엑셀이 사용 중이거나 잠겨 있습니다. 최신 파일을 다시 읽어 주세요.'],
      423: ['file-locked', 'SharePoint 엑셀이 사용 중이거나 잠겨 있습니다. 최신 파일을 다시 읽어 주세요.'],
      429: ['throttled', 'SharePoint 요청이 잠시 제한됐습니다. 잠시 후 다시 읽어 주세요.'] };
    const entry = choices[code] || ['sharepoint-failed', 'SharePoint 요청을 완료하지 못했습니다. 네트워크와 파일 상태를 확인해 주세요.']; const error = failure(...entry); error.httpStatus = code; return error;
  }
  function mapped(error) { return error && typeof error.code === 'string' ? error : failure('network-failed', 'SharePoint 요청을 완료하지 못했습니다. 로그인·네트워크 상태를 확인하고 원본을 다시 읽어 주세요.', error); }
  function rememberFailure(error) {
    if (['auth-required', 'permission-denied'].includes(error.code)) { cached = null; status.connected = false; status.needsPermission = true; status.requiresSignIn = error.code === 'auth-required'; }
    if (['not-found', 'invalid-target', 'invalid-version', 'wrong-origin', 'host-required'].includes(error.code)) { cached = null; status.connected = false; status.needsReconnect = true; }
    status.statusMessage = error.message;
  }
  async function request(url, settings) {
    context(); const destination = new URL(url);
    if (destination.origin !== origin()) throw failure('wrong-origin', '다른 사이트로 회사 인증을 보내지 않습니다.');
    const options = settings || {}, controller = new AbortController(), timer = setTimeout(() => controller.abort(), 60000);
    try {
      if (options.onAttempt) options.onAttempt();
      const response = await global.fetch(destination.href, { method: options.method || 'GET', credentials: 'same-origin', mode: 'same-origin', cache: 'no-store', redirect: 'error', signal: controller.signal,
        headers: { Accept: 'application/json;odata=nometadata', ...(options.headers || {}) }, body: options.body });
      if (!response.ok) throw httpError(response.status); return response;
    } finally { clearTimeout(timer); }
  }
  async function metadata(target) {
    const response = await request(endpoint(target, false)), raw = await response.json(), value = raw.d || raw;
    const uniqueId = String(value.UniqueId || '').toLowerCase(), size = Number(value.Length), path = value.ServerRelativeUrl;
    if (!GUID.test(uniqueId) || target.uniqueId && uniqueId !== target.uniqueId || path !== target.serverRelativeUrl || value.Exists === false || !Number.isSafeInteger(size) || size < 4 || size > LIMIT || value.Name !== target.name || !/\.xlsx$/i.test(value.Name)) throw failure('invalid-target', 'SharePoint 파일의 위치·ID·크기가 예상과 다릅니다. 다시 연결해 주세요.');
    if (![0, 1, 2].includes(value.CheckOutType)) throw failure('invalid-target', 'SharePoint 파일의 체크아웃 상태를 확인하지 못했습니다.');
    return { uniqueId, size, name: value.Name, eTag: versionTag(value.ETag || response.headers.get('etag')), checkOutType: value.CheckOutType };
  }
  async function snapshot(target, options) {
    const before = await metadata(target), canonical = { ...target, uniqueId: before.uniqueId };
    if (options?.allowCached && !status.uncertain && cached && referenceKey(canonical) === referenceKey(cached.target) && cached.target.uniqueId === before.uniqueId && cached.eTag === before.eTag && cached.bytes.length === before.size) {
      const bytes = cached.bytes.slice(); return { ...cached, target: canonical, bytes, checkOutType: before.checkOutType, file: new File([bytes], before.name, { type: MIME }) };
    }
    const response = await request(endpoint(canonical, true));
    if (Number(response.headers.get('content-length')) > LIMIT) throw failure('invalid-file', '공유 엑셀은 100 MB 이하로 사용해 주세요.');
    const bytes = bytesValue(await response.arrayBuffer()), after = await metadata(canonical);
    if (before.eTag !== after.eTag || before.size !== after.size || before.uniqueId !== after.uniqueId || bytes.length !== after.size) throw failure('conflict', '파일을 읽는 동안 다른 사용자가 변경했습니다. 최신 파일을 다시 읽어 주세요.');
    const hash = await digest(bytes), version = 'sharepoint-v1:' + targetKey(canonical) + ':' + encode(after.eTag) + ':' + hash;
    cached = { target: canonical, bytes: bytes.slice(), eTag: after.eTag, hash, version, checkOutType: after.checkOutType };
    return { ...cached, bytes, file: new File([bytes], after.name, { type: MIME }) };
  }
  function update(current) { selected = { ...current.target }; Object.assign(status, { connected: true, name: current.file.name, version: current.version, persistent: true, needsPermission: false, needsReconnect: false, requiresSignIn: false, uncertain: !!pendingLock, statusMessage: pendingLock ? 'SharePoint에서 본인이 방금 진행한 저장의 파일 상태를 확인한 뒤 원본을 다시 읽어 주세요. 다른 사람의 체크아웃은 해제하지 마세요.' : '' }); }
  async function reconcileLock(current) {
    pendingLock = await stored('locks', 'get', targetKey(current.target)) || null;
    if (pendingLock && current.checkOutType === 2) { await stored('locks', 'delete', targetKey(current.target)); pendingLock = null; }
  }
  async function journalLock(target, stage) {
    const value = { target: { ...target }, stage, createdAt: pendingLock?.createdAt || new Date().toISOString() };
    await stored('locks', 'put', targetKey(target), value); pendingLock = value;
  }
  async function clearLock(target) { await stored('locks', 'delete', targetKey(target)); pendingLock = null; }
  async function mutate(target, action, proof, onAttempt) { return request(endpoint(target, action), { method: 'POST', headers: { 'X-RequestDigest': proof }, onAttempt }); }
  async function formDigest(target) {
    const response = await request(target.siteUrl + '/_api/contextinfo', { method: 'POST', headers: { 'Content-Type': 'application/json;odata=nometadata' } });
    const raw = await response.json(), value = raw.d?.GetContextWebInformation || raw.GetContextWebInformation || raw.d || raw;
    if (typeof value.FormDigestValue !== 'string' || !value.FormDigestValue || value.FormDigestValue.length > 16384 || /[\r\n]/.test(value.FormDigestValue) || value.WebFullUrl && siteValue(value.WebFullUrl) !== target.siteUrl) throw failure('invalid-digest', 'SharePoint 저장 확인 정보를 받지 못했습니다. 페이지에 다시 로그인해 주세요.');
    return value.FormDigestValue;
  }
  async function connect(options) {
    return serial(async () => { try {
      if (!supported()) { context(); throw failure('unsupported', 'SharePoint 안에서 브라우저 저장소를 사용할 수 있는 창으로 열어 주세요.'); }
      const settings = options || {}; if (settings.target && (settings.fileUrl || settings.siteUrl)) throw failure('invalid-target', '파일 ID 또는 파일 주소 중 하나로 연결해 주세요.');
      const target = targetValue(settings.target || settings), current = await snapshot(target);
      if (typeof settings.validate === 'function') await settings.validate(current.file);
      await stored('targets', 'put', 'selected', current.target); await reconcileLock(current); update(current); backup = await stored('backups', 'get', targetKey(selected)) || null;
      return complete({ file: current.file, bytes: current.bytes, version: current.version });
    } catch (error) { const result = mapped(error); rememberFailure(result); throw result; } });
  }
  async function read() { return serial(async () => { try { if (!selected) throw failure('not-connected', 'SharePoint 엑셀을 먼저 연결해 주세요.'); const current = await snapshot(targetValue(selected), { allowCached: !status.uncertain }); await reconcileLock(current); update(current); return current; } catch (error) { const result = mapped(error); rememberFailure(result); throw result; } }); }
  async function restore() {
    return serial(async () => { try {
      context(); const saved = await stored('targets', 'get', 'selected'); if (!saved) return complete();
      const target = targetValue(saved), host = context();
      if (host.fileUrl && filePath(host.fileUrl, host.siteUrl) !== target.serverRelativeUrl) throw failure('invalid-target', '웹파트에 지정한 엑셀이 변경됐습니다. 지정된 파일에 다시 연결해 주세요.');
      selected = target; backup = await stored('backups', 'get', targetKey(target)) || null; const current = await snapshot(target); await reconcileLock(current); update(current); return complete({ file: current.file, bytes: current.bytes, version: current.version });
    } catch (error) { const result = mapped(error); rememberFailure(result); return complete(); } });
  }
  function checkRevision(verify) { if (typeof verify === 'function' && verify() !== true) throw failure('stale', '사진이나 연결 파일이 변경되어 이전 결과를 저장하지 않았습니다.'); }
  async function performSave(options) {
    const settings = options || {}, bytes = bytesValue(settings.bytes);
    if (!selected) throw failure('not-connected', 'SharePoint 엑셀을 먼저 연결해 주세요.');
    const target = targetValue(selected);
    if (status.uncertain || pendingLock) throw failure('requires-read', 'SharePoint에서 본인이 방금 진행한 저장의 체크인 상태를 확인한 뒤 원본을 다시 읽어 주세요. 다른 사람의 체크아웃은 해제하지 마세요.');
    const parts = typeof settings.expectedVersion === 'string' ? settings.expectedVersion.split(':') : [];
    if (parts.length !== 4 || parts[0] !== 'sharepoint-v1' || parts[1] !== targetKey(target) || !/^[a-f0-9]{64}$/.test(parts[3])) throw failure('conflict', '선택한 SharePoint 원본의 최신 버전을 다시 읽어 주세요.');
    checkRevision(settings.verifyCurrent); const latest = await snapshot(target, { allowCached: true });
    if (latest.checkOutType !== 2) throw failure('file-locked', '다른 작업에서 체크아웃한 엑셀입니다. 기존 작업을 체크인한 뒤 다시 읽어 주세요. 기존 작업은 변경하지 않았습니다.');
    if (latest.version !== settings.expectedVersion) throw failure('conflict', '다른 사용자나 Excel이 파일을 변경했습니다. 최신 파일을 다시 읽어 주세요.');
    const hash = await digest(bytes); checkRevision(settings.verifyCurrent);
    const original = { name: latest.file.name.replace(/\.xlsx$/i, '') + '_저장전백업.xlsx', bytes: latest.bytes.slice(), version: latest.version, timestamp: new Date().toISOString() };
    try { await stored('backups', 'put', targetKey(target), original); backup = original; } catch (error) { throw failure('backup-failed', '저장 전 원본을 기기에 백업하지 못해 SharePoint 저장을 중단했습니다. 원본은 변경하지 않았습니다.', error); }
    const proof = await formDigest(target); checkRevision(settings.verifyCurrent);
    let checkoutAttempted = false, checkoutOwned = false, writeStarted = false, writeRejected = false, checkinStarted = false;
    // SharePoint's file stream does not document If-Match as a complete
    // concurrency guarantee. Use its server checkout lock and verify the
    // original SHA again after acquiring it. Never reuse any existing checkout.
    await journalLock(target, 'acquiring');
    try {
      await mutate(target, 'CheckOut()', proof, () => { checkRevision(settings.verifyCurrent); checkoutAttempted = true; }); checkoutOwned = true;
      await journalLock(target, 'checked-out');
      const locked = await snapshot(target);
      if (locked.checkOutType === 2) throw failure('lock-failed', 'SharePoint의 파일 잠금을 확인하지 못했습니다. 저장하지 않았습니다.');
      if (locked.hash !== latest.hash) throw failure('conflict', '파일 잠금을 얻기 전에 다른 사용자가 엑셀을 변경했습니다. 덮어쓰지 않았습니다. 최신 파일을 다시 읽어 주세요.');
      checkRevision(settings.verifyCurrent); await journalLock(target, 'writing');
      let response;
      try {
        response = await request(endpoint(target, true), { method: 'POST', headers: { 'X-HTTP-Method': 'PUT', 'X-RequestDigest': proof, 'If-Match': locked.eTag, 'Content-Type': MIME }, body: bytes, onAttempt: () => { checkRevision(settings.verifyCurrent); writeStarted = true; } });
      } catch (error) { writeRejected = Number.isInteger(error?.httpStatus) && error.httpStatus >= 400 && error.httpStatus < 500; throw error; }
      const writtenTag = response.headers.get('etag'), written = await snapshot(target);
      if (written.hash !== hash || written.checkOutType === 2 || writtenTag && written.eTag !== versionTag(writtenTag) || hash !== locked.hash && written.eTag === locked.eTag) throw failure('verification-failed', '체크아웃 상태의 저장 내용을 확인하지 못했습니다. SharePoint에서 파일 상태를 확인해 주세요.');
      await journalLock(target, 'checking-in');
      await mutate(target, "CheckIn(comment='Tool Manager calibration update',checkintype=1)", proof, () => { checkinStarted = true; });
      const actual = await snapshot(target);
      if (actual.hash !== hash || actual.checkOutType !== 2) throw failure('verification-failed', 'SharePoint 체크인 후 파일 내용을 확인하지 못했습니다. 원본을 다시 확인해 주세요.');
      await clearLock(target); update(actual); status.lastSavedAt = new Date().toISOString(); return { name: actual.file.name, version: actual.version, backupAvailable: true };
    } catch (error) {
      const result = mapped(error), acquisitionRejected = !checkoutOwned && Number.isInteger(result.httpStatus) && result.httpStatus >= 400 && result.httpStatus < 500;
      let released = false;
      if (checkoutOwned && (!writeStarted || writeRejected) && !checkinStarted) {
        // Only the lock acquired by this operation may be undone. In
        // particular, neither an existing checkout nor a rejected acquisition
        // is ever undone. Verify the exact file identity before releasing.
        try {
          const current = await metadata(target);
          if (current.checkOutType !== 2) await mutate(target, 'UndoCheckOut()', proof);
          if ((await metadata(target)).checkOutType !== 2) throw failure('lock-release-failed', '체크아웃 해제 결과를 확인하지 못했습니다.');
          await clearLock(target); released = true;
        } catch (_) { /* Keep the durable journal for explicit user recovery. */ }
      } else if (!checkoutAttempted || acquisitionRejected) { await clearLock(target); released = true; }
      cached = null;
      if (!released) {
        status.uncertain = true;
        result.message = '저장 또는 체크아웃 완료 여부를 확인할 수 없습니다. 자동 재시도·변경 취소하지 않습니다. SharePoint에서 본인이 방금 진행한 저장의 파일 상태를 확인한 뒤 원본을 다시 읽어 주세요. 다른 사람의 체크아웃은 해제하지 마세요. 저장 전 백업은 기기에 보관되어 있습니다.';
      } else status.uncertain = false;
      throw result;
    }
  }
  async function save(options) { return serial(async () => { try { return global.navigator?.locks?.request ? await global.navigator.locks.request(LOCK, { mode: 'exclusive' }, () => performSave(options)) : await performSave(options); } catch (error) { const result = mapped(error); rememberFailure(result); throw result; } }); }
  async function takeBackup() { if (!backup && selected) try { backup = await stored('backups', 'get', targetKey(selected)) || null; } catch (_) { return null; } return backup ? { ...backup, bytes: new Uint8Array(backup.bytes).slice() } : null; }
  async function disconnect(options) { return serial(async () => { await stored('targets', 'delete', 'selected'); if (options?.clearBackup && selected) await stored('backups', 'delete', targetKey(selected)); selected = null; cached = null; pendingLock = null; if (options?.clearBackup) backup = null; status = { ...initial(), busy: true }; return complete(); }); }
  global.ToolSharePointFile = Object.freeze({ connect, read, save, restore, requestAccess: read, disconnect, takeBackup, inspect });
})(window);
