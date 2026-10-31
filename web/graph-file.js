/* Microsoft Graph workbook connection. Tokens and download URLs stay in memory. */
(function (global) {
  'use strict';
  const GRAPH = 'https://graph.microsoft.com/v1.0', DATABASE = 'tool-management-graph-file-v1';
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const LIMIT = 100 * 1024 * 1024, LOCK = 'tool-management-graph-workbook';
  let selected = null, databasePromise = null, queue = Promise.resolve(), backup = null, cachedSnapshot = null;
  const initial = () => ({ mode: 'microsoft365', connected: false, name: '', needsPermission: false, needsReconnect: false,
    requiresSignIn: false, persistent: false, busy: false, uncertain: false, version: '', lastSavedAt: '', statusMessage: '' });
  let status = initial();
  function failure(code, message, cause) { const error = new Error(message); error.code = code; if (cause) error.cause = cause; return error; }
  function supported() { return global.isSecureContext !== false && !!(global.fetch && global.crypto && global.crypto.subtle && global.indexedDB); }
  function inspect() { return Object.assign({}, status, { supported: supported(), backupAvailable: !!backup, target: selected ? { driveId: selected.driveId, itemId: selected.itemId, name: selected.name } : null }); }
  const complete = (extra) => Object.assign(inspect(), { busy: false }, extra || {});
  function serial(action) {
    const operation = queue.then(async () => { status.busy = true; try { return await action(); } finally { status.busy = false; } });
    queue = operation.catch(() => {}); return operation;
  }
  function openDatabase() {
    if (databasePromise) return databasePromise;
    databasePromise = new Promise((resolve, reject) => {
      if (!global.indexedDB) { reject(failure('storage-unavailable', '저장 전 백업을 보관할 브라우저 저장소가 없습니다.')); return; }
      let request;
      try { request = global.indexedDB.open(DATABASE, 1); } catch (error) { reject(error); return; }
      request.onupgradeneeded = () => { for (const name of ['targets', 'backups']) if (!request.result.objectStoreNames.contains(name)) request.result.createObjectStore(name); };
      request.onsuccess = () => { const db = request.result; db.onversionchange = () => { db.close(); databasePromise = null; }; resolve(db); };
      request.onerror = () => reject(failure('storage-unavailable', '공유 엑셀 연결·백업 저장소를 열지 못했습니다.', request.error));
      request.onblocked = () => reject(failure('storage-unavailable', '다른 앱 창을 닫은 후 공유 엑셀 연결을 다시 시도해 주세요.'));
    });
    databasePromise.catch(() => { databasePromise = null; }); return databasePromise;
  }
  async function stored(store, operation, key, value) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
      let transaction, request, result;
      try {
        try { transaction = db.transaction(store, operation === 'get' ? 'readonly' : 'readwrite', operation === 'get' ? undefined : { durability: 'strict' }); }
        catch (error) { if (error instanceof TypeError) transaction = db.transaction(store, operation === 'get' ? 'readonly' : 'readwrite'); else throw error; }
        const objectStore = transaction.objectStore(store);
        request = operation === 'get' ? objectStore.get(key) : operation === 'delete' ? objectStore.delete(key) : objectStore.put(value, key);
        request.onsuccess = () => { result = request.result; };
        request.onerror = () => {};
        transaction.oncomplete = () => resolve(result);
        transaction.onabort = transaction.onerror = () => reject(transaction.error || request.error || failure('storage-failed', '공유 엑셀 백업을 저장하지 못했습니다.'));
      } catch (error) { reject(error); }
    });
  }
  async function digest(bytes) {
    if (!supported()) throw failure('unsupported', '이 앱은 HTTPS와 브라우저 저장소를 사용할 수 있는 창에서 열어 주세요.');
    return Array.from(new Uint8Array(await global.crypto.subtle.digest('SHA-256', bytes)), (n) => n.toString(16).padStart(2, '0')).join('');
  }
  function base64url(value) {
    const bytes = new TextEncoder().encode(value); let binary = '';
    for (const byte of bytes) binary += String.fromCharCode(byte);
    return global.btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function validId(value) { return typeof value === 'string' && value.length > 0 && value.length <= 1024 && !/[\u0000-\u001f\u007f]/.test(value); }
  function targetValue(value) {
    if (!value || !validId(value.driveId) || !validId(value.itemId) || typeof value.name !== 'string' || !/\.xlsx$/i.test(value.name)) throw failure('invalid-target', '선택한 문서가 정상적인 .xlsx 엑셀이 아닙니다.');
    return { driveId: value.driveId, itemId: value.itemId, name: value.name };
  }
  const targetKey = (target) => base64url(JSON.stringify([target.driveId, target.itemId]));
  const itemPath = (target) => '/drives/' + encodeURIComponent(target.driveId) + '/items/' + encodeURIComponent(target.itemId);
  const select = '?$select=id,name,eTag,size,file,folder,parentReference,@microsoft.graph.downloadUrl';
  function etag(value) {
    if (typeof value !== 'string' || !value || value.length > 4096 || /[\r\n]/.test(value)) throw failure('invalid-version', '공유 문서의 변경 버전을 확인하지 못했습니다. 다시 연결해 주세요.');
    return value;
  }
  function bytesValue(value) {
    const bytes = value instanceof Uint8Array ? value.slice() : value instanceof ArrayBuffer ? new Uint8Array(value.slice(0)) : null;
    if (!bytes || bytes.length < 4 || bytes.length > LIMIT || bytes[0] !== 80 || bytes[1] !== 75 || bytes[2] !== 3 || bytes[3] !== 4) throw failure('invalid-file', '100 MB 이하의 정상적인 .xlsx 엑셀을 사용해 주세요.');
    return bytes;
  }
  function graphError(code) {
    if (code === 401) return failure('auth-required', '회사 Microsoft 365 로그인이 만료됐습니다. 다시 로그인한 후 원본을 읽어 주세요.');
    if (code === 403) return failure('permission-denied', '회사 계정의 문서 접근·수정 권한 또는 앱 사용 동의가 없습니다. 회사 IT의 권한 설정을 확인해 주세요.');
    if (code === 404) return failure('not-found', '공유 엑셀을 찾지 못했습니다. 문서의 공유 링크를 다시 연결해 주세요.');
    if (code === 412) return failure('conflict', '다른 사용자나 Excel이 공유 문서를 변경했습니다. 덮어쓰지 않았습니다. 최신 파일을 다시 읽어 주세요.');
    if (code === 409 || code === 423) return failure('file-locked', '회사 엑셀이 잠겨 있거나 다른 저장 작업 중입니다. 잠시 후 최신 파일을 다시 읽어 주세요.');
    if (code === 429) return failure('throttled', 'Microsoft 365 요청이 잠시 제한됐습니다. 잠시 후 원본을 다시 읽어 주세요.');
    return failure('cloud-failed', 'Microsoft 365 요청을 완료하지 못했습니다. 네트워크와 회사 서비스 상태를 확인해 주세요.');
  }
  function rememberFailure(error) {
    if (['auth-required', 'expired', 'auth-expired'].includes(error.code)) { cachedSnapshot = null; status.connected = false; status.requiresSignIn = true; status.needsPermission = true; }
    if (error.code === 'permission-denied') { cachedSnapshot = null; status.connected = false; status.needsPermission = true; }
    if (['not-found', 'invalid-target', 'invalid-version'].includes(error.code)) { cachedSnapshot = null; status.connected = false; status.needsReconnect = true; }
    status.statusMessage = error.message;
  }
  function mapError(error) {
    if (error && typeof error.code === 'string') return error;
    return failure('network-failed', '공유 엑셀 요청을 완료하지 못했습니다. 네트워크를 확인하고 원본을 다시 읽어 주세요.', error);
  }
  async function accessToken() {
    const auth = global.ToolM365Auth;
    if (!auth || typeof auth.getAccessToken !== 'function') throw failure('auth-required', '회사 Microsoft 365 로그인 연결을 먼저 설정해 주세요.');
    let token;
    try { token = await auth.getAccessToken({ minimumValidityMs: 60000 }); } catch (error) { throw failure('auth-required', '회사 Microsoft 365 로그인이 필요합니다. 다시 로그인해 주세요.', error); }
    if (typeof token !== 'string' || !token || /[\r\n]/.test(token)) throw failure('auth-required', '회사 Microsoft 365 로그인이 필요합니다.');
    return token;
  }
  async function graph(path, options) {
    if (typeof path !== 'string' || !path.startsWith('/') || path.startsWith('//') || /[\r\n]/.test(path)) throw failure('invalid-request', 'Microsoft 365 요청 주소가 올바르지 않습니다.');
    const settings = options || {}, token = await accessToken(), controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 60000);
    try {
      if (typeof settings.onAttempt === 'function') settings.onAttempt();
      const response = await global.fetch(GRAPH + path, {
        method: settings.method || 'GET', cache: 'no-store', credentials: 'omit', redirect: 'error', signal: controller.signal,
        headers: Object.assign({ Authorization: 'Bearer ' + token, Accept: 'application/json' }, settings.headers || {}), body: settings.body
      });
      if (!response.ok) { const error = graphError(response.status); error.httpStatus = response.status; throw error; }
      return response;
    } finally { clearTimeout(timer); }
  }
  async function metadata(target) {
    const value = await (await graph(itemPath(target) + select)).json();
    if (value.id !== target.itemId || !value.parentReference || value.parentReference.driveId !== target.driveId || value.folder || !value.file || !/\.xlsx$/i.test(String(value.name || '')) || !Number.isFinite(value.size) || value.size < 4 || value.size > LIMIT) throw failure('invalid-target', '공유 문서가 100 MB 이하의 .xlsx 엑셀인지 확인해 주세요.');
    etag(value.eTag); return value;
  }
  async function download(value) {
    let url;
    try { url = new URL(value['@microsoft.graph.downloadUrl']); } catch (_) { throw failure('download-unavailable', '문서 다운로드 주소를 확인하지 못했습니다. 공유 링크를 다시 연결해 주세요.'); }
    if (url.protocol !== 'https:' || url.username || url.password) throw failure('download-unavailable', 'Microsoft 365 문서 다운로드 주소가 올바르지 않습니다.');
    const controller = new AbortController(), timer = setTimeout(() => controller.abort(), 60000);
    try {
      // A Graph-provided download URL already has its own short-lived proof.
      // Never send the Graph bearer token to SharePoint or another host.
      const response = await global.fetch(url.href, { method: 'GET', credentials: 'omit', cache: 'no-store', redirect: 'follow', signal: controller.signal });
      if (!response.ok) throw failure('download-failed', '공유 문서 내용을 내려받지 못했습니다. 회사 로그인·네트워크를 확인하고 다시 읽어 주세요.');
      const size = response.headers.get('content-length');
      if (size && Number(size) > LIMIT) throw failure('invalid-file', '공유 엑셀은 100 MB 이하로 사용해 주세요.');
      return bytesValue(await response.arrayBuffer());
    } finally { clearTimeout(timer); }
  }
  function cachedValue(target, value) {
    // Keep one verified workbook in memory, never a bearer or a preauthenticated
    // URL. Returned bytes are separate so callers cannot mutate this proof.
    cachedSnapshot = { targetKey: targetKey(target), name: value.file.name, bytes: value.bytes.slice(), version: value.version, eTag: value.eTag, hash: value.hash };
  }
  async function snapshot(target, options) {
    const before = await metadata(target), cached = cachedSnapshot;
    if (options && options.allowCached && !status.uncertain && cached && cached.targetKey === targetKey(target)
      && cached.eTag === before.eTag && cached.bytes.length === before.size) {
      const bytes = cached.bytes.slice();
      return { file: new File([bytes], before.name, { type: MIME }), bytes, version: cached.version, eTag: cached.eTag, hash: cached.hash };
    }
    const bytes = await download(before), after = await metadata(target);
    if (before.eTag !== after.eTag || before.size !== after.size) throw failure('conflict', '공유 엑셀을 읽는 동안 다른 사용자가 수정했습니다. 최신 파일을 다시 읽어 주세요.');
    if (bytes.length !== after.size) throw failure('verification-failed', '공유 문서의 실제 크기가 예상과 다릅니다. 원본을 다시 읽어 주세요.');
    const hash = await digest(bytes), file = new File([bytes], after.name, { type: MIME });
    const result = { file, bytes, version: 'graph-v1:' + targetKey(target) + ':' + base64url(etag(after.eTag)) + ':' + hash, eTag: after.eTag, hash };
    cachedValue(target, result); return result;
  }
  function versionParts(version, target) {
    const parts = typeof version === 'string' ? version.split(':') : [];
    if (parts.length !== 4 || parts[0] !== 'graph-v1' || parts[1] !== targetKey(target) || !/^[a-f0-9]{64}$/.test(parts[3])) throw failure('conflict', '저장 전 선택한 공유 엑셀을 다시 읽어 주세요.');
    return parts;
  }
  function checkRevision(verify) { if (typeof verify === 'function' && verify() !== true) throw failure('stale', '사진이나 연결 파일이 변경되어 이전 결과를 저장하지 않았습니다. 다시 처리해 주세요.'); }
  async function selectedTarget() { if (!selected) throw failure('not-connected', '회사 공유 엑셀을 먼저 연결해 주세요.'); await accessToken(); return selected; }
  async function connect(options) {
    return serial(async () => {
      try {
        if (!supported()) throw failure('unsupported', 'HTTPS 앱에서 브라우저 저장소를 허용한 후 다시 연결해 주세요.');
        const settings = typeof options === 'string' ? { shareUrl: options } : options || {};
        if (typeof settings !== 'object' || Array.isArray(settings)) throw failure('invalid-target', '회사 Excel 연결 정보를 확인해 주세요.');
        let reference;
        if ('target' in settings) {
          if ('shareUrl' in settings) throw failure('invalid-target', '문서 ID와 공유 링크 중 하나로만 연결해 주세요.');
          const target = settings.target;
          if (!target || typeof target !== 'object' || !validId(target.driveId) || !validId(target.itemId)
            || target.name !== undefined && (typeof target.name !== 'string' || !/\.xlsx$/i.test(target.name))) throw failure('invalid-target', '지정된 회사 Excel 문서 ID를 확인해 주세요.');
          reference = { driveId: target.driveId, itemId: target.itemId };
          await accessToken(); cachedSnapshot = null;
        } else {
          let url;
          try { url = new URL(String(settings.shareUrl || '').trim()); } catch (_) { throw failure('invalid-link', '회사 Excel의 HTTPS 공유 링크를 입력해 주세요.'); }
          if (url.protocol !== 'https:' || url.username || url.password || url.href.length > 12000) throw failure('invalid-link', '회사 Excel의 HTTPS 공유 링크를 입력해 주세요.');
          await accessToken(); cachedSnapshot = null;
          const resolved = await (await graph('/shares/u!' + base64url(url.href) + '/driveItem' + select, { headers: { Prefer: 'redeemSharingLinkIfNecessary' } })).json();
          const item = resolved.remoteItem || resolved;
          reference = targetValue({ driveId: item.parentReference && item.parentReference.driveId, itemId: item.id, name: item.name });
        }
        const current = await snapshot(reference);
        const candidate = targetValue({ driveId: reference.driveId, itemId: reference.itemId, name: current.file.name });
        if (typeof settings.validate === 'function') await settings.validate(current.file);
        await stored('targets', 'put', 'selected', candidate);
        selected = candidate; backup = await stored('backups', 'get', targetKey(candidate)) || null;
        Object.assign(status, { connected: true, name: current.file.name, persistent: true, version: current.version, needsPermission: false, requiresSignIn: false, needsReconnect: false, uncertain: false, statusMessage: '' });
        return complete({ file: current.file, bytes: current.bytes, version: current.version });
      } catch (error) { const mapped = mapError(error); rememberFailure(mapped); throw mapped; }
    });
  }
  async function read() {
    return serial(async () => {
      try {
        const target = await selectedTarget(), current = await snapshot(target, { allowCached: !status.uncertain });
        selected = Object.assign({}, target, { name: current.file.name });
        Object.assign(status, { connected: true, name: current.file.name, version: current.version, uncertain: false, requiresSignIn: false, needsPermission: false, needsReconnect: false, statusMessage: '' });
        return current;
      } catch (error) { const mapped = mapError(error); rememberFailure(mapped); throw mapped; }
    });
  }
  async function restore() {
    return serial(async () => {
      try {
        const saved = await stored('targets', 'get', 'selected');
        if (!saved) return complete();
        selected = targetValue(saved); cachedSnapshot = null; backup = await stored('backups', 'get', targetKey(selected)) || null;
        Object.assign(status, { name: selected.name, persistent: true });
        try { await accessToken(); } catch (error) { rememberFailure(error); return complete(); }
        const current = await snapshot(selected);
        Object.assign(status, { connected: true, name: current.file.name, version: current.version, requiresSignIn: false, needsPermission: false, needsReconnect: false, statusMessage: '' });
        return complete({ file: current.file, bytes: current.bytes, version: current.version });
      } catch (error) { const mapped = mapError(error); rememberFailure(mapped); return complete(); }
    });
  }
  async function performSave(options) {
    const settings = options || {}, bytes = bytesValue(settings.bytes), target = await selectedTarget();
    if (status.uncertain) throw failure('requires-read', '이전 저장 완료 여부를 먼저 확인해야 합니다. 공유 원본을 다시 읽어 주세요.');
    versionParts(settings.expectedVersion, target); checkRevision(settings.verifyCurrent);
    const latest = await snapshot(target, { allowCached: true });
    if (latest.version !== settings.expectedVersion) throw failure('conflict', '다른 사용자나 Excel이 원본을 변경했습니다. 최신 파일을 다시 읽고 사진을 확인해 주세요.');
    const desiredHash = await digest(bytes); checkRevision(settings.verifyCurrent);
    const original = { name: latest.file.name.replace(/\.xlsx$/i, '') + '_저장전백업.xlsx', bytes: latest.bytes.slice(), version: latest.version, timestamp: new Date().toISOString(), target: targetValue(target) };
    try { await stored('backups', 'put', targetKey(target), original); backup = original; }
    catch (error) { throw failure('backup-failed', '저장 전 원본 백업을 기기에 보관하지 못해 공유 엑셀 저장을 중단했습니다. 원본은 변경하지 않았습니다.', error); }
    checkRevision(settings.verifyCurrent);
    const beforeWrite = await metadata(target);
    if (beforeWrite.eTag !== latest.eTag) throw failure('conflict', '저장 준비 중 다른 사용자가 공유 엑셀을 변경했습니다. 덮어쓰지 않았습니다. 최신 파일을 다시 읽어 주세요.');
    checkRevision(settings.verifyCurrent);
    let writeStarted = false;
    try {
      const response = await graph(itemPath(target) + '/content', { method: 'PUT', headers: { 'Content-Type': MIME, 'If-Match': latest.eTag }, body: bytes, onAttempt: () => { checkRevision(settings.verifyCurrent); writeStarted = true; } });
      const written = await response.json();
      if (!written || written.id !== target.itemId) throw failure('verification-failed', '공유 엑셀 저장 결과를 확인하지 못했습니다. 원본을 다시 읽어 주세요.');
      etag(written.eTag);
      const actual = await snapshot(target);
      if (actual.hash !== desiredHash || actual.eTag !== written.eTag) throw failure('verification-failed', '저장 후 공유 파일 내용이 예상과 다릅니다. 자동 재시도하지 않습니다. 원본을 다시 읽어 주세요.');
      selected = Object.assign({}, target, { name: actual.file.name });
      Object.assign(status, { connected: true, name: actual.file.name, version: actual.version, uncertain: false, requiresSignIn: false, needsPermission: false, needsReconnect: false, statusMessage: '', lastSavedAt: new Date().toISOString() });
      return { name: actual.file.name, version: actual.version, backupAvailable: true };
    } catch (error) {
      const mapped = mapError(error), rejected = [401, 403, 404, 409, 412, 423, 429].includes(mapped.httpStatus);
      if (writeStarted && !rejected) { cachedSnapshot = null; status.uncertain = true; mapped.message = '저장 완료 여부를 확인할 수 없습니다. 자동 재시도하지 않습니다. 원본을 다시 읽어 확인해 주세요. 저장 전 백업은 기기에 보관되어 있습니다.'; }
      throw mapped;
    }
  }
  async function save(options) {
    return serial(async () => {
      try {
        if (global.navigator && global.navigator.locks && typeof global.navigator.locks.request === 'function') return await global.navigator.locks.request(LOCK, { mode: 'exclusive' }, () => performSave(options));
        return await performSave(options);
      } catch (error) { const mapped = mapError(error); rememberFailure(mapped); throw mapped; }
    });
  }
  async function takeBackup() {
    if (!backup && selected) { try { backup = await stored('backups', 'get', targetKey(selected)) || null; } catch (_) { return null; } }
    return backup ? { name: backup.name, bytes: new Uint8Array(backup.bytes).slice(), version: backup.version, timestamp: backup.timestamp } : null;
  }
  async function disconnect(options) {
    return serial(async () => {
      await stored('targets', 'delete', 'selected');
      if (options && options.clearBackup && selected) await stored('backups', 'delete', targetKey(selected));
      selected = null; cachedSnapshot = null; if (options && options.clearBackup) backup = null;
      status = Object.assign(initial(), { busy: true }); return complete();
    });
  }
  async function requestAccess() { await accessToken(); return read(); }
  global.ToolGraphFile = Object.freeze({ connect, read, save, restore, requestAccess, disconnect, takeBackup, inspect });
})(window);
