/* The user selects a local OneDrive-synced workbook. No URL or remote API is used. */
(function (global) {
  'use strict';
  const DATABASE = 'tool-management-connected-file-v1';
  const DATABASE_VERSION = 2;
  const LOCK = 'tool-management-connected-workbook';
  const MODE = 'onedrive-local';
  let selectedHandle = null, databasePromise = null, queue = Promise.resolve();
  let backup = null;
  let status = {
    connected: false, name: '', mode: MODE, needsPermission: false, needsReconnect: false,
    persistent: false, statusMessage: '', busy: false, uncertain: false, version: '', lastSavedAt: ''
  };
  function failure(code, message, cause) {
    const error = new Error(message);
    error.code = code;
    if (cause) error.cause = cause;
    return error;
  }
  function supported() {
    return global.isSecureContext !== false && typeof global.showOpenFilePicker === 'function' && !!(global.crypto && global.crypto.subtle);
  }
  function inspect() { return Object.assign({}, status, { supported: supported(), backupAvailable: !!backup }); }
  function completed(extra) { return Object.assign(inspect(), { busy: false }, extra || {}); }
  function serial(action) {
    const operation = queue.then(async () => {
      status.busy = true;
      try { return await action(); } finally { status.busy = false; }
    });
    queue = operation.catch(() => {});
    return operation;
  }
  function openDatabase() {
    if (databasePromise) return databasePromise;
    databasePromise = new Promise((resolve, reject) => {
      if (!global.indexedDB) { reject(failure('storage-unavailable', '브라우저 저장소를 사용할 수 없습니다. Chrome 또는 Edge의 일반 창에서 다시 열어 주세요.')); return; }
      let request;
      try { request = global.indexedDB.open(DATABASE, DATABASE_VERSION); }
      catch (error) { reject(failure('storage-unavailable', '브라우저 저장소를 열 수 없습니다.', error)); return; }
      request.onupgradeneeded = () => {
        for (const name of ['handles', 'backups']) if (!request.result.objectStoreNames.contains(name)) request.result.createObjectStore(name);
      };
      request.onsuccess = () => {
        const database = request.result;
        database.onversionchange = () => { database.close(); databasePromise = null; };
        resolve(database);
      };
      request.onerror = () => reject(failure('storage-unavailable', '브라우저 저장소를 열 수 없습니다.', request.error));
      request.onblocked = () => reject(failure('storage-unavailable', '다른 프로그램 탭이 저장소 변경을 막고 있습니다. 다른 탭을 닫고 다시 시도해 주세요.'));
    });
    databasePromise.catch(() => { databasePromise = null; });
    return databasePromise;
  }
  async function stored(store, operation, key, value) {
    const database = await openDatabase();
    return new Promise((resolve, reject) => {
      let transaction, request, result;
      const writing = operation !== 'get';
      try {
        try { transaction = database.transaction(store, writing ? 'readwrite' : 'readonly', writing ? { durability: 'strict' } : undefined); }
        catch (error) {
          if (error instanceof TypeError) transaction = database.transaction(store, writing ? 'readwrite' : 'readonly');
          else throw error;
        }
        const objectStore = transaction.objectStore(store);
        request = operation === 'put' ? objectStore.put(value, key) : operation === 'delete' ? objectStore.delete(key) : objectStore.get(key);
        request.onsuccess = () => { result = request.result; };
        request.onerror = () => {};
        transaction.oncomplete = () => resolve(result);
        transaction.onerror = () => reject(transaction.error || request.error || failure('storage-failed', '브라우저 저장소 작업에 실패했습니다.'));
        transaction.onabort = () => reject(transaction.error || request.error || failure('storage-failed', '브라우저 저장소 작업이 취소되었습니다.'));
      } catch (error) { reject(error); }
    });
  }
  function markDisconnected(value) {
    try {
      if (value) global.localStorage.setItem('tool-management-file-disconnected', '1');
      else global.localStorage.removeItem('tool-management-file-disconnected');
    } catch (_) { /* The handle deletion remains the primary disconnection operation. */ }
  }
  function isDisconnected() {
    try { return global.localStorage.getItem('tool-management-file-disconnected') === '1'; }
    catch (_) { return false; }
  }
  function permissionState(value) { return value === 'granted' ? 'granted' : value === 'denied' ? 'denied' : 'prompt'; }
  async function query(handle) {
    if (!handle || typeof handle.queryPermission !== 'function') return 'prompt';
    try { return permissionState(await handle.queryPermission({ mode: 'readwrite' })); }
    catch (error) { throw mapError(error, 'permission'); }
  }
  async function requestPermission(handle) {
    if (await query(handle) === 'granted') return 'granted';
    if (typeof handle.requestPermission !== 'function') return 'prompt';
    try { return permissionState(await handle.requestPermission({ mode: 'readwrite' })); }
    catch (error) { throw mapError(error, 'permission'); }
  }
  function mapError(error, phase) {
    if (error && error.code && typeof error.code === 'string') return error;
    const name = error && error.name;
    if (name === 'NotFoundError') return failure('not-found', '연결한 엑셀 파일을 찾을 수 없습니다. OneDrive 동기화 상태를 확인하고 파일을 다시 연결해 주세요.', error);
    if (name === 'NotAllowedError') return failure('permission-denied', '엑셀 파일의 읽기·쓰기 권한이 없습니다. 「연결 권한 허용」으로 권한을 허용해 주세요.', error);
    if (name === 'SecurityError') return failure('unsupported', '이 창에서 파일 연결을 허용하지 않습니다. Chrome 또는 Edge의 일반 창에서 프로그램을 열어 주세요.', error);
    if (name === 'NoModificationAllowedError') return failure('file-locked', '엑셀 파일이 다른 창이나 프로그램에서 사용 중입니다. Excel을 닫은 뒤 파일을 다시 읽고 시도해 주세요.', error);
    if (name === 'AbortError' && phase === 'picker') return failure('cancelled', '파일 선택을 취소했습니다.', error);
    return failure(phase === 'write' ? 'write-failed' : 'read-failed', phase === 'write' ? '엑셀 파일을 저장할 수 없습니다. 파일 권한과 OneDrive 동기화 상태를 확인해 주세요.' : '엑셀 파일을 읽을 수 없습니다. 파일을 다시 연결해 주세요.', error);
  }
  function updateFailure(error) {
    if (['permission-denied', 'permission-required'].includes(error.code)) { status.connected = false; status.needsPermission = true; }
    if (error.code === 'not-found') { status.connected = false; status.needsReconnect = true; }
    status.statusMessage = error.message;
  }
  async function basicFile(handle) {
    let file;
    try { file = await handle.getFile(); } catch (error) { throw mapError(error, 'read'); }
    if (!/\.xlsx$/i.test(file.name)) throw failure('invalid-file', '매크로 없는 .xlsx 엑셀 파일을 선택해 주세요.');
    if (file.size > 100 * 1024 * 1024) throw failure('invalid-file', '연결할 엑셀 파일은 100 MB 이하여야 합니다.');
    const signature = new Uint8Array(await file.slice(0, 4).arrayBuffer());
    if (signature.length !== 4 || signature[0] !== 80 || signature[1] !== 75 || signature[2] !== 3 || signature[3] !== 4) throw failure('invalid-file', '정상적인 .xlsx 파일이 아닙니다. 암호 없는 엑셀 원본을 선택해 주세요.');
    return file;
  }
  async function digest(bytes) {
    if (!(global.crypto && global.crypto.subtle)) throw failure('unsupported', '이 창에서는 파일 무결성 검사를 사용할 수 없습니다. Chrome 또는 Edge의 일반 창에서 다시 열어 주세요.');
    const output = await global.crypto.subtle.digest('SHA-256', bytes);
    return Array.from(new Uint8Array(output), (n) => n.toString(16).padStart(2, '0')).join('');
  }
  const versionDigest = (version) => { const match = /^sha256:([a-f0-9]{64}):\d+:\d+$/i.exec(String(version || '')); return match ? match[1].toLowerCase() : ''; };
  async function snapshot(handle) {
    const file = await basicFile(handle);
    const bytes = new Uint8Array(await file.arrayBuffer());
    const hash = await digest(bytes);
    return { file, bytes, version: 'sha256:' + hash + ':' + file.size + ':' + Math.max(0, Number(file.lastModified) || 0) };
  }
  function resetConnection() {
    selectedHandle = null;
    status = { connected: false, name: '', mode: MODE, needsPermission: false, needsReconnect: false, persistent: false, statusMessage: '', busy: status.busy, uncertain: false, version: '', lastSavedAt: '' };
  }
  async function connect(options) {
    if (!supported()) throw failure('unsupported', '자동 원본 저장은 Chrome 또는 Edge에서 지원합니다. 지원되는 브라우저로 열고 OneDrive에 동기화된 .xlsx 파일을 선택해 주세요.');
    let handles;
    // Keep the picker within the user's click activation, before waiting for other work.
    try {
      handles = await global.showOpenFilePicker({ multiple: false, excludeAcceptAllOption: true, types: [{ description: 'Excel 통합 문서 (.xlsx)', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }], mode: 'readwrite' });
    } catch (error) { throw mapError(error, 'picker'); }
    const candidate = handles && handles[0];
    if (!candidate || candidate.kind !== 'file') throw failure('invalid-file', '연결할 엑셀 파일을 선택해 주세요.');
    return serial(async () => {
      if (await requestPermission(candidate) !== 'granted') throw failure('permission-denied', '자동 저장을 사용하려면 선택한 파일의 읽기·쓰기 권한을 허용해 주세요.');
      const file = await basicFile(candidate);
      if (options && typeof options.validate === 'function') await options.validate(file);
      let persistent = true, message = '';
      try { await stored('handles', 'put', 'workbook', candidate); markDisconnected(false); }
      catch (_) { persistent = false; message = '이번 창에서만 연결되었습니다. 브라우저가 연결 정보를 저장하지 못했습니다. 원본 자동 저장 전에 백업 저장소를 다시 확인합니다.'; }
      selectedHandle = candidate;
      status = { connected: true, name: file.name, mode: MODE, needsPermission: false, needsReconnect: false, persistent, statusMessage: message, busy: true, uncertain: false, version: '', lastSavedAt: '' };
      return completed({ file });
    });
  }
  async function loadBackup() {
    const saved = await stored('backups', 'get', 'latest');
    if (saved && saved.bytes && saved.name) backup = { name: String(saved.name), bytes: new Uint8Array(saved.bytes), version: String(saved.version || ''), timestamp: String(saved.timestamp || '') };
    return backup;
  }
  async function restore() {
    return serial(async () => {
      let handle;
      try {
        await loadBackup();
        if (isDisconnected()) return completed();
        handle = await stored('handles', 'get', 'workbook');
      } catch (_) {
        status.statusMessage = '이 브라우저에서는 연결 정보·백업 저장소를 열 수 없습니다. 파일을 다시 선택할 수 있지만 자동 저장은 백업이 저장되는 경우에만 가능합니다.';
        return completed();
      }
      if (!handle) return completed();
      selectedHandle = handle;
      status.name = handle.name || '';
      status.persistent = true;
      status.needsReconnect = false;
      try {
        const access = await query(handle);
        if (access !== 'granted') {
          status.connected = false; status.needsPermission = true;
          status.statusMessage = '이전에 연결한 엑셀이 있습니다. 「연결 권한 허용」을 눌러 읽기·쓰기를 다시 허용해 주세요.';
          return completed();
        }
        const file = await basicFile(handle);
        status.connected = true; status.needsPermission = false; status.name = file.name; status.statusMessage = '';
        return completed({ file });
      } catch (error) {
        const mapped = mapError(error, 'read'); updateFailure(mapped);
        if (!['permission-denied', 'permission-required'].includes(mapped.code)) { status.connected = false; status.needsReconnect = true; }
        return completed();
      }
    });
  }
  async function requestAccess() {
    return serial(async () => {
      if (!selectedHandle) throw failure('not-connected', 'OneDrive에 동기화된 엑셀 파일을 먼저 연결해 주세요.');
      try {
        const access = await requestPermission(selectedHandle);
        if (access !== 'granted') {
          status.connected = false; status.needsPermission = true;
          status.statusMessage = '읽기·쓰기 권한이 허용되지 않았습니다. 권한을 허용하거나 파일을 다시 연결해 주세요.';
          return completed();
        }
        const file = await basicFile(selectedHandle);
        status.connected = true; status.needsPermission = false; status.needsReconnect = false; status.name = file.name; status.statusMessage = '';
        return completed({ file });
      } catch (error) { const mapped = mapError(error, 'permission'); updateFailure(mapped); return completed(); }
    });
  }
  async function accessibleHandle() {
    if (!selectedHandle) throw failure('not-connected', 'OneDrive에 동기화된 엑셀 파일을 먼저 연결해 주세요.');
    const permission = await query(selectedHandle);
    if (permission !== 'granted') throw failure(permission === 'denied' ? 'permission-denied' : 'permission-required', '연결한 파일의 읽기·쓰기 권한을 다시 허용해 주세요.');
    return selectedHandle;
  }
  async function read() {
    return serial(async () => {
      try {
        const result = await snapshot(await accessibleHandle());
        status.connected = true; status.needsPermission = false; status.needsReconnect = false;
        status.name = result.file.name; status.version = result.version; status.uncertain = false; status.statusMessage = '';
        return result;
      } catch (error) { const mapped = mapError(error, 'read'); updateFailure(mapped); throw mapped; }
    });
  }
  function checkRevision(verify) {
    if (typeof verify !== 'function') return;
    let current = false;
    try { current = verify() === true; } catch (_) { /* A failed revision check cannot authorize a write. */ }
    if (!current) throw failure('stale', '저장 준비 중 사진이나 연결한 엑셀이 바뀌었습니다. 원본을 수정하지 않았습니다. 최신 내용을 확인하고 다시 시도해 주세요.');
  }
  async function writerFor(handle) {
    try { return await handle.createWritable({ keepExistingData: true, mode: 'exclusive' }); }
    catch (error) {
      // Older engines ignore unknown dictionary members or reject the mode enum.
      if (error instanceof TypeError && /mode|exclusive|enum/i.test(error.message || '')) return handle.createWritable({ keepExistingData: true });
      throw error;
    }
  }
  function backupName(name, timestamp) {
    return String(name).replace(/\.xlsx$/i, '') + '_자동저장전_' + timestamp.replace(/[-:]/g, '').replace('T', '_').slice(0, 15) + '.xlsx';
  }
  async function performSave(options) {
    if (!options || !(options.bytes instanceof Uint8Array) || !options.bytes.length) throw failure('invalid-file', '저장할 엑셀 내용이 없습니다.');
    const expected = versionDigest(options.expectedVersion);
    if (!expected) throw failure('conflict', '저장 전 원본 엑셀을 다시 읽어 주세요.');
    if (status.uncertain) throw failure('requires-read', '이전 저장의 완료 여부를 먼저 확인해야 합니다. 연결한 원본을 다시 읽은 뒤 저장해 주세요.');
    const bytes = options.bytes.slice();
    if (bytes.length > 100 * 1024 * 1024 || bytes.length < 4 || bytes[0] !== 80 || bytes[1] !== 75 || bytes[2] !== 3 || bytes[3] !== 4) throw failure('invalid-file', '저장할 내용이 정상적인 .xlsx 파일이 아닙니다.');
    const handle = await accessibleHandle();
    const desiredDigest = await digest(bytes);
    let writable = null, writeAttempted = false, closeAttempted = false, closeCompleted = false;
    try {
      checkRevision(options.verifyCurrent);
      writable = await writerFor(handle);
      const latest = await snapshot(handle);
      if (versionDigest(latest.version) !== expected) throw failure('conflict', '다른 사용자나 Excel이 원본을 변경했습니다. 원본을 덮어쓰지 않았습니다. 최신 파일을 다시 읽고 사진을 확인한 뒤 저장해 주세요.');
      checkRevision(options.verifyCurrent);
      const timestamp = new Date().toISOString();
      backup = { name: backupName(latest.file.name, timestamp), bytes: latest.bytes.slice(), version: latest.version, timestamp };
      try { await stored('backups', 'put', 'latest', backup); }
      catch (error) { throw failure('backup-failed', '저장 전 원본 백업을 브라우저에 보관하지 못해 자동 저장을 중단했습니다. 원본은 수정하지 않았습니다. 브라우저 저장 공간·설정을 확인하거나 수정본을 다운로드해 주세요.', error); }
      checkRevision(options.verifyCurrent);
      writeAttempted = true;
      await writable.write({ type: 'write', position: 0, data: bytes });
      await writable.truncate(bytes.length);
      // Staged bytes are still uncommitted. Detect an outside writer during
      // backup creation or staging before atomically replacing the original.
      const beforeCommit = await snapshot(handle);
      if (versionDigest(beforeCommit.version) !== expected) throw failure('conflict', '저장 준비 중 다른 사용자나 Excel이 원본을 변경했습니다. 원본을 덮어쓰지 않았습니다. 최신 파일을 다시 읽어 주세요.');
      checkRevision(options.verifyCurrent);
      closeAttempted = true;
      await writable.close();
      closeCompleted = true; writable = null;
      const actual = await snapshot(handle);
      if (versionDigest(actual.version) !== desiredDigest) throw failure('verification-failed', '저장 후 파일 내용이 예상과 다릅니다. 다른 프로그램의 변경 여부를 확인하고 원본을 다시 읽어 주세요. 자동 재시도하지 않습니다.');
      status.connected = true; status.needsPermission = false; status.needsReconnect = false;
      status.name = actual.file.name; status.version = actual.version; status.lastSavedAt = new Date().toISOString(); status.uncertain = false; status.statusMessage = '';
      return { name: actual.file.name, version: actual.version, backupAvailable: true };
    } catch (error) {
      let aborted = false;
      if (writable && typeof writable.abort === 'function') {
        try { await writable.abort(); aborted = true; } catch (_) { /* A failed abort leaves the outcome uncertain. */ }
      }
      let mapped = mapError(error, 'write');
      const safelyCancelledBeforeCommit = ['stale', 'conflict'].includes(mapped.code) && aborted && !closeAttempted;
      if ((writeAttempted || closeAttempted || closeCompleted) && !safelyCancelledBeforeCommit) {
        status.uncertain = true;
        if (mapped.code !== 'verification-failed') mapped = failure(mapped.code === 'write-failed' ? 'write-failed' : mapped.code, '저장 완료 여부를 확인할 수 없습니다. 자동 재시도하지 않습니다. 연결한 원본을 다시 읽어 확인해 주세요. 저장 전 백업은 보관되어 있습니다.', mapped);
      }
      updateFailure(mapped);
      throw mapped;
    }
  }
  async function save(options) {
    return serial(async () => {
      try {
        if (global.navigator && global.navigator.locks && typeof global.navigator.locks.request === 'function') return await global.navigator.locks.request(LOCK, { mode: 'exclusive' }, () => performSave(options));
        return await performSave(options);
      } catch (error) { const mapped = mapError(error, 'write'); updateFailure(mapped); throw mapped; }
    });
  }
  async function takeBackup() {
    if (!backup) {
      try { await loadBackup(); } catch (_) { return null; }
    }
    return backup ? { name: backup.name, bytes: backup.bytes.slice(), version: backup.version, timestamp: backup.timestamp } : null;
  }
  async function disconnect() {
    return serial(async () => {
      markDisconnected(true);
      let message = '';
      try { await stored('handles', 'delete', 'workbook'); }
      catch (_) { message = '이번 창의 연결은 해제했습니다. 브라우저의 이전 연결 정보 삭제는 완료하지 못했습니다. 다음 실행 시 연결 상태를 확인해 주세요.'; }
      resetConnection(); status.statusMessage = message;
      return completed();
    });
  }
  global.ToolConnectedFile = Object.freeze({ connect, restore, requestAccess, read, save, disconnect, takeBackup, inspect });
})(window);
