/* Device workbook storage for browsers without original-file write access.
 * This module never uploads a workbook or changes a OneDrive/SharePoint original.
 */
(function (global) {
  'use strict';
  const DATABASE = 'tool-management-device-workbook-v1';
  const STORE = 'workspace', MODE = 'device-workbook', LIMIT = 100 * 1024 * 1024;
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  let databasePromise = null, queue = Promise.resolve(), backupAvailable = false;
  const empty = () => ({ connected: false, name: '', mode: MODE, needsPermission: false, needsReconnect: false,
    persistent: false, statusMessage: '', busy: false, uncertain: false, version: '', lastSavedAt: '' });
  let status = empty();
  function failure(code, message, cause) {
    const error = new Error(message); error.code = code; if (cause) error.cause = cause; return error;
  }
  function supported() {
    return global.isSecureContext !== false && !!(global.indexedDB && global.crypto && global.crypto.subtle && global.crypto.getRandomValues);
  }
  function inspect() { return Object.assign({}, status, { supported: supported(), backupAvailable }); }
  function completed(extra) { return Object.assign(inspect(), { busy: false }, extra || {}); }
  function serial(action) {
    const next = queue.then(async () => { status.busy = true; try { return await action(); } finally { status.busy = false; } });
    queue = next.catch(() => {}); return next;
  }
  function storageError(error, code) {
    if (error && typeof error.code === 'string') return error;
    if (error && error.name === 'QuotaExceededError') return failure('storage-full', '이 기기의 앱 저장 공간이 부족해 저장하지 못했습니다. 이전 엑셀은 유지됩니다. 수정본을 저장·공유한 뒤 저장 공간을 확인해 주세요.', error);
    return failure(code || 'storage-failed', '이 기기의 앱 저장소를 사용할 수 없습니다. 일반 브라우저 창에서 다시 열거나 수정본을 저장·공유해 주세요.', error);
  }
  function openDatabase() {
    if (databasePromise) return databasePromise;
    databasePromise = new Promise((resolve, reject) => {
      if (!supported()) { reject(failure('unsupported', '이 창에서는 기기에 엑셀을 보관할 수 없습니다. HTTPS 앱을 일반 브라우저 창에서 열어 주세요.')); return; }
      let request;
      try { request = global.indexedDB.open(DATABASE, 1); }
      catch (error) { reject(storageError(error)); return; }
      request.onupgradeneeded = () => { if (!request.result.objectStoreNames.contains(STORE)) request.result.createObjectStore(STORE); };
      request.onsuccess = () => {
        const database = request.result;
        database.onversionchange = () => { database.close(); databasePromise = null; };
        resolve(database);
      };
      request.onerror = () => reject(storageError(request.error));
      request.onblocked = () => reject(failure('storage-blocked', '다른 앱 창이 저장소 변경을 막고 있습니다. 다른 창을 닫은 뒤 다시 시도해 주세요.'));
    });
    databasePromise.catch(() => { databasePromise = null; }); return databasePromise;
  }
  function transaction(database, mode) {
    if (mode === 'readonly') return database.transaction(STORE, mode);
    try { return database.transaction(STORE, mode, { durability: 'strict' }); }
    catch (error) { if (error instanceof TypeError) return database.transaction(STORE, mode); throw error; }
  }
  async function records() {
    const database = await openDatabase();
    return new Promise((resolve, reject) => {
      let tx, workbook, backup;
      try {
        tx = transaction(database, 'readonly');
        const store = tx.objectStore(STORE), bookRequest = store.get('workbook'), backupRequest = store.get('backup');
        bookRequest.onsuccess = () => { workbook = bookRequest.result; };
        backupRequest.onsuccess = () => { backup = backupRequest.result; };
        tx.oncomplete = () => resolve({ workbook, backup });
        tx.onerror = tx.onabort = () => reject(storageError(tx.error));
      } catch (error) { reject(storageError(error)); }
    });
  }
  function bytesOf(value) {
    let bytes;
    if (value instanceof Uint8Array) bytes = value.slice();
    else if (value instanceof ArrayBuffer) bytes = new Uint8Array(value.slice(0));
    else throw failure('invalid-file', '저장할 엑셀 내용이 없습니다.');
    if (bytes.length < 4 || bytes.length > LIMIT || bytes[0] !== 80 || bytes[1] !== 75 || bytes[2] !== 3 || bytes[3] !== 4)
      throw failure('invalid-file', '암호 없는 .xlsx 엑셀 파일을 100 MB 이하로 선택해 주세요.');
    return bytes;
  }
  async function hash(bytes) {
    if (!supported()) throw failure('unsupported', '파일 무결성 검사를 사용할 수 없습니다. HTTPS 앱으로 다시 열어 주세요.');
    return Array.from(new Uint8Array(await global.crypto.subtle.digest('SHA-256', bytes)), value => value.toString(16).padStart(2, '0')).join('');
  }
  function revision() {
    const bytes = new Uint8Array(16); global.crypto.getRandomValues(bytes);
    return Array.from(bytes, value => value.toString(16).padStart(2, '0')).join('');
  }
  function version(digest, id) { return 'device:sha256:' + digest + ':' + id; }
  function checkRevision(verify) {
    if (typeof verify !== 'function') return;
    let valid = false; try { valid = verify() === true; } catch (_) {}
    if (!valid) throw failure('stale', '저장 준비 중 사진이나 엑셀이 바뀌었습니다. 최신 내용을 확인한 뒤 다시 시도해 주세요.');
  }
  function usableRecord(record) {
    return !!record && typeof record.name === 'string' && /\.xlsx$/i.test(record.name) && typeof record.digest === 'string'
      && /^[a-f0-9]{64}$/.test(record.digest) && typeof record.revision === 'string' && /^[a-f0-9]{32}$/.test(record.revision)
      && record.version === version(record.digest, record.revision);
  }
  function adopt(record, backup) {
    backupAvailable = !!backup;
    Object.assign(status, { connected: true, name: record.name, persistent: true, needsPermission: false, needsReconnect: false,
      statusMessage: '이 기기의 앱에 엑셀을 보관합니다. 회사 공유 원본에 적용하려면 수정본을 저장·공유하세요.',
      version: record.version, lastSavedAt: record.savedAt || '' });
  }
  async function snapshot(record) {
    if (!usableRecord(record)) throw failure('not-connected', '이 기기에 사용할 엑셀 파일을 먼저 가져와 주세요.');
    const bytes = bytesOf(record.bytes);
    if (await hash(bytes) !== record.digest) throw failure('storage-corrupt', '기기에 보관된 엑셀을 확인할 수 없습니다. 원본 파일을 다시 가져와 주세요.');
    return { file: new File([bytes], record.name, { type: MIME, lastModified: Math.max(0, Number(record.lastModified) || 0) }), bytes, version: record.version };
  }
  async function atomic(change) {
    const database = await openDatabase();
    return new Promise((resolve, reject) => {
      let tx, result, ownError;
      const abort = (error) => { ownError = error; try { tx.abort(); } catch (_) {} };
      try {
        tx = transaction(database, 'readwrite');
        tx.oncomplete = () => resolve(result);
        tx.onerror = tx.onabort = () => reject(ownError || storageError(tx.error));
        const store = tx.objectStore(STORE), request = store.get('workbook');
        request.onsuccess = () => {
          try { result = change(store, request.result, abort); }
          catch (error) { abort(storageError(error)); }
        };
      } catch (error) { reject(storageError(error)); }
    });
  }
  async function connect(file, options) {
    return serial(async () => {
      if (!(file instanceof Blob) || !/\.xlsx$/i.test(String(file.name || '')) || file.size > LIMIT)
        throw failure('invalid-file', '암호 없는 .xlsx 엑셀 파일을 100 MB 이하로 선택해 주세요.');
      const bytes = bytesOf(await file.arrayBuffer());
      if (options && typeof options.validate === 'function') await options.validate(file);
      const digest = await hash(bytes), id = revision();
      const record = { name: file.name.replace(/[\\/\x00-\x1f]/g, '_'), bytes, digest, revision: id,
        version: version(digest, id), lastModified: Math.max(0, Number(file.lastModified) || Date.now()), savedAt: '' };
      await atomic((store) => { store.put(record, 'workbook'); store.delete('backup'); });
      adopt(record, null); return completed({ file: (await snapshot(record)).file });
    });
  }
  async function read() {
    return serial(async () => {
      try {
        const found = await records(), result = await snapshot(found.workbook);
        adopt(found.workbook, found.backup); return result;
      } catch (error) {
        if (['not-connected', 'storage-corrupt'].includes(error.code)) Object.assign(status, { connected: false, needsReconnect: true, statusMessage: error.message });
        throw error;
      }
    });
  }
  async function restore() {
    return serial(async () => {
      try {
        const found = await records();
        if (!found.workbook) { status = Object.assign(empty(), { busy: true }); backupAvailable = false; return completed(); }
        const result = await snapshot(found.workbook); adopt(found.workbook, found.backup); return completed({ file: result.file });
      } catch (error) {
        Object.assign(status, { connected: false, persistent: false, needsReconnect: true, statusMessage: error.message });
        throw error;
      }
    });
  }
  async function save(options) {
    return serial(async () => {
      if (!options || typeof options.expectedVersion !== 'string' || !/^device:sha256:[a-f0-9]{64}:[a-f0-9]{32}$/.test(options.expectedVersion))
        throw failure('conflict', '저장 전 기기에 보관된 엑셀을 다시 읽어 주세요.');
      const bytes = bytesOf(options.bytes), digest = await hash(bytes), id = revision(), stamp = new Date().toISOString();
      checkRevision(options.verifyCurrent);
      const saved = await atomic((store, current, abort) => {
        if (!usableRecord(current) || current.version !== options.expectedVersion) {
          abort(failure('conflict', '다른 앱 창에서 엑셀이 변경됐습니다. 최신 파일을 다시 읽고 사진을 확인해 주세요.')); return;
        }
        checkRevision(options.verifyCurrent);
        const backup = { name: current.name.replace(/\.xlsx$/i, '') + '_저장전백업.xlsx', bytes: current.bytes,
          version: current.version, timestamp: stamp };
        const next = { name: current.name, bytes, digest, revision: id, version: version(digest, id), lastModified: Date.now(), savedAt: stamp };
        try { store.put(backup, 'backup'); }
        catch (error) { abort(storageError(error, 'backup-failed')); return; }
        try { store.put(next, 'workbook'); checkRevision(options.verifyCurrent); }
        catch (error) { abort(storageError(error, 'write-failed')); return; }
        return { next, backup };
      });
      adopt(saved.next, saved.backup); return completed({ version: saved.next.version, bytes: bytes.slice() });
    });
  }
  function commit(expectedVersion, bytes, options) { return save(Object.assign({}, options || {}, { expectedVersion, bytes })); }
  async function takeBackup() {
    return serial(async () => {
      const found = await records(), saved = found.backup; backupAvailable = !!saved;
      return saved && saved.bytes && saved.name ? { name: String(saved.name), bytes: bytesOf(saved.bytes), version: String(saved.version || ''), timestamp: String(saved.timestamp || '') } : null;
    });
  }
  async function disconnect() {
    return serial(async () => {
      await atomic(store => { store.clear(); });
      status = Object.assign(empty(), { busy: true }); backupAvailable = false; return completed();
    });
  }
  const api = Object.freeze({ supported, inspect, connect, read, save, commit, restore, disconnect, clear: disconnect, takeBackup });
  global.ToolMobileFile = api; global.ToolDeviceFile = api;
})(typeof window !== 'undefined' ? window : globalThis);
