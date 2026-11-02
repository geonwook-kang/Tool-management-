/* Short-lived, device-only checkpoint made immediately before Microsoft login.
 * Auth tokens/configuration and browser URLs/handles are never checkpointed.
 */
(function (global) {
  'use strict';
  const DATABASE = 'tool-management-auth-resume-v1', STORE = 'pending', KEY = 'queue';
  const TTL = 10 * 60 * 1000, MAX_PHOTOS = 100, MAX_BYTES = 128 * 1024 * 1024;
  let databasePromise = null, queue = Promise.resolve(), memoryOwner = '';
  function failure(code, message, cause) { const error = new Error(message); error.code = code; if (cause) error.cause = cause; return error; }
  function plain(value) { return value != null && Object.getPrototypeOf(value) === Object.prototype; }
  function unavailable(error) {
    if (error && typeof error.code === 'string') return error;
    const full = error && error.name === 'QuotaExceededError';
    return failure(full ? 'resume-storage-full' : 'resume-storage-failed', full
      ? '로그인 전에 사진을 기기에 보관할 공간이 부족합니다. 로그인으로 이동하지 않았습니다. 사진을 줄이거나 저장 공간을 확인해 주세요.'
      : '로그인 전에 사진을 기기에 보관하지 못했습니다. 현재 화면의 사진을 유지한 채 저장소 설정을 확인해 주세요.', error);
  }
  function serial(action) { const next = queue.then(action); queue = next.catch(() => {}); return next; }
  function owner() {
    if (memoryOwner) return memoryOwner;
    try { const saved = global.sessionStorage.getItem('tool-manager-resume-owner'); if (/^[a-f0-9]{32}$/.test(saved || '')) { memoryOwner = saved; return saved; } } catch (_) {}
    if (!global.crypto || typeof global.crypto.getRandomValues !== 'function') throw unavailable();
    const bytes = new Uint8Array(16); global.crypto.getRandomValues(bytes);
    memoryOwner = Array.from(bytes, value => value.toString(16).padStart(2, '0')).join('');
    try { global.sessionStorage.setItem('tool-manager-resume-owner', memoryOwner); } catch (_) {}
    return memoryOwner;
  }
  function openDatabase() {
    if (databasePromise) return databasePromise;
    databasePromise = new Promise((resolve, reject) => {
      if (!global.indexedDB || global.isSecureContext === false) { reject(unavailable()); return; }
      let request;
      try { request = global.indexedDB.open(DATABASE, 1); } catch (error) { reject(unavailable(error)); return; }
      request.onupgradeneeded = () => { if (!request.result.objectStoreNames.contains(STORE)) request.result.createObjectStore(STORE); };
      request.onsuccess = () => { const database = request.result; database.onversionchange = () => { database.close(); databasePromise = null; }; resolve(database); };
      request.onerror = () => reject(unavailable(request.error));
      request.onblocked = () => reject(failure('resume-storage-blocked', '다른 앱 창을 닫은 뒤 로그인을 다시 시도해 주세요. 현재 사진은 이 화면에 남아 있습니다.'));
    });
    databasePromise.catch(() => { databasePromise = null; }); return databasePromise;
  }
  function contextOf(value) {
    if (!plain(value) || value.mode !== 'microsoft365') throw failure('resume-context-invalid', '사진을 이어서 처리할 회사 엑셀 연결을 먼저 확인해 주세요.');
    const result = { mode: 'microsoft365', driveId: '', itemId: '', shareUrl: '', uncertain: value.uncertain === true };
    for (const key of ['driveId', 'itemId', 'shareUrl']) if (typeof value[key] === 'string') {
      if (value[key].length > (key === 'shareUrl' ? 8192 : 1024)) throw failure('resume-context-invalid', '회사 엑셀 연결 정보가 올바르지 않습니다.');
      result[key] = value[key];
    }
    if (result.shareUrl) {
      let parsed; try { parsed = new URL(result.shareUrl); } catch (_) { throw failure('resume-context-invalid', '회사 엑셀 HTTPS 링크를 확인해 주세요.'); }
      if (parsed.protocol !== 'https:' || parsed.username || parsed.password) throw failure('resume-context-invalid', '회사 엑셀 HTTPS 링크를 확인해 주세요.');
      parsed.hash = ''; result.shareUrl = parsed.href;
    }
    if (!(result.driveId && result.itemId) && !result.shareUrl) throw failure('resume-context-invalid', '사진을 이어서 처리할 회사 엑셀을 먼저 선택해 주세요.');
    return result;
  }
  function sameContext(a, b) {
    if (a.mode !== b.mode) return false;
    if (a.driveId && a.itemId && b.driveId && b.itemId) return a.driveId === b.driveId && a.itemId === b.itemId;
    return !!a.shareUrl && !!b.shareUrl && a.shareUrl === b.shareUrl;
  }
  function charge(budget, amount) {
    budget.bytes += amount;
    if (budget.bytes > MAX_BYTES) throw failure('resume-too-large', '로그인 전에 보관할 사진이 너무 많습니다. 현재 사진을 유지한 채 사진을 나누어 처리해 주세요.');
  }
  function fields(value, strings, numbers, booleans, budget) {
    const result = {}; if (!plain(value)) return result;
    for (const key of strings || []) if (typeof value[key] === 'string') { charge(budget, value[key].length * 2); result[key] = value[key]; }
    for (const key of numbers || []) if (Number.isFinite(value[key]) && typeof value[key] === 'number') result[key] = value[key];
    for (const key of booleans || []) if (typeof value[key] === 'boolean') result[key] = value[key];
    return result;
  }
  function stringList(values, budget) { return (Array.isArray(values) ? values : []).filter(value => typeof value === 'string').map(value => { charge(budget, value.length * 2); return value; }); }
  function box(value) {
    if (!plain(value)) return undefined;
    const result = {}; for (const key of ['x0', 'y0', 'x1', 'y1']) { if (typeof value[key] !== 'number' || !Number.isFinite(value[key])) return undefined; result[key] = value[key]; }
    return result;
  }
  function imageOf(value, budget) {
    if (!plain(value)) return null;
    let bytes;
    if (value.photoBytes instanceof Uint8Array) bytes = value.photoBytes.slice();
    else if (value.photoBytes instanceof ArrayBuffer) bytes = new Uint8Array(value.photoBytes.slice(0));
    else return null;
    charge(budget, bytes.byteLength);
    const image = fields(value, ['photoExtension'], ['photoWidth', 'photoHeight'], [], budget);
    if (!image.photoExtension || !/^(?:jpe?g|png|webp)$/i.test(image.photoExtension) || !(image.photoWidth > 0) || !(image.photoHeight > 0))
      throw failure('resume-photo-invalid', '처리 중인 사진 정보를 보관하지 못했습니다. 현재 화면에서 다시 확인해 주세요.');
    image.photoBytes = bytes; return image;
  }
  function resultOf(value, budget) {
    if (!plain(value)) return null;
    const result = fields(value, ['text', 'documentType', 'error', 'errorCode', 'code'], ['confidence'], ['failed', 'engineFailure', 'cancelled', 'identityRecoveryIncomplete', 'fastPath'], budget);
    if (value.confidence === null) result.confidence = null;
    result.warnings = stringList(value.warnings, budget);
    const passes = Array.isArray(value.passes) ? value.passes : [];
    if (passes.length > 100) throw failure('resume-too-large', '보관할 인식 기록이 너무 많습니다. 현재 화면에서 사진을 나누어 처리해 주세요.');
    result.passes = passes.map(pass => {
      const clean = fields(pass, ['text', 'source', 'smaxRole', 'smaxLabelText'], ['confidence'], ['smaxLabelVerified'], budget);
      if (pass && pass.confidence === null) clean.confidence = null;
      for (const key of ['smaxLabelBBox', 'smaxValueBBox']) { const value = box(pass && pass[key]); if (value) clean[key] = value; }
      const words = pass && Array.isArray(pass.words) ? pass.words : [];
      budget.words += words.length;
      if (budget.words > 100000) throw failure('resume-too-large', '보관할 인식 기록이 너무 많습니다. 현재 사진을 나누어 처리해 주세요.');
      clean.words = words.map(word => {
        const copied = fields(word, ['text'], ['confidence'], [], budget);
        if (word && word.confidence === null) copied.confidence = null;
        for (const key of ['bbox', 'lineBox']) { const value = box(word && word[key]); if (value) copied[key] = value; }
        return copied;
      });
      return clean;
    });
    return result;
  }
  function photoOf(value, budget) {
    if (!plain(value) || !(value.file instanceof File)) throw failure('resume-photo-invalid', '로그인 전에 원본 사진 파일을 보관할 수 없습니다. 현재 화면에서 사진을 다시 확인해 주세요.');
    charge(budget, value.file.size);
    const photo = fields(value, ['id', 'approval', 'error', 'photoKindOverride'], [], ['edited', 'manualEntry', 'sharedConflict'], budget);
    photo.file = new File([value.file], value.file.name, { type: value.file.type, lastModified: value.file.lastModified });
    photo.image = imageOf(value.image, budget);
    const update = ['toolId', 'name', 'serial', 'last', 'due', 'photoKind'];
    photo.draft = fields(value.draft, update, [], ['updateDue'], budget);
    photo.result = resultOf(value.result, budget);
    if (plain(value.review)) {
      photo.review = fields(value.review, update, [], ['updateDue'], budget);
      if (photo.image) Object.assign(photo.review, photo.image);
    } else photo.review = null;
    photo.assessment = plain(value.assessment) ? fields(value.assessment, ['toolId', 'name', 'serial', 'observedSerial', 'last', 'due', 'photoKind', 'matchMethod', 'dateMethod', 'ownerEnglish'], [], ['eligible', 'updateDue'], budget) : null;
    if (photo.assessment) { photo.assessment.reasons = stringList(value.assessment.reasons, budget); photo.assessment.warnings = stringList(value.assessment.warnings, budget); }
    photo.saved = plain(value.saved) ? fields(value.saved, ['id', 'sheet', 'name', 'ownerBaseName', 'serial', 'asset', 'last', 'due', 'cycle', 'manufacturer', 'model', 'description', 'photoKind', 'version', 'bindingIdentity', 'sourceName'], ['row', 'interval', 'months', 'photoCount', 'smaxPhotoCount'], ['numericSerial', 'hasSmaxPhotoColumn'], budget) : null;
    return photo;
  }
  function expired(record) { const now = Date.now(); return !record || !Number.isFinite(record.createdAt) || !Number.isFinite(record.expiresAt) || record.expiresAt - record.createdAt !== TTL || now >= record.expiresAt || now < record.createdAt - 60000; }
  async function atomic(action) {
    const database = await openDatabase();
    return new Promise((resolve, reject) => {
      let tx, result, ownError;
      const abort = error => { ownError = error; try { tx.abort(); } catch (_) {} };
      try {
        try { tx = database.transaction(STORE, 'readwrite', { durability: 'strict' }); }
        catch (error) { if (error instanceof TypeError) tx = database.transaction(STORE, 'readwrite'); else throw error; }
        tx.oncomplete = () => resolve(result);
        tx.onerror = tx.onabort = () => reject(ownError || unavailable(tx.error));
        const store = tx.objectStore(STORE), request = store.get(KEY);
        request.onsuccess = () => { try { result = action(store, request.result, abort); } catch (error) { abort(unavailable(error)); } };
      } catch (error) { reject(unavailable(error)); }
    });
  }
  function details(record) { return { pending: true, count: record.photos.length, createdAt: record.createdAt, expiresAt: record.expiresAt, sameTab: record.owner === owner() }; }
  async function save(options) {
    return serial(async () => {
      if (!plain(options) || !Array.isArray(options.photos) || options.photos.length > MAX_PHOTOS) throw failure('resume-photo-invalid', '로그인 전에 보관할 사진 목록을 확인해 주세요.');
      const context = contextOf(options.context), budget = { bytes: 0, words: 0 }, photos = options.photos.map(photo => photoOf(photo, budget));
      const stamp = Date.now(), record = { context, photos, owner: owner(), createdAt: stamp, expiresAt: stamp + TTL };
      await atomic((store, previous) => {
        if (previous && !expired(previous)) {
          if (previous.owner !== record.owner) throw failure('resume-other-tab', '다른 앱 창에서 로그인을 기다리는 사진이 있습니다. 그 창에서 로그인을 마치거나 사진 보관을 지운 뒤 다시 시도해 주세요.');
          if (!sameContext(previous.context, context)) throw failure('resume-context-conflict', '다른 회사 엑셀로 이어서 처리할 사진이 보관돼 있습니다. 현재 사진을 확인한 뒤 보관을 지워 주세요.');
          if (previous.photos.length && !photos.length) throw failure('resume-pending-not-empty', '로그인을 기다리는 사진이 보관돼 있습니다. 보관한 사진을 복원하거나 지운 뒤 다시 로그인해 주세요.');
        }
        store.put(record, KEY);
      });
      return details(record);
    });
  }
  async function inspect() { return serial(() => atomic((store, record) => { if (expired(record)) { if (record) store.delete(KEY); return { pending: false, count: 0 }; } return details(record); })); }
  async function peek() { return serial(() => atomic((store, record) => { if (expired(record)) { if (record) store.delete(KEY); return null; } return Object.assign(details(record), { context: Object.assign({}, record.context) }); })); }
  async function take(expectedContext) {
    const expected = expectedContext == null ? null : contextOf(expectedContext);
    return serial(() => atomic((store, record) => {
      if (expired(record)) { if (record) store.delete(KEY); return null; }
      if (record.owner !== owner()) return null;
      if (expected && !sameContext(record.context, expected)) return null;
      store.delete(KEY); return { context: record.context, photos: record.photos, createdAt: record.createdAt, expiresAt: record.expiresAt };
    }));
  }
  async function clear() { return serial(() => atomic(store => { store.delete(KEY); return { pending: false, count: 0 }; })); }
  global.ToolResume = Object.freeze({ save, inspect, peek, take, restore: take, clear });
})(typeof window !== 'undefined' ? window : globalThis);
