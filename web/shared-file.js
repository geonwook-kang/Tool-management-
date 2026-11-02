/* Shared workbook transport. The owner's cloud settings choose the destination;
 * browser-local remembered files and caller-provided paths never choose it.
 */
(function (global) {
  'use strict';
  let binding = null, lastRead = null, validator = null, queue = Promise.resolve();
  const initial = () => ({ mode: 'shared', sourceMode: 'shared', connected: false, name: '', version: '', busy: false,
    needsPermission: false, needsReconnect: false, requiresSignIn: false, persistent: false,
    uncertain: false, destinationChanged: false, configurationVerified: false, statusMessage: '', lastSavedAt: '' });
  let status = initial();
  const failure = (code, message) => { const error = new Error(message); error.code = code; return error; };
  function graph() {
    if (!global.ToolGraphFile) throw failure('shared-unavailable', '공동 엑셀 연결 모듈을 준비하지 못했습니다. 프로그램을 다시 열어 주세요.');
    return global.ToolGraphFile;
  }
  function config() {
    if (!global.ToolSharedWorkspace || typeof global.ToolSharedWorkspace.refresh !== 'function') throw failure('shared-unavailable', '관리자의 공동 연결 설정을 먼저 열어 주세요.');
    return global.ToolSharedWorkspace;
  }
  const cloneTarget = value => value ? { driveId: value.driveId, itemId: value.itemId, name: value.name } : null;
  const sameTarget = (a, b) => !!(a && b && a.driveId === b.driveId && a.itemId === b.itemId);
  const safeString = (value, maximum = 1024) => typeof value === 'string' && value.length > 0 && value.length <= maximum && !/[\u0000-\u001f\u007f]/.test(value);
  function workspaceValue(raw) {
    if (!raw || raw.v !== 1 || !safeString(raw.workspaceId, 200) || !safeString(raw.ownerObjectId, 200)
      || !raw.workbook || !safeString(raw.workbook.driveId) || !safeString(raw.workbook.itemId)
      || !safeString(raw.workbook.name, 255) || !/\.xlsx$/i.test(raw.workbook.name)
      || !safeString(raw.configVersion, 4096) || raw.configVersion === '*') throw failure('shared-config-invalid', '공동 저장 대상 설정의 파일 ID와 버전을 확인하지 못했습니다.');
    const target = cloneTarget(raw.workbook);
    const identity = JSON.stringify([raw.workspaceId, raw.ownerObjectId, target.driveId, target.itemId]);
    return { workspaceId: raw.workspaceId, ownerObjectId: raw.ownerObjectId, target, identity,
      fingerprint: JSON.stringify([identity, raw.configVersion]), configVersion: raw.configVersion };
  }
  async function authoritative() {
    try { return workspaceValue(await config().refresh()); }
    catch (error) { status.connected = false; throw error; }
  }
  function supported() { try { return !!(global.ToolSharedWorkspace && global.ToolGraphFile && graph().inspect().supported); } catch (_) { return false; } }
  function inspect() {
    const remote = global.ToolGraphFile ? global.ToolGraphFile.inspect() : {}, matches = sameTarget(remote.target, binding && binding.target);
    return { ...status, supported: supported(), connected: !!(status.connected && matches && remote.connected),
      busy: status.busy || !!(matches && remote.busy), uncertain: status.uncertain || !!(matches && remote.uncertain),
      backupAvailable: !!(matches && remote.backupAvailable), persistent: !!(binding && remote.persistent),
      workspaceId: binding ? binding.workspaceId : '', ownerObjectId: binding ? binding.ownerObjectId : '',
      bindingIdentity: binding ? binding.identity : '', bindingFingerprint: binding ? binding.fingerprint : '',
      configVersion: binding ? binding.configVersion : '', target: cloneTarget(binding && binding.target) };
  }
  const complete = extra => ({ ...inspect(), busy: false, ...(extra || {}) });
  function serial(action) { const task = queue.then(async () => { status.busy = true; try { return await action(); } finally { status.busy = false; } }); queue = task.catch(() => {}); return task; }
  function rememberFailure(error) {
    const remote = global.ToolGraphFile ? global.ToolGraphFile.inspect() : {}, matches = sameTarget(remote.target, binding && binding.target);
    status.statusMessage = error.message;
    status.uncertain = status.uncertain || !!(matches && remote.uncertain);
    if (['auth-required', 'auth-expired', 'expired'].includes(error.code) || error.httpStatus === 401) {
      status.connected = false; status.needsPermission = true; status.requiresSignIn = true;
    } else if (error.httpStatus === 403 || error.code === 'permission-denied') {
      status.connected = false; status.needsPermission = true;
    } else if (['shared-destination-changed', 'shared-config-invalid', 'shared-not-connected', 'not-found', 'invalid-target'].includes(error.code)) {
      status.connected = false; status.needsReconnect = true;
    }
  }
  function mapped(error) { return error && typeof error.code === 'string' ? error : failure('shared-request-failed', '공동 엑셀 요청을 완료하지 못했습니다. 회사 로그인과 연결 상태를 확인해 주세요.'); }
  function copyBytes(value) {
    const bytes = value instanceof Uint8Array ? value.slice() : value instanceof ArrayBuffer ? new Uint8Array(value.slice(0)) : null;
    if (!bytes || bytes.length < 4 || bytes.length > 100 * 1024 * 1024 || bytes[0] !== 80 || bytes[1] !== 75 || bytes[2] !== 3 || bytes[3] !== 4) throw failure('invalid-file', '정상적인 공동 관리 엑셀 내용을 사용해 주세요.');
    return bytes;
  }
  function checkRevision(check) { if (typeof check === 'function' && check() !== true) throw failure('stale', '사진이나 작업 대상이 변경되어 이전 결과를 저장하지 않았습니다.'); }
  function changed() { return failure('shared-destination-changed', '관리자가 공동 저장 대상을 변경했습니다. 최신 공동 엑셀을 다시 읽은 뒤 사진을 확인해 주세요.'); }
  async function confirmDestination(expected) {
    const current = await authoritative();
    if (current.fingerprint !== expected.fingerprint) throw changed();
    if (!sameTarget(graph().inspect().target, expected.target)) throw changed();
    return current;
  }
  async function freshSnapshot(expected) {
    const current = expected || await authoritative(), previous = binding;
    const remote = graph().inspect();
    if (status.uncertain && previous && previous.identity !== current.identity) throw changed();
    if (!sameTarget(remote.target, current.target) || !remote.connected) {
      const validate = validator || (global.ToolWorkbook && (async file => global.ToolWorkbook.load(await file.arrayBuffer())));
      await graph().connect({ target: current.target, ...(validate ? { validate } : {}) });
    }
    const snapshot = await graph().read();
    if (!sameTarget(graph().inspect().target, current.target)) throw changed();
    binding = { ...current, target: cloneTarget(current.target) };
    lastRead = { version: snapshot.version, fingerprint: current.fingerprint, identity: current.identity };
    const remoteState = graph().inspect();
    Object.assign(status, { connected: true, name: snapshot.file.name, version: snapshot.version,
      needsPermission: false, needsReconnect: false, requiresSignIn: false, uncertain: !!remoteState.uncertain,
      destinationChanged: false, configurationVerified: true, persistent: remoteState.persistent, statusMessage: '' });
    return { ...snapshot, bindingIdentity: current.identity, bindingFingerprint: current.fingerprint,
      configVersion: current.configVersion, workspaceId: current.workspaceId,
      bindingChanged: !!(previous && previous.fingerprint !== current.fingerprint), target: cloneTarget(current.target) };
  }
  async function connect(options) {
    return serial(async () => { try {
      // Staff never select a local file or another sharing link. Only the
      // owner's authoritative cloud configuration selects the target.
      if (options && typeof options.validate === 'function') validator = options.validate;
      const snapshot = await freshSnapshot(); return complete(snapshot);
    } catch (error) { const result = mapped(error); rememberFailure(result); throw result; } });
  }
  async function read() { return serial(async () => { try { return await freshSnapshot(); } catch (error) { const result = mapped(error); rememberFailure(result); throw result; } }); }
  async function restore() { try { return await connect(); } catch (_) { return complete(); } }
  async function performSave(options) {
    const settings = options || {}, expected = binding && { ...binding, target: cloneTarget(binding.target) };
    if (!expected || !lastRead || settings.expectedVersion !== lastRead.version || lastRead.fingerprint !== expected.fingerprint) throw failure('shared-not-connected', '최신 공동 엑셀을 먼저 읽은 뒤 사진을 저장해 주세요.');
    if (inspect().uncertain) throw failure('requires-read', '이전 저장 완료 여부를 확인한 뒤 공동 원본을 다시 읽어 주세요.');
    checkRevision(settings.verifyCurrent); await confirmDestination(expected); checkRevision(settings.verifyCurrent);
    let bytes = copyBytes(settings.bytes), version = settings.expectedVersion, retries = 0;
    while (true) {
      try {
        const saved = await graph().save({ bytes, expectedVersion: version, verifyCurrent: settings.verifyCurrent,
          returnSnapshot: true, verifyDestination: async () => {
            checkRevision(settings.verifyCurrent); await confirmDestination(expected); checkRevision(settings.verifyCurrent); return true;
          } });
        binding = { ...expected, target: cloneTarget(expected.target) };
        lastRead = { version: saved.version, fingerprint: expected.fingerprint, identity: expected.identity };
        Object.assign(status, { connected: true, name: saved.name, version: saved.version, uncertain: false,
          needsPermission: false, needsReconnect: false, requiresSignIn: false, statusMessage: '',
          destinationChanged: false, configurationVerified: true,
          lastSavedAt: graph().inspect().lastSavedAt || new Date().toISOString() });
        // A different cloud item holds the owner's settings, so its update
        // cannot be atomic with the workbook PUT. Once Graph has verified the
        // actual workbook commit, keep that receipt even if settings changed
        // in flight or can no longer be read. Never retry that saved photo.
        try {
          const current = await authoritative();
          if (current.fingerprint !== expected.fingerprint) {
            Object.assign(status, { connected: false, needsReconnect: true, destinationChanged: true,
              statusMessage: '사진은 기존 공동 엑셀에 저장했습니다. 관리자가 변경한 공동 저장 대상을 다시 읽어 주세요.' });
          }
        } catch (error) {
          Object.assign(status, { connected: false, needsReconnect: true, configurationVerified: false });
          rememberFailure(mapped(error));
          status.statusMessage = '사진은 공동 엑셀에 저장했습니다. 현재 공동 연결 설정을 다시 확인해야 합니다. ' + status.statusMessage;
        }
        return { ...saved, bytes: copyBytes(saved.bytes || bytes), target: cloneTarget(expected.target),
          workspaceId: expected.workspaceId, bindingIdentity: expected.identity,
          bindingFingerprint: expected.fingerprint, configVersion: expected.configVersion, rebaseCount: retries,
          destinationChanged: status.destinationChanged, configurationVerified: status.configurationVerified,
          needsReconnect: status.needsReconnect, needsPermission: status.needsPermission, requiresSignIn: status.requiresSignIn };
      } catch (error) {
        const retryable = error && error.code === 'conflict' && (error.httpStatus === undefined || error.httpStatus === 412);
        if (!retryable || graph().inspect().uncertain || retries >= 2 || typeof settings.rebase !== 'function') throw error;
        checkRevision(settings.verifyCurrent); await confirmDestination(expected); checkRevision(settings.verifyCurrent);
        const snapshot = await freshSnapshot(expected);
        const rebased = await settings.rebase(snapshot);
        checkRevision(settings.verifyCurrent); await confirmDestination(expected); checkRevision(settings.verifyCurrent);
        bytes = copyBytes(rebased && rebased.bytes !== undefined ? rebased.bytes : rebased);
        version = snapshot.version; retries++;
      }
    }
  }
  async function save(options) { return serial(async () => { try { return await performSave(options); } catch (error) { const result = mapped(error); rememberFailure(result); throw result; } }); }
  async function requestAccess() { return read(); }
  async function takeBackup() { return graph().takeBackup(); }
  async function disconnect(options) { return serial(async () => { await graph().disconnect(options); binding = null; lastRead = null; validator = null; status = { ...initial(), busy: true }; return complete(); }); }
  global.ToolSharedFile = Object.freeze({ connect, read, save, restore, requestAccess, disconnect, takeBackup, inspect });
})(window);
