/* One owner-controlled workspace pointer in OneDrive, shared read-only with
 * the organization. Authentication remains with the official Graph endpoint.
 * Neither workbook contents nor access tokens are put in a public app URL.
 */
(function (global) {
  'use strict';
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const KEY = 'tool-manager-shared-workspace-v1', LIMIT = 16384;
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const SELECT = '?$select=id,name,size,eTag,file,folder,parentReference,@microsoft.graph.downloadUrl';
  let bootstrap = null, workspace = null, configVersion = '', persistent = true;
  let generation = 0, busy = false, lastError = '', queue = Promise.resolve(), provenOwner = '', provenActor = '';
  const clone = value => value == null ? null : JSON.parse(JSON.stringify(value));
  function failure(code, message, cause) { const error = new Error(message); error.code = code; if (cause) error.cause = cause; return error; }
  const validId = value => typeof value === 'string' && /^[A-Za-z0-9!._~-]{1,1024}$/.test(value);
  function fileReference(value, extension) {
    if (!value || !validId(value.driveId) || !validId(value.itemId) || typeof value.name !== 'string'
      || !value.name || value.name.length > 255 || /[\\/\u0000-\u001f\u007f]/.test(value.name)
      || !value.name.toLowerCase().endsWith(extension)) throw failure('workspace-target', '공동 저장 파일의 ID와 이름을 확인해 주세요.');
    return { driveId: value.driveId, itemId: value.itemId, name: value.name };
  }
  function shareURL(value) {
    let url; try { url = new URL(String(value || '').trim()); } catch (_) { throw failure('workspace-link', '정상적인 회사 HTTPS 공유 링크를 입력해 주세요.'); }
    if (url.protocol !== 'https:' || url.username || url.password || url.href.length > 12000) throw failure('workspace-link', '정상적인 회사 HTTPS 공유 링크를 입력해 주세요.');
    return url.href;
  }
  function normalizeBootstrap(value) {
    if (!value || value.v !== 1 || !GUID.test(String(value.clientId || '')) || !GUID.test(String(value.workspaceId || '')))
      throw failure('workspace-invite', '공동 작업 링크의 형식이 올바르지 않습니다.');
    const tenantId = String(value.tenantId || '').trim().toLowerCase();
    if (!GUID.test(tenantId) && tenantId !== 'organizations' && !/^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.onmicrosoft\.com$/.test(tenantId))
      throw failure('workspace-invite', '공동 작업 링크의 회사 연결 정보를 확인해 주세요.');
    const settings = fileReference(value.settings, '.json'); settings.shareUrl = shareURL(value.settings.shareUrl);
    return { v: 1, clientId: value.clientId.toLowerCase(), tenantId, workspaceId: value.workspaceId.toLowerCase(), settings };
  }
  function normalizeWorkspace(value, expected) {
    if (!value || value.v !== 1 || value.workspaceId !== expected.workspaceId || !validId(value.ownerObjectId)
      || Object.keys(value).some(key => !['v', 'workspaceId', 'ownerObjectId', 'workbook'].includes(key)))
      throw failure('workspace-settings', '공동 설정 파일이 선택한 작업 공간과 맞지 않습니다.');
    return { v: 1, workspaceId: value.workspaceId, ownerObjectId: value.ownerObjectId, workbook: fileReference(value.workbook, '.xlsx') };
  }
  function tag(value) { if (typeof value !== 'string' || !value || value === '*' || value.length > 4096 || /[\r\n]/.test(value)) throw failure('workspace-version', '공동 설정의 변경 버전을 확인하지 못했습니다.'); return value; }
  const encode = value => {
    let text = ''; for (const byte of new TextEncoder().encode(value)) text += String.fromCharCode(byte);
    return global.btoa(text).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  };
  const itemPath = value => '/drives/' + encodeURIComponent(value.driveId) + '/items/' + encodeURIComponent(value.itemId);
  function authInfo() {
    const info = global.ToolM365Auth?.inspect();
    if (!info?.configured) throw failure('auth-required', '담당자의 회사 앱 연결 설정과 회사 로그인이 필요합니다.');
    return info;
  }
  function sameAuth(a, b) { return a.clientId === b.clientId && a.tenantId === b.tenantId; }
  const actorIdentity = info => {
    const id = info.account?.objectId || info.account?.username;
    return id ? info.tenantId + '|' + String(id).toLowerCase() : '';
  };
  function assertCurrent(sequence, auth) {
    const current = authInfo();
    if (sequence !== generation || !sameAuth(auth, current) || auth.signedIn === true && current.signedIn !== true
      || actorIdentity(auth) !== actorIdentity(current)) throw failure('workspace-changed', '작업 중 회사 계정 또는 공동 작업 공간이 변경됐습니다. 다시 연결해 주세요.');
  }
  function httpError(status) {
    const values = {
      401: ['auth-required', '회사 로그인이 필요합니다. 다시 로그인해 주세요.'],
      403: ['workspace-denied', '현재 회사 계정으로 공동 설정에 접근하거나 수정할 수 없습니다.'],
      404: ['workspace-missing', '공동 설정 파일을 찾지 못했습니다. 담당자의 작업 링크를 확인해 주세요.'],
      409: ['workspace-conflict', '공동 설정 파일이 사용 중입니다. 최신 내용을 다시 읽어 주세요.'],
      412: ['workspace-conflict', '다른 작업에서 공동 설정을 변경했습니다. 덮어쓰지 않았습니다. 최신 내용을 다시 읽어 주세요.'],
      429: ['workspace-throttled', '회사 문서 요청이 잠시 제한됐습니다. 잠시 후 다시 시도해 주세요.']
    };
    const error = failure(...(values[status] || ['workspace-request', '회사 공동 설정 요청을 완료하지 못했습니다.'])); error.httpStatus = status; return error;
  }
  async function session() {
    const sequence = generation, before = authInfo();
    if (bootstrap && !sameAuth(before, bootstrap)) throw failure('workspace-account', '공동 작업 링크에 지정된 회사 앱으로 로그인해 주세요.');
    const token = await global.ToolM365Auth.getAccessToken({ minimumValidityMs: 60000 });
    if (typeof token !== 'string' || !token || /[\r\n]/.test(token)) throw failure('auth-required', '정상적인 회사 로그인 연결이 필요합니다.');
    const auth = authInfo();
    if (sequence !== generation || !sameAuth(before, auth) || before.signedIn === true && actorIdentity(before) !== actorIdentity(auth))
      throw failure('workspace-changed', '회사 앱 또는 공동 작업 연결이 변경됐습니다. 다시 시도해 주세요.');
    return { auth, token, sequence };
  }
  async function graph(path, context, options) {
    if (!path.startsWith('/') || path.startsWith('//') || /[\r\n]/.test(path)) throw failure('workspace-request', '공동 설정 요청 경로가 올바르지 않습니다.');
    assertCurrent(context.sequence, context.auth);
    const controller = new AbortController(), timer = global.setTimeout(() => controller.abort(), 30000), settings = options || {};
    try {
      const response = await global.fetch(GRAPH + path, { method: settings.method || 'GET', cache: 'no-store', credentials: 'omit', redirect: 'error', signal: controller.signal,
        headers: { Accept: 'application/json', Authorization: 'Bearer ' + context.token, ...(settings.headers || {}) }, body: settings.body });
      if (!response.ok) throw httpError(response.status);
      assertCurrent(context.sequence, context.auth); return response;
    } finally { global.clearTimeout(timer); }
  }
  async function boundedBytes(response) {
    if (Number(response.headers.get('content-length')) > LIMIT) throw failure('workspace-size', '공동 설정 파일이 너무 큽니다.');
    if (!response.body?.getReader) { const bytes = new Uint8Array(await response.arrayBuffer()); if (bytes.length > LIMIT) throw failure('workspace-size', '공동 설정 파일이 너무 큽니다.'); return bytes; }
    const reader = response.body.getReader(), pieces = []; let length = 0;
    while (true) { const chunk = await reader.read(); if (chunk.done) break; length += chunk.value.length; if (length > LIMIT) { await reader.cancel(); throw failure('workspace-size', '공동 설정 파일이 너무 큽니다.'); } pieces.push(chunk.value); }
    const bytes = new Uint8Array(length); let offset = 0; for (const piece of pieces) { bytes.set(piece, offset); offset += piece.length; } return bytes;
  }
  async function settingsMetadata(reference, context, redeem) {
    let value;
    try { value = await (await graph(itemPath(reference) + SELECT, context)).json(); }
    catch (error) {
      if (!redeem || ![403, 404].includes(error.httpStatus)) throw error;
      const resolved = await (await graph('/shares/u!' + encode(reference.shareUrl) + '/driveItem' + SELECT, context, { headers: { Prefer: 'redeemSharingLinkIfNecessary' } })).json();
      const actual = resolved.remoteItem || resolved;
      if (actual.id !== reference.itemId || actual.parentReference?.driveId !== reference.driveId) throw failure('workspace-settings', '공유 링크가 다른 공동 설정 파일을 가리킵니다.');
      value = await (await graph(itemPath(reference) + SELECT, context)).json();
    }
    if (value.id !== reference.itemId || value.parentReference?.driveId !== reference.driveId || !value.file || value.folder || !/\.json$/i.test(value.name || '')
      || !Number.isSafeInteger(value.size) || value.size < 2 || value.size > LIMIT) throw failure('workspace-settings', '공동 설정 파일의 ID·형식·크기를 확인해 주세요.');
    tag(value.eTag); return value;
  }
  async function readSettings(pointer, context) {
    const before = await settingsMetadata(pointer.settings, context, true);
    let url; try { url = new URL(before['@microsoft.graph.downloadUrl']); } catch (_) { throw failure('workspace-settings', '공동 설정 다운로드 주소를 확인하지 못했습니다.'); }
    if (url.protocol !== 'https:' || url.username || url.password) throw failure('workspace-settings', '공동 설정 다운로드 주소가 올바르지 않습니다.');
    const controller = new AbortController(), timer = global.setTimeout(() => controller.abort(), 30000);
    let bytes;
    try {
      // The short-lived Graph download URL is already authorized. Never send
      // our Graph bearer, browser cookies or a cached response to this host.
      const response = await global.fetch(url.href, { credentials: 'omit', cache: 'no-store', redirect: 'follow', signal: controller.signal });
      if (!response.ok) throw failure('workspace-download', '공동 설정을 내려받지 못했습니다. 다시 로그인해 주세요.');
      bytes = await boundedBytes(response);
    } finally { global.clearTimeout(timer); }
    assertCurrent(context.sequence, context.auth);
    const after = await settingsMetadata(pointer.settings, context, false);
    if (before.eTag !== after.eTag || before.size !== after.size || bytes.length !== after.size) throw failure('workspace-conflict', '공동 설정을 읽는 동안 담당자가 변경했습니다. 다시 읽어 주세요.');
    let value; try { value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes)); } catch (_) { throw failure('workspace-settings', '공동 설정 JSON 형식이 올바르지 않습니다.'); }
    return { value: normalizeWorkspace(value, pointer), version: after.eTag };
  }
  async function resolveWorkbook(raw, context) {
    const url = shareURL(raw), resolved = await (await graph('/shares/u!' + encode(url) + '/driveItem' + SELECT, context, { headers: { Prefer: 'redeemSharingLinkIfNecessary' } })).json();
    const item = resolved.remoteItem || resolved, reference = fileReference({ driveId: item.parentReference?.driveId, itemId: item.id, name: item.name }, '.xlsx');
    const fresh = await (await graph(itemPath(reference) + SELECT, context)).json();
    if (fresh.id !== reference.itemId || fresh.parentReference?.driveId !== reference.driveId || !fresh.file || fresh.folder || !/\.xlsx$/i.test(fresh.name || '')
      || !Number.isSafeInteger(fresh.size) || fresh.size < 4 || fresh.size > 100 * 1024 * 1024) throw failure('workspace-target', '정상적인 100 MB 이하의 회사 .xlsx 파일을 선택해 주세요.');
    return fileReference({ ...reference, name: fresh.name }, '.xlsx');
  }
  async function ownDrive(context) {
    const drive = await (await graph('/me/drive?$select=id,owner,driveType', context)).json();
    const ownerObjectId = drive.owner?.user?.id || context.auth.account?.objectId;
    if (!validId(drive.id) || !validId(ownerObjectId)) throw failure('workspace-owner', '담당자 본인의 OneDrive와 계정 ID를 확인하지 못했습니다.');
    return { driveId: drive.id, ownerObjectId };
  }
  async function ownerOnly(reference, owner, context) {
    const response = await (await graph(itemPath(reference) + '/permissions', context)).json();
    if (!Array.isArray(response.value) || response['@odata.nextLink']) throw failure('workspace-owner', '공동 설정의 전체 읽기·수정 권한을 확인하지 못했습니다.');
    for (const permission of response.value) {
      const roles = permission.roles || [];
      if (!roles.some(role => ['write', 'owner'].includes(role))) continue;
      const identities = permission.grantedToIdentitiesV2 || permission.grantedToIdentities || [permission.grantedToV2 || permission.grantedTo];
      if (!identities.length || identities.some(identity => identity?.user?.id !== owner)) throw failure('workspace-owner', '공동 설정에 담당자 이외의 수정 권한이 있습니다. 읽기 전용 공유로 설정해 주세요.');
    }
  }
  function setCurrent(result) { workspace = result.value; configVersion = result.version; lastError = ''; return { ...clone(workspace), configVersion }; }
  function inspect() {
    const auth = global.ToolM365Auth?.inspect() || {}, account = auth.account;
    return { configured: !!bootstrap, persistent, busy, bootstrap: clone(bootstrap), workspace: clone(workspace), workbook: clone(workspace?.workbook), configVersion,
      isOwner: !!(workspace && auth.signedIn !== false && (account?.objectId === workspace.ownerObjectId || provenOwner === workspace.ownerObjectId && provenActor && provenActor === actorIdentity(auth))), lastError };
  }
  function configure(value) {
    const normalized = normalizeBootstrap(value), encoded = JSON.stringify(normalized);
    try { global.localStorage.setItem(KEY, encoded); if (global.localStorage.getItem(KEY) !== encoded) throw new Error('storage'); }
    catch (_) { persistent = false; throw failure('workspace-storage', '공동 작업 링크를 이 기기에 기억하지 못했습니다. 브라우저 저장소를 확인해 주세요.'); }
    if (JSON.stringify(bootstrap) !== encoded) { generation++; workspace = null; configVersion = ''; provenOwner = ''; provenActor = ''; }
    bootstrap = normalized; persistent = true; lastError = ''; return inspect();
  }
  function clear() { try { global.localStorage.removeItem(KEY); } catch (_) { throw failure('workspace-storage', '기기의 공동 작업 연결을 지우지 못했습니다.'); } bootstrap = null; workspace = null; configVersion = ''; provenOwner = ''; provenActor = ''; generation++; lastError = ''; return inspect(); }
  function serial(action) {
    const promise = queue.then(async () => { busy = true; try { return await action(); } catch (error) { workspace = null; configVersion = ''; lastError = error.message; throw error; } finally { busy = false; } }); queue = promise.catch(() => {}); return promise;
  }
  function refresh() { return serial(async () => { if (!bootstrap) throw failure('workspace-required', '담당자의 공동 작업 링크로 열어 주세요.'); const pointer = clone(bootstrap), context = await session(); const result = await readSettings(pointer, context); assertCurrent(context.sequence, context.auth); return setCurrent(result); }); }
  function create(options) {
    return serial(async () => {
      const context = await session(), workbook = await resolveWorkbook(options?.workbookShareUrl, context), owner = await ownDrive(context), workspaceId = global.crypto.randomUUID();
      const settings = { v: 1, workspaceId, ownerObjectId: owner.ownerObjectId, workbook }, filename = 'ToolManagerShared-' + workspaceId + '.json', body = JSON.stringify(settings);
      const made = await (await graph('/me/drive/root:/' + encodeURIComponent(filename) + ':/content', context, { method: 'PUT', headers: { 'Content-Type': 'application/json', 'If-None-Match': '*' }, body })).json();
      const reference = fileReference({ driveId: made.parentReference?.driveId, itemId: made.id, name: made.name }, '.json');
      if (reference.driveId !== owner.driveId) throw failure('workspace-owner', '담당자 본인의 저장소에 공동 설정을 만들지 못했습니다.');
      const permission = await (await graph(itemPath(reference) + '/createLink', context, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ type: 'view', scope: 'organization' }) })).json();
      if (permission.link?.scope !== 'organization' || permission.link?.type !== 'view' || !Array.isArray(permission.roles) || !permission.roles.includes('read') || permission.roles.some(role => role !== 'read'))
        throw failure('workspace-sharing', '회사 내부의 읽기 전용 공동 설정 링크를 만들지 못했습니다.');
      reference.shareUrl = shareURL(permission.link.webUrl);
      await ownerOnly(reference, owner.ownerObjectId, context);
      const next = normalizeBootstrap({ v: 1, clientId: context.auth.clientId, tenantId: context.auth.tenantId, workspaceId, settings: reference });
      const verified = await readSettings(next, context);
      if (JSON.stringify(verified.value) !== JSON.stringify(settings)) throw failure('workspace-verification', '공동 설정의 저장 내용을 확인하지 못했습니다.');
      assertCurrent(context.sequence, context.auth); configure(next); provenOwner = owner.ownerObjectId; provenActor = actorIdentity(context.auth); setCurrent(verified); return clone(next);
    });
  }
  function changeWorkbook(options) {
    return serial(async () => {
      if (!bootstrap) throw failure('workspace-required', '공동 작업 공간을 먼저 연결해 주세요.');
      const pointer = clone(bootstrap), expectedVersion = options?.expectedVersion || configVersion, context = await session();
      const current = await readSettings(pointer, context), owner = await ownDrive(context);
      assertCurrent(context.sequence, context.auth);
      if (owner.driveId !== pointer.settings.driveId || owner.ownerObjectId !== current.value.ownerObjectId) throw failure('workspace-owner', '저장할 엑셀은 이 공동 작업 공간을 만든 담당자만 변경할 수 있습니다.');
      if (expectedVersion && current.version !== expectedVersion) throw failure('workspace-conflict', '다른 작업에서 공동 설정을 변경했습니다. 최신 설정을 다시 읽어 주세요.');
      const workbook = await resolveWorkbook(options?.workbookShareUrl, context);
      if (JSON.stringify(workbook) === JSON.stringify(current.value.workbook)) { provenOwner = owner.ownerObjectId; provenActor = actorIdentity(context.auth); return setCurrent(current); }
      await ownerOnly(pointer.settings, owner.ownerObjectId, context);
      const value = { ...current.value, workbook }, response = await graph(itemPath(pointer.settings) + '/content', context, { method: 'PUT', headers: { 'Content-Type': 'application/json', 'If-Match': current.version }, body: JSON.stringify(value) });
      const written = await response.json();
      if (written.id !== pointer.settings.itemId) throw failure('workspace-verification', '공동 설정 변경 결과를 확인하지 못했습니다.');
      const actual = await readSettings(pointer, context);
      assertCurrent(context.sequence, context.auth);
      if (actual.version !== tag(written.eTag) || JSON.stringify(actual.value) !== JSON.stringify(value)) throw failure('workspace-verification', '공동 설정 저장 후 내용이 예상과 다릅니다. 다시 읽어 주세요.');
      provenOwner = owner.ownerObjectId; provenActor = actorIdentity(context.auth); return setCurrent(actual);
    });
  }
  function captureInvitation() {
    try {
      const values = new URLSearchParams(global.location.hash.slice(1)).getAll('workspace');
      if (!values.length) return { value: null, error: null };
      if (values.length !== 1 || values[0].length > 24576 || !/^[A-Za-z0-9_-]+$/.test(values[0])) throw new Error('encoding');
      const text = values[0].replace(/-/g, '+').replace(/_/g, '/'), bytes = Uint8Array.from(global.atob(text + '='.repeat((4 - text.length % 4) % 4)), character => character.charCodeAt(0));
      return { value: normalizeBootstrap(JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes))), error: null };
    } catch (_) { return { value: null, error: failure('workspace-invite', '공동 작업 링크를 읽지 못했습니다. 담당자의 새 링크를 받아 주세요.') }; }
  }
  let invitation = captureInvitation(), invitationHash = global.location.hash, consumedInvitation = false;
  function takeInvitation() {
    if (global.location.hash !== invitationHash) { invitationHash = global.location.hash; invitation = captureInvitation(); consumedInvitation = false; }
    if (consumedInvitation) return null; consumedInvitation = true;
    if (invitation.error) throw invitation.error; return clone(invitation.value);
  }
  function invitationLink(baseURL) {
    if (!bootstrap) throw failure('workspace-required', '공동 작업 공간을 먼저 만드세요.');
    const url = new URL(baseURL || global.location.href, global.location.href), current = new URL(global.location.href);
    const local = url.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(url.hostname);
    if (url.origin !== current.origin || !(url.protocol === 'https:' || local) || url.username || url.password) throw failure('workspace-link', '현재 프로그램과 같은 HTTPS 주소를 사용해 주세요.');
    url.search = ''; url.hash = 'workspace=' + encode(JSON.stringify(bootstrap)); return url.href;
  }
  try { const saved = global.localStorage.getItem(KEY); if (saved) bootstrap = normalizeBootstrap(JSON.parse(saved)); } catch (_) { bootstrap = null; persistent = false; }
  global.ToolSharedWorkspace = Object.freeze({ configure, clear, inspect, refresh, create, changeWorkbook, takeInvitation, invitationLink });
})(window);
