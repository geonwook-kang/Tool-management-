/* A manager can share public app IDs and one workbook target with the team.
 * The invitation is kept in a URL fragment, removed before other app code runs,
 * and never grants Microsoft 365 permissions or contains an access token.
 */
(function (global) {
  'use strict';
  const KEY = 'tool-manager-team-config-v1';
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const LIMIT = 18000;
  let persistent = true, configuration = null;

  function failure(code, message) { const error = new Error(message); error.code = code; return error; }
  function supported() {
    const location = global.location;
    return global.isSecureContext !== false && (location.protocol === 'https:' || location.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(location.hostname));
  }
  function validate(value) {
    if (!value || typeof value !== 'object' || Array.isArray(value) || value.v !== undefined && value.v !== 1) throw failure('team-config', '공용 엑셀 연결 정보의 형식이 맞지 않습니다.');
    const clientId = String(value.clientId || '').trim().toLowerCase();
    const tenantId = String(value.tenantId || '').trim().toLowerCase();
    if (!GUID.test(clientId) || !GUID.test(tenantId) && tenantId !== 'organizations' && !/^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.onmicrosoft\.com$/.test(tenantId)) throw failure('team-config', '관리자가 회사 앱 ID와 테넌트 ID를 먼저 설정해야 합니다.');
    const name = String(value.name || '').trim();
    if (name.length > 120 || /[\u0000-\u001f\u007f]/.test(name)) throw failure('team-config', '공용 엑셀 이름은 120자 이내로 입력해 주세요.');
    const config = { v: 1, clientId, tenantId, name };
    if (value.target != null) {
      if (value.shareUrl) throw failure('team-config', '공용 엑셀 대상은 파일 ID 또는 공유 링크 한 가지로만 지정해 주세요.');
      const target = value.target;
      if (!target || typeof target !== 'object' || Array.isArray(target)) throw failure('team-config', '공용 엑셀의 파일 ID가 올바르지 않습니다.');
      const driveId = String(target.driveId || ''), itemId = String(target.itemId || ''), targetName = String(target.name || '');
      if (![driveId, itemId].every(id => id.length > 0 && id.length <= 200 && /^[A-Za-z0-9!._~-]+$/.test(id)) || !targetName || targetName.length > 255 || /[\u0000-\u001f\u007f/\\]/.test(targetName) || !/\.xlsx$/i.test(targetName)) throw failure('team-config', '공용 엑셀의 드라이브·파일 ID와 .xlsx 파일 이름을 확인해 주세요.');
      config.target = { driveId, itemId, name: targetName };
    } else {
      let shareUrl;
      try { shareUrl = new URL(String(value.shareUrl || '').trim()); } catch (_) { throw failure('team-config', '공용 엑셀의 HTTPS 공유 링크가 필요합니다.'); }
      if (shareUrl.protocol !== 'https:' || shareUrl.username || shareUrl.password || shareUrl.href.length > 12000) throw failure('team-config', '공용 엑셀의 정상적인 HTTPS 공유 링크가 필요합니다.');
      config.shareUrl = shareUrl.href;
    }
    return config;
  }
  function clone(value) { return value ? JSON.parse(JSON.stringify(value)) : null; }
  function captureInvitation() {
    const url = new URL(global.location.href), parts = url.hash.slice(1).split('&');
    const invitations = [], remaining = [];
    for (const part of parts) {
      const parsed = new URLSearchParams(part);
      if (parsed.has('team')) invitations.push(...parsed.getAll('team'));
      else remaining.push(part);
    }
    if (!invitations.length) return { value: null, error: null };
    // Keep unrelated fragment bytes unchanged. Do this even for rejected links.
    url.hash = remaining.join('&');
    try { global.history.replaceState(null, '', url.pathname + url.search + url.hash); }
    catch (_) { return { value: null, error: failure('team-link', '공용 연결 링크를 안전하게 열지 못했습니다. 일반 브라우저에서 다시 열어 주세요.') }; }
    try {
      if (!supported() || invitations.length !== 1 || invitations[0].length > LIMIT || !/^[A-Za-z0-9_-]+$/.test(invitations[0])) throw new Error('encoding');
      const encoded = invitations[0].replace(/-/g, '+').replace(/_/g, '/');
      const bytes = Uint8Array.from(global.atob(encoded + '='.repeat((4 - encoded.length % 4) % 4)), c => c.charCodeAt(0));
      const data = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
      if (data.v !== 1) throw new Error('version');
      return { value: validate(data), error: null };
    } catch (_) { return { value: null, error: failure('team-link', '공용 엑셀 연결 링크가 잘못되었거나 너무 깁니다. 관리자에게 새 연결 링크를 받아 주세요.') }; }
  }
  // Capture before DOMContentLoaded, service-worker setup or app navigation.
  let invitation = captureInvitation();
  try { const value = global.localStorage.getItem(KEY); configuration = value ? validate(JSON.parse(value)) : null; }
  catch (_) { persistent = false; configuration = null; }

  function takeInvitation() {
    // Opening another invite may navigate only the fragment of an already open
    // PWA. The app owns hashchange handling; consuming it here also scrubs that
    // fresh fragment before the app can apply or reject the new destination.
    if (!invitation.value && !invitation.error) invitation = captureInvitation();
    const result = invitation; invitation = { value: null, error: null };
    if (result.error) throw result.error;
    return clone(result.value);
  }
  function create(value) {
    if (!supported()) throw failure('team-unsupported', '공용 연결 링크는 HTTPS 앱에서 만들 수 있습니다.');
    const config = validate(value), bytes = new TextEncoder().encode(JSON.stringify(config));
    let binary = ''; for (const byte of bytes) binary += String.fromCharCode(byte);
    const encoded = global.btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    if (encoded.length > LIMIT) throw failure('team-config', '공용 엑셀 공유 링크가 너무 깁니다. 짧은 공유 링크를 다시 복사해 주세요.');
    const url = new URL(global.location.href); url.search = ''; url.hash = 'team=' + encoded;
    return url.href;
  }
  function remember(value) {
    if (!supported()) throw failure('team-unsupported', 'HTTPS 앱에서 공용 엑셀 연결을 저장해 주세요.');
    const config = validate(value), encoded = JSON.stringify(config);
    try {
      global.localStorage.setItem(KEY, encoded);
      if (global.localStorage.getItem(KEY) !== encoded) throw new Error('storage');
    } catch (_) { persistent = false; throw failure('team-storage', '공용 엑셀 연결을 이 기기에 저장하지 못했습니다. 일반 브라우저의 저장소 설정을 확인해 주세요.'); }
    configuration = config; persistent = true; return clone(config);
  }
  function clear() {
    try { global.localStorage.removeItem(KEY); }
    catch (_) { persistent = false; throw failure('team-storage', '기기에 저장한 공용 연결 정보를 지우지 못했습니다.'); }
    configuration = null; return inspect();
  }
  function exportSettings(value) {
    const config = validate(value);
    if (!config.target) throw failure('team-config', '먼저 회사 공유 엑셀을 연결해 파일 ID를 확인한 뒤 공용 설정을 내보내 주세요.');
    return JSON.stringify(Object.assign({ enabled: true }, config), null, 2) + '\n';
  }
  async function loadPublished() {
    if (!supported() || global.navigator && global.navigator.onLine === false) return null;
    const controller = new AbortController(), timer = global.setTimeout(() => controller.abort(), 5000);
    let readingNetwork = true;
    try {
      const url = new URL('./team-settings.json', global.location.href);
      const response = await global.fetch(url.href, { cache: 'no-store', credentials: 'omit', redirect: 'error', signal: controller.signal, headers: { Accept: 'application/json' } });
      if (response.status === 404) return null;
      if (!response.ok) throw failure('team-published', '공용 엑셀 설정을 불러오지 못했습니다. 네트워크 연결을 확인해 주세요.');
      const maximum = 16 * 1024;
      if (Number(response.headers.get('content-length')) > maximum) throw failure('team-published', '공용 엑셀 설정 파일이 너무 큽니다.');
      let bytes;
      if (response.body && typeof response.body.getReader === 'function') {
        const reader = response.body.getReader(), chunks = []; let size = 0;
        while (true) {
          const chunk = await reader.read(); if (chunk.done) break;
          size += chunk.value.byteLength;
          if (size > maximum) { await reader.cancel(); throw failure('team-published', '공용 엑셀 설정 파일이 너무 큽니다.'); }
          chunks.push(chunk.value);
        }
        bytes = new Uint8Array(size); let offset = 0;
        for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
      } else {
        bytes = new Uint8Array(await response.arrayBuffer());
        if (bytes.length > maximum) throw failure('team-published', '공용 엑셀 설정 파일이 너무 큽니다.');
      }
      readingNetwork = false;
      const value = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
      if (value && value.enabled === false) return null;
      if (!value || value.enabled !== true || value.shareUrl || !value.target) throw failure('team-published', '관리자가 게시한 공용 엑셀 설정의 형식이 맞지 않습니다.');
      return validate(value);
    } catch (error) {
      if (error && typeof error.code === 'string') throw error;
      if (readingNetwork && error && ['TypeError', 'AbortError'].includes(error.name)) throw failure('team-unavailable', '공용 연결 설정 서버에 연결할 수 없습니다. 이 기기에 기억한 연결을 사용합니다.');
      throw failure('team-published', '공용 엑셀 설정을 읽지 못했습니다. 관리자에게 연결 설정을 확인해 주세요.');
    } finally { global.clearTimeout(timer); }
  }
  function inspect() { return { configured: !!configuration, configuration: clone(configuration), persistent, supported: supported() }; }
  global.ToolTeamConfig = Object.freeze({ create, takeInvitation, remember, inspect, clear, loadPublished, exportSettings });
})(window);
