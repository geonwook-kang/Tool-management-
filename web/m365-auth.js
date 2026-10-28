/* Microsoft 365 public-client authentication. Tokens never enter persistent storage. */
(function (global) {
  'use strict';
  const CONFIG_KEY = 'tool-manager-m365-config-v1';
  const PENDING_KEY = 'tool-manager-m365-pkce-v1';
  const PENDING_TTL = 10 * 60 * 1000;
  const AUTHORITY = 'https://login.microsoftonline.com';
  const GRAPH_SCOPE = 'https://graph.microsoft.com/Files.ReadWrite.All';
  const SCOPES = ['openid', 'profile', GRAPH_SCOPE];
  const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const CALLBACK_KEYS = ['code', 'state', 'error', 'error_description', 'error_uri', 'session_state', 'client_info'];
  let configuration = null, token = null, initialization = null, generation = 0, lastError = '';

  function failure(code, message) { const error = new Error(message); error.code = code; return error; }
  function supported() {
    const location = global.location;
    const local = location.protocol === 'http:' && ['localhost', '127.0.0.1', '[::1]'].includes(location.hostname);
    return (location.protocol === 'https:' || local) && global.isSecureContext !== false && !!(global.crypto && global.crypto.subtle && global.crypto.getRandomValues);
  }
  function currentRedirect() {
    const url = new URL(global.location.href); url.search = ''; url.hash = ''; return url.href;
  }
  function validateConfiguration(value) {
    if (!value || typeof value !== 'object' || !GUID.test(String(value.clientId || '').trim())) throw failure('auth-config', '회사 앱의 애플리케이션(클라이언트) ID를 확인해 주세요.');
    const clientId = String(value.clientId).trim().toLowerCase(), tenantId = String(value.tenantId || 'organizations').trim().toLowerCase();
    if (!GUID.test(tenantId) && tenantId !== 'organizations' && !/^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.onmicrosoft\.com$/.test(tenantId)) throw failure('auth-config', '테넌트 ID 또는 회사의 onmicrosoft.com 테넌트 이름을 확인해 주세요.');
    const redirectUri = String(value.redirectUri || currentRedirect());
    let redirect;
    try { redirect = new URL(redirectUri); } catch (_) { throw failure('auth-config', '로그인 회신 주소를 확인해 주세요.'); }
    if (redirect.href !== currentRedirect() || redirect.username || redirect.password || redirect.search || redirect.hash) throw failure('auth-redirect', '로그인 회신 주소는 현재 앱의 주소와 정확히 같아야 합니다.');
    return { clientId, tenantId, redirectUri: redirect.href };
  }
  function loadConfiguration() {
    try { const saved = global.localStorage.getItem(CONFIG_KEY); return saved ? validateConfiguration(JSON.parse(saved)) : null; }
    catch (_) { return null; }
  }
  function clearPending() { try { global.sessionStorage.removeItem(PENDING_KEY); } catch (_) {} }
  function inspect() {
    const signedIn = !!(token && token.expiresAt > Date.now());
    return {
      configured: !!configuration, supported: supported(), signedIn, requiresSignIn: !signedIn,
      clientId: configuration ? configuration.clientId : '', tenantId: configuration ? configuration.tenantId : '',
      redirectUri: configuration ? configuration.redirectUri : currentRedirect(),
      expiresAt: token ? token.expiresAt : 0, account: signedIn ? Object.assign({}, token.account) : null, lastError
    };
  }
  function configure(value) {
    if (!supported()) throw failure('auth-unsupported', 'Microsoft 365 로그인은 HTTPS 앱 또는 localhost에서 사용할 수 있습니다.');
    const next = validateConfiguration(value);
    try { global.localStorage.setItem(CONFIG_KEY, JSON.stringify(next)); }
    catch (_) { throw failure('auth-storage', '회사 앱 연결 설정을 이 브라우저에 저장하지 못했습니다. 일반 창의 저장소 설정을 확인해 주세요.'); }
    if (!configuration || JSON.stringify(configuration) !== JSON.stringify(next)) { generation++; token = null; initialization = null; clearPending(); }
    configuration = next; lastError = ''; return inspect();
  }
  function base64url(bytes) {
    let binary = ''; for (const value of bytes) binary += String.fromCharCode(value);
    return global.btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function random() { return base64url(global.crypto.getRandomValues(new Uint8Array(32))); }
  async function signIn() {
    if (!supported()) throw failure('auth-unsupported', 'Microsoft 365 로그인은 HTTPS 앱에서 열어 주세요.');
    if (!configuration) throw failure('auth-config', '회사 IT에서 받은 앱 연결 정보를 먼저 저장해 주세요.');
    const verifier = random(), state = random(), nonce = random();
    const challenge = base64url(new Uint8Array(await global.crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier))));
    const pending = { verifier, state, nonce, createdAt: Date.now(), configuration: Object.assign({}, configuration) };
    try {
      global.sessionStorage.setItem(PENDING_KEY, JSON.stringify(pending));
      if (global.sessionStorage.getItem(PENDING_KEY) !== JSON.stringify(pending)) throw new Error('storage');
    } catch (_) { clearPending(); throw failure('auth-storage', '로그인 확인 정보를 임시 저장하지 못했습니다. 일반 브라우저 창에서 다시 시도해 주세요.'); }
    const url = new URL(AUTHORITY + '/' + encodeURIComponent(configuration.tenantId) + '/oauth2/v2.0/authorize');
    url.search = new URLSearchParams({
      client_id: configuration.clientId, response_type: 'code', response_mode: 'query',
      redirect_uri: configuration.redirectUri, scope: SCOPES.join(' '), state, nonce,
      code_challenge: challenge, code_challenge_method: 'S256', prompt: 'select_account'
    }).toString();
    lastError = ''; global.location.assign(url.href); return { redirecting: true };
  }

  // Scrub authorization codes before app initialization/service-worker setup.
  // The response stays only in this closure until initialize consumes it.
  function captureCallback() {
    const url = new URL(global.location.href), search = url.searchParams, hash = new URLSearchParams(url.hash.replace(/^#/, ''));
    const hasSearch = ['code', 'error', 'access_token', 'id_token'].some((key) => search.has(key));
    const hasHash = ['code', 'error', 'access_token', 'id_token'].some((key) => hash.has(key));
    if (!hasSearch && !hasHash) return null;
    const response = { values: {}, invalid: false };
    for (const key of CALLBACK_KEYS) {
      const values = search.getAll(key); if (values.length > 1) response.invalid = true;
      if (values.length === 1) response.values[key] = values[0];
    }
    if (hasHash || search.has('access_token') || search.has('id_token') || search.has('code') && search.has('error')) response.invalid = true;
    for (const key of CALLBACK_KEYS.concat(['access_token', 'id_token', 'token_type', 'expires_in'])) url.searchParams.delete(key);
    if (hasHash) url.hash = '';
    try { global.history.replaceState(null, '', url.pathname + url.search + url.hash); }
    catch (_) { response.invalid = true; }
    return response;
  }
  let callback = captureCallback();
  configuration = loadConfiguration();

  function identityClaims(value, pending) {
    // Claims are accepted only from the trusted HTTPS token endpoint response,
    // never from a caller-provided token or an implicit-flow URL fragment.
    try {
      const segments = value.split('.'); if (segments.length !== 3 || !segments[2]) throw new Error('jwt');
      const encoded = segments[1].replace(/-/g, '+').replace(/_/g, '/');
      const bytes = Uint8Array.from(global.atob(encoded + '='.repeat((4 - encoded.length % 4) % 4)), (c) => c.charCodeAt(0));
      const claims = JSON.parse(new TextDecoder().decode(bytes)), now = Date.now() / 1000;
      if (claims.aud !== pending.configuration.clientId || claims.nonce !== pending.nonce || !GUID.test(String(claims.tid || '')) || claims.iss !== AUTHORITY + '/' + String(claims.tid).toLowerCase() + '/v2.0' || !Number.isFinite(claims.exp) || claims.exp <= now || Number.isFinite(claims.nbf) && claims.nbf > now + 120) throw new Error('claims');
      if (GUID.test(pending.configuration.tenantId) && pending.configuration.tenantId !== String(claims.tid).toLowerCase()) throw new Error('tenant');
      return { name: String(claims.name || '').slice(0, 200), username: String(claims.preferred_username || '').slice(0, 300), tenantId: String(claims.tid).toLowerCase() };
    } catch (_) { throw failure('auth-response', '회사 로그인 응답의 계정 확인 정보가 맞지 않습니다. 다시 로그인해 주세요.'); }
  }
  async function initialize() {
    if (initialization) return initialization;
    initialization = (async () => {
      if (!callback) return inspect();
      const response = callback; callback = null;
      let pending;
      try { pending = JSON.parse(global.sessionStorage.getItem(PENDING_KEY) || 'null'); }
      catch (_) { pending = null; }
      clearPending();
      const values = response.values;
      if (!supported()) throw failure('auth-unsupported', 'Microsoft 365 로그인은 HTTPS 앱에서 열어 주세요.');
      if (response.invalid || !pending || typeof pending.state !== 'string' || !values.state || values.state !== pending.state || typeof pending.verifier !== 'string' || !/^[A-Za-z0-9_-]{43}$/.test(pending.verifier) || typeof pending.nonce !== 'string' || !Number.isFinite(pending.createdAt) || Date.now() - pending.createdAt > PENDING_TTL || pending.createdAt > Date.now() + 30000) throw failure('auth-state', '로그인 확인 정보가 만료되었거나 맞지 않습니다. 앱에서 다시 로그인을 시작해 주세요.');
      const expected = validateConfiguration(pending.configuration);
      if (!configuration || JSON.stringify(configuration) !== JSON.stringify(expected)) throw failure('auth-state', '로그인 도중 회사 앱 설정이 변경되었습니다. 다시 로그인해 주세요.');
      if (values.error) throw failure(values.error === 'access_denied' ? 'auth-denied' : 'auth-response', values.error === 'access_denied' ? '회사 계정의 접근 권한이 허용되지 않았습니다. 회사 IT의 앱 권한 승인을 확인해 주세요.' : 'Microsoft 365 로그인을 완료하지 못했습니다. 회사 앱 설정과 로그인 권한을 확인해 주세요.');
      if (!values.code || values.code.length > 8192) throw failure('auth-response', '로그인 응답에 인증 코드가 없습니다. 다시 로그인해 주세요.');
      const ownGeneration = generation, controller = new AbortController();
      const timeout = global.setTimeout(() => controller.abort(), 30000);
      let responseData;
      try {
        const result = await global.fetch(AUTHORITY + '/' + encodeURIComponent(expected.tenantId) + '/oauth2/v2.0/token', {
          method: 'POST', mode: 'cors', credentials: 'omit', cache: 'no-store', redirect: 'error', signal: controller.signal,
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
          body: new URLSearchParams({ client_id: expected.clientId, grant_type: 'authorization_code', code: values.code,
            redirect_uri: expected.redirectUri, code_verifier: pending.verifier, scope: SCOPES.join(' ') }).toString()
        });
        if (!result.ok) throw failure(result.status === 400 ? 'auth-consent' : 'auth-response', '회사 로그인 토큰을 받지 못했습니다. 앱의 SPA 회신 주소와 Microsoft Graph 위임 권한·회사 승인을 확인해 주세요.');
        responseData = await result.json();
      } catch (error) {
        if (error && error.code) throw error;
        throw failure('auth-network', '회사 로그인 서버에 연결하지 못했습니다. 인터넷 연결과 회사 로그인 정책을 확인해 주세요.');
      } finally { global.clearTimeout(timeout); }
      if (ownGeneration !== generation) throw failure('auth-cancelled', '로그인이 취소되었거나 회사 앱 설정이 변경되었습니다.');
      const expiry = Number(responseData.expires_in), granted = String(responseData.scope || '').split(/\s+/);
      if (responseData.token_type !== 'Bearer' || typeof responseData.access_token !== 'string' || !responseData.access_token || responseData.access_token.length > 32768 || !Number.isFinite(expiry) || expiry <= 0 || expiry > 86400 || typeof responseData.id_token !== 'string') throw failure('auth-response', '회사 로그인 응답이 정상적인 액세스 토큰 형식이 아닙니다.');
      if (!granted.includes('Files.ReadWrite.All') && !granted.includes(GRAPH_SCOPE)) throw failure('auth-permission', '회사 공유 파일에 필요한 Microsoft Graph Files.ReadWrite.All 위임 권한이 허용되지 않았습니다. 회사 IT에 확인해 주세요.');
      const account = identityClaims(responseData.id_token, pending);
      token = { value: responseData.access_token, expiresAt: Date.now() + expiry * 1000, account }; lastError = '';
      return inspect();
    })().catch((error) => { token = null; lastError = error.message; throw error; });
    return initialization;
  }
  async function getAccessToken(options) {
    await initialize();
    const minimum = options && Number.isFinite(options.minimumValidityMs) ? Math.max(0, options.minimumValidityMs) : 60000;
    if (!token) throw failure('auth-required', '회사 공유 엑셀을 사용하려면 Microsoft 365에 로그인해 주세요.');
    if (token.expiresAt <= Date.now() + minimum) { token = null; throw failure('auth-expired', '회사 로그인이 만료되어 다시 로그인해야 합니다. 아직 저장되지 않은 결과를 먼저 보관해 주세요.'); }
    return token.value;
  }
  function signOut() { generation++; token = null; initialization = null; callback = null; clearPending(); lastError = ''; return inspect(); }
  global.ToolM365Auth = Object.freeze({ configure, inspect, initialize, signIn, getAccessToken, signOut });
})(window);
