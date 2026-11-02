/* One-time owner setup and company sign-in for the common-workbook entry. */
(function (global) {
  'use strict';
  const OWNER_KEY = 'tool-manager-shared-owner-intent-v1';
  const DRAFT_KEY = 'tool-manager-shared-owner-draft-v1';
  const SSO_KEY = 'tool-manager-shared-sso-attempt-v1:';
  const TTL = 10 * 60 * 1000;
  const sharedEntry = !!global.document.querySelector('meta[name="tool-manager-shared-entry"]');
  const element = id => global.document.getElementById(id);
  let callbacks = {}, initializing = null, bound = false, internalBusy = false, navigating = false;
  let ownerIntent = false, rejectedInvitation = false, lastError = '', interactionRequired = false;
  let view = { busy: false, connected: false, needsSignIn: false, error: '' };
  function failure(code, message) { const error = new Error(message); error.code = code; return error; }
  function text(id, value) { const node = element(id); if (node) node.textContent = value; }
  function hide(id, value) { const node = element(id); if (node) node.hidden = !!value; }
  function disable(id, value) { const node = element(id); if (node) node.disabled = !!value; }
  function status(message, kind) { if (typeof callbacks.status === 'function') callbacks.status(message, kind || ''); }
  function refresh() { render(view); if (typeof callbacks.refresh === 'function') callbacks.refresh(); }
  function workspaceInfo() { return global.ToolSharedWorkspace.inspect(); }
  function authInfo() { return global.ToolM365Auth.inspect(); }
  function rememberOwnerIntent() {
    ownerIntent = true;
    try { global.sessionStorage.setItem(OWNER_KEY, 'true'); } catch (_) {}
  }
  try {
    ownerIntent = new URL(global.location.href).searchParams.get('setup') === '1' || global.sessionStorage.getItem(OWNER_KEY) === 'true';
    if (ownerIntent) rememberOwnerIntent();
  } catch (_) {}
  function fillClient() {
    const auth = authInfo();
    if (element('shared-client-id')) element('shared-client-id').value = auth.clientId || '';
    if (element('shared-tenant-id')) element('shared-tenant-id').value = auth.tenantId || 'organizations';
    if (element('shared-redirect')) element('shared-redirect').value = auth.redirectUri;
  }
  function saveClient() {
    if (workspaceInfo().configured) throw failure('shared-config-fixed', '공동 작업 링크에 지정된 회사 앱을 사용합니다. 엑셀 변경은 아래 저장 대상 변경을 사용해 주세요.');
    const clientId = element('shared-client-id')?.value.trim() || '';
    const tenantId = element('shared-tenant-id')?.value.trim() || 'organizations';
    const result = global.ToolM365Auth.configure({ clientId, tenantId });
    rememberOwnerIntent(); fillClient(); lastError = ''; interactionRequired = false; refresh(); return result;
  }
  function configurePointerAuth() {
    const pointer = workspaceInfo().bootstrap;
    if (pointer) global.ToolM365Auth.configure({ clientId: pointer.clientId, tenantId: pointer.tenantId });
    else if (!authInfo().configured) saveClient();
  }
  function saveOwnerDraft() {
    if (!ownerIntent && !workspaceInfo().isOwner) return;
    const value = element('shared-workbook-url')?.value.trim() || '';
    if (value.length > 12000) return;
    try { global.sessionStorage.setItem(DRAFT_KEY, JSON.stringify({ createdAt: Date.now(), workbookUrl: value })); } catch (_) {}
  }
  function restoreOwnerDraft() {
    try {
      const value = JSON.parse(global.sessionStorage.getItem(DRAFT_KEY) || 'null');
      if (value && typeof value.workbookUrl === 'string' && value.workbookUrl.length <= 12000 && Number.isFinite(value.createdAt)
        && Date.now() - value.createdAt >= 0 && Date.now() - value.createdAt <= TTL && element('shared-workbook-url')) element('shared-workbook-url').value = value.workbookUrl;
      else global.sessionStorage.removeItem(DRAFT_KEY);
    } catch (_) {}
  }
  function preserveBookmark() {
    if (!workspaceInfo().configured) return;
    // Restore only after Auth.initialize consumed and scrubbed the callback.
    // A stable fragment keeps the same workspace on a home-screen shortcut.
    try { global.history.replaceState(null, '', global.ToolSharedWorkspace.invitationLink()); } catch (_) {}
  }
  function render(value) {
    if (!sharedEntry) return;
    if (value) view = { ...view, ...value };
    const info = workspaceInfo(), auth = authInfo(), locked = !!(view.busy || internalBusy || navigating || info.busy);
    const newOwner = !info.configured && ownerIntent, canManage = newOwner || info.isOwner;
    hide('shared-panel', false);
    hide('shared-owner-panel', !canManage);
    hide('shared-owner-toggle', info.configured && !info.isOwner);
    hide('shared-setup-link', info.configured || ownerIntent);
    const panel = element('shared-owner-panel');
    if (panel && newOwner) panel.open = true;
    const setup = element('shared-setup-link');
    if (setup) { const link = new URL(global.location.href); link.search = '?setup=1'; link.hash = ''; setup.href = link.href; }
    hide('shared-signin', auth.signedIn && !view.needsSignIn || !auth.configured || rejectedInvitation);
    hide('shared-create', info.configured);
    hide('shared-change-target', !info.isOwner);
    hide('shared-save-client', info.configured);
    for (const id of ['shared-signin', 'shared-copy-link', 'shared-share-link', 'shared-owner-toggle', 'shared-save-client', 'shared-client-id', 'shared-tenant-id', 'shared-workbook-url', 'shared-create', 'shared-change-target']) disable(id, locked);
    disable('shared-client-id', locked || info.configured); disable('shared-tenant-id', locked || info.configured);
    disable('shared-create', locked || !newOwner || !auth.signedIn);
    disable('shared-change-target', locked || !info.isOwner || !auth.signedIn);
    disable('shared-copy-link', locked || !info.configured); disable('shared-share-link', locked || !info.configured);
    disable('shared-refresh', locked || !auth.signedIn || !info.configured);
    disable('shared-signout', locked || !auth.signedIn);
    hide('shared-signout', !auth.signedIn);
    if (navigating) text('shared-auth-status', '회사 로그인으로 이동하고 있습니다…');
    else if (lastError) text('shared-auth-status', lastError);
    else if (view.needsSignIn && info.configured) text('shared-auth-status', '공유 엑셀을 사용하려면 회사 로그인을 확인해 주세요. 엑셀 연결 정보는 유지됩니다.');
    else if (auth.signedIn) text('shared-auth-status', '회사 계정 로그인됨' + (auth.account?.username ? ' · ' + auth.account.username : ''));
    else if (interactionRequired) text('shared-auth-status', '회사 계정으로 계속하려면 로그인을 확인해 주세요. 엑셀 연결 정보는 유지되어 있습니다.');
    else if (info.configured) text('shared-auth-status', '회사 계정으로 계속하면 담당자가 지정한 공유 엑셀에 자동 연결됩니다.');
    else text('shared-auth-status', ownerIntent ? '담당자의 최초 연결 정보를 저장하고 회사 계정으로 로그인하세요.' : '담당자가 보내준 공동 작업 링크를 열어 주세요.');
    text('shared-target', info.workbook ? '공동 저장 대상: ' + info.workbook.name : info.configured ? '담당자가 지정한 공동 엑셀을 확인합니다.' : '공동 엑셀 최초 설정 전입니다.');
    text('shared-owner-note', info.isOwner ? '저장 대상을 바꾸면 같은 공동 작업 링크를 사용하는 모든 직원에게 반영됩니다.' : '처음 한 번 연결한 뒤 직원용 공동 작업 링크를 공유하세요. 직원은 앱 ID나 파일 경로를 입력하지 않습니다.');
    text('shared-remember-note', info.persistent === false ? '이 브라우저에 링크를 기억하지 못했습니다. 담당자의 공동 작업 링크 자체를 보관해 주세요.' : '엑셀 연결은 공동 설정에 보관합니다. 회사 로그인 확인이 다시 필요해도 지정한 파일은 유지됩니다.');
    hide('shared-link-row', !info.configured);
    if (info.configured && element('shared-link-text')) {
      try { element('shared-link-text').value = global.ToolSharedWorkspace.invitationLink(); } catch (_) { element('shared-link-text').value = ''; }
    }
  }
  async function signIn(options) {
    const silent = options?.silent === true;
    if (!sharedEntry || internalBusy || navigating || view.busy) return false;
    if (rejectedInvitation) { status('담당자의 올바른 공동 작업 링크를 다시 열어 주세요.', 'error'); return false; }
    internalBusy = true; lastError = ''; interactionRequired = false; refresh();
    try {
      configurePointerAuth(); saveOwnerDraft();
      if (typeof callbacks.beforeSignIn === 'function' && await callbacks.beforeSignIn({ silent }) === false) throw failure('shared-signin-cancelled', '현재 사진을 보관한 뒤 로그인을 계속해 주세요.');
      navigating = true; refresh();
      await global.ToolM365Auth.signIn({ prompt: silent ? 'none' : '' });
      return true;
    } catch (error) {
      navigating = false; lastError = error.message;
      if (typeof callbacks.onSignInFailure === 'function') callbacks.onSignInFailure(error);
      status('회사 로그인 준비: ' + error.message, 'error'); return false;
    } finally { internalBusy = false; refresh(); }
  }
  async function automaticSignIn() {
    const info = workspaceInfo(); if (!info.configured || authInfo().signedIn || rejectedInvitation) return false;
    const key = SSO_KEY + info.bootstrap.workspaceId;
    try {
      const previous = Number(global.sessionStorage.getItem(key));
      if (Number.isFinite(previous) && previous > 0 && Date.now() - previous >= 0 && Date.now() - previous <= TTL) return false;
      const stamp = String(Date.now()); global.sessionStorage.setItem(key, stamp);
      if (global.sessionStorage.getItem(key) !== stamp) return false;
    } catch (_) { return false; }
    return signIn({ silent: true });
  }
  async function copyLink() {
    try {
      const link = global.ToolSharedWorkspace.invitationLink();
      if (global.navigator.clipboard?.writeText) {
        try { await global.navigator.clipboard.writeText(link); status('공동 작업 링크를 복사했습니다. 직원에게 이 링크를 보내세요.', 'success'); return true; } catch (_) {}
      }
      const field = element('shared-link-text'); if (field) { field.value = link; field.focus(); field.select(); }
      status('공동 작업 링크를 길게 눌러 복사한 뒤 직원에게 보내세요.'); return false;
    } catch (error) { status(error.message, 'error'); return false; }
  }
  async function shareLink() {
    try {
      const url = global.ToolSharedWorkspace.invitationLink();
      if (typeof global.navigator.share === 'function') await global.navigator.share({ title: '툴 매니저 공동 저장', text: '회사 계정으로 열고 툴 사진을 올려 주세요.', url });
      else await copyLink();
    } catch (error) { if (error.name !== 'AbortError') status('공동 링크 공유: ' + error.message, 'error'); }
  }
  async function changeTarget(create) {
    if (internalBusy || navigating || view.busy) return;
    internalBusy = true; lastError = ''; refresh();
    try {
      if (!create) {
        if (typeof callbacks.beforeChangeTarget === 'function') { if (await callbacks.beforeChangeTarget() === false) return; }
        else if (global.ToolApp?.state?.photos.some(photo => !photo.saved)) throw failure('shared-pending-photos', '현재 사진 작업을 완료한 뒤 저장 대상을 변경해 주세요.');
      }
      const workbookShareUrl = element('shared-workbook-url')?.value.trim() || '';
      if (create) await global.ToolSharedWorkspace.create({ workbookShareUrl });
      else await global.ToolSharedWorkspace.changeWorkbook({ workbookShareUrl, expectedVersion: workspaceInfo().configVersion });
      try { global.sessionStorage.removeItem(DRAFT_KEY); } catch (_) {}
      preserveBookmark();
      if (typeof callbacks.connect === 'function') await callbacks.connect({ targetChanged: !create });
      if (view.error) return;
      status(create ? '공동 저장 연결을 만들었습니다. 아래 공동 작업 링크를 직원에게 공유하세요.' : '공동 저장 대상을 변경했습니다. 직원들은 같은 링크를 계속 사용할 수 있습니다.', 'success');
    } catch (error) { lastError = error.message; status(error.message, 'error'); }
    finally { internalBusy = false; refresh(); }
  }
  function bind() {
    if (bound) return; bound = true;
    if (element('shared-signin')) element('shared-signin').onclick = () => signIn();
    if (element('shared-save-client')) element('shared-save-client').onclick = () => { try { saveClient(); status('최초 연결 정보를 저장했습니다. 회사 계정으로 계속하세요.', 'success'); } catch (error) { lastError = error.message; status(error.message, 'error'); render(); } };
    if (element('shared-copy-link')) element('shared-copy-link').onclick = copyLink;
    if (element('shared-share-link')) element('shared-share-link').onclick = shareLink;
    if (element('shared-create')) element('shared-create').onclick = () => changeTarget(true);
    if (element('shared-change-target')) element('shared-change-target').onclick = () => changeTarget(false);
    if (element('shared-owner-toggle')) element('shared-owner-toggle').onclick = () => { const panel = element('shared-owner-panel'), wasOpen = !!(panel && !panel.hidden && panel.open); rememberOwnerIntent(); render(); if (panel && !panel.hidden) panel.open = !wasOpen; };
  }
  async function initialize(options) {
    callbacks = { ...callbacks, ...(options || {}) };
    if (!sharedEntry) return { sharedEntry: false };
    if (initializing) return initializing;
    const run = (async () => {
      rejectedInvitation = false; lastError = ''; interactionRequired = false;
      bind(); render();
      try {
        const invitation = global.ToolSharedWorkspace.takeInvitation();
        if (invitation) global.ToolSharedWorkspace.configure(invitation);
      } catch (error) { rejectedInvitation = true; lastError = error.message; status(error.message, 'error'); refresh(); return { error }; }
      const pointer = workspaceInfo().bootstrap;
      if (pointer) global.ToolM365Auth.configure({ clientId: pointer.clientId, tenantId: pointer.tenantId });
      fillClient(); restoreOwnerDraft();
      let authError = null;
      try { await global.ToolM365Auth.initialize(); }
      catch (error) {
        authError = error;
        if (error.code === 'auth-interaction-required') { interactionRequired = true; status('회사 계정으로 계속하면 지정된 엑셀을 사용합니다.'); }
        else { lastError = error.message; status('회사 로그인 확인: ' + error.message, 'error'); }
      }
      // A successful round trip ends this attempt. A later page open has no
      // memory token and may use the company's existing SSO session again.
      if (!authError && authInfo().signedIn && workspaceInfo().bootstrap) {
        try { global.sessionStorage.removeItem(SSO_KEY + workspaceInfo().bootstrap.workspaceId); } catch (_) {}
      }
      preserveBookmark(); refresh();
      if (!authError && authInfo().signedIn && workspaceInfo().configured) {
        try { if (typeof callbacks.connect === 'function') await callbacks.connect(); }
        catch (error) { lastError = error.message; status(error.message, 'error'); }
        refresh();
      } else if (!authError && workspaceInfo().configured) await automaticSignIn();
      return { configured: workspaceInfo().configured, signedIn: authInfo().signedIn, error: authError };
    })();
    initializing = run;
    try { return await run; }
    finally { if (initializing === run) initializing = null; }
  }
  global.ToolSharedUI = Object.freeze({ initialize, render, signIn, copyLink });
})(window);
