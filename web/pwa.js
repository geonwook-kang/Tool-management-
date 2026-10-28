/* Install/update controls for the HTTPS application shell. No user file data
   enters service-worker messages or the public application cache. */
(function (global) {
  'use strict';
  const BUILD = '__PWA_BUILD__';
  let registration = null, promptEvent = null, requestedReload = false, reloadGuard = null, installationConfirmed = false;
  const details = { version: BUILD, supported: false, installed: false, offlineReady: false, updateAvailable: false, error: '' };
  const $ = id => document.getElementById(id);
  const installed = () => !!(global.matchMedia && global.matchMedia('(display-mode: standalone)').matches || global.navigator.standalone === true);
  const ios = () => /iPhone|iPad|iPod/i.test(navigator.userAgent) || navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1;
  function isBusy() {
    const app = global.ToolApp && global.ToolApp.state;
    if (app && (app.busy || app.preparing || app.source && app.source.saving)) return true;
    if (reloadGuard) { try { return reloadGuard() !== true; } catch (_) { return true; } }
    return !!(app && Array.isArray(app.photos) && app.photos.some(photo => !photo.saved));
  }
  function state(text) { if ($('pwa-state')) $('pwa-state').textContent = text; }
  function render() {
    details.installed = installed() || installationConfirmed;
    if ($('pwa-install')) {
      $('pwa-install').hidden = details.installed;
      $('pwa-install').textContent = promptEvent ? '앱 설치' : '홈 화면에 추가';
    }
    if ($('pwa-update')) $('pwa-update').hidden = !details.updateAvailable;
  }
  function message(worker, type, timeout = 6000) {
    return new Promise((resolve, reject) => {
      if (!worker) { reject(new Error('앱 저장을 준비하고 있습니다.')); return; }
      const channel = new MessageChannel(), timer = setTimeout(() => { channel.port1.close(); reject(new Error('앱 상태를 확인하지 못했습니다.')); }, timeout);
      channel.port1.onmessage = event => { clearTimeout(timer); channel.port1.close(); resolve(event.data); };
      try { worker.postMessage({ type }, [channel.port2]); }
      catch (error) { clearTimeout(timer); channel.port1.close(); reject(error); }
    });
  }
  async function checkOffline() {
    try {
      const result = await message(registration && registration.active, 'TOOL_PWA_STATUS');
      details.offlineReady = !!(result && result.complete && result.version === BUILD);
      if (details.offlineReady) state('앱 저장 완료 · 다음 실행부터 인터넷 없이 사진을 인식할 수 있습니다.');
      else if (result && result.complete) state('기존 앱이 저장돼 있습니다. 새 버전을 적용하면 현재 앱도 오프라인으로 사용할 수 있습니다.');
      else state('오프라인 앱 저장을 완료하지 못했습니다. 인터넷 연결 상태에서 다시 열어 주세요.');
    } catch (error) { details.error = error.message; state(error.message); }
    render();
  }
  async function install() {
    const help = $('pwa-install-help');
    if (details.installed) return true;
    if (promptEvent) {
      const event = promptEvent; promptEvent = null;
      try { await event.prompt(); const choice = await event.userChoice; render(); return choice && choice.outcome === 'accepted'; }
      catch (_) { /* A browser menu remains available if a prompt expires. */ }
    }
    if (help) {
      help.hidden = false;
      help.textContent = ios() ? '아이폰·아이패드: Safari에서 이 주소를 열고 공유 버튼 → 「홈 화면에 추가」 → 「추가」를 누르세요.' : 'Chrome 또는 Edge의 메뉴에서 「앱 설치」 또는 「홈 화면에 추가」를 선택하세요. 설치 메뉴가 없으면 이 주소를 즐겨찾기로 사용할 수 있습니다.';
    }
    return false;
  }
  async function applyUpdate() {
    if (isBusy()) { state('사진 처리·저장 또는 미저장 작업이 있습니다. 작업을 마친 뒤 새 버전을 적용하세요.'); return false; }
    if (!registration || !registration.waiting) {
      if (details.updateAvailable) { requestedReload = true; global.location.reload(); return true; }
      return false;
    }
    try {
      requestedReload = true;
      const result = await message(registration.waiting, 'TOOL_PWA_ACTIVATE', 10000);
      if (!result || !result.accepted) {
        requestedReload = false; state('다른 앱 창에 처리·미저장 작업이 있습니다. 모든 앱 창의 작업을 마친 뒤 다시 적용하세요.'); return false;
      }
      return true;
    } catch (error) { requestedReload = false; state(error.message); return false; }
  }
  function waiting() {
    details.updateAvailable = !!(registration && registration.waiting);
    if (details.updateAvailable) state('새 버전이 준비됐습니다. 사진 작업을 마친 뒤 「새 버전 적용」을 누르세요.');
    render();
  }
  async function boot() {
    render();
    if ($('pwa-install')) $('pwa-install').onclick = install;
    if ($('pwa-update')) $('pwa-update').onclick = applyUpdate;
    if (!('serviceWorker' in navigator) || !global.isSecureContext || !/^https?:$/.test(location.protocol)) {
      state('홈 화면 설치와 오프라인 앱 저장은 HTTPS 앱 주소에서 사용할 수 있습니다.'); return;
    }
    details.supported = true;
    state('오프라인 인식 엔진을 휴대폰에 저장하는 중… 처음에는 인터넷 연결을 유지해 주세요.');
    navigator.serviceWorker.addEventListener('message', event => {
      if (event.data && event.data.type === 'TOOL_PWA_CAN_ACTIVATE' && event.ports[0]) event.ports[0].postMessage({ idle: !isBusy() });
    });
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (requestedReload && !isBusy()) { global.location.reload(); return; }
      // Other open windows retain their in-memory workbook/photos. Applying an
      // update in one window must not reload another window's unsaved work.
      if (details.offlineReady) {
        details.updateAvailable = true; state('새 버전이 적용됐습니다. 작업을 마친 뒤 「새 버전 적용」으로 이 창을 새로 여세요.'); render();
      } else checkOffline();
    });
    try {
      registration = await navigator.serviceWorker.register('./service-worker.js', { scope: './', updateViaCache: 'none' });
      registration.addEventListener('updatefound', () => {
        const worker = registration.installing;
        if (worker) worker.addEventListener('statechange', () => {
          if (worker.state === 'installed') {
            if (registration.active && registration.waiting) waiting();
            else navigator.serviceWorker.ready.then(checkOffline);
          }
          if (worker.state === 'redundant' && !details.offlineReady) state('앱 저장을 완료하지 못했습니다. 기존에 저장된 앱은 계속 사용할 수 있습니다. 인터넷 연결 상태에서 다시 열어 주세요.');
        });
      });
      waiting();
      if (registration.active) await checkOffline();
      else navigator.serviceWorker.ready.then(checkOffline).catch(() => {});
    } catch (error) {
      // A saved shell remains usable when a browser cannot check its script
      // for updates while offline. Query its existing active registration.
      try {
        const existing = await navigator.serviceWorker.getRegistration('./');
        if (existing && existing.active) { registration = existing; waiting(); await checkOffline(); return; }
      } catch (_) {}
      details.error = error.message; state('오프라인 앱을 저장하지 못했습니다. 이 화면은 계속 사용할 수 있습니다. ' + error.message);
    }
  }
  global.addEventListener('beforeinstallprompt', event => { event.preventDefault(); promptEvent = event; render(); });
  global.addEventListener('appinstalled', () => { promptEvent = null; installationConfirmed = true; details.installed = true; render(); state(details.offlineReady ? '앱 설치와 오프라인 저장을 완료했습니다.' : '홈 화면에 추가됐습니다. 오프라인 앱 저장도 완료될 때까지 기다려 주세요.'); });
  global.ToolPWA = Object.freeze({ inspect: () => Object.assign({}, details), isBusy, install, applyUpdate, setReloadGuard: guard => { reloadGuard = typeof guard === 'function' ? guard : null; } });
  if (document.readyState === 'loading') global.addEventListener('DOMContentLoaded', boot); else boot();
})(window);
