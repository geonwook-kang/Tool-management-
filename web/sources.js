/* One active workbook source; local copies and Microsoft 365 stay distinct. */
(function (global) {
  'use strict';
  const key = 'tool-manager-workbook-source';
  const mobile = /Android|iPhone|iPad|iPod/i.test(global.navigator.userAgent) || (global.navigator.maxTouchPoints > 1 && /Macintosh/i.test(global.navigator.userAgent));
  const hosted = !!global.ToolSharePointContext;
  let mode = mobile ? 'sharepoint' : 'desktop';
  try { const saved = global.localStorage.getItem(key); if (['desktop', 'microsoft365', 'device', 'sharepoint'].includes(saved)) mode = saved; } catch (_) {}
  if (hosted || mobile && mode === 'microsoft365' && !global.ToolM365Auth?.inspect().configured) mode = 'sharepoint';
  const adapter = () => ({ desktop: global.ToolConnectedFile, microsoft365: global.ToolGraphFile, device: global.ToolMobileFile, sharepoint: global.ToolSharePointFile })[mode];
  function use(value) {
    if (!['desktop', 'microsoft365', 'device', 'sharepoint'].includes(value)) throw new Error('지원하지 않는 엑셀 연결 방식입니다.');
    if (hosted && value !== 'sharepoint') throw new Error('이 SharePoint 페이지에서는 지정된 공유 엑셀을 사용합니다.');
    mode = value;
    try { global.localStorage.setItem(key, mode); } catch (_) {}
    return inspect();
  }
  function inspect() {
    const backend = adapter();
    return Object.assign({ connected: false, persistent: false }, backend && backend.inspect(), { sourceMode: mode, mobile, hosted });
  }
  function call(name, args) {
    const backend = adapter();
    if (!backend || typeof backend[name] !== 'function') throw new Error('엑셀 연결 모듈을 준비하지 못했습니다. 앱을 새로 열어 주세요.');
    return backend[name](...args);
  }
  global.ToolSources = Object.freeze({
    mode: () => mode, use, inspect,
    connect: (...args) => call('connect', args), restore: (...args) => call('restore', args),
    requestAccess: (...args) => call('requestAccess', args), read: (...args) => call('read', args),
    save: (...args) => call('save', args), disconnect: (...args) => call('disconnect', args),
    takeBackup: (...args) => call('takeBackup', args)
  });
})(window);
