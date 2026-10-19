/* Decode the offline OCR payload only when OCR is first requested. */
(function (global) {
  'use strict';
  let pending = null;

  function decodeBase64(value, language) {
    if (typeof value !== 'string' || value.length < 100) {
      throw new Error(language + ' 인식 데이터가 없습니다. 프로그램 HTML 파일을 다시 받아 주세요.');
    }
    let binary;
    try { binary = global.atob(value); }
    catch (_) { throw new Error(language + ' 인식 데이터가 손상되었습니다. 프로그램 파일을 다시 받아 주세요.'); }
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
    if (bytes[0] !== 0x1f || bytes[1] !== 0x8b) {
      throw new Error(language + ' 인식 데이터 형식이 올바르지 않습니다.');
    }
    return bytes;
  }

  async function readRuntime() {
    const element = document.getElementById('runtime-assets');
    if (!element) throw new Error('오프라인 문자 인식 엔진이 없습니다. 전체 프로그램 HTML 파일을 열어 주세요.');
    let data;
    try { data = JSON.parse(element.textContent); }
    catch (_) { throw new Error('문자 인식 엔진 정보가 손상되었습니다. 프로그램 파일을 다시 받아 주세요.'); }
    if (!data || typeof data.workerSource !== 'string' || data.workerSource.length < 1000 ||
        typeof data.coreSource !== 'string' || data.coreSource.indexOf('data:application/octet-stream;base64,AGFzbQ') === -1 ||
        !data.languages) {
      throw new Error('오프라인 문자 인식 엔진이 완전하지 않습니다. 프로그램 파일을 다시 받아 주세요.');
    }
    // Yield once so the UI can paint its loading state before decoding models.
    await new Promise(function (resolve) { setTimeout(resolve, 0); });
    return {
      workerSource: data.workerSource,
      coreSource: data.coreSource,
      languages: { kor: decodeBase64(data.languages.kor, '한글'), eng: decodeBase64(data.languages.eng, '영문') }
    };
  }

  global.ToolRuntime = {
    ready: function () {
      if (!pending) pending = readRuntime().catch(function (error) { pending = null; throw error; });
      return pending;
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
