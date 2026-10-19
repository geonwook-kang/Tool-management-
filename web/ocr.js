/* Offline browser OCR. Photos and workbook contents remain on this device. */
(function (global) {
  'use strict';
  let workerPromise = null, activeWorker = null, workerGeneration = 0;
  let queue = Promise.resolve(), progressListener = null, progressStart = 0, progressSpan = 1;
  let workerURLs = [], rawWorkers = [], currentPhase = '사진을 준비하는 중', currentProgress = 0;
  let activeCancellation = null, cancellationGeneration = 0;
  const unique = (values) => Array.from(new Set(values));
  const cleanSerial = (value) => String(value == null ? '' : value).normalize('NFKC').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const cleanName = (value) => String(value || '').normalize('NFKC').replace(/\s/g, '').toUpperCase();
  const score = (value) => value != null && value !== '' && Number.isFinite(Number(value)) ? Math.max(0, Math.min(100, Number(value))) : null;
  const escapeRegex = (value) => value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  function emit(status, progress) {
    currentPhase = status; currentProgress = Math.max(0, Math.min(1, progress || 0));
    if (typeof progressListener === 'function') {
      try { progressListener({ status, progress: currentProgress }); } catch (_) { /* A display must not interrupt OCR. */ }
    }
  }
  function dateOccurrences(text) {
    const result = [], pattern = /(?:^|[^0-9])(20\d{2})[ \t]*[.\-/년][ \t]*(\d{1,2})[ \t]*[.\-/월][ \t]*(\d{1,2})(?!\d)/g;
    let found;
    while ((found = pattern.exec(text))) {
      const year = Number(found[1]), month = Number(found[2]), day = Number(found[3]);
      const date = new Date(Date.UTC(year, month - 1, day));
      if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) {
        result.push({ value: [year, String(month).padStart(2, '0'), String(day).padStart(2, '0')].join('-'), index: found.index + found[0].search(/20\d{2}/), end: pattern.lastIndex });
      }
    }
    return result;
  }
  const datesIn = (text) => unique(dateOccurrences(text).map((date) => date.value)).sort();
  function identityPattern(value, serial) {
    const normalized = serial ? cleanSerial(value) : cleanName(value);
    return new RegExp('(?:^|[^' + (serial ? 'A-Z0-9' : '가-힣A-Za-z') + '])' + Array.from(normalized).map(escapeRegex).join(serial ? '[ \\t._\\-]*' : '[ \\t]*') + '(?![' + (serial ? 'A-Z0-9' : '가-힣A-Za-z') + '])', 'i');
  }
  function hasIdentity(text, value, serial) {
    const pattern = identityPattern(value, serial);
    if (!serial) return pattern.test(text);
    const dates = dateOccurrences(text), repeated = new RegExp(pattern.source, 'gi');
    let match;
    while ((match = repeated.exec(text))) {
      if (!dates.some((date) => match.index < date.end && repeated.lastIndex > date.index)) return true;
    }
    return false;
  }
  function valueConfidence(pass, value, kind) {
    if (!pass.words.length) return pass.confidence;
    const normalize = kind === 'name' ? cleanName : cleanSerial, target = normalize(value);
    let best = null;
    for (let i = 0; i < pass.words.length; i++) {
      let combined = '', minimum = 100;
      for (let j = i; j < Math.min(i + 12, pass.words.length); j++) {
        const word = pass.words[j], confidence = score(word.confidence);
        combined += String(word.text || '');
        if (confidence == null) break;
        minimum = Math.min(minimum, confidence);
        const matches = kind === 'date' ? datesIn(combined).includes(value) : normalize(combined) === target;
        if (matches) best = best == null ? minimum : Math.max(best, minimum);
        if (kind !== 'date' && normalize(combined).length > target.length) break;
      }
    }
    return best;
  }
  function labelsIn(text) {
    const labels = [], expressions = [
      ['due', /차\s*기\s*(?:검\s*)?교\s*정\s*(?:예\s*정\s*)?일\s*자?|(?:the\s+)?due\s+date(?:\s+of)?(?:\s+next\s+calibration)?|next\s+calibration(?:\s+date)?|cal\s+next/gi],
      ['last', /(?:검\s*)?교\s*정\s*일\s*자?|date\s+of\s+calibration|calibration\s+date|cal\s+last/gi]
    ];
    expressions.forEach((entry) => {
      let match;
      while ((match = entry[1].exec(text))) {
        const start = match.index, end = start + match[0].length;
        if (!labels.some((label) => start < label.end && end > label.index)) labels.push({ role: entry[0], index: start, end });
      }
    });
    return labels.sort((a, b) => a.index - b.index);
  }
  function datePass(pass) {
    const occurrences = dateOccurrences(pass.text), labels = labelsIn(pass.text), roles = { last: [], due: [] };
    occurrences.forEach((date, index) => {
      const before = labels.filter((label) => label.end <= date.index && date.index - label.end <= 140).pop();
      if (before && (!index || occurrences[index - 1].index < before.end)) roles[before.role].push({ value: date.value, confidence: valueConfidence(pass, date.value, 'date') });
    });
    return { dates: unique(occurrences.map((date) => date.value)).sort(), roles, dueLabel: labels.some((label) => label.role === 'due') };
  }
  function parse(text, tools, options) {
    options = options || {};
    const source = String(text || '').normalize('NFKC');
    tools = Array.isArray(tools) ? tools : [];
    const supplied = Array.isArray(options.passes) && options.passes.length ? options.passes : [{ text: source, confidence: options.confidence, words: options.words }];
    const passes = supplied.map((pass) => {
      if (typeof pass === 'string') pass = { text: pass };
      return { text: String(pass.text || '').normalize('NFKC'), source: String(pass.source || ''), confidence: score(pass.confidence == null ? options.confidence : pass.confidence), words: (Array.isArray(pass.words) ? pass.words : []).map((word) => ({ text: String(word.text || ''), confidence: score(word.confidence) })) };
    });
    text = unique([source].concat(passes.map((pass) => pass.text))).filter(Boolean).join('\n');
    const named = tools.filter((tool) => cleanName(tool.name).length >= 2 && hasIdentity(text, tool.name, false));
    const numbered = tools.filter((tool) => cleanSerial(tool.serial).length >= 3 && hasIdentity(text, tool.serial, true));
    const names = unique(named.map((tool) => String(tool.name))), serials = unique(numbered.map((tool) => String(tool.serial)));
    const observedNames = names.slice(), observedSerials = serials.map(cleanSerial);
    const ignoredNames = /^(?:교정필증|교정일자|교정일|차기교정|성적서|성적서번호|예정일자|다음교정|인정기구|한국인정|시리얼|이름|담당자|엔지니어|검교정|검교정일|교정|일자|예정|차기|인정|번호|기구|필증|날짜|성명|라벨|검사|관리|담당|도구|장비|한국|동아)$/;
    text.split(/\r?\n/).forEach((line) => {
      const name = /^\s*(?:(?:성명|이름|담당자|엔지니어|owner|name)\s*[:：]?\s*)?([가-힣](?:[ \t]*[가-힣]){1,3})(?:\s*|[ \t]+[0-9][A-Za-z0-9 ._-]*)$/i.exec(line);
      if (name && !ignoredNames.test(cleanName(name[1]))) observedNames.push(cleanName(name[1]));
      const explicitName = /(?:성명|이름|담당자|엔지니어|owner|name)\s*[:：]\s*([가-힣A-Za-z][가-힣A-Za-z \t]{1,30})/i.exec(line);
      if (explicitName && !ignoredNames.test(cleanName(explicitName[1]))) observedNames.push(cleanName(explicitName[1]));
      const explicitSerial = /(?:serial(?:\s*(?:number|no\.?))?|시리얼(?:\s*넘버|\s*번호)?|s\s*\/\s*n)\s*[:：#]?\s*([A-Z0-9][A-Z0-9._-]{2,})/i.exec(line);
      if (explicitSerial) observedSerials.push(cleanSerial(explicitSerial[1]));
      const dates = dateOccurrences(line), numeric = /(?:^|[^A-Z0-9])([A-Z0-9]{5,20})(?![A-Z0-9])/gi;
      let match;
      while ((match = numeric.exec(line))) {
        const index = match.index + match[0].indexOf(match[1]);
        if (/\d/.test(match[1]) && !dates.some((date) => index >= date.index && index < date.end)) {
          if (/^\d{5,20}$/.test(match[1]) || /^\s*[A-Z0-9]{5,20}\s*$/i.test(line)) observedSerials.push(cleanSerial(match[1]));
        }
      }
    });
    function bestConfidence(value, kind) {
      const values = passes.filter((pass) => hasIdentity(pass.text, value, kind === 'serial')).map((pass) => valueConfidence(pass, value, kind)).filter((value) => value != null);
      return values.length ? Math.max.apply(Math, values) : null;
    }
    const identityEvidence = {
      names, serials, observedNames: unique(observedNames.map(cleanName)), observedSerials: unique(observedSerials),
      nameMatches: named.map((tool) => ({ toolId: tool.id, name: tool.name, confidence: bestConfidence(tool.name, 'name') })),
      serialMatches: numbered.map((tool) => ({ toolId: tool.id, serial: tool.serial, confidence: bestConfidence(tool.serial, 'serial') }))
    };
    const readings = passes.map(datePass);
    const sourceCovered = passes.some((pass) => pass.text === source) || passes.map((pass) => pass.text).join('\n') === source;
    const allReadings = sourceCovered ? readings : readings.concat([datePass({ text: source, confidence: score(options.confidence), words: [] })]);
    const dateEvidence = { dates: datesIn(text), last: '', due: '', method: '', confidence: null, explicitLast: '', explicitDue: '', lastConfidence: null, dueConfidence: null, dueLabelSeen: allReadings.some((reading) => reading.dueLabel), ambiguous: false, repeatCount: 0 };
    ['last', 'due'].forEach((role) => {
      const values = unique(allReadings.flatMap((reading) => reading.roles[role].map((date) => date.value)));
      if (values.length > 1) dateEvidence.ambiguous = true;
      if (values.length === 1) {
        const scores = readings.flatMap((reading) => reading.roles[role]).filter((date) => date.value === values[0] && date.confidence != null).map((date) => date.confidence);
        dateEvidence[role === 'last' ? 'explicitLast' : 'explicitDue'] = values[0];
        dateEvidence[role + 'Confidence'] = scores.length ? Math.max.apply(Math, scores) : null;
      }
    });
    if (!dateEvidence.ambiguous && dateEvidence.explicitLast && dateEvidence.explicitDue) {
      dateEvidence.last = dateEvidence.explicitLast; dateEvidence.due = dateEvidence.explicitDue; dateEvidence.method = 'labelled';
      if (dateEvidence.lastConfidence != null && dateEvidence.dueConfidence != null) dateEvidence.confidence = Math.min(dateEvidence.lastConfidence, dateEvidence.dueConfidence);
    } else if (!dateEvidence.ambiguous && dateEvidence.dates.length === 2) {
      const pair = dateEvidence.dates.join(',');
      const agreements = readings.map((reading, index) => {
        if (reading.dates.length !== 2 || reading.dates.join(',') !== pair) return null;
        const confidences = reading.dates.map((date) => valueConfidence(passes[index], date, 'date'));
        return confidences.every((value) => value != null) ? Math.min.apply(Math, confidences) : null;
      }).filter((value) => value != null && value >= 70);
      dateEvidence.repeatCount = agreements.length;
      const compatible = (!dateEvidence.explicitLast || dateEvidence.explicitLast === dateEvidence.dates[0]) && (!dateEvidence.explicitDue || dateEvidence.explicitDue === dateEvidence.dates[1]);
      if (agreements.length >= 2 && compatible) {
        dateEvidence.last = dateEvidence.dates[0]; dateEvidence.due = dateEvidence.dates[1]; dateEvidence.method = 'repeated-dates'; dateEvidence.confidence = Math.min.apply(Math, agreements);
      }
    }
    const overlap = named.filter((tool) => numbered.some((other) => other.id === tool.id));
    const candidates = overlap.length ? overlap : named.concat(numbered), warnings = [];
    if (identityEvidence.observedNames.length > 1 || identityEvidence.observedSerials.length > 1) warnings.push('서로 다른 이름 또는 시리얼 번호가 읽혀 자동 반영을 보류했습니다.');
    if (!names.length && !serials.length) warnings.push('엑셀과 정확히 일치하는 이름 또는 시리얼 번호를 읽지 못했습니다.');
    if (named.length && numbered.length && !overlap.length) warnings.push('사진의 이름과 시리얼 번호가 서로 다른 엑셀 항목을 가리킵니다.');
    if (!dateEvidence.method) warnings.push('교정일·다음 교정일의 라벨 또는 두 번의 일치하는 날짜 인식이 필요합니다.');
    if (dateEvidence.ambiguous) warnings.push('교정일 또는 다음 교정일이 서로 다르게 읽혔습니다.');
    return { text, name: names.length === 1 ? names[0] : '', serial: serials.length === 1 ? serials[0] : '', last: dateEvidence.last || dateEvidence.explicitLast || '', due: dateEvidence.due || dateEvidence.explicitDue || '', failed: false, candidateIds: unique(candidates.map((tool) => tool.id)), warnings, confidence: score(options.confidence), identityEvidence, dateEvidence, passes };
  }

  function deadline(promise, ms, message) {
    let timer;
    const racers = [promise, new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(message)), ms); })];
    if (activeCancellation) racers.push(activeCancellation.promise);
    return Promise.race(racers).finally(() => clearTimeout(timer));
  }
  function discardWorker() {
    workerGeneration += 1; workerPromise = null;
    if (activeWorker) { try { activeWorker.terminate(); } catch (_) {} activeWorker = null; }
    rawWorkers.forEach((worker) => { try { worker.terminate(); } catch (_) {} }); rawWorkers = [];
    workerURLs.forEach((url) => URL.revokeObjectURL(url)); workerURLs = [];
  }
  function cancel() {
    cancellationGeneration += 1;
    if (activeCancellation) {
      activeCancellation.cancelled = true;
      activeCancellation.reject(new Error('인식을 중지했습니다.'));
    }
    discardWorker();
  }
  function createTrackedWorker(languages, options) {
    // Tesseract creates its Worker synchronously, but returns the public handle
    // only after initialization. Capture that one constructor call so Cancel
    // can terminate even a worker stuck before initialization completes.
    const NativeWorker = global.Worker;
    if (typeof NativeWorker !== 'function') return global.Tesseract.createWorker(languages, 1, options);
    function TrackedWorker(url, settings) {
      const worker = new NativeWorker(url, settings); rawWorkers.push(worker); return worker;
    }
    TrackedWorker.prototype = NativeWorker.prototype;
    let replaced = false;
    try { global.Worker = TrackedWorker; replaced = global.Worker === TrackedWorker; } catch (_) {}
    try { return global.Tesseract.createWorker(languages, 1, options); }
    finally { if (replaced && global.Worker === TrackedWorker) global.Worker = NativeWorker; }
  }
  async function runtimeAssets() {
    const runtime = global.ToolRuntime;
    if (!runtime) throw new Error('오프라인 문자 인식 데이터가 없는 구버전입니다. 최신 프로그램 파일을 다시 열어 주세요.');
    const assets = typeof runtime.ready === 'function' ? await runtime.ready() : runtime;
    if (!assets || !assets.workerSource || !assets.coreSource || !assets.languages || !assets.languages.kor || !assets.languages.eng) throw new Error('내장 문자 인식 데이터가 손상되었습니다. 프로그램 파일을 다시 내려받아 주세요.');
    return assets;
  }
  function getWorker() {
    if (!global.Tesseract || typeof global.Tesseract.createWorker !== 'function') return Promise.reject(new Error('문자 인식 모듈이 없습니다. 최신 프로그램 파일을 다시 열어 주세요.'));
    if (!workerPromise) {
      const generation = workerGeneration;
      workerPromise = (async () => {
        const assets = await runtimeAssets();
        if (generation !== workerGeneration) throw new Error('문자 인식 준비가 취소되었습니다.');
        emit('내장 한글·영문 인식 엔진을 준비하는 중', 0.05);
        // Tesseract 5.1.1 loads {code,data} language objects correctly, but its
        // initialize handler mistakenly uses .data as the language filename.
        // Adapt only that message, before the library's normal message handler.
        const languageAdapter = 'self.addEventListener("message",function(event){var message=event.data;if(message&&message.action==="initialize"&&message.payload&&Array.isArray(message.payload.langs)){message.payload.langs=message.payload.langs.map(function(lang){return typeof lang==="string"?lang:lang.code;});}});\n';
        const workerURL = URL.createObjectURL(new Blob([languageAdapter, assets.workerSource], { type: 'application/javascript' }));
        const coreURL = URL.createObjectURL(new Blob([assets.coreSource], { type: 'application/javascript' }));
        workerURLs.push(workerURL, coreURL);
        let initializationReject;
        const failed = new Promise((_, reject) => { initializationReject = reject; });
        const creation = createTrackedWorker([{ code: 'kor', data: assets.languages.kor }, { code: 'eng', data: assets.languages.eng }], {
          workerPath: workerURL, workerBlobURL: false, corePath: coreURL + '#core.js', cacheMethod: 'none', gzip: true,
          logger: (event) => {
            if (generation !== workerGeneration) return;
            const statuses = { 'loading tesseract core': '내장 문자 인식 엔진 준비 중', 'initializing tesseract': '문자 인식 엔진 초기화 중', 'loading language traineddata': '내장 한글·영문 데이터 준비 중', 'initializing api': '문자 인식 준비 중', 'recognizing text': '사진의 이름·시리얼·날짜를 읽는 중' };
            const progress = event.status === 'recognizing text' ? progressStart + (Number(event.progress) || 0) * progressSpan : Math.max(0.03, Math.min(0.13, (Number(event.progress) || 0) * 0.13));
            emit(statuses[event.status] || '문자 인식 준비 중', progress);
          },
          errorHandler: (error) => initializationReject(new Error(String(error || '문자 인식 엔진 오류')))
        }).then((worker) => {
          if (generation !== workerGeneration) { worker.terminate(); throw new Error('문자 인식 준비가 취소되었습니다.'); }
          activeWorker = worker; return worker;
        });
        return Promise.race([creation, failed]);
      })();
    }
    return deadline(workerPromise, 45000, '문자 인식 엔진을 45초 안에 준비하지 못했습니다. 파일을 Chrome 또는 Edge에서 다시 열어 주세요.');
  }
  async function imageCanvas(blob) {
    if (!(blob instanceof Blob)) throw new Error('사진 파일을 선택해 주세요.');
    let image;
    if (typeof global.createImageBitmap === 'function') { try { image = await global.createImageBitmap(blob, { imageOrientation: 'from-image' }); } catch (_) {} }
    if (!image) image = await new Promise((resolve, reject) => {
      const img = new Image(), url = URL.createObjectURL(blob);
      img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('사진을 열지 못했습니다. JPG, PNG 또는 WebP 사진을 사용해 주세요.')); };
      img.src = url;
    });
    const width = image.width, height = image.height;
    if (!width || !height || width * height > 40000000) { if (image.close) image.close(); throw new Error('사진은 4천만 픽셀 이하로 올려 주세요.'); }
    const scale = Math.min(2400 / Math.max(width, height), Math.max(1, Math.min(2, 1400 / width)));
    const canvas = document.createElement('canvas'); canvas.width = Math.round(width * scale); canvas.height = Math.round(height * scale);
    const context = canvas.getContext('2d'); context.fillStyle = '#fff'; context.fillRect(0, 0, canvas.width, canvas.height);
    context.filter = 'grayscale(1) contrast(1.15)'; context.drawImage(image, 0, 0, canvas.width, canvas.height);
    if (image.close) image.close(); return canvas;
  }
  function cropCanvas(canvas, left, top, width, height) {
    const crop = document.createElement('canvas'); crop.width = Math.round(canvas.width * width); crop.height = Math.round(canvas.height * height);
    crop.getContext('2d').drawImage(canvas, Math.round(canvas.width * left), Math.round(canvas.height * top), crop.width, crop.height, 0, 0, crop.width, crop.height); return crop;
  }
  async function runRecognition(blob, tools, onProgress) {
    progressListener = onProgress; progressStart = 0.15; progressSpan = 0.64;
    const cancellation = { cancelled: false, reject: null, promise: null };
    cancellation.promise = new Promise((_, reject) => { cancellation.reject = reject; });
    cancellation.promise.catch(() => {});
    activeCancellation = cancellation;
    const began = Date.now(), passes = [];
    const heartbeat = setInterval(() => {
      if (typeof progressListener === 'function') {
        try { progressListener({ status: currentPhase + ' · ' + Math.floor((Date.now() - began) / 1000) + '초', progress: currentProgress }); } catch (_) {}
      }
    }, 2000);
    function record(response, source) {
      const data = response.data || {};
      const words = Array.isArray(data.words) ? data.words : (data.blocks || []).flatMap((block) => (block.paragraphs || []).flatMap((paragraph) => (paragraph.lines || []).flatMap((line) => line.words || [])));
      passes.push({ text: data.text || '', confidence: data.confidence, source, words: words.map((word) => ({ text: word.text || '', confidence: word.confidence })) });
    }
    const current = () => parse(passes.map((pass) => pass.text).join('\n'), tools, { passes });
    const strongIdentity = (result) => result.identityEvidence.serialMatches.concat(result.identityEvidence.nameMatches).some((match) => match.confidence != null && match.confidence >= 70);
    try {
      emit('사진을 준비하는 중', 0.01);
      const canvas = await deadline(imageCanvas(blob), 15000, '사진을 15초 안에 열지 못했습니다. 다른 JPG 또는 PNG 사진으로 다시 시도해 주세요.');
      const worker = await getWorker();
      await deadline(worker.setParameters({ tessedit_pageseg_mode: '11', preserve_interword_spaces: '1' }), 5000, '문자 인식 설정을 적용하지 못했습니다. 다시 시도해 주세요.');
      record(await deadline(worker.recognize(canvas, {}, { text: true, blocks: true }), 60000, '사진을 60초 안에 읽지 못했습니다. 라벨을 크게 촬영한 사진으로 다시 시도해 주세요.'), 'original');
      let preliminary = current();
      if (!strongIdentity(preliminary) || preliminary.dateEvidence.method !== 'labelled' || !(preliminary.dateEvidence.confidence >= 70)) {
        progressStart = 0.8; progressSpan = 0.1;
        emit('교정 라벨을 한 번 더 읽는 중', progressStart);
        await deadline(worker.setParameters({ tessedit_pageseg_mode: '6' }), 5000, '문자 인식 설정에 실패했습니다.');
        record(await deadline(worker.recognize(cropCanvas(canvas, 0.05, 0.4, 0.9, 0.6), {}, { text: true, blocks: true }), 45000, '날짜 추가 인식 시간이 초과되었습니다. 날짜를 확인해 주세요.'), 'label-crop');
        preliminary = current();
      }
      if (!strongIdentity(preliminary)) {
        progressStart = 0.91; progressSpan = 0.08;
        emit('이름과 시리얼 표지를 한 번 더 읽는 중', progressStart);
        await deadline(worker.setParameters({ tessedit_pageseg_mode: '6' }), 5000, '문자 인식 설정에 실패했습니다.');
        record(await deadline(worker.recognize(cropCanvas(canvas, 0.08, 0.14, 0.8, 0.48), {}, { text: true, blocks: true }), 45000, '이름·시리얼 추가 인식 시간이 초과되었습니다.'), 'identity-crop');
      }
      const result = current(); emit('이름·시리얼·교정 날짜 인식 완료', 1); return result;
    } catch (error) {
      discardWorker();
      const result = current(); result.failed = true; result.cancelled = cancellation.cancelled;
      result.error = String(error && error.message || error || '문자 인식 오류').slice(0, 240);
      result.warnings.unshift(cancellation.cancelled ? '인식을 중지했습니다. 다시 인식 버튼으로 시작할 수 있습니다.' : '자동 인식을 완료하지 못했습니다. 다시 시도하거나 사진을 보며 직접 입력할 수 있습니다.'); result.warnings.push(result.error);
      emit(result.error, 1); return result;
    } finally { clearInterval(heartbeat); progressListener = null; if (activeCancellation === cancellation) activeCancellation = null; }
  }
  global.ToolOCR = {
    parse,
    recognize(blob, tools, onProgress) {
      const generation = cancellationGeneration;
      const next = queue.then(() => {
        if (generation !== cancellationGeneration) return Object.assign(parse('', tools), { failed: true, cancelled: true, error: '인식을 중지했습니다.' });
        return runRecognition(blob, tools, onProgress);
      });
      queue = next.catch(() => {}); return next;
    },
    cancel,
    reset: cancel
  };
})(typeof window !== 'undefined' ? window : globalThis);
