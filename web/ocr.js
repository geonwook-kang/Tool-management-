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
  const median = (values) => { const sorted = values.slice().sort((a, b) => a - b); return sorted.length ? sorted[Math.floor(sorted.length / 2)] : 0; };
  function wordBox(value) {
    if (!value) return null;
    const box = { x0: Number(value.x0), y0: Number(value.y0), x1: Number(value.x1), y1: Number(value.y1) };
    return Object.values(box).every(Number.isFinite) && box.x1 > box.x0 && box.y1 > box.y0 ? box : null;
  }
  function normalizeWord(word) {
    const result = { text: String(word.text || '').normalize('NFKC'), confidence: score(word.confidence) };
    const box = wordBox(word.bbox), lineBox = wordBox(word.lineBox || word.line && word.line.bbox);
    if (box) result.bbox = box;
    if (lineBox) result.lineBox = lineBox;
    return result;
  }
  const unionBox = (words) => ({ x0: Math.min(...words.map((w) => w.bbox.x0)), y0: Math.min(...words.map((w) => w.bbox.y0)), x1: Math.max(...words.map((w) => w.bbox.x1)), y1: Math.max(...words.map((w) => w.bbox.y1)) });
  const centerY = (box) => (box.y0 + box.y1) / 2;
  const boxHeight = (box) => box.y1 - box.y0;

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
      ['last', /(?:검\s*)?교\s*정\s*일\s*자?|date\s+of\s+calibration|calibration\s+date|cal\s+last/gi],
      ['other', /제\s*조\s*일\s*자?|발\s*행\s*일\s*자?|성\s*적\s*서\s*발\s*행\s*일|date\s+of\s+(?:manufacture|issue)|manufactur(?:e|ing)\s+date|certificate\s+date/gi]
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
  function physicalRows(words) {
    const rows = [];
    words.slice().sort((a, b) => centerY(a.bbox) - centerY(b.bbox) || a.bbox.x0 - b.bbox.x0).forEach((word) => {
      const cy = centerY(word.bbox), h = boxHeight(word.bbox), separator = /^[.\-/]+$/.test(word.text);
      // Median anchors prevent one tall/rotated word from chaining two rows.
      const candidates = rows.map((row) => ({ row, distance: Math.abs(cy - median(row.map((w) => centerY(w.bbox)))), height: median(row.map((w) => boxHeight(w.bbox))) }))
        .filter((item) => item.distance <= Math.max(2, separator ? item.height * 0.75 : Math.min(h, item.height) * 0.6)).sort((a, b) => a.distance - b.distance);
      if (candidates.length) candidates[0].row.push(word); else rows.push([word]);
    });
    return rows.map((row) => row.sort((a, b) => a.bbox.x0 - b.bbox.x0)).sort((a, b) => median(a.map((w) => centerY(w.bbox))) - median(b.map((w) => centerY(w.bbox))));
  }
  function wordText(words) {
    let text = '';
    const spans = words.map((word) => { if (text) text += ' '; const index = text.length; text += word.text; return { word, index, end: text.length }; });
    return { text, spans };
  }
  function spanWords(joined, span) { return joined.spans.filter((item) => item.index < span.end && item.end > span.index).map((item) => item.word); }
  function occurrenceConfidence(words) {
    const scores = words.map((word) => score(word.confidence));
    return scores.length && scores.every((value) => value != null) ? Math.min(...scores) : null;
  }
  function spatialDates(pass) {
    const words = pass.words.filter((word) => word.bbox && (dateOccurrences(word.text).length || /^[\d.\-/년월일\s]+$/.test(word.text)));
    return physicalRows(words).flatMap((row) => {
      const joined = wordText(row);
      return dateOccurrences(joined.text).map((date) => {
        const selected = spanWords(joined, date);
        return { value: date.value, bbox: unionBox(selected), words: selected, confidence: occurrenceConfidence(selected) };
      });
    });
  }
  function spatialLabels(pass) {
    const words = pass.words.filter((word) => word.bbox && /[가-힣A-Za-z]/.test(word.text));
    const rows = physicalRows(words), candidates = [];
    const sequences = [words, rows.flat()].concat(rows);
    sequences.forEach((sequence) => {
      const joined = wordText(sequence);
      labelsIn(joined.text).forEach((label) => {
        const selected = spanWords(joined, label), bbox = unionBox(selected);
        if (!selected.length || bbox.y1 - bbox.y0 > median(selected.map((word) => boxHeight(word.bbox))) * 5) return;
        if (!candidates.some((item) => item.role === label.role && ['x0', 'x1', 'y0', 'y1'].every((key) => Math.abs(item.bbox[key] - bbox[key]) < 1))) candidates.push({ role: label.role, bbox, words: selected });
      });
    });
    return candidates;
  }
  function geometryDatePass(pass) {
    const dates = spatialDates(pass), labels = spatialLabels(pass), roles = { last: [], due: [] }, chosen = [];
    let ambiguous = false;
    labels.forEach((label) => {
      const edges = dates.map((date) => {
        const h = Math.max(boxHeight(date.bbox), median(label.words.map((word) => boxHeight(word.bbox))));
        const verticalGap = Math.max(0, label.bbox.y0 - date.bbox.y1, date.bbox.y0 - label.bbox.y1);
        const horizontalGap = date.bbox.x0 - label.bbox.x1;
        const shared = label.words.some((word) => date.words.includes(word) && labelsIn(word.text).filter((item) => item.role !== 'other').length === 1 && dateOccurrences(word.text).length === 1);
        const right = horizontalGap >= -h * 0.2 && verticalGap <= h * 0.55;
        const below = date.bbox.y0 >= label.bbox.y1 - h * 0.2 && verticalGap <= h * 2.5 && Math.abs(date.bbox.x0 - label.bbox.x0) <= h * 1.5;
        if (!shared && !right && !below) return null;
        if (right && labels.some((other) => other !== label && other.role !== label.role && other.bbox.x0 >= label.bbox.x1 && other.bbox.x1 <= date.bbox.x0 && Math.abs(centerY(other.bbox) - centerY(date.bbox)) <= h)) return null;
        const distance = shared ? -1 : right ? Math.abs(centerY(label.bbox) - centerY(date.bbox)) / h + Math.max(0, horizontalGap) / h * 0.01 : verticalGap / h + 0.5;
        return { label, date, distance };
      }).filter(Boolean).sort((a, b) => a.distance - b.distance);
      if (!edges.length) return;
      // Two dates on the same row without a separating label are unresolved.
      if (edges[1] && edges[1].date.value !== edges[0].date.value && edges[1].distance - edges[0].distance < 0.3) return;
      chosen.push(edges[0]);
    });
    dates.forEach((date) => {
      const matches = chosen.filter((edge) => edge.date === date), assignedRoles = unique(matches.map((edge) => edge.label.role));
      if (assignedRoles.length > 1) { ambiguous = true; return; }
      if (assignedRoles.length === 1 && assignedRoles[0] !== 'other') roles[assignedRoles[0]].push({ value: date.value, confidence: date.confidence, bbox: date.bbox });
    });
    return { dates: unique(dates.map((date) => date.value)).sort(), roles, dueLabel: labels.some((label) => label.role === 'due'), lastLabel: labels.some((label) => label.role === 'last'), otherLabel: labels.some((label) => label.role === 'other'), ambiguous, geometry: true };
  }
  function datePass(pass) {
    if (pass.words.some((word) => word.bbox)) return geometryDatePass(pass);
    const occurrences = dateOccurrences(pass.text), labels = labelsIn(pass.text), roles = { last: [], due: [] };
    occurrences.forEach((date, index) => {
      const previous = index ? occurrences[index - 1].end : 0;
      const pending = labels.filter((label) => label.end <= date.index && label.index >= previous && date.index - label.end <= 140);
      const before = pending[pending.length - 1];
      // Text column order (all labels, then all dates) is not row evidence.
      if (before && before.role !== 'other' && unique(pending.map((label) => label.role)).length === 1) roles[before.role].push({ value: date.value, confidence: valueConfidence(pass, date.value, 'date') });
    });
    return { dates: unique(occurrences.map((date) => date.value)).sort(), roles, dueLabel: labels.some((label) => label.role === 'due'), lastLabel: labels.some((label) => label.role === 'last'), otherLabel: labels.some((label) => label.role === 'other'), ambiguous: false, geometry: false };
  }
  function parse(text, tools, options) {
    options = options || {};
    const source = String(text || '').normalize('NFKC');
    tools = Array.isArray(tools) ? tools : [];
    const supplied = Array.isArray(options.passes) && options.passes.length ? options.passes : [{ text: source, confidence: options.confidence, words: options.words }];
    const passes = supplied.map((pass) => {
      if (typeof pass === 'string') pass = { text: pass };
      return { text: String(pass.text || '').normalize('NFKC'), source: String(pass.source || ''), confidence: score(pass.confidence == null ? options.confidence : pass.confidence), words: (Array.isArray(pass.words) ? pass.words : []).map(normalizeWord) };
    });
    const passTexts = unique(passes.map((pass) => pass.text)).filter(Boolean);
    const sourceCovered = passes.some((pass) => pass.text === source) || passes.map((pass) => pass.text).join('\n') === source || passTexts.join('\n') === source;
    const identityPasses = passes.filter((pass) => !/^(?:date|label)[-_]crop/.test(pass.source));
    const identityText = unique((sourceCovered ? [] : [source]).concat(identityPasses.map((pass) => pass.text))).join('\n');
    text = sourceCovered ? passTexts.join('\n') : unique([source].concat(passTexts)).filter(Boolean).join('\n');
    const named = tools.filter((tool) => cleanName(tool.name).length >= 2 && hasIdentity(identityText, tool.name, false));
    const numbered = tools.filter((tool) => cleanSerial(tool.serial).length >= 3 && hasIdentity(identityText, tool.serial, true));
    const names = unique(named.map((tool) => String(tool.name))), serials = unique(numbered.map((tool) => String(tool.serial)));
    const observedNames = names.slice(), observedSerials = serials.map(cleanSerial);
    const nameEvidence = [], serialEvidence = [];
    function observeName(name, kind, pass) {
      name = cleanName(name);
      const confidence = valueConfidence(pass, name, 'name');
      let entry = nameEvidence.find((item) => item.name === name && item.kind === kind);
      if (!entry) { entry = { name, kind, confidence: null, matchedToolIds: tools.filter((tool) => cleanName(tool.name) === name).map((tool) => tool.id), passSources: [] }; nameEvidence.push(entry); }
      if (confidence != null) entry.confidence = entry.confidence == null ? confidence : Math.max(entry.confidence, confidence);
      if (!entry.passSources.includes(pass.source || 'text')) entry.passSources.push(pass.source || 'text');
    }
    function observeSerial(serial, kind, pass) {
      serial = cleanSerial(serial);
      const confidence = valueConfidence(pass, serial, 'serial');
      let entry = serialEvidence.find((item) => item.serial === serial && item.kind === kind);
      if (!entry) { entry = { serial, kind, confidence: null, matchedToolIds: tools.filter((tool) => cleanSerial(tool.serial) === serial).map((tool) => tool.id), passSources: [] }; serialEvidence.push(entry); }
      if (confidence != null) entry.confidence = entry.confidence == null ? confidence : Math.max(entry.confidence, confidence);
      if (!entry.passSources.includes(pass.source || 'text')) entry.passSources.push(pass.source || 'text');
    }
    const ignoredNames = /^(?:교정필증|교정일자|교정일|차기교정|성적서|성적서번호|예정일자|다음교정|인정기구|한국인정|시리얼|이름|담당자|엔지니어|검교정|검교정일|교정|일자|예정|차기|인정|번호|기구|필증|날짜|성명|라벨|검사|관리|담당|도구|장비|한국|동아)$/;
    const observationPasses = sourceCovered ? identityPasses : identityPasses.concat([{ text: source, confidence: score(options.confidence), words: [], source: 'text' }]);
    observationPasses.forEach((pass) => {
      named.forEach((tool) => { if (hasIdentity(pass.text, tool.name, false)) observeName(tool.name, 'workbook', pass); });
      numbered.forEach((tool) => { if (hasIdentity(pass.text, tool.serial, true)) observeSerial(tool.serial, 'workbook', pass); });
      pass.text.split(/\r?\n/).forEach((line) => {
      const name = /^\s*(?:(?:성명|이름|담당자|엔지니어|owner|name)\s*[:：]?\s*)?([가-힣](?:[ \t]*[가-힣]){1,3})(?:\s*|[ \t]+[0-9][A-Za-z0-9 ._-]*)$/i.exec(line);
      if (name && !ignoredNames.test(cleanName(name[1]))) { observedNames.push(cleanName(name[1])); observeName(name[1], /(?:성명|이름|담당자|엔지니어|owner|name)/i.test(line) ? 'labelled' : 'standalone', pass); }
      const explicitName = /(?:성명|이름|담당자|엔지니어|owner|name)\s*[:：]\s*([가-힣]{2,4}|[A-Za-z][A-Za-z \t]{1,30})(?![가-힣])/i.exec(line);
      if (explicitName && !ignoredNames.test(cleanName(explicitName[1]))) { observedNames.push(cleanName(explicitName[1])); observeName(explicitName[1], 'labelled', pass); }
      const explicitSerial = /(?:serial(?:\s*(?:number|no\.?))?|시리얼(?:\s*넘버|\s*번호)?|s\s*\/\s*n)\s*[:：#]?\s*([A-Z0-9][A-Z0-9._-]{2,})/i.exec(line);
      if (explicitSerial) { observedSerials.push(cleanSerial(explicitSerial[1])); observeSerial(explicitSerial[1], 'labelled', pass); }
      const dates = dateOccurrences(line), numeric = /(?:^|[^A-Z0-9])([A-Z0-9]{5,20})(?![A-Z0-9])/gi;
      let match;
      while ((match = numeric.exec(line))) {
        const index = match.index + match[0].indexOf(match[1]);
        if (/\d/.test(match[1]) && !dates.some((date) => index >= date.index && index < date.end)) {
          if (/^\d{5,20}$/.test(match[1]) || /^\s*[A-Z0-9]{5,20}\s*$/i.test(line)) { observedSerials.push(cleanSerial(match[1])); observeSerial(match[1], 'standalone', pass); }
        }
      }
      });
    });
    function bestConfidence(value, kind) {
      const values = identityPasses.filter((pass) => hasIdentity(pass.text, value, kind === 'serial')).map((pass) => valueConfidence(pass, value, kind)).filter((value) => value != null);
      return values.length ? Math.max.apply(Math, values) : null;
    }
    const identityEvidence = {
      names, serials, observedNames: unique(observedNames.map(cleanName)), observedNameEvidence: nameEvidence, observedSerials: unique(observedSerials), observedSerialEvidence: serialEvidence,
      nameMatches: named.map((tool) => ({ toolId: tool.id, name: tool.name, confidence: bestConfidence(tool.name, 'name') })),
      serialMatches: numbered.map((tool) => ({ toolId: tool.id, serial: tool.serial, confidence: bestConfidence(tool.serial, 'serial') }))
    };
    const datePasses = passes.filter((pass) => !/^identity[-_]crop/.test(pass.source));
    const dateText = (sourceCovered ? [] : [source]).concat(datePasses.map((pass) => pass.text)).join('\n');
    const readings = datePasses.map(datePass);
    const allReadings = sourceCovered ? readings : readings.concat([datePass({ text: source, confidence: score(options.confidence), words: [] })]);
    const dateEvidence = { dates: unique(datesIn(dateText).concat(readings.flatMap((reading) => reading.dates))).sort(), last: '', due: '', method: '', confidence: null, explicitLast: '', explicitDue: '', lastConfidence: null, dueConfidence: null, dueLabelSeen: allReadings.some((reading) => reading.dueLabel), lastLabelSeen: allReadings.some((reading) => reading.lastLabel), otherLabelSeen: allReadings.some((reading) => reading.otherLabel), ambiguous: allReadings.some((reading) => reading.ambiguous), repeatCount: 0 };
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
    } else if (!dateEvidence.ambiguous && !dateEvidence.otherLabelSeen && (!dateEvidence.lastLabelSeen || dateEvidence.explicitLast) && (!dateEvidence.dueLabelSeen || dateEvidence.explicitDue) && dateEvidence.dates.length === 2) {
      const pair = dateEvidence.dates.join(',');
      const agreements = readings.map((reading, index) => {
        if (reading.dates.length !== 2 || reading.dates.join(',') !== pair) return null;
        const confidences = reading.dates.map((date) => valueConfidence(datePasses[index], date, 'date'));
        return confidences.every((value) => value != null) ? Math.min.apply(Math, confidences) : null;
      }).filter((value) => value != null && value >= 70);
      dateEvidence.repeatCount = agreements.length;
      const compatible = (!dateEvidence.explicitLast || dateEvidence.explicitLast === dateEvidence.dates[0]) && (!dateEvidence.explicitDue || dateEvidence.explicitDue === dateEvidence.dates[1]);
      if (agreements.length >= 2 && compatible) {
        dateEvidence.last = dateEvidence.dates[0]; dateEvidence.due = dateEvidence.dates[1]; dateEvidence.method = 'repeated-dates'; dateEvidence.confidence = Math.min.apply(Math, agreements);
      }
    }
    const overlap = named.filter((tool) => numbered.some((other) => other.id === tool.id));
    if (dateEvidence.method && dateEvidence.confidence != null && dateEvidence.confidence >= 70) {
      const confirmedDates = [dateEvidence.last, dateEvidence.due].map((value) => value.replace(/-/g, ''));
      // A serial-only pass may lose the dots in a date already confirmed by
      // date OCR. It is not another serial unless explicitly labelled as such
      // or present as a real serial in the workbook.
      identityEvidence.observedSerialEvidence = serialEvidence.filter((item) => !(item.kind === 'standalone' && !item.matchedToolIds.length && confirmedDates.includes(item.serial)));
      identityEvidence.observedSerials = identityEvidence.observedSerials.filter((serial) => identityEvidence.observedSerialEvidence.some((item) => item.serial === serial));
    }
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
        // Local Windows file:// pages have an opaque origin. Importing a second
        // blob:null URL from their worker can fail even though creating the
        // first worker succeeded. Put the verified core and worker in the same
        // classic script, so getCore sees the factory without importScripts.
        // Fail before the worker bundle if the embedded core is unavailable;
        // there must be no network fallback for an offline program.
        const registerCore = '\n;if(typeof TesseractCore!=="function"){throw new Error("내장 문자 인식 엔진을 불러오지 못했습니다.");}self.TesseractCore=TesseractCore;\n';
        const workerURL = URL.createObjectURL(new Blob([assets.coreSource, registerCore, languageAdapter, assets.workerSource], { type: 'application/javascript' }));
        workerURLs.push(workerURL);
        let initializationReject;
        const failed = new Promise((_, reject) => { initializationReject = reject; });
        const creation = createTrackedWorker([{ code: 'kor', data: assets.languages.kor }, { code: 'eng', data: assets.languages.eng }], {
          workerPath: workerURL, workerBlobURL: false, cacheMethod: 'none', gzip: true,
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
  function dateCrop(canvas, pass) {
    const dates = spatialDates(pass), labels = spatialLabels(pass).filter((label) => label.role !== 'other');
    let region;
    if (dates.length >= 2 && dates.length <= 8) {
      const bbox = unionBox(dates.map((date) => ({ bbox: date.bbox })));
      const width = median(dates.map((date) => date.bbox.x1 - date.bbox.x0)), height = median(dates.map((date) => boxHeight(date.bbox)));
      // Re-read the date values independently of tiny table text. Retain every
      // observed date; never select two by chronological order or discard a
      // third date. Semantic labels remain available from the original pass.
      region = { x0: bbox.x0 - width * 0.12, x1: bbox.x1 + width * 0.12, y0: bbox.y0 - height * 0.45, y1: bbox.y1 + height * 0.45 };
    } else if (labels.length >= 2) {
      const bbox = unionBox(labels.map((label) => ({ bbox: label.bbox })));
      const width = bbox.x1 - bbox.x0, height = median(labels.map((label) => boxHeight(label.bbox)));
      region = { x0: bbox.x0 - height, x1: bbox.x1 + width * 2.5, y0: bbox.y0 - height, y1: bbox.y1 + height };
    }
    if (!region) return cropCanvas(canvas, 0.05, 0.4, 0.9, 0.6);
    const x = Math.max(0, Math.floor(region.x0)), y = Math.max(0, Math.floor(region.y0));
    const width = Math.min(canvas.width, Math.ceil(region.x1)) - x, height = Math.min(canvas.height, Math.ceil(region.y1)) - y;
    if (width < 1 || height < 1) return cropCanvas(canvas, 0.05, 0.4, 0.9, 0.6);
    const factor = dates.length >= 2 ? Math.min(2, 1600 / Math.max(width, height)) : Math.min(5, 2200 / Math.max(width, height), Math.max(1, 1600 / Math.max(width, height)));
    const crop = document.createElement('canvas'); crop.width = Math.round(width * factor); crop.height = Math.round(height * factor);
    crop.getContext('2d').drawImage(canvas, x, y, width, height, 0, 0, crop.width, crop.height);
    return crop;
  }
  async function runRecognition(blob, tools, onProgress) {
    progressListener = onProgress; progressStart = 0.15; progressSpan = 0.64;
    const cancellation = { cancelled: false, reject: null, promise: null };
    cancellation.promise = new Promise((_, reject) => { cancellation.reject = reject; });
    cancellation.promise.catch(() => {});
    activeCancellation = cancellation;
    const began = Date.now(), passes = [];
    let stage = 'image-load';
    const heartbeat = setInterval(() => {
      if (typeof progressListener === 'function') {
        try { progressListener({ status: currentPhase + ' · ' + Math.floor((Date.now() - began) / 1000) + '초', progress: currentProgress }); } catch (_) {}
      }
    }, 2000);
    function record(response, source) {
      const data = response.data || {};
      const words = Array.isArray(data.words) ? data.words : (data.blocks || []).flatMap((block) => (block.paragraphs || []).flatMap((paragraph) => (paragraph.lines || []).flatMap((line) => line.words || [])));
      passes.push({ text: data.text || '', confidence: data.confidence, source, words: words.map(normalizeWord) });
    }
    const current = () => parse(passes.map((pass) => pass.text).join('\n'), tools, { passes });
    const strongIdentity = (result) => result.identityEvidence.serialMatches.concat(result.identityEvidence.nameMatches).some((match) => match.confidence != null && match.confidence >= 70);
    try {
      emit('사진을 준비하는 중', 0.01);
      const canvas = await deadline(imageCanvas(blob), 15000, '사진을 15초 안에 열지 못했습니다. 다른 JPG 또는 PNG 사진으로 다시 시도해 주세요.');
      stage = 'engine-init';
      const worker = await getWorker();
      stage = 'recognition';
      await deadline(worker.setParameters({ tessedit_pageseg_mode: '11', preserve_interword_spaces: '1' }), 5000, '문자 인식 설정을 적용하지 못했습니다. 다시 시도해 주세요.');
      record(await deadline(worker.recognize(canvas, {}, { text: true, blocks: true }), 60000, '사진을 60초 안에 읽지 못했습니다. 라벨을 크게 촬영한 사진으로 다시 시도해 주세요.'), 'original');
      let preliminary = current();
      if (!strongIdentity(preliminary) || preliminary.dateEvidence.method !== 'labelled' || !(preliminary.dateEvidence.confidence >= 70)) {
        progressStart = 0.8; progressSpan = 0.1;
        emit('교정 라벨을 한 번 더 읽는 중', progressStart);
        await deadline(worker.setParameters({ tessedit_pageseg_mode: '6' }), 5000, '문자 인식 설정에 실패했습니다.');
        record(await deadline(worker.recognize(dateCrop(canvas, passes[0]), {}, { text: true, blocks: true }), 45000, '날짜 추가 인식 시간이 초과되었습니다. 날짜를 확인해 주세요.'), 'date-crop');
        preliminary = current();
      }
      if (!strongIdentity(preliminary)) {
        progressStart = 0.91; progressSpan = 0.08;
        emit('이름과 시리얼 표지를 한 번 더 읽는 중', progressStart);
        await deadline(worker.setParameters({ tessedit_pageseg_mode: '6' }), 5000, '문자 인식 설정에 실패했습니다.');
        record(await deadline(worker.recognize(cropCanvas(canvas, 0.08, 0.14, 0.8, 0.48), {}, { text: true, blocks: true }), 45000, '이름·시리얼 추가 인식 시간이 초과되었습니다.'), 'identity-crop');
        preliminary = current();
        if (!strongIdentity(preliminary)) {
          emit('사진 하단의 시리얼 표지를 읽는 중', 0.97);
          record(await deadline(worker.recognize(cropCanvas(canvas, 0.05, 0.4, 0.9, 0.6), {}, { text: true, blocks: true }), 45000, '하단 시리얼 추가 인식 시간이 초과되었습니다.'), 'identity-crop-lower');
        }
      }
      const result = current(); result.engineFailure = false; result.code = ''; result.errorCode = '';
      emit('이름·시리얼·교정 날짜 인식 완료', 1); return result;
    } catch (error) {
      discardWorker();
      const result = current(); result.failed = true; result.cancelled = cancellation.cancelled;
      result.engineFailure = !cancellation.cancelled && stage === 'engine-init';
      result.code = cancellation.cancelled ? 'cancelled' : result.engineFailure ? 'engine-init' : stage === 'image-load' ? 'image-load' : 'ocr-failed';
      result.errorCode = result.code;
      result.error = String(error && error.message || error || '문자 인식 오류').slice(0, 240);
      result.warnings.unshift(cancellation.cancelled ? '인식을 중지했습니다. 다시 인식 버튼으로 시작할 수 있습니다.' : result.engineFailure ? '문자 인식 엔진을 시작하지 못해 사진의 글자를 읽지 못했습니다. 오류를 해결한 뒤 다시 시도해 주세요.' : '자동 인식을 완료하지 못했습니다. 다시 시도하거나 사진을 보며 직접 입력할 수 있습니다.'); result.warnings.push(result.error);
      emit(result.error, 1); return result;
    } finally { clearInterval(heartbeat); progressListener = null; if (activeCancellation === cancellation) activeCancellation = null; }
  }
  global.ToolOCR = {
    parse,
    recognize(blob, tools, onProgress) {
      const generation = cancellationGeneration;
      const next = queue.then(() => {
        if (generation !== cancellationGeneration) return Object.assign(parse('', tools), { failed: true, cancelled: true, engineFailure: false, code: 'cancelled', errorCode: 'cancelled', error: '인식을 중지했습니다.' });
        return runRecognition(blob, tools, onProgress);
      });
      queue = next.catch(() => {}); return next;
    },
    cancel,
    reset: cancel
  };
})(typeof window !== 'undefined' ? window : globalThis);
