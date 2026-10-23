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
  const identityRows = new WeakMap();
  function identityProof(pass, value, serial) {
    const normalize = serial ? cleanSerial : cleanName, target = normalize(value);
    let proof = hasIdentity(pass.text, value, serial) ? { confidence: valueConfidence(pass, value, serial ? 'serial' : 'name'), words: [] } : null;
    let geometricMatch = false, geometricAccepted = false;
    if (!identityRows.has(pass)) identityRows.set(pass, physicalRows(pass.words.filter((word) => word.bbox)));
    identityRows.get(pass).forEach((row) => {
      for (let i = 0; i < row.length; i++) {
        const selected = []; let combined = '';
        for (let j = i; j < Math.min(row.length, i + 12); j++) {
          const word = row[j], previous = selected[selected.length - 1];
          if (!(serial ? /^[A-Za-z0-9 ._\-]+$/ : /^[가-힣A-Za-z \t]+$/).test(word.text)) break;
          if (previous && word.bbox.x0 - previous.bbox.x1 > Math.max(boxHeight(word.bbox), boxHeight(previous.bbox)) * 1.5) break;
          selected.push(word); combined += word.text;
          const actual = normalize(combined);
          if (!target.startsWith(actual)) break;
          const connected = (neighbor, edge) => neighbor && !/^(?:S\/?N|SERIAL|NO|NAME|OWNER|성명|이름|담당자|엔지니어|시리얼)$/i.test(neighbor.text) && Math.max(0, neighbor.bbox.x0 - edge.bbox.x1, edge.bbox.x0 - neighbor.bbox.x1) <= Math.max(boxHeight(neighbor.bbox), boxHeight(edge.bbox)) * 1.5 && (serial ? /^[A-Za-z0-9._\-]+$/.test(neighbor.text) && (/\d/.test(neighbor.text) || cleanSerial(neighbor.text).length <= 3) : /[가-힣]/.test(target) ? /^[가-힣]+$/.test(neighbor.text) : /^[A-Za-z]+$/.test(neighbor.text));
          if (actual === target) geometricMatch = true;
          if (actual === target && (selected.length > 1 || hasIdentity(pass.text, value, serial)) && !connected(row[i - 1], selected[0]) && !connected(row[j + 1], word) && (!serial || !dateOccurrences(selected.map((item) => item.text).join(' ')).length)) {
            geometricAccepted = true;
            const confidence = occurrenceConfidence(selected);
            if (!proof || proof.confidence == null || confidence != null && confidence >= proof.confidence) proof = { confidence, words: selected.slice() };
            break;
          }
        }
      }
    });
    return geometricMatch && !geometricAccepted ? null : proof;
  }
  function explicitSerials(line, tools) {
    const label = /(?:serial(?:\s*(?:number|no\.?))?|시리얼(?:\s*넘버|\s*번호)?|s\s*\/\s*n)\s*[:：#]?\s*/gi;
    const values = []; let match;
    while ((match = label.exec(line))) {
      const tail = line.slice(label.lastIndex), first = /^([A-Z0-9][A-Z0-9._-]{2,})/i.exec(tail);
      if (!first) continue;
      let value = first[1], rest = tail.slice(value.length);
      // Preserve actual spaced serial fragments. Never complete a value from
      // workbook characters, and never consume a following labelled field.
      while (/^\d+$/.test(cleanSerial(value))) {
        const next = /^[ \t]+(\d+)(?![A-Z0-9.\/-])/i.exec(rest);
        if (!next || cleanSerial(value).length + next[1].length > 20) break;
        value += next[1]; rest = rest.slice(next[0].length);
      }
      const suffix = /^[ \t]+([A-Z]{1,3})(?![A-Z0-9])/i.exec(rest);
      if (suffix && !/^(?:CAL|DUE|S|N|NO)$/i.test(suffix[1])) value += suffix[1];
      values.push(cleanSerial(value));
    }
    return values;
  }
  function boundarySerials(pass, tools) {
    const result = [];
    if (!identityRows.has(pass)) identityRows.set(pass, physicalRows(pass.words.filter((word) => word.bbox)));
    identityRows.get(pass).forEach((row) => {
      let group = [];
      const finish = () => {
        if (group.length > 1) {
          const value = cleanSerial(group.map((word) => word.text).join(''));
          if (value.length <= 24 && tools.some((tool) => { const serial = cleanSerial(tool.serial); return serial.length >= 3 && value !== serial && value.includes(serial) && !identityProof(pass, tool.serial, true); })) result.push({ serial: value, confidence: occurrenceConfidence(group) });
        }
        group = [];
      };
      row.forEach((word) => {
        const previous = group[group.length - 1];
        if (!/^[A-Za-z0-9._\-]+$/.test(word.text) || !/\d/.test(word.text) && cleanSerial(word.text).length > 3 || /^(?:SN|NO|CAL|DUE)$/i.test(word.text) || dateOccurrences(word.text).length) { finish(); return; }
        if (previous && word.bbox.x0 - previous.bbox.x1 > Math.max(boxHeight(word.bbox), boxHeight(previous.bbox)) * 1.5) finish();
        group.push(word);
      });
      finish();
    });
    return result;
  }
  function spatialDates(pass) {
    const words = pass.words.filter((word) => word.bbox && (dateOccurrences(word.text).length || /^[\d.\-/년월일\s]+$/.test(word.text)));
    const rawDates = datesIn(pass.text), completedWords = new Set(), lineGroups = new Map(), result = [];
    words.filter((word) => word.lineBox).forEach((word) => {
      const key = ['x0', 'y0', 'x1', 'y1'].map((part) => word.lineBox[part]).join(':');
      if (!lineGroups.has(key)) lineGroups.set(key, []);
      lineGroups.get(key).push(word);
    });
    function readRow(row, reconstructed) {
      row = row.slice().sort((a, b) => a.bbox.x0 - b.bbox.x0);
      const joined = wordText(row);
      return dateOccurrences(joined.text).map((date) => {
        const selected = spanWords(joined, date);
        const confidence = occurrenceConfidence(selected);
        if (selected.some((word, index) => index && word.bbox.x0 - selected[index - 1].bbox.x1 > Math.max(boxHeight(word.bbox), boxHeight(selected[index - 1].bbox)) * 1.5)) return null;
        // Do not invent a new calendar value from an unrelated faint fragment
        // that happens to overlap another line. Low-confidence dates actually
        // present in OCR text are retained for review and conflict handling.
        if (reconstructed && selected.length > 1 && !rawDates.includes(date.value) && (confidence == null || confidence < 70)) return null;
        return { value: date.value, bbox: unionBox(selected), words: selected, confidence: occurrenceConfidence(selected) };
      }).filter(Boolean);
    }
    lineGroups.forEach((row) => {
      const found = readRow(row, false);
      if (found.length) { result.push(...found); row.forEach((word) => completedWords.add(word)); }
    });
    // Prefer the recognizer's complete lines before using physical alignment
    // to recover words split into separate OCR blocks.
    physicalRows(words.filter((word) => !completedWords.has(word))).forEach((row) => result.push(...readRow(row, true)));
    return result;
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
      const normalized = { text: String(pass.text || '').normalize('NFKC'), source: String(pass.source || ''), confidence: score(pass.confidence == null ? options.confidence : pass.confidence), words: (Array.isArray(pass.words) ? pass.words : []).map(normalizeWord) };
      if (['serial', 'owner', 'due'].includes(pass.smaxRole)) normalized.smaxRole = pass.smaxRole;
      const labelBox = wordBox(pass.smaxLabelBBox), valueBox = wordBox(pass.smaxValueBBox);
      if (labelBox) normalized.smaxLabelBBox = labelBox;
      if (valueBox) normalized.smaxValueBBox = valueBox;
      if (typeof pass.smaxLabelText === 'string') normalized.smaxLabelText = pass.smaxLabelText;
      if (pass.smaxLabelVerified === true) normalized.smaxLabelVerified = true;
      return normalized;
    });
    const passTexts = unique(passes.map((pass) => pass.text)).filter(Boolean);
    const sourceCovered = passes.some((pass) => pass.text === source) || passes.map((pass) => pass.text).join('\n') === source || passTexts.join('\n') === source;
    const identityPasses = passes.filter((pass) => !/^(?:date|label)[-_](?:crop|probe)/.test(pass.source));
    const matchingPasses = sourceCovered ? identityPasses : identityPasses.concat([{ text: source, confidence: score(options.confidence), words: [], source: 'text' }]);
    const isProbe = (pass) => /^identity[-_]probe/.test(pass.source);
    const acceptedProof = (pass, value, serial) => { const proof = identityProof(pass, value, serial); return proof && (!isProbe(pass) || proof.confidence != null && proof.confidence >= 70) ? proof : null; };
    text = sourceCovered ? passTexts.join('\n') : unique([source].concat(passTexts)).filter(Boolean).join('\n');
    const named = tools.filter((tool) => cleanName(tool.name).length >= 2 && matchingPasses.some((pass) => acceptedProof(pass, tool.name, false)));
    const numbered = tools.filter((tool) => cleanSerial(tool.serial).length >= 3 && matchingPasses.some((pass) => acceptedProof(pass, tool.serial, true)));
    const names = unique(named.map((tool) => String(tool.name))), serials = unique(numbered.map((tool) => String(tool.serial)));
    const observedNames = names.slice(), observedSerials = serials.map(cleanSerial);
    const nameEvidence = [], serialEvidence = [], fragmentEvidence = [];
    function observeName(name, kind, pass) {
      name = cleanName(name);
      const proof = identityProof(pass, name, false), confidence = proof ? proof.confidence : valueConfidence(pass, name, 'name');
      let entry = nameEvidence.find((item) => item.name === name && item.kind === kind);
      if (!entry) { entry = { name, kind, confidence: null, matchedToolIds: tools.filter((tool) => cleanName(tool.name) === name).map((tool) => tool.id), passSources: [] }; nameEvidence.push(entry); }
      if (confidence != null) entry.confidence = entry.confidence == null ? confidence : Math.max(entry.confidence, confidence);
      if (!entry.passSources.includes(pass.source || 'text')) entry.passSources.push(pass.source || 'text');
    }
    function observeSerial(serial, kind, pass, suppliedConfidence) {
      serial = cleanSerial(serial);
      const proof = identityProof(pass, serial, true), confidence = suppliedConfidence != null ? suppliedConfidence : proof ? proof.confidence : valueConfidence(pass, serial, 'serial');
      let entry = serialEvidence.find((item) => item.serial === serial && item.kind === kind);
      if (!entry) { entry = { serial, kind, confidence: null, matchedToolIds: tools.filter((tool) => cleanSerial(tool.serial) === serial).map((tool) => tool.id), passSources: [], sourceReadings: [] }; serialEvidence.push(entry); }
      if (confidence != null) entry.confidence = entry.confidence == null ? confidence : Math.max(entry.confidence, confidence);
      const source = pass.source || 'text';
      if (!entry.passSources.includes(source)) entry.passSources.push(source);
      let reading = entry.sourceReadings.find((item) => item.source === source);
      if (!reading) { reading = { source, confidence: null }; entry.sourceReadings.push(reading); }
      if (confidence != null) reading.confidence = reading.confidence == null ? confidence : Math.max(reading.confidence, confidence);
    }
    function observeFragment(fragment, pass) {
      const proof = identityProof(pass, fragment, true);
      if (!proof) return; // A nearby digit/letter continuation is not a token.
      const confidence = proof.confidence, source = pass.source || 'text';
      let entry = fragmentEvidence.find((item) => item.fragment === fragment);
      if (!entry) { entry = { fragment, confidence: null, passSources: [], sourceReadings: [] }; fragmentEvidence.push(entry); }
      if (confidence != null) entry.confidence = entry.confidence == null ? confidence : Math.max(entry.confidence, confidence);
      if (!entry.passSources.includes(source)) entry.passSources.push(source);
      let reading = entry.sourceReadings.find((item) => item.source === source);
      if (!reading) { reading = { source, confidence: null }; entry.sourceReadings.push(reading); }
      if (confidence != null) reading.confidence = reading.confidence == null ? confidence : Math.max(reading.confidence, confidence);
    }
    const ignoredNames = /^(?:교정필증|교정일자|교정일|차기교정|성적서|성적서번호|예정일자|예정일|다음교정|인정기구|인정기관|교정기관|한국인정|주식회사|시리얼|이름|담당자|엔지니어|검교정|검교정일|교정|일자|예정|차기|인정|번호|기구|필증|날짜|성명|라벨|검사|관리|담당|도구|장비|한국|동아)$/;
    const probeSerials = [];
    const observationPasses = matchingPasses;
    observationPasses.forEach((pass) => {
      named.forEach((tool) => { if (acceptedProof(pass, tool.name, false)) observeName(tool.name, 'workbook', pass); });
      numbered.forEach((tool) => { if (acceptedProof(pass, tool.serial, true)) observeSerial(tool.serial, 'workbook', pass); });
      const matchedWords = new Set(numbered.flatMap((tool) => { const proof = acceptedProof(pass, tool.serial, true); return proof && proof.words.length > 1 ? proof.words : []; }));
      const isMatchedFragment = (value) => { const words = pass.words.filter((word) => cleanSerial(word.text) === cleanSerial(value)); return words.length && words.every((word) => matchedWords.has(word)); };
      boundarySerials(pass, tools).forEach((item) => {
        if (isProbe(pass)) { if (item.confidence != null && item.confidence >= 70) probeSerials.push({ serial: item.serial, confidence: item.confidence, pass }); }
        else { observedSerials.push(item.serial); observeSerial(item.serial, 'standalone', pass, item.confidence); }
      });
      pass.text.split(/\r?\n/).forEach((line, lineIndex, lines) => {
      const fragment = /^\s*(\d(?:[ \t]*\d){3})\s*$/.exec(line) || /^\s*(?:serial(?:\s*(?:number|no\.?))?|시리얼(?:\s*넘버|\s*번호)?|s\s*\/\s*n)\s*[:：#]?\s*(\d(?:[ \t]*\d){3})\s*$/i.exec(line);
      const fragmentContext = lines.slice(Math.max(0, lineIndex - 2), lineIndex).join(' ');
      if (fragment && !/incomplete/.test(pass.source) && !/^20\d{2}$/.test(cleanSerial(fragment[1])) && !/성\s*적\s*서|certificate|report\s*(?:no|number)|교\s*정\s*(?:예\s*정\s*)?일|date\s+of|cal\s+(?:last|next)|due\s+date|next\s+calibration|manufactur/i.test(fragmentContext)) observeFragment(cleanSerial(fragment[1]), pass);
      const name = /^\s*(?:(?:성명|이름|담당자|엔지니어|owner|name)\s*[:：]?\s*)?([가-힣](?:[ \t]*[가-힣]){1,3})(?:\s*|[ \t]+[0-9][A-Za-z0-9 ._-]*)$/i.exec(line);
      if (name && !isProbe(pass) && !ignoredNames.test(cleanName(name[1]))) { observedNames.push(cleanName(name[1])); observeName(name[1], /(?:성명|이름|담당자|엔지니어|owner|name)/i.test(line) ? 'labelled' : 'standalone', pass); }
      const explicitName = /(?:성명|이름|담당자|엔지니어|owner|name)\s*[:：]\s*([가-힣]{2,4}|[A-Za-z][A-Za-z \t]{1,30})(?![가-힣])/i.exec(line);
      if (explicitName && !ignoredNames.test(cleanName(explicitName[1]))) { observedNames.push(cleanName(explicitName[1])); observeName(explicitName[1], 'labelled', pass); }
      explicitSerials(line, tools).forEach((serial) => { observedSerials.push(serial); observeSerial(serial, 'labelled', pass); });
      const dates = dateOccurrences(line), numeric = /(?:^|[^A-Z0-9])([A-Z0-9]{5,20})(?![A-Z0-9])/gi;
      let match;
      while ((match = numeric.exec(line))) {
        const index = match.index + match[0].indexOf(match[1]);
        if (/\d/.test(match[1]) && !dates.some((date) => index >= date.index && index < date.end)) {
          if ((/^\d{5,20}$/.test(match[1]) || /^\s*[A-Z0-9]{5,20}\s*$/i.test(line)) && !isMatchedFragment(match[1])) {
            if (isProbe(pass)) {
              const proof = identityProof(pass, match[1], true);
              const confidence = proof ? proof.confidence : valueConfidence(pass, match[1], 'serial');
              if (confidence != null && confidence >= 70) probeSerials.push({ serial: cleanSerial(match[1]), confidence, pass });
            } else { observedSerials.push(cleanSerial(match[1])); observeSerial(match[1], 'standalone', pass); }
          }
        }
      }
      });
    });
    const hasConfidentSerial = numbered.some((tool) => matchingPasses.some((pass) => { const proof = acceptedProof(pass, tool.serial, true); return proof && proof.confidence != null && proof.confidence >= 70; }));
    unique(probeSerials.map((item) => item.serial)).forEach((serial) => {
      const observations = probeSerials.filter((item) => item.serial === serial);
      if (!hasConfidentSerial || unique(observations.map((item) => item.pass.source)).length >= 2) observations.forEach((item) => { observedSerials.push(serial); observeSerial(serial, 'standalone', item.pass, item.confidence); });
    });
    function bestConfidence(value, kind) {
      const values = matchingPasses.map((pass) => acceptedProof(pass, value, kind === 'serial')).filter(Boolean).map((proof) => proof.confidence).filter((value) => value != null);
      return values.length ? Math.max.apply(Math, values) : null;
    }
    const identityEvidence = {
      names, serials, observedNames: unique(observedNames.map(cleanName)), observedNameEvidence: nameEvidence, observedSerials: unique(observedSerials), observedSerialEvidence: serialEvidence, observedSerialFragments: fragmentEvidence,
      nameMatches: named.map((tool) => ({ toolId: tool.id, name: tool.name, confidence: bestConfidence(tool.name, 'name') })),
      serialMatches: numbered.map((tool) => ({ toolId: tool.id, serial: tool.serial, confidence: bestConfidence(tool.serial, 'serial') }))
    };
    const datePasses = passes.filter((pass) => !/^(?:identity[-_](?:crop|probe)|date[-_]probe|smax[-_]field)/.test(pass.source));
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
    const result = { text, name: names.length === 1 ? names[0] : '', serial: serials.length === 1 ? serials[0] : '', last: dateEvidence.last || dateEvidence.explicitLast || '', due: dateEvidence.due || dateEvidence.explicitDue || '', failed: false, identityRecoveryIncomplete: passes.some((pass) => pass.source === 'identity-probe-incomplete'), candidateIds: unique(candidates.map((tool) => tool.id)), warnings, confidence: score(options.confidence), identityEvidence, dateEvidence, passes, documentType: 'calibration' };
    if (global.ToolSMAX && typeof global.ToolSMAX.analyze === 'function') {
      let documentWarnings = [];
      try {
        const analysis = global.ToolSMAX.analyze(passes, tools);
        if (analysis && ['smax', 'unknown-smax', 'calibration'].includes(analysis.documentType)) {
          const calibrationWarnings = result.warnings.slice();
          Object.assign(result, analysis);
          if (result.documentType === 'calibration') result.warnings = unique(calibrationWarnings.concat(Array.isArray(analysis.warnings) ? analysis.warnings : []));
        }
        if (analysis && Array.isArray(analysis.warnings)) documentWarnings = analysis.warnings.slice();
      } catch (_) {
        result.documentType = 'unknown-smax'; result.smaxEvidence = { incomplete: true };
        documentWarnings.push('사진 유형과 필드를 확인하지 못해 자동 반영을 보류했습니다.');
      }
      if (result.documentType !== 'calibration') { result.last = ''; result.due = ''; result.warnings = documentWarnings; }
    }
    return result;
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
    // Keep native pixels for small labels. Enlarging already small strokes can
    // turn a 6 into an 8; targeted re-reads must also have the original scale.
    const nativeScale = Math.min(1, 3200 / Math.max(width, height));
    if (Math.abs(scale / nativeScale - 1) > 0.05) {
      const native = document.createElement('canvas'); native.width = Math.round(width * nativeScale); native.height = Math.round(height * nativeScale);
      const nativeContext = native.getContext('2d'); nativeContext.fillStyle = '#fff'; nativeContext.fillRect(0, 0, native.width, native.height);
      nativeContext.filter = 'grayscale(1) contrast(1.15)'; nativeContext.drawImage(image, 0, 0, native.width, native.height);
      canvas.nativeCanvas = native;
    }
    if (image.close) image.close(); return canvas;
  }
  function cropCanvas(canvas, left, top, width, height) {
    const crop = document.createElement('canvas'); crop.width = Math.round(canvas.width * width); crop.height = Math.round(canvas.height * height);
    crop.getContext('2d').drawImage(canvas, Math.round(canvas.width * left), Math.round(canvas.height * top), crop.width, crop.height, 0, 0, crop.width, crop.height); return crop;
  }
  function dateCrop(canvas, pass, nativeSize) {
    const dates = spatialDates(pass), labels = spatialLabels(pass).filter((label) => label.role !== 'other');
    let region;
    if (dates.length >= 1 && dates.length <= 8) {
      const bbox = unionBox(dates.map((date) => ({ bbox: date.bbox })));
      const width = median(dates.map((date) => date.bbox.x1 - date.bbox.x0)), height = median(dates.map((date) => boxHeight(date.bbox)));
      // Re-read the date values independently of tiny table text. Retain every
      // observed date; never select two by chronological order or discard a
      // third date. Semantic labels remain available from the original pass.
      const verticalMargin = dates.length === 1 ? height * 4 : height * 0.45;
      region = { x0: bbox.x0 - width * 0.12, x1: bbox.x1 + width * 0.12, y0: bbox.y0 - verticalMargin, y1: bbox.y1 + verticalMargin };
    } else if (labels.length >= 2) {
      const bbox = unionBox(labels.map((label) => ({ bbox: label.bbox })));
      const width = bbox.x1 - bbox.x0, height = median(labels.map((label) => boxHeight(label.bbox)));
      region = { x0: bbox.x0 - height, x1: bbox.x1 + width * 2.5, y0: bbox.y0 - height, y1: bbox.y1 + height };
    }
    if (!region) return cropCanvas(canvas, 0.05, 0.4, 0.9, 0.6);
    const x = Math.max(0, Math.floor(region.x0)), y = Math.max(0, Math.floor(region.y0));
    const width = Math.min(canvas.width, Math.ceil(region.x1)) - x, height = Math.min(canvas.height, Math.ceil(region.y1)) - y;
    if (width < 1 || height < 1) return cropCanvas(canvas, 0.05, 0.4, 0.9, 0.6);
    const factor = nativeSize ? 1 : dates.length >= 2 ? Math.min(2, 1600 / Math.max(width, height)) : Math.min(5, 2200 / Math.max(width, height), Math.max(1, 1600 / Math.max(width, height)));
    const crop = document.createElement('canvas'); crop.width = Math.round(width * factor); crop.height = Math.round(height * factor);
    crop.getContext('2d').drawImage(canvas, x, y, width, height, 0, 0, crop.width, crop.height);
    return crop;
  }
  function boundedCrop(canvas, region, nativeSize, factor) {
    const source = nativeSize && canvas.nativeCanvas ? canvas.nativeCanvas : canvas;
    const sx = source.width / canvas.width, sy = source.height / canvas.height;
    const x = Math.max(0, Math.floor(region.x0 * sx)), y = Math.max(0, Math.floor(region.y0 * sy));
    const width = Math.min(source.width, Math.ceil(region.x1 * sx)) - x, height = Math.min(source.height, Math.ceil(region.y1 * sy)) - y;
    if (width < 1 || height < 1) return null;
    const scale = Math.min(factor || 1, 2200 / Math.max(width, height));
    const crop = document.createElement('canvas'); crop.width = Math.max(1, Math.round(width * scale)); crop.height = Math.max(1, Math.round(height * scale));
    crop.getContext('2d').drawImage(source, x, y, width, height, 0, 0, crop.width, crop.height);
    return crop;
  }
  function dateProbe(canvas, pass) {
    const valid = spatialDates(pass), yearWords = pass.words.filter((word) => word.bbox && /^20\d{2}(?:\s*[.\/\-년]\s*\d{0,2}(?:\s*[.\/\-월]\s*\d{0,2})?)?\.?$/.test(word.text.trim()));
    if (valid.length || yearWords.length) {
      const boxes = valid.length ? valid.map((item) => ({ bbox: item.bbox })) : yearWords;
      const bbox = unionBox(boxes), width = median(boxes.map((item) => item.bbox.x1 - item.bbox.x0)), height = median(boxes.map((item) => boxHeight(item.bbox)));
      // A year or year/month fragment is a location hint, not a full date
      // width. Keep enough pixels to its right for the missing month/day.
      return boundedCrop(canvas, { x0: bbox.x0 - width * 0.15, x1: valid.length ? bbox.x1 + width * 0.15 : Math.max(bbox.x1 + width * 0.15, bbox.x0 + Math.max(width * 2.5, height * 12)), y0: bbox.y0 - height * 4, y1: bbox.y1 + height * 4 }, true);
    }
    const hints = pass.words.filter((word) => word.bbox && /calib?rat|al[i1]?brat|교\s*정|예\s*정|due|next/i.test(word.text));
    if (hints.length) {
      const bbox = unionBox(hints), width = bbox.x1 - bbox.x0, height = median(hints.map((word) => boxHeight(word.bbox)));
      return boundedCrop(canvas, { x0: bbox.x0 - width * 0.2, x1: Math.max(bbox.x1 + width * 1.5, bbox.x0 + height * 16), y0: bbox.y0 - height * 0.5, y1: bbox.y1 + Math.max(height * 10, (bbox.y1 - bbox.y0) * 2) }, true);
    }
    return cropCanvas(canvas.nativeCanvas || canvas, 0.05, 0.25, 0.9, 0.7);
  }
  function dateColumnFromLabels(canvas, pass) {
    // A tiny table may retain only "일 자" from each left-hand label while
    // suppressing the numbers as table cells. Use those row locations solely
    // to propose the right column; the numbers still need two clean OCR reads.
    const rows = physicalRows(pass.words.filter((word) => word.bbox && /[가-힣]/.test(word.text)))
      .filter((row) => /일자/.test(cleanName(row.map((word) => word.text).join(''))));
    if (rows.length < 2 || rows.length > 5) return null;
    const boxes = rows.map(unionBox), h = median(rows.flat().map((word) => boxHeight(word.bbox)));
    const right = median(boxes.map((box) => box.x1));
    const crop = boundedCrop(canvas, { x0: right + h * 0.2, x1: canvas.width, y0: Math.min(...boxes.map((box) => box.y0)) - h * 0.5, y1: Math.max(...boxes.map((box) => box.y1)) + h * 0.5 }, true);
    return crop ? withoutVerticalBorders(crop, h) : null;
  }
  function withoutVerticalBorders(canvas, characterHeight) {
    // Closed spreadsheet-like cells can cause Tesseract to suppress their
    // contents. Trim only long vertical rules near the crop's outer edges.
    // Pixel values and recognized characters are never rewritten.
    try {
      const width = canvas.width, height = canvas.height;
      const pixels = canvas.getContext('2d').getImageData(0, 0, width, height).data;
      const radius = Math.max(1, Math.min(6, Math.round(characterHeight * 0.1)));
      const occupied = new Uint32Array(width), longest = new Uint32Array(width), current = new Uint32Array(width), prefix = new Uint32Array(width + 1);
      for (let y = 0; y < height; y++) {
        prefix[0] = 0;
        for (let x = 0; x < width; x++) {
          const index = (y * width + x) * 4;
          const dark = pixels[index + 3] > 128 && (pixels[index] + pixels[index + 1] + pixels[index + 2]) / 3 < 140;
          prefix[x + 1] = prefix[x] + Number(dark);
        }
        for (let x = 0; x < width; x++) {
          const dark = prefix[Math.min(width, x + radius + 1)] > prefix[Math.max(0, x - radius)];
          if (dark) { occupied[x]++; current[x]++; longest[x] = Math.max(longest[x], current[x]); } else current[x] = 0;
        }
      }
      const runs = [];
      for (let x = 0; x < width; x++) {
        if (occupied[x] < height * 0.72 || longest[x] < height * 0.65) continue;
        const previous = runs[runs.length - 1];
        if (previous && previous.end === x - 1) previous.end = x; else runs.push({ start: x, end: x });
      }
      const narrow = runs.filter((run) => run.end - run.start + 1 <= characterHeight * 0.6);
      const left = narrow.find((run) => run.end < Math.min(width * 0.35, characterHeight * 2));
      const right = narrow.find((run) => run.start > Math.max(width * 0.5, width - characterHeight * 3));
      const padding = Math.max(1, Math.round(characterHeight * 0.1));
      const x0 = left ? left.end + padding : 0, x1 = right ? right.start - padding : width;
      if ((x0 || x1 !== width) && x1 - x0 >= Math.max(width * 0.5, characterHeight * 4)) return boundedCrop(canvas, { x0, x1, y0: 0, y1: height }, false);
    } catch (_) { /* Keep the original crop if pixels are unavailable. */ }
    return canvas;
  }
  function identityCrop(canvas, pass, tools) {
    const numeric = pass.words.filter((word) => word.bbox && /^[0-9][0-9 ._-]{5,}$/.test(word.text) && !dateOccurrences(word.text).length && !/^20\d{6}$/.test(cleanSerial(word.text)));
    const names = tools.filter((tool) => hasIdentity(pass.text, tool.name, false)).map((tool) => cleanName(tool.name));
    const named = pass.words.filter((word) => word.bbox && names.some((name) => name.includes(cleanName(word.text)) && /[가-힣A-Za-z]/.test(word.text)));
    const serialLabels = pass.words.filter((word) => word.bbox && /serial|s\s*\/\s*n|시리얼/i.test(word.text));
    const besideSerialLabel = (word) => serialLabels.some((label) => word.bbox.x0 >= label.bbox.x1 && Math.abs(centerY(word.bbox) - centerY(label.bbox)) <= Math.max(boxHeight(word.bbox), boxHeight(label.bbox)) * 1.5);
    if (numeric.length && (!named.length || numeric.some(besideSerialLabel))) {
      const word = numeric.slice().sort((a, b) => Number(besideSerialLabel(b)) - Number(besideSerialLabel(a)) || (b.confidence || 0) - (a.confidence || 0))[0], box = word.bbox, width = box.x1 - box.x0, height = boxHeight(box);
      return { canvas: boundedCrop(canvas, { x0: box.x0 - width * 0.25, x1: box.x1 + width * 0.25, y0: box.y0 - height * 0.4, y1: box.y1 + height * 0.4 }, true, 1.5), psm: '7' };
    }
    if (named.length) {
      const box = unionBox(named), width = box.x1 - box.x0, height = boxHeight(box);
      return { canvas: boundedCrop(canvas, { x0: box.x0 - width * 0.3, x1: box.x1 + width * 3, y0: box.y0 - height, y1: box.y1 + height * 5 }, true, 1.5), psm: '6' };
    }
    return { canvas: cropCanvas(canvas, 0.03, 0.02, 0.94, 0.5), psm: '6' };
  }
  function identityPanels(canvas) {
    const native = canvas.nativeCanvas || canvas, regions = [];
    // White paper tags are location proposals only. Their pixels must still
    // produce exact, confident OCR; no workbook text is painted into a crop.
    try {
      const scale = Math.min(1, 640 / native.width), width = Math.max(1, Math.round(native.width * scale)), height = Math.max(1, Math.round(native.height * 0.74 * scale));
      const sample = document.createElement('canvas'); sample.width = width; sample.height = height;
      const context = sample.getContext('2d', { willReadFrequently: true }); context.drawImage(native, 0, 0, native.width, native.height * 0.74, 0, 0, width, height);
      const data = context.getImageData(0, 0, width, height).data, mask = new Uint8Array(width * height), pending = new Int32Array(width * height);
      for (let index = 0; index < mask.length; index++) mask[index] = data[index * 4 + 3] > 128 && (data[index * 4] + data[index * 4 + 1] + data[index * 4 + 2]) / 3 >= 172 ? 1 : 0;
      for (let start = 0; start < mask.length; start++) {
        if (!mask[start]) continue;
        let head = 0, tail = 1, count = 0, x0 = width, y0 = height, x1 = 0, y1 = 0; pending[0] = start; mask[start] = 0;
        while (head < tail) {
          const index = pending[head++], x = index % width, y = Math.floor(index / width); count++;
          x0 = Math.min(x0, x); x1 = Math.max(x1, x); y0 = Math.min(y0, y); y1 = Math.max(y1, y);
          if (x && mask[index - 1]) { mask[index - 1] = 0; pending[tail++] = index - 1; }
          if (x + 1 < width && mask[index + 1]) { mask[index + 1] = 0; pending[tail++] = index + 1; }
          if (y && mask[index - width]) { mask[index - width] = 0; pending[tail++] = index - width; }
          if (y + 1 < height && mask[index + width]) { mask[index + width] = 0; pending[tail++] = index + width; }
        }
        const w = x1 - x0 + 1, h = y1 - y0 + 1;
        if (count < 24 || w < 12 || w / scale < 64 || h < 3 || w / h < 1.3 || w / h > 35 || count < w * h * 0.3 || w * h > width * height * 0.35 || h > height * 0.3) continue;
        const pad = Math.max(2, Math.min(h * 0.1, 5));
        regions.push({ x0: (x0 - pad) / scale, x1: (x1 + 1 + pad) / scale, y0: (y0 - pad) / scale, y1: (y1 + 1 + pad) / scale, area: w * h });
      }
    } catch (_) { /* Overlapping upper bands remain available without pixels. */ }
    const panels = regions.sort((a, b) => b.area - a.area).filter((region, index, sorted) => !sorted.slice(0, index).some((other) => region.x0 >= other.x0 && region.x1 <= other.x1 && region.y0 >= other.y0 && region.y1 <= other.y1)).slice(0, 2).sort((a, b) => a.y0 - b.y0);
    // A detected white component can be a letter hole or a reflection. Reserve
    // independent coverage even when such a component was found.
    [[0.01, 0.42], [0.34, 0.74]].forEach(([top, bottom]) => panels.push({ x0: native.width * 0.03, x1: native.width * 0.97, y0: native.height * top, y1: native.height * bottom }));
    return panels.map((region) => ({ canvas: boundedCrop(native, region, false), region })).filter((item) => item.canvas);
  }
  function identityLines(canvas) {
    const lines = [];
    try {
      const width = canvas.width, height = canvas.height, data = canvas.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, width, height).data;
      const ranges = []; let active = null;
      for (let y = 1; y < height - 1; y++) {
        let count = 0, left = width, right = 0;
        for (let x = 1; x < width - 1; x++) {
          const index = (y * width + x) * 4;
          if ((data[index] + data[index + 1] + data[index + 2]) / 3 < 130) { count++; left = Math.min(left, x); right = Math.max(right, x); }
        }
        if (count >= Math.max(2, width * 0.006)) {
          if (active && y - active.y1 <= Math.max(2, height * 0.025)) { active.y1 = y; active.x0 = Math.min(active.x0, left); active.x1 = Math.max(active.x1, right); }
          else { active = { x0: left, x1: right, y0: y, y1: y }; ranges.push(active); }
        }
      }
      ranges.filter((range) => range.y1 - range.y0 >= 3 && range.x1 - range.x0 >= 10).sort((a, b) => b.x1 - b.x0 - (a.x1 - a.x0)).slice(0, 2).forEach((range) => {
        const h = range.y1 - range.y0 + 1;
        lines.push({ x0: range.x0 - h * 0.4, x1: range.x1 + h * 0.4, y0: range.y0 - h * 0.35, y1: range.y1 + h * 0.35 });
      });
    } catch (_) {}
    return lines.length ? lines : [{ x0: 0, x1: canvas.width, y0: 0, y1: canvas.height }];
  }
  async function runRecognition(blob, tools, onProgress) {
    progressListener = onProgress; progressStart = 0.15; progressSpan = 0.64;
    const cancellation = { cancelled: false, reject: null, promise: null };
    cancellation.promise = new Promise((_, reject) => { cancellation.reject = reject; });
    cancellation.promise.catch(() => {});
    activeCancellation = cancellation;
    const began = Date.now(), passes = [], recoveryWarnings = [];
    let stage = 'image-load';
    const heartbeat = setInterval(() => {
      if (typeof progressListener === 'function') {
        try { progressListener({ status: currentPhase + ' · ' + Math.floor((Date.now() - began) / 1000) + '초', progress: currentProgress }); } catch (_) {}
      }
    }, 2000);
    function record(response, source, metadata) {
      const data = response.data || {};
      const words = Array.isArray(data.words) ? data.words : (data.blocks || []).flatMap((block) => (block.paragraphs || []).flatMap((paragraph) => (paragraph.lines || []).flatMap((line) => line.words || [])));
      passes.push(Object.assign({ text: data.text || '', confidence: data.confidence, source, words: words.map(normalizeWord) }, metadata || {}));
    }
    const current = () => parse(passes.map((pass) => pass.text).join('\n'), tools, { passes });
    const strongIdentity = (result) => {
      const serials = unique(result.identityEvidence.serialMatches.filter((match) => match.confidence != null && match.confidence >= 70).map((match) => match.toolId));
      const names = unique(result.identityEvidence.nameMatches.filter((match) => match.confidence != null && match.confidence >= 70).map((match) => match.toolId));
      return serials.length === 1 || (!serials.length && names.length === 1);
    };
    try {
      emit('사진을 준비하는 중', 0.01);
      const canvas = await deadline(imageCanvas(blob), 15000, '사진을 15초 안에 열지 못했습니다. 다른 JPG 또는 PNG 사진으로 다시 시도해 주세요.');
      stage = 'engine-init';
      const worker = await getWorker();
      stage = 'recognition';
      await deadline(worker.setParameters({ tessedit_pageseg_mode: '11', preserve_interword_spaces: '1' }), 5000, '문자 인식 설정을 적용하지 못했습니다. 다시 시도해 주세요.');
      record(await deadline(worker.recognize(canvas, {}, { text: true, blocks: true }), 60000, '사진을 60초 안에 읽지 못했습니다. 라벨을 크게 촬영한 사진으로 다시 시도해 주세요.'), 'original');
      let preliminary = current();
      if (preliminary.documentType === 'smax' || preliminary.documentType === 'unknown-smax') {
        const beganFields = Date.now(), fieldBudget = 60000; let fieldCalls = 0;
        const markIncomplete = () => { if (!passes.some((pass) => pass.source === 'smax-field-incomplete')) passes.push({ text: '', confidence: null, source: 'smax-field-incomplete', words: [] }); };
        try {
          const proposed = global.ToolSMAX && typeof global.ToolSMAX.regionsForPass === 'function' ? global.ToolSMAX.regionsForPass(passes[0], canvas.width, canvas.height) : [];
          const regions = (Array.isArray(proposed) ? proposed : []).filter((region) => ['serial', 'owner', 'due'].includes(region.role) && region.labelVerified === true && wordBox(region.bbox) && wordBox(region.labelBBox) && typeof region.labelText === 'string' && region.labelText.trim())
            .sort((a, b) => ['serial', 'owner', 'due'].indexOf(a.role) - ['serial', 'owner', 'due'].indexOf(b.role));
          const serialAlreadyRead = (preliminary.smaxEvidence && preliminary.smaxEvidence.serialReadings || []).some((reading) => reading.labelVerified === true && score(reading.confidence) >= 70 && tools.some((tool) => cleanSerial(tool.serial) === cleanSerial(reading.value)));
          for (let index = 0; index < regions.length; index++) {
            const region = regions[index];
            if (region.role === 'serial' && serialAlreadyRead) continue;
            const input = boundedCrop(canvas, region.bbox, true, 1.5);
            if (!input) continue;
            for (const psm of region.role === 'owner' ? ['6'] : region.role === 'serial' ? ['7', '11'] : ['6', '11']) {
              if (fieldCalls >= 6 || Date.now() - beganFields >= fieldBudget) { markIncomplete(); break; }
              fieldCalls++;
              emit('SMAX의 ' + ({ serial: 'Serial/Asset Number', owner: 'Tool Owner', due: 'Calibration Due Date' }[region.role]) + ' 필드를 확인하는 중 · ' + fieldCalls + '/6', 0.82 + fieldCalls * 0.025);
              await deadline(worker.setParameters({ tessedit_pageseg_mode: psm }), Math.min(5000, fieldBudget - (Date.now() - beganFields)), 'SMAX 필드 인식 설정 시간이 초과되었습니다.');
              const metadata = { smaxRole: region.role, smaxLabelVerified: true, smaxLabelText: region.labelText, smaxLabelBBox: wordBox(region.labelBBox), smaxValueBBox: wordBox(region.bbox) };
              record(await deadline(worker.recognize(input, {}, { text: true, blocks: true }), Math.min(20000, Math.max(1, fieldBudget - (Date.now() - beganFields))), 'SMAX 필드 추가 확인 시간이 초과되었습니다.'), 'smax-field-' + region.role + '-' + psm + '-' + index, metadata);
            }
            if (fieldCalls >= 6 || Date.now() - beganFields >= fieldBudget) { if (index + 1 < regions.length) markIncomplete(); break; }
          }
        } catch (error) {
          if (cancellation.cancelled) throw error;
          markIncomplete(); discardWorker();
          recoveryWarnings.push('SMAX 필드를 모두 확인하지 못해 자동 반영을 보류했습니다. 다시 올려 주세요.');
        }
        const result = current(); result.engineFailure = false; result.code = ''; result.errorCode = '';
        result.warnings.push(...recoveryWarnings);
        emit('SMAX 필드 인식 완료', 1); return result;
      }
      if (!strongIdentity(preliminary) || preliminary.dateEvidence.method !== 'labelled' || !(preliminary.dateEvidence.confidence >= 70)) {
        progressStart = 0.8; progressSpan = 0.1;
        emit('교정 라벨을 한 번 더 읽는 중', progressStart);
        if (spatialDates(passes[0]).length >= 2 || preliminary.dateEvidence.method === 'labelled' && preliminary.dateEvidence.confidence >= 70) {
          await deadline(worker.setParameters({ tessedit_pageseg_mode: '6' }), 5000, '문자 인식 설정에 실패했습니다.');
          record(await deadline(worker.recognize(dateCrop(canvas, passes[0]), {}, { text: true, blocks: true }), 45000, '날짜 추가 인식 시간이 초과되었습니다. 날짜를 확인해 주세요.'), 'date-crop');
        } else {
          // Discovery scans locate small date lines; they are not proof of a
          // date. Only the subsequent clean, independently segmented re-reads
          // can supply automatic dates, so an exploratory 06/08 error cannot
          // silently become an approved due date.
          const probe = dateProbe(canvas, passes[0]);
          await deadline(worker.setParameters({ tessedit_pageseg_mode: '11' }), 5000, '문자 인식 설정에 실패했습니다.');
          record(await deadline(worker.recognize(probe, {}, { text: true, blocks: true }), 45000, '작은 교정 라벨 탐색 시간이 초과되었습니다.'), 'date-probe');
          let discovered = passes[passes.length - 1];
          let focused = dateColumnFromLabels(probe, discovered) || (spatialDates(discovered).length ? dateCrop(probe, discovered, true) : null);
          if (!focused) {
            await deadline(worker.setParameters({ tessedit_pageseg_mode: '6' }), 5000, '문자 인식 설정에 실패했습니다.');
            record(await deadline(worker.recognize(probe, {}, { text: true, blocks: true }), 45000, '작은 교정 표 재탐색 시간이 초과되었습니다.'), 'date-probe-lines');
            discovered = passes[passes.length - 1];
            focused = dateColumnFromLabels(probe, discovered) || (spatialDates(discovered).length ? dateCrop(probe, discovered, true) : null);
          }
          if (focused) {
            for (const psm of ['6', '11']) {
              emit('찾은 날짜를 원본 크기로 확인하는 중', 0.86);
              await deadline(worker.setParameters({ tessedit_pageseg_mode: psm }), 5000, '문자 인식 설정에 실패했습니다.');
              record(await deadline(worker.recognize(focused, {}, { text: true, blocks: true }), 45000, '확대 날짜 확인 시간이 초과되었습니다.'), 'date-crop-refine-' + psm);
            }
          }
        }
        preliminary = current();
        const dateEvidence = preliminary.dateEvidence;
        const clearDatePair = dateEvidence.method && dateEvidence.last && dateEvidence.due && dateEvidence.confidence >= 70;
        const clearLastOnly = dateEvidence.explicitLast && dateEvidence.lastConfidence >= 70 && !dateEvidence.dueLabelSeen && !dateEvidence.explicitDue && !dateEvidence.ambiguous && dateEvidence.dates.length === 1;
        if (!clearDatePair && !clearLastOnly) {
          // An isolated year/header can point at the wrong place or crop away
          // the day. Search an independent wider area once; its locator text
          // is excluded from approval just like the first date probe.
          emit('날짜 라벨의 넓은 영역을 다시 찾는 중', 0.89);
          const wide = cropCanvas(canvas.nativeCanvas || canvas, 0.03, 0.18, 0.94, 0.8);
          await deadline(worker.setParameters({ tessedit_pageseg_mode: '6' }), 5000, '문자 인식 설정에 실패했습니다.');
          record(await deadline(worker.recognize(wide, {}, { text: true, blocks: true }), 20000, '넓은 날짜 라벨 탐색 시간이 초과되었습니다.'), 'date-probe-wide');
          const discovered = passes[passes.length - 1];
          const focused = dateColumnFromLabels(wide, discovered) || (spatialDates(discovered).length ? dateCrop(wide, discovered, true) : null);
          if (focused) {
            for (const psm of ['6', '11']) {
              await deadline(worker.setParameters({ tessedit_pageseg_mode: psm }), 5000, '문자 인식 설정에 실패했습니다.');
              record(await deadline(worker.recognize(focused, {}, { text: true, blocks: true }), 20000, '찾은 교정 날짜 확인 시간이 초과되었습니다.'), 'date-crop-wide-refine-' + psm);
            }
          }
          preliminary = current();
        }
      }
      if (!strongIdentity(preliminary)) {
        progressStart = 0.91; progressSpan = 0.08;
        emit('이름과 시리얼 표지를 한 번 더 읽는 중', progressStart);
        const identity = identityCrop(canvas, passes[0], tools);
        await deadline(worker.setParameters({ tessedit_pageseg_mode: identity.psm }), 5000, '문자 인식 설정에 실패했습니다.');
        record(await deadline(worker.recognize(identity.canvas, {}, { text: true, blocks: true }), 45000, '이름·시리얼 추가 인식 시간이 초과되었습니다.'), 'identity-crop');
        preliminary = current();
        if (!strongIdentity(preliminary)) {
          emit('사진 하단의 시리얼 표지를 읽는 중', 0.97);
          await deadline(worker.setParameters({ tessedit_pageseg_mode: '6' }), 5000, '문자 인식 설정에 실패했습니다.');
          record(await deadline(worker.recognize(cropCanvas(canvas, 0.05, 0.4, 0.9, 0.6), {}, { text: true, blocks: true }), 45000, '하단 시리얼 추가 인식 시간이 초과되었습니다.'), 'identity-crop-lower');
        }
        if (!strongIdentity(current())) {
          const recoveryStart = Date.now(), budget = 90000, panels = identityPanels(canvas); let calls = 0;
          const readTag = async (input, psm, source) => {
            if (calls >= 10 || Date.now() - recoveryStart >= budget) return false;
            const remaining = budget - (Date.now() - recoveryStart); calls++;
            emit('상단 이름·시리얼 표지를 각각 확인하는 중 · ' + calls + '/10', 0.97);
            await deadline(worker.setParameters({ tessedit_pageseg_mode: psm }), Math.min(5000, remaining), '표지 인식 설정 시간이 초과되었습니다.');
            record(await deadline(worker.recognize(input, {}, { text: true, blocks: true }), Math.min(20000, Math.max(1, budget - (Date.now() - recoveryStart))), '상단 표지 추가 확인 시간이 초과되었습니다.'), source);
            return true;
          };
          try {
            // Read each separate paper tag once before stopping on a match;
            // this also retains conflicts between two actual known serials.
            for (let index = 0; index < panels.length; index++) {
              const panel = panels[index].canvas;
              if (!await readTag(boundedCrop(panel, { x0: 0, y0: 0, x1: panel.width, y1: panel.height }, false, 1.5), '6', 'identity-probe-panel-' + index)) {
                passes.push({ text: '', confidence: null, source: 'identity-probe-incomplete', words: [] });
                recoveryWarnings.push('표지 후보를 모두 확인하지 못해 자동 반영을 보류했습니다.');
                break;
              }
            }
            if (!strongIdentity(current())) {
              const lineCandidates = panels.flatMap((panel, index) => identityLines(panel.canvas).map((region, line) => ({ canvas: panel.canvas, region, index, line })));
              for (const candidate of lineCandidates) {
                for (const variant of [{ psm: '7', scale: 1.5 }, { psm: '13', scale: 2 }]) {
                  if (!await readTag(boundedCrop(candidate.canvas, candidate.region, false, variant.scale), variant.psm, 'identity-probe-line-' + candidate.index + '-' + candidate.line + '-' + variant.psm)) break;
                  if (strongIdentity(current())) break;
                }
                if (strongIdentity(current()) || calls >= 10 || Date.now() - recoveryStart >= budget) break;
              }
            }
          } catch (error) {
            if (cancellation.cancelled) throw error;
            // Preserve successful date reads when an optional tag retry runs
            // out of time. It remains unapproved unless identity is proven.
            recoveryWarnings.push('표지 추가 확인을 끝내지 못했습니다. 이름·시리얼 라벨을 크게 찍어 다시 올려 주세요.');
            passes.push({ text: '', confidence: null, source: 'identity-probe-incomplete', words: [] });
            discardWorker();
          }
        }
      }
      const result = current(); result.engineFailure = false; result.code = ''; result.errorCode = '';
      result.warnings.push(...recoveryWarnings);
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
