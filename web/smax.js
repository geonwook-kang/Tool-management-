/* SMAX fields are read by their labels. Workbook identity remains authoritative. */
(function (global) {
  'use strict';
  const serialKey = value => String(value == null ? '' : value).normalize('NFKC').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const high = value => value != null && Number.isFinite(Number(value)) && Number(value) >= 70 && Number(value) <= 100;
  const score = value => value != null && value !== '' && Number.isFinite(Number(value)) && Number(value) >= 0 && Number(value) <= 100 ? Number(value) : null;
  const unique = values => Array.from(new Set(values));
  const definitions = [
    { role: 'serial', source: '\\bserial\\s*(?:[/|iIl1]\\s*)?asset\\s*(?:number|no\\.?)\\b' },
    { role: 'owner', source: '\\btool\\s*owner\\b' },
    { role: 'due', source: '\\bcalibration\\s*due\\s*date\\b' },
    { role: 'structure', source: '\\b(?:product|tool\\s*details?)\\b' },
    { role: 'other', source: '\\b(?:SGPN|tool\\s*code|(?:last\\s*)?(?:created|modified|updated)(?:\\s*(?:by|on|date))?|serial\\s*number|asset\\s*number)\\b' }
  ];
  const canonicalPattern = role => new RegExp('^(?:' + definitions.find(item => item.role === role).source + ')[:\\s]*$', 'i');
  function box(value) {
    if (!value) return null;
    const result = Object.fromEntries(['x0', 'y0', 'x1', 'y1'].map(key => [key, Number(value[key])]));
    return Object.values(result).every(Number.isFinite) && result.x1 > result.x0 && result.y1 > result.y0 ? result : null;
  }
  const height = value => value.y1 - value.y0, cy = value => (value.y0 + value.y1) / 2;
  const union = words => ({ x0: Math.min(...words.map(w => w.bbox.x0)), y0: Math.min(...words.map(w => w.bbox.y0)), x1: Math.max(...words.map(w => w.bbox.x1)), y1: Math.max(...words.map(w => w.bbox.y1)) });
  const confidence = words => words.length && words.every(w => score(w.confidence) != null) ? Math.min(...words.map(w => score(w.confidence))) : null;
  function joined(words) {
    let text = '';
    const spans = words.map(word => { if (text) text += ' '; const start = text.length; text += word.text; return { word, start, end: text.length }; });
    return { text, spans };
  }
  function labels(text) {
    const result = [];
    for (const definition of definitions) {
      const regex = new RegExp(definition.source, 'gi'); let match;
      while ((match = regex.exec(text))) {
        const start = match.index, end = start + match[0].length;
        if (!result.some(item => start < item.end && end > item.start)) result.push({ role: definition.role, text: match[0], start, end });
      }
    }
    return result.sort((a, b) => a.start - b.start);
  }
  function rows(words) {
    const result = [];
    for (const word of words.slice().sort((a, b) => cy(a.bbox) - cy(b.bbox) || a.bbox.x0 - b.bbox.x0)) {
      const target = result.find(row => Math.abs(cy(word.bbox) - cy(row[0].bbox)) <= Math.min(height(word.bbox), height(row[0].bbox)) * 0.6);
      if (target) target.push(word); else result.push([word]);
    }
    return result.map(row => row.sort((a, b) => a.bbox.x0 - b.bbox.x0));
  }
  function geometry(pass) {
    const words = (pass.words || []).filter(w => box(w.bbox)).map(w => ({ ...w, bbox: box(w.bbox) })), result = [];
    for (const row of rows(words)) {
      const text = joined(row), found = labels(text.text).map(label => ({ ...label, words: text.spans.filter(item => item.start < label.end && item.end > label.start).map(item => item.word) }));
      for (const label of found) {
        if (!['serial', 'owner', 'due'].includes(label.role) || !label.words.length) continue;
        const labelBBox = union(label.words), next = found.filter(other => other.start >= label.end && other.words.length).map(other => union(other.words).x0);
        const right = next.length ? Math.min(...next) : Infinity;
        const values = row.filter(w => w.bbox.x0 >= labelBBox.x1 - 1 && w.bbox.x1 <= right && !label.words.includes(w));
        result.push({ role: label.role, labelText: label.text, labelBBox, right, valueWords: values, valueBBox: values.length ? union(values) : null });
      }
    }
    return result;
  }
  function dateValue(raw) {
    const value = String(raw || '').trim().replace(/\s*([.\/-])\s*/g, '$1');
    let match = /^(20\d{2})[.\/-](\d{1,2})[.\/-](\d{1,2})$/.exec(value), y, m, d;
    if (match) [, y, m, d] = match;
    else { match = /^(\d{1,2})\/(\d{1,2})\/(20\d{2})$/.exec(value); if (!match) return ''; [, m, d, y] = match; }
    const date = new Date(Date.UTC(Number(y), Number(m) - 1, Number(d)));
    return date.getUTCFullYear() === Number(y) && date.getUTCMonth() + 1 === Number(m) && date.getUTCDate() === Number(d) ? [y, String(m).padStart(2, '0'), String(d).padStart(2, '0')].join('-') : '';
  }
  const iso = value => /^20\d{2}-\d{2}-\d{2}$/.test(String(value || '')) && dateValue(value) === value;
  function valuesFor(role, raw) {
    raw = String(raw || '');
    if (role === 'serial') raw = raw.replace(/^[\s©®ⓘℹ\uFE0F@]+/u, '');
    raw = raw.normalize('NFKC').replace(/^[:#\s]+/, '').trim();
    if (!raw) return [];
    if (role === 'due') {
      const found = Array.from(raw.matchAll(/(?:20\d{2}\s*[.\/-]\s*\d{1,2}\s*[.\/-]\s*\d{1,2}|\d{1,2}\s*\/\s*\d{1,2}\s*\/\s*20\d{2})/g)).filter(match => !/[\dA-Za-z]/.test(raw[match.index - 1] || '') && !/[\dA-Za-z]/.test(raw[match.index + match[0].length] || '')).map(match => match[0]);
      return found.length ? found.map(text => ({ raw: text, value: dateValue(text) })) : [{ raw: raw.slice(0, 120), value: '' }];
    }
    const first = raw.split(/\r?\n/).map(line => line.trim()).find(Boolean) || '';
    if (role === 'serial') return [{ raw: first.slice(0, 120), value: /^[A-Za-z0-9][A-Za-z0-9._\- \t]{1,100}$/.test(first) ? serialKey(first) : '' }];
    return [{ raw: first.slice(0, 120), value: first.replace(/\s+/g, ' ').slice(0, 120) }];
  }
  function analyze(passes, tools) {
    passes = Array.isArray(passes) ? passes : [];
    const evidence = { serialReadings: [], dueReadings: [], ownerEnglish: [], labels: [], incomplete: false }, observed = [];
    function add(role, raw, pass, certainty, bbox, labelText, labelBBox) {
      const list = role === 'serial' ? evidence.serialReadings : role === 'due' ? evidence.dueReadings : evidence.ownerEnglish;
      for (const item of valuesFor(role, raw)) {
        const reading = { ...item, role, source: String(pass.source || 'text'), confidence: score(certainty), bbox: box(bbox), labelText, labelBBox: box(labelBBox), labelVerified: true };
        if (!list.some(old => old.value === reading.value && old.raw === reading.raw && old.source === reading.source && old.confidence === reading.confidence)) list.push(reading);
      }
    }
    for (const pass of passes) {
      const source = String(pass.source || '');
      if (/^smax-field-incomplete/i.test(source) || pass.smaxIncomplete) evidence.incomplete = true;
      if (/^(?:date|label)[-_]probe/i.test(source) || /(?:^|[-_])incomplete(?:[-_]|$)/i.test(source) || pass.smaxIncomplete) continue;
      const text = String(pass.text || '').normalize('NFKC'), allLabels = labels(text);
      observed.push(text); evidence.labels.push(...allLabels.map(label => ({ role: label.role, text: label.text, source: String(pass.source || 'text') })));
      const role = String(pass.smaxRole || ''), metadata = /^smax-field-(serial|owner|due)(?:-|$)/i.exec(String(pass.source || ''));
      if (metadata && role === metadata[1].toLowerCase() && pass.smaxLabelVerified === true && box(pass.smaxLabelBBox) && canonicalPattern(role).test(String(pass.smaxLabelText || '').trim())) {
        evidence.labels.push({ role, text: pass.smaxLabelText, source: pass.source });
        const bodyLabels = labels(text), matching = bodyLabels.find(label => label.role === role);
        if (!matching && bodyLabels.length) continue;
        const raw = matching ? text.slice(matching.end, bodyLabels.find(label => label.start >= matching.end)?.start) : text;
        const words = (pass.words || []).filter(w => !matching || !canonicalPattern(role).test(String(w.text || '')));
        add(role, raw, pass, words.length ? confidence(words) : score(pass.confidence), pass.smaxValueBBox, pass.smaxLabelText, pass.smaxLabelBBox);
        continue;
      }
      const geometric = geometry(pass);
      for (const field of geometric) if (field.valueWords.length) add(field.role, field.valueWords.map(w => w.text).join(' '), pass, confidence(field.valueWords), field.valueBBox, field.labelText, field.labelBBox);
      // Text-only input may prove a field. With word boxes, text fallback has
      // no independent confidence and cannot override its geometric reading.
      for (let index = 0; index < allLabels.length; index++) {
        const label = allLabels[index]; if (!['serial', 'owner', 'due'].includes(label.role)) continue;
        const raw = text.slice(label.end, index + 1 < allLabels.length ? allLabels[index + 1].start : label.end + 180);
        const wordData = (pass.words || []);
        add(label.role, raw, pass, wordData.length ? null : score(pass.confidence), null, label.text, null);
      }
    }
    const serialLabel = evidence.labels.some(label => label.role === 'serial'), structure = evidence.labels.some(label => ['owner', 'due', 'structure'].includes(label.role));
    const owner = evidence.labels.some(label => label.role === 'owner'), due = evidence.labels.some(label => label.role === 'due');
    const anchors = unique(evidence.labels.map(label => /\bSGPN\b/i.test(label.text) ? 'sgpn' : /\btool\s*code\b/i.test(label.text) ? 'toolcode' : /\bproduct\b/i.test(label.text) ? 'product' : '').filter(Boolean));
    const hint = serialLabel || observed.some(text => /\bsmax\b|\btool\s*(?:details?|information)\b/i.test(text)) || owner && due && anchors.length >= 2;
    const documentType = serialLabel && structure ? 'smax' : hint ? 'unknown-smax' : 'calibration';
    return { documentType, smaxEvidence: evidence, warnings: documentType === 'unknown-smax' ? ['SMAX 화면의 Serial / Asset Number와 관련 필드 라벨을 확실히 읽지 못해 자동 반영을 보류합니다.'] : [] };
  }
  function regionsForPass(pass, width, heightValue) {
    const widthValue = Number(width), h = Number(heightValue);
    if (!(widthValue > 0 && h > 0)) return [];
    return geometry(pass).map(field => {
      const label = field.labelBBox, rowHeight = height(label), x0 = Math.max(0, label.x1 + rowHeight * 0.1), x1 = Math.min(widthValue, field.right - rowHeight * 0.1, field.valueBBox ? field.valueBBox.x1 + rowHeight : widthValue);
      const bbox = { x0, x1, y0: Math.max(0, label.y0 - rowHeight * 0.35), y1: Math.min(h, label.y1 + rowHeight * 0.35) };
      return { role: field.role, bbox, labelBBox: label, labelText: field.labelText, labelVerified: true };
    }).filter(field => box(field.bbox) && field.bbox.x1 - field.bbox.x0 > 5);
  }
  function assess(model, result) {
    const output = { eligible: false, toolId: '', name: '', serial: '', last: '', due: '', reasons: [], warnings: [], photoKind: 'smax', updateDue: false, matchMethod: 'smax-serial', dateMethod: 'unchanged-dates', ownerEnglish: '' };
    const reject = reason => { if (!output.reasons.includes(reason)) output.reasons.push(reason); };
    const tools = model && Array.isArray(model.tools) ? model.tools : [], data = result && result.smaxEvidence;
    if (!result || result.failed || result.documentType !== 'smax' || !data) { reject('SMAX 필수 필드 라벨과 시리얼을 확인해야 합니다.'); return output; }
    if (data.incomplete || result.identityRecoveryIncomplete) reject('SMAX 필드 추가 인식을 끝내지 못해 자동 반영을 보류합니다.');
    const serials = unique((data.serialReadings || []).filter(r => r.labelVerified && high(r.confidence)).map(r => serialKey(r.value)).filter(Boolean));
    if (serials.length !== 1) { reject(serials.length ? 'SMAX Serial / Asset Number가 서로 다르게 읽혔습니다.' : 'SMAX Serial / Asset Number를 신뢰도 있게 읽지 못했습니다.'); return output; }
    const matches = tools.filter(tool => serialKey(tool.serial) === serials[0]);
    if (matches.length !== 1) { reject('SMAX 시리얼과 정확히 일치하는 툴이 엑셀에서 하나로 정해지지 않습니다.'); return output; }
    const tool = matches[0]; Object.assign(output, { toolId: tool.id, name: tool.name, serial: tool.serial, last: tool.last || '', due: tool.due || '' });
    if (!tool.name || /폐기|disposed|retired/i.test(String(tool.due || ''))) reject('폐기되었거나 담당자가 없는 툴은 자동 변경하지 않습니다.');
    if (tool.hasSmaxPhotoColumn === false) reject('이 시트에 SMAX Tool Photo 열이 없습니다.');
    if (Number(tool.smaxPhotoCount) > 1) reject('SMAX 사진 위치에 여러 사진이 있어 교체 대상을 확정할 수 없습니다.');
    const owners = (data.ownerEnglish || []).slice().sort((a, b) => (b.confidence || 0) - (a.confidence || 0));
    output.ownerEnglish = owners[0]?.value || '';
    const dues = unique((data.dueReadings || []).filter(r => r.labelVerified && high(r.confidence) && iso(r.value)).map(r => r.value));
    if (dues.length > 1) reject('SMAX Calibration Due Date가 서로 다르게 읽혔습니다.');
    else if (dues.length === 1) {
      const due = dues[0]; output.due = due; output.updateDue = true; output.dateMethod = 'smax-due';
      if (iso(tool.last) && due <= tool.last || iso(tool.due) && due < tool.due) reject('SMAX 차기 교정일이 기존 검교정 기록보다 오래되었습니다.');
    } else output.warnings.push('Calibration Due Date를 확실히 읽지 못해 기존 Cal Last와 Cal Next를 보존합니다.');
    output.eligible = !output.reasons.length; return output;
  }
  global.ToolSMAX = Object.freeze({ analyze, assess, regionsForPass });
})(typeof window !== 'undefined' ? window : globalThis);
