/* Automatic row selection uses observed OCR evidence, never workbook guesses. */
(function (global) {
  'use strict';
  const normalizeSerial = (value) => String(value == null ? '' : value).normalize('NFKC').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const normalizeName = (value) => String(value || '').normalize('NFKC').replace(/\s/g, '').toUpperCase();
  const high = (confidence) => confidence != null && Number.isFinite(Number(confidence)) && Number(confidence) >= 70;
  function isoDate(value) {
    if (!/^20\d{2}-\d{2}-\d{2}$/.test(String(value || ''))) return false;
    const parts = value.split('-').map(Number), date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    return date.getUTCFullYear() === parts[0] && date.getUTCMonth() + 1 === parts[1] && date.getUTCDate() === parts[2];
  }
  function addMonths(value, months) {
    if (!isoDate(value) || !Number.isInteger(months) || months <= 0 || months > 1200) return '';
    const parts = value.split('-').map(Number);
    const date = new Date(Date.UTC(parts[0], parts[1] - 1 + months + 1, 0));
    date.setUTCDate(Math.min(parts[2], date.getUTCDate()));
    return date.toISOString().slice(0, 10);
  }
  function monthInterval(tool) {
    if (Number.isInteger(tool.months) && tool.months > 0) return tool.months;
    const interval = Number(tool.interval), cycle = String(tool.cycle || '').trim().toUpperCase();
    if (!Number.isInteger(interval) || interval < 1) return null;
    if (['MONTH', 'MONTHS', '개월', '월'].includes(cycle)) return interval;
    if (['YEAR', 'YEARS', '년'].includes(cycle)) return interval * 12;
    return null;
  }
  function assess(model, result) {
    if (result && /smax/i.test(result.documentType || '')) {
      if (global.ToolSMAX && typeof global.ToolSMAX.assess === 'function') return global.ToolSMAX.assess(model, result);
      return { eligible: false, toolId: '', name: '', serial: '', last: '', due: '', photoKind: 'smax', updateDue: false,
        matchMethod: '', dateMethod: '', reasons: ['SMAX 화면 처리 모듈을 불러오지 못했습니다. 최신 프로그램으로 다시 열어 주세요.'] };
    }
    const output = { eligible: false, toolId: '', name: '', serial: '', last: '', due: '', reasons: [], matchMethod: '', dateMethod: '' };
    const reject = (reason) => { if (!output.reasons.includes(reason)) output.reasons.push(reason); };
    const tools = model && Array.isArray(model.tools) ? model.tools : [];
    if (!tools.length) { reject('관리 엑셀을 먼저 선택해 주세요.'); return output; }
    if (!result || result.failed) { reject('사진 인식을 완료하지 못했습니다.'); return output; }
    const identity = result.identityEvidence, dates = result.dateEvidence;
    if (!identity || !dates) { reject('자동 반영에 필요한 문자 인식 근거가 없습니다.'); return output; }
    if (result.identityRecoveryIncomplete) reject('이름·시리얼 추가 인식을 끝내지 못했습니다. 다른 툴과 혼동할 수 있어 자동 반영을 보류합니다.');
    const serialMatches = tools.filter((tool) => (identity.serialMatches || []).some((match) => match.toolId === tool.id && normalizeSerial(match.serial) === normalizeSerial(tool.serial)));
    const nameMatches = tools.filter((tool) => (identity.nameMatches || []).some((match) => match.toolId === tool.id && normalizeName(match.name) === normalizeName(tool.name)));
    const strongNameMatches = nameMatches.filter((tool) => (identity.nameMatches || []).some((match) => match.toolId === tool.id && high(match.confidence)));
    // Unlabelled Korean words on a calibration sticker can be company names
    // or headings. They must not become contradictory owner names merely
    // because they have the same length as a person's name.
    const conflictingNames = Array.isArray(identity.observedNameEvidence)
      ? identity.observedNameEvidence.filter((proof) => ['workbook', 'labelled'].includes(proof.kind) && high(proof.confidence)).map((proof) => proof.name)
      : identity.observedNames || [];
    const conflictingSerials = Array.isArray(identity.observedSerialEvidence)
      ? identity.observedSerialEvidence.filter((proof) => high(proof.confidence)).map((proof) => proof.serial)
      : identity.observedSerials || [];
    const clearDatePair = !dates.ambiguous && ['labelled', 'repeated-dates'].includes(dates.method) && high(dates.confidence)
      && isoDate(dates.last) && isoDate(dates.due) && dates.due > dates.last;
    const observedSerialEvidence = Array.isArray(identity.observedSerialEvidence) ? identity.observedSerialEvidence : [];
    const usableSource = (value) => {
      const source = typeof value === 'string' ? value.trim() : '';
      return source && !/^(?:text|unknown|null|undefined)$/i.test(source)
        && !/^(?:date|label)[-_]/i.test(source) && !/(?:^|[-_])incomplete(?:$|[-_])/i.test(source) ? source : '';
    };
    function proofSources(proof) {
      const listed = new Set((Array.isArray(proof.passSources) ? proof.passSources : []).map(usableSource).filter(Boolean));
      return (Array.isArray(proof.sourceReadings) ? proof.sourceReadings : []).filter((reading) => {
        const source = usableSource(reading && reading.source);
        return source && listed.has(source) && high(reading.confidence);
      }).map((reading) => usableSource(reading.source));
    }
    function actualSources(value) {
      const sources = new Set();
      for (const proof of observedSerialEvidence) {
        if (proof.learned || !high(proof.confidence) || !['standalone', 'labelled', 'workbook'].includes(proof.kind)
          || normalizeSerial(proof.serial) !== value) continue;
        // The aggregate maximum and passSources cannot prove that each pass
        // read the fragment confidently. Older evidence therefore stays pending.
        for (const source of proofSources(proof)) sources.add(source);
      }
      return Array.from(sources);
    }
    const exactOwner = (target) => (identity.nameMatches || []).some((proof) => !proof.learned && proof.toolId === target.id
      && normalizeName(proof.name) === normalizeName(target.name) && high(proof.confidence));
    function exactSerialSources(target) {
      const exact = (identity.serialMatches || []).some((proof) => !proof.learned && proof.toolId === target.id
        && normalizeSerial(proof.serial) === normalizeSerial(target.serial) && high(proof.confidence));
      return exact ? actualSources(normalizeSerial(target.serial)) : [];
    }
    // Both selection and conflict handling use this one global candidate rule.
    // No characters are changed: a fragment must be a literal contiguous part
    // of precisely one serial in the entire workbook, including disposed tools.
    function resolvePartialSerial(value, options) {
      const fragment = normalizeSerial(value), settings = options || {};
      if (!fragment || tools.some((candidate) => normalizeSerial(candidate.serial) === fragment)) return null;
      const candidates = tools.filter((candidate) => {
        const canonical = normalizeSerial(candidate.serial);
        return canonical.length > fragment.length && canonical.includes(fragment);
      });
      if (candidates.length !== 1) return null;
      const target = candidates[0], canonical = normalizeSerial(target.serial);
      const withOwner = !settings.exactConfirmation && exactOwner(target);
      if (fragment.length < (withOwner ? 5 : 6) || fragment.length / canonical.length < (withOwner ? 0.5 : 0.6)) return null;
      const sources = actualSources(fragment);
      if (settings.exactConfirmation) {
        const exactSources = exactSerialSources(target);
        if (!sources.some((source) => exactSources.some((other) => source !== other))) return null;
      } else {
        if (sources.length < 2 || !clearDatePair || addMonths(dates.last, monthInterval(target)) !== dates.due) return null;
      }
      return { tool: target, fragment, sources };
    }
    const partialMatches = Array.from(new Set(observedSerialEvidence.map((proof) => normalizeSerial(proof.serial))))
      .map((fragment) => resolvePartialSerial(fragment)).filter(Boolean);
    const partialTools = tools.filter((candidate) => partialMatches.some((match) => match.tool.id === candidate.id));
    function ownerFragmentState() {
      const names = Array.from(new Set(tools.filter(exactOwner).map((candidate) => normalizeName(candidate.name))));
      const empty = { matches: [], target: null, blocked: false };
      if (names.length !== 1) return empty;
      const evidence = Array.isArray(identity.observedSerialFragments) ? identity.observedSerialFragments : [];
      const observed = new Map();
      for (const proof of evidence) {
        const fragment = normalizeSerial(proof.fragment);
        // This separate channel contains actual four-digit tokens, never a
        // guessed edit, a substring extracted from a longer reading, or a year.
        if (proof.learned || !high(proof.confidence) || !/^\d{4}$/.test(fragment) || /^20\d{2}$/.test(fragment)) continue;
        if (!observed.has(fragment)) observed.set(fragment, new Set());
        for (const source of proofSources(proof)) observed.get(fragment).add(source);
      }
      const fragments = Array.from(observed).filter(([, sources]) => sources.size >= 2).map(([fragment]) => fragment);
      if (!fragments.length) return empty;
      const ownerTools = tools.filter((candidate) => normalizeName(candidate.name) === names[0]);
      const readings = fragments.map((fragment) => {
        // Keep disposed tools in this comparison. A calibration interval must
        // never break a tie between two serials containing the same fragment.
        const known = tools.some((candidate) => normalizeSerial(candidate.serial) === fragment);
        const candidates = ownerTools.filter((candidate) => {
          const canonical = normalizeSerial(candidate.serial);
          return canonical.length > fragment.length && canonical.includes(fragment);
        });
        return { fragment, candidates, known };
      });
      const present = readings.filter((reading) => reading.candidates.length);
      const ambiguous = readings.some((reading) => reading.known || reading.candidates.length > 1);
      const targets = Array.from(new Set(present.map((reading) => reading.candidates[0].id)));
      const blocked = ambiguous || targets.length > 1 || (present.length > 0 && present.length !== readings.length);
      if (blocked || !present.length) return { matches: [], target: null, blocked };
      const target = present[0].candidates[0];
      if (present.some((reading) => reading.fragment.length / normalizeSerial(target.serial).length < 0.4)) return empty;
      return { matches: present.map((reading) => ({ tool: target, fragment: reading.fragment })), target, blocked: false };
    }
    const ownerFragments = ownerFragmentState();
    const ownerFragmentMatch = ownerFragments.target && clearDatePair
      && addMonths(dates.last, monthInterval(ownerFragments.target)) === dates.due ? ownerFragments : null;
    const resolveSerial = (serial, target) => {
      if ((identity.resolvedSerials || []).some((proof) => proof.toolId === target.id && high(proof.confidence)
        && normalizeSerial(proof.alias) === normalizeSerial(serial) && normalizeSerial(proof.serial) === normalizeSerial(target.serial))) return true;
      // Only this exact, separately verified four-digit observation may be
      // reconciled with an explicit S/N reading. Other full serial conflicts stay.
      if (ownerFragmentMatch && ownerFragmentMatch.target.id === target.id
        && ownerFragmentMatch.matches.some((match) => match.fragment === normalizeSerial(serial))) return true;
      const partial = resolvePartialSerial(serial) || resolvePartialSerial(serial, { exactConfirmation: true });
      return !!partial && partial.tool.id === target.id;
    };
    const resolveName = (name, target) => (identity.resolvedNames || []).some((proof) => proof.toolId === target.id && high(proof.confidence)
      && normalizeName(proof.alias) === normalizeName(name) && normalizeName(proof.name) === normalizeName(target.name));
    let tool = null;
    if (serialMatches.length === 1) {
      tool = serialMatches[0];
      const proof = identity.serialMatches.find((match) => match.toolId === tool.id);
      output.matchMethod = proof?.learned ? 'remembered-serial' : 'serial';
      if (!high(proof && proof.confidence)) reject('시리얼 번호의 인식 신뢰도가 낮아 확인이 필요합니다.');
    } else if (serialMatches.length > 1) reject('인식한 시리얼 번호가 여러 툴과 일치합니다.');
    else if (partialTools.length === 1) { tool = partialTools[0]; output.matchMethod = 'partial-serial'; }
    else if (partialTools.length > 1) reject('부분 시리얼 번호가 서로 다른 툴을 가리킵니다. 자동 반영을 보류합니다.');
    else if (ownerFragmentMatch) { tool = ownerFragmentMatch.target; output.matchMethod = 'owner-partial-serial'; }
    else if (nameMatches.length === 1 && tools.filter((other) => normalizeName(other.name) === normalizeName(nameMatches[0].name)).length === 1) {
      tool = nameMatches[0];
      const proof = identity.nameMatches.find((match) => match.toolId === tool.id);
      output.matchMethod = proof?.learned ? 'remembered-name' : 'name';
      if (!high(proof && proof.confidence)) reject('이름의 인식 신뢰도가 낮아 확인이 필요합니다.');
    } else if (nameMatches.length) {
      // The workbook is the reference for every interval, including tools
      // calibrated less often than a torque wrench or a multimeter.
      const owners = new Set(strongNameMatches.map((candidate) => normalizeName(candidate.name)));
      const cycleMatches = clearDatePair && owners.size === 1 ? tools.filter((candidate) => owners.has(normalizeName(candidate.name))
        && addMonths(dates.last, monthInterval(candidate)) === dates.due) : [];
      if (cycleMatches.length === 1) { tool = cycleMatches[0]; output.matchMethod = 'name-cycle'; }
      else reject('이 엔지니어에게 툴이 여러 개 있어 시리얼 번호 또는 구분되는 교정 주기가 필요합니다.');
    }
    else reject('사진의 이름 또는 시리얼 번호와 정확히 일치하는 툴이 없습니다.');
    if (tool) {
      output.toolId = tool.id; output.name = tool.name; output.serial = tool.serial;
      if (['name', 'remembered-name', 'name-cycle', 'owner-partial-serial'].includes(output.matchMethod)) {
        if (ownerFragments.blocked) reject('같은 담당자의 부분 번호가 여러 툴을 가리키거나 서로 다릅니다. 주기만으로 선택하지 않습니다.');
        if (ownerFragments.target && ownerFragments.target.id !== tool.id) reject('사진의 부분 번호와 검교정 주기가 서로 다른 툴을 가리킵니다.');
      }
      if (conflictingNames.some((name) => normalizeName(name) !== normalizeName(tool.name) && !resolveName(name, tool))) reject('사진의 이름이 시리얼 번호의 담당자와 다르게 읽혔습니다.');
      if (conflictingSerials.some((serial) => normalizeSerial(serial) !== normalizeSerial(tool.serial) && !resolveSerial(serial, tool))) reject('사진에서 서로 다른 시리얼 번호가 읽혔습니다.');
      if (tools.filter((other) => normalizeSerial(other.serial) === normalizeSerial(tool.serial)).length !== 1) reject('엑셀에 동일한 시리얼 번호가 중복되어 있습니다.');
      if (strongNameMatches.length && !strongNameMatches.some((other) => other.id === tool.id)) reject('사진의 이름과 시리얼 번호가 서로 다른 툴을 가리킵니다.');
      if (!tool.name || !normalizeSerial(tool.serial)) reject('엑셀 항목에 담당자 또는 시리얼 번호가 없습니다.');
      if (/폐기|disposed|retired/i.test(String(tool.due || ''))) reject('폐기 처리된 툴은 자동 변경하지 않습니다.');
      if (Number(tool.photoCount) > 1) reject('이 툴의 사진 위치에 여러 사진이 있어 자동 교체할 수 없습니다.');
    }
    if (dates.ambiguous) reject('교정일 또는 다음 교정일이 서로 다르게 읽혔습니다.');
    if (['labelled', 'repeated-dates'].includes(dates.method) && high(dates.confidence)) {
      output.last = dates.last; output.due = dates.due; output.dateMethod = dates.method;
    } else if (tool && !dates.ambiguous && dates.explicitLast && high(dates.lastConfidence) && !dates.explicitDue && !dates.dueLabelSeen && (dates.dates || []).length === 1) {
      const months = monthInterval(tool), due = months ? addMonths(dates.explicitLast, months) : '';
      if (due) { output.last = dates.explicitLast; output.due = due; output.dateMethod = 'calculated-cycle'; }
      else reject('차기 교정일을 읽지 못했고 엑셀의 검교정 주기도 없습니다.');
    } else reject('교정일·다음 교정일의 라벨과 인식 신뢰도를 확인해야 합니다.');
    if (output.last && output.due) {
      if (!isoDate(output.last) || !isoDate(output.due)) reject('유효한 교정 날짜를 읽지 못했습니다.');
      else if (output.due <= output.last) reject('차기 교정일이 교정일보다 늦어야 합니다.');
      if (tool && isoDate(tool.last) && output.last < tool.last) reject('사진의 교정일이 기존 기록보다 오래되어 자동 반영하지 않습니다.');
    }
    output.eligible = Boolean(tool && output.last && output.due && !output.reasons.length);
    return output;
  }
  global.ToolAutomation = Object.freeze({ assess });
})(typeof window !== 'undefined' ? window : globalThis);
