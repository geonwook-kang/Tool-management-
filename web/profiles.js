/* Only reviewed tool identities and narrow OCR variants are kept in this browser. */
(function (global) {
  'use strict';
  const KEY = 'tool-management-profiles-v1', MAX_PROFILES = 500, MAX_ALIASES = 10, MAX_CONFLICTS = 5000;
  let profiles = [], blockedNames = [], blockedSerials = [], suspended = false;
  let persistent = true, storageError = '';
  const serial = (value) => String(value == null ? '' : value).normalize('NFKC').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const owner = (value) => String(value || '').normalize('NFKC').replace(/\s/g, '').toUpperCase();
  const asset = (value) => String(value || '').normalize('NFKC').trim().replace(/\s+/g, ' ');
  const unique = (values) => Array.from(new Set(values));
  const high = (value) => value != null && Number.isFinite(Number(value)) && Number(value) >= 70 && Number(value) <= 100;
  function months(tool) {
    if (Number.isInteger(tool.months) && tool.months > 0 && tool.months <= 1200) return tool.months;
    const n = Number(tool.interval), cycle = String(tool.cycle || '').trim().toUpperCase();
    if (!Number.isInteger(n) || n <= 0) return null;
    const count = ['YEAR', 'YEARS', '년'].includes(cycle) ? n * 12 : ['MONTH', 'MONTHS', '월', '개월'].includes(cycle) ? n : null;
    return count && count <= 1200 ? count : null;
  }
  function date(value) {
    const match = /^(20\d{2})-(\d{2})-(\d{2})$/.exec(String(value || ''));
    if (!match) return null;
    const parts = match.slice(1).map(Number), result = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    return result.getUTCFullYear() === parts[0] && result.getUTCMonth() + 1 === parts[1] && result.getUTCDate() === parts[2] ? result : null;
  }
  function addMonths(value, count) {
    const start = date(value);
    if (!start || !Number.isInteger(count) || count < 1 || count > 1200) return '';
    const result = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth() + count + 1, 0));
    result.setUTCDate(Math.min(start.getUTCDate(), result.getUTCDate()));
    return result.toISOString().slice(0, 10);
  }
  function oneEdit(a, b) {
    const left = Array.from(a), right = Array.from(b);
    if (Math.abs(left.length - right.length) > 1 || a === b) return false;
    let i = 0, j = 0, edits = 0;
    while (i < left.length && j < right.length) {
      if (left[i] === right[j]) { i++; j++; continue; }
      if (++edits > 1) return false;
      if (left.length >= right.length) i++;
      if (right.length >= left.length) j++;
    }
    return edits + Number(i < left.length || j < right.length) === 1;
  }
  function profileKey(profile) { return JSON.stringify([profile.serial, profile.owner, profile.asset, profile.months]); }
  function canonical(tool) { return { serial: serial(tool.serial), owner: owner(tool.name), asset: asset(tool.asset), months: months(tool) }; }
  function canonicalValid(value) {
    return value.serial.length >= 3 && value.serial.length <= 64 && value.owner.length >= 2 && value.owner.length <= 80 && value.asset.length <= 512 && Number.isInteger(value.months) && value.months > 0 && value.months <= 1200;
  }
  function available(tool) { return !/폐기|disposed|retired/i.test(String(tool.due || '')) && !(Number(tool.photoCount) > 1); }
  function activeProfile(profile, tools) {
    if (!canonicalValid(profile) || tools.filter((tool) => serial(tool.serial) === profile.serial).length !== 1) return null;
    const matches = tools.filter((tool) => profileKey(canonical(tool)) === profileKey(profile) && available(tool));
    return matches.length === 1 ? matches[0] : null;
  }
  function sanitizeProfile(value) {
    if (!value || typeof value !== 'object') return null;
    const result = { serial: serial(value.serial), owner: owner(value.owner), asset: asset(value.asset), months: Number(value.months), nameAliases: [], serialAliases: [] };
    if (!canonicalValid(result)) return null;
    result.nameAliases = unique((Array.isArray(value.nameAliases) ? value.nameAliases : []).map(owner)).filter((v) => v.length >= 2 && v.length <= 80 && oneEdit(v, result.owner)).slice(0, MAX_ALIASES);
    result.serialAliases = unique((Array.isArray(value.serialAliases) ? value.serialAliases : []).map(serial)).filter((v) => v.length === result.serial.length && oneEdit(v, result.serial)).slice(0, MAX_ALIASES - result.nameAliases.length);
    return result;
  }
  function load() {
    try {
      const raw = global.localStorage.getItem(KEY);
      if (!raw) return;
      if (raw.length > 1500000) throw new Error('size');
      const saved = JSON.parse(raw);
      if (!saved || saved.version !== 1 || !Array.isArray(saved.profiles)) throw new Error('format');
      const found = new Set();
      profiles = saved.profiles.slice(-MAX_PROFILES).map(sanitizeProfile).filter((p) => p && !found.has(profileKey(p)) && found.add(profileKey(p)));
      blockedNames = unique((Array.isArray(saved.blockedNames) ? saved.blockedNames : []).slice(0, MAX_CONFLICTS).map(owner)).filter((v) => v.length >= 2 && v.length <= 80);
      blockedSerials = unique((Array.isArray(saved.blockedSerials) ? saved.blockedSerials : []).slice(0, MAX_CONFLICTS).map(serial)).filter((v) => v.length >= 3 && v.length <= 64);
      suspended = saved.suspended === true;
    } catch (_) {
      profiles = []; blockedNames = []; blockedSerials = []; suspended = false;
      persistent = false; storageError = '이 브라우저의 툴 기억 정보를 읽지 못했습니다. 사진 인식은 계속할 수 있습니다.';
    }
  }
  function persist() {
    try {
      global.localStorage.setItem(KEY, JSON.stringify({ version: 1, profiles, blockedNames, blockedSerials, suspended }));
      persistent = true; storageError = '';
    } catch (_) {
      persistent = false; storageError = '툴 기억 정보를 저장하지 못했습니다. 이번 창에서만 기억하며 사진 인식은 계속할 수 있습니다.';
    }
  }
  function inspect() {
    return {
      count: profiles.length, aliasCount: profiles.reduce((n, p) => n + p.nameAliases.length + p.serialAliases.length, 0),
      conflictCount: blockedNames.length + blockedSerials.length, persistent, error: storageError, suspended
    };
  }
  function clear() {
    profiles = []; blockedNames = []; blockedSerials = []; suspended = false;
    try { global.localStorage.removeItem(KEY); persistent = true; storageError = ''; }
    catch (_) { persistent = false; storageError = '이번 창의 기억은 삭제했습니다. 브라우저 저장소의 이전 기억은 삭제하지 못했습니다.'; }
    return inspect();
  }
  function exactSerial(tool, identity) {
    return (identity.serialMatches || []).some((proof) => !proof.learned && proof.toolId === tool.id && serial(proof.serial) === serial(tool.serial) && high(proof.confidence));
  }
  function exactName(tool, identity) {
    return (identity.nameMatches || []).some((proof) => !proof.learned && proof.toolId === tool.id && owner(proof.name) === owner(tool.name) && high(proof.confidence));
  }
  function observedExactSerial(tool, identity) {
    return (identity.observedSerialEvidence || []).some((proof) => !proof.learned && high(proof.confidence) && serial(proof.serial) === serial(tool.serial));
  }
  function observedExactName(tool, identity) {
    return (identity.observedNameEvidence || []).some((proof) => !proof.learned && high(proof.confidence) && owner(proof.name) === owner(tool.name));
  }
  function dateProof(result, tool, review) {
    const evidence = result && result.dateEvidence;
    if (!evidence || evidence.ambiguous) return false;
    if (['labelled', 'repeated-dates'].includes(evidence.method) && high(evidence.confidence) && evidence.last === review.last && evidence.due === review.due) return true;
    return !!(evidence.explicitLast === review.last && high(evidence.lastConfidence) && !evidence.explicitDue && !evidence.dueLabelSeen && (evidence.dates || []).length === 1 && addMonths(review.last, months(tool)) === review.due);
  }
  function independentNameCycle(tool, tools, result, review) {
    const evidence = result.dateEvidence || {}, identity = result.identityEvidence || {};
    if (!exactName(tool, identity) || !['labelled', 'repeated-dates'].includes(evidence.method) || !high(evidence.confidence) || evidence.ambiguous || evidence.last !== review.last || evidence.due !== review.due) return false;
    const matches = tools.filter((candidate) => owner(candidate.name) === owner(tool.name) && addMonths(review.last, months(candidate)) === review.due && available(candidate));
    return matches.length === 1 && matches[0].id === tool.id;
  }
  function certificateVariant(result, variant, proof) {
    if (/cert|certificate|성적서/i.test(String(proof.kind || ''))) return true;
    const texts = [result.text || ''].concat((result.passes || []).map((pass) => pass.text || ''));
    const label = /성적\s*서|성적\s*번호|certificate|cert\.?\s*(?:no\.?|number)/i;
    return texts.some((text) => {
      const lines = String(text).split(/\r?\n/);
      return lines.some((line, index) => serial(line).includes(variant) && (label.test(line) || (index > 0 && label.test(lines[index - 1]) && serial(line) === variant)));
    });
  }
  function dateVariant(result, value) {
    if (/^20\d{6}$/.test(value) && date(value.slice(0, 4) + '-' + value.slice(4, 6) + '-' + value.slice(6, 8))) return true;
    const evidence = result.dateEvidence || {};
    return (evidence.dates || []).concat([evidence.last, evidence.due, evidence.explicitLast, evidence.explicitDue]).filter(Boolean).some((v) => serial(v) === value);
  }
  function block(kind, variant) {
    const list = kind === 'nameAliases' ? blockedNames : blockedSerials;
    if (!list.includes(variant)) {
      if (list.length >= MAX_CONFLICTS) suspended = true;
      else list.push(variant);
    }
    for (const profile of profiles) profile[kind] = profile[kind].filter((v) => v !== variant);
  }
  function addAlias(profile, kind, variant, tools) {
    if (suspended || (kind === 'nameAliases' ? blockedNames : blockedSerials).includes(variant)) return;
    const other = profiles.filter((p) => p !== profile && p[kind].includes(variant));
    if (other.length) {
      // One owner may have several tools. A name variant belongs to that owner.
      if (kind === 'nameAliases' && other.every((p) => p.owner === profile.owner)) {
        if (other.some((p) => activeProfile(p, tools))) return;
        // A reviewed change of asset or interval invalidates the old binding.
        for (const previous of other) previous[kind] = previous[kind].filter((v) => v !== variant);
      } else { block(kind, variant); return; }
    }
    if (!profile[kind].includes(variant) && profile.nameAliases.length + profile.serialAliases.length < MAX_ALIASES) profile[kind].push(variant);
  }
  function remember(model, photo) {
    try {
      const tools = model && Array.isArray(model.tools) ? model.tools : [], review = photo && photo.review;
      if (!review || !['automatic', 'manual'].includes(photo.approval)) return inspect();
      const candidates = tools.filter((tool) => tool.id === review.toolId), tool = candidates.length === 1 ? candidates[0] : null;
      if (!tool || !available(tool) || serial(review.serial) !== serial(tool.serial) || owner(review.name) !== owner(tool.name) || !date(review.last) || !date(review.due) || review.due <= review.last || (date(tool.last) && review.last < tool.last)) return inspect();
      const base = canonical(tool);
      if (!canonicalValid(base) || tools.filter((candidate) => serial(candidate.serial) === base.serial).length !== 1) return inspect();
      // A manually reviewed identity is usable after an engine failure, but
      // partial OCR output from that failure must never teach a variant.
      const result = photo.result && !photo.result.failed ? photo.result : {}, identity = result.identityEvidence || {};
      const exact = exactSerial(tool, identity), nameCycle = independentNameCycle(tool, tools, result, review);
      if (photo.approval === 'automatic' && (!dateProof(result, tool, review) || !(exact || nameCycle))) return inspect();
      if (photo.approval === 'automatic') {
        const otherSerial = tools.some((candidate) => candidate.id !== tool.id && (exactSerial(candidate, identity) || observedExactSerial(candidate, identity)));
        const otherOwner = tools.some((candidate) => owner(candidate.name) !== base.owner && (exactName(candidate, identity) || observedExactName(candidate, identity)));
        if (otherSerial || otherOwner) return inspect();
      }
      let profile = profiles.find((p) => profileKey(p) === profileKey(base));
      if (!profile) { profile = Object.assign(base, { nameAliases: [], serialAliases: [] }); profiles.push(profile); }
      if (exact) for (const proof of identity.observedNameEvidence || []) {
        if (proof.learned || !['standalone', 'labelled'].includes(proof.kind)) continue;
        const variant = owner(proof.name);
        if (variant.length < 2 || variant.length > 80 || !oneEdit(variant, base.owner) || tools.some((candidate) => owner(candidate.name) === variant && owner(candidate.name) !== base.owner)) continue;
        addAlias(profile, 'nameAliases', variant, tools);
      }
      if (photo.approval === 'manual' || nameCycle) for (const proof of identity.observedSerialEvidence || []) {
        if (proof.learned || !['standalone', 'labelled'].includes(proof.kind)) continue;
        const variant = serial(proof.serial);
        if (variant.length !== base.serial.length || !oneEdit(variant, base.serial) || tools.some((candidate) => serial(candidate.serial) === variant) || dateVariant(result, variant) || certificateVariant(result, variant, proof)) continue;
        addAlias(profile, 'serialAliases', variant, tools);
      }
      if (profiles.length > MAX_PROFILES) profiles = profiles.slice(-MAX_PROFILES);
      persist(); return inspect();
    } catch (_) { storageError = '이번 사진의 기억 정보를 만들지 못했습니다. 사진 인식과 엑셀 반영은 계속할 수 있습니다.'; return inspect(); }
  }
  function enhance(model, result) {
    if (!result || result.failed || !result.identityEvidence) return result;
    try {
      const tools = model && Array.isArray(model.tools) ? model.tools : [], original = result.identityEvidence;
      const identity = Object.assign({}, original, {
        nameMatches: (original.nameMatches || []).filter((p) => !p.learned).map((p) => Object.assign({}, p)),
        serialMatches: (original.serialMatches || []).filter((p) => !p.learned).map((p) => Object.assign({}, p)),
        resolvedNames: (original.resolvedNames || []).filter((p) => !p.learned).map((p) => Object.assign({}, p)),
        resolvedSerials: (original.resolvedSerials || []).filter((p) => !p.learned).map((p) => Object.assign({}, p))
      });
      const baseIds = result.profileBaseCandidateIds || result.candidateIds || [];
      const output = Object.assign({}, result, { identityEvidence: identity, profileBaseCandidateIds: baseIds.slice(), candidateIds: baseIds.filter((id) => tools.some((tool) => tool.id === id)) });
      if (suspended) return output;
      const active = profiles.map((profile) => ({ profile, tool: activeProfile(profile, tools) })).filter((entry) => entry.tool);
      const exactTools = tools.filter((tool) => exactSerial(tool, identity) || observedExactSerial(tool, identity));
      const exactOwners = unique(tools.filter((tool) => exactName(tool, identity) || observedExactName(tool, identity)).map((tool) => owner(tool.name)));
      for (const proof of original.observedSerialEvidence || []) {
        if (proof.learned || !high(proof.confidence) || !['standalone', 'labelled'].includes(proof.kind)) continue;
        const variant = serial(proof.serial);
        if (blockedSerials.includes(variant) || tools.some((tool) => serial(tool.serial) === variant) || dateVariant(result, variant) || certificateVariant(result, variant, proof)) continue;
        const matches = active.filter((entry) => entry.profile.serialAliases.includes(variant));
        if (matches.length !== 1) continue;
        const entry = matches[0], tool = entry.tool;
        if (exactTools.some((other) => other.id !== tool.id) || exactOwners.some((name) => name !== owner(tool.name))) continue;
        if (!identity.serialMatches.some((p) => p.toolId === tool.id)) identity.serialMatches.push({ toolId: tool.id, serial: tool.serial, confidence: Number(proof.confidence), learned: true, alias: variant });
        if (!identity.resolvedSerials.some((p) => p.alias === variant && p.toolId === tool.id)) identity.resolvedSerials.push({ toolId: tool.id, serial: tool.serial, confidence: Number(proof.confidence), learned: true, alias: variant });
        output.candidateIds.push(tool.id);
      }
      for (const proof of original.observedNameEvidence || []) {
        if (proof.learned || !high(proof.confidence) || !['standalone', 'labelled'].includes(proof.kind)) continue;
        const variant = owner(proof.name);
        if (blockedNames.includes(variant) || tools.some((tool) => owner(tool.name) === variant)) continue;
        const matches = active.filter((entry) => entry.profile.nameAliases.includes(variant));
        if (matches.length !== 1) continue;
        const canonicalOwner = matches[0].profile.owner;
        if (exactTools.some((tool) => owner(tool.name) !== canonicalOwner) || exactOwners.some((name) => name !== canonicalOwner)) continue;
        for (const tool of tools.filter((candidate) => owner(candidate.name) === canonicalOwner && available(candidate))) {
          if (!identity.nameMatches.some((p) => p.toolId === tool.id)) identity.nameMatches.push({ toolId: tool.id, name: tool.name, confidence: Number(proof.confidence), learned: true, alias: variant });
          if (!identity.resolvedNames.some((p) => p.alias === variant && p.toolId === tool.id)) identity.resolvedNames.push({ toolId: tool.id, name: tool.name, confidence: Number(proof.confidence), learned: true, alias: variant });
          output.candidateIds.push(tool.id);
        }
      }
      output.candidateIds = unique(output.candidateIds);
      return output;
    } catch (_) { return result; }
  }
  load();
  global.ToolProfiles = Object.freeze({ remember, enhance, inspect, clear });
})(typeof window !== 'undefined' ? window : globalThis);
