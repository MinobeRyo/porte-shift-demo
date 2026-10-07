import { daysInMonth, normalizeMonthlyProposal, applyMonthlyEntries } from './monthly-core.js';

const labels = {
  experience: ['未設定', '新人', 'ベテラン'],
  category: ['未設定', '学生', '社会人'],
  preference: ['未設定', '少なめ', '通常', '多め']
};
const own = (value, key) => Object.prototype.hasOwnProperty.call(value, key);
function object(value, label) {
  if (!value || typeof value !== 'object' || Array.isArray(value) || ![Object.prototype, null].includes(Object.getPrototypeOf(value))) throw new Error(`${label}の形式が正しくありません。`);
}
function name(value) {
  if (typeof value !== 'string' || !value.trim() || value.trim().length > 40) throw new Error('スタッフ名は1〜40文字で指定してください。');
  return value.trim();
}
function condition(value, month) {
  object(value, '基本条件');
  const result = { name: name(value.name) };
  for (const [key, allowed] of Object.entries(labels)) {
    if (!allowed.includes(value[key])) throw new Error(`${result.name}の基本条件（${key}）が正しくありません。`);
    result[key] = value[key];
  }
  if (!Array.isArray(value.weekdays) || value.weekdays.length > 7 || value.weekdays.some(day => !Number.isInteger(day) || day < 0 || day > 6) || new Set(value.weekdays).size !== value.weekdays.length) throw new Error(`${result.name}の基本曜日が正しくありません。`);
  result.weekdays = [...value.weekdays];
  // Explicitly reject missing times instead of silently applying the core defaults.
  if (typeof value.start !== 'string' || !value.start || typeof value.end !== 'string' || !value.end) throw new Error(`${result.name}の基本開始・終了時刻を指定してください。`);
  const normalized = normalizeMonthlyProposal({ needsClarification: false, explanation: '', entries: [{ day: 1, status: 'available', start: value.start, end: value.end }] }, { month });
  result.start = normalized.entries[0].start;
  result.end = normalized.entries[0].end;
  return result;
}
function createStaff(staffName) {
  return { id: crypto.randomUUID(), name: staffName, experience: '未設定', category: '未設定', preference: '未設定', weekdays: [], start: '09:00', end: '18:00' };
}

// Produce one reviewable, atomic change. The caller saves nextState only after confirmation.
export function buildExcelImportPlan(state, parsed, { mapping, conditions = [], importConditions = false, filename = '' } = {}) {
  object(state, '保存済みデータ');
  object(parsed, 'Excelの読み取り結果');
  daysInMonth(parsed.month);
  if (!Array.isArray(state.staff) || state.staff.length > 30 || !Array.isArray(state.history)) throw new Error('保存済みスタッフ・履歴の形式が正しくありません。');
  object(state.months, '保存済みの希望表');
  if (!Array.isArray(parsed.staff) || !parsed.staff.length) throw new Error('取り込むスタッフがいません。');
  if (!Array.isArray(mapping) || !Array.isArray(conditions)) throw new Error('スタッフの対応表・基本条件の形式が正しくありません。');
  if (typeof importConditions !== 'boolean' || typeof filename !== 'string') throw new Error('取り込み設定の形式が正しくありません。');

  const existingById = new Map(), usedNames = new Set();
  for (const staff of state.staff) {
    object(staff, '保存済みスタッフ');
    const staffName = name(staff.name);
    if (typeof staff.id !== 'string' || !staff.id || ['new', 'skip', '__proto__', 'constructor', 'prototype'].includes(staff.id) || existingById.has(staff.id) || usedNames.has(staffName)) throw new Error('保存済みスタッフのIDまたは名前が重複・不正です。');
    existingById.set(staff.id, staff);
    usedNames.add(staffName);
  }
  const sources = new Map();
  for (const source of parsed.staff) {
    object(source, '取り込み元スタッフ');
    const sourceName = name(source.name);
    if (sources.has(sourceName)) throw new Error(`取り込み元に同じ名前（${sourceName}）があります。`);
    if (!Array.isArray(source.entries)) throw new Error(`${sourceName}の勤務希望の形式が正しくありません。`);
    sources.set(sourceName, source);
  }
  const conditionByName = new Map();
  for (const raw of conditions) {
    const normalized = condition(raw, parsed.month);
    if (conditionByName.has(normalized.name)) throw new Error(`基本条件に同じ名前（${normalized.name}）があります。`);
    conditionByName.set(normalized.name, normalized);
  }
  const targets = new Map(), usedTargets = new Set();
  let peopleCount = 0, newStaffCount = 0;
  for (const item of mapping) {
    object(item, 'スタッフの対応');
    const sourceName = name(item.sourceName), targetId = item.targetId;
    if (!sources.has(sourceName) || targets.has(sourceName)) throw new Error('対応表に存在しないスタッフ、または重複した指定があります。');
    if (typeof targetId !== 'string' || !['new', 'skip'].includes(targetId) && !existingById.has(targetId)) throw new Error(`${sourceName}の対応先を選び直してください。`);
    targets.set(sourceName, targetId);
    if (targetId === 'skip') continue;
    peopleCount++;
    if (targetId === 'new') {
      if (usedNames.has(sourceName)) throw new Error(`${sourceName}は登録済みです。既存スタッフを対応先に選んでください。`);
      usedNames.add(sourceName); newStaffCount++;
    } else {
      if (usedTargets.has(targetId)) throw new Error('複数の取り込み元を同じスタッフに対応させることはできません。');
      usedTargets.add(targetId);
    }
  }
  if (targets.size !== sources.size) throw new Error('すべての取り込み元スタッフについて対応先を選んでください。');
  if (!peopleCount) throw new Error('取り込むスタッフを1人以上選んでください。');
  if (state.staff.length + newStaffCount > 30) throw new Error('スタッフは新規登録を含めて30名までです。');

  const nextState = structuredClone(state), changes = [], records = [];
  nextState.month = parsed.month;
  if (!own(nextState.months, parsed.month)) nextState.months[parsed.month] = {};
  object(nextState.months[parsed.month], '対象月の希望表');
  let overwrites = 0, entryCount = 0;
  const text = filename.slice(0, 1500), at = new Date().toISOString();
  for (const [sourceName, source] of sources) {
    const targetId = targets.get(sourceName);
    if (targetId === 'skip') continue;
    const staff = targetId === 'new' ? createStaff(sourceName) : nextState.staff.find(person => person.id === targetId);
    if (targetId === 'new') nextState.staff.push(staff);
    if (importConditions && conditionByName.has(sourceName)) {
      const { name: ignoredName, ...values } = conditionByName.get(sourceName);
      Object.assign(staff, structuredClone(values));
    }
    const options = { month: parsed.month, defaultStart: staff.start, defaultEnd: staff.end };
    const proposal = normalizeMonthlyProposal({ needsClarification: source.entries.length === 0, explanation: '', entries: source.entries }, options);
    const entries = proposal.entries;
    const existing = own(nextState.months[parsed.month], staff.id) ? nextState.months[parsed.month][staff.id] : {};
    const updated = applyMonthlyEntries(existing, entries, { month: parsed.month, source: 'excel', text });
    const previous = {};
    for (const entry of entries) {
      const before = own(existing, entry.day) ? structuredClone(existing[entry.day]) : null;
      previous[entry.day] = before;
      changes.push({ staffName: staff.name, day: entry.day, previous: before, next: structuredClone(entry) });
      if (before !== null) overwrites++;
    }
    nextState.months[parsed.month][staff.id] = updated;
    entryCount += entries.length;
    records.push({ at, month: parsed.month, staffId: staff.id, staffName: staff.name, source: 'excel', text, entries: structuredClone(entries), previous });
  }
  nextState.history = [...records, ...nextState.history].slice(0, 100);
  return { changes, overwrites, newStaffCount, entryCount, peopleCount, nextState };
}
