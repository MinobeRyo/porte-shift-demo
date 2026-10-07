const ENTRY_KEYS = ['day', 'status', 'start', 'end'];
const FLAGS = ['startDefaulted', 'endDefaulted'];
const own = (value, key) => Object.prototype.hasOwnProperty.call(value, key);
function checkObject(value, required, optional = [], label = 'データ') {
  const prototype = value !== null && typeof value === 'object' ? Object.getPrototypeOf(value) : undefined;
  if (Array.isArray(value) || (prototype !== Object.prototype && prototype !== null)) throw new Error(`${label}の形式が正しくありません。`);
  const allowed = new Set([...required, ...optional]);
  if (required.some(key => !own(value, key)) || Reflect.ownKeys(value).some(key => !allowed.has(key))) throw new Error(`${label}に必要な項目がないか、不明な項目があります。`);
}
export function daysInMonth(month) {
  if (typeof month !== 'string' || !/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error('対象月はYYYY-MM形式で指定してください。');
  const year = Number(month.slice(0, 4));
  if (year === 0) throw new Error('対象年は0001〜9999年で指定してください。');
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  return [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][Number(month.slice(5)) - 1];
}
function time(value, label) {
  if (typeof value !== 'string' || !/^(?:[01]?\d|2[0-3]):[0-5]\d$/.test(value)) throw new Error(`${label}は09:00のような時刻で指定してください。`);
  return value.padStart(5, '0');
}
function defaults(start = '09:00', end = '18:00') {
  const result = { start: time(start, '既定の開始時刻'), end: time(end, '既定の終了時刻') };
  if (result.start >= result.end) throw new Error('既定の終了時刻は開始時刻より後にしてください。');
  return result;
}
function normalizeEntry(entry, count, fallback, acceptFlags = false) {
  checkObject(entry, ENTRY_KEYS, acceptFlags ? FLAGS : [], '勤務希望');
  if (!Number.isInteger(entry.day) || entry.day < 1 || entry.day > count) throw new Error(`日付は対象月の1〜${count}日の整数で指定してください。`);
  if (!['available', 'unavailable'].includes(entry.status)) throw new Error('勤務可否が正しくありません。');
  if (typeof entry.start !== 'string' || typeof entry.end !== 'string') throw new Error('開始・終了時刻は文字列で指定してください。');
  for (const key of FLAGS) if (own(entry, key) && typeof entry[key] !== 'boolean') throw new Error('時刻の補完情報が正しくありません。');
  if (entry.status === 'unavailable') {
    if (entry.start !== '' || entry.end !== '') throw new Error('勤務不可の日には時刻を指定できません。');
    if (entry.startDefaulted === true || entry.endDefaulted === true) throw new Error('勤務不可の日には時刻の補完情報を指定できません。');
    return { day: entry.day, status: 'unavailable', start: '', end: '', startDefaulted: false, endDefaulted: false };
  }
  const startEmpty = entry.start === '', endEmpty = entry.end === '';
  if ((startEmpty && entry.startDefaulted === false) || (endEmpty && entry.endDefaulted === false)) throw new Error('空の時刻と補完情報が矛盾しています。');
  const start = startEmpty ? fallback.start : time(entry.start, '開始時刻'), end = endEmpty ? fallback.end : time(entry.end, '終了時刻');
  if (start >= end) throw new Error(`${entry.day}日の終了時刻は開始時刻より後にしてください。`);
  return { day: entry.day, status: 'available', start, end, startDefaulted: own(entry, 'startDefaulted') ? entry.startDefaulted : startEmpty, endDefaulted: own(entry, 'endDefaulted') ? entry.endDefaulted : endEmpty };
}
function normalizeEntries(entries, count, fallback, acceptFlags = false) {
  if (!Array.isArray(entries) || entries.length > count) throw new Error(`勤務希望は${count}件以内の配列で指定してください。`);
  const seen = new Set();
  return entries.map(entry => { const result = normalizeEntry(entry, count, fallback, acceptFlags); if (seen.has(result.day)) throw new Error(`${result.day}日の勤務希望が重複しています。`); seen.add(result.day); return result; }).sort((a, b) => a.day - b.day);
}
export function normalizeMonthlyProposal(raw, { month, defaultStart = '09:00', defaultEnd = '18:00' } = {}) {
  const count = daysInMonth(month), fallback = defaults(defaultStart, defaultEnd);
  checkObject(raw, ['needsClarification', 'explanation', 'entries'], [], 'AI提案');
  if (typeof raw.needsClarification !== 'boolean') throw new Error('追加確認の要否が正しくありません。');
  if (typeof raw.explanation !== 'string' || raw.explanation.length > 2000) throw new Error('説明は2000文字以内の文字列で指定してください。');
  const entries = normalizeEntries(raw.entries, count, fallback);
  if (raw.needsClarification && entries.length) throw new Error('追加確認が必要な提案には勤務希望を含められません。');
  if (!raw.needsClarification && entries.length === 0) throw new Error('勤務希望がありません。希望日を具体的に指定してください。');
  return { needsClarification: raw.needsClarification, explanation: raw.explanation.trim(), entries };
}
export function applyMonthlyEntries(existing, entries, { month, source = 'message', text = '' } = {}) {
  const count = daysInMonth(month);
  if (typeof source !== 'string' || !source.trim() || source.length > 100) throw new Error('入力元は100文字以内の文字列で指定してください。');
  if (typeof text !== 'string' || text.length > 1500) throw new Error('元のメッセージは1500文字以内にしてください。');
  const prototype = existing !== null && typeof existing === 'object' ? Object.getPrototypeOf(existing) : undefined;
  if (Array.isArray(existing) || (prototype !== Object.prototype && prototype !== null)) throw new Error('既存の希望表の形式が正しくありません。');
  const fallback = defaults();
  for (const key of Reflect.ownKeys(existing)) {
    if (typeof key !== 'string' || !/^[1-9]\d?$/.test(key) || Number(key) > count) throw new Error('既存の希望表に不正な日付があります。');
    const record = existing[key];
    checkObject(record, ['status', 'start', 'end'], [...FLAGS, 'source', 'text'], '既存の勤務希望');
    normalizeEntry({ day: Number(key), status: record.status, start: record.start, end: record.end, ...(own(record, 'startDefaulted') ? { startDefaulted: record.startDefaulted } : {}), ...(own(record, 'endDefaulted') ? { endDefaulted: record.endDefaulted } : {}) }, count, fallback, true);
    if (own(record, 'source') && (typeof record.source !== 'string' || record.source.length > 100)) throw new Error('既存の入力元が正しくありません。');
    if (own(record, 'text') && (typeof record.text !== 'string' || record.text.length > 1500)) throw new Error('既存のメッセージが正しくありません。');
  }
  const validated = normalizeEntries(entries, count, fallback, true), next = structuredClone(existing);
  for (const entry of validated) { const { day, ...record } = entry; next[String(day)] = { ...record, source: source.trim(), text }; }
  return next;
}
export function makeWeekdayProposal({ month, weekdays, defaultStart = '09:00', defaultEnd = '18:00' } = {}) {
  const count = daysInMonth(month), fallback = defaults(defaultStart, defaultEnd);
  if (!Array.isArray(weekdays) || weekdays.length === 0 || weekdays.some(day => !Number.isInteger(day) || day < 0 || day > 6)) throw new Error('基本曜日を1つ以上選んで保存してください。');
  const selected = new Set(weekdays), entries = [];
  for (let day = 1; day <= count; day++) { const weekday = new Date(`${month}-${String(day).padStart(2, '0')}T12:00:00.000Z`).getUTCDay(); if (selected.has(weekday)) entries.push({ day, status: 'available', start: fallback.start, end: fallback.end }); }
  return { needsClarification: false, explanation: '選択した基本曜日から勤務希望の候補を作成しました。', entries };
}

// A fully matched date-only message is unambiguous; avoid letting the model invent dates.
export function parseSimpleDateList(text, options) {
  if (typeof text !== 'string') return null;
  const normalized = text.normalize('NFKC').trim();
  const match = normalized.match(/^(\d{1,2}日?(?:\s*[,、・と]\s*\d{1,2}日?)*)(?:\s*[,、])?\s*(?:でお願いします|お願いします|で[ー〜～]?|希望です|希望|に入れます|入れます)?[!。\s]*$/);
  if (!match) return null;
  const days = [...new Set(match[1].match(/\d+/g).map(Number))];
  return normalizeMonthlyProposal({ needsClarification: false, explanation: '日付の一覧として読み取り、基本時間を補いました。', entries: days.map(day => ({ day, status: 'available', start: '', end: '' })) }, options);
}
