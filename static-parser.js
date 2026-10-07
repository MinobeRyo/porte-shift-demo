import { makeWeekdayProposal, normalizeMonthlyProposal, parseSimpleDateList } from './monthly-core.js';

const guidance = '日付と具体的な時刻、または曜日を確認できませんでした。「1、3、4日は10時から16時」のように書き換えてください。';
const timeToken = '(?:\\d{1,2}(?::\\d{2}|時(?:\\d{1,2}分)?)?)';
const suffix = '(?:入れます|でお願いします|です)?';
const range = new RegExp(`^(${timeToken})(?:から|〜|~|-)(${timeToken})(?:まで)?${suffix}$`);
const from = new RegExp(`^(${timeToken})から${suffix}$`);
const until = new RegExp(`^(${timeToken})まで${suffix}$`);
const dayAtom = '\\d{1,2}日?(?:(?:〜|~|-|から)\\d{1,2}日?)?';
const dateClause = new RegExp(`^(${dayAtom}(?:(?:、|,|・|と)${dayAtom})*)(?:は|に)?(.*)$`);
const weekdayClause = /^(?:毎週)?(土日|[日月火水木金土](?:曜日?)?(?:(?:と|、|,|・)[日月火水木金土](?:曜日?)?)*)(?:なら|は|に)?(.+)$/;

// Callers may also use this before an AI request: uncertainty must stay a question.
export function requestClarificationReason(prompt) {
  if (typeof prompt !== 'string') return null;
  const text = prompt.normalize('NFKC');
  if (/(?:午前|午後|夕方|夜|朝|昼)/.test(text)) return '「午後」などの表現だけでは勤務時間を確定できません。開始・終了時刻を具体的に指定してください。';
  if (/(?:再来週|来週|今週|再来月|来月|今月|明後日|あさって|明日|あした|今日|週末)/.test(text)) return '相対的な日付の基準が不明です。対象月の日付、または「毎週土曜」のような曜日を指定してください。';
  if (/(?:たぶん|多分|おそらく|相談|未定|かも|できれば|可能なら|くらい|頃|ごろ)/.test(text)) return '未確定の希望が含まれています。勤務できる日と時刻が決まってから入力してください。';
  return null;
}

function clock(text) {
  const match = text.match(/^(\d{1,2})(?::(\d{2})|時(?:(\d{1,2})分)?)?$/);
  const hour = Number(match?.[1]), minute = Number(match?.[2] ?? match?.[3] ?? 0);
  if (!match || hour > 23 || minute > 59) throw new Error('時刻が正しくありません。0〜23時・0〜59分で指定してください。');
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function availability(content) {
  if (/^(?:お?休み(?:です)?|入れません|勤務不可)$/.test(content)) return { status: 'unavailable', start: '', end: '' };
  if (/^(?:入れます|でお願いします|お願いします|で[ー〜~]?|希望です|希望|です)?$/.test(content)) return { status: 'available', start: '', end: '' };
  let start = '', end = '', times;
  if ((times = content.match(range))) { start = clock(times[1]); end = clock(times[2]); }
  else if ((times = content.match(from))) start = clock(times[1]);
  else if ((times = content.match(until))) end = clock(times[1]);
  else return null;
  return { status: 'available', start, end };
}

function dateEntries(clause) {
  const match = clause.match(dateClause);
  if (!match || (!match[2] && /[はに]$/.test(clause))) return null;
  const value = availability(match[2]);
  if (!value) return null;
  const days = new Set();
  for (const atom of match[1].split(/[、,・と]/)) {
    const [, first, last] = atom.match(/^(\d{1,2})日?(?:(?:〜|~|-|から)(\d{1,2})日?)?$/);
    const start = Number(first), end = Number(last ?? first);
    if (start > end) throw new Error('日付の範囲が逆になっています。開始日を終了日より前にしてください。');
    for (let day = start; day <= end; day++) days.add(day);
  }
  return [...days].map(day => ({ day, ...value }));
}

function parseUnit(part, options) {
  const clause = part.replace(/!+$/, '').trim().replace(/[ \t]+/g, '');
  const weekdays = clause.match(weekdayClause);
  if (weekdays) {
    const value = availability(weekdays[2]);
    if (!value) return null;
    const selected = weekdays[1] === '土日' ? [6, 0] : weekdays[1].split(/[と、,・]/).map(day => '日月火水木金土'.indexOf(day[0]));
    return { weekly: true, entries: makeWeekdayProposal({ ...options, weekdays: selected }).entries.map(entry => ({ day: entry.day, ...value })) };
  }
  const entries = dateEntries(clause);
  return entries ? { weekly: false, entries } : null;
}

// A comma may join dates sharing one time, or separate complete requests.
// Keep the longest complete request together, and require every remaining part to match.
function splitRequests(sentence, options, memo = new Map()) {
  if (memo.has(sentence)) return memo.get(sentence);
  const whole = parseUnit(sentence, options);
  if (whole) return [whole];
  for (let index = sentence.length - 1; index >= 0; index--) {
    if (!/[、,]/.test(sentence[index])) continue;
    const first = parseUnit(sentence.slice(0, index), options);
    if (!first) continue;
    const rest = splitRequests(sentence.slice(index + 1), options, memo);
    if (rest) return [first, ...rest];
  }
  memo.set(sentence, null);
  return null;
}

// Deliberately small grammar: every clause must match before any entries are returned.
export function parsePublicRequest(prompt, { month, defaultStart = '09:00', defaultEnd = '18:00' } = {}) {
  const options = { month, defaultStart, defaultEnd };
  const clarify = explanation => ({ needsClarification: true, explanation, entries: [] });
  try {
    normalizeMonthlyProposal(clarify(guidance), options);
    if (typeof prompt !== 'string' || !prompt.trim()) return clarify('勤務希望の日付または曜日を入力してください。');
    if (prompt.length > 1500) return clarify('勤務希望は1500文字以内で入力してください。');
    const reason = requestClarificationReason(prompt);
    if (reason) return clarify(reason);
    const simple = parseSimpleDateList(prompt, options);
    if (simple) return { ...simple, explanation: '日付一覧を確認しました。時刻は基本時間で補いました。' };
    const sentences = prompt.normalize('NFKC').trim().split(/[。\n]+/).map(value => value.trim()).filter(Boolean);
    if (sentences.length === 0) return clarify(guidance);
    let weekly = null;
    const explicit = new Map();
    for (const sentence of sentences) {
      const units = splitRequests(sentence, options);
      if (!units) return clarify(guidance);
      for (const unit of units) {
        if (unit.weekly) {
          if (weekly) return clarify('曜日の指定が複数に分かれています。「毎週火曜と土曜に入れます」のように1つにまとめてください。');
          weekly = unit.entries;
          continue;
        }
        for (const entry of unit.entries) {
          if (explicit.has(entry.day)) return clarify(`${entry.day}日の希望が複数あります。勤務可否と時間を1つにまとめてください。`);
          explicit.set(entry.day, entry);
        }
      }
    }
    const byDay = new Map((weekly ?? []).map(entry => [entry.day, entry]));
    for (const [day, entry] of explicit) byDay.set(day, entry);
    return normalizeMonthlyProposal({ needsClarification: false, explanation: '明記された希望を確認しました。曜日は対象月の日付に展開し、空欄の時刻は基本時間で補います。', entries: [...byDay.values()] }, options);
  } catch (error) {
    return clarify(error instanceof Error ? error.message : guidance);
  }
}
