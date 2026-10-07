import { makeWeekdayProposal, normalizeMonthlyProposal, parseSimpleDateList } from './monthly-core.js';

const guidance = '公開版は簡易読み取りです。日付と具体的な時刻に書き換えるか、ローカル版でお試しください。';
const timeToken = '(?:\\d{1,2}(?::\\d{2}|時(?:\\d{1,2}分)?)?)';
const suffix = '(?:入れます|でお願いします|です)?';
const range = new RegExp(`^(${timeToken})(?:から|〜|~|-)(${timeToken})(?:まで)?${suffix}$`);
const from = new RegExp(`^(${timeToken})から${suffix}$`);
const until = new RegExp(`^(${timeToken})まで${suffix}$`);
const weekdayClause = /^毎週([日月火水木金土](?:曜日?)?(?:(?:と|、|,|・)[日月火水木金土](?:曜日?)?)*)に入れます$/;

function clock(text) {
  const match = text.match(/^(\d{1,2})(?::(\d{2})|時(?:(\d{1,2})分)?)?$/);
  const hour = Number(match?.[1]), minute = Number(match?.[2] ?? match?.[3] ?? 0);
  if (!match || hour > 23 || minute > 59) throw new Error('時刻が正しくありません。');
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function dateEntry(clause) {
  const match = clause.match(/^(\d{1,2})日(?:は)?(.+)$/);
  if (!match) return null;
  const day = Number(match[1]), content = match[2];
  if (/^(?:お?休み(?:です)?|入れません|勤務不可)$/.test(content)) return { day, status: 'unavailable', start: '', end: '' };
  if (content === '入れます') return { day, status: 'available', start: '', end: '' };
  let start = '', end = '', times;
  if ((times = content.match(range))) { start = clock(times[1]); end = clock(times[2]); }
  else if ((times = content.match(from))) start = clock(times[1]);
  else if ((times = content.match(until))) end = clock(times[1]);
  else return null;
  return { day, status: 'available', start, end };
}

// Deliberately small grammar: every clause must match before any entries are returned.
export function parsePublicRequest(prompt, { month, defaultStart = '09:00', defaultEnd = '18:00' } = {}) {
  const options = { month, defaultStart, defaultEnd };
  const clarification = { needsClarification: true, explanation: guidance, entries: [] };
  try {
    normalizeMonthlyProposal(clarification, options);
    if (typeof prompt !== 'string' || !prompt.trim() || prompt.length > 1500) return clarification;
    const simple = parseSimpleDateList(prompt, options);
    if (simple) return { ...simple, explanation: `公開版の簡易読み取りで日付一覧を確認しました。時刻は基本時間で補いました。` };
    const text = prompt.normalize('NFKC').trim();
    const clauses = text.split(/[。\n]+|[、,](?=\s*\d{1,2}日(?:は)?)/).map(value => value.trim()).filter(Boolean);
    if (clauses.length === 0) return clarification;
    let weekly = null;
    const explicit = new Map();
    for (const part of clauses) {
      const clause = part.replace(/!+$/, '').trim().replace(/[ \t]+/g, '');
      const weekdays = clause.match(weekdayClause);
      if (weekdays) {
        if (weekly) return clarification;
        const selected = weekdays[1].split(/[と、,・]/).map(value => '日月火水木金土'.indexOf(value[0]));
        weekly = makeWeekdayProposal({ ...options, weekdays: selected }).entries.map(entry => ({ ...entry, start: '', end: '' }));
        continue;
      }
      const entry = dateEntry(clause);
      if (!entry || explicit.has(entry.day)) return clarification;
      explicit.set(entry.day, entry);
    }
    const byDay = new Map((weekly ?? []).map(entry => [entry.day, entry]));
    for (const [day, entry] of explicit) byDay.set(day, entry);
    return normalizeMonthlyProposal({ needsClarification: false, explanation: '公開版の簡易読み取りで明記された希望を確認しました。空欄の時刻は基本時間で補います。', entries: [...byDay.values()] }, options);
  } catch {
    return clarification;
  }
}
