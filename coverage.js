import { daysInMonth } from './monthly-core.js';

const own = (value, key) => Object.prototype.hasOwnProperty.call(value, key);
function object(value, label) {
  const prototype = value !== null && typeof value === 'object' ? Object.getPrototypeOf(value) : undefined;
  if (Array.isArray(value) || (prototype !== Object.prototype && prototype !== null)) throw new Error(`${label}の形式が正しくありません。`);
}
function minutes(value, label) {
  if (typeof value !== 'string' || !/^(?:[01]?\d|2[0-3]):[0-5]\d$/.test(value)) throw new Error(`${label}は09:00のような時刻で指定してください。`);
  const [hour, minute] = value.split(':').map(Number);
  return hour * 60 + minute;
}
const clock = value => `${String(Math.floor(value / 60)).padStart(2, '0')}:${String(value % 60).padStart(2, '0')}`;

/**
 * 勤務希望に基づく参考集計。確定配置・休憩・優先順位は扱わない。
 * availableCount / veteranCount は当日の希望登録人数（対象時間外も含む）。
 * minimumAvailable / shortages は settings.start 以上 settings.end 未満の同時人数。
 * 未登録は勤務不可と区別する。引数は変更しない。
 */
export function summarizeCoverage({ month, staff, requests, settings } = {}) {
  const count = daysInMonth(month);
  if (!Array.isArray(staff)) throw new Error('スタッフは配列で指定してください。');
  object(requests, '勤務希望一覧');
  object(settings, '必要人数の設定');
  if (!Number.isInteger(settings.required) || settings.required < 0 || settings.required > 30) throw new Error('必要人数は0〜30人の整数で指定してください。');
  const from = minutes(settings.start, '対象の開始時刻'), to = minutes(settings.end, '対象の終了時刻');
  if (from >= to) throw new Error('対象の終了時刻は開始時刻より後にしてください。');
  const ids = new Set();
  for (const person of staff) {
    object(person, 'スタッフ');
    if (typeof person.id !== 'string' || !person.id.trim() || ids.has(person.id)) throw new Error('スタッフIDが空か重複しています。');
    if (typeof person.name !== 'string' || !person.name.trim() || typeof person.experience !== 'string' || !person.experience.trim()) throw new Error('スタッフの名前・経験を指定してください。');
    ids.add(person.id);
  }
  const validated = new Map();
  for (const id of Reflect.ownKeys(requests)) {
    if (!ids.has(id)) throw new Error('勤務希望に未登録のスタッフIDがあります。');
    const entries = requests[id];
    object(entries, 'スタッフの勤務希望');
    const days = new Map();
    for (const day of Reflect.ownKeys(entries)) {
      if (typeof day !== 'string' || !/^[1-9]\d?$/.test(day) || Number(day) > count) throw new Error('勤務希望に対象月外または不正な日付があります。');
      const entry = entries[day];
      object(entry, '勤務希望');
      if (!['available', 'unavailable'].includes(entry.status) || !own(entry, 'start') || !own(entry, 'end')) throw new Error('勤務希望の可否・時刻が正しくありません。');
      if (entry.status === 'unavailable') {
        if (entry.start !== '' || entry.end !== '') throw new Error('勤務不可の日には時刻を指定できません。');
        days.set(Number(day), { status: 'unavailable' });
      } else {
        const start = minutes(entry.start, `${day}日の開始時刻`), end = minutes(entry.end, `${day}日の終了時刻`);
        if (start >= end) throw new Error(`${day}日の終了時刻は開始時刻より後にしてください。`);
        days.set(Number(day), { status: 'available', start, end });
      }
    }
    validated.set(id, days);
  }

  return Array.from({ length: count }, (_, index) => {
    const day = index + 1, intervals = [];
    let availableCount = 0, unreportedCount = 0, veteranCount = 0;
    for (const person of staff) {
      const entry = validated.get(person.id)?.get(day);
      if (!entry) { unreportedCount++; continue; }
      if (entry.status !== 'available') continue;
      availableCount++;
      if (person.experience === 'ベテラン') veteranCount++;
      const start = Math.max(from, entry.start), end = Math.min(to, entry.end);
      if (start < end) intervals.push({ start, end });
    }
    const boundaries = [...new Set([from, to, ...intervals.flatMap(entry => [entry.start, entry.end])])].sort((a, b) => a - b);
    const shortages = [];
    let minimumAvailable = Infinity;
    for (let i = 0; i < boundaries.length - 1; i++) {
      const start = boundaries[i], end = boundaries[i + 1];
      const available = intervals.filter(entry => entry.start <= start && entry.end > start).length;
      minimumAvailable = Math.min(minimumAvailable, available);
      if (available >= settings.required) continue;
      const previous = shortages.at(-1);
      if (previous && previous.end === clock(start) && previous.available === available) previous.end = clock(end);
      else shortages.push({ start: clock(start), end: clock(end), available, missing: settings.required - available });
    }
    return { day, availableCount, unreportedCount, minimumAvailable, shortages, veteranCount };
  });
}
