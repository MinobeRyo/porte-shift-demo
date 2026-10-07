import { daysInMonth } from './monthly-core.js';

const HEADERS = ['対象月', '名前', '日', '可否', '開始', '終了'];
const OFF = new Set(['×', '✕', '休', '勤務不可', 'unavailable']);
const ON = new Set(['○', '〇', '勤務希望', 'available']);
const blank = value => value === undefined || value === null || (typeof value === 'string' && !value.trim());
const text = value => typeof value === 'string' ? value.normalize('NFKC').trim() : String(value ?? '');
const nonempty = row => row.some(value => !blank(value));
const fail = (row, message) => { throw new Error(`${row + 1}行目：${message}`); };

function monthValue(value, row) {
  if (value instanceof Date) {
    const year = value.getUTCFullYear();
    const result = `${String(year).padStart(4, '0')}-${String(value.getUTCMonth() + 1).padStart(2, '0')}`;
    daysInMonth(result);
    return result;
  }
  const match = text(value).match(/^(\d{4})(?:-(\d{1,2})|年\s*(\d{1,2})月)$/);
  if (!match) fail(row, '対象月は2026-08または2026年8月の形で指定してください。');
  const result = `${match[1]}-${String(Number(match[2] || match[3])).padStart(2, '0')}`;
  daysInMonth(result);
  return result;
}
function titleMonth(value, row) {
  if (typeof value !== 'string') return null;
  const match = text(value).match(/^(\d{4}年\s*\d{1,2}月)\s*(?:シフト表|勤務希望表)(?:\s*[（(]未確定[）)])?$/);
  return match ? monthValue(match[1], row) : null;
}
function dayValue(value, month, row) {
  const normalized = typeof value === 'number' ? value : Number(text(value).replace(/日$/, ''));
  if (!/^[1-9]\d?(?:日)?$/.test(text(value)) || !Number.isInteger(normalized) || normalized > daysInMonth(month)) fail(row, '対象月に存在する日を整数で指定してください。');
  return normalized;
}
function nameValue(value, row) {
  if (typeof value !== 'string' || !value.trim() || [...value.trim()].length > 40) fail(row, '名前は1〜40文字の文字列で指定してください。');
  return value.trim();
}
function timeValue(value, row) {
  if (blank(value)) return '';
  let hour, minute;
  if (value instanceof Date) {
    hour = value.getUTCHours(); minute = value.getUTCMinutes();
  } else if (typeof value === 'number') {
    if (value < 0 || value >= 1) fail(row, '数値の時刻は0以上1未満のExcel時刻で指定してください。');
    const total = Math.round(value * 1440);
    hour = Math.floor(total / 60); minute = total % 60;
  } else {
    const match = text(value).match(/^(\d{1,2})(?::(\d{2})|時(?:\s*(\d{1,2})分?)?)$/);
    if (!match) fail(row, `「${text(value).slice(0, 30)}」を時刻として読み取れません。`);
    hour = Number(match[1]); minute = Number(match[2] || match[3] || 0);
  }
  if (hour > 23 || minute > 59) fail(row, '時刻は00:00〜23:59で指定してください。');
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}
export { timeValue as parseExcelTime };
function entryValue(day, status, start, end, row) {
  if (!ON.has(text(status)) && !OFF.has(text(status))) fail(row, '可否は勤務希望または勤務不可で指定してください。');
  if (OFF.has(text(status))) {
    if (!blank(start) || !blank(end)) fail(row, '勤務不可の日に時刻が入っています。');
    return { day, status: 'unavailable', start: '', end: '' };
  }
  const entry = { day, status: 'available', start: timeValue(start, row), end: timeValue(end, row) };
  if (entry.start && entry.end && entry.start >= entry.end) fail(row, '終了時刻は開始時刻より後にしてください。');
  return entry;
}
function gridEntry(day, upper, lower, row) {
  if (blank(upper) && blank(lower)) return null;
  const cells = [upper, lower].filter(value => !blank(value));
  if (cells.some(value => OFF.has(text(value)))) {
    if (!cells.every(value => OFF.has(text(value)))) fail(row, `${day}日の勤務不可と時刻・勤務希望が矛盾しています。`);
    return entryValue(day, '勤務不可', '', '', row);
  }
  return entryValue(day, '勤務希望', ON.has(text(upper)) ? '' : upper, ON.has(text(lower)) ? '' : lower, row);
}

/** Excelライブラリ非依存の変換。全体を検証してから返し、元データは変更しない。 */
export function parseWorksheetRows(rows, { fallbackMonth } = {}) {
  if (!Array.isArray(rows) || !rows.length || rows.length > 1000) throw new Error('Excelシートは1〜1000行で指定してください。');
  rows.forEach((row, index) => {
    if (!Array.isArray(row) || row.length > 40) fail(index, '1行は40列以内の配列で指定してください。');
    for (const value of row) {
      if (blank(value)) continue;
      if (typeof value === 'string') continue;
      if (typeof value === 'number' && Number.isFinite(value)) continue;
      if (value instanceof Date && Number.isFinite(value.getTime())) continue;
      fail(index, 'セルには文字列・数値・日時・空欄だけを指定してください。数式などのオブジェクトは取り込めません。');
    }
  });
  const people = new Map(), warnings = [];
  let month;
  const useMonth = (value, row) => {
    if (month && value !== month) fail(row, '異なる対象月が混在しています。1か月ごとに取り込んでください。');
    month = value;
  };
  const person = (name, row) => {
    if (!people.has(name)) {
      if (people.size >= 30) fail(row, 'スタッフは30名以内にしてください。');
      people.set(name, new Map());
    }
    return people.get(name);
  };
  const add = (name, entry, row) => {
    const days = person(name, row);
    if (!entry) return;
    if (days.has(entry.day)) fail(row, `${name}さんの${entry.day}日が重複しています。`);
    days.set(entry.day, entry);
    if (entry.status === 'available' && (!entry.start || !entry.end)) warnings.push(`${name}さんの${entry.day}日：${!entry.start && !entry.end ? '開始・終了' : !entry.start ? '開始' : '終了'}時刻がありません。基本時間で補完するため、反映前に確認してください。`);
  };
  const records = HEADERS.every((header, index) => text(rows[0][index]) === header);
  if (records) {
    rows.forEach((row, index) => {
      if (row.slice(6).some(value => !blank(value))) fail(index, '取込用表に未対応の列があります。');
      if (index === 0 || !nonempty(row)) return;
      if (blank(row[0])) fail(index, '取込用表では各行の対象月が必要です。');
      useMonth(monthValue(row[0], index), index);
      const name = nameValue(row[1], index);
      if (blank(row[2])) {
        if (row.slice(3, 6).some(value => !blank(value))) fail(index, '日付が空の行に可否・時刻があります。');
        person(name, index);
        return;
      }
      const day = dayValue(row[2], month, index);
      add(name, entryValue(day, row[3], row[4], row[5], index), index);
    });
    if (!month) throw new Error('取込用表に対象月と勤務希望の行がありません。');
  } else {
    rows.forEach((row, index) => row.forEach(value => { const found = titleMonth(value, index); if (found) useMonth(found, index); }));
    if (!month) {
      if (blank(fallbackMonth)) throw new Error('シフト表から対象月を取得できません。対象月を選択してください。');
      month = monthValue(fallbackMonth, 0);
      warnings.push(`表から対象月を取得できないため、選択中の${month}を使用しました。日付と月をご確認ください。`);
    }
    let header = null, foundHeader = false;
    const headerAt = row => row.findIndex(value => text(value) === '名前');
    const titleRow = row => row.some(value => titleMonth(value, 0));
    const checkOtherColumns = (row, permitted, index) => row.forEach((value, column) => { if (!permitted.has(column) && !blank(value)) fail(index, '日付・名前以外の列に読み取れない値があります。'); });
    for (let index = 0; index < rows.length; index++) {
      const row = rows[index];
      if (!nonempty(row)) continue;
      if (titleRow(row)) {
        if (row.some(value => !blank(value) && !titleMonth(value, index))) fail(index, 'タイトル行に読み取れない値があります。');
        header = null;
        continue;
      }
      const nameColumn = headerAt(row);
      if (nameColumn >= 0) {
        const days = new Map(), seen = new Set();
        row.forEach((value, column) => {
          if (column === nameColumn || blank(value)) return;
          if (column < nameColumn) fail(index, '名前より前の列に読み取れない値があります。');
          const day = dayValue(value, month, index);
          if (seen.has(day)) fail(index, 'ヘッダーの日付が重複しています。');
          days.set(column, day); seen.add(day);
        });
        if (!days.size) fail(index, '名前の右側に日付のヘッダーが必要です。');
        header = { nameColumn, days, permitted: new Set([nameColumn, ...days.keys()]) };
        foundHeader = true;
        const weekdayRow = rows[index + 1];
        if (weekdayRow && [...days.keys()].some(column => !blank(weekdayRow[column])) && [...days.keys()].every(column => blank(weekdayRow[column]) || /^[（(]?[日月火水木金土](?:曜日)?[）)]?$/.test(text(weekdayRow[column]))) && (blank(weekdayRow[nameColumn]) || ['名前', '曜日'].includes(text(weekdayRow[nameColumn])))) {
          checkOtherColumns(weekdayRow, header.permitted, index + 1);
          index++;
        }
        continue;
      }
      if (!header) fail(index, '名前・日付のヘッダーより前に読み取れない値があります。');
      checkOtherColumns(row, header.permitted, index);
      const name = nameValue(row[header.nameColumn], index);
      const next = rows[index + 1];
      const paired = next && !titleRow(next) && headerAt(next) < 0 && (blank(next[header.nameColumn]) || text(next[header.nameColumn]) === text(name));
      const lower = paired ? next : [];
      if (paired) checkOtherColumns(lower, header.permitted, index + 1);
      person(name, index);
      for (const [column, day] of header.days) add(name, gridEntry(day, row[column], lower[column], index), index);
      if (paired) index++;
    }
    if (!foundHeader) throw new Error('対応するシフト表のヘッダーが見つかりません。');
  }
  if (!people.size) throw new Error('取り込めるスタッフがありません。');
  return { month, staff: [...people].map(([name, days]) => ({ name, entries: [...days.values()].sort((a, b) => a.day - b.day) })), warnings, format: records ? 'records' : 'split-grid' };
}
