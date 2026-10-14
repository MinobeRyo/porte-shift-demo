// Staff basic conditions read from a DM. Only what the text actually says becomes a candidate;
// the manager checks each item before it is saved.
export const PROFILE_FIELDS = ['category', 'experience', 'preference', 'weekdays', 'time'];
const VALUES = { category: ['学生', '社会人'], experience: ['新人', 'ベテラン'], preference: ['少なめ', '通常', '多め'] };
const clock = value => {
  const match = /^(\d{1,2}):(\d{2})$/.exec(value ?? '');
  if (!match || Number(match[1]) > 23 || Number(match[2]) > 59) return '';
  return `${match[1].padStart(2, '0')}:${match[2]}`;
};

/** Validates a raw profile (from the model or the rules). Empty values mean "not mentioned". */
export function normalizeStaffProfile(raw) {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new Error('基本情報の読み取り結果の形式が正しくありません。');
  const fields = {};
  for (const key of ['category', 'experience', 'preference']) {
    if (raw[key] === undefined || raw[key] === '' || raw[key] === null) continue;
    if (!VALUES[key].includes(raw[key])) throw new Error('基本情報の読み取り結果に対応していない値があります。');
    fields[key] = raw[key];
  }
  if (raw.weekdays !== undefined) {
    if (!Array.isArray(raw.weekdays) || raw.weekdays.some(d => !Number.isInteger(d) || d < 0 || d > 6)) throw new Error('基本曜日の読み取り結果が正しくありません。');
    const days = [...new Set(raw.weekdays)].sort((a, b) => a - b);
    if (days.length) fields.weekdays = days;
  }
  const start = clock(raw.start), end = clock(raw.end);
  if (start && end) {
    if (start >= end) throw new Error('基本時間の終了は開始より後にしてください。');
    fields.time = { start, end };
  }
  return { fields, explanation: typeof raw.explanation === 'string' ? raw.explanation.slice(0, 300) : '' };
}

const WEEKDAY = { 日: 0, 月: 1, 火: 2, 水: 3, 木: 4, 金: 5, 土: 6 };
const NEGATIVE = /(無理|むり|入れない|入れません|NG|ダメ|だめ|休み|不可|できない|できません)/;
const toClock = (hour, minute = '00') => `${String(Number(hour)).padStart(2, '0')}:${minute.padStart(2, '0')}`;

/** Simple rules for the public demo and as a fallback when the local AI is unavailable. */
export function parseStaffProfileRules(text) {
  if (typeof text !== 'string') return normalizeStaffProfile({});
  const raw = { weekdays: [] };
  if (/(大学生|高校生|専門学生|短大生|大学院生|学生|大学\d年|高校\d年)/.test(text)) raw.category = '学生';
  else if (/(社会人|会社員|主婦|主夫|パート)/.test(text)) raw.category = '社会人';
  if (/(ベテラン|経験者|経験あり|\d+年(?:以上)?(?:働|勤め|経験))/.test(text)) raw.experience = 'ベテラン';
  else if (/(新人|未経験|初めて|はじめて|入ったばかり)/.test(text)) raw.experience = '新人';
  if (/(多め|多く|たくさん|がっつり|いっぱい|稼ぎたい)/.test(text)) raw.preference = '多め';
  else if (/(少なめ|少しだけ|控えめ|ちょっとだけ|少なく)/.test(text)) raw.preference = '少なめ';
  else if (/(普通|通常|今まで通り)/.test(text)) raw.preference = '通常';
  for (const sentence of text.split(/[。\n！!？?]/)) {
    if (NEGATIVE.test(sentence)) continue;
    if (/平日/.test(sentence)) raw.weekdays.push(1, 2, 3, 4, 5);
    if (/(土日|週末)/.test(sentence)) raw.weekdays.push(6, 0);
    for (const match of sentence.matchAll(/([月火水木金土日])曜/g)) raw.weekdays.push(WEEKDAY[match[1]]);
    // 「月水金」のように曜日の文字だけが並ぶ書き方（3文字以上）
    for (const match of sentence.matchAll(/(?<![0-9０-９])([月火水木金土日]{3,7})(?![0-9０-９曜日])/g)) for (const ch of match[1]) raw.weekdays.push(WEEKDAY[ch]);
  }
  const time = /(\d{1,2})(?:[:：](\d{2}))?時?(?:半)?\s*(?:から|〜|~|-|ー|－)\s*(\d{1,2})(?:[:：](\d{2}))?時?/.exec(text);
  if (time && !/(\d{1,2})日\s*(?:は|の)?\s*$/.test(text.slice(0, time.index))) { raw.start = toClock(time[1], time[2]); raw.end = toClock(time[3], time[4]); }
  try { return normalizeStaffProfile(raw); }
  catch { return normalizeStaffProfile({ ...raw, start: '', end: '' }); }
}
