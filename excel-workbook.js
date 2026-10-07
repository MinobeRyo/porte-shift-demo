import { daysInMonth } from './monthly-core.js';
import { parseExcelTime } from './excel-core.js';

const headers = ['対象月', '名前', '日', '可否', '開始', '終了'];
const rosterHeaders = ['名前', '経験', '区分', '勤務量', '基本曜日', '基本開始', '基本終了'];
let library;
async function excel() {
  library ||= import('./exceljs.min.js').then(() => globalThis.ExcelJS);
  const result = await library;
  if (!result?.Workbook) throw new Error('Excelの処理を読み込めませんでした。ページを再読み込みしてください。');
  return result;
}
function simpleValue(value) {
  if (value == null || typeof value === 'string' || typeof value === 'number' || value instanceof Date) return value;
  if (value.richText) return value.richText.map(part => part.text).join('');
  if (value.formula || value.sharedFormula) {
    if (value.result == null) throw new Error('計算結果が保存されていない数式があります。Excelで再計算・保存してから読み込んでください。');
    return simpleValue(value.result);
  }
  if (typeof value.text === 'string' && value.hyperlink) return value.text;
  throw new Error('エラー値など、読み取れないセルがあります。Excelで内容を確認してください。');
}
export function worksheetRows(sheet) {
  if (!sheet || sheet.rowCount > 1000 || sheet.columnCount > 40) throw new Error('取り込むシートは1000行・40列以内にしてください。');
  const rows = [];
  for (let row = 1; row <= sheet.rowCount; row++) {
    const values = [];
    for (let col = 1; col <= sheet.columnCount; col++) {
      try { values.push(simpleValue(sheet.getCell(row, col).value)); }
      catch (error) { throw new Error(`${sheet.name} ${sheet.getCell(row, col).address}: ${error.message}`); }
    }
    rows.push(values);
  }
  return rows;
}
// Check ZIP's declared expanded size before loading an XLSX into memory.
function checkArchive(buffer) {
  if (buffer.byteLength > 5 * 1024 * 1024) throw new Error('Excelファイルは5MB以内にしてください。');
  const data = new DataView(buffer);
  let end = -1;
  for (let p = data.byteLength - 22; p >= Math.max(0, data.byteLength - 65557); p--) if (data.getUint32(p, true) === 0x06054b50) { end = p; break; }
  if (end < 0) throw new Error('通常の.xlsxファイルを選んでください。暗号化・パスワード付きのファイルには対応していません。');
  const count = data.getUint16(end + 10, true);
  if (count > 2000 || count === 65535) throw new Error('ファイルの構成が大きすぎます。必要なシートだけを別ファイルに保存してください。');
  let pos = data.getUint32(end + 16, true), total = 0;
  for (let i = 0; i < count; i++) {
    if (pos + 46 > data.byteLength || data.getUint32(pos, true) !== 0x02014b50) throw new Error('Excelファイルの構成が正しくありません。');
    total += data.getUint32(pos + 24, true);
    if (total > 30 * 1024 * 1024) throw new Error('展開後の内容が大きすぎます。必要なシートだけを別ファイルに保存してください。');
    pos += 46 + data.getUint16(pos + 28, true) + data.getUint16(pos + 30, true) + data.getUint16(pos + 32, true);
  }
}
export async function readExcelFile(file) {
  if (!/\.xlsx$/i.test(file.name)) throw new Error('.xlsx形式を選んでください。.xlsや.csvはExcelで.xlsxに保存し直してください。');
  if (file.size > 5 * 1024 * 1024) throw new Error('Excelファイルは5MB以内にしてください。');
  const buffer = await file.arrayBuffer(); checkArchive(buffer);
  const ExcelJS = await excel(), workbook = new ExcelJS.Workbook();
  try { await workbook.xlsx.load(buffer); }
  catch { throw new Error('Excelを開けませんでした。壊れていない、パスワードなしの.xlsxか確認してください。'); }
  if (!workbook.worksheets.length || workbook.worksheets.length > 40) throw new Error('シート数は1〜40枚にしてください。');
  return workbook;
}
export function readStaffConditions(workbook) {
  const sheet = workbook.getWorksheet('基本条件');
  if (!sheet) return [];
  const rows = worksheetRows(sheet);
  if (!rosterHeaders.every((h, i) => String(rows[0]?.[i] ?? '').trim() === h)) throw new Error('「基本条件」シートの列見出しが対応形式と異なります。');
  const result = [];
  for (let index = 1; index < rows.length; index++) {
    const row = rows[index]; if (row.every(v => v == null || v === '')) continue;
    const [name, experience, category, preference, weekdays] = row.map(v => String(v ?? '').trim());
    const start = parseExcelTime(row[5], index), end = parseExcelTime(row[6], index);
    if (!name || name.length > 40) throw new Error(`基本条件 ${index + 1}行: 名前は1〜40文字で指定してください。`);
    if (!['未設定', '新人', 'ベテラン'].includes(experience) || !['未設定', '学生', '社会人'].includes(category) || !['未設定', '少なめ', '通常', '多め'].includes(preference)) throw new Error(`基本条件 ${index + 1}行: 経験・区分・勤務量を確認してください。`);
    const tokens = weekdays ? weekdays.split(/[・、,\s]+/).filter(Boolean) : [];
    if (tokens.some(v => !/^[日月火水木金土](?:曜(?:日)?)?$/.test(v))) throw new Error(`基本条件 ${index + 1}行: 曜日は「月・火」の形式で指定してください。`);
    if (!/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(start) || !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(end) || start >= end) throw new Error(`基本条件 ${index + 1}行: 基本開始・終了を09:00の形式で指定してください。`);
    if (result.some(p => p.name === name)) throw new Error(`基本条件で名前「${name}」が重複しています。`);
    result.push({ name, experience, category, preference, weekdays: [...new Set(tokens.map(v => '日月火水木金土'.indexOf(v[0])))], start, end });
  }
  return result;
}
function formatTable(sheet, widths) {
  sheet.columns = widths.map(width => ({ width }));
  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.getRow(1).font = { name: 'Yu Gothic', bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF355C96' } };
  sheet.getRow(1).height = 25;
  sheet.eachRow((row, index) => {
    if (index > 1) row.font = { name: 'Yu Gothic', size: 10 };
    row.alignment = { vertical: 'middle', wrapText: true };
    if (index > 1 && index % 2 === 0) row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F5F8' } };
  });
  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: Math.max(1, sheet.rowCount), column: widths.length } };
}
export async function createExcelWorkbook(state, { template = false, staffId = '' } = {}) {
  const month = state.month, count = daysInMonth(month), staff = staffId ? state.staff.filter(s => s.id === staffId) : state.staff;
  if (!staff.length) throw new Error('出力するスタッフを選んでください。');
  const ExcelJS = await excel(), workbook = new ExcelJS.Workbook();
  workbook.creator = 'ポルテ 勤務希望デモ';
  const grid = workbook.addWorksheet('勤務希望表', { pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1 } });
  grid.columns = [{ width: 18 }, ...Array.from({ length: 16 }, () => ({ width: 8 }))];
  grid.mergeCells('A1:Q1'); grid.getCell('A1').value = `${Number(month.slice(0, 4))}年${Number(month.slice(5))}月 勤務希望表（未確定）`;
  grid.getCell('A1').font = { name: 'Yu Gothic', size: 16, bold: true }; grid.getRow(1).height = 30;
  const data = template ? {} : state.months[month] || {};
  let row = 4;
  for (const [from, to] of [[1, Math.min(16, count)], [17, count]]) {
    grid.mergeCells(row, 1, row + 1, 1); grid.getCell(row, 1).value = '名前';
    for (let day = from; day <= to; day++) {
      const col = day - from + 2, weekday = new Date(`${month}-${String(day).padStart(2, '0')}T12:00:00Z`).getUTCDay();
      grid.getCell(row, col).value = day; grid.getCell(row + 1, col).value = '日月火水木金土'[weekday];
      for (const r of [row, row + 1]) grid.getCell(r, col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: weekday === 0 ? 'FFFFE5E5' : weekday === 6 ? 'FFE3EBFA' : 'FFEDEFF3' } };
    }
    const top = row; row += 2;
    staff.forEach((s, index) => {
      grid.mergeCells(row, 1, row + 1, 1); grid.getCell(row, 1).value = s.name;
      for (let day = from; day <= to; day++) {
        const col = day - from + 2, entry = data[s.id]?.[day];
        if (entry?.status === 'unavailable') grid.getCell(row, col).value = '×';
        else if (entry) { grid.getCell(row, col).value = entry.start; grid.getCell(row + 1, col).value = entry.end; }
        if (index % 2 === 1) for (const r of [row, row + 1]) grid.getCell(r, col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } };
      }
      row += 2;
    });
    for (let r = top; r < row; r++) {
      grid.getRow(r).height = 19;
      for (let c = 1; c <= to - from + 2; c++) {
        const cell = grid.getCell(r, c); cell.font = { name: 'Yu Gothic', size: 10 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = Object.fromEntries(['top', 'left', 'bottom', 'right'].map(side => [side, { style: 'thin', color: { argb: 'FFA0A0A0' } }]));
      }
    }
    row += 2;
  }
  grid.pageSetup.printArea = `A1:Q${row - 2}`;
  const records = workbook.addWorksheet('取込用'); records.addRow(headers);
  for (const s of staff) {
    const entries = Object.entries(data[s.id] || {}).sort((a, b) => Number(a[0]) - Number(b[0]));
    if (!entries.length) records.addRow([month, s.name, '', '', '', '']);
    for (const [day, e] of entries) records.addRow([month, s.name, Number(day), e.status === 'available' ? '勤務希望' : '勤務不可', e.start, e.end]);
  }
  formatTable(records, [14, 20, 8, 14, 12, 12]);
  const roster = workbook.addWorksheet('基本条件'); roster.addRow(rosterHeaders);
  staff.forEach(s => roster.addRow([s.name, s.experience, s.category, s.preference, [1, 2, 3, 4, 5, 6, 0].filter(d => s.weekdays.includes(d)).map(d => '日月火水木金土'[d]).join('・'), s.start, s.end]));
  formatTable(roster, [20, 14, 14, 14, 25, 14, 14]);
  const history = workbook.addWorksheet('取り込み履歴'); history.addRow(['記録日時', '対象月', '名前', '入力元', '元の内容', '反映した希望']);
  if (!template) (state.history || []).filter(h => h.month === month && staff.some(s => s.id === h.staffId)).forEach(h => history.addRow([h.at, h.month, h.staffName, h.source, h.text, h.entries.map(e => `${e.day}日 ${e.status === 'unavailable' ? '×' : `${e.start}〜${e.end}`}`).join(' ／ ')]));
  formatTable(history, [27, 14, 20, 14, 55, 65]);
  return workbook;
}
export async function exportExcel(state, options) {
  const workbook = await createExcelWorkbook(state, options);
  return workbook.xlsx.writeBuffer();
}
