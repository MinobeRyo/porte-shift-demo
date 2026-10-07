import { daysInMonth, normalizeMonthlyProposal, applyMonthlyEntries, makeWeekdayProposal } from './monthly-core.js';
import { PUBLIC_DEMO } from './runtime-config.js';
import { parsePublicRequest } from './static-parser.js';
import { summarizeCoverage } from './coverage.js';
import { readExcelFile, worksheetRows, readStaffConditions, exportExcel } from './excel-workbook.js';
import { parseWorksheetRows } from './excel-core.js';
import { buildExcelImportPlan } from './excel-plan.js';

const KEY = PUBLIC_DEMO ? 'porte-pages-demo-v1' : 'porte-monthly-requests-v1';
const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const labels = { experience: ['未設定', '新人', 'ベテラン'], category: ['未設定', '学生', '社会人'], preference: ['未設定', '少なめ', '通常', '多め'] };
const createStaff = (name, id = crypto.randomUUID()) => ({ id, name, experience: '未設定', category: '未設定', preference: '未設定', weekdays: [], start: '09:00', end: '18:00' });
const initial = () => ({ version: 1, month: '2026-10', staff: ['スタッフA', 'スタッフB', 'スタッフC'].map((name, i) => createStaff(name, `demo-${i}`)), months: {}, history: [] });
let state = initial();
try { const saved = JSON.parse(localStorage.getItem(KEY)); if (saved?.version === 1 && Array.isArray(saved.staff) && saved.staff.length && saved.months && Array.isArray(saved.history)) { daysInMonth(saved.month); state = saved; } } catch { /* Start with a clean demo if the saved format cannot be read. */ }
let view = 'month', selectedStaff = state.staff[0].id, draft = '', candidate = null, busy = false, timer, staffDraft = null;
let excelSession = null, period = 0, applied = null, demoMode = false;
const intakeDrafts = new Map();
const demoDrafts = new Map();
const draftKey = () => `${state.month}/${selectedStaff}`;
function keepIntake() { (demoMode ? demoDrafts : intakeDrafts).set(draftKey(), { draft, candidate }); }
function restoreIntake() { ({ draft, candidate } = (demoMode ? demoDrafts : intakeDrafts).get(draftKey()) || { draft: '', candidate: null }); applied = null; }
function selectPerson(id) { keepIntake(); selectedStaff = id; demoMode = false; restoreIntake(); view = 'intake'; render(); }
function go(nextView) { if (busy) return; keepIntake(); view = nextView; render(); document.querySelector('.app-menu').open = false; }
const coverageSettings = () => state.coverage?.[state.month] || { required: 2, start: '09:00', end: '18:00' };
const coverageRows = () => summarizeCoverage({ month: state.month, staff: state.staff, requests: state.months[state.month] || {}, settings: coverageSettings() });
const content = document.querySelector('#content');
const monthInput = document.querySelector('#target-month');
const modal = document.querySelector('#modal');
monthInput.value = state.month;
if (PUBLIC_DEMO) {
  document.querySelector('#usage-note').textContent = '勤務希望の下書き · このブラウザに保存';
  document.querySelector('#demo-description').textContent = '公開版はAIなしの簡易読み取りです。入力は他の端末と共有されません。確定シフトではありません。';
}
const member = () => state.staff.find(s => s.id === selectedStaff) || state.staff[0];
const requests = id => state.months[state.month]?.[id] || {};
const caption = entry => !entry ? '未登録' : entry.status === 'unavailable' ? '× 勤務不可' : `${entry.start}〜${entry.end}`;
const dow = day => { const [y, m] = state.month.split('-').map(Number); return new Date(y, m - 1, day).getDay(); };
const stamp = () => JSON.stringify({ month: state.month, staff: member(), requests: requests(selectedStaff) });
function notify(text) { document.querySelector('#message').textContent = text; clearTimeout(timer); timer = setTimeout(() => { document.querySelector('#message').textContent = ''; }, 6500); }
function save(next) {
  try { localStorage.setItem(KEY, JSON.stringify(next)); state = next; }
  catch { throw new Error('ブラウザに保存できませんでした。空き容量やブラウザの保存設定を確認してください。'); }
}
function options(values, current) { return values.map(v => `<option${v === current ? ' selected' : ''}>${esc(v)}</option>`).join(''); }
function render() {
  document.querySelectorAll('nav button').forEach(b => b.setAttribute('aria-current', b.dataset.view === (view === 'staff-edit' ? 'staff' : view === 'coverage' ? 'month' : view) ? 'page' : 'false'));
  monthInput.disabled = busy;
  document.querySelectorAll('[data-view], #previous-month, #next-month').forEach(b => b.disabled = busy);
  if (view === 'intake') renderIntake();
  if (view === 'month') renderMonth();
  if (view === 'staff') renderStaff();
  if (view === 'staff-edit') renderStaffEdit();
  if (view === 'history') renderHistory();
  if (view === 'coverage') renderCoverage();
  if (view === 'excel') renderExcel();
}
function renderIntake() {
  const s = member();
  const dmPanel = `<div class="dm-demo${demoMode ? ' is-active' : ''}"><div><div class="dm-demo-title"><span class="draft-badge">デモ</span><strong>Instagram DMの取り込み体験</strong></div><p>${demoMode ? `${esc(s.name)}からのテストDM · ${esc(state.month)}の希望` : '選択中のスタッフから届くサンプルの希望を読み取ります。'}</p><p class="muted">アカウント不要・Instagramとの送受信はありません。保存すると、このブラウザの希望表に反映されます。</p></div><button class="secondary" id="${demoMode ? 'leave-dm-demo' : 'receive-test-dm'}" ${busy ? 'disabled' : ''}>${demoMode ? '通常の入力に戻る' : 'テストDMを受信'}</button></div>`;
  content.innerHTML = `<section class="intake-view"><div class="section-heading"><h2>希望を入力</h2><span class="muted">${Number(state.month.split('-')[1])}月</span></div><div class="form-top"><label>スタッフ<select id="intake-staff" ${busy ? 'disabled' : ''}>${state.staff.map(p => `<option value="${esc(p.id)}"${p.id === s.id ? ' selected' : ''}>${esc(p.name)}${Object.keys(requests(p.id)).length ? ' · 登録あり' : ''}</option>`).join('')}</select></label><span class="muted">基本 ${esc(s.start)}–${esc(s.end)}</span></div>${dmPanel}${applied ? `<div class="intake-success" role="status"><strong>${applied.demo ? 'テストDM · ' : ''}${esc(applied.name)} · ${applied.count}日分を保存しました</strong><div class="actions"><button id="next-person" ${state.staff.length < 2 ? 'hidden' : ''}>次のスタッフへ</button><button class="secondary" id="view-saved">シフト表を見る</button></div></div>` : ''}<label for="request-text">${demoMode ? '受信したテストDM（編集できます）' : '届いたメッセージを貼り付け'}</label><textarea id="request-text" maxlength="1500" placeholder="1,3,4でー&#10;10日は13時から17時" ${busy ? 'disabled' : ''}>${esc(draft)}</textarea><div class="actions intake-actions"><button id="parse" ${busy ? 'disabled' : ''}>${busy ? '読み取り中…' : '内容を確認する'}</button><button class="secondary" id="from-weekdays" ${demoMode ? 'hidden' : ''} ${busy || !s.weekdays.length ? 'disabled' : ''} title="スタッフに登録した基本曜日から候補を作ります">基本曜日から入力</button></div><details class="help-details" ${demoMode ? 'hidden' : ''}><summary>入力例・読み取りについて</summary><div class="example-buttons">${['1,3,4でー', '10,17,25,31でお願いします！！', '3日は10時から16時、4日は15時から', '毎週火曜と木曜に入れます。15日は休みです。'].map((t, i) => `<button class="secondary" data-example="${i}" ${busy ? 'disabled' : ''}>${esc(t)}</button>`).join('')}</div><p>時刻の指定がない日は基本時間を使います。</p><span class="muted" id="ai-status">${PUBLIC_DEMO ? '公開版は対応書式の簡易読み取りです。AIは使いません。' : '接続状況を確認中'}</span></details><div id="preview"></div></section>`;
  if (demoMode) document.querySelector('#leave-dm-demo').onclick = leaveDmDemo;
  else document.querySelector('#receive-test-dm').onclick = receiveTestDm;
  document.querySelector('#intake-staff').onchange = e => selectPerson(e.target.value);
  document.querySelector('#request-text').oninput = e => { draft = e.target.value; candidate = null; applied = null; keepIntake(); document.querySelector('#preview').innerHTML = ''; document.querySelector('.intake-success')?.remove(); };
  document.querySelectorAll('[data-example]').forEach(b => b.onclick = () => { draft = b.textContent; candidate = null; applied = null; keepIntake(); renderIntake(); });
  document.querySelector('#parse').onclick = parse;
  document.querySelector('#from-weekdays').onclick = () => {
    try { const raw = makeWeekdayProposal({ month: state.month, weekdays: s.weekdays, defaultStart: s.start, defaultEnd: s.end }); candidate = { ...normalizeMonthlyProposal(raw, { month: state.month, defaultStart: s.start, defaultEnd: s.end }), snapshot: stamp(), text: `基本曜日: ${s.weekdays.map(d => '日月火水木金土'[d]).join('・')}`, source: 'weekday' }; applied = null; renderPreview(); document.querySelector('#preview').scrollIntoView({ block: 'nearest' }); }
    catch (e) { notify(e.message); }
  };
  if (applied) {
    document.querySelector('#view-saved').onclick = () => go('month');
    document.querySelector('#next-person').onclick = () => selectPerson(state.staff[(state.staff.findIndex(p => p.id === selectedStaff) + 1) % state.staff.length].id);
  }
  if (!PUBLIC_DEMO) fetch('/api/ai/status').then(r => r.json()).then(data => { const el = document.querySelector('#ai-status'); if (el) el.textContent = data.ready ? 'Qwen 3.5 接続済み' : (data.message || data.error || 'Ollamaの起動を確認してください'); }).catch(() => { const el = document.querySelector('#ai-status'); if (el) el.textContent = '接続できません'; });
  renderPreview();
}
async function parse() {
  if (!draft.trim()) { notify('勤務希望の文章を入力してください。'); return; }
  const s = member(), snapshot = stamp(), text = draft, source = demoMode ? 'instagram-demo' : 'message';
  busy = true; candidate = null; applied = null; render();
  try {
    let result;
    if (PUBLIC_DEMO) {
      result = parsePublicRequest(text, { month: state.month, defaultStart: s.start, defaultEnd: s.end });
    } else {
      const response = await fetch('/api/monthly/parse', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ month: state.month, staffName: s.name, prompt: text, defaultStart: s.start, defaultEnd: s.end }) });
      result = await response.json(); if (!response.ok) throw new Error(result.error || '読み取りに失敗しました。');
    }
    if (snapshot !== stamp() || draft !== text) throw new Error('条件が変わりました。もう一度読み取ってください。');
    candidate = { ...result, snapshot, text, source };
  } catch (e) { notify(e.message); }
  finally { busy = false; keepIntake(); render(); document.querySelector('#preview')?.scrollIntoView({ block: 'nearest' }); }
}
function receiveTestDm() {
  if (busy) return;
  keepIntake(); demoMode = true; restoreIntake();
  if (!draft.trim() && !candidate) draft = '3日は10時から16時\n4日は15時から18時\n5日は入れません';
  keepIntake();
  render();
  if (!candidate) parse();
}
function leaveDmDemo() {
  if (busy) return;
  keepIntake(); demoMode = false; restoreIntake(); render();
}
function renderPreview() {
  const el = document.querySelector('#preview'); if (!el || !candidate) return;
  const c = candidate;
  const sourceLabel = c.source === 'instagram-demo' ? '<p class="dm-source">テストDMの読み取り結果 · 保存前に日付と時刻を確認してください。</p>' : '';
  if (c.needsClarification) { el.innerHTML = `<div class="preview"><h2>確認が必要です</h2>${sourceLabel}<p class="feedback">${esc(c.explanation)}</p><p>上の文章を補足して、もう一度読み取ってください。</p></div>`; return; }
  el.innerHTML = `<div class="preview"><div class="section-heading"><h3>読み取り結果 · ${c.entries.length}日分</h3><span class="muted">まだ保存されていません</span></div>${sourceLabel}<div class="preview-cards">${c.entries.map((e, index) => `<div class="preview-card" data-preview-index="${index}"><label class="preview-include"><input type="checkbox" aria-label="${e.day}日を反映" ${e.included === false ? '' : 'checked'}><span>反映</span></label><label class="preview-date"><span class="sr-only">日付</span><input type="number" min="1" max="${daysInMonth(state.month)}" value="${e.day}" aria-label="候補${index + 1}の日付">日</label><select aria-label="${e.day}日の可否"><option value="available"${e.status === 'available' ? ' selected' : ''}>勤務希望</option><option value="unavailable"${e.status === 'unavailable' ? ' selected' : ''}>勤務不可</option></select><div class="preview-time"><input type="time" aria-label="${e.day}日の開始" data-part="start" value="${esc(e.start)}" ${e.status === 'unavailable' ? 'disabled' : ''}><span>–</span><input type="time" aria-label="${e.day}日の終了" data-part="end" value="${esc(e.end)}" ${e.status === 'unavailable' ? 'disabled' : ''}></div><div class="preview-note">${e.startDefaulted || e.endDefaulted ? '<span class="inferred">基本時間</span>' : ''}${requests(selectedStaff)[e.day] ? `<span class="conflict">上書き前：${esc(caption(requests(selectedStaff)[e.day]))}</span>` : ''}</div></div>`).join('')}</div><p class="muted">反映する日だけチェック。日付と時刻は直接直せます。</p><button id="apply">この内容を保存</button><details class="help-details"><summary>読み取りの詳細</summary><p>${esc(c.explanation)}</p></details></div>`;
  el.querySelectorAll('[data-preview-index]').forEach(row => {
    const remember = () => {
      const entry = c.entries[Number(row.dataset.previewIndex)];
      entry.included = row.querySelector('input[type=checkbox]').checked;
      entry.day = row.querySelector('input[type=number]').value;
      entry.status = row.querySelector('select').value;
      for (const part of ['start', 'end']) { const value = row.querySelector(`[data-part=${part}]`).value; if (value !== entry[part]) entry[`${part}Defaulted`] = false; entry[part] = value; }
      keepIntake();
    };
    row.querySelector('select').onchange = e => { const disabled = e.target.value === 'unavailable'; row.querySelectorAll('input[type=time]').forEach(input => { input.disabled = disabled; if (!disabled && !input.value) input.value = member()[input.dataset.part]; }); remember(); };
    row.addEventListener('input', remember);
  });
  document.querySelector('#apply').onclick = () => {
    try {
      if (c.snapshot !== stamp()) throw new Error('条件が変わりました。もう一度読み取ってください。');
      const entries = readPreviewEntries();
      const normalCandidateIsCurrent = demoMode && intakeDrafts.get(draftKey())?.candidate?.snapshot === c.snapshot;
      commitEntries(selectedStaff, entries, c.source, c.text);
      candidate = null; draft = ''; keepIntake();
      if (demoMode) {
        demoMode = false; restoreIntake();
        // Only this demo save changed the baseline. Keep edits and show the latest overwrite warnings.
        if (normalCandidateIsCurrent && candidate) candidate.snapshot = stamp();
      }
      applied = { name: member().name, count: entries.length, demo: c.source === 'instagram-demo' };
      period = Math.floor((entries[0].day - 1) / 7); render(); document.querySelector('.intake-success').scrollIntoView({ block: 'nearest' });
    } catch (e) { notify(e.message); }
  };
}
function readPreviewEntries() {
  const rows = document.querySelectorAll('[data-preview-index]');
  const entries = [...rows].filter(row => row.querySelector('input[type=checkbox]').checked).map(row => {
    const old = candidate.entries[Number(row.dataset.previewIndex)], day = Number(row.querySelector('input[type=number]').value), status = row.querySelector('select').value;
    const start = status === 'available' ? row.querySelector('[data-part=start]').value : '', end = status === 'available' ? row.querySelector('[data-part=end]').value : '';
    if (status === 'available' && (!start || !end)) throw new Error(`${day}日の開始と終了を入力してください。`);
    return { day, status, start, end, startDefaulted: status === 'available' && start === old.start && old.startDefaulted, endDefaulted: status === 'available' && end === old.end && old.endDefaulted };
  });
  if (!entries.length) throw new Error('反映する日を1つ以上選んでください。');
  applyMonthlyEntries({}, entries, { month: state.month });
  return entries;
}
function commitEntries(id, entries, source, text) {
  const next = structuredClone(state); next.months[state.month] ||= {};
  next.months[state.month][id] = applyMonthlyEntries(requests(id), entries, { month: state.month, source, text });
  next.history.unshift({ at: new Date().toISOString(), month: state.month, staffId: id, staffName: state.staff.find(s => s.id === id).name, source, text, entries, previous: Object.fromEntries(entries.map(e => [e.day, requests(id)[e.day] || null])) });
  next.history = next.history.slice(0, 100); save(next);
}
function renderMonth() {
  const count = daysInMonth(state.month), settings = coverageSettings(), rows = coverageRows();
  const registered = state.staff.filter(s => Object.keys(requests(s.id)).length).length;
  const ranges = period === 'all' ? [[1, Math.min(16, count)], [17, count]] : [[period * 7 + 1, Math.min(period * 7 + 7, count)]];
  content.innerHTML = `<section class="schedule-view"><div class="section-heading"><div><h2>${Number(state.month.split('-')[1])}月の勤務希望 <span class="draft-badge">未確定</span></h2><span class="muted">${registered} / ${state.staff.length}人 登録あり</span></div><div class="actions"><button class="secondary" id="open-excel">Excel</button><button id="open-intake">希望を入力</button></div></div>${!registered ? '<div class="schedule-empty">届いた希望を入力するか、<button class="text-button" id="empty-excel">Excelを取り込む</button></div>' : ''}<div class="schedule-controls"><div class="periods" aria-label="表示する日付">${Array.from({ length: Math.ceil(count / 7) }, (_, i) => `<button class="secondary" data-period="${i}" aria-pressed="${period === i}">${i * 7 + 1}–${Math.min(i * 7 + 7, count)}日</button>`).join('')}<button class="secondary" data-period="all" aria-pressed="${period === 'all'}">月全体</button></div><button class="text-button" id="open-coverage">必要${settings.required}人 · 条件変更</button></div>${ranges.map(([from, to]) => `<div class="table-scroll"><table class="month-table ${period === 'all' ? 'full-month' : 'week-table'}"><thead><tr><th>スタッフ</th>${Array.from({ length: to - from + 1 }, (_, i) => { const d = i + from, w = dow(d); return `<th class="${w === 0 ? 'sun' : w === 6 ? 'sat' : ''}"><button class="date-header" data-open-day="${d}" aria-label="${d}日のスタッフと不足を確認">${d}<small>${'日月火水木金土'[w]}</small></button></th>`; }).join('')}</tr></thead><tbody>${state.staff.map(s => `<tr><th scope="row"><button class="staff-shortcut" data-input-person="${esc(s.id)}" aria-label="${esc(s.name)}の希望を入力">${esc(s.name)}</button></th>${Array.from({ length: to - from + 1 }, (_, i) => { const day = from + i, entry = requests(s.id)[day]; return `<td><button class="day-cell ${entry?.status === 'available' ? 'has-request' : ''}" data-id="${esc(s.id)}" data-day="${day}" aria-label="${esc(s.name)} ${day}日 ${esc(caption(entry))}">${!entry ? '<span class="cell-empty">–</span>' : entry.status === 'unavailable' ? '<span class="off">×</span>' : `<span>${esc(entry.start)}</span><span>${esc(entry.end)}</span>`}</button></td>`; }).join('')}</tr>`).join('')}</tbody><tfoot><tr><th scope="row">不足人数</th>${rows.slice(from - 1, to).map(r => { const missing = Math.max(0, ...r.shortages.map(s => s.missing)); const blank = r.unreportedCount === state.staff.length; return `<td><button class="shortage-cell ${blank ? '' : missing ? 'is-shortage' : 'is-covered'}" data-open-day="${r.day}" aria-label="${r.day}日のスタッフと不足を確認">${settings.required === 0 ? '–' : blank ? '–' : missing ? `あと${missing}` : 'OK'}</button></td>`; }).join('')}</tr></tfoot></table></div>`).join('')}<div class="schedule-legend"><span>セルを押して編集 · – 未登録 · × 勤務不可</span><button class="text-button" id="month-coverage">月の不足を見る</button></div></section>`;
  document.querySelector('#open-intake').onclick = () => go('intake');
  document.querySelector('#open-excel').onclick = () => go('excel');
  if (document.querySelector('#empty-excel')) document.querySelector('#empty-excel').onclick = () => go('excel');
  document.querySelector('#open-coverage').onclick = () => { go('coverage'); document.querySelector('#coverage-options').open = true; };
  document.querySelector('#month-coverage').onclick = () => go('coverage');
  document.querySelectorAll('[data-period]').forEach(b => b.onclick = () => { period = b.dataset.period === 'all' ? 'all' : Number(b.dataset.period); renderMonth(); });
  document.querySelectorAll('[data-input-person]').forEach(b => b.onclick = () => selectPerson(b.dataset.inputPerson));
  document.querySelectorAll('.day-cell').forEach(b => b.onclick = () => editCell(b.dataset.id, Number(b.dataset.day)));
  document.querySelectorAll('[data-open-day]').forEach(b => b.onclick = () => openCoverageDay(rows.find(r => r.day === Number(b.dataset.openDay)), settings));
}
function editCell(id, day, returnToDay = false) {
  if (modal.open) modal.close();
  const s = state.staff.find(s => s.id === id), entry = requests(id)[day];
  document.querySelector('#dialog-content').innerHTML = `<h2 id="dialog-title">${esc(s.name)} ／ ${day}日</h2><form id="cell-form"><div class="small-form"><label>可否<select id="cell-status"><option value="available"${entry?.status !== 'unavailable' ? ' selected' : ''}>勤務希望</option><option value="unavailable"${entry?.status === 'unavailable' ? ' selected' : ''}>勤務不可</option></select></label><label>開始<input id="cell-start" type="time" value="${esc(entry?.start || s.start)}" required></label><label>終了<input id="cell-end" type="time" value="${esc(entry?.end || s.end)}" required></label></div><div class="actions"><button type="button" class="secondary" id="close-dialog">キャンセル</button><button type="button" class="danger" id="clear-cell">未登録に戻す</button><button>保存</button></div></form>`;
  const toggle = () => document.querySelectorAll('#cell-form input').forEach(i => i.disabled = document.querySelector('#cell-status').value === 'unavailable'); toggle();
  document.querySelector('#cell-status').onchange = toggle;
  document.querySelector('#close-dialog').onclick = () => { modal.close(); if (returnToDay) openCoverageDay(coverageRows().find(r => r.day === day), coverageSettings()); };
  document.querySelector('#cell-form').onsubmit = e => { e.preventDefault(); try { const status = document.querySelector('#cell-status').value; commitEntries(id, [{ day, status, start: status === 'available' ? document.querySelector('#cell-start').value : '', end: status === 'available' ? document.querySelector('#cell-end').value : '' }], 'manual', `${day}日を画面で手修正`); modal.close(); render(); if (returnToDay) openCoverageDay(coverageRows().find(r => r.day === day), coverageSettings()); notify('保存しました。'); } catch (error) { notify(error.message); } };
  document.querySelector('#clear-cell').onclick = () => { try { const next = structuredClone(state); delete next.months[state.month]?.[id]?.[day]; next.history.unshift({ at: new Date().toISOString(), month: state.month, staffId: id, staffName: s.name, source: 'manual', text: `${day}日を未登録に戻しました`, entries: [], previous: { [day]: entry || null } }); next.history = next.history.slice(0, 100); save(next); modal.close(); render(); if (returnToDay) openCoverageDay(coverageRows().find(r => r.day === day), coverageSettings()); } catch (e) { notify(e.message); } };
  modal.showModal();
}
function renderStaff() {
  content.innerHTML = `<section><div class="section-heading"><h2>スタッフ <span class="muted">${state.staff.length}人</span></h2><button id="open-staff-edit">${staffDraft ? '編集中の内容へ' : 'スタッフを編集'}</button></div><div class="staff-cards">${state.staff.map(s => `<article class="staff-card"><h3>${esc(s.name)}</h3><p>${esc(s.start)}–${esc(s.end)}</p><p class="muted">${[1, 2, 3, 4, 5, 6, 0].filter(d => s.weekdays.includes(d)).map(d => '日月火水木金土'[d]).join('・') || '基本曜日なし'}</p><div class="staff-tags">${['experience', 'category', 'preference'].filter(key => s[key] !== '未設定').map(key => `<span>${esc(s[key])}</span>`).join('')}</div><button class="secondary" data-staff-input="${esc(s.id)}">希望を入力</button></article>`).join('')}</div><details class="help-details"><summary>基本条件について</summary><p>基本曜日・時間は希望入力に使います。経験・区分・勤務量は管理用で、自動割り当てには使いません。</p></details></section>`;
  document.querySelector('#open-staff-edit').onclick = () => go('staff-edit');
  document.querySelectorAll('[data-staff-input]').forEach(b => b.onclick = () => selectPerson(b.dataset.staffInput));
}
function renderStaffEdit() {
  staffDraft ||= structuredClone(state.staff);
  content.innerHTML = `<section><h2>スタッフを編集</h2><form id="staff-form"><div class="staff-edit-list">${staffDraft.map(s => `<fieldset class="staff-edit-card" data-staff-id="${esc(s.id)}"><legend>${esc(s.name)}</legend><label>名前<input aria-label="${esc(s.name)}の名前" name="name" type="text" maxlength="40" required value="${esc(s.name)}"></label><div class="staff-edit-attributes">${['experience', 'category', 'preference'].map(key => `<label>${{ experience: '経験', category: '区分', preference: '勤務量' }[key]}<select aria-label="${esc(s.name)}の${{ experience: '経験', category: '区分', preference: '勤務量' }[key]}" name="${key}">${options(labels[key], s[key])}</select></label>`).join('')}</div><div class="weekday-checks">${[1, 2, 3, 4, 5, 6, 0].map(d => `<label>${'日月火水木金土'[d]}<input type="checkbox" value="${d}" aria-label="${esc(s.name)} ${'日月火水木金土'[d]}曜日" ${s.weekdays.includes(d) ? 'checked' : ''}></label>`).join('')}</div><div class="staff-edit-time"><label>基本開始<input type="time" name="start" aria-label="${esc(s.name)}の基本開始" value="${esc(s.start)}" required></label><label>基本終了<input type="time" name="end" aria-label="${esc(s.name)}の基本終了" value="${esc(s.end)}" required></label></div></fieldset>`).join('')}</div><button class="secondary" type="button" id="add-staff">スタッフを追加</button><div class="actions sticky-actions"><span class="muted">保存するまで未反映</span><button class="secondary" type="button" id="cancel-staff">キャンセル</button><button>保存</button></div></form></section>`;
  const readDraft = () => [...document.querySelectorAll('[data-staff-id]')].map(row => {
    const s = { ...staffDraft.find(person => person.id === row.dataset.staffId) };
    ['name', 'experience', 'category', 'preference', 'start', 'end'].forEach(k => s[k] = row.querySelector(`[name=${k}]`).value.trim());
    s.weekdays = [...row.querySelectorAll('input:checked')].map(i => Number(i.value)); return s;
  });
  const form = document.querySelector('#staff-form');
  form.addEventListener('input', () => { staffDraft = readDraft(); });
  form.addEventListener('change', () => { staffDraft = readDraft(); });
  form.onsubmit = e => {
    e.preventDefault();
    try {
      const staff = readDraft();
      for (const s of staff) {
        if (!s.name) throw new Error('スタッフ名を入力してください。');
        if (!s.start || !s.end) throw new Error('基本開始と基本終了を入力してください。');
        normalizeMonthlyProposal({ explanation: '', needsClarification: false, entries: [{ day: 1, status: 'available', start: s.start, end: s.end }] }, { month: state.month });
      }
      if (new Set(staff.map(s => s.name)).size !== staff.length) throw new Error('同じ名前のスタッフがいます。区別できる名前にしてください。');
      save({ ...state, staff }); staffDraft = null; view = 'staff'; render(); notify('基本条件を保存しました。');
    } catch (error) { notify(error.message); }
  };
  document.querySelector('#add-staff').onclick = () => {
    if (staffDraft.length >= 30) { notify('デモでは30名まで登録できます。'); return; }
    staffDraft = readDraft(); let number = staffDraft.length + 1;
    while (staffDraft.some(s => s.name === `スタッフ${number}`)) number++;
    staffDraft.push(createStaff(`スタッフ${number}`)); render();
  };
  document.querySelector('#cancel-staff').onclick = () => { staffDraft = null; view = 'staff'; render(); notify('入力中の変更を取り消しました。'); };
}
function renderHistory() {
  const entries = state.history.filter(h => h.month === state.month);
  content.innerHTML = `<section><h2>履歴</h2><p class="muted">${esc(state.month)} · ${entries.length}件</p>${entries.map(h => `<details class="history-entry"><summary><strong>${esc(h.staffName)}</strong><span>${esc(new Date(h.at).toLocaleString('ja-JP'))}</span><span>${h.entries.length ? `${h.entries.length}日分` : '未登録に変更'}</span>${h.source === 'instagram-demo' ? '<span class="draft-badge">テストDM</span>' : ''}</summary><p class="history-message">${esc(h.text)}</p><p>${h.entries.map(e => `${e.day}日 ${caption(e)}${e.startDefaulted || e.endDefaulted ? '（基本時間）' : ''}`).map(esc).join(' ／ ') || 'セルを未登録に変更'}</p></details>`).join('') || '<p>まだ履歴はありません。</p>'}<details class="help-details"><summary>保存する内容</summary><p>反映した元文章と変更内容を、全月合わせて最新100件まで保存します。</p></details></section>`;
}
function renderCoverage() {
  const settings = coverageSettings();
  const rows = summarizeCoverage({ month: state.month, staff: state.staff, requests: state.months[state.month] || {}, settings });
  const missingStaff = state.staff.filter(s => Object.keys(requests(s.id)).length === 0);
  const shortageDays = rows.filter(row => row.shortages.length);
  const configured = settings.required > 0;
  const leadingDays = dow(1);
  const monthNumber = Number(state.month.split('-')[1]);
  content.innerHTML = `<section class="coverage-view"><div class="coverage-heading"><h2>人数チェック</h2><button class="text-button" id="back-schedule">シフト表に戻る</button><p class="muted">${esc(state.month)} · 勤務希望からの目安です</p></div><div class="coverage-summary"><div class="coverage-stat${configured ? shortageDays.length ? ' is-warning' : ' is-good' : ''}"><span>不足がある日</span><strong>${configured ? `${shortageDays.length}<small>日</small>` : '—'}</strong></div><div class="coverage-stat${missingStaff.length ? ' is-warning' : ' is-good'}"><span>希望未登録</span><strong>${missingStaff.length}<small>人</small></strong></div><div class="coverage-stat"><span>必要人数</span><strong>${configured ? `${settings.required}<small>人</small>` : '未設定'}</strong></div></div><details id="coverage-options" class="coverage-options"><summary>必要${settings.required}人 · ${esc(settings.start)}–${esc(settings.end)} <span>条件を変更</span></summary><form id="coverage-form" class="coverage-settings"><label>必要人数<input name="required" type="number" min="0" max="30" value="${settings.required}" required></label><label>開始<input name="start" type="time" value="${esc(settings.start)}" required></label><span class="coverage-time-separator" aria-hidden="true">〜</span><label>終了<input name="end" type="time" value="${esc(settings.end)}" required></label><button>保存</button><span class="coverage-settings-note muted">全日に適用 · 0人で解除</span></form></details>${missingStaff.length ? `<details class="coverage-unsubmitted"><summary>希望未登録のスタッフ <span>${missingStaff.length}人</span></summary><p class="muted">今月の希望がまだ1件もありません。</p><div class="coverage-names">${missingStaff.map(s => `<button class="secondary" data-missing-person="${esc(s.id)}">${esc(s.name)} · 入力</button>`).join('')}</div></details>` : ''}${!configured ? '<div class="coverage-empty"><h3>何人必要ですか？</h3><p>上で人数を設定すると、不足する日がわかります。</p></div>' : `<div class="coverage-results"><div class="coverage-toolbar"><h3>日ごとの状況</h3><span class="muted">日付を押すと詳細</span></div><div class="coverage-calendar" aria-label="${esc(state.month)}の人数"><div class="coverage-weekday is-sunday">日</div>${['月', '火', '水', '木', '金'].map(day => `<div class="coverage-weekday">${day}</div>`).join('')}<div class="coverage-weekday is-saturday">土</div>${'<span class="coverage-empty-cell" aria-hidden="true"></span>'.repeat(leadingDays)}${rows.map(r => `<button type="button" class="coverage-day-button ${r.shortages.length ? 'is-shortage' : 'is-covered'}" data-coverage-day="${r.day}" aria-haspopup="dialog" aria-label="${monthNumber}月${r.day}日（${'日月火水木金土'[dow(r.day)]}） ${r.shortages.length ? `最大${Math.max(...r.shortages.map(s => s.missing))}人不足` : '充足'}"><strong>${r.day}</strong><span>${r.shortages.length ? `あと${Math.max(...r.shortages.map(s => s.missing))}人` : '充足'}</span></button>`).join('')}</div></div>`}<details class="help-details"><summary>集計について</summary><p>確定シフトではなく、登録された勤務希望を集計しています。未登録は勤務不可とは区別します。希望あり・ベテランの人数は当日全体、最少人数は設定した時間内に同時に入れる人数です。</p></details></section>`;
  document.querySelector('#back-schedule').onclick = () => go('month');
  document.querySelectorAll('[data-missing-person]').forEach(b => b.onclick = () => selectPerson(b.dataset.missingPerson));
  document.querySelectorAll('[data-coverage-day]').forEach(button => button.onclick = () => openCoverageDay(rows.find(row => row.day === Number(button.dataset.coverageDay)), settings));
  document.querySelector('#coverage-form').onsubmit = e => {
    e.preventDefault();
    try {
      const form = new FormData(e.currentTarget), nextSettings = { required: Number(form.get('required')), start: form.get('start'), end: form.get('end') };
      summarizeCoverage({ month: state.month, staff: state.staff, requests: state.months[state.month] || {}, settings: nextSettings });
      save({ ...state, coverage: { ...state.coverage, [state.month]: nextSettings } }); render(); notify('人数の設定を保存しました。');
    } catch (error) { notify(error.message); }
  };
}
function openCoverageDay(row, settings) {
  const monthNumber = Number(state.month.split('-')[1]);
  document.querySelector('#dialog-content').innerHTML = `<h2 id="dialog-title">${monthNumber}月${row.day}日（${'日月火水木金土'[dow(row.day)]}）</h2>${settings.required === 0 ? '<p class="muted">人数チェックは解除中</p>' : row.shortages.length ? `<div class="coverage-shortages">${row.shortages.map(s => `<span class="coverage-interval"><span>${s.start}–${s.end}</span><strong>あと${s.missing}人</strong></span>`).join('')}</div>` : `<p class="good">${esc(settings.start)}–${esc(settings.end)} · 人数を満たしています</p>`}<div class="day-people">${state.staff.map(s => `<div><span><strong>${esc(s.name)}</strong><small>${esc(caption(requests(s.id)[row.day]))}</small></span><button class="secondary" data-day-edit="${esc(s.id)}" aria-label="${esc(s.name)} ${row.day}日を編集">編集</button></div>`).join('')}</div><details class="help-details"><summary>人数の内訳</summary><dl class="coverage-breakdown"><div><dt>希望あり</dt><dd>${row.availableCount}人</dd></div><div><dt>うちベテラン</dt><dd>${row.veteranCount}人</dd></div><div><dt>未登録</dt><dd>${row.unreportedCount}人</dd></div><div><dt>同時に入れる最少人数</dt><dd>${row.minimumAvailable}人</dd></div></dl></details><div class="actions"><button type="button" id="close-dialog" autofocus>閉じる</button></div>`;
  document.querySelector('#close-dialog').onclick = () => modal.close();
  document.querySelectorAll('[data-day-edit]').forEach(b => b.onclick = () => editCell(b.dataset.dayEdit, row.day, true));
  modal.showModal();
}

async function downloadExcel(options = {}) {
  busy = true; render();
  try {
    const buffer = await exportExcel(state, options);
    const url = URL.createObjectURL(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const link = document.createElement('a'); link.href = url;
    const suffix = options.staffId ? `_${state.staff.find(s => s.id === options.staffId)?.name || ''}` : '';
    link.download = `ポルテ_${state.month}${suffix}_${options.template ? '提出用' : '勤務希望表'}.xlsx`;
    document.body.append(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 30000);
    notify('Excelを出力しました。');
  } catch (error) { notify(error.message); }
  finally { busy = false; render(); }
}
function parseExcelSheet() {
  const session = excelSession;
  session.parsed = null; session.error = ''; session.plan = null;
  try {
    session.parsed = parseWorksheetRows(worksheetRows(session.workbook.getWorksheet(session.sheet)), { fallbackMonth: state.month });
    session.mapping = session.parsed.staff.map(person => ({ sourceName: person.name, targetId: state.staff.find(s => s.name === person.name)?.id || 'new' }));
  } catch (error) { session.error = error.message; }
}
function renderExcel() {
  const session = excelSession;
  content.innerHTML = `<section><h2>Excel</h2><p class="muted">ファイルはこのブラウザ内で処理します。</p><h3>書き出す</h3><div class="small-form"><label>スタッフ<select id="export-person"><option value="">全員</option>${state.staff.map(s => `<option value="${esc(s.id)}">${esc(s.name)}</option>`).join('')}</select></label><button id="export-full" ${busy ? 'disabled' : ''}>希望表を出力</button><button class="secondary" id="export-template" ${busy ? 'disabled' : ''}>空の提出用を出力</button></div><details class="help-details"><summary>提出用Excelの使い方</summary><p>スタッフに渡し、記入済みのファイルを下で取り込めます。「勤務希望表」「取込用」「基本条件」「取り込み履歴」の4シートを書き出します。</p></details><h3>取り込む</h3><p class="muted">.xlsx · 5MBまで · 内容を確認してから反映</p><details class="help-details"><summary>対応する形式</summary><p>「対象月・名前・日・可否・開始・終了」の一覧、または前半・後半に分かれた開始／終了2段の表に対応します。1シート1000行・40列まで。空欄は登録済みの希望を消しません。</p></details><label class="excel-file">Excelファイルを選択<input id="excel-file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ${busy ? 'disabled' : ''}></label>${busy ? '<p role="status">Excelを処理中です…</p>' : ''}${session?.workbook ? `<p>選択中：${esc(session.filename)}</p><label>取り込むシート<select id="excel-sheet">${session.workbook.worksheets.map(s => `<option${s.name === session.sheet ? ' selected' : ''}>${esc(s.name)}</option>`).join('')}</select></label><p class="muted">Excelで編集したシートを選んでください。履歴は取り込みません。</p>` : ''}<div id="excel-review"></div></section>`;
  document.querySelector('#export-full').onclick = () => downloadExcel({ staffId: document.querySelector('#export-person').value });
  document.querySelector('#export-template').onclick = () => downloadExcel({ template: true, staffId: document.querySelector('#export-person').value });
  document.querySelector('#excel-file').onchange = async e => {
    const file = e.target.files[0]; if (!file) return;
    excelSession = null; busy = true; render();
    try {
      const workbook = await readExcelFile(file);
      excelSession = { workbook, filename: file.name, sheet: workbook.getWorksheet('取込用')?.name || workbook.worksheets[0].name, conditions: [], conditionsError: '', importConditions: false };
      try { excelSession.conditions = readStaffConditions(workbook); } catch (error) { excelSession.conditionsError = error.message; }
      parseExcelSheet();
    } catch (error) { excelSession = { error: error.message }; }
    finally { busy = false; render(); }
  };
  const selector = document.querySelector('#excel-sheet');
  if (selector) selector.onchange = e => { session.sheet = e.target.value; parseExcelSheet(); renderExcelReview(); };
  renderExcelReview();
}
function renderExcelReview() {
  const root = document.querySelector('#excel-review'), session = excelSession; if (!root || !session) return;
  if (session.error) { root.innerHTML = `<p class="import-errors" role="alert">${esc(session.error)}</p><p>希望表は変更していません。シートを選び直すか、Excelの内容を修正してください。</p>`; return; }
  if (!session.parsed) return;
  const parsed = session.parsed;
  root.innerHTML = `<div class="preview"><h3>取り込み先：${esc(parsed.month)}</h3><p class="feedback">${parsed.month !== state.month ? `画面の対象月（${esc(state.month)}）とは異なります。反映後は${esc(parsed.month)}へ切り替えます。` : '対象月を確認してください。'}</p>${parsed.warnings.length ? `<ul>${parsed.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul>` : ''}<h3>スタッフを確認</h3><div class="table-scroll"><table><thead><tr><th>Excelの名前</th><th>取り込み先</th></tr></thead><tbody>${parsed.staff.map((s, i) => `<tr><th>${esc(s.name)}</th><td><select data-import-person="${i}" aria-label="${esc(s.name)}の取り込み先"><option value="new">新しいスタッフ</option><option value="skip">取り込まない</option>${state.staff.map(p => `<option value="${esc(p.id)}">${esc(p.name)}</option>`).join('')}</select></td></tr>`).join('')}</tbody></table></div>${session.conditions.length ? `<label><input type="checkbox" id="excel-conditions" ${session.importConditions ? 'checked' : ''}>基本条件も更新（${session.conditions.length}人）</label>` : ''}${session.conditionsError ? `<p class="import-errors">基本条件は取り込めません：${esc(session.conditionsError)}</p>` : ''}<div id="excel-plan"></div></div>`;
  root.querySelectorAll('[data-import-person]').forEach(select => {
    const index = Number(select.dataset.importPerson); select.value = session.mapping[index].targetId;
    select.onchange = () => { session.mapping[index].targetId = select.value; renderExcelPlan(); };
  });
  const conditions = document.querySelector('#excel-conditions');
  if (conditions) conditions.onchange = () => { session.importConditions = conditions.checked; renderExcelPlan(); };
  renderExcelPlan();
}
function renderExcelPlan() {
  const root = document.querySelector('#excel-plan'), session = excelSession; session.plan = null;
  try {
    const plan = buildExcelImportPlan(state, session.parsed, { mapping: session.mapping, conditions: session.conditions, importConditions: session.importConditions, filename: session.filename });
    session.plan = plan; session.snapshot = JSON.stringify(state);
    root.innerHTML = `<h3>変更内容</h3><p>${plan.peopleCount}人・${plan.entryCount}日分 · 新規${plan.newStaffCount}人 · <strong>上書き${plan.overwrites}件</strong></p><p class="muted">ファイルにない日はそのまま残します。</p>${staffDraft ? '<p class="import-errors">スタッフ編集中です。先に保存するか、キャンセルしてください。</p>' : ''}<div class="excel-preview"><table><thead><tr><th>スタッフ</th><th>日</th><th>現在</th><th>反映後</th><th>確認</th></tr></thead><tbody>${plan.changes.map(c => `<tr><th>${esc(c.staffName)}</th><td>${c.day}日</td><td>${esc(caption(c.previous))}</td><td>${esc(caption(c.next))}</td><td>${c.next.startDefaulted || c.next.endDefaulted ? '基本時間で補完' : ''}${c.previous ? ' ／ 上書き' : ''}</td></tr>`).join('')}</tbody></table></div><div class="actions"><label><input type="checkbox" id="excel-confirm">対象月・名前・時刻・上書き内容を確認しました</label><button id="excel-apply" disabled>反映する</button></div>`;
    document.querySelector('#excel-confirm').onchange = e => { document.querySelector('#excel-apply').disabled = !e.target.checked || !!staffDraft; };
    document.querySelector('#excel-apply').onclick = () => {
      try {
        if (!document.querySelector('#excel-confirm').checked || staffDraft) throw new Error('確認と、入力中の基本条件の保存・取り消しを済ませてください。');
        if (session.snapshot !== JSON.stringify(state)) throw new Error('希望表が変わりました。取り込み内容を確認し直してください。');
        keepIntake(); save(plan.nextState); monthInput.value = state.month; excelSession = null;
        if (!state.staff.some(s => s.id === selectedStaff)) selectedStaff = state.staff[0].id;
        demoMode = false; restoreIntake(); period = 0; view = 'month'; render(); notify(`${plan.peopleCount}名・${plan.entryCount}日分をExcelから反映しました。`);
      } catch (error) { notify(error.message); }
    };
  } catch (error) { root.innerHTML = `<p class="import-errors" role="alert">${esc(error.message)}</p><p>希望表は変更していません。名前の対応やExcelの内容を確認してください。</p>`; }
}
function changeMonth(value) {
  try { daysInMonth(value); keepIntake(); save({ ...state, month: value }); monthInput.value = state.month; demoMode = false; restoreIntake(); period = 0; render(); }
  catch (e) { monthInput.value = state.month; notify(e.message); }
}
monthInput.onchange = () => changeMonth(monthInput.value);
for (const [id, delta] of [['previous-month', -1], ['next-month', 1]]) document.querySelector(`#${id}`).onclick = () => {
  const [year, month] = state.month.split('-').map(Number), date = new Date(year, month - 1 + delta, 1);
  changeMonth(`${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`);
};
document.querySelectorAll('[data-view]').forEach(b => b.onclick = () => go(b.dataset.view));
render();
