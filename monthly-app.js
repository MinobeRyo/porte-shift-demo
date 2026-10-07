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
let view = 'intake', selectedStaff = state.staff[0].id, draft = '', candidate = null, busy = false, timer, staffDraft = null;
let excelSession = null;
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
  document.querySelectorAll('nav button').forEach(b => b.setAttribute('aria-current', b.dataset.view === view ? 'page' : 'false'));
  monthInput.disabled = busy;
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
  content.innerHTML = `<section><h2>希望を入力</h2><div class="form-top"><label>スタッフ<select id="intake-staff" ${busy ? 'disabled' : ''}>${state.staff.map(p => `<option value="${esc(p.id)}"${p.id === s.id ? ' selected' : ''}>${esc(p.name)}</option>`).join('')}</select></label><span class="muted">時間指定なし：${esc(s.start)}〜${esc(s.end)}</span></div><details class="help-details"><summary>入力例を見る</summary><div class="example-buttons">${['1,3,4でー', '10,17,25,31でお願いします！！', '3日は10時から16時、4日は15時から', '毎週火曜と木曜に入れます。15日は休みです。'].map((t, i) => `<button class="secondary" data-example="${i}" ${busy ? 'disabled' : ''}>${esc(t)}</button>`).join('')}</div><p>1人分ずつ入力してください。名前はスタッフ編集で変更できます。</p></details><label for="request-text">メッセージ</label><textarea id="request-text" maxlength="1500" placeholder="例：1,3,4でー" ${busy ? 'disabled' : ''}>${esc(draft)}</textarea><div class="actions"><button id="parse" ${busy ? 'disabled' : ''}>${busy ? '読み取り中…' : '読み取る'}</button><button class="secondary" id="from-weekdays" ${busy ? 'disabled' : ''}>基本曜日を使う</button><span class="muted" id="ai-status">${PUBLIC_DEMO ? '簡易読み取り · AIなし' : '接続状況を確認中'}</span></div><div id="preview"></div></section>`;
  document.querySelector('#intake-staff').onchange = e => { selectedStaff = e.target.value; draft = ''; candidate = null; renderIntake(); };
  document.querySelector('#request-text').oninput = e => { draft = e.target.value; candidate = null; document.querySelector('#preview').innerHTML = ''; };
  document.querySelectorAll('[data-example]').forEach(b => b.onclick = () => { draft = b.textContent; candidate = null; renderIntake(); });
  document.querySelector('#parse').onclick = parse;
  document.querySelector('#from-weekdays').onclick = () => {
    try { const raw = makeWeekdayProposal({ month: state.month, weekdays: s.weekdays, defaultStart: s.start, defaultEnd: s.end }); candidate = { ...normalizeMonthlyProposal(raw, { month: state.month, defaultStart: s.start, defaultEnd: s.end }), snapshot: stamp(), text: `基本曜日: ${s.weekdays.map(d => '日月火水木金土'[d]).join('・')}`, source: 'weekday' }; renderPreview(); }
    catch (e) { notify(e.message); }
  };
  if (!PUBLIC_DEMO) fetch('/api/ai/status').then(r => r.json()).then(data => { const el = document.querySelector('#ai-status'); if (el) el.textContent = data.ready ? 'Qwen 3.5 接続済み' : (data.message || data.error || 'Ollamaの起動を確認してください'); }).catch(() => { const el = document.querySelector('#ai-status'); if (el) el.textContent = '接続できません'; });
  renderPreview();
}
async function parse() {
  if (!draft.trim()) { notify('勤務希望の文章を入力してください。'); return; }
  const s = member(), snapshot = stamp(), text = draft;
  busy = true; candidate = null; render();
  try {
    let result;
    if (PUBLIC_DEMO) {
      result = parsePublicRequest(text, { month: state.month, defaultStart: s.start, defaultEnd: s.end });
    } else {
      const response = await fetch('/api/monthly/parse', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ month: state.month, staffName: s.name, prompt: text, defaultStart: s.start, defaultEnd: s.end }) });
      result = await response.json(); if (!response.ok) throw new Error(result.error || '読み取りに失敗しました。');
    }
    if (snapshot !== stamp() || draft !== text) throw new Error('条件が変わりました。もう一度読み取ってください。');
    candidate = { ...result, snapshot, text, source: 'message' };
  } catch (e) { notify(e.message); }
  finally { busy = false; render(); }
}
function renderPreview() {
  const el = document.querySelector('#preview'); if (!el || !candidate) return;
  const c = candidate;
  if (c.needsClarification) { el.innerHTML = `<div class="preview"><h2>確認が必要です</h2><p class="feedback">${esc(c.explanation)}</p><p>上の文章を補足して、もう一度読み取ってください。</p></div>`; return; }
  el.innerHTML = `<div class="preview"><h2>${c.entries.length}日分の確認</h2><p class="muted">登録済みの日は上書きします。反映しない日はチェックを外してください。</p><details class="help-details"><summary>読み取りの詳細</summary><p>${esc(c.explanation)}</p><p>日付・可否・時刻は、この表で修正できます。</p></details><div class="table-scroll"><table><thead><tr><th>反映</th><th>日付</th><th>可否</th><th>開始</th><th>終了</th><th>補完・既存の内容</th></tr></thead><tbody>${c.entries.map((e, index) => `<tr data-preview-index="${index}"><td><input type="checkbox" aria-label="${e.day}日を反映" checked></td><td><input type="number" min="1" max="${daysInMonth(state.month)}" value="${e.day}" aria-label="候補${index + 1}の日付">日</td><td><select aria-label="${e.day}日の可否"><option value="available"${e.status === 'available' ? ' selected' : ''}>勤務希望</option><option value="unavailable"${e.status === 'unavailable' ? ' selected' : ''}>勤務不可</option></select></td><td><input type="time" aria-label="${e.day}日の開始" data-part="start" value="${esc(e.start)}" ${e.status === 'unavailable' ? 'disabled' : ''}></td><td><input type="time" aria-label="${e.day}日の終了" data-part="end" value="${esc(e.end)}" ${e.status === 'unavailable' ? 'disabled' : ''}></td><td>${e.startDefaulted || e.endDefaulted ? '<span class="inferred">基本時間で補完</span><br>' : ''}${requests(selectedStaff)[e.day] ? `<span class="conflict">現在：${esc(caption(requests(selectedStaff)[e.day]))}</span>` : '新規'}</td></tr>`).join('')}</tbody></table></div><button id="apply">希望表に反映</button></div>`;
  el.querySelectorAll('[data-preview-index]').forEach(row => row.querySelector('select').onchange = e => { const disabled = e.target.value === 'unavailable'; row.querySelectorAll('input[type=time]').forEach(input => { input.disabled = disabled; if (!disabled && !input.value) input.value = member()[input.dataset.part]; }); });
  document.querySelector('#apply').onclick = () => {
    try {
      if (c.snapshot !== stamp()) throw new Error('条件が変わりました。もう一度読み取ってください。');
      const entries = readPreviewEntries();
      commitEntries(selectedStaff, entries, c.source, c.text); candidate = null; view = 'month'; render(); notify(`${entries.length}日分を未確定の希望表に反映しました。`);
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
  const count = daysInMonth(state.month);
  content.innerHTML = `<section><div class="section-heading"><h2>${esc(state.month)} 希望表</h2><button id="export" ${busy ? 'disabled' : ''}>Excel出力</button></div><p class="muted">セルを押して編集 · 空欄は未登録 · ×は勤務不可</p><details class="help-details"><summary>表の見方</summary><p>未確定の勤務希望です。上段が開始、下段が終了。横スクロールで末日まで確認できます。</p></details>${[[1, Math.min(16, count)], [17, count]].map(([from, to]) => `<h3>${from}日〜${to}日</h3><div class="table-scroll"><table class="month-table"><thead><tr><th>名前</th>${Array.from({ length: to - from + 1 }, (_, i) => { const d = i + from, w = dow(d); return `<th class="${w === 0 ? 'sun' : w === 6 ? 'sat' : ''}">${d}<small>${'日月火水木金土'[w]}</small></th>`; }).join('')}</tr></thead><tbody>${state.staff.map(s => `<tr><th scope="row">${esc(s.name)}</th>${Array.from({ length: to - from + 1 }, (_, i) => { const day = from + i, entry = requests(s.id)[day]; return `<td><button class="day-cell" data-id="${esc(s.id)}" data-day="${day}" aria-label="${esc(s.name)} ${day}日 ${esc(caption(entry))}">${!entry ? '&nbsp;' : entry.status === 'unavailable' ? '<span class="off">×</span>' : `<span>${esc(entry.start)}</span><span>${esc(entry.end)}</span>`}</button></td>`; }).join('')}</tr>`).join('')}</tbody></table></div>`).join('')}</section>`;
  document.querySelector('#export').onclick = () => downloadExcel();
  document.querySelectorAll('.day-cell').forEach(b => b.onclick = () => editCell(b.dataset.id, Number(b.dataset.day)));
}
function editCell(id, day) {
  const s = state.staff.find(s => s.id === id), entry = requests(id)[day];
  document.querySelector('#dialog-content').innerHTML = `<h2 id="dialog-title">${esc(s.name)} ／ ${day}日</h2><form id="cell-form"><div class="small-form"><label>可否<select id="cell-status"><option value="available"${entry?.status !== 'unavailable' ? ' selected' : ''}>勤務希望</option><option value="unavailable"${entry?.status === 'unavailable' ? ' selected' : ''}>勤務不可</option></select></label><label>開始<input id="cell-start" type="time" value="${esc(entry?.start || s.start)}" required></label><label>終了<input id="cell-end" type="time" value="${esc(entry?.end || s.end)}" required></label></div><div class="actions"><button type="button" class="secondary" id="close-dialog">キャンセル</button><button type="button" class="danger" id="clear-cell">未登録に戻す</button><button>保存</button></div></form>`;
  const toggle = () => document.querySelectorAll('#cell-form input').forEach(i => i.disabled = document.querySelector('#cell-status').value === 'unavailable'); toggle();
  document.querySelector('#cell-status').onchange = toggle;
  document.querySelector('#close-dialog').onclick = () => modal.close();
  document.querySelector('#cell-form').onsubmit = e => { e.preventDefault(); try { const status = document.querySelector('#cell-status').value; commitEntries(id, [{ day, status, start: status === 'available' ? document.querySelector('#cell-start').value : '', end: status === 'available' ? document.querySelector('#cell-end').value : '' }], 'manual', `${day}日を画面で手修正`); candidate = null; modal.close(); render(); notify('保存しました。'); } catch (error) { notify(error.message); } };
  document.querySelector('#clear-cell').onclick = () => { try { const next = structuredClone(state); delete next.months[state.month]?.[id]?.[day]; next.history.unshift({ at: new Date().toISOString(), month: state.month, staffId: id, staffName: s.name, source: 'manual', text: `${day}日を未登録に戻しました`, entries: [], previous: { [day]: entry || null } }); next.history = next.history.slice(0, 100); save(next); candidate = null; modal.close(); render(); } catch (e) { notify(e.message); } };
  modal.showModal();
}
function renderStaff() {
  content.innerHTML = `<section><div class="section-heading"><h2>スタッフ</h2><button id="open-staff-edit">編集</button></div><p class="muted">${state.staff.length}人</p><div class="table-scroll"><table class="staff-read-table"><thead><tr><th>名前</th><th>経験</th><th>区分</th><th>勤務量</th><th>基本曜日</th><th>基本時間</th></tr></thead><tbody>${state.staff.map(s => `<tr><th scope="row">${esc(s.name)}</th><td>${esc(s.experience)}</td><td>${esc(s.category)}</td><td>${esc(s.preference)}</td><td>${[1, 2, 3, 4, 5, 6, 0].filter(d => s.weekdays.includes(d)).map(d => '日月火水木金土'[d]).join('・') || '未設定'}</td><td>${esc(s.start)}〜${esc(s.end)}</td></tr>`).join('')}</tbody></table></div><details class="help-details"><summary>基本条件の使い道</summary><p>曜日・時間は希望の候補作成に使います。経験・区分・勤務量による自動割り当てには、まだ対応していません。</p></details></section>`;
  document.querySelector('#open-staff-edit').onclick = () => { view = 'staff-edit'; render(); };
}
function renderStaffEdit() {
  staffDraft ||= structuredClone(state.staff);
  content.innerHTML = `<section><h2>スタッフを編集</h2><p class="muted">編集が終わったら保存してください。</p><form id="staff-form"><div class="table-scroll"><table class="staff-table"><thead><tr><th>名前</th><th>経験</th><th>区分</th><th>勤務量</th><th>基本曜日</th><th>基本開始</th><th>基本終了</th></tr></thead><tbody>${staffDraft.map(s => `<tr data-staff-id="${esc(s.id)}"><td><input aria-label="${esc(s.name)}の名前" name="name" type="text" maxlength="40" required value="${esc(s.name)}"></td>${['experience', 'category', 'preference'].map(key => `<td><select aria-label="${esc(s.name)}の${{ experience: '経験', category: '区分', preference: '勤務量' }[key]}" name="${key}">${options(labels[key], s[key])}</select></td>`).join('')}<td><div class="weekday-checks">${[1, 2, 3, 4, 5, 6, 0].map(d => `<label>${'日月火水木金土'[d]}<input type="checkbox" value="${d}" aria-label="${esc(s.name)} ${'日月火水木金土'[d]}曜日" ${s.weekdays.includes(d) ? 'checked' : ''}></label>`).join('')}</div></td><td><input type="time" name="start" aria-label="${esc(s.name)}の基本開始" value="${esc(s.start)}" required></td><td><input type="time" name="end" aria-label="${esc(s.name)}の基本終了" value="${esc(s.end)}" required></td></tr>`).join('')}</tbody></table></div><div class="actions"><button>保存</button><button class="secondary" type="button" id="add-staff">スタッフを追加</button><button class="secondary" type="button" id="cancel-staff">キャンセル</button></div><details class="help-details"><summary>保存について</summary><p>変更は保存するまで反映されません。基本時間は次の読み取りから使い、登録済みの希望はそのまま残します。入力途中で再読み込みすると、編集内容は失われます。</p></details></form></section>`;
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
      save({ ...state, staff }); candidate = null; staffDraft = null; view = 'staff'; render(); notify('基本条件を保存しました。');
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
  content.innerHTML = `<section><h2>履歴</h2><p class="muted">${esc(state.month)} · ${entries.length}件</p>${entries.map(h => `<details class="history-entry"><summary><strong>${esc(h.staffName)}</strong><span>${esc(new Date(h.at).toLocaleString('ja-JP'))}</span><span>${h.entries.length ? `${h.entries.length}日分` : '未登録に変更'}</span></summary><p class="history-message">${esc(h.text)}</p><p>${h.entries.map(e => `${e.day}日 ${caption(e)}${e.startDefaulted || e.endDefaulted ? '（基本時間）' : ''}`).map(esc).join(' ／ ') || 'セルを未登録に変更'}</p></details>`).join('') || '<p>まだ履歴はありません。</p>'}<details class="help-details"><summary>保存する内容</summary><p>反映した元文章と変更内容を、全月合わせて最新100件まで保存します。</p></details></section>`;
}
function renderCoverage() {
  const settings = state.coverage?.[state.month] || { required: 0, start: '09:00', end: '18:00' };
  const rows = summarizeCoverage({ month: state.month, staff: state.staff, requests: state.months[state.month] || {}, settings });
  const missingStaff = state.staff.filter(s => Object.keys(requests(s.id)).length === 0);
  const shortageDays = rows.filter(row => row.shortages.length);
  const configured = settings.required > 0;
  const leadingDays = dow(1);
  const monthNumber = Number(state.month.split('-')[1]);
  content.innerHTML = `<section class="coverage-view"><div class="coverage-heading"><h2>人数チェック</h2><p class="muted">${esc(state.month)} · 勤務希望からの目安です</p></div><div class="coverage-summary"><div class="coverage-stat${configured ? shortageDays.length ? ' is-warning' : ' is-good' : ''}"><span>不足がある日</span><strong>${configured ? `${shortageDays.length}<small>日</small>` : '—'}</strong></div><div class="coverage-stat${missingStaff.length ? ' is-warning' : ' is-good'}"><span>希望未登録</span><strong>${missingStaff.length}<small>人</small></strong></div><div class="coverage-stat"><span>必要人数</span><strong>${configured ? `${settings.required}<small>人</small>` : '未設定'}</strong></div></div><form id="coverage-form" class="coverage-settings"><label>必要人数<input name="required" type="number" min="0" max="30" value="${settings.required}" required></label><label>開始<input name="start" type="time" value="${esc(settings.start)}" required></label><span class="coverage-time-separator" aria-hidden="true">〜</span><label>終了<input name="end" type="time" value="${esc(settings.end)}" required></label><button>保存</button><span class="coverage-settings-note muted">全日に適用 · 0人で解除</span></form>${missingStaff.length ? `<details class="coverage-unsubmitted"><summary>希望未登録のスタッフ <span>${missingStaff.length}人</span></summary><p class="muted">今月の希望がまだ1件もありません。</p><div class="coverage-names">${missingStaff.map(s => `<span>${esc(s.name)}</span>`).join('')}</div></details>` : ''}${!configured ? '<div class="coverage-empty"><h3>何人必要ですか？</h3><p>上で人数を設定すると、不足する日がわかります。</p></div>' : `<div class="coverage-results"><div class="coverage-toolbar"><h3>日ごとの状況</h3><span class="muted">日付を押すと詳細</span></div><div class="coverage-calendar" aria-label="${esc(state.month)}の人数"><div class="coverage-weekday is-sunday">日</div>${['月', '火', '水', '木', '金'].map(day => `<div class="coverage-weekday">${day}</div>`).join('')}<div class="coverage-weekday is-saturday">土</div>${'<span class="coverage-empty-cell" aria-hidden="true"></span>'.repeat(leadingDays)}${rows.map(r => `<button type="button" class="coverage-day-button ${r.shortages.length ? 'is-shortage' : 'is-covered'}" data-coverage-day="${r.day}" aria-haspopup="dialog" aria-label="${monthNumber}月${r.day}日（${'日月火水木金土'[dow(r.day)]}） ${r.shortages.length ? `最大${Math.max(...r.shortages.map(s => s.missing))}人不足` : '充足'}"><strong>${r.day}</strong><span>${r.shortages.length ? `あと${Math.max(...r.shortages.map(s => s.missing))}人` : '充足'}</span></button>`).join('')}</div></div>`}<details class="help-details"><summary>集計について</summary><p>確定シフトではなく、登録された勤務希望を集計しています。未登録は勤務不可とは区別します。希望あり・ベテランの人数は当日全体、最少人数は設定した時間内に同時に入れる人数です。</p></details></section>`;
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
  document.querySelector('#dialog-content').innerHTML = `<h2 id="dialog-title">${monthNumber}月${row.day}日（${'日月火水木金土'[dow(row.day)]}）</h2>${row.shortages.length ? `<div class="coverage-shortages">${row.shortages.map(s => `<span class="coverage-interval"><span>${s.start}〜${s.end}</span><strong>あと${s.missing}人</strong></span>`).join('')}</div>` : `<p class="coverage-status is-good">${esc(settings.start)}〜${esc(settings.end)} · 人数を満たしています</p>`}<details class="coverage-day-details"><summary>人数の内訳</summary><dl class="coverage-breakdown"><div><dt>希望あり</dt><dd>${row.availableCount}人</dd></div><div><dt>うちベテラン</dt><dd>${row.veteranCount}人</dd></div><div><dt>未登録</dt><dd>${row.unreportedCount}人</dd></div><div><dt>同時に入れる最少人数</dt><dd>${row.minimumAvailable}人</dd></div></dl></details><div class="actions"><button type="button" id="close-dialog" autofocus>閉じる</button></div>`;
  document.querySelector('#close-dialog').onclick = () => modal.close();
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
        save(plan.nextState); monthInput.value = state.month; candidate = null; excelSession = null;
        if (!state.staff.some(s => s.id === selectedStaff)) selectedStaff = state.staff[0].id;
        view = 'month'; render(); notify(`${plan.peopleCount}名・${plan.entryCount}日分をExcelから反映しました。`);
      } catch (error) { notify(error.message); }
    };
  } catch (error) { root.innerHTML = `<p class="import-errors" role="alert">${esc(error.message)}</p><p>希望表は変更していません。名前の対応やExcelの内容を確認してください。</p>`; }
}
monthInput.onchange = () => { try { daysInMonth(monthInput.value); save({ ...state, month: monthInput.value }); candidate = null; render(); } catch (e) { monthInput.value = state.month; notify(e.message); } };
document.querySelectorAll('nav button').forEach(b => b.onclick = () => { if (busy) { notify('読み取りが終わってから画面を切り替えてください。'); return; } try { if (view === 'intake' && candidate && !candidate.needsClarification) candidate.entries = readPreviewEntries(); view = b.dataset.view; render(); } catch (e) { notify(e.message); } });
render();
