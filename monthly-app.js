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
if (PUBLIC_DEMO) document.querySelector('#usage-note').textContent = '仮公開デモ：未確定の勤務希望表です。日付だけなら基本9:00〜18:00。空欄は未登録、×は勤務不可です。文章は簡易ルールで読み取り、AIは使いません。入力はこのブラウザ内だけに保存され、他の端末とは共有されません。';
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
  content.innerHTML = `<section><h2>1. メッセージを入力</h2><div class="form-top"><label>対象スタッフ<select id="intake-staff" ${busy ? 'disabled' : ''}>${state.staff.map(p => `<option value="${esc(p.id)}"${p.id === s.id ? ' selected' : ''}>${esc(p.name)}</option>`).join('')}</select></label><span class="muted">${esc(state.month)} ／ 時刻省略時 ${esc(s.start)}〜${esc(s.end)}</span></div><p class="muted">1人分のメッセージを貼り付けてください。スタッフ名は基本条件の画面で変更できます。</p><div class="example-buttons">${['1,3,4でー', '10,17,25,31でお願いします！！', '3日は10時から16時、4日は15時から', '毎週火曜と木曜に入れます。15日は休みです。'].map((t, i) => `<button class="secondary" data-example="${i}" ${busy ? 'disabled' : ''}>${esc(t)}</button>`).join('')}</div><label for="request-text">勤務希望の文章</label><textarea id="request-text" maxlength="1500" placeholder="例：1,3,4でー" ${busy ? 'disabled' : ''}>${esc(draft)}</textarea><div class="actions"><button id="parse" ${busy ? 'disabled' : ''}>${busy ? '読み取り中…' : '文章を読み取る'}</button><button class="secondary" id="from-weekdays" ${busy ? 'disabled' : ''}>基本曜日から候補を作る</button><span class="muted" id="ai-status">${PUBLIC_DEMO ? '公開版：簡易読み取り（AIなし）' : '接続状況を確認中'}</span></div><div id="preview"></div></section>`;
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
  el.innerHTML = `<div class="preview"><h2>2. 読み取り結果を確認</h2><p>${esc(c.explanation)}</p><p class="muted">${c.entries.length}日分。日付・可否・時刻を修正でき、反映のチェックを外した日は除外します。同じ日に既存の内容がある場合、この候補で更新します。</p><div class="table-scroll"><table><thead><tr><th>反映</th><th>日付</th><th>可否</th><th>開始</th><th>終了</th><th>補完・既存の内容</th></tr></thead><tbody>${c.entries.map((e, index) => `<tr data-preview-index="${index}"><td><input type="checkbox" aria-label="${e.day}日を反映" checked></td><td><input type="number" min="1" max="${daysInMonth(state.month)}" value="${e.day}" aria-label="候補${index + 1}の日付">日</td><td><select aria-label="${e.day}日の可否"><option value="available"${e.status === 'available' ? ' selected' : ''}>勤務希望</option><option value="unavailable"${e.status === 'unavailable' ? ' selected' : ''}>勤務不可</option></select></td><td><input type="time" aria-label="${e.day}日の開始" data-part="start" value="${esc(e.start)}" ${e.status === 'unavailable' ? 'disabled' : ''}></td><td><input type="time" aria-label="${e.day}日の終了" data-part="end" value="${esc(e.end)}" ${e.status === 'unavailable' ? 'disabled' : ''}></td><td>${e.startDefaulted || e.endDefaulted ? '<span class="inferred">時刻省略のため基本時間を補完</span><br>' : ''}${requests(selectedStaff)[e.day] ? `<span class="conflict">現在：${esc(caption(requests(selectedStaff)[e.day]))}</span>` : '新規'}</td></tr>`).join('')}</tbody></table></div><button id="apply">確認した内容を月間希望表に反映</button></div>`;
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
  content.innerHTML = `<section><div class="section-heading"><h2>${esc(state.month)} 勤務希望表（未確定）</h2><button id="export" ${busy ? 'disabled' : ''}>Excel出力</button></div><p class="muted">セルを押すと手修正できます。上段が開始、下段が終了です。空欄＝未登録、×＝勤務不可。</p><p class="scroll-hint muted">横にスクロールして末日まで確認できます。</p>${[[1, Math.min(16, count)], [17, count]].map(([from, to]) => `<h3>${from}日〜${to}日</h3><div class="table-scroll"><table class="month-table"><thead><tr><th>名前</th>${Array.from({ length: to - from + 1 }, (_, i) => { const d = i + from, w = dow(d); return `<th class="${w === 0 ? 'sun' : w === 6 ? 'sat' : ''}">${d}<small>${'日月火水木金土'[w]}</small></th>`; }).join('')}</tr></thead><tbody>${state.staff.map(s => `<tr><th scope="row">${esc(s.name)}</th>${Array.from({ length: to - from + 1 }, (_, i) => { const day = from + i, entry = requests(s.id)[day]; return `<td><button class="day-cell" data-id="${esc(s.id)}" data-day="${day}" aria-label="${esc(s.name)} ${day}日 ${esc(caption(entry))}">${!entry ? '&nbsp;' : entry.status === 'unavailable' ? '<span class="off">×</span>' : `<span>${esc(entry.start)}</span><span>${esc(entry.end)}</span>`}</button></td>`; }).join('')}</tr>`).join('')}</tbody></table></div>`).join('')}</section>`;
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
  content.innerHTML = `<section><div class="section-heading"><h2>基本条件一覧</h2><button id="open-staff-edit">基本条件を入力・編集</button></div><p class="muted">保存した条件を確認する画面です。${state.staff.length}名を登録しています。</p><div class="table-scroll"><table class="staff-read-table"><thead><tr><th>名前</th><th>経験</th><th>区分</th><th>勤務量</th><th>基本曜日</th><th>基本時間</th></tr></thead><tbody>${state.staff.map(s => `<tr><th scope="row">${esc(s.name)}</th><td>${esc(s.experience)}</td><td>${esc(s.category)}</td><td>${esc(s.preference)}</td><td>${[1, 2, 3, 4, 5, 6, 0].filter(d => s.weekdays.includes(d)).map(d => '日月火水木金土'[d]).join('・') || '未設定'}</td><td>${esc(s.start)}〜${esc(s.end)}</td></tr>`).join('')}</tbody></table></div><p class="muted">基本曜日・時間は希望の候補作成に使います。経験・区分・勤務量による自動割り当ては今後の対応です。</p></section>`;
  document.querySelector('#open-staff-edit').onclick = () => { view = 'staff-edit'; render(); };
}
function renderStaffEdit() {
  staffDraft ||= structuredClone(state.staff);
  content.innerHTML = `<section><h2>基本条件を入力・編集</h2><p class="muted">名前・経験・区分・希望する勤務量・基本曜日・時間を入力し、最後に保存してください。変更中の内容は、保存するまで一覧や月間表に反映されません。</p><form id="staff-form"><div class="table-scroll"><table class="staff-table"><thead><tr><th>名前</th><th>経験</th><th>区分</th><th>勤務量</th><th>基本曜日</th><th>基本開始</th><th>基本終了</th></tr></thead><tbody>${staffDraft.map(s => `<tr data-staff-id="${esc(s.id)}"><td><input aria-label="${esc(s.name)}の名前" name="name" type="text" maxlength="40" required value="${esc(s.name)}"></td>${['experience', 'category', 'preference'].map(key => `<td><select aria-label="${esc(s.name)}の${{ experience: '経験', category: '区分', preference: '勤務量' }[key]}" name="${key}">${options(labels[key], s[key])}</select></td>`).join('')}<td><div class="weekday-checks">${[1, 2, 3, 4, 5, 6, 0].map(d => `<label>${'日月火水木金土'[d]}<input type="checkbox" value="${d}" aria-label="${esc(s.name)} ${'日月火水木金土'[d]}曜日" ${s.weekdays.includes(d) ? 'checked' : ''}></label>`).join('')}</div></td><td><input type="time" name="start" aria-label="${esc(s.name)}の基本開始" value="${esc(s.start)}" required></td><td><input type="time" name="end" aria-label="${esc(s.name)}の基本終了" value="${esc(s.end)}" required></td></tr>`).join('')}</tbody></table></div><div class="actions"><button>保存して一覧を見る</button><button class="secondary" type="button" id="add-staff">スタッフを追加</button><button class="secondary" type="button" id="cancel-staff">変更を取り消して一覧へ</button></div><p class="muted">保存後の基本時間は次の読み取りから使います。登録済みの勤務希望は変更しません。入力途中の内容は画面の切り替えでは残り、ブラウザの再読み込みで失われます。</p></form></section>`;
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
  content.innerHTML = `<section><h2>取り込み履歴</h2><p class="muted">対象月の記録です。全月合わせて最新100件を保存します。元文章は確認済みの反映時に保存します。</p>${state.history.filter(h => h.month === state.month).map(h => `<article><h3>${esc(h.staffName)} ／ ${esc(new Date(h.at).toLocaleString('ja-JP'))}</h3><p class="history-message">${esc(h.text)}</p><p>${h.entries.map(e => `${e.day}日 ${caption(e)}${e.startDefaulted || e.endDefaulted ? '（基本時間を補完）' : ''}`).map(esc).join(' ／ ') || 'セルを未登録に変更'}</p><hr></article>`).join('') || '<p>まだ反映した内容はありません。</p>'}</section>`;
}
function renderCoverage() {
  const settings = state.coverage?.[state.month] || { required: 0, start: '09:00', end: '18:00' };
  const rows = summarizeCoverage({ month: state.month, staff: state.staff, requests: state.months[state.month] || {}, settings });
  const missingStaff = state.staff.filter(s => Object.keys(requests(s.id)).length === 0);
  content.innerHTML = `<section><h2>${esc(state.month)} 不足・未登録の確認</h2><p>勤務希望をもとにした参考です。確定シフトの人数ではありません。必要人数を設定すると、時間帯ごとの不足候補を表示します。</p><form id="coverage-form" class="small-form"><label>必要人数（0＝未設定）<input name="required" type="number" min="0" max="30" value="${settings.required}" required></label><label>確認する開始<input name="start" type="time" value="${esc(settings.start)}" required></label><label>確認する終了<input name="end" type="time" value="${esc(settings.end)}" required></label><button>この月の確認条件を保存</button></form><p class="feedback">希望が未登録のスタッフ：${missingStaff.length ? missingStaff.map(s => esc(s.name)).join('、') : 'なし'}</p><p class="muted">勤務希望ありは当日の人数、時間内の最少人数は同時に勤務できる人数です。未登録の日は勤務不可と区別しています。必要人数はこの月の全日に共通です。</p><div class="table-scroll"><table><thead><tr><th>日付</th><th>勤務希望あり</th><th>うちベテラン</th><th>未登録</th><th>時間内の最少人数</th><th>不足候補</th></tr></thead><tbody>${rows.map(r => `<tr class="${r.shortages.length ? 'shortage-row' : ''}"><th>${r.day}日（${'日月火水木金土'[dow(r.day)]}）</th><td>${r.availableCount}人</td><td>${r.veteranCount}人</td><td>${r.unreportedCount}人</td><td>${r.minimumAvailable}人</td><td>${settings.required === 0 ? '必要人数を設定してください' : r.shortages.map(s => `${s.start}〜${s.end}：${s.missing}人不足`).join('<br>') || '希望人数は足りています'}</td></tr>`).join('')}</tbody></table></div></section>`;
  document.querySelector('#coverage-form').onsubmit = e => {
    e.preventDefault();
    try {
      const form = new FormData(e.currentTarget), nextSettings = { required: Number(form.get('required')), start: form.get('start'), end: form.get('end') };
      summarizeCoverage({ month: state.month, staff: state.staff, requests: state.months[state.month] || {}, settings: nextSettings });
      save({ ...state, coverage: { ...state.coverage, [state.month]: nextSettings } }); render(); notify('人数を確認する条件を保存しました。');
    } catch (error) { notify(error.message); }
  };
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
    notify('Excelファイルを出力しました。勤務希望表・取込用・基本条件・取り込み履歴の4シートです。');
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
  content.innerHTML = `<section><h2>Excel取り込み・出力</h2><p>Excelファイルはこのブラウザ内で読み書きします。取り込みは、対象月・スタッフ・変更内容を確認してから反映します。</p><h3>Excelへ出力</h3><div class="small-form"><label>出力するスタッフ<select id="export-person"><option value="">全員</option>${state.staff.map(s => `<option value="${esc(s.id)}">${esc(s.name)}</option>`).join('')}</select></label><button id="export-full" ${busy ? 'disabled' : ''}>現在の希望をExcel出力</button><button class="secondary" id="export-template" ${busy ? 'disabled' : ''}>空の提出用Excelを出力</button></div><p class="muted">提出用ファイルをスタッフに渡し、記入後のファイルを下で集約できます。出力は「勤務希望表」「取込用」「基本条件」「取り込み履歴」の4シートです。</p><h3>Excelから取り込み</h3><p class="muted">.xlsx（5MB以内・1シート1000行／40列以内）。「対象月・名前・日・可否・開始・終了」の一覧、または前半・後半に分かれた開始／終了2段の表に対応します。空欄は既存の希望を消しません。</p><label class="excel-file">Excelファイルを選択<input id="excel-file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ${busy ? 'disabled' : ''}></label>${busy ? '<p role="status">Excelを処理中です…</p>' : ''}${session?.workbook ? `<p>選択中：${esc(session.filename)}</p><label>取り込むシート<select id="excel-sheet">${session.workbook.worksheets.map(s => `<option${s.name === session.sheet ? ' selected' : ''}>${esc(s.name)}</option>`).join('')}</select></label><p class="muted">「取込用」があれば最初に選択します。「勤務希望表」をExcelで編集した場合は、上のシートも「勤務希望表」に切り替えてください。履歴は取り込み対象にしません。</p>` : ''}<div id="excel-review"></div></section>`;
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
  root.innerHTML = `<div class="preview"><h3>取り込み先：${esc(parsed.month)}</h3><p class="feedback">${parsed.month !== state.month ? `画面の対象月（${esc(state.month)}）とは異なります。反映後は${esc(parsed.month)}へ切り替えます。` : '画面と同じ対象月です。'} ファイルを選んだだけでは反映しません。</p>${parsed.warnings.length ? `<ul>${parsed.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul>` : ''}<h3>名前の対応を確認</h3><div class="table-scroll"><table><thead><tr><th>Excelの名前</th><th>取り込み先</th></tr></thead><tbody>${parsed.staff.map((s, i) => `<tr><th>${esc(s.name)}</th><td><select data-import-person="${i}" aria-label="${esc(s.name)}の取り込み先"><option value="new">新しいスタッフとして追加</option><option value="skip">取り込まない</option>${state.staff.map(p => `<option value="${esc(p.id)}">${esc(p.name)}</option>`).join('')}</select></td></tr>`).join('')}</tbody></table></div>${session.conditions.length ? `<label><input type="checkbox" id="excel-conditions" ${session.importConditions ? 'checked' : ''}>基本条件も取り込む（${session.conditions.length}名分。該当者の経験・区分・勤務量・曜日・基本時間を更新）</label>` : ''}${session.conditionsError ? `<p class="import-errors">基本条件は取り込めません：${esc(session.conditionsError)}</p>` : ''}<div id="excel-plan"></div></div>`;
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
    root.innerHTML = `<h3>反映前の確認</h3><p>${plan.peopleCount}名・${plan.entryCount}日分、新しいスタッフ${plan.newStaffCount}名。<strong>既存の希望への上書き${plan.overwrites}件</strong>です。ファイルに書かれていない日はそのまま残します。</p>${staffDraft ? '<p class="import-errors">基本条件に入力途中の内容があります。先に「基本条件を入力」で保存するか、変更を取り消してください。</p>' : ''}<div class="excel-preview"><table><thead><tr><th>スタッフ</th><th>日</th><th>現在</th><th>反映後</th><th>確認</th></tr></thead><tbody>${plan.changes.map(c => `<tr><th>${esc(c.staffName)}</th><td>${c.day}日</td><td>${esc(caption(c.previous))}</td><td>${esc(caption(c.next))}</td><td>${c.next.startDefaulted || c.next.endDefaulted ? '基本時間で補完' : ''}${c.previous ? ' ／ 上書き' : ''}</td></tr>`).join('')}</tbody></table></div><div class="actions"><label><input type="checkbox" id="excel-confirm">対象月・名前・時刻・上書き内容を確認しました</label><button id="excel-apply" disabled>確認した内容をまとめて反映</button></div><p class="muted">取り込み履歴にファイル名と反映内容を保存します。</p>`;
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
