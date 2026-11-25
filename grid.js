// 工番ページの表：Excel のようにセルを選んでそのまま入力する
//  - クリックで選択 → 文字を打つと入力開始（日本語入力も可）／ダブルクリック・F2 で今の値を編集
//  - Tab・Shift+Tab で左右、Enter・Shift+Enter で上下、矢印で移動。Shift+矢印・ドラッグで範囲選択
//  - Ctrl+C・Ctrl+X・Ctrl+V で Excel との間でコピー・貼り付け。Delete で消去。Ctrl+Z で元に戻す
//  - 入力したセルはその場で保存。一番下の空行に入力すると新しいタスクになる
// app.js の _issues・_currentProject・_currentTab・TABS などを使う

// w：固定の幅（px）。flex があれば画面の幅に合わせて伸び縮みする列（w はその最小の幅、flex は配分の割合）
const GRID_COLS = [
    { key: 'task_no',             label: 'No.',          w: 40,  type: 'no' },
    { key: 'found_date',          label: '日付',         w: 84,  type: 'date' },
    { key: 'finder',              label: '発見者',       w: 60,  type: 'person' },
    { key: 'owner',               label: '担当者',       w: 60,  type: 'person' },
    { key: 'machine',             label: '機械',         w: 52,  type: 'text', opts: () => _machineOpts },
    { key: 'unit',                label: 'ユニット',     w: 60,  type: 'text', opts: () => _unitOpts },
    { key: 'category',            label: '分類',         w: 72,  type: 'choice', opts: () => CATEGORY_LIST },
    { key: 'item',                label: '項目',         w: 120, flex: 4, type: 'long' },
    { key: 'design_owner',        label: '設計担当',     w: 60,  type: 'person' },
    { key: 'construction_owner',  label: '施工担当',     w: 60,  type: 'person' },
    { key: 'status',              label: '進捗状況',     w: 104, type: 'choice', opts: () => STATUS_LIST.map(s => s.value) },
    { key: 'current_process_due', label: '現工程完了日', w: 84,  type: 'date' },
    { key: 'due_date',            label: '完了予定日',   w: 84,  type: 'date' },
    { key: 'completed_date',      label: '完了日',       w: 84,  type: 'date' },
    { key: 'detail',              label: '詳細・対応策', w: 100, flex: 3, type: 'long' },
    { key: '__note',              label: '状況（最新）', w: 100, flex: 3, type: 'note' },
];
const COL_ITEM = GRID_COLS.findIndex(c => c.key === 'item');

let _gridRows = [];                 // 表示中の行（最後の「新規行」は含めない。新規行の番号 = _gridRows.length）
let _sel = { r: 0, c: COL_ITEM, r2: 0, c2: COL_ITEM };  // r,c = アクティブセル / r2,c2 = 範囲のもう一方の端
let _edit = null;                   // 入力中 { r, c, mode: 'enter'|'edit', orig }
let _undoStack = [];
let _suggest = { items: [], idx: -1 };
let _dragging = false;
let _machineOpts = [];
let _unitOpts = [];
let _assigneeNames = [];

const $ed = () => document.getElementById('grid_editor');
const $body = () => document.getElementById('grid_body');

// ===== 表の描画 =====
function renderGrid() {
    const tab = TABS.find(t => t.key === _currentTab) || TABS[0];
    const machine = document.getElementById('pv_machine').value;
    const kw = val('pv_keyword').toLowerCase();
    _gridRows = _issues.filter(tab.test)
        .filter(r => !machine || r.machine === machine)
        .filter(r => !kw || [r.unit, r.item, r.detail, r.owner, r.finder, r.design_owner, r.construction_owner]
            .some(v => (v || '').toLowerCase().includes(kw)));
    drawGrid();
}

function drawGrid() {
    document.getElementById('grid_cols').innerHTML = GRID_COLS.map(() => '<col>').join('');
    document.getElementById('grid_head').innerHTML = '<tr>' + GRID_COLS.map(c => `<th>${escapeHtml(c.label)}</th>`).join('') + '</tr>';
    const html = _gridRows.map((issue, r) => rowHtml(issue, r)).join('') + rowHtml(null, _gridRows.length);
    $body().innerHTML = html;
    fitColumns();   // 行を入れて縦のスクロールバーが出た後の幅で合わせる
    document.getElementById('issue_count').textContent = `${_gridRows.length} 件`;
    clampSel();
    paintSel();
    placeEditor();
}

function rowHtml(issue, r) {
    const cls = issue ? (isDone(issue) ? 'g-row row-done' : 'g-row') : 'g-row g-new';
    return `<tr class="${cls}" data-r="${r}">` + GRID_COLS.map(col =>
        `<td class="${cellClass(issue, col)}">${cellHtml(issue, col)}</td>`).join('') + '</tr>';
}

function cellClass(issue, col) {
    const cls = ['g-cell', 'g-' + col.type];
    if (issue && col.key === 'due_date' && isOverdue(issue)) cls.push('overdue');
    return cls.join(' ');
}

function cellHtml(issue, col) {
    if (!issue) return col.type === 'no' ? '<span class="g-plus" title="ここに入力すると新しいタスクになります">＋</span>' : '';
    if (col.type === 'no') return escapeHtml(issue.task_no ?? '');
    if (col.type === 'date') return fmtDate(issue[col.key]);
    if (col.type === 'note') {
        const n = _latestNotes[issue.id];
        if (!n) return '';
        return `<span class="note-date">${escapeHtml(fmtDate(n.latest.noted_at))}</span> ${escapeHtml(n.latest.body)}` +
            (n.count > 1 ? `<span class="note-more">他${n.count - 1}件</span>` : '');
    }
    if (col.key === 'status') {
        if (!issue.status) return '';
        const st = statusInfo(issue.status);
        return `<span class="status-badge" style="color:${st.color};background:${st.bg}">${escapeHtml(issue.status)}</span>`;
    }
    return escapeHtml(issue[col.key]);
}

function redrawRow(issue) {
    const r = _gridRows.indexOf(issue);
    const tr = $body().rows[r];
    if (!tr) return;
    tr.className = isDone(issue) ? 'g-row row-done' : 'g-row';
    GRID_COLS.forEach((col, c) => {
        tr.cells[c].className = cellClass(issue, col);
        tr.cells[c].innerHTML = cellHtml(issue, col);
    });
    paintSel();
}

// ===== 選択 =====
function normSel() {
    return {
        r0: Math.min(_sel.r, _sel.r2), r1: Math.max(_sel.r, _sel.r2),
        c0: Math.min(_sel.c, _sel.c2), c1: Math.max(_sel.c, _sel.c2),
    };
}

function clampSel() {
    const maxR = _gridRows.length;
    const cl = (v, max) => Math.max(0, Math.min(max, v));
    _sel = { r: cl(_sel.r, maxR), c: cl(_sel.c, GRID_COLS.length - 1), r2: cl(_sel.r2, maxR), c2: cl(_sel.c2, GRID_COLS.length - 1) };
}

function cellTd(r, c) {
    const tr = $body().rows[r];
    return tr ? tr.cells[c] : null;
}

function paintSel() {
    const body = $body();
    body.querySelectorAll('.g-sel, .g-active').forEach(td => td.classList.remove('g-sel', 'g-active'));
    const { r0, r1, c0, c1 } = normSel();
    const multi = r0 !== r1 || c0 !== c1;
    for (let r = r0; r <= r1; r++) {
        for (let c = c0; c <= c1; c++) {
            const td = cellTd(r, c);
            if (td && multi) td.classList.add('g-sel');
        }
    }
    const act = cellTd(_sel.r, _sel.c);
    if (act) act.classList.add('g-active');
}

function setActive(r, c, extend) {
    if (extend) { _sel.r2 = r; _sel.c2 = c; } else _sel = { r, c, r2: r, c2: c };
    clampSel();
    paintSel();
    placeEditor();
    const td = cellTd(extend ? _sel.r2 : _sel.r, extend ? _sel.c2 : _sel.c);
    if (td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

function moveActive(dr, dc, extend) {
    const baseR = extend ? _sel.r2 : _sel.r;
    const baseC = extend ? _sel.c2 : _sel.c;
    setActive(baseR + dr, baseC + dc, extend);
}

// 入力欄（textarea）は常にアクティブセルの上にいて、キー入力・日本語入力・コピー貼り付けを受け取る
function placeEditor() {
    const ed = $ed();
    const td = cellTd(_sel.r, _sel.c);
    if (!td) return;
    const wrap = document.getElementById('grid_wrap');
    const wr = wrap.getBoundingClientRect();
    const tr = td.getBoundingClientRect();
    ed.style.left = (tr.left - wr.left + wrap.scrollLeft) + 'px';
    ed.style.top = (tr.top - wr.top + wrap.scrollTop) + 'px';
    ed.style.width = Math.max(tr.width, editorMinWidth()) + 'px';
    ed.style.height = tr.height + 'px';
    if (!_edit) {
        ed.value = selectionText() || ' ';
        ed.select();
    }
}

function editorMinWidth() {
    if (!_edit) return 0;
    const t = GRID_COLS[_edit.c].type;
    return (t === 'long' || t === 'note') ? 320 : (t === 'person' || t === 'choice' || t === 'text') ? 140 : 110;
}

// ===== セルの値と文字の変換 =====
function ymd(s) {
    if (!s) return '';
    const [y, m, d] = s.slice(0, 10).split('-').map(Number);
    return `${y}/${m}/${d}`;
}

function cellText(issue, col) {
    if (!issue) return '';
    if (col.type === 'no') return String(issue.task_no ?? '');
    if (col.type === 'date') return ymd(issue[col.key]);
    if (col.type === 'note') return _latestNotes[issue.id]?.latest.body || '';
    return issue[col.key] ?? '';
}

function nfkc(s) { return String(s ?? '').normalize('NFKC').trim(); }

function parseDateText(t) {
    const s = nfkc(t).split(/\s+/)[0];
    if (!s) return { value: null };
    let y, m, d, mm;
    if ((mm = s.match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?$/))) [y, m, d] = [mm[1], mm[2], mm[3]].map(Number);
    else if ((mm = s.match(/^(\d{1,2})[\/\-.月](\d{1,2})日?$/))) { y = new Date().getFullYear(); m = Number(mm[1]); d = Number(mm[2]); }
    else if ((mm = s.match(/^(\d{4})(\d{2})(\d{2})$/))) [y, m, d] = [mm[1], mm[2], mm[3]].map(Number);
    const dt = y ? new Date(y, m - 1, d) : null;
    if (!dt || dt.getMonth() !== m - 1 || dt.getDate() !== d) {
        return { error: `「${t}」は日付として読めません（例：11/27、2026/11/27）` };
    }
    return { value: `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}` };
}

function parseStatus(t) {
    const s = nfkc(t);
    if (!s) return { value: STATUS_LIST[0].value };
    const values = STATUS_LIST.map(x => x.value);
    if (values.includes(s)) return { value: s };
    const num = s.match(/^(\d+)/);
    if (num && Number(num[1]) < values.length) return { value: values[Number(num[1])] };
    const byName = values.find(v => v.split(' ')[1] === s) || values.find(v => v.includes(s));
    if (byName) return { value: byName };
    return { error: `「${t}」は進捗状況にありません（0〜10 の番号でも入力できます）` };
}

function parseCategory(t) {
    const s = nfkc(t);
    if (!s) return { value: '' };
    if (s.includes('出荷前') || s === '1') return { value: '出荷前' };
    if (s.includes('現地工事') || s === '2') return { value: '現地工事' };
    if (s.includes('出荷後') || s.includes('現地問題') || s === '3') return { value: '出荷後対応' };
    return { error: `「${t}」は分類にありません（出荷前・現地工事・出荷後対応）` };
}

function normalizeCell(col, text) {
    const t = String(text ?? '').replace(/\r/g, '');
    if (col.type === 'date') return parseDateText(t);
    if (col.key === 'status') return parseStatus(t);
    if (col.key === 'category') return parseCategory(t);
    if (col.type === 'long') return { value: t.replace(/\s+$/, '').replace(/^\n+/, '') };
    return { value: t.trim() };
}

// ===== クリップボード（Excel と同じタブ区切り） =====
function tsvCell(s) {
    s = String(s ?? '');
    return /[\t\n"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function selectionText() {
    const { r0, r1, c0, c1 } = normSel();
    const lines = [];
    for (let r = r0; r <= r1; r++) {
        const cells = [];
        for (let c = c0; c <= c1; c++) cells.push(tsvCell(cellText(_gridRows[r], GRID_COLS[c])));
        lines.push(cells.join('\t'));
    }
    return lines.join('\n');
}

function parseTsv(text) {
    text = text.replace(/\r\n?/g, '\n');
    if (text.endsWith('\n')) text = text.slice(0, -1);
    const rows = [[]];
    let cur = '', i = 0, quoted = false;
    while (i < text.length) {
        const ch = text[i];
        if (quoted) {
            if (ch === '"' && text[i + 1] === '"') { cur += '"'; i += 2; continue; }
            if (ch === '"') { quoted = false; i++; continue; }
            cur += ch; i++; continue;
        }
        if (ch === '"' && cur === '') { quoted = true; i++; continue; }
        if (ch === '\t') { rows[rows.length - 1].push(cur); cur = ''; i++; continue; }
        if (ch === '\n') { rows[rows.length - 1].push(cur); rows.push([]); cur = ''; i++; continue; }
        cur += ch; i++;
    }
    rows[rows.length - 1].push(cur);
    return rows;
}

function onCopy(e) {
    if (_edit) return; // 入力中は普通の文字コピー
    e.preventDefault();
    e.clipboardData.setData('text/plain', selectionText());
}

function onCut(e) {
    if (_edit) return;
    onCopy(e);
    clearSelection();
}

function onPaste(e) {
    if (_edit) return;
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    if (!text) return;
    const data = parseTsv(text);
    const { r0, r1, c0, c1 } = normSel();
    const cells = [];
    if (data.length === 1 && data[0].length === 1 && (r1 > r0 || c1 > c0)) {
        // 1つの値を選んだ範囲すべてに入れる（Excel と同じ）
        for (let r = r0; r <= Math.min(r1, _gridRows.length - 1); r++)
            for (let c = c0; c <= c1; c++) cells.push({ r, c, text: data[0][0] });
        applyChanges(cells);
        return;
    }
    let maxC = c0;
    data.forEach((row, i) => row.forEach((v, j) => {
        const c = c0 + j;
        if (c >= GRID_COLS.length) return;
        cells.push({ r: r0 + i, c, text: v });
        maxC = Math.max(maxC, c);
    }));
    applyChanges(cells).then(() => {
        _sel = { r: r0, c: c0, r2: Math.min(r0 + data.length - 1, _gridRows.length), c2: maxC };
        clampSel(); paintSel(); placeEditor();
    });
}

function clearSelection() {
    const { r0, r1, c0, c1 } = normSel();
    const cells = [];
    for (let r = r0; r <= Math.min(r1, _gridRows.length - 1); r++)
        for (let c = c0; c <= c1; c++) {
            const col = GRID_COLS[c];
            if (col.type === 'no' || col.type === 'note') continue;
            cells.push({ r, c, text: '' });
        }
    if (cells.length) applyChanges(cells);
}

// ===== 保存 =====
let _saveSeq = 0;
function setGridStatus(text, cls) {
    const el = document.getElementById('grid_status');
    el.textContent = text;
    el.className = 'grid-status ' + (cls || '');
}

async function applyChanges(cells) {
    const byRow = new Map();
    for (const { r, c, text } of cells) {
        const col = GRID_COLS[c];
        if (!col || col.type === 'no') continue;
        if (!byRow.has(r)) byRow.set(r, []);
        byRow.get(r).push({ col, text });
    }
    const errors = [];
    const updates = [];
    const notes = [];
    const creates = [];
    for (const [r, list] of [...byRow.entries()].sort((a, b) => a[0] - b[0])) {
        const issue = _gridRows[r];
        const patch = {};
        const rowNotes = [];
        for (const { col, text } of list) {
            if (col.type === 'note') { const b = String(text ?? '').trim(); if (b) rowNotes.push(b); continue; }
            const n = normalizeCell(col, text);
            if (n.error) { errors.push(n.error); continue; }
            if (issue && (issue[col.key] ?? '') === (n.value ?? '')) continue;
            patch[col.key] = n.value;
        }
        if (patch.status === DONE_STATUS && !(patch.completed_date || issue?.completed_date)) patch.completed_date = today();
        if (issue) {
            if (Object.keys(patch).length) updates.push({ issue, patch });
            rowNotes.forEach(body => notes.push({ issue, body }));
        } else if (Object.keys(patch).length || rowNotes.length) {
            creates.push({ patch, notes: rowNotes });
        }
    }
    if (errors.length) showToast(errors[0] + (errors.length > 1 ? `（ほか ${errors.length - 1} 件）` : ''));
    if (!updates.length && !notes.length && !creates.length) return;

    const seq = ++_saveSeq;
    setGridStatus('保存中…', 'saving');
    const me = window._getCurrentUserEmail();
    const now = new Date().toISOString();
    let failed = null;

    // 既存の行：画面を先に書き換えてから保存
    const undoBatch = [];
    for (const { issue, patch } of updates) {
        const old = {};
        Object.keys(patch).forEach(k => { old[k] = issue[k] ?? null; });
        undoBatch.push({ id: issue.id, old });
        Object.assign(issue, patch, { updated_by: me, updated_at: now });
        redrawRow(issue);
    }
    if (undoBatch.length) { _undoStack.push(undoBatch); if (_undoStack.length > 50) _undoStack.shift(); }
    await Promise.all(updates.map(async ({ issue, patch }) => {
        const { error } = await supabaseClient.from('tasklist_issues')
            .update({ ...patch, updated_by: me, updated_at: now }).eq('id', issue.id);
        if (error) failed = error;
    }));

    // 新しい行
    if (creates.length) {
        await ensureProject(_currentProject);
        let no = Math.max(await nextTaskNo(_currentProject), ..._issues.map(r => (r.task_no || 0) + 1));
        const defaults = {
            project_number: _currentProject,
            category: CATEGORY_LIST.includes(_currentTab) ? _currentTab : '',
            status: STATUS_LIST[0].value,
            found_date: today(),
            finder: window._getCurrentUserName(),
            machine: document.getElementById('pv_machine').value || '',
            created_by: me, updated_by: me,
        };
        const payloads = creates.map(cr => ({ ...defaults, ...cr.patch, task_no: no++ }));
        const { data, error } = await supabaseClient.from('tasklist_issues').insert(payloads).select();
        if (error) failed = error;
        else {
            (data || []).forEach((row, i) => {
                _issues.push(row);
                _gridRows.push(row);
                creates[i].notes.forEach(body => notes.push({ issue: row, body }));
            });
            drawGrid();
        }
    }

    // 状況（履歴）の追記
    if (notes.length) {
        const { error } = await supabaseClient.from('tasklist_issue_notes').insert(
            notes.map(n => ({ issue_id: n.issue.id, noted_at: today(), body: n.body, created_by: me })));
        if (error) failed = error;
        else notes.forEach(n => {
            const cur = _latestNotes[n.issue.id];
            _latestNotes[n.issue.id] = { latest: { noted_at: today(), body: n.body }, count: (cur ? cur.count : 0) + 1 };
            redrawRow(n.issue);
        });
    }

    renderTabs();
    if (failed) {
        setGridStatus('保存できませんでした', 'error');
        showToast('保存に失敗しました: ' + failed.message + '（最新の内容を読み直します）');
        loadProject();
    } else if (seq === _saveSeq) {
        setGridStatus('保存しました', 'saved');
    }
}

async function undoLast() {
    const batch = _undoStack.pop();
    if (!batch) { showToast('元に戻せる変更はありません'); return; }
    const me = window._getCurrentUserEmail();
    const now = new Date().toISOString();
    setGridStatus('保存中…', 'saving');
    let failed = null;
    await Promise.all(batch.map(async ({ id, old }) => {
        const issue = _issues.find(r => r.id === id);
        if (!issue) return;
        Object.assign(issue, old, { updated_by: me, updated_at: now });
        redrawRow(issue);
        const { error } = await supabaseClient.from('tasklist_issues').update({ ...old, updated_by: me, updated_at: now }).eq('id', id);
        if (error) failed = error;
    }));
    renderTabs();
    setGridStatus(failed ? '保存できませんでした' : '元に戻しました', failed ? 'error' : 'saved');
    if (failed) showToast('元に戻せませんでした: ' + failed.message);
}

// ===== 入力 =====
function startEdit(mode, showAll) {
    const col = GRID_COLS[_sel.c];
    const issue = _gridRows[_sel.r];
    if (col.type === 'no') { if (issue) openDialog(issue.id); return; }
    _sel.r2 = _sel.r; _sel.c2 = _sel.c; paintSel();
    _edit = { r: _sel.r, c: _sel.c, mode, orig: col.type === 'note' ? '' : cellText(issue, col) };
    const ed = $ed();
    ed.value = _edit.orig;
    showEditor();
    ed.focus();
    ed.setSelectionRange(ed.value.length, ed.value.length);
    updateSuggest(!!showAll);
}

// 選んだセルでそのまま文字を打ち始めたとき（日本語入力の開始も含む）
function beginTyping() {
    const col = GRID_COLS[_sel.c];
    if (col.type === 'no') return false;
    _sel.r2 = _sel.r; _sel.c2 = _sel.c; paintSel();
    _edit = { r: _sel.r, c: _sel.c, mode: 'enter', orig: col.type === 'note' ? '' : cellText(_gridRows[_sel.r], col) };
    showEditor();
    return true;
}

function showEditor() {
    const ed = $ed();
    const col = GRID_COLS[_edit.c];
    ed.classList.add('editing');
    ed.placeholder = col.type === 'note' ? '状況を追記（今日の日付で記録）' : '';
    placeEditor();
    autosize();
}

function autosize() {
    const ed = $ed();
    const td = cellTd(_sel.r, _sel.c);
    const h = td ? td.getBoundingClientRect().height : 30;
    ed.style.height = 'auto';
    ed.style.height = Math.max(h, ed.scrollHeight + 2) + 'px';
}

function finishEdit(commit, dr, dc) {
    if (!_edit) return;
    const { r, c, orig } = _edit;
    const text = $ed().value;
    _edit = null;
    const ed = $ed();
    ed.classList.remove('editing');
    ed.placeholder = '';
    hideSuggest();
    if (commit && text !== orig) applyChanges([{ r, c, text }]);
    if (dr || dc) moveActive(dr || 0, dc || 0, false);
    else placeEditor();
}

// ===== 候補リスト（分類・進捗状況・担当者・機械・ユニット） =====
function optionsFor(col) {
    if (col.type === 'person') return _assigneeNames;
    return col.opts ? col.opts() : null;
}

function updateSuggest(showAll) {
    const box = document.getElementById('grid_suggest');
    if (!_edit) { hideSuggest(); return; }
    const col = GRID_COLS[_edit.c];
    const all = optionsFor(col);
    if (!all || !all.length) { hideSuggest(); return; }
    const value = $ed().value;
    const q = nfkc(value).toLowerCase();
    let items;
    if (showAll || (col.type === 'choice' && (!q || value === _edit.orig))) items = all.slice();
    else if (!q) items = [];
    else items = all.filter(o => nfkc(o).toLowerCase().includes(q));
    if (items.length === 1 && items[0] === value && col.type !== 'choice') items = [];
    items = items.slice(0, 40);
    if (!items.length) { hideSuggest(); return; }
    let idx = -1;
    if (col.type === 'choice') { idx = items.indexOf(value); if (idx < 0) idx = 0; }
    _suggest = { items, idx };
    box.innerHTML = items.map((o, i) => `<div class="sg-item${i === idx ? ' active' : ''}" data-i="${i}">${escapeHtml(o)}</div>`).join('');
    const ed = $ed();
    box.style.left = ed.style.left;
    box.style.top = (parseFloat(ed.style.top) + ed.offsetHeight) + 'px';
    box.style.minWidth = ed.style.width;
    box.style.display = 'block';
    const act = box.querySelector('.sg-item.active');
    if (act) act.scrollIntoView({ block: 'nearest' });
}

function hideSuggest() {
    _suggest = { items: [], idx: -1 };
    const box = document.getElementById('grid_suggest');
    if (box) box.style.display = 'none';
}

function moveSuggest(d) {
    if (!_suggest.items.length) return false;
    _suggest.idx = Math.max(0, Math.min(_suggest.items.length - 1, _suggest.idx + d));
    const box = document.getElementById('grid_suggest');
    box.querySelectorAll('.sg-item').forEach((el, i) => el.classList.toggle('active', i === _suggest.idx));
    const act = box.querySelector('.sg-item.active');
    if (act) act.scrollIntoView({ block: 'nearest' });
    return true;
}

function takeSuggest() {
    if (_suggest.idx >= 0 && _suggest.items[_suggest.idx] != null) $ed().value = _suggest.items[_suggest.idx];
}

// ===== キー操作 =====
function onEditorKeyDown(e) {
    if (e.isComposing || e.keyCode === 229) return; // 日本語の変換中は何もしない
    if (_edit) return onEditKey(e);
    onNavKey(e);
}

function onNavKey(e) {
    const ctrl = e.ctrlKey || e.metaKey;
    const maxR = _gridRows.length;
    if (ctrl) {
        const k = e.key.toLowerCase();
        if (k === 'z') { e.preventDefault(); undoLast(); return; }
        if (k === 'a') { e.preventDefault(); _sel = { r: 0, c: 0, r2: maxR, c2: GRID_COLS.length - 1 }; paintSel(); placeEditor(); return; }
        if (e.key === 'ArrowUp') { e.preventDefault(); setActive(0, e.shiftKey ? _sel.c2 : _sel.c, e.shiftKey); return; }
        if (e.key === 'ArrowDown') { e.preventDefault(); setActive(maxR, e.shiftKey ? _sel.c2 : _sel.c, e.shiftKey); return; }
        if (e.key === 'ArrowLeft') { e.preventDefault(); setActive(e.shiftKey ? _sel.r2 : _sel.r, 0, e.shiftKey); return; }
        if (e.key === 'ArrowRight') { e.preventDefault(); setActive(e.shiftKey ? _sel.r2 : _sel.r, GRID_COLS.length - 1, e.shiftKey); return; }
        if (e.key === 'Home') { e.preventDefault(); setActive(0, 0, false); return; }
        return; // Ctrl+C / X / V はブラウザのコピー・貼り付けイベントで処理
    }
    switch (e.key) {
        case 'ArrowUp':    e.preventDefault(); moveActive(-1, 0, e.shiftKey); return;
        case 'ArrowDown':
            e.preventDefault();
            if (e.altKey) { startEdit('edit', true); return; }
            moveActive(1, 0, e.shiftKey); return;
        case 'ArrowLeft':  e.preventDefault(); moveActive(0, -1, e.shiftKey); return;
        case 'ArrowRight': e.preventDefault(); moveActive(0, 1, e.shiftKey); return;
        case 'Tab':        e.preventDefault(); moveActive(0, e.shiftKey ? -1 : 1, false); return;
        case 'Enter':      e.preventDefault(); moveActive(e.shiftKey ? -1 : 1, 0, false); return;
        case 'Home':       e.preventDefault(); setActive(_sel.r, 0, e.shiftKey); return;
        case 'End':        e.preventDefault(); setActive(_sel.r, GRID_COLS.length - 1, e.shiftKey); return;
        case 'PageDown':   e.preventDefault(); moveActive(15, 0, e.shiftKey); return;
        case 'PageUp':     e.preventDefault(); moveActive(-15, 0, e.shiftKey); return;
        case 'F2':         e.preventDefault(); startEdit('edit'); return;
        case 'Delete':     e.preventDefault(); clearSelection(); return;
        case 'Backspace':
            e.preventDefault();
            if (beginTyping()) { $ed().value = ''; autosize(); updateSuggest(false); }
            return;
        case 'Escape':     _sel.r2 = _sel.r; _sel.c2 = _sel.c; paintSel(); placeEditor(); return;
    }
    // それ以外の文字キーはそのまま textarea に入り、input イベントで入力開始になる
}

function onEditKey(e) {
    const sugOpen = _suggest.items.length > 0;
    switch (e.key) {
        case 'Enter':
            e.preventDefault();
            if (e.altKey) { // Alt+Enter でセル内改行（Excel と同じ）
                const ed = $ed();
                ed.setRangeText('\n', ed.selectionStart, ed.selectionEnd, 'end');
                autosize();
                return;
            }
            if (sugOpen) takeSuggest();
            finishEdit(true, e.shiftKey ? -1 : 1, 0);
            return;
        case 'Tab':
            e.preventDefault();
            if (sugOpen) takeSuggest();
            finishEdit(true, 0, e.shiftKey ? -1 : 1);
            return;
        case 'Escape':
            e.preventDefault();
            if (sugOpen && GRID_COLS[_edit.c].type !== 'choice') { hideSuggest(); return; }
            finishEdit(false);
            return;
        case 'ArrowDown':
        case 'ArrowUp':
            if (sugOpen) { e.preventDefault(); moveSuggest(e.key === 'ArrowDown' ? 1 : -1); return; }
            if (_edit.mode === 'enter') { e.preventDefault(); finishEdit(true, e.key === 'ArrowDown' ? 1 : -1, 0); }
            return;
        case 'ArrowLeft':
        case 'ArrowRight':
            if (_edit.mode === 'enter') { e.preventDefault(); finishEdit(true, 0, e.key === 'ArrowRight' ? 1 : -1); }
            return;
    }
}

function onEditorInput() {
    if (!_edit) {
        if (!beginTyping()) { placeEditor(); return; }
    }
    autosize();
    updateSuggest(false);
}

function onCompositionStart() {
    if (!_edit) beginTyping();
}

function onCompositionEnd() {
    if (!_edit) placeEditor(); // No. 列など入力できないセルで変換したときは元に戻す
}

// ===== マウス =====
function cellPos(target) {
    const td = target.closest('td');
    if (!td || !td.parentElement.dataset.r) return null;
    return { r: Number(td.parentElement.dataset.r), c: td.cellIndex };
}

function onGridMouseDown(e) {
    if (e.button !== 0) return;
    const p = cellPos(e.target);
    if (!p) return;
    e.preventDefault();
    if (_edit) finishEdit(true);
    if (e.shiftKey) setActive(p.r, p.c, true);
    else setActive(p.r, p.c, false);
    _dragging = true;
    $ed().focus({ preventScroll: true });
}

function onGridMouseOver(e) {
    if (!_dragging) return;
    const p = cellPos(e.target);
    if (!p || (p.r === _sel.r2 && p.c === _sel.c2)) return;
    _sel.r2 = p.r; _sel.c2 = p.c;
    paintSel();
}

function onGridDblClick(e) {
    const p = cellPos(e.target);
    if (!p) return;
    _sel = { r: p.r, c: p.c, r2: p.r, c2: p.c };
    startEdit('edit', GRID_COLS[p.c].type === 'choice');
}

// 「＋ タスクを登録」：一番下の空行の「項目」に入力を始める
// 列の幅を表の枠の幅に合わせる：固定の列はそのまま、項目・詳細・状況の列で残りを分け合う。
// 画面が狭くて最小の幅に収まらないときだけ横にスクロールする
function fitColumns() {
    const cols = document.querySelectorAll('#grid_cols col');
    if (cols.length !== GRID_COLS.length) return;
    const wrap = document.getElementById('grid_wrap');
    const avail = wrap.clientWidth;
    const fixed = GRID_COLS.reduce((a, c) => a + c.w, 0);
    const flexSum = GRID_COLS.reduce((a, c) => a + (c.flex || 0), 0);
    const extra = Math.max(0, avail - fixed);
    let total = 0;
    GRID_COLS.forEach((c, i) => {
        const w = c.flex ? Math.floor(c.w + extra * c.flex / flexSum) : c.w;
        cols[i].style.width = w + 'px';
        total += w;
    });
    document.getElementById('grid').style.width = total + 'px';
}

function gridAddRow() {
    setActive(_gridRows.length, COL_ITEM, false);
    $ed().focus({ preventScroll: true });
    startEdit('edit');
}

function initGrid() {
    const ed = $ed();
    ed.addEventListener('keydown', onEditorKeyDown);
    ed.addEventListener('input', onEditorInput);
    ed.addEventListener('compositionstart', onCompositionStart);
    ed.addEventListener('compositionend', onCompositionEnd);
    ed.addEventListener('copy', onCopy);
    ed.addEventListener('cut', onCut);
    ed.addEventListener('paste', onPaste);
    ed.addEventListener('blur', () => { if (_edit) finishEdit(true); });
    const body = $body();
    body.addEventListener('mousedown', onGridMouseDown);
    body.addEventListener('mouseover', onGridMouseOver);
    body.addEventListener('dblclick', onGridDblClick);
    document.addEventListener('mouseup', () => { if (_dragging) { _dragging = false; placeEditor(); } });
    const box = document.getElementById('grid_suggest');
    box.addEventListener('mousedown', (e) => {
        e.preventDefault();
        const it = e.target.closest('.sg-item');
        if (!it) return;
        $ed().value = _suggest.items[Number(it.dataset.i)];
        finishEdit(true);
    });
    // 枠の幅が変わったら（画面の大きさ・スクロールバーの出入り）列の幅を合わせ直す
    let lastW = 0;
    new ResizeObserver(() => {
        const w = document.getElementById('grid_wrap').clientWidth;
        if (w === lastW || !w) return;
        lastW = w;
        fitColumns();
        placeEditor();
    }).observe(document.getElementById('grid_wrap'));
    window.addEventListener('resize', () => placeEditor());
}

document.addEventListener('DOMContentLoaded', initGrid);
