// タスクリスト（2000番台工事の外観検査以降の指摘・出荷後の不具合）
// supabaseClient は auth.js で生成済み

// Excelの「進捗状況＋色」シートと同じ11段階（値・並びを変えない）
const STATUS_LIST = [
    { value: '0 未着手',       color: '#8a94a6', bg: '#eef2f7' },
    { value: '1 設計中',       color: '#6b4fbb', bg: '#f1ecfb' },
    { value: '2 製作・手配中', color: '#2a6fb0', bg: '#e8f1fa' },
    { value: '3 納品済み',     color: '#1565c0', bg: '#e3f2fd' },
    { value: '4 施工中',       color: '#b4780a', bg: '#fdf1df' },
    { value: '5 施工済み',     color: '#9c6b05', bg: '#faf0db' },
    { value: '6 操業確認中',   color: '#0d8a72', bg: '#e1f5f0' },
    { value: '7 ペンディング', color: '#b05a00', bg: '#fdeadb' },
    { value: '8 キャンセル',   color: '#999999', bg: '#f0f0f0' },
    { value: '9 現地対応',     color: '#c2410c', bg: '#fdeadb' },
    { value: '10 完了',        color: '#2e7d32', bg: '#e8f5e9' },
];
const DONE_STATUS = '10 完了';
const CANCELLED_STATUS = '8 キャンセル';

function statusInfo(value) {
    return STATUS_LIST.find(s => s.value === value) || STATUS_LIST[0];
}

// ===== 状態 =====
let _issues = [];
let _currentIssueId = null; // null = 新規登録
let _currentNotes = [];

// ===== 初期化 =====
document.addEventListener('DOMContentLoaded', () => {
    fillStatusOptions();
    bindEvents();
    const pn = new URLSearchParams(location.search).get('pn');
    if (pn) document.getElementById('filter_project').value = pn;
});

window.onAuthReady = function (loggedIn) {
    if (!loggedIn) return;
    loadAssigneeOptions();
    loadIssues();
};

function fillStatusOptions() {
    const filterSel = document.getElementById('filter_status');
    filterSel.innerHTML = '<option value="">すべてのステータス</option>' +
        STATUS_LIST.map(s => `<option value="${escapeHtml(s.value)}">${escapeHtml(s.value)}</option>`).join('');

    const formSel = document.getElementById('f_status');
    formSel.innerHTML = STATUS_LIST.map(s => `<option value="${escapeHtml(s.value)}">${escapeHtml(s.value)}</option>`).join('');
}

function bindEvents() {
    document.getElementById('filter_project').addEventListener('input', debounce(loadIssues, 300));
    document.getElementById('filter_keyword').addEventListener('input', debounce(loadIssues, 300));
    document.getElementById('filter_status').addEventListener('change', loadIssues);
    document.getElementById('filter_open_only').addEventListener('change', loadIssues);
    document.getElementById('filter_post_shipment_only').addEventListener('change', loadIssues);
    document.getElementById('new_issue_btn').addEventListener('click', () => openDialog(null));
    document.getElementById('dialog_close_btn').addEventListener('click', closeDialog);
    document.getElementById('cancel_btn').addEventListener('click', closeDialog);
    document.getElementById('save_btn').addEventListener('click', saveIssue);
    document.getElementById('archive_btn').addEventListener('click', archiveIssue);
    document.getElementById('add_note_btn').addEventListener('click', addNote);
    document.getElementById('f_project_number').addEventListener('change', () => loadMachineUnitOptions(document.getElementById('f_project_number').value.trim()));
    document.getElementById('detail_overlay').addEventListener('click', (e) => { if (e.target.id === 'detail_overlay') closeDialog(); });
}

function debounce(fn, ms) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ===== 一覧読み込み =====
async function loadIssues() {
    const project = document.getElementById('filter_project').value.trim();
    const keyword = document.getElementById('filter_keyword').value.trim();
    const status = document.getElementById('filter_status').value;
    const openOnly = document.getElementById('filter_open_only').checked;
    const postShipmentOnly = document.getElementById('filter_post_shipment_only').checked;

    let q = supabaseClient.from('tasklist_issues').select('*').eq('is_archived', false);
    if (project) q = q.ilike('project_number', `%${project}%`);
    if (status) q = q.eq('status', status);
    if (openOnly) q = q.not('status', 'in', `("${DONE_STATUS}","${CANCELLED_STATUS}")`);
    if (postShipmentOnly) q = q.eq('category', '出荷後');
    q = q.order('project_number', { ascending: true }).order('due_date', { ascending: true, nullsFirst: false });

    const { data, error } = await q;
    if (error) { showToast('読み込みに失敗しました: ' + error.message); return; }

    let rows = data || [];
    if (keyword) {
        const kw = keyword.toLowerCase();
        rows = rows.filter(r => [r.machine, r.unit, r.item, r.owner, r.design_owner, r.construction_owner]
            .some(v => (v || '').toLowerCase().includes(kw)));
    }
    _issues = rows;
    renderTable(rows);
}

async function loadNoteCounts(ids) {
    if (!ids.length) return {};
    const { data, error } = await supabaseClient.from('tasklist_issue_notes').select('issue_id').in('issue_id', ids);
    if (error || !data) return {};
    const counts = {};
    data.forEach(n => { counts[n.issue_id] = (counts[n.issue_id] || 0) + 1; });
    return counts;
}

async function renderTable(rows) {
    const tbody = document.getElementById('issue_tbody');
    document.getElementById('issue_count').textContent = `${rows.length} 件`;
    document.getElementById('empty_state').style.display = rows.length ? 'none' : '';
    if (!rows.length) { tbody.innerHTML = ''; return; }

    const counts = await loadNoteCounts(rows.map(r => r.id));

    tbody.innerHTML = rows.map(r => {
        const st = statusInfo(r.status);
        const muParts = [r.machine, r.unit].filter(Boolean).join(' / ');
        return `<tr data-id="${r.id}" class="issue-row">
            <td>${escapeHtml(r.project_number)}</td>
            <td>${escapeHtml(muParts)}</td>
            <td>${escapeHtml(r.category)}</td>
            <td class="col-item">${escapeHtml(r.item)}</td>
            <td>${escapeHtml(r.owner)}</td>
            <td><span class="status-badge" style="color:${st.color};background:${st.bg}">${escapeHtml(r.status)}</span></td>
            <td>${escapeHtml(r.due_date || '')}</td>
            <td class="col-notes">${counts[r.id] ? counts[r.id] + '件' : ''}</td>
        </tr>`;
    }).join('');

    tbody.querySelectorAll('.issue-row').forEach(tr => {
        tr.addEventListener('click', () => openDialog(Number(tr.dataset.id)));
    });
}

// ===== 担当者・機械/ユニット候補 =====
async function loadAssigneeOptions() {
    const { data, error } = await supabaseClient.from('assignee_directory').select('name,email').order('name');
    if (error || !data) return;
    const dl = document.getElementById('assignee_options');
    const seen = new Set();
    dl.innerHTML = data.filter(p => p.name && !seen.has(p.name) && seen.add(p.name))
        .map(p => `<option value="${escapeHtml(p.name)}"></option>`).join('');
}

async function loadMachineUnitOptions(projectNumber) {
    const machineDl = document.getElementById('machine_options');
    const unitDl = document.getElementById('unit_options');
    machineDl.innerHTML = '';
    unitDl.innerHTML = '';
    if (!projectNumber) return;
    const { data, error } = await supabaseClient.from('tasks').select('machine,unit').eq('project_number', projectNumber);
    if (error || !data) return;
    const machines = [...new Set(data.map(r => r.machine).filter(Boolean))];
    const units = [...new Set(data.map(r => r.unit).filter(Boolean))];
    machineDl.innerHTML = machines.map(m => `<option value="${escapeHtml(m)}"></option>`).join('');
    unitDl.innerHTML = units.map(u => `<option value="${escapeHtml(u)}"></option>`).join('');
}

// ===== 詳細・編集ダイアログ =====
function openDialog(issueId) {
    _currentIssueId = issueId;
    const issue = issueId ? _issues.find(r => r.id === issueId) : null;

    document.getElementById('dialog_title').textContent = issue ? `タスクの編集（No.${issue.id}）` : 'タスクの登録';
    document.getElementById('archive_btn').style.display = issue ? '' : 'none';
    document.getElementById('notes_section').style.display = issue ? '' : 'none';

    setVal('f_project_number', issue?.project_number || document.getElementById('filter_project').value.trim());
    setVal('f_category', issue?.category || '');
    setVal('f_machine', issue?.machine || '');
    setVal('f_unit', issue?.unit || '');
    setVal('f_found_date', issue?.found_date || today());
    setVal('f_finder', issue?.finder || window._getCurrentUserName());
    setVal('f_owner', issue?.owner || '');
    setVal('f_design_owner', issue?.design_owner || '');
    setVal('f_construction_owner', issue?.construction_owner || '');
    setVal('f_status', issue?.status || STATUS_LIST[0].value);
    setVal('f_current_process_due', issue?.current_process_due || '');
    setVal('f_due_date', issue?.due_date || '');
    setVal('f_completed_date', issue?.completed_date || '');
    setVal('f_item', issue?.item || '');
    setVal('f_detail', issue?.detail || '');
    setVal('f_new_note', '');

    loadMachineUnitOptions(document.getElementById('f_project_number').value.trim());

    if (issue) loadNotesFor(issue.id); else renderNotes([]);

    document.getElementById('detail_overlay').classList.add('open');
}

function closeDialog() {
    document.getElementById('detail_overlay').classList.remove('open');
    _currentIssueId = null;
}

function setVal(id, v) { document.getElementById(id).value = v || ''; }
function val(id) { return document.getElementById(id).value.trim(); }
function today() { return new Date().toISOString().slice(0, 10); }

async function saveIssue() {
    const projectNumber = val('f_project_number');
    if (!projectNumber) { showToast('工番を入力してください'); return; }

    const payload = {
        project_number: projectNumber,
        category: val('f_category'),
        machine: val('f_machine'),
        unit: val('f_unit'),
        found_date: val('f_found_date') || null,
        finder: val('f_finder'),
        owner: val('f_owner'),
        design_owner: val('f_design_owner'),
        construction_owner: val('f_construction_owner'),
        status: val('f_status'),
        current_process_due: val('f_current_process_due') || null,
        due_date: val('f_due_date') || null,
        completed_date: val('f_completed_date') || null,
        item: val('f_item'),
        detail: val('f_detail'),
        updated_by: window._getCurrentUserEmail(),
        updated_at: new Date().toISOString(),
    };

    let error;
    if (_currentIssueId) {
        ({ error } = await supabaseClient.from('tasklist_issues').update(payload).eq('id', _currentIssueId));
    } else {
        payload.created_by = window._getCurrentUserEmail();
        ({ error } = await supabaseClient.from('tasklist_issues').insert([payload]));
    }
    if (error) { showToast('保存に失敗しました: ' + error.message); return; }

    showToast('保存しました');
    closeDialog();
    loadIssues();
}

async function archiveIssue() {
    if (!_currentIssueId) return;
    if (!confirm('このタスクを一覧から外します。よろしいですか？')) return;
    const { error } = await supabaseClient.from('tasklist_issues')
        .update({ is_archived: true, updated_by: window._getCurrentUserEmail(), updated_at: new Date().toISOString() })
        .eq('id', _currentIssueId);
    if (error) { showToast('削除に失敗しました: ' + error.message); return; }
    showToast('一覧から外しました');
    closeDialog();
    loadIssues();
}

// ===== 状況（履歴）=====
async function loadNotesFor(issueId) {
    const { data, error } = await supabaseClient.from('tasklist_issue_notes')
        .select('*').eq('issue_id', issueId).order('noted_at', { ascending: false }).order('created_at', { ascending: false });
    if (error) { renderNotes([]); return; }
    _currentNotes = data || [];
    renderNotes(_currentNotes);
}

function renderNotes(notes) {
    const list = document.getElementById('notes_list');
    if (!notes.length) { list.innerHTML = '<div class="notes-empty">まだ記録がありません</div>'; return; }
    list.innerHTML = notes.map(n => `<div class="note-item">
        <div class="note-meta">${escapeHtml(n.noted_at || '')}${n.created_by ? '　' + escapeHtml(n.created_by) : ''}</div>
        <div class="note-body">${escapeHtml(n.body)}</div>
    </div>`).join('');
}

async function addNote() {
    if (!_currentIssueId) return;
    const body = val('f_new_note');
    if (!body) return;
    const { error } = await supabaseClient.from('tasklist_issue_notes').insert([{
        issue_id: _currentIssueId,
        noted_at: today(),
        body,
        created_by: window._getCurrentUserEmail(),
    }]);
    if (error) { showToast('追記に失敗しました: ' + error.message); return; }
    setVal('f_new_note', '');
    loadNotesFor(_currentIssueId);
}

// ===== トースト =====
let _toastTimer;
function showToast(msg) {
    const el = document.getElementById('toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(_toastTimer);
    _toastTimer = setTimeout(() => el.classList.remove('show'), 2500);
}
