// タスクリスト（2000番台工事の外観検査以降の指摘・出荷後の不具合）
// 画面は2つ：ダッシュボード（工番の一覧）→ 工番ごとのタスク一覧（Excel のシートのように分類ごとのタブ）
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

// Excel の分類（１出荷前／2 現地工事／3 現地問題）に対応
const CATEGORY_LIST = ['出荷前', '現地工事', '出荷後対応'];

function statusInfo(value) {
    return STATUS_LIST.find(s => s.value === value) || STATUS_LIST[0];
}
function isDone(r) { return r.status === DONE_STATUS || r.status === CANCELLED_STATUS; }
function isOverdue(r) { return !isDone(r) && r.due_date && r.due_date < today(); }

// 工番ページのタブ（Excel のシートの代わり）
const TABS = [
    { key: 'open',       label: '未完了すべて', test: r => !isDone(r) },
    { key: '出荷前',     label: '出荷前',       test: r => !isDone(r) && r.category === '出荷前' },
    { key: '現地工事',   label: '現地工事',     test: r => !isDone(r) && r.category === '現地工事' },
    { key: '出荷後対応', label: '出荷後対応',   test: r => !isDone(r) && r.category === '出荷後対応' },
    { key: 'none',       label: '分類なし',     test: r => !isDone(r) && !CATEGORY_LIST.includes(r.category), hideIfEmpty: true },
    { key: 'done',       label: '完了済み',     test: r => isDone(r) },
    { key: 'all',        label: 'すべて',       test: () => true },
];

// ===== 状態 =====
let _dashMode = 'open';        // 'open' = 進行中 / 'closed' = 過去の工事
let _projects = [];            // ダッシュボード用にまとめた工番の一覧
let _currentProject = null;    // 表示中の工番（文字列）
let _currentProjectRow = null; // 表示中の工番の tasklist_projects の行（＋autoTitle）
let _currentTab = 'open';
let _issues = [];              // 表示中の工番のタスク
let _latestNotes = {};         // issue_id → { latest, count }
let _currentIssueId = null;    // null = 新規登録
let _projectDialogMode = 'new';

// ===== 初期化 =====
document.addEventListener('DOMContentLoaded', () => {
    fillSelectOptions();
    bindEvents();
});

window.onAuthReady = function (loggedIn) {
    if (!loggedIn) return;
    loadAssigneeOptions();
    routeFromUrl();
};

window.addEventListener('popstate', routeFromUrl);

function routeFromUrl() {
    const params = new URLSearchParams(location.search);
    const pn = params.get('pn');
    if (pn) openProject(pn, params.get('tab') || 'open', false);
    else showDashboard(false);
}

function fillSelectOptions() {
    const opts = STATUS_LIST.map(s => `<option value="${escapeHtml(s.value)}">${escapeHtml(s.value)}</option>`).join('');
    document.getElementById('f_status').innerHTML = opts;
    document.getElementById('f_category').innerHTML = '<option value="">（分類なし）</option>' +
        CATEGORY_LIST.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
}

function bindEvents() {
    // ダッシュボード
    document.querySelectorAll('#dash_seg .seg-btn').forEach(b => b.addEventListener('click', () => {
        _dashMode = b.dataset.mode;
        document.querySelectorAll('#dash_seg .seg-btn').forEach(x => x.classList.toggle('active', x === b));
        renderDashboard();
    }));
    document.getElementById('dash_search').addEventListener('input', debounce(renderDashboard, 200));
    document.getElementById('new_project_btn').addEventListener('click', () => openProjectDialog('new'));

    // 工番ページ
    document.getElementById('back_btn').addEventListener('click', () => showDashboard(true));
    document.getElementById('edit_project_btn').addEventListener('click', () => openProjectDialog('edit'));
    document.getElementById('new_issue_btn').addEventListener('click', gridAddRow);
    document.getElementById('pv_machine').addEventListener('change', renderGrid);
    document.getElementById('pv_keyword').addEventListener('input', debounce(renderGrid, 200));

    // タスクのダイアログ
    document.getElementById('dialog_close_btn').addEventListener('click', closeDialog);
    document.getElementById('cancel_btn').addEventListener('click', closeDialog);
    document.getElementById('save_btn').addEventListener('click', saveIssue);
    document.getElementById('archive_btn').addEventListener('click', archiveIssue);
    document.getElementById('add_note_btn').addEventListener('click', addNote);
    document.getElementById('f_project_number').addEventListener('change', () => loadMachineUnitOptions(val('f_project_number')));
    document.getElementById('detail_overlay').addEventListener('click', (e) => { if (e.target.id === 'detail_overlay') closeDialog(); });

    // 工事のダイアログ
    document.getElementById('project_close_btn').addEventListener('click', closeProjectDialog);
    document.getElementById('project_cancel_btn').addEventListener('click', closeProjectDialog);
    document.getElementById('project_save_btn').addEventListener('click', saveProject);
    document.getElementById('p_close_btn').addEventListener('click', toggleProjectClosed);
    document.getElementById('project_overlay').addEventListener('click', (e) => { if (e.target.id === 'project_overlay') closeProjectDialog(); });

    document.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape') return;
        if (document.getElementById('detail_overlay').classList.contains('open')) closeDialog();
        else if (document.getElementById('project_overlay').classList.contains('open')) closeProjectDialog();
    });
}

// ===== 共通 =====
function debounce(fn, ms) {
    let t;
    return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function setVal(id, v) { document.getElementById(id).value = v || ''; }
function val(id) { return document.getElementById(id).value.trim(); }

// 端末の日付（日本時間）で YYYY-MM-DD
function today() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// 今年なら M/D、それ以外は YYYY/M/D
function fmtDate(s) {
    if (!s) return '';
    const [y, m, d] = s.slice(0, 10).split('-').map(Number);
    return y === new Date().getFullYear() ? `${m}/${d}` : `${y}/${m}/${d}`;
}

// Supabase は1回に1000行までしか返さないので、ページを分けて全部読む
async function fetchAll(makeQuery) {
    const PAGE = 1000;
    let all = [];
    for (let from = 0; ; from += PAGE) {
        const { data, error } = await makeQuery().range(from, from + PAGE - 1);
        if (error) return { data: null, error };
        all = all.concat(data || []);
        if (!data || data.length < PAGE) break;
    }
    return { data: all, error: null };
}

function pushUrl(params) {
    const qs = new URLSearchParams(params).toString();
    history.pushState(null, '', qs ? `?${qs}` : location.pathname);
}

// ===== ダッシュボード =====
async function showDashboard(push) {
    _currentProject = null;
    if (push) pushUrl({});
    document.getElementById('dashboard_view').style.display = '';
    document.getElementById('project_view').style.display = 'none';
    document.title = 'タスクリスト';
    await loadDashboard();
}

async function loadDashboard() {
    const [issuesRes, projectsRes] = await Promise.all([
        fetchAll(() => supabaseClient.from('tasklist_issues')
            .select('id,project_number,category,status,due_date,updated_at')
            .eq('is_archived', false).order('id')),
        supabaseClient.from('tasklist_projects').select('*'),
    ]);
    if (issuesRes.error) { showToast('読み込みに失敗しました: ' + issuesRes.error.message); return; }
    if (projectsRes.error) { showToast('工事一覧の読み込みに失敗しました: ' + projectsRes.error.message); return; }

    const map = new Map();
    const ensure = (pn) => {
        if (!map.has(pn)) map.set(pn, {
            project_number: pn, title: '', is_closed: false, closed_at: null,
            total: 0, done: 0, overdue: 0, byCat: { '出荷前': 0, '現地工事': 0, '出荷後対応': 0, 'none': 0 },
            nextDue: null, lastUpdated: null,
        });
        return map.get(pn);
    };
    (projectsRes.data || []).forEach(p => Object.assign(ensure(p.project_number), {
        title: p.title || '', is_closed: !!p.is_closed, closed_at: p.closed_at,
    }));
    issuesRes.data.forEach(r => {
        const p = ensure(r.project_number);
        p.total++;
        if (isDone(r)) p.done++;
        else {
            p.byCat[CATEGORY_LIST.includes(r.category) ? r.category : 'none']++;
            if (isOverdue(r)) p.overdue++;
            if (r.due_date && (!p.nextDue || r.due_date < p.nextDue)) p.nextDue = r.due_date;
        }
        if (r.updated_at && (!p.lastUpdated || r.updated_at > p.lastUpdated)) p.lastUpdated = r.updated_at;
    });
    _projects = [...map.values()].sort((a, b) => b.project_number.localeCompare(a.project_number, 'ja', { numeric: true }));

    await fillCustomerNames(_projects.filter(p => !p.title));
    renderDashboard();
}

// 客先・件名が未入力の工番は、工程表（tasks）の客先名・工事内容を表示する
async function fillCustomerNames(projects) {
    if (!projects.length) return;
    const nums = projects.map(p => p.project_number);
    const { data, error } = await fetchAll(() => supabaseClient.from('tasks')
        .select('id,project_number,customer_name,project_details')
        .in('project_number', nums).not('customer_name', 'is', null).order('id'));
    if (error || !data) return;
    const found = {};
    data.forEach(t => {
        if (!t.customer_name || found[t.project_number]) return;
        found[t.project_number] = [t.customer_name, t.project_details].filter(Boolean).join('　');
    });
    projects.forEach(p => { p.autoTitle = found[p.project_number] || ''; });
}

function renderDashboard() {
    const kw = val('dash_search').toLowerCase();
    const openCount = _projects.filter(p => !p.is_closed).length;
    document.getElementById('dash_count_open').textContent = openCount;
    document.getElementById('dash_count_closed').textContent = _projects.length - openCount;

    const list = _projects
        .filter(p => _dashMode === 'closed' ? p.is_closed : !p.is_closed)
        .filter(p => !kw || [p.project_number, p.title, p.autoTitle].some(v => (v || '').toLowerCase().includes(kw)));

    const grid = document.getElementById('project_grid');
    const empty = document.getElementById('dash_empty');
    empty.style.display = list.length ? 'none' : '';
    empty.textContent = kw ? '該当する工事はありません'
        : _dashMode === 'closed' ? '過去の工事はまだありません（工番ページの「工事の設定」から移せます）'
        : '進行中の工事はありません。「＋ 工事を追加」から登録してください';

    grid.innerHTML = list.map(p => {
        const open = p.total - p.done;
        const pct = p.total ? Math.round(p.done / p.total * 100) : 0;
        const title = p.title || p.autoTitle || '';
        const cats = [['出荷前', '出荷前'], ['現地工事', '現地工事'], ['出荷後対応', '出荷後対応']]
            .map(([k, label]) => `<div class="cat-cell"><div class="cat-num${p.byCat[k] ? '' : ' zero'}">${p.byCat[k]}</div><div class="cat-label">${label}</div></div>`).join('')
            + `<div class="cat-cell"><div class="cat-num done${p.done ? '' : ' zero'}">${p.done}</div><div class="cat-label">完了</div></div>`;
        const flags = [
            p.overdue ? `<span class="flag flag-late">期限切れ ${p.overdue}</span>` : '',
            p.byCat.none ? `<span class="flag">分類なし ${p.byCat.none}</span>` : '',
            p.is_closed && p.closed_at ? `<span class="flag">${fmtDate(p.closed_at)} 完了</span>` : '',
        ].join('');
        return `<button type="button" class="project-card${p.is_closed ? ' closed' : ''}" data-pn="${escapeHtml(p.project_number)}">
            <div class="card-head">
                <span class="card-number">${escapeHtml(p.project_number)}</span>
                <span class="card-title">${escapeHtml(title)}</span>
            </div>
            <div class="card-open"><b>${open}</b> 件 未完了 <span class="card-total">／ 全 ${p.total} 件</span></div>
            <div class="cat-row">${cats}</div>
            <div class="progress"><div class="progress-bar" style="width:${pct}%"></div></div>
            <div class="card-foot">
                <span>${p.nextDue ? '次の完了予定 ' + fmtDate(p.nextDue) : ''}</span>
                <span class="flags">${flags}</span>
            </div>
        </button>`;
    }).join('');

    grid.querySelectorAll('.project-card').forEach(el => el.addEventListener('click', () => openProject(el.dataset.pn, 'open', true)));
}

// ===== 工番ごとのタスク一覧 =====
async function openProject(pn, tab, push) {
    _currentProject = pn;
    _currentTab = TABS.some(t => t.key === tab) ? tab : 'open';
    if (push) pushUrl({ pn, tab: _currentTab });
    document.getElementById('dashboard_view').style.display = 'none';
    document.getElementById('project_view').style.display = '';
    document.getElementById('pv_keyword').value = '';
    document.getElementById('grid_body').innerHTML = '';
    document.title = `${pn} タスクリスト`;
    await loadProject();
    document.getElementById('grid_editor').focus({ preventScroll: true });
}

async function loadProject() {
    const pn = _currentProject;
    const [issuesRes, projRes] = await Promise.all([
        fetchAll(() => supabaseClient.from('tasklist_issues').select('*')
            .eq('project_number', pn).eq('is_archived', false).order('id')),
        supabaseClient.from('tasklist_projects').select('*').eq('project_number', pn).maybeSingle(),
    ]);
    if (pn !== _currentProject) return; // 読み込み中に別の工番へ移った
    if (issuesRes.error) { showToast('読み込みに失敗しました: ' + issuesRes.error.message); return; }

    const proj = projRes.data || { project_number: pn, title: '', is_closed: false };
    if (!proj.title) {
        const tmp = { project_number: pn };
        await fillCustomerNames([tmp]);
        proj.autoTitle = tmp.autoTitle || '';
    }
    _currentProjectRow = proj;
    document.getElementById('pv_number').textContent = pn;
    document.getElementById('pv_title').textContent = proj.title || proj.autoTitle || '';
    document.getElementById('pv_closed_tag').style.display = proj.is_closed ? '' : 'none';

    _issues = issuesRes.data.sort((a, b) =>
        (a.task_no ?? 1e9) - (b.task_no ?? 1e9) || a.id - b.id);
    _latestNotes = await loadLatestNotes(_issues.map(r => r.id));

    const machines = [...new Set(_issues.map(r => r.machine).filter(Boolean))].sort();
    const sel = document.getElementById('pv_machine');
    const prev = sel.value;
    sel.innerHTML = '<option value="">すべての機械</option>' + machines.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
    sel.value = machines.includes(prev) ? prev : '';

    renderTabs();
    renderGrid();
    loadMachineUnitOptions(pn);
}

async function loadLatestNotes(ids) {
    const result = {};
    for (let i = 0; i < ids.length; i += 100) {
        const chunk = ids.slice(i, i + 100);
        const { data, error } = await supabaseClient.from('tasklist_issue_notes')
            .select('issue_id,noted_at,body,created_at').in('issue_id', chunk)
            .order('noted_at', { ascending: false, nullsFirst: false }).order('created_at', { ascending: false });
        if (error || !data) continue;
        data.forEach(n => {
            if (!result[n.issue_id]) result[n.issue_id] = { latest: n, count: 0 };
            result[n.issue_id].count++;
        });
    }
    return result;
}

function renderTabs() {
    const el = document.getElementById('pv_tabs');
    el.innerHTML = TABS.map(t => {
        const n = _issues.filter(t.test).length;
        if (t.hideIfEmpty && !n && _currentTab !== t.key) return '';
        return `<button type="button" class="tab${t.key === _currentTab ? ' active' : ''}${t.key === 'done' ? ' tab-done' : ''}" data-tab="${t.key}">${t.label}<span class="tab-count">${n}</span></button>`;
    }).join('');
    el.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => {
        _currentTab = b.dataset.tab;
        history.replaceState(null, '', `?${new URLSearchParams({ pn: _currentProject, tab: _currentTab })}`);
        renderTabs();
        renderGrid();
    }));
}

// ===== 担当者・機械/ユニット候補 =====
async function loadAssigneeOptions() {
    const { data, error } = await supabaseClient.from('assignee_directory').select('name,email').order('name');
    if (error || !data) return;
    const dl = document.getElementById('assignee_options');
    const seen = new Set();
    _assigneeNames = data.filter(p => p.name && !seen.has(p.name) && seen.add(p.name)).map(p => p.name);
    dl.innerHTML = _assigneeNames.map(n => `<option value="${escapeHtml(n)}"></option>`).join('');
}

async function loadMachineUnitOptions(projectNumber) {
    const machineDl = document.getElementById('machine_options');
    const unitDl = document.getElementById('unit_options');
    machineDl.innerHTML = '';
    unitDl.innerHTML = '';
    if (!projectNumber) return;
    const machines = new Set();
    const units = new Set();
    if (projectNumber === _currentProject) {
        _issues.forEach(r => { if (r.machine) machines.add(r.machine); if (r.unit) units.add(r.unit); });
    }
    const { data } = await supabaseClient.from('tasks').select('machine,unit').eq('project_number', projectNumber).limit(1000);
    (data || []).forEach(r => { if (r.machine) machines.add(r.machine); if (r.unit) units.add(r.unit); });
    if (projectNumber === _currentProject) { _machineOpts = [...machines].sort(); _unitOpts = [...units].sort(); }
    machineDl.innerHTML = [...machines].sort().map(m => `<option value="${escapeHtml(m)}"></option>`).join('');
    unitDl.innerHTML = [...units].sort().map(u => `<option value="${escapeHtml(u)}"></option>`).join('');
}

// ===== タスクの詳細・編集ダイアログ =====
function openDialog(issueId) {
    _currentIssueId = issueId;
    const issue = issueId ? _issues.find(r => r.id === issueId) : null;

    document.getElementById('dialog_title').textContent = issue
        ? `タスクの編集（${issue.project_number} No.${issue.task_no ?? '-'}）` : 'タスクの登録';
    document.getElementById('archive_btn').style.display = issue ? '' : 'none';
    document.getElementById('notes_section').style.display = issue ? '' : 'none';

    // 新規登録は、いま開いているタブの分類を初期値にする
    const tabCategory = CATEGORY_LIST.includes(_currentTab) ? _currentTab : '';
    setVal('f_project_number', issue?.project_number || _currentProject || '');
    setVal('f_category', issue ? issue.category : tabCategory);
    if (issue && issue.category && !CATEGORY_LIST.includes(issue.category)) {
        // 古い分類名のまま残っているものも選べるように一時的に足す
        const sel = document.getElementById('f_category');
        sel.insertAdjacentHTML('beforeend', `<option value="${escapeHtml(issue.category)}" data-temp="1">${escapeHtml(issue.category)}</option>`);
        sel.value = issue.category;
    }
    setVal('f_machine', issue?.machine || document.getElementById('pv_machine').value || '');
    setVal('f_unit', issue?.unit || '');
    setVal('f_found_date', issue ? issue.found_date : today());
    setVal('f_finder', issue ? issue.finder : window._getCurrentUserName());
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

    loadMachineUnitOptions(val('f_project_number'));

    if (issue) loadNotesFor(issue.id); else renderNotes([]);

    document.getElementById('detail_overlay').classList.add('open');
}

function closeDialog() {
    document.getElementById('detail_overlay').classList.remove('open');
    document.querySelectorAll('#f_category option[data-temp]').forEach(o => o.remove());
    _currentIssueId = null;
}

async function ensureProject(projectNumber) {
    await supabaseClient.from('tasklist_projects').upsert(
        [{ project_number: projectNumber, created_by: window._getCurrentUserEmail(), updated_by: window._getCurrentUserEmail() }],
        { onConflict: 'project_number', ignoreDuplicates: true });
}

async function nextTaskNo(projectNumber) {
    const { data } = await supabaseClient.from('tasklist_issues').select('task_no')
        .eq('project_number', projectNumber).not('task_no', 'is', null)
        .order('task_no', { ascending: false }).limit(1);
    return ((data && data[0] && data[0].task_no) || 0) + 1;
}

async function saveIssue() {
    const projectNumber = val('f_project_number');
    if (!projectNumber) { showToast('工番を入力してください'); return; }
    if (!val('f_item')) { showToast('項目（タスク内容）を入力してください'); return; }

    const status = val('f_status');
    let completedDate = val('f_completed_date') || null;
    // 完了にしたのに完了日が空なら、今日を入れる
    if (status === DONE_STATUS && !completedDate) completedDate = today();

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
        status,
        current_process_due: val('f_current_process_due') || null,
        due_date: val('f_due_date') || null,
        completed_date: completedDate,
        item: val('f_item'),
        detail: val('f_detail'),
        updated_by: window._getCurrentUserEmail(),
        updated_at: new Date().toISOString(),
    };

    const btn = document.getElementById('save_btn');
    btn.disabled = true;
    try {
        await ensureProject(projectNumber);
        const current = _currentIssueId ? _issues.find(r => r.id === _currentIssueId) : null;
        // 新規、または別の工番へ移したときは、その工番の最後の No. の次を振る
        if (!current || current.project_number !== projectNumber) payload.task_no = await nextTaskNo(projectNumber);

        let error;
        if (current) {
            ({ error } = await supabaseClient.from('tasklist_issues').update(payload).eq('id', _currentIssueId));
        } else {
            payload.created_by = window._getCurrentUserEmail();
            ({ error } = await supabaseClient.from('tasklist_issues').insert([payload]));
        }
        if (error) { showToast('保存に失敗しました: ' + error.message); return; }
    } finally {
        btn.disabled = false;
    }

    showToast('保存しました');
    closeDialog();
    if (projectNumber !== _currentProject) openProject(projectNumber, _currentTab, true);
    else loadProject();
}

async function archiveIssue() {
    if (!_currentIssueId) return;
    if (!confirm('このタスクを一覧から外します。よろしいですか？\n（完了したタスクは「進捗状況」を 10 完了 にしてください。完了済みタブに残ります）')) return;
    const { error } = await supabaseClient.from('tasklist_issues')
        .update({ is_archived: true, updated_by: window._getCurrentUserEmail(), updated_at: new Date().toISOString() })
        .eq('id', _currentIssueId);
    if (error) { showToast('削除に失敗しました: ' + error.message); return; }
    showToast('一覧から外しました');
    closeDialog();
    loadProject();
}

// ===== 状況（履歴）=====
async function loadNotesFor(issueId) {
    const { data, error } = await supabaseClient.from('tasklist_issue_notes')
        .select('*').eq('issue_id', issueId)
        .order('noted_at', { ascending: false, nullsFirst: false }).order('created_at', { ascending: false });
    if (issueId !== _currentIssueId) return;
    renderNotes(error ? [] : (data || []));
}

function renderNotes(notes) {
    const list = document.getElementById('notes_list');
    if (!notes.length) { list.innerHTML = '<div class="notes-empty">まだ記録がありません</div>'; return; }
    list.innerHTML = notes.map(n => `<div class="note-item">
        <div class="note-meta">${escapeHtml(n.noted_at || '日付なし')}${n.created_by ? '　' + escapeHtml(n.created_by) : ''}</div>
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
    _latestNotes = await loadLatestNotes(_issues.map(r => r.id));
    renderGrid();
}

// ===== 工事の追加・設定 =====
function openProjectDialog(mode) {
    _projectDialogMode = mode;
    const editing = mode === 'edit' && _currentProjectRow;
    document.getElementById('project_dialog_title').textContent = editing ? `工事の設定（${_currentProject}）` : '工事を追加';
    setVal('p_number', editing ? _currentProject : '');
    document.getElementById('p_number').disabled = !!editing;
    setVal('p_title', editing ? (_currentProjectRow.title || '') : '');
    document.getElementById('p_title').placeholder = editing && _currentProjectRow.autoTitle
        ? `空欄なら「${_currentProjectRow.autoTitle}」（工程表の客先名）` : '空欄なら工程表の客先名を表示';

    const closeRow = document.getElementById('p_close_row');
    closeRow.style.display = editing ? '' : 'none';
    if (editing) {
        const closed = !!_currentProjectRow.is_closed;
        document.getElementById('p_close_label').textContent = closed ? 'この工事は「過去の工事」です' : '工事が終わったら';
        document.getElementById('p_close_help').textContent = closed
            ? '進行中の一覧に戻すこともできます'
            : 'ダッシュボードの「過去の工事」に移します。タスクのデータはそのまま保管されます';
        const btn = document.getElementById('p_close_btn');
        btn.textContent = closed ? '進行中に戻す' : '工事完了（過去の工事へ）';
        btn.className = closed ? 'btn' : 'btn btn-danger';
    }
    document.getElementById('project_overlay').classList.add('open');
    setTimeout(() => document.getElementById(editing ? 'p_title' : 'p_number').focus(), 0);
}

function closeProjectDialog() {
    document.getElementById('project_overlay').classList.remove('open');
}

async function saveProject() {
    const pn = val('p_number');
    if (!pn) { showToast('工番を入力してください'); return; }
    const row = { project_number: pn, title: val('p_title'), updated_by: window._getCurrentUserEmail(), updated_at: new Date().toISOString() };
    if (_projectDialogMode === 'new') row.created_by = window._getCurrentUserEmail();
    const { error } = await supabaseClient.from('tasklist_projects').upsert([row], { onConflict: 'project_number' });
    if (error) { showToast('保存に失敗しました: ' + error.message); return; }
    closeProjectDialog();
    showToast('保存しました');
    openProject(pn, 'open', _projectDialogMode === 'new');
}

async function toggleProjectClosed() {
    const closed = !!_currentProjectRow?.is_closed;
    if (!closed) {
        const openCount = _issues.filter(r => !isDone(r)).length;
        const msg = openCount
            ? `未完了のタスクが ${openCount} 件あります。それでも「過去の工事」に移しますか？`
            : 'この工事を「過去の工事」に移します。よろしいですか？';
        if (!confirm(msg)) return;
    }
    const { error } = await supabaseClient.from('tasklist_projects').upsert([{
        project_number: _currentProject,
        is_closed: !closed,
        closed_at: closed ? null : today(),
        updated_by: window._getCurrentUserEmail(),
        updated_at: new Date().toISOString(),
    }], { onConflict: 'project_number' });
    if (error) { showToast('変更に失敗しました: ' + error.message); return; }
    closeProjectDialog();
    showToast(closed ? '進行中に戻しました' : '過去の工事に移しました');
    loadProject();
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
