"""Excel のタスクリスト（タスクリスト_2810ML‗（三五）.xlsm など）を Supabase に移すための SQL を作る。

使い方（PowerShell / コマンドプロンプト）:
    py -3 -I tools\\excel_to_sql.py <xlsmファイル or フォルダ> [...] -o import.sql
    py -3 -I tools\\excel_to_sql.py <フォルダ> --list                 … 対象ファイルと件数の確認だけ
    py -3 -I tools\\excel_to_sql.py <運用中フォルダ> <過去分フォルダ> --closed <過去分フォルダ> -o import.sql

- フォルダを渡すと、サブフォルダも含めて「タスクリスト_<工番>…」の Excel をすべて読む（原紙・社外秘・極秘は読まない）
- 同じ工番＋機械のファイルが複数ある（「… 2.xlsm」のようなコピー）ときは、更新日時が新しいものだけ使う
- ファイルが多いときは import_01.sql, import_02.sql … に分けて書き出す。順に Supabase の SQL Editor で実行する
- 同じファイルを2回流しても、同じ「タスクリスト_<工番><機械>」＋No. のタスクは二重に入らない
- 読むシート：「タスクリスト」（進行中）と「完了タスク」。どちらも見出し行に「No.」「項目」がある行を探して読む
- 工番：タイトル（B2「2810ML 三五 タスクリスト」）またはファイル名の先頭の数字
- 客先：タイトルの2語目、無ければファイル名の（ ）の中
- --closed に指定したフォルダの工番、--auto-close なら全タスクが完了・キャンセルの工番を「過去の工事」にする

外部に何も送らない（SQL ファイルを書くだけ）。必要なのは openpyxl のみ。
"""
import argparse
import datetime as dt
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl

STATUS_LIST = ['0 未着手', '1 設計中', '2 製作・手配中', '3 納品済み', '4 施工中', '5 施工済み',
               '6 操業確認中', '7 ペンディング', '8 キャンセル', '9 現地対応', '10 完了']
DONE_STATUSES = ('10 完了', '8 キャンセル')

# Excel の見出し → DB の列
HEADER_MAP = {
    'No.': 'no', 'No': 'no',
    '日付': 'found_date', '記入日': 'found_date',
    '発見者': 'finder', '記入者': 'finder',
    '担当者': 'owner',
    '機械': 'machine',
    'ユニット': 'unit',
    '分類': 'category',
    '項目': 'item',
    '設計担当': 'design_owner',
    '施工担当': 'construction_owner',
    '進捗状況': 'status',
    '現工程完了日': 'current_process_due',
    '完了予定日': 'due_date',
    '完了日': 'completed_date',
    '詳細・対応策': 'detail',
    '状況': 'notes',
    '追加通知済み': 'notified',
}

SKIP_WORDS = ('社外秘', '極秘', '原紙')


def norm(s):
    """全角数字・全角スペースなどをそろえる"""
    return unicodedata.normalize('NFKC', str(s)).strip() if s is not None else ''


def map_category(v):
    s = norm(v)
    if not s:
        return ''
    if '出荷前' in s:
        return '出荷前'
    if '現地工事' in s:
        return '現地工事'
    if '現地問題' in s or '出荷後' in s:
        return '出荷後対応'
    return s


# 古いファイルの8段階（0 未着手 … 5 確認中・6 完了・7 ペンディング）の名前の読み替え
STATUS_ALIAS = {'確認中': '6 操業確認中', '完了': '10 完了', 'キャンセル': '8 キャンセル', 'ペンディング': '7 ペンディング'}


def map_status(v):
    """番号は時期によって意味が違うため、名前で対応づける（名前が無いときだけ番号で）"""
    s = norm(v)
    if not s:
        return STATUS_LIST[0]
    name = re.sub(r'^\d+\s*', '', s)
    for st in STATUS_LIST:
        if s == st or name == st.split(' ', 1)[1]:
            return st
    if name in STATUS_ALIAS:
        return STATUS_ALIAS[name]
    m = re.match(r'^(\d+)$', s)
    if m and int(m.group(1)) < len(STATUS_LIST):
        return STATUS_LIST[int(m.group(1))]
    return s


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = norm(v)
    m = re.match(r'^(\d{4})[/-](\d{1,2})[/-](\d{1,2})', s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


NOTE_DATE_RE = re.compile(r'^\s*(\d{1,2})/(\d{1,2})[\s　]*(.*)$')


def split_notes(text, base_date):
    """「状況」欄を履歴に分ける。行頭の M/D を日付として扱い、日付のない行は直前の履歴に続ける"""
    text = (str(text) if text is not None else '').strip()
    if not text:
        return []
    base = base_date or dt.date.today()
    notes = []
    cur = None
    for line in text.splitlines():
        m = NOTE_DATE_RE.match(line)
        d = None
        if m:
            try:
                d = dt.date(base.year, int(m.group(1)), int(m.group(2)))
                if d < base - dt.timedelta(days=60):
                    d = d.replace(year=base.year + 1)
            except ValueError:
                d = None
        if d:
            cur = {'noted_at': d, 'body': m.group(3).strip() or line.strip()}
            notes.append(cur)
        elif cur is None:
            cur = {'noted_at': None, 'body': line.strip()}
            notes.append(cur)
        else:
            cur['body'] += '\n' + line.strip()
    for n in notes:
        n['body'] = n['body'].strip()
    return [n for n in notes if n['body']]


def project_info(path, wb):
    title = ''
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            for v in row:
                if isinstance(v, str) and 'タスクリスト' in v:
                    title = norm(v)
                    break
            if title:
                break
        if title:
            break
    stem = norm(path.stem)
    pn = ''
    # 工番：4桁の数字か「3A84」のような 数字＋英字＋数字2桁（2810ML の ML は機械なので含めない）
    pn_re = r'(\d{4}|\d[A-Z]\d{2})'
    m = re.match(r'^タスクリスト[_\s]*' + pn_re, stem)
    if m:
        pn = m.group(1)
    elif title:
        m = re.match(pn_re, title)
        pn = m.group(1) if m else ''
    return pn, customer_from(stem, title)


def customer_from(stem, title):
    """客先：ファイル名の（ ）の中 → ファイル名の最後の「_」の後ろ → タイトルの2語目"""
    m = re.search(r'\((.+)\)', stem)
    if m and re.search(r'[^\d,\s]', m.group(1)):
        return re.sub(r'\s+', ' ', m.group(1)).strip()
    parts = [p.strip() for p in re.split(r'[_]', stem)[1:]]
    parts = [p for p in parts if p and not re.match(r'^[\d,\s()A-Z]+$', p)]
    if parts:
        return re.sub(r'\s+', ' ', ' '.join(parts))
    if title:
        words = [w for w in re.split(r'\s+', title.replace('タスクリスト', '')) if w]
        if len(words) >= 2:
            return words[1]
    return ''


def read_sheet(ws):
    header_row = None
    cols = {}
    for r in range(1, min(ws.max_row, 15) + 1):
        vals = [norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        if 'No.' in vals and '項目' in vals:
            header_row = r
            for c, v in enumerate(vals, start=1):
                if v in HEADER_MAP:
                    cols[HEADER_MAP[v]] = c
            break
    if header_row is None:
        return []
    # 「完了タスク」シートは「追加通知済み」の見出しが無い。進行中シートと同じく「状況」の右隣とみなす
    if 'notified' not in cols and 'notes' in cols:
        cols['notified'] = cols['notes'] + 1
    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        rec = {k: ws.cell(r, c).value for k, c in cols.items()}
        if not any(norm(v) for k, v in rec.items() if k not in ('no', 'notified', 'status', 'category')):
            continue  # No. だけ振ってある空行（分類・進捗状況だけ初期値が入っている行も空とみなす）
        rows.append(rec)
    return rows


def sql_str(v):
    if v is None or v == '':
        return 'null'
    return "'" + str(v).replace("'", "''") + "'"


def sql_date(d):
    return f"'{d.isoformat()}'" if d else 'null'


def sheet_key(path):
    """二重取り込みを防ぐキー。「タスクリスト_2783EN‗（丸一名古屋） 2.xlsm」のようなコピーも「タスクリスト_2783EN」にそろえる"""
    stem = norm(path.stem)
    m = re.match(r'^タスクリスト[_\s]*([0-9A-Za-z\-]+)', stem)
    if m:
        return 'タスクリスト_' + m.group(1)
    return re.sub(r'(\s*\(\d+\)|\s+\d+)$', '', stem)


def file_sql(path):
    """1ファイル分の SQL 行と集計を返す（工番が読めない・タスクが無いときは None）"""
    wb = openpyxl.load_workbook(path, data_only=True, keep_vba=False)
    pn, customer = project_info(path, wb)
    if not pn:
        print(f'[スキップ] 工番が読み取れません: {path}', file=sys.stderr)
        return None
    records = [(ws.title, rec) for ws in wb.worksheets for rec in read_sheet(ws)]
    if not records:
        print(f'[スキップ] タスクがありません: {path.name}', file=sys.stderr)
        return None
    # No. が空・重複している行には、そのファイルの最後の No. の次を振る
    def as_int(v):
        try:
            return int(v)
        except (TypeError, ValueError):
            return None
    used = set()
    numbered = []
    pending = []
    for sheet, rec in records:
        no = as_int(rec.get('no'))
        if no is None or no in used:
            pending.append((sheet, rec))
        else:
            used.add(no)
            numbered.append((no, rec))
    next_no = max(used, default=0) + 1
    for sheet, rec in pending:
        print(f'[注意] {path.name} / {sheet}: No. が空か重複している行に No.{next_no} を振りました（項目: {norm(rec.get("item"))[:30]}）', file=sys.stderr)
        numbered.append((next_no, rec))
        next_no += 1

    source_file = sheet_key(path)
    out = [f'-- ===== {path.name}（工番 {pn} {customer}） =====']
    out.append(
        f"insert into public.tasklist_projects (project_number, title) values ({sql_str(pn)}, {sql_str(customer)}) "
        f"on conflict (project_number) do update set title = coalesce(nullif(public.tasklist_projects.title, ''), excluded.title);")
    count = 0
    open_count = 0
    note_count = 0
    last_date = None
    for no, rec in numbered:
        found = to_date(rec.get('found_date'))
        status = map_status(rec.get('status'))
        completed = to_date(rec.get('completed_date'))
        if status not in DONE_STATUSES:
            open_count += 1
        for d in (found, completed):
            if d and (not last_date or d > last_date):
                last_date = d
        values = {
            'project_number': sql_str(pn),
            'task_no': str(no),
            'found_date': sql_date(found),
            'finder': sql_str(norm(rec.get('finder'))),
            'owner': sql_str(norm(rec.get('owner'))),
            'machine': sql_str(norm(rec.get('machine'))),
            'unit': sql_str(norm(rec.get('unit'))),
            'category': sql_str(map_category(rec.get('category'))),
            'item': sql_str(str(rec.get('item') or '').strip()),
            'design_owner': sql_str(norm(rec.get('design_owner'))),
            'construction_owner': sql_str(norm(rec.get('construction_owner'))),
            'status': sql_str(status),
            'current_process_due': sql_date(to_date(rec.get('current_process_due'))),
            'due_date': sql_date(to_date(rec.get('due_date'))),
            'completed_date': sql_date(completed),
            'detail': sql_str(str(rec.get('detail') or '').strip()),
            'notified': 'true' if norm(rec.get('notified')) else 'false',
            'source_file': sql_str(source_file),
            'source_no': str(no),
            'created_by': sql_str('Excel移行'),
            'updated_by': sql_str('Excel移行'),
        }
        cols = ', '.join(values.keys())
        vals = ', '.join(values.values())
        out.append(
            f'insert into public.tasklist_issues ({cols}) select {vals} '
            f'where not exists (select 1 from public.tasklist_issues where source_file = {sql_str(source_file)} and source_no = {no});')
        for n in split_notes(rec.get('notes'), found):
            out.append(
                'insert into public.tasklist_issue_notes (issue_id, noted_at, body, created_by) '
                f"select i.id, {sql_date(n['noted_at'])}, {sql_str(n['body'])}, 'Excel移行' from public.tasklist_issues i "
                f'where i.source_file = {sql_str(source_file)} and i.source_no = {no} '
                'and not exists (select 1 from public.tasklist_issue_notes x where x.issue_id = i.id '
                f"and x.body = {sql_str(n['body'])} and x.noted_at is not distinct from {sql_date(n['noted_at'])});")
            note_count += 1
        count += 1
    out.append('')
    return {'pn': pn, 'customer': customer, 'lines': out, 'count': count, 'open': open_count,
            'notes': note_count, 'last_date': last_date, 'path': path}


def header():
    return ['-- タスクリスト Excel 移行（excel_to_sql.py で生成 ' + dt.datetime.now().strftime('%Y-%m-%d %H:%M') + '）',
            '-- 同じファイル（工番＋機械）＋No. のタスクがすでにあれば、そのタスクと履歴は追加しない', 'begin;', '']


def collect(inputs):
    """ファイルとフォルダ（サブフォルダも）から対象を集め、同じキーのコピーは更新日時が新しいものだけ残す"""
    found = []
    for p in map(Path, inputs):
        if p.is_dir():
            found += list(p.rglob('タスクリスト*.xls*'))
        elif p.exists():
            found.append(p)
        else:
            print(f'[スキップ] 見つかりません: {p}', file=sys.stderr)
    paths = []
    for x in found:
        if x.name.startswith('~$') or x.suffix.lower() not in ('.xlsx', '.xlsm'):
            continue
        if any(w in part for part in x.parts for w in SKIP_WORDS):
            print(f'[スキップ] 原紙・社外秘など: {x.name}', file=sys.stderr)
            continue
        if not re.match(r'^タスクリスト[_\s]*\d', norm(x.stem)):
            continue  # 「タスクリストリマインド休暇日リスト」など工番の無いファイル
        paths.append(x)
    best = {}
    for x in paths:
        k = sheet_key(x)
        if k not in best:
            best[k] = x
            continue
        newer, older = (x, best[k]) if x.stat().st_mtime > best[k].stat().st_mtime else (best[k], x)
        print(f'[重複] {older} は {newer.name} の方が新しいので使いません', file=sys.stderr)
        best[k] = newer
    return sorted(best.values(), key=sheet_key)


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser(description='Excel タスクリストを Supabase 取り込み用 SQL に変換')
    ap.add_argument('inputs', nargs='+', help='xlsm/xlsx ファイル、またはそれらが入ったフォルダ（サブフォルダもたどる）')
    ap.add_argument('-o', '--output', default='import.sql',
                    help='出力する SQL ファイル。多いときは import_01.sql, import_02.sql … に分ける')
    ap.add_argument('--per-file', type=int, default=30, help='SQL 1本に入れる Excel の数（既定 30）')
    ap.add_argument('--closed', action='append', default=[],
                    help='このフォルダ（の下）にあるファイルの工番は「過去の工事」にする。何回でも指定できる')
    ap.add_argument('--auto-close', action='store_true', help='全タスクが完了・キャンセルの工番を「過去の工事」にする')
    ap.add_argument('--list', action='store_true', help='SQL は作らず、対象ファイルと件数だけ表示する')
    args = ap.parse_args()

    paths = collect(args.inputs)
    if not paths:
        print('読み込むファイルがありません', file=sys.stderr)
        return 1
    closed_dirs = [Path(c).resolve() for c in args.closed]

    def in_closed_dir(path):
        rp = path.resolve()
        return any(d == rp or d in rp.parents for d in closed_dirs)

    results = []
    for path in paths:
        try:
            res = file_sql(path)
        except Exception as e:  # 壊れたファイル・パスワード付きなど
            print(f'[スキップ] 読めません: {path}（{e}）', file=sys.stderr)
            continue
        if res:
            results.append(res)

    # 「過去の工事」は工番単位：その工番のファイルが全部過去分フォルダにある（--auto-close なら全部完了）ときだけ
    by_pn = {}
    for r in results:
        by_pn.setdefault(r['pn'], []).append(r)
    closed_pns = {}
    for pn, rs in by_pn.items():
        if all(in_closed_dir(r['path']) for r in rs) or (args.auto_close and all(r['open'] == 0 for r in rs)):
            dates = [r['last_date'] for r in rs if r['last_date']]
            closed_pns[pn] = max(dates) if dates else None

    total = 0
    for r in results:
        total += r['count']
        print(f"{r['path'].name}: 工番 {r['pn']}（{r['customer'] or '客先なし'}） タスク {r['count']} 件（未完了 {r['open']}）"
              f"・状況の履歴 {r['notes']} 件{'　→ 過去の工事' if r['pn'] in closed_pns else ''}")
    print(f'合計 {len(results)} ファイル・工番 {len(by_pn)} 件（うち過去の工事 {len(closed_pns)} 件）・タスク {total} 件')
    if args.list:
        return 0

    chunks = [[r['lines'] for r in results[i:i + args.per_file]] for i in range(0, len(results), args.per_file)]
    chunks[-1].append([
        f"update public.tasklist_projects set is_closed = true, closed_at = coalesce(closed_at, {sql_date(d)}) "
        f"where project_number = {sql_str(pn)} and not is_closed;" for pn, d in sorted(closed_pns.items())])

    out = Path(args.output)
    names = []
    for i, chunk in enumerate(chunks, start=1):
        name = out if len(chunks) == 1 else out.with_name(f'{out.stem}_{i:02d}{out.suffix}')
        body = header() + [line for lines in chunk for line in lines] + ['commit;']
        name.write_text('\n'.join(body) + '\n', encoding='utf-8')
        names.append(name)
    print('→ 書き出しました。上から順に Supabase の SQL Editor で実行してください:')
    for n in names:
        print(f'   {n}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
