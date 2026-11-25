-- タスクリスト：工番ごとのダッシュボード・過去の工事の保管・Excel移行に必要な変更
-- supabase_setup.sql を実行した後に、Supabase の SQL Editor で1回だけ実行してください。
-- （何度実行しても壊れないように if not exists で書いています）

begin;

-- ===== 工番（工事）ごとの情報 =====
create table if not exists public.tasklist_projects (
  project_number text primary key,
  title text,                         -- 客先・件名（例：三五）。空なら工程表の tasks.customer_name を表示
  is_closed boolean not null default false,  -- true = 過去の工事（データはそのまま保管）
  closed_at date,
  created_by text,
  updated_by text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
comment on table public.tasklist_projects is 'タスクリストの工番一覧（ダッシュボード用）。is_closed=true は過去の工事';

alter table public.tasklist_projects enable row level security;
do $$ begin
  if not exists (select 1 from pg_policies where tablename = 'tasklist_projects' and policyname = 'authenticated read/write') then
    create policy "authenticated read/write" on public.tasklist_projects
      for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
  end if;
end $$;

-- ===== タスクに Excel 由来の項目を追加 =====
alter table public.tasklist_issues add column if not exists task_no integer;       -- 工番内の No.（Excel の No. を引き継ぐ）
alter table public.tasklist_issues add column if not exists notified boolean not null default false; -- Excel の「追加通知済み」
alter table public.tasklist_issues add column if not exists source_file text;     -- 移行元の Excel ファイル名（二重取り込み防止）
alter table public.tasklist_issues add column if not exists source_no integer;    -- 移行元の Excel の No.

create unique index if not exists tasklist_issues_source_uq
  on public.tasklist_issues (source_file, source_no) where source_file is not null;
create index if not exists tasklist_issues_project_task_no on public.tasklist_issues (project_number, task_no);

-- Excel の「状況」欄には日付のない書き込みもあるため、日付なしの履歴を許す
alter table public.tasklist_issue_notes alter column noted_at drop not null;

-- 分類の呼び方を Excel（１出荷前／2 現地工事／3 現地問題）に合わせて「出荷前／現地工事／出荷後対応」に統一
update public.tasklist_issues set category = '出荷後対応' where category in ('出荷後', '現地問題', '3 現地問題');

-- 既存のタスクに工番があって tasklist_projects に無いものを登録しておく
insert into public.tasklist_projects (project_number)
select distinct project_number from public.tasklist_issues
on conflict (project_number) do nothing;

commit;
