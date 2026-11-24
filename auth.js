// Supabase 設定（他アプリと同じプロジェクト。anon key は公開前提の鍵）
const S_URL = "https://dgekjzkrybrswsxlcbvh.supabase.co";
const S_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRnZWtqemtyeWJyc3dzeGxjYnZoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Njg4ODQ3MjIsImV4cCI6MjA4NDQ2MDcyMn0.BsEj53lV3p76yE9fMPTaLn7ocKTNzYPTqIAnBafYItU";

const supabaseClient = supabase.createClient(S_URL, S_KEY, {
    auth: {
        persistSession: true,
        detectSessionInUrl: true,
        storage: typeof window !== 'undefined' ? window.localStorage : undefined,
        autoRefreshToken: true,
    },
});

// このアプリはサーバー側（main.py）がログイン必須を強制しているため、ここに来た時点でほぼログイン済み。
// タスクの登録・編集・削除は、ログインしている社内の人なら誰でもできる（担当者ごとの許可リストは設けない）。
let _isLoggedIn = false;
let _currentUserEmail = '';
let _currentUserName = '';

window._getCurrentUserEmail = () => _currentUserEmail;
window._getCurrentUserName = () => _currentUserName || (_currentUserEmail ? _currentUserEmail.split('@')[0] : '');
window._isLoggedIn = () => _isLoggedIn;

function _updateUIForAuth(loggedIn) {
    _isLoggedIn = loggedIn;
    document.body.classList.toggle('not-logged-in', !loggedIn);
    if (typeof window.onAuthReady === 'function') window.onAuthReady(loggedIn);
}

supabaseClient.auth.onAuthStateChange((_event, session) => {
    const user = session?.user || null;
    _currentUserEmail = user?.email || '';
    const meta = user?.user_metadata || {};
    _currentUserName = meta.full_name || meta.name || '';
    _updateUIForAuth(!!user);
});

async function doLogout() {
    if (!confirm('ログアウトしますか？')) return;
    await supabaseClient.auth.signOut();
    location.reload();
}

// ログイン中は「ログイン」ボタンの代わりに、ユーザー表示（アイコン＋フルネーム）を出す
M365Login.mountUserChip(supabaseClient, { mountAfter: '#app_title', canEdit: () => _isLoggedIn });
