import os
from pathlib import Path

import jwt
from jwt import PyJWKClient
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, RedirectResponse, JSONResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://dgekjzkrybrswsxlcbvh.supabase.co")
JWKS_URL = f"{SUPABASE_URL}/auth/v1/.well-known/jwks.json"
ISSUER = f"{SUPABASE_URL}/auth/v1"
AUDIENCE = "authenticated"
COOKIE_NAME = "sb_at"
ALLOWED_EMAIL_DOMAIN = os.environ.get("ALLOWED_EMAIL_DOMAIN", "@kusakabe.com").strip().lower()
INDEX_FILE = "index.html"
APP_TITLE = os.environ.get("APP_TITLE", "タスクリスト")

# ログインなしで見てよいパス（/login 自身と、そこが使う supabase.min.js、/api/me は中で自前判定する）
PUBLIC_PATHS = {"/login", "/api/me", "/supabase.min.js"}

# ローカルの画面確認（kdeploy preview）のときだけ有効。Cloud Run では K_SERVICE が必ず入るので効かない
LOCAL = os.environ.get("APP_ENV") == "local" and not os.environ.get("K_SERVICE")
DEV_LOGIN_EMAIL = os.environ.get("DEV_LOGIN_EMAIL", "").strip().lower() if LOCAL else ""

# Cloud Run では Dockerfile が static/ にコピーする。手元の確認ではこのフォルダ自身をそのまま見る
STATIC_DIR = Path(os.environ.get("STATIC_DIR", str(Path(__file__).resolve().parent)))

_jwk_client = PyJWKClient(JWKS_URL, cache_keys=True, lifespan=3600)


def verify_token(token: str) -> dict | None:
    try:
        signing_key = _jwk_client.get_signing_key_from_jwt(token)
        return jwt.decode(token, signing_key.key, algorithms=["ES256"], audience=AUDIENCE, issuer=ISSUER)
    except Exception:
        return None


def current_user(request: Request) -> dict | None:
    token = request.cookies.get(COOKIE_NAME)
    if not token:
        return None
    user = verify_token(token)
    if not user:
        return None
    # 同じ Supabase プロジェクトの誰のトークンでも署名検証は通るため、社内のメールドメインだけに絞る
    email = (user.get("email") or "").strip().lower()
    if not email.endswith(ALLOWED_EMAIL_DOMAIN):
        return None
    return user


app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)


@app.middleware("http")
async def security_headers(request: Request, call_next):
    response = await call_next(request)
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["Strict-Transport-Security"] = "max-age=31536000; includeSubDomains"
    return response


@app.middleware("http")
async def auth_gate(request: Request, call_next):
    path = request.url.path
    if path in PUBLIC_PATHS or current_user(request) or DEV_LOGIN_EMAIL:
        return await call_next(request)
    if path == "/":
        return RedirectResponse("/login", status_code=302)
    return JSONResponse({"error": "unauthorized"}, status_code=401)


@app.get("/")
def index():
    return FileResponse(STATIC_DIR / INDEX_FILE)


@app.get("/login")
def login_page():
    html = (Path(__file__).parent / "login.html").read_text(encoding="utf-8")
    html = html.replace("__SUPABASE_URL__", SUPABASE_URL).replace("__SUPABASE_ANON_KEY__", _anon_key())
    return HTMLResponse(html.replace("__APP_TITLE__", APP_TITLE))


@app.get("/api/me")
def api_me(request: Request):
    user = current_user(request)
    if not user and not DEV_LOGIN_EMAIL:
        return JSONResponse({"error": "unauthorized"}, status_code=401)
    return {"email": (user or {}).get("email", DEV_LOGIN_EMAIL), "local": LOCAL and bool(DEV_LOGIN_EMAIL) and not user}


def _anon_key() -> str:
    # 画面側（html/js のどこか）に書かれているのと同じ anon key（公開前提の鍵）をそのまま使う
    marker = 'const S_KEY = "'
    for path in STATIC_DIR.rglob("*"):
        if path.suffix not in (".html", ".js"):
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            continue
        if marker in text:
            i = text.index(marker) + len(marker)
            return text[i:text.index('"', i)]
    raise RuntimeError("S_KEY が見つかりません（画面側のファイルに const S_KEY = \"...\" がありません）")


app.mount("/", StaticFiles(directory=STATIC_DIR), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
