"""
Microsoft Graph の認証(device code flow + トークンキャッシュ)
--------------------------------------------------------------
サーバ本体(outlook_server.py)と対話ログイン(login.py)の
両方から使う。ここにはMCPの概念を持ち込まない。

方針:
  - パスワードは扱わない。ブラウザでコードを入力するだけの device code flow。
  - 一度ログインすればリフレッシュトークンがキャッシュされ、以後は無人で更新される。
  - キャッシュはこのファイルの隣に平文JSONで置く(msalの標準形式)。
    実質的にメールボックスの鍵なので、.gitignore と権限で守る。
"""

from __future__ import annotations

import os
from pathlib import Path

import msal
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# BASE_SCOPES  = 読み取り + フォルダの作成/改名/移動/削除 + メール移動 / 既読化
# EXTRA_SCOPES = 受信トレイの振分ルール(messageRules)の読み書き
# 送信(Mail.Send)は要求しない。
# offline_access / openid / profile はmsalが自動で付けるので書かない(書くとエラー)。
BASE_SCOPES = ["Mail.ReadWrite"]
EXTRA_SCOPES = ["MailboxSettings.ReadWrite"]
SCOPES = BASE_SCOPES + EXTRA_SCOPES  # login.py はこの全部で同意を取る

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class AuthError(RuntimeError):
    """設定不足・未ログインなど、利用者が手を動かせば直るもの。"""


def client_id() -> str:
    v = os.environ.get("OUTLOOK_CLIENT_ID", "").strip()
    if not v:
        raise AuthError(
            f".env に OUTLOOK_CLIENT_ID がありません。{BASE_DIR / '.env'} を確認してください。"
            " 値は Azure でアプリ登録すると得られます(README参照)。"
        )
    return v


def authority() -> str:
    # common   = 個人アカウント + 職場/学校アカウント
    # consumers= 個人(Hotmail/Outlook.com)のみ
    tenant = os.environ.get("OUTLOOK_TENANT", "common").strip() or "common"
    return f"https://login.microsoftonline.com/{tenant}"


def cache_path() -> Path:
    raw = os.environ.get("OUTLOOK_TOKEN_CACHE", "token_cache.json").strip()
    p = Path(raw).expanduser()
    return p if p.is_absolute() else BASE_DIR / p


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    p = cache_path()
    if p.exists():
        cache.deserialize(p.read_text(encoding="utf-8"))
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if not cache.has_state_changed:
        return
    p = cache_path()
    p.write_text(cache.serialize(), encoding="utf-8")
    try:
        p.chmod(0o600)  # Windowsではほぼ無視されるが、置いておいて損はない
    except OSError:
        pass


def build_app(cache: msal.SerializableTokenCache | None = None) -> tuple[msal.PublicClientApplication, msal.SerializableTokenCache]:
    cache = cache if cache is not None else _load_cache()
    app = msal.PublicClientApplication(
        client_id(), authority=authority(), token_cache=cache
    )
    return app, cache


def acquire_token_silent() -> str:
    """キャッシュ済みトークンを返す。無ければ AuthError。

    サーバ本体からはこれしか呼ばない。stdioサーバの中で対話ログインを
    始めると、クライアントのタイムアウトと噛み合わず必ず事故る。
    """
    app, cache = build_app()
    accounts = app.get_accounts()
    if not accounts:
        raise AuthError(
            "まだログインしていません。ターミナルで一度だけ次を実行してください:\n"
            f"    python {BASE_DIR / 'login.py'}"
        )
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        # 権限を増やす前に取ったトークンには EXTRA_SCOPES が入っていない。
        # そこで全体を止めるのは過剰なので、基本権限だけで取り直す。
        # ルール系ツールだけが Graph から 403 を受け、他はそのまま動く。
        result = app.acquire_token_silent(BASE_SCOPES, account=accounts[0])
    _save_cache(cache)
    if not result or "access_token" not in result:
        raise AuthError(
            "トークンの更新に失敗しました(期限切れ、またはアクセス許可の取り消し)。"
            f"再ログインしてください:\n    python {BASE_DIR / 'login.py'}"
        )
    return result["access_token"]


def signed_in_account() -> str | None:
    """ログイン中のアカウント名。未ログインなら None。診断用。"""
    try:
        app, _ = build_app()
    except AuthError:
        return None
    accounts = app.get_accounts()
    return accounts[0].get("username") if accounts else None
