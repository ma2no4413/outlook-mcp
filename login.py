"""
対話ログイン(最初に一度だけ実行する)
--------------------------------------
    python login.py

ブラウザで表示されるコードを入力するとトークンがキャッシュされ、
以後 MCP サーバは無人でトークンを更新する。
アカウントを変えたいときは --logout してからもう一度。
"""

from __future__ import annotations

import sys

from outlook_auth import (
    SCOPES,
    AuthError,
    build_app,
    cache_path,
    _save_cache,
    signed_in_account,
)

# Windows の英語環境ではコンソールが cp1252 になり、日本語の案内を出した時点で
# UnicodeEncodeError で落ちる。ログインは利用者が必ず一度は通る場所なので、
# ここで確実に読める形にしておく。
for _stream in (sys.stdout, sys.stderr):
    _reconfigure = getattr(_stream, "reconfigure", None)
    if _reconfigure is not None:
        _reconfigure(encoding="utf-8", errors="replace")


def logout() -> int:
    app, cache = build_app()
    for acc in app.get_accounts():
        app.remove_account(acc)
    _save_cache(cache)
    p = cache_path()
    if p.exists():
        p.unlink()
    print("ログアウトしました。キャッシュを削除:", p)
    return 0


def login() -> int:
    app, cache = build_app()

    # ログイン済みでも、要求スコープが増えていたら同意を取り直す必要がある。
    # ここを「ログイン済みだから何もしない」で済ませると、Azure側で権限を足した
    # あとも古いトークンが使われ続け、原因の分かりにくい403になる。
    who = signed_in_account()
    if who:
        accounts = app.get_accounts()
        got = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None
        if got and "access_token" in got:
            print(f"すでに {who} でログイン済みで、必要な権限も揃っています。")
            print(f"許可済み: {' '.join(SCOPES)}")
            print("切り替えるなら: python login.py --logout")
            return 0
        print(f"{who} でログイン済みですが、必要な権限が不足しています。")
        print(f"要求する権限: {' '.join(SCOPES)}")
        print("同意を取り直します。\n")

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("device code flow を開始できませんでした。", file=sys.stderr)
        print(flow.get("error_description", flow), file=sys.stderr)
        print(
            "\nAzureのアプリ登録で「パブリック クライアント フローを許可する」が"
            "有効になっているか確認してください(README参照)。",
            file=sys.stderr,
        )
        return 1

    print()
    print(flow["message"])  # 「https://microsoft.com/devicelogin を開いて XXXX を入力」
    print()
    print("完了するまでここで待ちます...")

    result = app.acquire_token_by_device_flow(flow)  # ここでブロックしてポーリング
    _save_cache(cache)

    if "access_token" not in result:
        print("ログインに失敗しました。", file=sys.stderr)
        print(result.get("error_description", result), file=sys.stderr)
        return 1

    print(f"\nOK: {signed_in_account()} でログインしました。")
    print("トークンキャッシュ:", cache_path())
    print("このファイルはメールボックスの鍵に相当します。共有しないでください。")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(logout() if "--logout" in sys.argv[1:] else login())
    except AuthError as e:
        print(f"設定エラー: {e}", file=sys.stderr)
        sys.exit(1)
