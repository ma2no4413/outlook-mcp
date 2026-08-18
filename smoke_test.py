"""
スモークテスト — サーバを実際に stdio で起動して外形を検証する
----------------------------------------------------------------
    .venv/bin/python smoke_test.py

pytest(test_outlook.py)が関数単体を見るのに対し、こちらは
「MCPクライアントから見てサーバがどう振る舞うか」を見る。

Microsoft Graph には接続しない。ログイン前でも実行でき、
むしろ未ログイン状態でも例外で落ちないことを確認するのが主目的。
ログイン済みの環境でも同じ判定基準で通る。

終了コード: 全項目パスなら 0、ひとつでも落ちたら 1。
"""

from __future__ import annotations

import asyncio
import sys
from datetime import datetime, timezone
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def force_utf8_output() -> None:
    """出力を UTF-8 に固定する。

    Windows のコンソール既定は環境の言語で決まり、英語環境では cp1252 になる。
    このファイルの出力は日本語なので、そのままだと最初の print で
    UnicodeEncodeError になり、テストが1件も走らないまま落ちる。
    日本語版 Windows(cp932)では通ってしまうため、開発機では気づけない。
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


force_utf8_output()

BASE_DIR = Path(__file__).resolve().parent

EXPECTED_TOOLS = {
    "check_config": {"read_only": True, "destructive": None, "required": []},
    "list_folders": {"read_only": True, "destructive": None, "required": []},
    "search_messages": {"read_only": True, "destructive": None, "required": []},
    "get_message": {"read_only": True, "destructive": None, "required": ["message_id"]},
    "list_rules": {"read_only": True, "destructive": None, "required": []},
    "create_draft": {"read_only": False, "destructive": False, "required": ["to", "subject", "body"]},
    "draft_reply": {"read_only": False, "destructive": False, "required": ["message_id", "body"]},
    "create_folder": {"read_only": False, "destructive": False, "required": ["name"]},
    "rename_folder": {"read_only": False, "destructive": False, "required": ["folder", "new_name"]},
    "move_folder": {"read_only": False, "destructive": False, "required": ["folder"]},
    "move_messages": {"read_only": False, "destructive": False, "required": ["message_ids", "folder"]},
    "move_by_search": {"read_only": False, "destructive": False, "required": ["dest"]},
    "archive_messages": {"read_only": False, "destructive": False, "required": ["message_ids"]},
    "mark_messages_read": {"read_only": False, "destructive": False, "required": ["message_ids"]},
    "mark_read_by_search": {"read_only": False, "destructive": True, "required": []},
    "create_rule": {"read_only": False, "destructive": False, "required": ["name"]},
    "move_to_trash": {"read_only": False, "destructive": True, "required": ["message_ids"]},
    "delete_folder": {"read_only": False, "destructive": True, "required": ["folder"]},
    "delete_rule": {"read_only": False, "destructive": True, "required": ["rule"]},
}

results: list[tuple[str, str, bool, str]] = []


def check(test_id: str, name: str, ok: bool, detail: str = "") -> None:
    results.append((test_id, name, ok, detail))
    mark = "PASS" if ok else "FAIL"
    print(f"[{mark}] {test_id} {name}")
    if detail:
        for line in detail.splitlines():
            print(f"         {line}")


def text_of(result) -> str:
    """call_tool の戻りから本文テキストを取り出す。"""
    parts = [getattr(c, "text", "") for c in (result.content or [])]
    return "\n".join(p for p in parts if p)


async def run() -> int:
    params = StdioServerParameters(
        command=sys.executable, args=[str(BASE_DIR / "outlook_server.py")]
    )
    async with stdio_client(params) as (reader, writer):
        async with ClientSession(reader, writer) as session:
            init = await session.initialize()

            # --- S-01 ハンドシェイク ---
            # バージョンも見る。空のまま公開すると、クライアントにもレジストリにも
            # 「版が分からないサーバ」として出る。実際 v0.2.0 まで空だった。
            import outlook_server

            name = init.server_info.name
            version = init.server_info.version
            check("S-01", "stdio でハンドシェイクが成立し、サーバ名とバージョンを申告する",
                  name == "outlook" and version == outlook_server.__version__ and bool(version),
                  f"server_info = {name!r} / {version!r}")

            # --- S-02 ツールが過不足なく公開される ---
            tools = {t.name: t for t in (await session.list_tools()).tools}
            check("S-02", f"ツールが {len(EXPECTED_TOOLS)} 個、名前が一致",
                  set(tools) == set(EXPECTED_TOOLS),
                  f"公開: {sorted(tools)}")

            # --- S-03 入力スキーマが引数から正しく生成される ---
            # handle_errors デコレータが functools.wraps を失うと、ここが
            # required=['args','kwargs'] に化けて全ツールが呼べなくなる。
            bad = []
            for tname, want in EXPECTED_TOOLS.items():
                t = tools.get(tname)
                if not t:
                    bad.append(f"{tname}: 未公開")
                    continue
                got = sorted((t.input_schema or {}).get("required", []))
                if got != sorted(want["required"]):
                    bad.append(f"{tname}: required={got} (期待 {sorted(want['required'])})")
            check("S-03", "各ツールの required 引数がシグネチャどおり",
                  not bad, "\n".join(bad) or f"全{len(EXPECTED_TOOLS)}ツール一致")

            # --- S-04 破壊的操作の申告 ---
            bad = []
            for tname, want in EXPECTED_TOOLS.items():
                ann = tools[tname].annotations if tname in tools else None
                ro = getattr(ann, "read_only_hint", None) if ann else None
                de = getattr(ann, "destructive_hint", None) if ann else None
                if ro != want["read_only"] or de != want["destructive"]:
                    bad.append(f"{tname}: read_only={ro}, destructive={de} "
                               f"(期待 {want['read_only']}, {want['destructive']})")
            destructive = [n for n, w in EXPECTED_TOOLS.items() if w["destructive"]]
            check("S-04", "read_only_hint / destructive_hint が正しく申告される",
                  not bad, "\n".join(bad) or f"destructive=True: {', '.join(destructive)}")

            # --- S-05 送信できないことの保証 ---
            # 下書きの作成は許す(create_draft / draft_reply)。下書きはメールボックスに
            # 置かれるだけで外へ出ない。禁じるのは実際に送出する経路だけ。
            # ツール名の検査は補助でしかない。本当の保証は「Mail.Send を要求していない」
            # ことなので、そちらも見る。権限が無ければ、仮にコードが壊れても送れない。
            import outlook_auth
            forbidden = [n for n in tools if any(
                k in n.lower() for k in ("send", "submit", "permanent")
            )]
            scopes = " ".join(outlook_auth.SCOPES).lower()
            send_scope = "mail.send" in scopes
            check("S-05", "送信の経路が存在しない(ツール名 + 要求スコープ)",
                  not forbidden and not send_scope,
                  f"送出系ツール: {forbidden or 'なし'} / 要求スコープ: {' '.join(outlook_auth.SCOPES)}")

            # --- S-06〜S-08 例外を投げず、案内文字列を返す ---
            # MCPのエラー応答ではなく通常のテキストで返ること = LLMが次の手を打てる。
            async def expect_text(test_id: str, label: str, tool: str, args: dict) -> None:
                res = await session.call_tool(tool, args)
                body = text_of(res)
                crashed = (
                    getattr(res, "is_error", False)
                    or "Error executing tool" in body
                    or "Traceback" in body
                    or "想定外のエラー" in body
                )
                check(test_id, label, bool(body) and not crashed,
                      f"{tool} -> {body.splitlines()[0][:110] if body else '(空)'}")

            await expect_text("S-06", "未設定/未ログインでも check_config が案内を返す",
                              "check_config", {})
            await expect_text("S-07", "未設定/未ログインでも search_messages が案内を返す",
                              "search_messages", {"query": "テスト"})
            await expect_text("S-08", "未知の短縮ID を案内文字列で拒否する",
                              "get_message", {"message_id": "#999999"})

            # --- S-09 一括操作の上限 ---
            over = ",".join(f"raw-id-{i}" for i in range(30))
            res = await session.call_tool("move_to_trash", {"message_ids": over})
            body = text_of(res)
            check("S-09", "25件を超える一括指定を拒否する",
                  "25 件まで" in body, f"move_to_trash(30件) -> {body[:110]}")

            # --- S-10 読み取り専用モード ---
            # サーバは起動時ではなく呼び出しごとに環境変数を見るが、
            # 子プロセスの環境は変えられないためここでは表示のみ確認する。
            res = await session.call_tool("check_config", {})
            body = text_of(res)
            check("S-10", "check_config が書き込み可否を報告する",
                  "書き込み:" in body,
                  "\n".join(ln for ln in body.splitlines() if "書き込み" in ln) or "(該当行なし)")

    passed = sum(1 for _, _, ok, _ in results if ok)
    print()
    print(f"結果: {passed}/{len(results)} パス")
    return 0 if passed == len(results) else 1


if __name__ == "__main__":
    print("outlook-mcp スモークテスト")
    print(f"実行時刻: {datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')}")
    print(f"Python:   {sys.version.split()[0]}")
    print(f"対象:     {BASE_DIR / 'outlook_server.py'}")
    print("-" * 72)
    sys.exit(asyncio.run(run()))
