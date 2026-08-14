"""
純粋関数のテスト(ネットワークにもメールボックスにも触らない)
    .venv/bin/pytest -q
"""

from __future__ import annotations

from typing import Any

import pytest

import outlook_server as s


# ---------- $filter / KQL の組み立て ----------

def test_filter_none_when_no_conditions():
    assert s.build_odata_filter() is None


def test_filter_combines_conditions():
    f = s.build_odata_filter(since="2026-01-01", unread_only=True)
    assert "isRead eq false" in f
    assert "receivedDateTime ge 2026-01-01T00:00:00Z" in f
    assert f.count(" and ") == 1


def test_filter_until_is_inclusive():
    # 1/31 まで = 2/1 より前
    assert "receivedDateTime lt 2026-02-01T00:00:00Z" in s.build_odata_filter(until="2026-01-31")


def test_bad_date_is_rejected():
    with pytest.raises(ValueError, match="YYYY-MM-DD"):
        s.build_odata_filter(since="2026/01/01")


def test_kql_merges_everything():
    q = s.build_kql("領収書", from_address="billing@example.com", since="2026-01-01", until="2026-01-31")
    assert q == "領収書 AND from:billing@example.com AND received>=2026-01-01 AND received<=2026-01-31"


def test_kql_strips_inner_quotes():
    # 式全体を " で囲うので、内側に " が残るとKQLが壊れる
    assert '"' not in s.build_kql('請求"書', from_address='a"b@example.com')


def test_kql_from_only():
    assert s.build_kql(None, from_address="amazon.co.jp") == "from:amazon.co.jp"


# ---------- HTML → 平文 ----------

def test_html_to_text_drops_markup_and_script():
    html = "<html><head><style>p{color:red}</style></head><body><p>こんにちは</p><script>evil()</script><p>本文</p></body></html>"
    text = s.html_to_text(html)
    assert "evil" not in text
    assert "color:red" not in text
    assert "こんにちは" in text and "本文" in text


def test_html_to_text_decodes_entities():
    assert "A&B" in s.html_to_text("<div>A&amp;B</div>")


def test_truncate_marks_omission():
    out = s.truncate("あ" * 100, limit=10)
    assert out.startswith("あ" * 10)
    assert "90 文字を省略" in out
    assert s.truncate("短い", limit=10) == "短い"


# ---------- フォルダ解決 ----------

FOLDERS = [
    {"id": "id-inbox", "displayName": "受信トレイ", "path": "受信トレイ", "total": 3, "unread": 1},
    {"id": "id-rcpt", "displayName": "領収書", "path": "受信トレイ/領収書", "total": 0, "unread": 0},
    {"id": "id-work", "displayName": "仕事", "path": "受信トレイ/仕事", "total": 5, "unread": 0},
    {"id": "id-work2", "displayName": "仕事", "path": "アーカイブ/仕事", "total": 1, "unread": 0},
]


def test_resolve_by_exact_name():
    assert s.resolve_folder("領収書", FOLDERS) == "id-rcpt"


def test_resolve_by_full_path_disambiguates():
    assert s.resolve_folder("アーカイブ/仕事", FOLDERS) == "id-work2"


def test_resolve_ambiguous_name_raises_with_paths():
    with pytest.raises(ValueError, match="複数あります"):
        s.resolve_folder("仕事", FOLDERS)


def test_resolve_well_known_when_not_in_list():
    assert s.resolve_folder("ゴミ箱", []) == "deleteditems"


def test_resolve_unknown_raises():
    with pytest.raises(ValueError, match="見つかりません"):
        s.resolve_folder("存在しない棚", FOLDERS)


# ---------- 短縮ID ----------

def test_handles_are_stable_and_not_reused():
    h1 = s.register_handle("AAAA")
    h2 = s.register_handle("BBBB")
    assert h1 != h2
    assert s.register_handle("AAAA") == h1  # 再検索でも同じメールは同じ番号
    assert s.resolve_message_id(h1) == "AAAA"


def test_unknown_handle_raises():
    with pytest.raises(ValueError, match="未知"):
        s.resolve_message_id("#999999")


def test_raw_id_passes_through():
    assert s.resolve_message_id("AQMkAGxyz") == "AQMkAGxyz"


def test_parse_refs_accepts_separators():
    a, b = s.register_handle("M1"), s.register_handle("M2")
    assert s.parse_refs(f"{a}, {b}") == ["M1", "M2"]
    assert s.parse_refs([a, b]) == ["M1", "M2"]


def test_parse_refs_rejects_empty_and_oversized():
    with pytest.raises(ValueError, match="1件も指定されていません"):
        s.parse_refs("")
    too_many = ",".join(s.register_handle(f"X{i}") for i in range(s.MAX_IDS_PER_CALL + 1))
    with pytest.raises(ValueError, match=f"{s.MAX_IDS_PER_CALL} 件まで"):
        s.parse_refs(too_many)


# ---------- 表示 ----------

MSG = {
    "id": "M-fmt",
    "subject": "ご注文の確認",
    "from": {"emailAddress": {"name": "Amazon.co.jp", "address": "auto@amazon.co.jp"}},
    "receivedDateTime": "2026-08-09T05:03:00Z",
    "isRead": False,
    "hasAttachments": True,
    "bodyPreview": "この度はご注文ありがとうございます。",
}


def test_format_line_marks_unread_and_attachment():
    line = s.format_line(MSG, "#7", "受信トレイ", preview=False)
    assert line.startswith("#7 ")
    assert "●" in line and "📎" in line
    assert "Amazon.co.jp | ご注文の確認 | 受信トレイ" in line
    assert "ご注文ありがとう" not in line  # preview=False なので出ない


def test_format_line_includes_preview_when_asked():
    assert "ご注文ありがとう" in s.format_line(MSG, "#7", None, preview=True)


def test_format_line_handles_missing_fields():
    line = s.format_line({"id": "x"}, "#1", None, preview=True)
    assert "(件名なし)" in line and "(差出人不明)" in line


# ---------- 書き込みロック ----------

def test_readonly_mode_blocks_writes(monkeypatch):
    monkeypatch.setenv("OUTLOOK_READONLY", "true")
    assert s.readonly_mode() is True
    with pytest.raises(PermissionError):
        s.ensure_writable()


def test_writes_allowed_by_default(monkeypatch):
    monkeypatch.delenv("OUTLOOK_READONLY", raising=False)
    assert s.readonly_mode() is False
    s.ensure_writable()  # 例外が出なければOK


# ---------- ツール本体(Graph呼び出しは差し替える) ----------

class FakeGraph(list):
    """呼び出しを記録するスタブ。responses[部分パス] = 返すJSON。"""

    def __init__(self) -> None:
        super().__init__()
        self.responses: dict[str, Any] = {}
        self.fail_on: str | None = None

    def __call__(self, method: str, path: str, **kwargs: Any) -> dict:
        self.append((method, path, kwargs))
        if self.fail_on and self.fail_on in path:
            raise s.GraphError("対象が見つかりません。")
        for key, value in self.responses.items():
            if key in path:
                return value
        return {}


@pytest.fixture
def calls(monkeypatch) -> FakeGraph:
    """s.graph を記録用スタブに差し替える。ネットワークには出ない。"""
    fake = FakeGraph()
    monkeypatch.setattr(s, "graph", fake)
    monkeypatch.setattr(s, "fetch_folders", lambda force=False: FOLDERS)
    monkeypatch.setattr(s, "folder_name_map", lambda: {f["id"]: f["path"] for f in FOLDERS})
    monkeypatch.delenv("OUTLOOK_READONLY", raising=False)
    return fake


def test_search_with_query_uses_search_not_filter(calls):
    calls.responses["/messages"] = {"value": [MSG]}
    out = s.search_messages(query="領収書", from_address="amazon.co.jp", limit=5)
    _, path, kwargs = calls[0]
    params = kwargs["params"]
    assert path == "/me/messages"
    assert params["$search"] == '"領収書 AND from:amazon.co.jp"'
    assert "$filter" not in params and "$orderby" not in params  # 併用するとGraphが400を返す
    assert "ご注文の確認" in out


def test_search_without_query_uses_filter_and_orderby(calls):
    calls.responses["/messages"] = {"value": [MSG]}
    s.search_messages(unread_only=True, since="2026-08-01", limit=5)
    params = calls[0][2]["params"]
    assert "$search" not in params
    assert params["$filter"] == "isRead eq false and receivedDateTime ge 2026-08-01T00:00:00Z"
    assert params["$orderby"] == "receivedDateTime desc"


def test_search_scopes_to_folder(calls):
    calls.responses["/messages"] = {"value": []}
    s.search_messages(folder="領収書")
    assert calls[0][1] == "/me/mailFolders/id-rcpt/messages"


def test_search_applies_unread_filter_clientside_when_searching(calls):
    read_msg = dict(MSG, id="M-read", isRead=True)
    calls.responses["/messages"] = {"value": [read_msg]}
    out = s.search_messages(query="確認", unread_only=True)
    assert "一致するメールはありません" in out


def test_search_caps_limit(calls):
    calls.responses["/messages"] = {"value": []}
    s.search_messages(limit=9999)
    assert calls[0][2]["params"]["$top"] == s.MAX_RESULTS


def test_move_messages_posts_move_and_counts(calls):
    h = s.register_handle("MOVE-1")
    out = s.move_messages(h, "領収書")
    method, path, kwargs = calls[-1]
    assert (method, path) == ("POST", "/me/messages/MOVE-1/move")
    assert kwargs["json"] == {"destinationId": "id-rcpt"}
    assert "1件を「領収書」へ移動" in out


def test_move_to_trash_uses_wellknown_folder(calls):
    out = s.move_to_trash(s.register_handle("TRASH-1"))
    assert calls[-1][2]["json"] == {"destinationId": "deleteditems"}
    assert "ゴミ箱" in out and "元に戻せます" in out


def test_partial_failure_is_reported_not_raised(calls):
    ok_h, bad_h = s.register_handle("OK-1"), s.register_handle("BAD-1")
    calls.fail_on = "BAD-1"
    out = s.mark_messages_read(f"{ok_h},{bad_h}")
    assert "1件を既読にしました" in out  # 成功したぶんは巻き戻さない
    assert bad_h in out and "失敗" in out


def test_write_tools_return_error_string_in_readonly(monkeypatch):
    monkeypatch.setenv("OUTLOOK_READONLY", "true")
    out = s.move_messages(s.register_handle("RO-1"), "領収書")
    assert out.startswith("エラー:") and "OUTLOOK_READONLY" in out


# ---------- 下書き(送信はしない) ----------

def test_send_scope_is_never_requested():
    """このプロジェクトの中核となる保証。壊れたら送信できてしまう。"""
    import outlook_auth
    joined = " ".join(outlook_auth.SCOPES).lower()
    assert "mail.send" not in joined


def test_create_draft_posts_to_messages_not_send(calls):
    out = s.create_draft(to="a@example.com, b@example.com", subject="件名", body="本文")
    method, path, kwargs = calls[-1]
    assert (method, path) == ("POST", "/me/messages")  # /sendMail ではない
    body = kwargs["json"]
    assert [r["emailAddress"]["address"] for r in body["toRecipients"]] == ["a@example.com", "b@example.com"]
    assert body["subject"] == "件名"
    assert body["body"] == {"contentType": "Text", "content": "本文"}
    assert "ccRecipients" not in body
    assert "送信はしていません" in out


def test_create_draft_includes_cc_when_given(calls):
    s.create_draft(to="a@example.com", subject="件名", body="本文", cc="c@example.com")
    assert calls[-1][2]["json"]["ccRecipients"] == [{"emailAddress": {"address": "c@example.com"}}]


def test_create_draft_rejects_empty_recipient_and_subject(calls):
    assert "宛先が空です" in s.create_draft(to=" , ", subject="件名", body="本文")
    assert "件名が空です" in s.create_draft(to="a@example.com", subject="  ", body="本文")
    assert not any(m == "POST" for m, _, _ in calls)  # 何も作っていない


def test_draft_reply_uses_create_reply_endpoint(calls):
    h = s.register_handle("REPLY-1")
    out = s.draft_reply(h, body="承知しました")
    method, path, kwargs = calls[-1]
    assert (method, path) == ("POST", "/me/messages/REPLY-1/createReply")
    assert kwargs["json"] == {"comment": "承知しました"}
    assert "送信はしていません" in out


def test_draft_reply_all_uses_reply_all_endpoint(calls):
    s.draft_reply(s.register_handle("REPLY-2"), body="はい", reply_all=True)
    assert calls[-1][1].endswith("/createReplyAll")


def test_draft_tools_blocked_in_readonly(monkeypatch):
    monkeypatch.setenv("OUTLOOK_READONLY", "true")
    assert s.create_draft(to="a@example.com", subject="x", body="y").startswith("エラー:")
    assert s.draft_reply("#1", body="y").startswith("エラー:")


# ---------- フォルダ管理 ----------

def test_system_folder_is_protected():
    inbox = {"id": "id-inbox", "displayName": "受信トレイ", "path": "受信トレイ", "total": 3, "unread": 1}
    with pytest.raises(ValueError, match="システムフォルダ"):
        s.ensure_not_system(inbox)
    s.ensure_not_system(FOLDERS[1])  # 普通のフォルダは通る


def test_rename_folder_patches_display_name(calls):
    out = s.rename_folder("領収書", "Receipts")
    method, path, kwargs = calls[-1]
    assert (method, path) == ("PATCH", "/me/mailFolders/id-rcpt")
    assert kwargs["json"] == {"displayName": "Receipts"}
    assert "改名しました" in out


def test_rename_folder_rejects_path_separator(calls):
    out = s.rename_folder("領収書", "親/子")
    assert out.startswith("エラー:") and "「/」は使えません" in out


def test_rename_folder_refuses_system_folder(calls):
    assert "システムフォルダ" in s.rename_folder("受信トレイ", "べつの名前")


def test_move_folder_to_root_uses_msgfolderroot(calls):
    s.move_folder("領収書")
    assert calls[-1][2]["json"] == {"destinationId": "msgfolderroot"}


def test_move_folder_rejects_moving_into_own_descendant(calls, monkeypatch):
    # 受信トレイ/仕事 を、その子である 受信トレイ/仕事/下請け の下へは動かせない
    nested = FOLDERS + [
        {"id": "id-deep", "displayName": "下請け", "path": "受信トレイ/仕事/下請け", "total": 0, "unread": 0}
    ]
    monkeypatch.setattr(s, "fetch_folders", lambda force=False: nested)
    out = s.move_folder("受信トレイ/仕事", "下請け")
    assert out.startswith("エラー:") and "自分の子孫" in out
    assert not any("/move" in p for _, p, _ in calls)


def test_move_folder_rejects_moving_into_itself(calls):
    out = s.move_folder("領収書", "領収書")
    assert out.startswith("エラー:") and "自分自身" in out


def test_delete_folder_refuses_nonempty_without_force(calls):
    out = s.delete_folder("受信トレイ/仕事")  # total=5
    assert out.startswith("エラー:")
    assert "force=True" in out
    assert not any(m == "DELETE" for m, _, _ in calls)  # 何も消していない


def test_delete_folder_with_force_calls_delete(calls):
    out = s.delete_folder("受信トレイ/仕事", force=True)
    assert calls[-1][0] == "DELETE"
    assert "削除しました" in out


def test_delete_empty_folder_needs_no_force(calls):
    s.delete_folder("領収書")  # total=0
    assert calls[-1][0] == "DELETE"


# ---------- 一括移動 ----------

def test_batch_move_splits_into_chunks_of_20(calls):
    ids = [f"M{i}" for i in range(45)]
    calls.responses["/$batch"] = {"responses": [{"id": str(n), "status": 200} for n in range(20)]}
    s.batch_move(ids, "id-rcpt")
    batches = [c for c in calls if c[1] == "/$batch"]
    assert len(batches) == 3  # 20 + 20 + 5
    assert len(batches[0][2]["json"]["requests"]) == 20
    assert len(batches[-1][2]["json"]["requests"]) == 5


def test_batch_move_counts_per_item_status(calls):
    calls.responses["/$batch"] = {
        "responses": [
            {"id": "0", "status": 200},
            {"id": "1", "status": 404, "body": {"error": {"message": "見つかりません"}}},
            {"id": "2", "status": 200},
        ]
    }
    ok, failed = s.batch_move(["A", "B", "C"], "id-rcpt")
    assert ok == 2
    assert len(failed) == 1 and "見つかりません" in failed[0]


def test_collect_filters_by_sender_substring(calls):
    calls.responses["/messages"] = {
        "value": [
            {"id": "A", "subject": "x", "from": {"emailAddress": {"address": "a@blockdag.network", "name": "BLOCKDAG"}}},
            {"id": "B", "subject": "x", "from": {"emailAddress": {"address": "b@example.com", "name": "Other"}}},
        ]
    }
    ids, scanned = s.collect_message_ids("/me/messages", None, "blockdag", None, 100)
    assert ids == ["A"] and scanned == 2


def test_collect_respects_cap(calls):
    calls.responses["/messages"] = {
        "value": [{"id": f"M{i}", "subject": "s", "from": {}} for i in range(10)]
    }
    ids, _ = s.collect_message_ids("/me/messages", None, None, None, 3)
    assert len(ids) == 3


def test_move_by_search_requires_a_condition(calls):
    out = s.move_by_search(dest="領収書")
    assert out.startswith("エラー:") and "絞り込み条件がありません" in out


def test_move_by_search_dry_run_moves_nothing(calls):
    calls.responses["/messages"] = {
        "value": [{"id": "A", "subject": "s", "from": {"emailAddress": {"address": "x@spam.com"}}}]
    }
    out = s.move_by_search(dest="領収書", from_address="spam.com")
    assert "まだ動かしていません" in out
    assert not any(p == "/$batch" for _, p, _ in calls)


def test_move_by_search_executes_when_not_dry_run(calls):
    calls.responses["/messages"] = {
        "value": [{"id": "A", "subject": "s", "from": {"emailAddress": {"address": "x@spam.com"}}}]
    }
    calls.responses["/$batch"] = {"responses": [{"id": "0", "status": 200}]}
    out = s.move_by_search(dest="領収書", from_address="spam.com", dry_run=False)
    assert any(p == "/$batch" for _, p, _ in calls)
    assert "1件を「領収書」へ移動しました" in out


def test_batch_mark_read_patches_is_read(calls):
    calls.responses["/$batch"] = {"responses": [{"id": "0", "status": 200}]}
    ok, failed = s.batch_mark_read(["A"], read=True)
    req = calls[-1][2]["json"]["requests"][0]
    assert req["method"] == "PATCH"
    assert req["url"] == "/me/messages/A"
    assert req["body"] == {"isRead": True}
    assert ok == 1 and not failed


def test_mark_read_by_search_filters_to_unread_only(calls):
    calls.responses["/messages"] = {"value": []}
    s.mark_read_by_search(folder="領収書")
    # 既読にするのだから、拾うのは未読だけでよい
    assert "isRead eq false" in calls[0][2]["params"]["$filter"]


def test_mark_read_by_search_unmark_targets_read_mail(calls):
    calls.responses["/messages"] = {"value": []}
    s.mark_read_by_search(folder="領収書", read=False)
    assert "isRead eq true" in calls[0][2]["params"]["$filter"]


def test_mark_read_by_search_dry_run_changes_nothing(calls):
    calls.responses["/messages"] = {"value": [{"id": "A", "subject": "s", "from": {}}]}
    out = s.mark_read_by_search(folder="領収書")
    assert "まだ変更していません" in out
    assert "復元できません" in out  # 不可逆であることを告げる
    assert not any(p == "/$batch" for _, p, _ in calls)


def test_mark_read_by_search_executes_when_not_dry_run(calls):
    calls.responses["/messages"] = {"value": [{"id": "A", "subject": "s", "from": {}}]}
    calls.responses["/$batch"] = {"responses": [{"id": "0", "status": 200}]}
    out = s.mark_read_by_search(folder="領収書", dry_run=False)
    assert "1件を既読にしました" in out


def test_mark_read_by_search_allows_whole_mailbox(calls):
    # move と違い、既読化は居場所を変えないので条件なしでも許す
    calls.responses["/messages"] = {"value": []}
    out = s.mark_read_by_search()
    assert not out.startswith("エラー:")
    assert calls[0][1] == "/me/messages"


def test_move_by_search_rejects_same_source_and_dest(calls):
    out = s.move_by_search(dest="領収書", folder="領収書")
    assert out.startswith("エラー:") and "同じフォルダ" in out


# ---------- 振分ルール ----------

def test_create_rule_builds_conditions_and_actions(calls):
    calls.responses["messageRules"] = {"value": [{"sequence": 4}]}
    out = s.create_rule(
        name="BLOCKDAG を捨てる", from_contains="blockdag.network, blockdag.io", to_trash=True
    )
    post = [c for c in calls if c[0] == "POST" and "messageRules" in c[1]][-1]
    body = post[2]["json"]
    assert body["conditions"] == {"senderContains": ["blockdag.network", "blockdag.io"]}
    assert body["actions"]["delete"] is True
    assert body["actions"]["stopProcessingRules"] is True
    assert body["sequence"] == 5  # 既存の最大 +1
    assert "作成しました" in out


def test_create_rule_resolves_destination_folder(calls):
    calls.responses["messageRules"] = {"value": []}
    s.create_rule(name="領収書へ", from_contains="billing@", move_to="領収書", mark_read=True)
    body = [c for c in calls if c[0] == "POST" and "messageRules" in c[1]][-1][2]["json"]
    assert body["actions"]["moveToFolder"] == "id-rcpt"
    assert body["actions"]["markAsRead"] is True
    assert body["sequence"] == 1  # ルールが1つも無ければ 1 から


def test_create_rule_requires_condition_and_action(calls):
    assert "条件が1つもありません" in s.create_rule(name="空", move_to="領収書")
    assert "動作が1つもありません" in s.create_rule(name="空", from_contains="a@b.com")


def test_create_rule_rejects_move_and_trash_together(calls):
    out = s.create_rule(name="両方", from_contains="a@b.com", move_to="領収書", to_trash=True)
    assert out.startswith("エラー:") and "同時に指定できません" in out


def test_list_rules_renders_conditions(calls):
    calls.responses["messageRules"] = {
        "value": [{
            "id": "r1", "displayName": "広告", "sequence": 1, "isEnabled": True,
            "conditions": {"senderContains": ["ads.example"]},
            "actions": {"moveToFolder": "id-rcpt", "markAsRead": True},
        }]
    }
    out = s.list_rules()
    assert "広告" in out and "ads.example" in out
    assert "受信トレイ/領収書" in out  # フォルダIDが名前に解決されている
    assert "既読化" in out


def test_describe_conditions_reads_address_form():
    # Web UI 製のルールは fromAddresses 形式。これを読めないと
    # 「条件なし = 全メール一致」と誤読させる。
    conds = {"fromAddresses": [{"emailAddress": {"name": "afb", "address": "info@afi-b.com"}}]}
    describe = s.describe_conditions(conds)
    assert "info@afi-b.com" in describe
    assert "なし" not in describe


def test_describe_conditions_surfaces_unknown_keys():
    out = s.describe_conditions({"someFutureCondition": ["x"]})
    assert "someFutureCondition" in out  # 黙って落とさない


def test_describe_conditions_marks_truly_empty_as_catch_all():
    assert "全メールに一致" in s.describe_conditions({})


def test_describe_actions_flags_missing_folder():
    out = s.describe_actions({"moveToFolder": "id-gone"}, {})
    assert "存在しない" in out


def test_describe_actions_renders_forwarding():
    out = s.describe_actions(
        {"forwardTo": [{"emailAddress": {"address": "x@example.com"}}]}, {}
    )
    assert "転送" in out and "x@example.com" in out


def test_delete_rule_matches_by_name(calls):
    calls.responses["messageRules"] = {
        "value": [{"id": "r-abc", "displayName": "広告", "sequence": 1}]
    }
    out = s.delete_rule("広告")
    method, path, _ = calls[-1]
    assert method == "DELETE" and path.endswith("/r-abc")
    assert "削除しました" in out


def test_delete_rule_unknown_name_lists_existing(calls):
    calls.responses["messageRules"] = {"value": [{"id": "r1", "displayName": "広告"}]}
    out = s.delete_rule("存在しない")
    assert out.startswith("エラー:") and "広告" in out


def test_get_message_renders_html_body(calls):
    calls.responses["/me/messages/"] = {
        "subject": "請求書",
        "from": {"emailAddress": {"name": "経理", "address": "keiri@example.com"}},
        "toRecipients": [{"emailAddress": {"address": "me@example.com"}}],
        "receivedDateTime": "2026-08-09T05:03:00Z",
        "isRead": True,
        "parentFolderId": "id-inbox",
        "body": {"contentType": "html", "content": "<p>金額は<b>1,000円</b>です</p><script>x()</script>"},
    }
    out = s.get_message(s.register_handle("GET-1"))
    assert "件名: 請求書" in out
    assert "経理 <keiri@example.com>" in out
    assert "フォルダ: 受信トレイ" in out
    assert "1,000円" in out
    assert "<p>" not in out and "x()" not in out
