# テスト仕様・実施結果

outlook-mcp のテスト項目と、その実行エビデンスです。

| | |
|---|---|
| 実施日 | 2026-08-14（初回実施 2026-08-10 / 別マシンへ移設後の再実行 / 同日 v0.2 でツール追加） |
| 結果 | **100項目 / 100パス、0失敗** |
| 内訳 | 単体テスト 71項目 + スモークテスト 10項目 + **結合テスト（実アカウント）読み取り9項目・書き込み10項目** |
| 環境 | Windows 11 Home 10.0.22631 / Python 3.13.1 / mcp 2.0.0 / msal 1.37.0 / httpx 0.28.1 |

> **v0.2 でツールが 9 → 17 に増えました（2026-08-14）。**
> 追加分（フォルダ改名/移動/削除・一括移動・一括既読化・振分ルール3種）の単体テストは
> 36項目を新規に書き、6節にまとめています。スモークテストも公開ツール17個に更新しました。
> 追加分は実アカウント（約4万通・276フォルダ）に対する**実運用でも通していますが**、
> 4〜5節のような手順化された結合テストとしては記録していません。範囲は6節に明記しています。

> **移設後の再実行について（2026-08-14）**
> 別マシンへコピーしたため、環境を変えて1〜2節を実行し直しました。**Windows 10 / Python 3.12.10 → Windows 11 / Python 3.13.1** に変わっていますが、結果は初回と同じ 45/45 パスです。あわせて、同日に**4節の結合テスト（実アカウント9項目）を新規に実施**しました。
> `.venv` は Python 3.12 に結びついていて移設先では起動できなかったため、3.13 で作り直しています(`requirements.txt` はバージョン下限のみの指定なので、依存の解決結果は [evidence/environment.txt](evidence/environment.txt) を参照)。
> 唯一、応答文が変わったのは S-07 です。移設先では `.env` に `OUTLOOK_CLIENT_ID` が入っているため、案内が「`.env` に値がありません」から「まだログインしていません」に進みました。どちらも「例外で落ちずに次の手を案内する」という判定基準を満たします。

## 再実行の方法

```bash
# macOS / Linux
.venv/bin/pip install -r requirements.txt pytest
.venv/bin/pytest -v              # 単体テスト(71項目)
.venv/bin/python smoke_test.py   # スモークテスト(10項目、終了コードで判定)
```

```powershell
# Windows (PowerShell)
.venv\Scripts\pip install -r requirements.txt pytest
.venv\Scripts\pytest -v
.venv\Scripts\python.exe smoke_test.py
```

1〜2節は **Microsoft Graph に接続せず、メールボックスを変更しません**。認証情報も不要で、ログイン前の状態でそのまま実行できます。4〜5節の結合テストは実アカウントへの接続とログイン済みのトークンが必要です（再現用スクリプトはリポジトリに含めていません）。

---

## 1. 単体テスト（[test_outlook.py](../test_outlook.py) / v0.1 分 35項目）

> v0.2 で追加した36項目は **6節** にまとめています。合計71項目です。

Graph呼び出しは記録用スタブ `FakeGraph` に差し替えています。「どんなリクエストを組み立てたか」を検証対象にしているため、ネットワークなしで送信内容の正しさを確認できます。

### 1-1. 検索条件の組み立て

Graph が `$search` と `$filter` を併用できない制約に、実装が正しく従っているかを見る区分です。

| ID | 検証内容 | 期待結果 | テスト関数 | 結果 |
|---|---|---|---|---|
| U-01 | 条件なしのとき `$filter` を作らない | `None` を返す | `test_filter_none_when_no_conditions` | PASS |
| U-02 | 未読＋期間の複合条件 | `and` で連結された式になる | `test_filter_combines_conditions` | PASS |
| U-03 | `until` が終端日を含む | `1/31まで` → `lt 2026-02-01T00:00:00Z` | `test_filter_until_is_inclusive` | PASS |
| U-04 | 不正な日付形式を弾く | `ValueError`（YYYY-MM-DD を案内） | `test_bad_date_is_rejected` | PASS |
| U-05 | KQL に全条件を畳み込む | `領収書 AND from:… AND received>=… AND received<=…` | `test_kql_merges_everything` | PASS |
| U-06 | KQL内部の `"` を除去する | 式全体を `"` で囲うため内側に残さない | `test_kql_strips_inner_quotes` | PASS |
| U-07 | 差出人だけの指定 | `from:amazon.co.jp` | `test_kql_from_only` | PASS |

### 1-2. 本文の整形

| ID | 検証内容 | 期待結果 | テスト関数 | 結果 |
|---|---|---|---|---|
| U-08 | HTMLメールの平文化 | `<script>` `<style>` の中身が本文に混ざらない | `test_html_to_text_drops_markup_and_script` | PASS |
| U-09 | 文字実体参照の復元 | `&amp;` → `&` | `test_html_to_text_decodes_entities` | PASS |
| U-10 | 長文の切り詰め | 省略した文字数を明示する／短文はそのまま | `test_truncate_marks_omission` | PASS |

### 1-3. フォルダ解決

日本語フォルダ名・階層・同名フォルダの扱いを見ます。

| ID | 検証内容 | 期待結果 | テスト関数 | 結果 |
|---|---|---|---|---|
| U-11 | 完全一致の名前 | 「領収書」→ 該当ID | `test_resolve_by_exact_name` | PASS |
| U-12 | フルパスでの指定 | 「アーカイブ/仕事」→ 該当ID | `test_resolve_by_full_path_disambiguates` | PASS |
| U-13 | 同名フォルダが複数 | 誤爆せず、候補パスを添えて `ValueError` | `test_resolve_ambiguous_name_raises_with_paths` | PASS |
| U-14 | well-known名へのフォールバック | 「ゴミ箱」→ `deleteditems` | `test_resolve_well_known_when_not_in_list` | PASS |
| U-15 | 存在しないフォルダ | 既存一覧を添えて `ValueError` | `test_resolve_unknown_raises` | PASS |

### 1-4. 短縮ID（`#N`）

**再検索で `#3` の指す先が変わると、意図しないメールを移動してしまう**——そこを固定する区分です。

| ID | 検証内容 | 期待結果 | テスト関数 | 結果 |
|---|---|---|---|---|
| U-16 | 同一メールには同じ番号、別メールには別番号 | 番号を使い回さない | `test_handles_are_stable_and_not_reused` | PASS |
| U-17 | 未知の番号 | 取り直しを案内して `ValueError` | `test_unknown_handle_raises` | PASS |
| U-18 | 生のGraph IDも受け付ける | そのまま通す | `test_raw_id_passes_through` | PASS |
| U-19 | 区切り文字の許容 | `"#1, #2"` もリストも同じ結果 | `test_parse_refs_accepts_separators` | PASS |
| U-20 | 空指定と上限超過 | どちらも `ValueError`（26件目で拒否） | `test_parse_refs_rejects_empty_and_oversized` | PASS |

### 1-5. 一覧表示

| ID | 検証内容 | 期待結果 | テスト関数 | 結果 |
|---|---|---|---|---|
| U-21 | 未読・添付の記号表示 | `●` `📎` が付き、`preview=False` では本文が出ない | `test_format_line_marks_unread_and_attachment` | PASS |
| U-22 | プレビュー指定時 | 本文冒頭が付く | `test_format_line_includes_preview_when_asked` | PASS |
| U-23 | 欠損フィールド | 落ちずに「(件名なし)」「(差出人不明)」 | `test_format_line_handles_missing_fields` | PASS |

### 1-6. 書き込みロック

| ID | 検証内容 | 期待結果 | テスト関数 | 結果 |
|---|---|---|---|---|
| U-24 | `OUTLOOK_READONLY=true` | `PermissionError` で書き込みを止める | `test_readonly_mode_blocks_writes` | PASS |
| U-25 | 既定（未設定） | 書き込みを許可する | `test_writes_allowed_by_default` | PASS |

### 1-7. ツール本体の振る舞い

| ID | 検証内容 | 期待結果 | テスト関数 | 結果 |
|---|---|---|---|---|
| U-26 | キーワード検索の経路 | `$search` を使い、`$filter`/`$orderby` を**付けない** | `test_search_with_query_uses_search_not_filter` | PASS |
| U-27 | 条件のみの検索の経路 | `$filter` + `$orderby=receivedDateTime desc` | `test_search_without_query_uses_filter_and_orderby` | PASS |
| U-28 | フォルダ指定 | エンドポイントが `/me/mailFolders/{id}/messages` になる | `test_search_scopes_to_folder` | PASS |
| U-29 | `$search`時の未読絞り込み | サーバ側でできない分を取得後に適用する | `test_search_applies_unread_filter_clientside_when_searching` | PASS |
| U-30 | 件数上限 | `limit=9999` でも `$top` は50に丸める | `test_search_caps_limit` | PASS |
| U-31 | 移動 | `POST /me/messages/{id}/move` に `destinationId` を送り件数を報告 | `test_move_messages_posts_move_and_counts` | PASS |
| U-32 | ゴミ箱への移動 | `destinationId=deleteditems`／「元に戻せます」と明示 | `test_move_to_trash_uses_wellknown_folder` | PASS |
| U-33 | 一括操作の部分失敗 | 成功分は巻き戻さず、失敗した短縮IDを列挙する | `test_partial_failure_is_reported_not_raised` | PASS |
| U-34 | 読み取り専用時のツール応答 | 例外送出ではなく「エラー: …」文字列を返す | `test_write_tools_return_error_string_in_readonly` | PASS |
| U-35 | 本文取得 | 件名・差出人・フォルダを整形し、HTMLを平文化 | `test_get_message_renders_html_body` | PASS |

---

## 2. スモークテスト（[smoke_test.py](../smoke_test.py) / 10項目）

サーバを**実際に stdio で起動**し、MCPクライアントから見た外形を検証します。単体テストでは捕まえられない「クライアント越しに壊れている」状態を検出する層です。

| ID | 検証内容 | 期待結果 | 結果 |
|---|---|---|---|
| S-01 | stdio ハンドシェイク | `server_info.name == "outlook"` | PASS |
| S-02 | ツールの公開 | **17個**、名前が定義どおり | PASS |
| S-03 | 入力スキーマ | 各ツールの `required` がシグネチャどおり | PASS |
| S-04 | 操作種別の申告 | `read_only_hint` / `destructive_hint` が正しい（destructive は `move_to_trash` / `delete_folder` / `delete_rule` の3つ） | PASS |
| S-05 | 危険なツールの不在 | send / reply / forward / draft / 完全削除 に該当する名前が存在しない | PASS |
| S-06 | 未設定/未ログイン時の `check_config` | 例外で落ちず、設定手順を案内する | PASS |
| S-07 | 未設定/未ログイン時の `search_messages` | 例外で落ちず、不足しているもの（`.env` の値、またはログイン）を案内する | PASS |
| S-08 | 未知の短縮ID | 「取り直してください」と案内する | PASS |
| S-09 | 一括操作の上限 | 30件指定を「25件まで」と拒否する | PASS |
| S-10 | 書き込み可否の報告 | `check_config` が現在の書き込み可否を出力する | PASS |

S-06〜S-09 は「LLMが次の手を打てる文字列で返ること」を判定基準にしています。MCPのエラー応答やトレースバックで返るとエージェントが停止してしまうためです。

---

## 3. 回帰テストとして固定した不具合

実装中に実際に踏んだ3件です。いずれも再発防止のテストを紐付けてあります。

| # | 不具合 | 症状 | 検出したテスト |
|---|---|---|---|
| 1 | エラーハンドリングのデコレータに `functools.wraps` がなかった | MCP SDK がシグネチャからスキーマを起こすため、全9ツールの入力が `required: ['args','kwargs']` に化けて**呼び出し不能**になっていた。単体テストは全パスしたまま | **S-03** |
| 2 | `$search` の値を二重に引用符で囲っていた | `""領収書" AND from:"x""` という壊れたKQLを送っていた | U-05 / U-06 / **U-26** |
| 3 | `$filter` で `contains(from/emailAddress/address, …)` を使っていた | Graph の messages は `contains()` 非対応。差出人の部分一致が常に400になる想定 | U-07 / **U-26** |
| 4 | **README がリダイレクトURIを「空のまま」と案内していた** | 個人アカウントでのログインが `invalid_request: The provided request must include a 'redirect_uri' input parameter.` で必ず失敗する | **I-01**（自動テストでは検出不能） |

1件目は「単体テストが全部緑でも、クライアントからは1つも呼べない」という状態でした。スモークテストの層を分けた理由がこれです。

4件目はコードではなく**手順書の誤り**で、コードを1行も動かさずに人を止めるたぐいのものです。「パブリック クライアント フローを許可する = はい」は必要条件ですが十分ではなく、個人アカウント側の認証サーバ（`login.live.com`）は実体としてのリダイレクトURIの登録を別途要求します。職場・学校アカウントだけなら空でも通るため見落とされていました。device code の**発行**は成功し、ブラウザで認証した瞬間に落ちるので、原因の切り分けもしにくい部類です。自動テストで捕まえる手段が無く、実アカウントで一度通すまで発見できませんでした（[README.md](../README.md) 修正済み）。

---

## 4. 結合テスト（実アカウント / 2026-08-14）

1〜2節はGraphに接続しません。本節は**実際の Hotmail アカウントに接続して**確認した記録です。**読み取り系ツールのみを実行し、メールボックスは一切変更していません。**

| | |
|---|---|
| 実施日 | 2026-08-14 |
| アカウント | 個人 Microsoft アカウント（Hotmail） |
| 規模 | 受信トレイ 143件（未読 136件）／ 全 264 フォルダ・最大3階層 |
| エビデンス | [evidence/integration.txt](evidence/integration.txt) |

> 稼働中のメールボックスなので件数は実行のたびに動きます（本節の実行中にも 141→143 件へ増えました）。エビデンスには**件名・差出人・フォルダ名の実データを記録していません**。このリポジトリは公開前提のため、返り値の「形」だけを残しています。

| ID | 検証内容 | 結果 | 備考 |
|---|---|---|---|
| I-01 | `login.py` の device code flow | PASS | 下記「踏んだ不具合」参照。**README の手順に誤りがあり、修正した** |
| I-02 | `check_config` の疎通 | PASS | 受信トレイ 143件（未読 136件）を取得 |
| I-03 | 日本語フォルダ名・階層の解決 | PASS | 264フォルダ。`親/12.英数字混じり/日本語の名前` のような3階層＋日本語＋記号を正しく解決 |
| I-04 | `$filter` + `$orderby` 経路（条件なし検索） | PASS | 新着順で返る |
| I-05 | KQL フリーテキスト検索 | PASS | **Graph が受理**。該当100件超の警告表示も動作 |
| I-06 | KQL `from:` 検索 | PASS | Graph が受理 |
| I-07 | KQL 全条件（query＋期間＋未読） | PASS | Graph が受理。未読の絞り込みはクライアント側で適用される |
| I-08 | フォルダ指定検索（日本語名） | PASS | 「受信トレイ」でスコープが効く |
| I-09 | `get_message` の本文取得 | PASS | 件名・差出人・宛先・フォルダ・状態を整形して返す |

### 本節で分かったこと

**`html_to_text()` は実運用ではほぼ発火しません。** サーバは `Prefer: outlook.body-content-type="text"` を送っており、Graph 側が平文化して返すためです。直近25通を調べたところ **25通すべて `contentType=text`** で、HTMLのまま来たものはありませんでした。この関数はGraphが平文化に応じない場合の保険として残っていますが、**実メールでの平文化の読みやすさは依然として未確認**です（U-08〜U-09 で関数単体の正しさは確認済み）。

なお本文中に `<mailto:...>` `<https://...>` という山括弧が現れますが、これはHTMLの残骸ではなく、Graph が平文化する際に採用するリンク表記です。

---

## 5. 書き込み系ツールの結合テスト（実アカウント / 2026-08-14）

**既存のメールには一切触れていません。** テスト用フォルダとテスト用メッセージを自分で作り、それだけを動かして、最後に自分で作ったものだけを片付けています。終了時点でメールボックスは実行前と同じ状態です（エビデンス末尾の「残存確認」で検証）。

| ID | 検証内容 | 期待結果 | 結果 |
|---|---|---|---|
| W-01 | `create_folder` | フォルダが実際に作られる | PASS |
| W-02 | 作成物を `search_messages` が拾う | 短縮IDが振られる | PASS |
| W-03 | `mark_messages_read(read=True)` | Graph上の `isRead` が `True` になる | PASS |
| W-04 | `mark_messages_read(read=False)` | 未読に戻る | PASS |
| W-05 | `archive_messages` | アーカイブへ移動 | PASS |
| W-06 | **移動後に古い短縮IDを使う** | 例外ではなくエラー文字列を返す | PASS |
| W-07 | `move_messages` | 指定フォルダへ戻る | PASS |
| W-08 | 存在しない移動先 | エラー文字列で拒否し、**メールを動かさない** | PASS |
| W-09 | `move_to_trash` | 「削除済みアイテム」へ移動 | PASS |
| W-10 | ゴミ箱からの復元 | 元のフォルダへ戻せる（完全削除ではない） | PASS |

エビデンス: [evidence/integration-write.txt](evidence/integration-write.txt)

W-08 と W-10 は、READMEが掲げる設計方針（「壊せる範囲を先に狭める」「ゴミ箱へ移すだけ」）が実際に成立していることの確認です。W-10 で移動先を誤ってもメールが失われないことを実測しました。

### 分かった制約 — 移動すると短縮IDが失効する

**Graph は移動時にメッセージIDを再発行します。** そのため `move_messages` / `archive_messages` / `move_to_trash` を実行した瞬間、**動かしたメールの短縮IDは指す先を失います**（W-06）。

```
move_messages("#1,#2", "領収書")   → 成功
mark_messages_read("#1")           → エラー: 対象が見つかりません。
```

例外にはならず案内文字列が返るためエージェントは停止しませんが、「移動してから既読にする」のような連続操作は**移動後に `search_messages` を引き直す必要があります**。

発見時点では、サーバの `instructions` にも README にもこの説明がなく、エージェントは一度失敗してから学ぶ状態でした。**対処済みです**:

- `outlook_server.py` の `instructions` に失効の説明を追加
- 移動を伴う3ツール（`move_messages` / `archive_messages` / `move_to_trash`）の docstring に1行ずつ追加。MCPクライアントから見えるツール説明に反映されることを stdio 経由で確認済み（`mark_messages_read` は移動しないため対象外）
- [README.md](../README.md) の「短縮ID」節にも追記

いずれも説明文の変更のみで、挙動は変えていません。

### 依然として未検証の項目

- トークンの自動更新、失効時の挙動（長期運用しないと踏めない）
- 実際のHTMLメールに対する平文化の読みやすさ（4節のとおり `html_to_text()` が発火せず）

> レート制限と大量操作については、6節で実運用の記録を追記しました。

---

## 6. v0.2 追加ツールの検証（2026-08-14）

ツールを 9 → 17 に増やしました。単体テストを36項目追加し、**実アカウント（約4万通・276フォルダ）で実運用も通しています**。ただし4〜5節のように手順化した結合テストとしては記録していないため、確認できた範囲と残っている範囲を分けて書きます。

### 追加した単体テスト（36項目）

| 対象 | 主な検証内容 |
|---|---|
| フォルダ管理 | システムフォルダの保護、`new_name` への `/` 混入拒否、自分の子孫への移動拒否、空でないフォルダの `force` なし削除拒否 |
| 一括処理 | `/$batch` が20件ずつに分割されること、**レスポンスを1件ずつ status で判定**すること、差出人の部分一致、上限、`dry_run` で何も起きないこと、条件なし `move_by_search` の拒否 |
| 振分ルール | 条件/動作の組み立て、フォルダ名→ID解決、条件・動作が空のルールの拒否、`move_to` と `to_trash` の同時指定の拒否、名前によるルール削除 |
| ルールの表示 | `fromAddresses` 形式の読み取り、未知の条件キーを落とさないこと、本当に条件が空のときだけ「全メールに一致」と表示すること |

### 実運用で確認できたこと

| 項目 | 実績 |
|---|---|
| フォルダの改名・移動 | 274フォルダを9トップレベルへ再編。**フォルダIDは改名・移動をまたいで維持される**ことを確認（既存の振分ルール31件が移動後も正しい移動先を指し続けた） |
| 一括移動 | 単一フォルダ345件（`@odata.nextLink` による複数ページ）、受信トレイ140件を差出人別に振り分け |
| 一括既読化 | **14,617件**（731バッチ）。所要 約20分 |
| レート制限 | 上記のうち**4件が `MailboxConcurrency limit` で失敗**。1件ずつ status を見る設計のため、成功分を巻き戻さずに再実行で残り4件だけを処理できた |
| 振分ルール | 既存31件の読み取り、8件の新規作成、1件の削除 |

`MailboxConcurrency limit` は 429 ではなくバッチ内の個別レスポンスとして返りました。**バッチ全体を成否で判定していたら14,613件を再処理することになっていた**箇所です。

### この節で残っている未検証

- `delete_folder` の `force=True`（中身ごと削除）— 空フォルダの削除のみ実施
- `create_rule` の `body_contains` / `enabled=False` / `sequence` 明示指定以外の分岐
- `mark_read_by_search(read=False)`（既読→未読への一括戻し）— 単体テストのみ
- 2,000件を超える単一 `move_by_search`（上限に達したときの継続呼び出し）

---

## 7. エビデンス

| ファイル | 内容 |
|---|---|
| [evidence/pytest.txt](evidence/pytest.txt) | `pytest -v` の全出力（71項目の個別結果） |
| [evidence/junit.xml](evidence/junit.xml) | 同左の機械可読形式（CI取り込み用） |
| [evidence/smoke.txt](evidence/smoke.txt) | `smoke_test.py` の全出力（10項目、実際の応答文つき） |
| [evidence/integration.txt](evidence/integration.txt) | 実アカウントに対する結合テスト・読み取り系の全出力（9項目。メール本文・件名・差出人は伏せ字） |
| [evidence/integration-write.txt](evidence/integration-write.txt) | 同・書き込み系の全出力（10項目。後片付けの結果まで含む） |
| [evidence/environment.txt](evidence/environment.txt) | OS / Pythonバージョン / `pip freeze` 全依存 |
