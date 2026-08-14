# outlook-mcp

Hotmail / Outlook.com のメールボックスを、AIエージェントから自然言語で検索・整理できるようにする MCP サーバです。

```
あなた: 先月のAmazonからのメール、未読のやつを「買い物」フォルダにまとめて
Claude: (search_messages → list_folders → move_messages を自動で呼ぶ)
        7月のAmazonからの未読12件を「受信トレイ/買い物」へ移動しました。
```

Microsoft Graph API を使います。**メールの送信はできません。完全削除もできません。**

---

## できること / できないこと

| | |
|---|---|
| ✅ 検索 | 件名・本文・差出人・期間・未読で絞り込み |
| ✅ 読む | 本文の取得(HTMLメールは平文に変換) |
| ✅ 整理 | フォルダ移動、アーカイブ、既読/未読の切り替え |
| ✅ 一括整理 | 条件に一致するメールをまとめて移動・既読化(下見つき) |
| ✅ 棚の再編 | フォルダの作成・改名・移動・削除 |
| ✅ 自動振分 | Outlook側のルールを作る(MCPが起動していなくても効く) |
| ✅ 捨てる | ゴミ箱へ移動(**元に戻せます**) |
| ❌ 送信 | 実装していません。`Mail.Send` 権限を要求しません |
| ❌ 完全削除 | 実装していません。取り返しがつかないため |
| ❌ 添付の取得 | 未実装(添付の有無は 📎 で表示されます) |

要求する権限は **`Mail.ReadWrite`** と **`MailboxSettings.ReadWrite`**(振分ルール用)の2つだけです。

---

## ユースケース

実際に4万通・274フォルダのメールボックスで通した作業です。

### 受信トレイを空にする

差出人ごとに既存フォルダへ流し込みます。フォルダが無ければ作ります。

```
あなた: 受信トレイを差出人ごとに振り分けて空にして
Claude: 140件を確認しました。ぐるなび9件→06_Shop/Meal/ぐるなび、
        マネックス6件→02_Invest/マネックス証券、… 全140件を振り分けます。
```

### 何年も溜まったフォルダを再編する

`move_folder` は**メールを1通も動かさずに**棚ごと階層を変えます。数万通の再編が数十回のAPI呼び出しで終わります。

```
あなた: もう見ていない領域を 99_Archive にまとめて
Claude: Music / Communication / Game / Employment を 99_Archive の下へ移しました
        (約10,100通、メールは無変更)。
```

### 一日分のメールから有益な情報だけ拾う

```
あなた: 8月1日に届いたメールで有益なものをピックアップして
Claude: 76件のうち実質的なものは6件でした。
        後藤達也「Amazon/Apple 決算明暗の構図」— Amazonが設備投資を2200億ドルへ引き上げ…
        楽天銀行 8月優遇金利 0.38%、8月3日から +0.10% 改定予定 …
```

### 溜まる一方の購読を仕分ける

`create_rule` は Outlook サーバ側にルールを作るので、**MCPが起動していなくても24時間効きます**。

```
あなた: この差出人、今後は自動でこのフォルダに入れて
Claude: ルール「Meta広告の領収書」を作成しました(適用順 33)。
        今後届く分を 04_Business/Meta広告 へ移動 + 既読化します。
        既に届いている分は move_by_search で別途動かせます。
```

### 記録として保全する

削除ではなく格納だけ、既読にもしない、という指定ができます。未読件数の増加そのものが「継続的に送られてきている」記録になります。

```
あなた: この差出人は詐欺の疑いがあるので証拠として残したい
Claude: 格納のみ・既読化なしのルールを作りました。ヘッダはOutlook側に完全な形で残ります。
```

---

## 設計方針

**AIに受信箱を触らせる**以上、壊せる範囲を先に狭めておくべきだと考えました。

- **完全削除の手段を置かない。** ゴミ箱へ移すだけ。ツールとして存在しない操作は、どう指示されても起きません。
- **一度に触れるのは25件まで。** 誤爆したときの被害を有限にします。大量処理は専用ツールに分け、**下見を既定**にしました。
- **無条件の一括移動を拒否する。** 絞り込み条件のない `move_by_search` はエラーになります。
- **システムフォルダを守る。** 受信トレイ・迷惑メール等は改名・移動・削除できません。
- **空でないフォルダは `force` なしに消せない。** 退避(`move_folder`)で足りるなら、そちらを促します。
- **`OUTLOOK_READONLY=true` で書き込みを全面停止。** 読み取り専用サーバとして動かせます。
- **送信権限を要求しない。** 「AIが勝手にメールを出す」経路を原理的に作りません。
- **`destructive_hint` を正しく申告する。** 対応クライアントは `move_to_trash` を他と区別して扱えます。

---

## セットアップ

**必要なもの**: Python 3.10以上、Microsoftアカウント、Claude Code(または他のMCPクライアント)

作業は2つに分かれています。**人がやるのは1だけ**です。

### 1. Azure でアプリを登録する(人の手・初回のみ)

クライアントIDというGUIDを1個発行します。無料で、Azureのサブスクリプション契約は要りません。

ブラウザでのサインインと同意が絡むため、ここだけは自分の目で確認しながら進めてください。**あなたのメールボックスへのアクセス権を発行する操作**です。

→ **[docs/AZURE.md](docs/AZURE.md)**

### 2. 残り全部(Claude Code に任せる)

リポジトリを clone して Claude Code を起動し、こう言うだけです。

```
docs/SETUP-FOR-CLAUDE.md を読んでセットアップして
```

仮想環境の作成、依存関係、`.env`、MCP登録、疎通確認まで自動で進みます。途中1回だけ、**ログインのために手が止まります** — device code flow はブラウザでのサインインを伴うため、エージェントには完了できません。提示されたコマンドを自分で実行してください。

→ 手順書の中身: **[docs/SETUP-FOR-CLAUDE.md](docs/SETUP-FOR-CLAUDE.md)**

<details>
<summary>手作業で入れたい場合</summary>

```bash
# macOS / Linux
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
cp .env.example .env          # OUTLOOK_CLIENT_ID を書く
.venv/bin/python login.py     # ブラウザでコードを入力
claude mcp add outlook -- /abs/path/.venv/bin/python /abs/path/outlook_server.py
```

```powershell
# Windows (PowerShell)
py -3 -m venv .venv
.venv\Scripts\pip install -r requirements.txt
copy .env.example .env
.venv\Scripts\python.exe login.py
claude mcp add outlook -- "C:\abs\path\.venv\Scripts\python.exe" "C:\abs\path\outlook_server.py"
```

パスは**絶対パス**で。Windowsは `bin` ではなく `Scripts`、拡張子つきです。
つまずいたときの対応表は [docs/SETUP-FOR-CLAUDE.md](docs/SETUP-FOR-CLAUDE.md) にあります。

</details>

つながったら、まず `check_config` を呼ばせてみてください。

```
あなた: Outlookつながってる?
Claude: OK: 受信トレイ 3,412件(未読 87件)にアクセスできました。
        振分ルール: 読み書き可(12件設定済み)
```

---

## ツール

| ツール | 種別 | 説明 |
|---|---|---|
| `check_config` | 読取 | 設定・認証・接続の診断 |
| `list_folders` | 読取 | フォルダ一覧(件数・未読数つき) |
| `search_messages` | 読取 | 検索。キーワード / 差出人 / 期間 / 未読 / フォルダ |
| `get_message` | 読取 | 1通の本文と宛先を読む |
| `list_rules` | 読取 | 設定済みの振分ルールを一覧 |
| `create_folder` | 書込 | フォルダを作る |
| `rename_folder` | 書込 | フォルダを改名する(中身は動かない) |
| `move_folder` | 書込 | フォルダを別の親の下へ移す(中身ごと) |
| `move_messages` | 書込 | 指定フォルダへ移動(最大25件) |
| `move_by_search` | 書込 | 条件に一致するメールを一括移動(最大2,000件) |
| `mark_messages_read` | 書込 | 既読 / 未読の切り替え(最大25件) |
| `mark_read_by_search` | 書込 | 条件に一致するメールを一括既読化(最大25,000件) |
| `archive_messages` | 書込 | アーカイブへ移動 |
| `create_rule` | 書込 | 自動振分ルールを作る |
| `move_to_trash` | 破壊 | ゴミ箱へ移動(元に戻せる) |
| `delete_folder` | 破壊 | フォルダを削除(空でなければ `force` が要る) |
| `delete_rule` | 破壊 | 振分ルールを削除(メールは動かない) |

### 一括処理 — `move_by_search` / `mark_read_by_search`

1件ずつ扱うツールは1回25件です。数千通を動かすために、条件に一致するものをまとめて処理します。

**既定は下見(`dry_run=True`)で、何件動くかを数えるだけです。**

```
move_by_search(dest="99_Archive", folder="Music")
  → 元: Music / 走査 6,214件 → 該当 6,214件
    【下見のみ・まだ動かしていません】

move_by_search(dest="99_Archive", folder="Music", dry_run=False)
  → 6,214件を「99_Archive」へ移動しました。
```

内部では Graph の `/$batch` に20件ずつ束ねます。1通ずつ叩くと往復回数が現実的でないためです。**レスポンスは1件ずつ status を見ます** — バッチ全体を成否で判定すると、一部がスロットリングされただけで全件を再処理することになるためです。

- `move_by_search` は絞り込み条件を1つも指定しない呼び出しを**拒否**します(メールボックス全体を無条件に動かす事故を防ぐため)
- `mark_read_by_search` は居場所を変えないため条件なしを許し、上限も緩めてあります。ただし**既読/未読は「まだ見ていない」という情報そのもので、まとめて既読にすると復元できません**
- 差出人・件名の部分一致は手元で判定します(`$filter` が `contains()` を受け付けないため)。**アドレスと表示名の両方**を見るので、`ぐるなび` のような日本語の差出人名でも絞れます

> **棚ごと動かせるなら `move_folder` のほうが速いです。** メールを1通も動かさずに階層だけ変わります。

### 自動振分ルール — `create_rule`

Outlook のサーバ側に保存されるルールです。**このMCPが起動していなくても24時間効きます。**

```
create_rule(name="Meta広告の領収書", from_contains="facebookmail.com",
            subject_contains="領収書", move_to="04_Business/Meta広告", mark_read=True)
```

- 条件(`from_contains` / `subject_contains` / `body_contains`)は複数指定すると **AND**。値はカンマ区切りで複数渡せます
- **既に届いているメールには適用されません。** 過去分は `move_by_search` で別途動かします
- `to_trash` はゴミ箱へ移すだけで、完全削除ではありません
- 作る前に `list_rules` で既存を確認してください。Outlook の Web UI で作ったルールは条件が `fromAddresses` 形式で入っており、`create_rule` が使う `senderContains` とは別物です。同じ差出人に二重にルールを作ると、適用順の早い方が勝ちます

### 短縮ID

`search_messages` の各行は `#12` のような番号から始まります。Graph のメッセージIDは150文字前後あり、50件返すとそれだけで文脈を食い潰すためです。整理系ツールにはこの番号をそのまま渡します。

```
#12 2026-08-09 14:03 ●📎 Amazon.co.jp | ご注文の確認 | 受信トレイ
```

番号はサーバのプロセスが生きている間だけ有効で、**使い回されません**(再検索で `#3` の指す先が変わると事故になるため)。生のGraph IDも受け付けます。

> **メールを移動すると、その短縮IDは失効します。** Graph は移動時にメッセージIDを再発行するためです。
>
> ```
> move_messages("#1,#2", "領収書")   → 成功
> mark_messages_read("#1")           → エラー: 対象が見つかりません。
> ```
>
> 「移動してから既読にする」のような連続操作をするときは、**移動後に `search_messages` を引き直して
> 新しい番号を取り直してください**。エラーは案内文字列で返るため処理は止まりません。
>
> 既読化(`mark_messages_read` / `mark_read_by_search`)はメールを動かさないため、IDは失効しません。

---

## 既知の制約

**キーワード検索と厳密な新着順は両立しません。** Microsoft Graph は `$search` と `$filter` / `$orderby` を併用できない仕様です。このサーバは:

- キーワードや差出人の指定があるとき → `$search`(KQL)で検索し、**関連度順**で最大100件取得してから手元で日付順に並べ替える
- 指定がないとき → `$filter` + `$orderby` で**確実に新着順**

という切り替えをしています。前者で該当が100件を超える場合、古いものが取りこぼされる可能性があります。その旨は結果の末尾に表示されます。

**`search_messages` の `since` / `until` は UTC 基準です。** JST で「その日」を厳密に切りたい場合は、前後の日をまたいで取得し手元で絞り込んでください。

**フォルダ階層は3階層までしか列挙しません。** それより深いフォルダは `list_folders` に出ません。移動操作自体は深さに関係なく機能します。

**大量処理はスロットリングされることがあります。** Graph が `MailboxConcurrency limit` を返した分は失敗として報告されるので、同じ呼び出しを再実行すれば残りだけを拾えます。

---

## 開発 / テスト

```bash
# macOS / Linux
.venv/bin/pip install pytest
.venv/bin/pytest -q              # 単体テスト
.venv/bin/python smoke_test.py   # スモークテスト 10項目
```

```powershell
# Windows (PowerShell)
.venv\Scripts\pip install pytest
.venv\Scripts\pytest -q
.venv\Scripts\python.exe smoke_test.py
```

どちらも **Microsoft Graph に接続せず、メールボックスを変更しません**。認証情報も不要で、ログイン前の状態のまま実行できます。

- `test_outlook.py` — 関数単位。Graph呼び出しはスタブに差し替え、「どんなリクエストを組み立てたか」を検証します
- `smoke_test.py` — サーバを実際に stdio で起動し、MCPクライアントから見える外形(ツール一覧・入力スキーマ・`destructive_hint`・エラー時の応答)を検証します

テスト項目の一覧と実行エビデンスは **[docs/TEST.md](docs/TEST.md)** にあります。

実アカウントに対する結合テストも実施済みです。書き込み側は**既存のメールに触れず**、テスト用のフォルダとメッセージを自作して往復させ、最後に自分で作ったものだけを片付ける形で確認しています。未検証のまま残っている範囲も同ドキュメントに明記しています。

---

## ドキュメント

| | 読む人 | 内容 |
|---|---|---|
| このREADME | 人間 | 機能・ユースケース・設計方針・ツール一覧 |
| [docs/AZURE.md](docs/AZURE.md) | 人間 | Azureアプリ登録(唯一の手作業) |
| [docs/SETUP-FOR-CLAUDE.md](docs/SETUP-FOR-CLAUDE.md) | **エージェント** | セットアップ手順書。Claude Code に読ませる |
| [docs/TEST.md](docs/TEST.md) | 人間 | テスト項目とエビデンス |

---

## ライセンス

MIT
