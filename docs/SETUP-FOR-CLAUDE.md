# SETUP-FOR-CLAUDE — エージェント向けセットアップ手順書

このファイルは **Claude Code(または同等のコーディングエージェント)が読んで実行する**ための手順書です。人間向けの説明は [../README.md](../README.md) と [AZURE.md](AZURE.md) にあります。

利用者はこう言うだけで済みます:

```
docs/SETUP-FOR-CLAUDE.md を読んでセットアップして
```

---

## エージェントへの前提と禁止事項

**先に読むこと。以下は手順より優先されます。**

1. **`OUTLOOK_CLIENT_ID` を推測・捏造しない。** 利用者がAzureで発行したGUIDです。手元になければ利用者に尋ね、答えが得られるまで `.env` を書かないこと。プレースホルダのまま先に進むと、後段の失敗原因が分かりにくくなります。

2. **`login.py` は利用者が自分で実行する。** device code flow はブラウザでのサインインを伴い、エージェントには完了できません。**あなたが `login.py` を実行してはいけません**(応答待ちでハングします)。手順4で必ず手を止め、コマンドを提示して待ってください。

3. **デバイスコード・認証コード・トークンを利用者に尋ねない。** それらを受け取っても何もできませんし、受け取るべきでもありません。

4. **`.env` と `token_cache.json` を git に含めない、内容を出力しない。** `token_cache.json` はメールボックスの鍵に相当します。`.gitignore` 済みですが、`git add -A` の前に `git status` で確認してください。

5. **メールボックスを操作しない。** このファイルの範囲はセットアップだけです。疎通確認は `check_config` のみを使い、利用者が明示的に頼むまで検索・移動・削除は行わないこと。

---

## 前提条件の確認

次が揃っていなければ、揃うまで進めないでください。

| 条件 | 確認方法 | 揃っていないとき |
|---|---|---|
| Python 3.10 以上 | `python --version` / `py -3 --version` | 利用者に導入を依頼 |
| Azure アプリ登録済み | 利用者に `OUTLOOK_CLIENT_ID` を尋ねる | [AZURE.md](AZURE.md) を案内して中断 |
| リポジトリ直下にいる | `outlook_server.py` が存在するか | 正しいディレクトリへ移動 |

---

## 手順1 — OS を判定してパスを決める

以降のコマンドは OS で変わります。**最初に判定し、それ以降は一貫して同じ形式を使ってください。**

| | 仮想環境の python | pip |
|---|---|---|
| macOS / Linux | `.venv/bin/python` | `.venv/bin/pip` |
| Windows | `.venv\Scripts\python.exe` | `.venv\Scripts\pip.exe` |

Windows は `bin` ではなく `Scripts`、拡張子 `.exe` つきです。ここを取り違えるとこの後すべて失敗します。

---

## 手順2 — 仮想環境と依存関係

```bash
# macOS / Linux
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

```powershell
# Windows
py -3 -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

`.venv` が既にあれば作り直さず、`pip install` だけ実行してください。

---

## 手順3 — `.env` を作る

`.env.example` をコピーし、`OUTLOOK_CLIENT_ID` に利用者から受け取ったGUIDを書きます。

```ini
OUTLOOK_CLIENT_ID=<利用者から受け取ったGUID>
OUTLOOK_TENANT=common
OUTLOOK_READONLY=false
```

- `OUTLOOK_TENANT` — 既定は `common`。利用者が Hotmail / Outlook.com の**個人アカウントだけ**を使うと明言した場合は `consumers` を推奨してください。同じメールアドレスが「個人」と「職場」の2つの人格を持つ場合の取り違えを防げます(取り違えると Graph が `MailboxNotEnabledForRESTAPI` を返します)。
- `OUTLOOK_READONLY` — 読み取り専用で試したい場合は `true`。書き込み系ツールが全て無効になります。迷ったら利用者に確認してください。

---

## 手順4 — ここで手を止める(ログインは利用者が行う)

**あなたはこのコマンドを実行しません。** 利用者に提示して、完了の報告を待ってください。

```bash
# macOS / Linux
.venv/bin/python login.py
```

```powershell
# Windows
.venv\Scripts\python.exe login.py
```

利用者に次のことを伝えてください。

- 表示されたURLをブラウザで開き、コードを入力してMicrosoftアカウントでサインインすること
- **パスワードをこのツールに渡すことはない**こと
- 「通常は表示されないページにアクセスしました」という警告は**正常**であること。device code flow がフィッシングに悪用されてきたため、Microsoft は全員に警告を出します。**同意画面のアプリ名が自分でAzureに登録した名前(例 `outlook-mcp`)であることだけは確認する**よう伝えてください
- アカウント選択画面が出たら「**個人のアカウント**」を選ぶこと(メールボックスがあるのは個人の方です)

成功すると `token_cache.json` が生成されます。ファイルの存在で確認できますが、**中身は開かないでください。**

---

## 手順5 — MCP クライアントに登録する

**絶対パスで指定してください。** 相対パスだと別のディレクトリから起動したときに動きません。

```bash
# macOS / Linux
claude mcp add outlook -- /abs/path/outlook-mcp/.venv/bin/python /abs/path/outlook-mcp/outlook_server.py
```

```powershell
# Windows(パスに空白や日本語が含まれる場合に備え、必ず二重引用符で囲む)
claude mcp add outlook -- "C:\abs\path\outlook-mcp\.venv\Scripts\python.exe" "C:\abs\path\outlook-mcp\outlook_server.py"
```

スコープの判断:

- 既定は **local**(そのプロジェクト内でだけ有効)
- どのディレクトリからでも使いたいなら `-s user` を付ける
- 登録先を変えるときは先に `claude mcp remove outlook`

どちらにするか利用者に確認してください。メールボックス整理は特定のプロジェクトに紐づかない作業なので、`-s user` を選ぶ利用者が多いはずです。

登録後 `claude mcp list` で `connected` を確認します。

---

## 手順6 — 疎通確認

**MCPサーバを認識させるには再接続が必要です。** 利用者に `/mcp` から接続するよう伝えてください。

その後 `check_config` を1回だけ呼びます。期待する出力:

```
設定ファイル: /abs/path/outlook-mcp/.env
  存在: あり
書き込み: 有効
ログイン中: someone@example.com
OK: 受信トレイ 3,412件(未読 87件)にアクセスできました。
振分ルール: 読み書き可(12件設定済み)
```

**最後の2行が両方 OK であることを確認してください。** 片方だけ通ることがあります(下の対応表を参照)。

---

## つまずいたときの対応表

`check_config` の出力から原因を特定できます。**推測で設定を書き換える前に、必ず `check_config` を呼んでください。**

| 症状 | 原因 | 対応 |
|---|---|---|
| `NG: 未ログインです` | `login.py` 未実行 | 手順4に戻る |
| `振分ルール: 使えません` | Azureで `MailboxSettings.ReadWrite` 未追加、または権限追加後に同意し直していない | [AZURE.md](AZURE.md) 手順8を確認 → 利用者に `python login.py` を再実行してもらう |
| `NG: メールボックスに届きません` + `MailboxNotEnabledForRESTAPI` | ログイン時に「職場または学校アカウント」を選んだ | `.env` の `OUTLOOK_TENANT=consumers` に変更 → `login.py --logout` → `login.py` |
| `invalid_request: ... 'redirect_uri' ...`(ログイン時) | Azure手順7のリダイレクトURI未登録 | [AZURE.md](AZURE.md) 手順7 |
| `device code flow を開始できませんでした` | Azure手順6「パブリック クライアント フローを許可する」が無効 | [AZURE.md](AZURE.md) 手順6 |
| `.env に OUTLOOK_CLIENT_ID がありません` | 手順3が未完了、または値が空 | 利用者にGUIDを尋ねる |
| `書き込み: 無効 (OUTLOOK_READONLY=true)` | 読み取り専用モード | 意図的ならそのまま。書き込みたいなら `.env` を `false` に → `/mcp` で再接続 |
| ツールが `/mcp` に出てこない | 登録後に再接続していない | `/mcp` から再接続。それでも出ないなら `claude mcp list` で登録を確認 |

**`.env` を書き換えたら、必ず `/mcp` からの再接続が要ります。** サーバは起動時に読むためです。

---

## 動作確認(任意)

利用者が求めた場合のみ。Microsoft Graph には接続せず、メールボックスも変更しません。

```bash
.venv/bin/pip install pytest
.venv/bin/pytest -q              # 単体テスト
.venv/bin/python smoke_test.py   # stdio 経由の外形テスト
```

---

## 完了報告に含めること

セットアップが済んだら、利用者に次を伝えてください。

1. `check_config` の結果(接続先アカウント、受信トレイ件数、ルール機能の可否)
2. 登録スコープ(local / user)
3. `token_cache.json` は共有・コミットしないこと
4. できないこと — **メールの送信と完全削除は実装されていない**。削除はゴミ箱への移動のみ

利用者が次に何をしたいか(受信トレイの整理、フォルダ再編、振分ルールの作成)を尋ねる前に、**まず `list_folders` で現状を見せる**と話が早いです。
